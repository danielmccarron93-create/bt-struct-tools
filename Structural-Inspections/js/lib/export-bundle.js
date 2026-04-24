/**
 * Export bundle (Phase V2) — one-click archive in the north-star §9.5 layout.
 *
 * Produces a ZIP:
 *
 *   YYYY-MM-DD_<InspectionType>_<Inspector>/
 *   ├── Report.pdf                  # the generated PDF
 *   ├── Report_source.json          # editable JSON snapshot of the inspection
 *   └── Photos/
 *       ├── YYYY-MM-DD_<project>_<inspection>_Item01_photo1.jpg
 *       └── ...
 *
 * This lets the engineer drop one file into OneDrive / SharePoint without
 * waiting for the real MSAL-based sync. When OneDrive integration lands the
 * UI stays the same — this module gets replaced with a direct upload.
 */

import {
  getInspection, getProject, getDrawing, listItemsForInspection,
  listPhotosForItem, listHighlightsForInspection,
  getUserProfile, saveReport
} from '../db.js';
import { generateReport } from './report.js';
import { renderAnnotatedImage } from './annotate.js';

/**
 * Build the bundle as a Blob. Triggers a browser download via `saveAs`
 * style (the caller does the download).
 *
 * opts.onProgress(msg) optional — status pings for the UI.
 */
export async function buildExportBundle(inspectionId, { onProgress } = {}) {
  const notify = (msg) => { try { onProgress?.(msg); } catch {} };

  if (!window.JSZip) throw new Error('JSZip not loaded — check your network on first use.');

  notify('Generating report PDF\u2026');
  const { blob: pdfBlob, filename: pdfName, pageCount } =
    await generateReport(inspectionId);

  notify('Gathering source data\u2026');
  const inspection = await getInspection(inspectionId);
  const [project, drawing, profile, items, highlights] = await Promise.all([
    getProject(inspection.projectId),
    inspection.primaryDrawingId ? getDrawing(inspection.primaryDrawingId) : Promise.resolve(null),
    getUserProfile(),
    listItemsForInspection(inspection.id),
    listHighlightsForInspection(inspection.id)
  ]);

  const itemsWithPhotos = await Promise.all(items.map(async (it) => ({
    ...it,
    photos: await listPhotosForItem(it.id)
  })));

  // Build the JSON snapshot — strip blobs, only store metadata + counts there.
  const source = {
    exportedAt: new Date().toISOString(),
    schemaVersion: 1,
    project: project ? {
      id: project.id,
      jobNumber: project.jobNumber,
      name: project.name,
      client: project.client,
      siteAddress: project.siteAddress,
      notes: project.notes,
      createdAt: project.createdAt,
      updatedAt: project.updatedAt
    } : null,
    drawing: drawing ? {
      id: drawing.id,
      sheetNumber: drawing.sheetNumber,
      revision: drawing.revision,
      description: drawing.description,
      filename: drawing.filename,
      pageNumber: drawing.pageNumber,
      pageWidth: drawing.pageWidth,
      pageHeight: drawing.pageHeight,
      rotation: drawing.rotation,
      calibration: drawing.calibration || null,
      gridCalibration: drawing.gridCalibration || null
    } : null,
    inspection: {
      id: inspection.id,
      inspectionTypeKey: inspection.inspectionTypeKey,
      inspectionTypeName: inspection.inspectionTypeName,
      inspectionTypeCategory: inspection.inspectionTypeCategory,
      date: inspection.date,
      inspectorName: inspection.inspectorName,
      attendees: inspection.attendees,
      weather: inspection.weather,
      notes: inspection.notes,
      status: inspection.status,
      generalComments: inspection.generalComments || [],
      createdAt: inspection.createdAt,
      updatedAt: inspection.updatedAt
    },
    highlights: highlights.map((h) => ({
      id: h.id, drawingId: h.drawingId,
      pdfX: h.pdfX, pdfY: h.pdfY, pdfW: h.pdfW, pdfH: h.pdfH
    })),
    items: itemsWithPhotos.map((it) => ({
      id: it.id,
      itemNumber: it.itemNumber,
      severity: it.severity,
      status: it.status,
      gridRef: it.gridRef || '',
      pdfX: it.pdfX, pdfY: it.pdfY,
      comment: it.comment,
      libraryKey: it.libraryKey || null,
      asClause: it.asClause || '',
      createdAt: it.createdAt,
      updatedAt: it.updatedAt,
      photos: (it.photos || []).map((p) => ({
        id: p.id,
        caption: p.caption || '',
        width: p.width, height: p.height,
        annotations: p.annotations || [],
        createdAt: p.createdAt,
        // Path inside the bundle — matches what we write in Photos/
        file: photoFilename(project, inspection, it, p)
      }))
    })),
    inspector: {
      name:    profile?.name    || '',
      email:   profile?.email   || '',
      rpeq:    profile?.rpeq    || '',
      cpeng:   profile?.cpeng   || '',
      company: profile?.company || '',
      role:    profile?.role    || ''
    },
    report: {
      filename: pdfName,
      pageCount
    }
  };

  notify('Assembling ZIP\u2026');
  const zip = new window.JSZip();
  const folder = bundleFolderName(project, inspection);
  const root = zip.folder(folder);
  root.file('Report.pdf', pdfBlob);
  root.file('Report_source.json', JSON.stringify(source, null, 2));

  const photoDir = root.folder('Photos');
  for (const it of itemsWithPhotos) {
    for (const p of (it.photos || [])) {
      const fname = photoFilename(project, inspection, it, p);
      try {
        let blob = p.blob;
        // If annotated, bake the vector overlay so the exported photo
        // matches what's in the PDF.
        if (Array.isArray(p.annotations) && p.annotations.length) {
          const { dataUrl } = await renderAnnotatedImage(p.blob, p.annotations, { maxLongEdge: 2000 });
          blob = await (await fetch(dataUrl)).blob();
        }
        photoDir.file(fname, blob);
      } catch (err) {
        console.warn('Skipping photo in bundle:', err);
      }
    }
  }

  notify('Finalising\u2026');
  const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });

  // Persist the generated PDF under the inspection for quick re-open, matching
  // the behaviour of the plain "Generate report" button.
  try {
    await saveReport(inspectionId, {
      filename: pdfName,
      pdfBlob,
      sizeBytes: pdfBlob.size,
      pageCount,
      generatedBy: inspection.inspectorName || ''
    });
  } catch (err) {
    console.warn('Report persist (bundle) failed:', err);
  }

  notify('Done');
  return {
    blob: zipBlob,
    filename: `${folder}.zip`,
    folderName: folder
  };
}

function bundleFolderName(project, inspection) {
  const date  = (inspection?.date || new Date().toISOString().slice(0, 10));
  const type  = sanitize(inspection?.inspectionTypeName || 'Inspection');
  const insp  = sanitize(inspection?.inspectorName || 'Inspector');
  const job   = sanitize(project?.jobNumber || 'JOB');
  return `${date}_${job}_${type}_${insp}`;
}

function photoFilename(project, inspection, item, photo) {
  const date = (inspection?.date || new Date().toISOString().slice(0, 10));
  const job  = sanitize(project?.jobNumber || 'JOB');
  const type = sanitize(inspection?.inspectionTypeName || 'Inspection');
  const n    = String(item.itemNumber).padStart(2, '0');
  // Photos per item aren't explicitly numbered; index them by position.
  // We pass the original id as the suffix so filenames stay unique even if
  // the same item has multiple photos added over time.
  return `${date}_${job}_${type}_Item${n}_photo${photo.id}.jpg`;
}

function sanitize(s) {
  return String(s || '')
    .replace(/[^\w\-]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}
