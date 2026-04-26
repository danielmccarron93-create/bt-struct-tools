/**
 * .btproject bundle reader + validator.
 *
 * A `.btproject` file is a ZIP produced by Cowork via the procedure in
 * `<OneDrive>/22_Claude/Structural Inspections/Drawings/CLAUDE.md`. It contains
 * the structured project map plus copies of every source PDF, ready to import
 * into the BT Inspect PWA in one tap.
 *
 * Bundle layout:
 *
 *   project.btproject  (ZIP)
 *   ├── project-map.json
 *   ├── manifest.txt           (human-readable summary)
 *   └── Drawings/
 *       ├── BT 52 Second Av.pdf
 *       └── …
 *
 * This module:
 *   - Loads the bundle via JSZip (loaded globally in index.html).
 *   - Parses project-map.json.
 *   - Validates required fields (best-effort, not full JSON Schema — friendly
 *     errors on malformed bundles).
 *   - Reads each PDF as a Blob, ready to insert into pdfSources.
 *   - Returns a structured object the import flow can consume.
 *
 * Validation philosophy: catch the things that would corrupt the database or
 * make the app crash later, but tolerate cosmetic gaps. If `generalNotes` is
 * missing, that's fine — a Cowork run on a project without a Notes page
 * should still land cleanly.
 */

/* global JSZip */

const SUPPORTED_SCHEMA_VERSION = 1;

/**
 * Read a .btproject Blob and return the structured contents.
 *
 * @param {Blob} blob — the .btproject file
 * @returns {Promise<{
 *   projectMap: Object,            // the parsed project-map.json
 *   pdfBlobs: Array<{ filename: string, blob: Blob, sizeBytes: number }>,
 *   manifest:  string              // raw manifest.txt or '' if missing
 * }>}
 */
export async function readBundle(blob) {
  if (!blob || typeof blob.arrayBuffer !== 'function') {
    throw new Error('Pick a .btproject file to import.');
  }
  if (typeof JSZip === 'undefined') {
    throw new Error('JSZip not loaded — refresh the app and try again.');
  }

  let zip;
  try {
    zip = await JSZip.loadAsync(blob);
  } catch (err) {
    throw new Error(`Couldn't read that as a .btproject file (not a valid zip).`);
  }

  // 1) project-map.json (required)
  const mapEntry = zip.file('project-map.json');
  if (!mapEntry) {
    throw new Error('Bundle is missing project-map.json — was this produced by Cowork?');
  }
  let projectMap;
  try {
    const text = await mapEntry.async('string');
    projectMap = JSON.parse(text);
  } catch (err) {
    throw new Error(`Couldn't parse project-map.json: ${err.message || err}`);
  }
  validateProjectMap(projectMap);

  // 2) manifest.txt (optional)
  let manifest = '';
  const manifestEntry = zip.file('manifest.txt');
  if (manifestEntry) {
    try { manifest = await manifestEntry.async('string'); } catch {}
  }

  // 3) PDFs from Drawings/ folder. Match the filenames declared in
  // projectMap.sourceFiles[] so we can flag missing/extra files.
  const expected = (projectMap.sourceFiles || []).map((f) => f.filename);
  const actualFiles = [];
  zip.folder('Drawings')?.forEach((relPath, fileEntry) => {
    if (fileEntry.dir) return;
    if (!/\.pdf$/i.test(relPath)) return;
    actualFiles.push({ relPath, fileEntry });
  });

  const pdfBlobs = [];
  for (const entry of actualFiles) {
    const filename = basename(entry.relPath);
    const blob = await entry.fileEntry.async('blob');
    // JSZip returns a generic Blob — re-tag with the right MIME type.
    const pdfBlob = blob.type === 'application/pdf'
      ? blob
      : new Blob([await blob.arrayBuffer()], { type: 'application/pdf' });
    pdfBlobs.push({ filename, blob: pdfBlob, sizeBytes: pdfBlob.size });
  }

  // Cross-check: every sourceFiles entry should have a matching PDF.
  const haveFilenames = new Set(pdfBlobs.map((p) => p.filename));
  const missing = expected.filter((f) => !haveFilenames.has(f));
  if (missing.length > 0) {
    throw new Error(
      `Bundle is missing source PDF${missing.length === 1 ? '' : 's'}: ${missing.join(', ')}`
    );
  }
  // Extras (PDFs not in the map) are fine — log but don't fail.
  const extras = pdfBlobs
    .map((p) => p.filename)
    .filter((f) => !expected.includes(f));
  if (extras.length > 0) {
    // Not fatal; CLAUDE.md may have evolved between bundle generations.
    // We'll still import them as un-mapped sources.
    console.warn('Bundle has extra PDFs not referenced in project-map.json:', extras);
  }

  return { projectMap, pdfBlobs, manifest };
}

/**
 * Validate a parsed project-map object. Throws on hard failures (missing
 * required fields, wrong schema version). Returns a list of soft warnings
 * for things that look off but won't crash the import.
 */
export function validateProjectMap(projectMap) {
  if (!projectMap || typeof projectMap !== 'object') {
    throw new Error('project-map.json is not a valid JSON object.');
  }
  const sv = projectMap.schemaVersion;
  if (sv !== SUPPORTED_SCHEMA_VERSION) {
    throw new Error(
      `Bundle uses schema version ${sv}; this app supports v${SUPPORTED_SCHEMA_VERSION}. ` +
      `Re-run Cowork to regenerate against the current schema.`
    );
  }
  if (!projectMap.project || typeof projectMap.project !== 'object') {
    throw new Error('project-map.json is missing the "project" object.');
  }
  if (!String(projectMap.project.name || '').trim()) {
    throw new Error('project-map.json has no project name.');
  }
  if (!Array.isArray(projectMap.sourceFiles) || projectMap.sourceFiles.length === 0) {
    throw new Error('project-map.json has no sourceFiles[] — at least one PDF is required.');
  }
  if (!Array.isArray(projectMap.drawings)) {
    throw new Error('project-map.json is missing the drawings[] array.');
  }
  if (!Array.isArray(projectMap.inspectionPlan)) {
    throw new Error('project-map.json is missing the inspectionPlan[] array.');
  }

  // Soft sanity checks — surface in the import preview but don't block.
  const soft = [];
  if (projectMap.drawings.length === 0) {
    soft.push('inspectionPlan exists but no drawings were classified.');
  }
  if (projectMap.inspectionPlan.length === 0) {
    soft.push('No inspections were proposed in the plan.');
  }
  // Every drawing references a sourceFile that must exist
  const sourceNames = new Set(projectMap.sourceFiles.map((f) => f.filename));
  const orphans = projectMap.drawings.filter((d) => !sourceNames.has(d.sourceFile));
  if (orphans.length) {
    soft.push(`${orphans.length} drawing entr${orphans.length === 1 ? 'y' : 'ies'} reference a sourceFile that's not declared.`);
  }
  // Inspection plan drawingRefs should resolve to sheet numbers we have
  const sheetSet = new Set(
    projectMap.drawings.map((d) => (d.sheetNumber || '').toUpperCase()).filter(Boolean)
  );
  let unknownRefs = 0;
  for (const entry of projectMap.inspectionPlan) {
    for (const ref of (entry.drawingRefs || [])) {
      const s = (ref.sheetNumber || '').toUpperCase();
      if (s && !sheetSet.has(s)) unknownRefs += 1;
    }
  }
  if (unknownRefs > 0) {
    soft.push(`${unknownRefs} inspection-plan drawingRef${unknownRefs === 1 ? '' : 's'} point at sheet number${unknownRefs === 1 ? '' : 's'} not in the drawings list.`);
  }

  return soft;
}

/**
 * Produce a one-line summary of a parsed project-map for the import preview.
 */
export function summariseBundle(projectMap, pdfBlobs) {
  const p = projectMap.project || {};
  const lines = [];
  lines.push(`${p.jobNumber ? p.jobNumber + ' — ' : ''}${p.name || '(unnamed)'}`);
  if (p.siteAddress) lines.push(p.siteAddress);
  if (p.client)      lines.push(`Client: ${p.client}`);
  if (p.issueStatus) lines.push(`Issue: ${p.issueStatus}`);
  lines.push(`${projectMap.drawings.length} drawing${projectMap.drawings.length === 1 ? '' : 's'} across ${pdfBlobs.length} PDF${pdfBlobs.length === 1 ? '' : 's'}`);
  lines.push(`${projectMap.inspectionPlan.length} inspection${projectMap.inspectionPlan.length === 1 ? '' : 's'} proposed`);
  if ((projectMap.warnings || []).length) {
    lines.push(`${projectMap.warnings.length} warning${projectMap.warnings.length === 1 ? '' : 's'}`);
  }
  return lines.join('\n');
}

function basename(p) {
  return String(p || '').split('/').pop().split('\\').pop();
}
