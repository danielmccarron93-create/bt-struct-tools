/**
 * Sharing helpers (Phase 7).
 *
 * Two paths:
 *   - Mobile Safari / Android Chrome — navigator.share with a File attached,
 *     which brings up the OS share sheet (Mail, WhatsApp, AirDrop, Files…).
 *   - Desktop / older browsers — fall back to triggering a download and
 *     opening a pre-composed mailto (email client launches; user attaches
 *     the now-downloaded file manually — mailto specs don't allow file
 *     attachments directly).
 *
 * Exports:
 *   canShareFiles()               — feature-detect files-on-share
 *   shareReport(blob, filename,   — attempt share; fall back to download+mailto
 *                meta)
 *   buildEmailDraft(meta)         — { subject, body } strings
 *   triggerDownload(blob, name)   — plain download helper
 *   openPdfPreview(blob)          — open a blob in a new tab; returns the URL
 *                                    (caller must revoke later)
 */

/* --------------------------------------------------------------------------
   Feature detection
   -------------------------------------------------------------------------- */
export function canShareFiles() {
  if (typeof navigator === 'undefined') return false;
  if (typeof navigator.canShare !== 'function') return false;
  try {
    // A small sacrificial File just to probe capability.
    const probe = new File([new Blob(['x'])], 'probe.txt', { type: 'text/plain' });
    return navigator.canShare({ files: [probe] });
  } catch {
    return false;
  }
}

/* --------------------------------------------------------------------------
   Share a report blob
   --------------------------------------------------------------------------
   Returns { method: 'share' | 'download+mailto' | 'cancelled' | 'error', error? }.
   Callers can use this to decide whether to show toasts.
   -------------------------------------------------------------------------- */
export async function shareReport(blob, filename, meta = {}) {
  // Preferred: native share sheet with the file attached.
  if (canShareFiles()) {
    try {
      const file = new File([blob], filename, { type: 'application/pdf' });
      const { subject, body } = buildEmailDraft(meta);
      await navigator.share({
        files: [file],
        title: subject,
        text:  body
      });
      return { method: 'share' };
    } catch (err) {
      // User cancelled the share sheet — normal on iOS if they tap outside.
      if (err && (err.name === 'AbortError' || /abort/i.test(err.message || ''))) {
        return { method: 'cancelled' };
      }
      // Fall through to download+mailto fallback
      console.warn('navigator.share failed, falling back:', err);
    }
  }

  // Fallback: download the PDF + open mailto composer.
  try {
    triggerDownload(blob, filename);
    const { subject, body } = buildEmailDraft(meta, { noteFallback: true });
    // Use a slight delay so the download prompt doesn't race the mailto handler.
    setTimeout(() => {
      const url = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
      window.location.href = url;
    }, 150);
    return { method: 'download+mailto' };
  } catch (err) {
    return { method: 'error', error: err };
  }
}

/* --------------------------------------------------------------------------
   Email draft builder
   --------------------------------------------------------------------------
   meta = { project, inspection } — both may be null if absent.
   opts.noteFallback — true adds a line telling the recipient to attach the
                       file manually (only relevant on desktop fallback path).
   -------------------------------------------------------------------------- */
export function buildEmailDraft(meta = {}, { noteFallback = false } = {}) {
  const { project, inspection } = meta;
  const jobNum = project?.jobNumber || '';
  const projectName = project?.name || '';
  const typeName = inspection?.inspectionTypeName || 'Structural inspection';
  const date = inspection?.date
    ? new Date(inspection.date).toLocaleDateString('en-AU', { day: '2-digit', month: 'short', year: 'numeric' })
    : '';

  const subject = [
    jobNum ? `[${jobNum}]` : '',
    'BT Inspection Report —',
    typeName,
    date ? `(${date})` : ''
  ].filter(Boolean).join(' ').trim();

  const lines = [
    'Hello,',
    '',
    `Please find attached the structural inspection report for:`,
    jobNum || projectName
      ? `    ${jobNum ? jobNum + ' — ' : ''}${projectName}`
      : '',
    `    ${typeName}${date ? ', ' + date : ''}`,
    '',
    'Items requiring rectification are listed in the schedule. Re-inspection may be required once items are addressed.',
    ''
  ];
  if (noteFallback) {
    lines.push(
      '[The PDF has been downloaded separately — please attach it to this email before sending.]',
      ''
    );
  }
  lines.push(
    'Regards,',
    inspection?.inspectorName || 'Bligh Tanner Pty Ltd'
  );

  return { subject, body: lines.filter((l) => l !== null).join('\n') };
}

/* --------------------------------------------------------------------------
   Plain download helper — used on its own and by the fallback path.
   -------------------------------------------------------------------------- */
export function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  try {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.rel = 'noopener';
    document.body.appendChild(a);
    a.click();
    a.remove();
  } finally {
    // Give the browser a moment to actually start the download before revoking.
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }
}

/**
 * Open a blob in a new browser tab for preview. Caller is responsible for
 * revoking the URL at an appropriate later time.
 */
export function openPdfPreview(blob) {
  const url = URL.createObjectURL(blob);
  try {
    window.open(url, '_blank', 'noopener');
  } catch {
    // Popup blocked — fall back to triggering a download
    const a = document.createElement('a');
    a.href = url;
    a.target = '_blank';
    document.body.appendChild(a);
    a.click();
    a.remove();
  }
  return url;
}
