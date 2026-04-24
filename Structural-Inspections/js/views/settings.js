/**
 * Settings view (Phase 8).
 *
 * Editable inspector profile (name, RPEQ, CPEng, email, company, role) saved
 * into the settings table under key 'user.profile'. The profile's name flows
 * through to the "Inspector" field when starting a new inspection, and the
 * RPEQ/CPEng lines appear in the report sign-off block.
 *
 * Also: read-only diagnostics + destructive "wipe local database" button for
 * testing. Sync / OneDrive integration is still out of scope.
 */

import { db, getUserProfile, saveUserProfile } from '../db.js';
import { toast } from '../components/toast.js';

export async function render(root) {
  const [profile, inspectionTypesCount, drawingsCount, projectsCount,
         inspectionsCount, reportsCount] = await Promise.all([
    getUserProfile(),
    db.inspectionTypes.count(),
    db.drawings.count(),
    db.projects.count(),
    db.inspections.count(),
    db.reports.count()
  ]);

  const storageEstimate = await (navigator.storage?.estimate?.() ?? Promise.resolve(null));

  root.innerHTML = `
    <div class="section-heading">
      <h1>Settings</h1>
    </div>

    <div class="stack-lg">
      <div class="card">
        <h3>Inspector profile</h3>
        <p class="muted">Used to pre-fill the Inspector field on new inspections and to stamp the signature block on generated reports.</p>
        <form id="profile-form" class="form" autocomplete="on">
          <div class="form-field">
            <label for="pf-name">Name <span class="req">*</span></label>
            <input id="pf-name" name="name" type="text" required
              value="${escapeAttr(profile.name)}" placeholder="e.g. Dan McCarron"/>
          </div>
          <div class="form-field">
            <label for="pf-role">Role</label>
            <input id="pf-role" name="role" type="text"
              value="${escapeAttr(profile.role)}" placeholder="e.g. Senior Structural Engineer"/>
          </div>
          <div class="form-field">
            <label for="pf-rpeq">RPEQ number</label>
            <input id="pf-rpeq" name="rpeq" type="text" inputmode="numeric"
              value="${escapeAttr(profile.rpeq)}" placeholder="e.g. 12345"/>
            <div class="form-field__help muted">Printed below your name on the report sign-off block.</div>
          </div>
          <div class="form-field">
            <label for="pf-cpeng">CPEng number (optional)</label>
            <input id="pf-cpeng" name="cpeng" type="text"
              value="${escapeAttr(profile.cpeng)}" placeholder="e.g. NER 987654"/>
          </div>
          <div class="form-field">
            <label for="pf-email">Email</label>
            <input id="pf-email" name="email" type="email"
              value="${escapeAttr(profile.email)}" placeholder="e.g. dan.mccarron@blightanner.com.au"/>
          </div>
          <div class="form-field">
            <label for="pf-company">Company</label>
            <input id="pf-company" name="company" type="text"
              value="${escapeAttr(profile.company)}" placeholder="Bligh Tanner Pty Ltd"/>
          </div>

          <div class="form-field">
            <label for="pf-builder-email">Default builder email (optional)</label>
            <input id="pf-builder-email" name="builderEmail" type="email"
              value="${escapeAttr(profile.builderEmail || '')}" placeholder="e.g. super@itcconstructions.com.au"/>
            <div class="form-field__help muted">Used to pre-fill the close-out QR mailto link on generated reports.</div>
          </div>

          <div class="form-field">
            <label for="pf-anthropic">Anthropic API key (for AI comment expansion)</label>
            <input id="pf-anthropic" name="anthropicKey" type="password" autocomplete="off"
              value="${escapeAttr(profile.anthropicKey || '')}" placeholder="sk-ant-..."/>
            <div class="form-field__help muted">Stored on this device only. Used to polish short-form notes into AS-referenced comments. Leave blank to disable AI expansion.</div>
          </div>

          <div class="form-field">
            <label for="pf-signature">Digital signature (PNG / JPEG)</label>
            <div class="signature-field" id="pf-signature-field">
              <div class="signature-field__preview" id="pf-signature-preview">
                ${profile.signatureBlob
                  ? '<span class="muted">Loading preview\u2026</span>'
                  : '<span class="muted">No signature on file</span>'}
              </div>
              <div class="cluster">
                <input id="pf-signature" name="signature" type="file" accept="image/*" class="sr-only"/>
                <label class="btn btn--secondary btn--sm" for="pf-signature">Upload signature</label>
                <button type="button" class="btn btn--ghost btn--sm" id="pf-signature-clear"
                  ${profile.signatureBlob ? '' : 'disabled'}>Remove</button>
              </div>
            </div>
            <div class="form-field__help muted">
              Transparent-background PNG works best. Appears above your typed name in the report sign-off.
            </div>
          </div>

          <div class="cluster cluster--end">
            <button type="submit" class="btn btn--primary">Save profile</button>
          </div>
        </form>
      </div>

      <div class="card card--muted">
        <h3>Diagnostics</h3>
        <dl class="stack">
          <div><dt class="strong">App version</dt> <dd>v2.0.0 (Phase 10)</dd></div>
          <div><dt class="strong">Service worker</dt> <dd>${'serviceWorker' in navigator ? 'supported' : 'not supported'}</dd></div>
          <div><dt class="strong">Native share (files)</dt> <dd>${typeof navigator.canShare === 'function' ? 'supported' : 'not supported'}</dd></div>
          <div><dt class="strong">Inspection types seeded</dt> <dd>${inspectionTypesCount}</dd></div>
          <div><dt class="strong">Projects</dt> <dd>${projectsCount}</dd></div>
          <div><dt class="strong">Drawings</dt> <dd>${drawingsCount}</dd></div>
          <div><dt class="strong">Inspections</dt> <dd>${inspectionsCount}</dd></div>
          <div><dt class="strong">Reports stored</dt> <dd>${reportsCount}</dd></div>
          ${storageEstimate ? `
            <div><dt class="strong">Storage used</dt>
                 <dd>${formatBytes(storageEstimate.usage)} of ${formatBytes(storageEstimate.quota)} available
                 ${storageEstimate.quota ? `(${(100 * storageEstimate.usage / storageEstimate.quota).toFixed(1)}%)` : ''}</dd></div>
          ` : ''}
        </dl>
      </div>

      <div class="card">
        <h3>Cloud sync</h3>
        <p class="muted">Not wired for v1. Reports are saved locally under each inspection and shared via the native share sheet or email on generation. Cross-device sync via SharePoint / OneDrive is a v2 consideration.</p>
      </div>

      <div class="card">
        <h3 style="color: var(--color-danger);">Danger zone</h3>
        <p class="muted">Delete all local data — projects, drawings, inspections, items, photos, reports. Useful for fresh-start testing.</p>
        <button class="btn btn--danger" id="btn-wipe">Wipe local database</button>
      </div>
    </div>
  `;

  const form = root.querySelector('#profile-form');

  // --- Signature field: staged blob updates applied on submit ---
  const sigInput   = root.querySelector('#pf-signature');
  const sigClear   = root.querySelector('#pf-signature-clear');
  const sigPreview = root.querySelector('#pf-signature-preview');
  let stagedSignature = undefined;  // undefined = no change, null = clear, Blob = replace
  let stagedSignatureUrl = null;

  function paintSignaturePreview(blob) {
    if (stagedSignatureUrl) { URL.revokeObjectURL(stagedSignatureUrl); stagedSignatureUrl = null; }
    if (!blob) {
      sigPreview.innerHTML = '<span class="muted">No signature on file</span>';
      sigClear.disabled = true;
      return;
    }
    stagedSignatureUrl = URL.createObjectURL(blob);
    sigPreview.innerHTML = `<img src="${stagedSignatureUrl}" alt="Signature preview" />`;
    sigClear.disabled = false;
  }

  // Paint existing signature if present
  if (profile.signatureBlob) paintSignaturePreview(profile.signatureBlob);

  sigInput.addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file) return;
    if (!/^image\//.test(file.type)) {
      toast('Please pick a PNG or JPEG image', { kind: 'error' });
      return;
    }
    stagedSignature = file;
    paintSignaturePreview(file);
  });

  sigClear.addEventListener('click', () => {
    stagedSignature = null;
    paintSignaturePreview(null);
  });

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(form).entries());
    delete data.signature;   // FormData picks up the file input — we handle it via stagedSignature
    if (stagedSignature !== undefined) data.signatureBlob = stagedSignature;
    try {
      await saveUserProfile(data);
      toast('Profile saved', { kind: 'success' });
      // Reset staged state — saved signature is now canonical
      stagedSignature = undefined;
    } catch (err) {
      console.error(err);
      toast(`Couldn\u2019t save: ${err.message}`, { kind: 'error' });
    }
  });

  root.querySelector('#btn-wipe').addEventListener('click', async () => {
    if (!confirm('Wipe ALL local data? This cannot be undone.')) return;
    await db.delete();
    location.reload();
  });
}

function formatBytes(n) {
  if (!n && n !== 0) return '—';
  const units = ['B', 'KB', 'MB', 'GB'];
  let i = 0;
  while (n >= 1024 && i < units.length - 1) { n /= 1024; i++; }
  return `${n.toFixed(1)} ${units[i]}`;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}
function escapeAttr(s) { return escapeHtml(s).replace(/"/g, '&quot;'); }
