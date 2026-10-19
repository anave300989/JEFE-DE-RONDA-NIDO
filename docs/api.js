// Llama al backend de Google Apps Script. Se usa text/plain para evitar el bloqueo CORS.
async function api(accion, datos) {
  const res = await fetch(window.API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(Object.assign({ accion }, datos || {}))
  });
  const j = await res.json();
  if (!j.ok) throw new Error(j.error);
  return j.data;
}

function $(s, el) { return (el || document).querySelector(s); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]); }
function aviso(el, texto, tipo) { el.className = 'msg ' + (tipo || ''); el.textContent = texto; el.classList.remove('hidden'); }
