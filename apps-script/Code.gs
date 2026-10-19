/**
 * Sistema de Control del Jefe de Ronda - ESCONBOL
 * Backend en Google Apps Script (base de datos en Google Sheets, archivos en Google Drive).
 *
 * Hojas que usa (se crean solas con instalar()):
 *   Responsables : semana (lunes), grado_nombre, celular, pin
 *   SantoSena    : fecha, santo, sena
 *   Comandantes  : cargo, grado_nombre, celular, clave_callmebot
 *   Rondas       : id, fecha, grado_nombre, celular, token, estado, enviado, completado, pdf_id, fotos, novedades
 *
 * Propiedades del script (Configuración del proyecto > Propiedades del script):
 *   ADMIN_PASSWORD   : contraseña del administrador
 *   WEB_URL          : dirección de la página en GitHub Pages (ej. https://usuario.github.io/control-jefe-ronda/)
 *   WA_TOKEN, WA_PHONE_ID : (opcional) API oficial de WhatsApp Cloud. Si están, el envío es automático.
 */

const TZ = 'America/La_Paz';
const HORA_ENVIO = 8;

const LISTA = [
  { titulo: 'PREVENCIÓN DE GUARDIA Y CONTROL GENERAL', items: [
    'Identificación correcta ante el centinela e intercambio de Santo y Seña.',
    'Recepción del parte del Capitán de Servicio sobre novedades.',
    'Anotación personal en el Libro de Novedades con la hora exacta de la ronda.',
    'Existen planes y Reglamentos en la Prevención (Acción y Reacción, otros).',
    'Nombrados cuentan con munición (No letal y Guerra) en cargadores.',
    'El personal de soldados conoce su Santo y Seña.',
    'Libros de Novedades, Visitas, Vehículos, Sanciones y Encomiendas, correctamente llenados.',
    'Personal de la guardia que se encuentra descansando lo hace con el equipo y armamento listo y sin desvestirse.',
    'Centinelas se comunican entre sí pasando el número cada determinado tiempo o mediante señales establecidas.'
  ]},
  { titulo: 'MATERIAL BÉLICO Y POLVORINES', items: [
    'Centinelas resguardan y vigilan áreas críticas (Polvorines/Clase V).',
    'Iluminación del área de polvorines.',
    'Cierre verificado de las puertas de los depósitos C-V y material bélico.',
    'Material y equipos contra incendios (extintores, arena, etc.)'
  ]},
  { titulo: 'DEPÓSITOS DE COMBUSTIBLE Y TRANSPORTES', items: [
    'Cierre verificado de las puertas de los depósitos de combustible.',
    'Verificación de que no exista combustible derramado.',
    'Material y equipos contra incendios (extintores, arena, etc.)',
    'Vehículos en parqueos asignados con candados en los tanques de combustible'
  ]},
  { titulo: 'COCINAS', items: [
    'Cierre correcto de puertas de acceso a las cocinas.',
    'Corte del fluido de gas comprobado (llaves de paso de garrafas, mangueras bien conectadas)',
    'Ausencia de personal manipulando cocinas fuera de horario.',
    'Materiales y herramientas industriales correctamente guardados.',
    'Utensilios de cocina correctamente guardados.',
    'Material y los equipos contra incendios (extintores, arena, etc.)'
  ]},
  { titulo: 'SANIDAD', items: [
    'Personal de turno se encuentra presente y pernocta en la enfermería para atender cualquier emergencia nocturna.',
    'Orden, disciplina e higiene estricta en la enfermería y áreas de internación.',
    'Insumos médicos se encuentran resguardados.',
    'Soldados enfermos internados cumplen el reposo prescrito por el Médico de la Unidad.'
  ]},
  { titulo: 'CABALLERIZAS', items: [
    'Soldados de servicio de caballerizas descansando en el cuarto de los caballerizos.',
    'Caballos no se encuentran enredados, sueltos o en riesgo de lesión.',
    'Ausencia de conexiones eléctricas clandestinas o en deterioro'
  ]},
  { titulo: 'TALLERES', items: [
    'Cierre correcto de puertas de acceso a los talleres.',
    'Corte del fluido eléctrico comprobado (llaves, motores apagados).',
    'Existe personal manipulando maquinaria fuera de horario.',
    'Materiales y herramientas industriales correctamente guardados.',
    'Material y los equipos contra incendios (extintores, arena, etc.)'
  ]},
  { titulo: 'DORMITORIOS, CUADRAS Y RÉGIMEN INTERNO', items: [
    'Oficiales de Servicio de permanencia continua con el personal de Tropa de su compañía, con libro de novedades al día.',
    'Todos los soldados se encuentran descansando en la cuadra asignada.',
    'Disciplina de luces en las instalaciones de las cuadras',
    'Baños, duchas, lavamanos, tanques de agua correctamente cerrados',
    'Ausencia de conexiones eléctricas clandestinas o en deterioro'
  ]}
];

const HOJAS = {
  Responsables: ['semana', 'grado_nombre', 'celular', 'pin'],
  SantoSena: ['fecha', 'santo', 'sena'],
  Comandantes: ['cargo', 'grado_nombre', 'celular', 'clave_callmebot'],
  Rondas: ['id', 'fecha', 'grado_nombre', 'celular', 'token', 'estado', 'enviado', 'completado', 'pdf_id', 'fotos', 'novedades']
};

/* ---------- Instalación ---------- */

/** Ejecutar una sola vez desde el editor: crea hojas, carpeta y el envío diario de las 08:00. */
function instalar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HOJAS).forEach(n => {
    let sh = ss.getSheetByName(n);
    if (!sh) sh = ss.insertSheet(n);
    if (sh.getLastRow() === 0) sh.appendRow(HOJAS[n]);
    sh.getRange('A:Z').setNumberFormat('@');
  });
  const c = ss.getSheetByName('Comandantes');
  if (c.getLastRow() === 1) {
    c.appendRow(['CMDTE.', '', '']);
    c.appendRow(['2DO. CMDTE.', '', '']);
  }
  carpeta_();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('envioDiario').timeBased().atHour(HORA_ENVIO).everyDays(1).inTimezone(TZ).create();
}

function carpeta_(sub) {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('FOLDER_ID');
  let root;
  if (id) root = DriveApp.getFolderById(id);
  else {
    root = DriveApp.createFolder('Control Jefe de Ronda');
    props.setProperty('FOLDER_ID', root.getId());
  }
  if (!sub) return root;
  const it = root.getFoldersByName(sub);
  return it.hasNext() ? it.next() : root.createFolder(sub);
}

/* ---------- Utilidades de hojas ---------- */

function hoja_(n) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
  const cols = HOJAS[n].length;
  if (sh.getLastColumn() < cols) sh.getRange(1, 1, 1, cols).setValues([HOJAS[n]]);
  return sh;
}

function filas_(n) {
  const v = hoja_(n).getDataRange().getDisplayValues();
  const h = v.shift();
  return v.map((r, i) => { const o = { _fila: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}

function reemplazar_(n, pred, nuevas) {
  const sh = hoja_(n);
  const quedan = filas_(n).filter(r => !pred(r)).map(r => HOJAS[n].map(k => r[k]));
  const todas = quedan.concat(nuevas.map(o => HOJAS[n].map(k => o[k] == null ? '' : String(o[k]))));
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HOJAS[n].length).clearContent();
  if (todas.length) sh.getRange(2, 1, todas.length, HOJAS[n].length).setValues(todas);
}

function actualizarRonda_(id, cambios) {
  const sh = hoja_('Rondas');
  const r = filas_('Rondas').find(x => x.id === id);
  if (!r) return;
  Object.keys(cambios).forEach(k => sh.getRange(r._fila, HOJAS.Rondas.indexOf(k) + 1).setValue(String(cambios[k])));
}

function hoy_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ahora_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function lunes_(fecha) {
  const d = new Date(fecha + 'T12:00:00');
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
function token_() { return Utilities.getUuid().replace(/-/g, ''); }
function celular_(c) {
  let n = String(c || '').replace(/\D/g, '');
  if (n.length === 8) n = '591' + n;
  return n;
}

/* ---------- Envío diario ---------- */

/** Crea las rondas del día para los dos responsables de la semana y envía los mensajes. */
function envioDiario() {
  const fecha = hoy_();
  const resp = filas_('Responsables').filter(r => r.semana === lunes_(fecha));
  const ss = filas_('SantoSena').find(r => r.fecha === fecha) || { santo: '(sin cargar)', sena: '(sin cargar)' };
  const existentes = filas_('Rondas').filter(r => r.fecha === fecha);
  const url = PropertiesService.getScriptProperties().getProperty('WEB_URL') || '';
  resp.forEach(p => {
    let r = existentes.find(x => x.celular === celular_(p.celular));
    if (!r) {
      r = { id: fecha + '-' + celular_(p.celular), fecha: fecha, grado_nombre: p.grado_nombre,
            celular: celular_(p.celular), token: token_(), estado: 'PENDIENTE', enviado: '', completado: '', pdf_id: '', fotos: '', novedades: '' };
      hoja_('Rondas').appendRow(HOJAS.Rondas.map(k => r[k]));
    }
    const ok = enviarWhatsApp_(r.celular, mensajeRonda_(r, ss, url));
    if (ok) actualizarRonda_(r.id, { enviado: ahora_() });
  });
}

function mensajeRonda_(r, ss, url) {
  return '*CONTROL DE JEFE DE RONDA - ESCONBOL*\n' +
    r.grado_nombre + ', usted está designado como Jefe de Ronda para el ' + r.fecha + '.\n\n' +
    '*SANTO:* ' + ss.santo + '\n*SEÑA:* ' + ss.sena + '\n\n' +
    'Llene la lista de verificación (22:00 a 05:00) en:\n' + url + '?t=' + r.token;
}

function mensajeConfirmacion_(r) {
  return '*CONTROL DE JEFE DE RONDA - ESCONBOL*\n' +
    'Se pasó ronda el ' + r.fecha + '.\nJefe de Ronda: ' + r.grado_nombre + '\nLista de verificación registrada: ' + r.completado + '.';
}

/** Envía por la API oficial de WhatsApp si está configurada. Devuelve false si no lo está. */
function enviarWhatsApp_(numero, texto) {
  const p = PropertiesService.getScriptProperties();
  const tok = p.getProperty('WA_TOKEN'), phone = p.getProperty('WA_PHONE_ID');
  if (!tok || !phone || !numero) return false;
  const res = UrlFetchApp.fetch('https://graph.facebook.com/v20.0/' + phone + '/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + tok },
    payload: JSON.stringify({ messaging_product: 'whatsapp', to: numero, type: 'text', text: { body: texto } })
  });
  return res.getResponseCode() < 300;
}

/** Envía por CallMeBot (gratis). Cada destinatario activa su clave una sola vez. */
function enviarCallMeBot_(numero, clave, texto) {
  try {
    const res = UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=%2B' + numero +
      '&apikey=' + encodeURIComponent(clave) + '&text=' + encodeURIComponent(texto), { muteHttpExceptions: true });
    return res.getResponseCode() < 300 && /queued|sent/i.test(res.getContentText());
  } catch (e) { return false; }
}

/** Prueba: envía un mensaje de prueba por CallMeBot a cada comandante con clave. */
function probarAvisoComandantes() {
  filas_('Comandantes').filter(c => c.celular && c.clave_callmebot).forEach(c =>
    Logger.log(c.cargo + ': ' + (enviarCallMeBot_(celular_(c.celular), c.clave_callmebot, 'Prueba del sistema de Control de Jefe de Ronda.') ? 'enviado' : 'ERROR')));
}

function enlaceWa_(numero, texto) {
  return 'https://wa.me/' + numero + '?text=' + encodeURIComponent(texto);
}

/* ---------- API web ---------- */

function doGet() {
  return ContentService.createTextOutput('Control Jefe de Ronda: servicio activo.');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = ACCIONES[req.accion];
    if (!fn) throw new Error('Acción no válida');
    out = { ok: true, data: fn(req) };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

const ACCIONES = {
  /* Jefe de Ronda */
  abrir: req => {
    const r = rondaPorToken_(req.token, req.pin);
    return { fecha: r.fecha, grado_nombre: r.grado_nombre, estado: r.estado, lista: LISTA };
  },
  enviar: req => {
    const r = rondaPorToken_(req.token, req.pin);
    if (r.estado === 'COMPLETADO') throw new Error('Esta lista ya fue enviada.');
    validarEnvio_(req.datos);
    const carpeta = carpeta_(r.fecha);
    const fotos = req.datos.fotos.map((f, i) => {
      const blob = Utilities.newBlob(Utilities.base64Decode(f.split(',')[1]), 'image/jpeg', r.id + '-foto' + (i + 1) + '.jpg');
      return carpeta.createFile(blob).getId();
    });
    r.completado = ahora_();
    const pdf = generarPdf_(r, req.datos, carpeta);
    actualizarRonda_(r.id, { estado: 'COMPLETADO', completado: r.completado, pdf_id: pdf.getId(),
                             fotos: fotos.join(','), novedades: req.datos.novedades || '' });
    const msg = mensajeConfirmacion_(r);
    const avisos = filas_('Comandantes').filter(c => c.celular).map(c => {
      const n = celular_(c.celular);
      const auto = c.clave_callmebot ? enviarCallMeBot_(n, c.clave_callmebot, msg) : enviarWhatsApp_(n, msg);
      return { cargo: c.cargo, automatico: auto, enlace: enlaceWa_(n, msg) };
    });
    return { avisos: avisos.filter(a => !a.automatico).map(a => ({ cargo: a.cargo, enlace: a.enlace })) };
  },

  /* Administrador */
  login: req => {
    const pass = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
    if (!pass || req.password !== pass) { Utilities.sleep(1500); throw new Error('Contraseña incorrecta'); }
    const s = token_();
    CacheService.getScriptCache().put('adm_' + s, '1', 6 * 3600);
    return { sesion: s };
  },
  panel: req => {
    admin_(req);
    const semana = lunes_(req.fecha || hoy_());
    const fin = new Date(semana + 'T12:00:00'); fin.setDate(fin.getDate() + 6);
    const finS = Utilities.formatDate(fin, TZ, 'yyyy-MM-dd');
    const url = PropertiesService.getScriptProperties().getProperty('WEB_URL') || '';
    const santos = filas_('SantoSena').filter(r => r.fecha >= semana && r.fecha <= finS);
    return {
      semana: semana,
      responsables: filas_('Responsables').filter(r => r.semana === semana).map(limpiar_),
      santoSena: santos.map(limpiar_),
      comandantes: filas_('Comandantes').map(limpiar_),
      rondas: filas_('Rondas').filter(r => r.fecha >= semana && r.fecha <= finS).reverse().map(r => {
        const ss = santos.find(x => x.fecha === r.fecha) || { santo: '', sena: '' };
        return { id: r.id, fecha: r.fecha, grado_nombre: r.grado_nombre, estado: r.estado, enviado: r.enviado,
                 completado: r.completado, tienePdf: !!r.pdf_id,
                 enlaceWa: enlaceWa_(r.celular, mensajeRonda_(r, ss, url)) };
      })
    };
  },
  guardarSemana: req => {
    admin_(req);
    const semana = lunes_(req.semana);
    if ((req.responsables || []).length > 2) throw new Error('Solo se permiten dos responsables por semana.');
    reemplazar_('Responsables', r => r.semana === semana,
      req.responsables.map(p => ({ semana: semana, grado_nombre: p.grado_nombre, celular: celular_(p.celular), pin: p.pin })));
    const fin = new Date(semana + 'T12:00:00'); fin.setDate(fin.getDate() + 6);
    const finS = Utilities.formatDate(fin, TZ, 'yyyy-MM-dd');
    reemplazar_('SantoSena', r => r.fecha >= semana && r.fecha <= finS,
      req.santoSena.filter(x => x.fecha >= semana && x.fecha <= finS));
    return true;
  },
  guardarComandantes: req => {
    admin_(req);
    reemplazar_('Comandantes', () => true, req.comandantes.map(c => ({ cargo: c.cargo, grado_nombre: c.grado_nombre, celular: celular_(c.celular), clave_callmebot: c.clave_callmebot || '' })));
    return true;
  },
  generarHoy: req => { admin_(req); envioDiario(); return true; },
  historial: req => {
    admin_(req);
    return filas_('Rondas').filter(r => r.estado === 'COMPLETADO' && (!req.desde || r.fecha >= req.desde) && (!req.hasta || r.fecha <= req.hasta))
      .reverse().map(r => ({ id: r.id, fecha: r.fecha, grado_nombre: r.grado_nombre, completado: r.completado }));
  },
  pdf: req => {
    admin_(req);
    const r = filas_('Rondas').find(x => x.id === req.id);
    if (!r || !r.pdf_id) throw new Error('PDF no encontrado');
    const f = DriveApp.getFileById(r.pdf_id);
    return { nombre: f.getName(), base64: Utilities.base64Encode(f.getBlob().getBytes()) };
  }
};

function limpiar_(r) { const o = Object.assign({}, r); delete o._fila; return o; }

function admin_(req) {
  if (!req.sesion || !CacheService.getScriptCache().get('adm_' + req.sesion)) throw new Error('SESION_VENCIDA');
}

function rondaPorToken_(token, pin) {
  const r = filas_('Rondas').find(x => x.token && x.token === token);
  if (!r) throw new Error('Enlace no válido.');
  const p = filas_('Responsables').find(x => x.semana === lunes_(r.fecha) && celular_(x.celular) === r.celular);
  if (p && p.pin && String(p.pin) !== String(pin)) { Utilities.sleep(1500); throw new Error('PIN incorrecto.'); }
  const limite = new Date(r.fecha + 'T12:00:00'); limite.setDate(limite.getDate() + 1);
  if (hoy_() > Utilities.formatDate(limite, TZ, 'yyyy-MM-dd')) throw new Error('Este enlace ya venció.');
  return r;
}

function validarEnvio_(d) {
  const total = LISTA.reduce((s, sec) => s + sec.items.length, 0);
  const marcados = Object.keys(d.respuestas || {}).filter(k => d.respuestas[k] === 'SI' || d.respuestas[k] === 'NO');
  if (marcados.length !== total) throw new Error('Debe marcar SI o NO en todos los puntos.');
  if (!d.fotos || d.fotos.length !== 4) throw new Error('Debe adjuntar las 4 fotografías.');
  if (!d.firma) throw new Error('Falta la firma del Jefe de Ronda.');
  const img = /^data:image\/(jpeg|png);base64,[A-Za-z0-9+\/=]+$/;
  if (!img.test(d.firma) || d.fotos.some(f => !img.test(f))) throw new Error('Imagen no válida.');
}

/* ---------- PDF ---------- */

const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

function generarPdf_(r, d, carpeta) {
  const t = HtmlService.createTemplateFromFile('Plantilla');
  const p = r.fecha.split('-');
  t.dia = String(Number(p[2]));
  t.mes = MESES[Number(p[1]) - 1];
  t.anio = p[0];
  t.lista = LISTA;
  t.d = d;
  t.r = r;
  const html = t.evaluate().getContent();
  const blob = Utilities.newBlob(html, 'text/html', 'lista.html').getAs('application/pdf')
    .setName('Lista Jefe de Ronda ' + r.fecha + ' - ' + r.grado_nombre + '.pdf');
  return carpeta.createFile(blob);
}

/** Prueba: genera un PDF de ejemplo en la carpeta de Drive para revisar el formato. */
function probarPdf() {
  const resp = {};
  LISTA.forEach((s, i) => s.items.forEach((_, j) => resp[i + '-' + j] = j % 3 ? 'SI' : 'NO'));
  const pix = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
  const d = { respuestas: resp, obs: { '0-0': 'Sin novedad' }, novedades: 'Ninguna.', capitan: 'Cap. Ejemplo', cmdteGuardia: 'Sbtte. Ejemplo',
              oficialPatrullas: 'Tte. Ejemplo', fotos: [pix, pix, pix, pix], leyendas: ['Prevención', 'Polvorines', 'Cocinas', 'Talleres'], firma: pix };
  const f = generarPdf_({ fecha: hoy_(), grado_nombre: 'Tte. Prueba', completado: ahora_() }, d, carpeta_('pruebas'));
  Logger.log(f.getUrl());
}
