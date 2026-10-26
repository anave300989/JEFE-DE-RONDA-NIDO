/**
 * Sistema de Control del Jefe de Ronda - ESCONBOL
 * Backend en Google Apps Script (base de datos en Google Sheets, archivos en Google Drive).
 *
 * Hojas que usa (se crean solas con instalar()):
 *   Responsables : fecha, grado_nombre, celular, pin   (dos por día)
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
  Responsables: ['fecha', 'grado_nombre', 'celular', 'pin'],
  SantoSena: ['fecha', 'santo', 'sena'],
  Comandantes: ['cargo', 'grado_nombre', 'celular', 'clave_callmebot'],
  Rondas: ['id', 'fecha', 'grado_nombre', 'celular', 'token', 'estado', 'enviado', 'completado', 'pdf_id', 'fotos', 'novedades'],
  GuardiaRol: ['fecha', 'cargo', 'grado_nombre', 'celular', 'pin', 'token'],
  GuardiaFotos: ['id', 'fecha', 'cargo', 'turno', 'lugar', 'hora', 'recibida', 'archivo_id', 'grado_nombre'],
  EnlacesComando: ['semana', 'token', 'pin']
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(n);
  if (!sh) {
    sh = ss.insertSheet(n);
    sh.getRange('A:Z').setNumberFormat('@');
  }
  const cols = HOJAS[n].length;
  const enc = sh.getRange(1, 1, 1, cols).getDisplayValues()[0];
  if (enc.join() !== HOJAS[n].join()) sh.getRange(1, 1, 1, cols).setValues([HOJAS[n]]);
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

/** Crea las rondas del día para los dos responsables del día y envía los mensajes. */
function envioDiario() {
  const fecha = hoy_();
  const resp = filas_('Responsables').filter(r => r.fecha === fecha);
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
      responsables: filas_('Responsables').filter(r => r.fecha >= semana && r.fecha <= finS).map(limpiar_),
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
    const fin = new Date(semana + 'T12:00:00'); fin.setDate(fin.getDate() + 6);
    const finS = Utilities.formatDate(fin, TZ, 'yyyy-MM-dd');
    const resp = (req.responsables || []).filter(p => p.fecha >= semana && p.fecha <= finS);
    const porDia = {};
    resp.forEach(p => porDia[p.fecha] = (porDia[p.fecha] || 0) + 1);
    if (Object.keys(porDia).some(f => porDia[f] > 2)) throw new Error('Solo se permiten dos responsables por día.');
    reemplazar_('Responsables', r => r.fecha >= semana && r.fecha <= finS,
      resp.map(p => ({ fecha: p.fecha, grado_nombre: p.grado_nombre, celular: celular_(p.celular), pin: p.pin })));
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
  const p = filas_('Responsables').find(x => x.fecha === r.fecha && celular_(x.celular) === r.celular);
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

/* ---------- Guardia: Capitán de Servicio y Oficial de Patrulla ---------- */

const CARGOS_GUARDIA = {
  CAPITAN: { nombre: 'CAPITÁN DE SERVICIO', lugares: ['Polvorín', 'Material Bélico', 'Comando', 'Plana Mayor', 'Perezosos', 'Transportes'] },
  PATRULLA: { nombre: 'OFICIAL DE PATRULLA', lugares: ['Fuerte 1', 'Fuerte 2', 'Campo Salado', 'Confital', 'Granja 1', 'Barrio Chino'] }
};

// Turno de 24 h: de 07:00 a 07:00 del día siguiente.
const TURNOS = [
  { id: 'MANANA', nombre: 'Mañana', horario: '07:00 a 13:00' },
  { id: 'TARDE', nombre: 'Tarde', horario: '13:00 a 19:00' },
  { id: 'NOCHE', nombre: 'Noche', horario: '19:00 a 01:00' },
  { id: 'MADRUGADA', nombre: 'Madrugada', horario: '01:00 a 07:00' }
];

/** Fecha de guardia y turno que corresponden a un instante (en hora de Bolivia). */
function turnoDe_(ms) {
  const d = new Date(ms);
  const h = Number(Utilities.formatDate(d, TZ, 'H'));
  const fecha = Utilities.formatDate(h < 7 ? new Date(ms - 12 * 3600 * 1000) : d, TZ, 'yyyy-MM-dd');
  const turno = h >= 7 && h < 13 ? 'MANANA' : h >= 13 && h < 19 ? 'TARDE' : (h >= 19 || h < 1) ? 'NOCHE' : 'MADRUGADA';
  return { fecha: fecha, turno: turno };
}

function asignacionGuardia_(token, pin) {
  const a = filas_('GuardiaRol').find(x => x.token && x.token === token);
  if (!a) throw new Error('Enlace no válido.');
  if (a.pin && String(a.pin) !== String(pin)) { Utilities.sleep(1500); throw new Error('PIN incorrecto.'); }
  return a;
}

function conteoGuardia_(fecha, cargo) {
  const c = {};
  filas_('GuardiaFotos').filter(f => f.fecha === fecha && f.cargo === cargo)
    .forEach(f => { c[f.turno] = c[f.turno] || {}; c[f.turno][f.lugar] = (c[f.turno][f.lugar] || 0) + 1; });
  return c;
}

function mensajeGuardia_(a, url) {
  return '*CONTROL DE GUARDIA - ESCONBOL*\n' + a.grado_nombre + ', usted está designado como ' + CARGOS_GUARDIA[a.cargo].nombre +
    ' para la guardia del ' + a.fecha + ' (07:00 a 07:00).\n\nRegistre las fotografías de cada turno en:\n' + url + 'guardia.html?t=' + a.token;
}

function mensajeComando_(e, url) {
  return '*CONTROL DE GUARDIA Y RONDA - ESCONBOL*\nSemana del ' + e.semana + '.\nConsulte las fotografías de guardia y el registro de rondas en:\n' +
    url + 'comando.html?t=' + e.token;
}

function finSemana_(semana) {
  const d = new Date(semana + 'T12:00:00'); d.setDate(d.getDate() + 6);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function carpetaGuardia_(fecha) {
  const g = carpeta_('Guardia');
  const it = g.getFoldersByName(fecha);
  return it.hasNext() ? it.next() : g.createFolder(fecha);
}

Object.assign(ACCIONES, {
  /* Capitán de Servicio / Oficial de Patrulla */
  guardiaAbrir: req => {
    const a = asignacionGuardia_(req.token, req.pin);
    const actual = turnoDe_(Date.now());
    return { fecha: a.fecha, cargo: CARGOS_GUARDIA[a.cargo].nombre, grado_nombre: a.grado_nombre,
             lugares: CARGOS_GUARDIA[a.cargo].lugares, turnos: TURNOS,
             turnoActual: actual.fecha === a.fecha ? actual.turno : null,
             terminada: actual.fecha > a.fecha, ahora: Date.now(), conteo: conteoGuardia_(a.fecha, a.cargo) };
  },
  guardiaFoto: req => {
    const a = asignacionGuardia_(req.token, req.pin);
    const ahora = Date.now(), tomada = Number(req.tomada) || ahora;
    if (tomada > ahora + 2 * 60 * 1000 || tomada < ahora - 6 * 3600 * 1000) throw new Error('La hora de la foto no es válida.');
    const t = turnoDe_(tomada);
    if (t.fecha !== a.fecha || t.turno !== req.turno) throw new Error('La foto no corresponde a este turno.');
    if (CARGOS_GUARDIA[a.cargo].lugares.indexOf(req.lugar) < 0) throw new Error('Lugar no válido.');
    if (!/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(req.foto || '')) throw new Error('Imagen no válida.');
    const hora = Utilities.formatDate(new Date(tomada), TZ, 'yyyy-MM-dd HH:mm:ss');
    const turnoNombre = TURNOS.find(x => x.id === t.turno).nombre;
    const nombre = a.fecha + ' ' + turnoNombre + ' - ' + req.lugar + ' - ' + hora.slice(11).replace(/:/g, '.') + '.jpg';
    const archivo = carpetaGuardia_(a.fecha).createFile(Utilities.newBlob(Utilities.base64Decode(req.foto.split(',')[1]), 'image/jpeg', nombre));
    hoja_('GuardiaFotos').appendRow([token_(), a.fecha, a.cargo, t.turno, req.lugar, hora,
      Utilities.formatDate(new Date(ahora), TZ, 'yyyy-MM-dd HH:mm:ss'), archivo.getId(), a.grado_nombre]);
    return { hora: hora, conteo: conteoGuardia_(a.fecha, a.cargo) };
  },

  /* Administrador */
  guardiaPanel: req => {
    admin_(req);
    const semana = lunes_(req.fecha || hoy_()), finS = finSemana_(semana);
    const url = PropertiesService.getScriptProperties().getProperty('WEB_URL') || '';
    const e = filas_('EnlacesComando').find(x => x.semana === semana);
    const cmd = filas_('Comandantes').filter(c => c.celular);
    return {
      semana: semana,
      rol: filas_('GuardiaRol').filter(r => r.fecha >= semana && r.fecha <= finS).map(r => ({
        fecha: r.fecha, cargo: r.cargo, grado_nombre: r.grado_nombre, celular: r.celular, pin: r.pin,
        enlaceWa: r.token ? enlaceWa_(r.celular, mensajeGuardia_(r, url)) : '' })),
      comando: e ? { pin: e.pin, url: url + 'comando.html?t=' + e.token,
        avisos: cmd.map(c => ({ cargo: c.cargo, enlace: enlaceWa_(celular_(c.celular), mensajeComando_(e, url)) })) } : null
    };
  },
  guardarGuardia: req => {
    admin_(req);
    const semana = lunes_(req.semana), finS = finSemana_(semana);
    const previas = filas_('GuardiaRol').filter(r => r.fecha >= semana && r.fecha <= finS);
    const rol = (req.rol || []).filter(r => r.fecha >= semana && r.fecha <= finS && CARGOS_GUARDIA[r.cargo] && r.grado_nombre);
    reemplazar_('GuardiaRol', r => r.fecha >= semana && r.fecha <= finS, rol.map(r => {
      const cel = celular_(r.celular);
      const prev = previas.find(x => x.fecha === r.fecha && x.cargo === r.cargo && x.celular === cel);
      return { fecha: r.fecha, cargo: r.cargo, grado_nombre: r.grado_nombre, celular: cel, pin: r.pin, token: prev ? prev.token : token_() };
    }));
    return true;
  },
  crearEnlaceComando: req => {
    admin_(req);
    const semana = lunes_(req.semana);
    if (!/^\d{4,6}$/.test(String(req.pin || ''))) throw new Error('El PIN de los comandantes debe tener de 4 a 6 números.');
    const prev = filas_('EnlacesComando').find(x => x.semana === semana);
    reemplazar_('EnlacesComando', r => r.semana === semana, [{ semana: semana, token: prev ? prev.token : token_(), pin: req.pin }]);
    return true;
  },

  /* Cmdte y 2do Cmdte (solo lectura) */
  comandoAbrir: req => {
    const e = comandoEnlace_(req.token, req.pin);
    const semana = e.semana, finS = finSemana_(semana), hoy = hoy_();
    const rondas = filas_('Rondas'), resp = filas_('Responsables'), rol = filas_('GuardiaRol');
    const fotos = filas_('GuardiaFotos').filter(f => f.fecha >= semana && f.fecha <= finS);
    const dias = [];
    for (let i = 0; i < 7; i++) {
      const d = new Date(semana + 'T12:00:00'); d.setDate(d.getDate() + i);
      const fecha = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
      dias.push({
        fecha: fecha,
        rondas: resp.filter(p => p.fecha === fecha).map(p => {
          const r = rondas.find(x => x.fecha === fecha && x.celular === celular_(p.celular));
          return { grado_nombre: p.grado_nombre, completado: r && r.estado === 'COMPLETADO' ? r.completado : '',
                   estado: r && r.estado === 'COMPLETADO' ? 'PASÓ RONDA' : (fecha >= hoy ? 'PENDIENTE' : 'NO PASÓ RONDA') };
        }),
        guardia: Object.keys(CARGOS_GUARDIA).map(c => {
          const a = rol.find(x => x.fecha === fecha && x.cargo === c);
          return { cargo: c, cargoNombre: CARGOS_GUARDIA[c].nombre, grado_nombre: a ? a.grado_nombre : '', lugares: CARGOS_GUARDIA[c].lugares,
                   fotos: fotos.filter(f => f.fecha === fecha && f.cargo === c).map(f => ({ id: f.id, turno: f.turno, lugar: f.lugar, hora: f.hora, recibida: f.recibida })) };
        })
      });
    }
    return { semana: semana, fin: finS, turnos: TURNOS, dias: dias };
  },
  comandoFoto: req => {
    const e = comandoEnlace_(req.token, req.pin);
    const f = filas_('GuardiaFotos').find(x => x.id === req.id);
    if (!f || f.fecha < e.semana || f.fecha > finSemana_(e.semana)) throw new Error('Foto no encontrada.');
    return { base64: Utilities.base64Encode(DriveApp.getFileById(f.archivo_id).getBlob().getBytes()) };
  }
});

function comandoEnlace_(token, pin) {
  const e = filas_('EnlacesComando').find(x => x.token && x.token === token);
  if (!e) throw new Error('Enlace no válido.');
  if (String(e.pin) !== String(pin)) { Utilities.sleep(1500); throw new Error('PIN incorrecto.'); }
  return e;
}
