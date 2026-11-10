/**
 * DESAFÍO 360 - Servidor en Google Apps Script
 * Fuente única de verdad: el Google Sheet (hojas Config, Equipos, Log).
 * Todas las llamadas son GET con parámetros (evita problemas de CORS desde GitHub Pages).
 */

/* ============ CONFIGURACIÓN (lo único que debes editar) ============ */
const ADMIN_PASS = 'cambia-esta-clave';   // <-- CAMBIA ESTO por tu contraseña de admin
/* =================================================================== */

const HOJA_CFG = 'Config';
const HOJA_EQ = 'Equipos';
const HOJA_LOG = 'Log';
const N_BLOQUES = 6;
const MAX_PTS = 100;       // máximo por bloque (6 x 100 = 600)
const CLASIFICAN = 6;      // equipos que clasifican
const NOMBRES_ORIGINALES = [   // nombres de fábrica: se restauran al reiniciar el juego (por id 1..10)
  'Dream Team', 'Mineroraptors de la PGN', 'Sirius', 'Cracks gpt', 'Como tú',
  'El combo del pechi', 'Los últimos del examen, primeros en confianza',
  'La Nasa nos busca pero no contestamos', 'Obsidiana', '😎Sin Talento, pero con Actitud 😎'
];
const HB_VIVO_MS = 20000;  // un equipo se ve "conectado" si hubo señal en los últimos 20 s

// Columnas (1-based) de la hoja Equipos
const C = { id: 1, nombre: 2, ocupado: 3, token: 4, bloque: 5, pts: 6, seg: 12,
            total: 18, tiempo: 19, terminado: 20, fin: 21, ult: 22, clasif: 23,
            expulsado: 24, mov: 25 };   // 24 y 25 se crean solas si no existen
const N_COLS = 25;
const CLAVES_CFG = ['estado', 't0_ms', 'pausa_inicio_ms', 'pausa_acumulada_ms', 'version'];

/* ----------------------------- Entrada ----------------------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    out = manejar_(p);
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function manejar_(p) {
  const accion = p.accion || 'estado';
  switch (accion) {
    case 'estado':    return estado_(p);
    case 'elegir':    return conLock_(() => elegir_(p));
    case 'puntaje':   return conLock_(() => puntaje_(p));
    // ---- acciones de administrador ----
    case 'iniciar':   return admin_(p, iniciar_);
    case 'pausar':    return admin_(p, pausar_);
    case 'reanudar':  return admin_(p, reanudar_);
    case 'finalizar': return admin_(p, finalizar_);
    case 'reiniciar': return admin_(p, reiniciar_);
    case 'renombrar': return admin_(p, renombrar_);
    case 'liberar':   return admin_(p, liberar_);
    case 'corregir':  return admin_(p, corregir_);
    case 'clasificar':return admin_(p, clasificar_);
    case 'expulsar':  return admin_(p, expulsar_);
    case 'readmitir': return admin_(p, readmitir_);
    case 'mover':     return admin_(p, mover_);
    default: throw new Error('Acción desconocida: ' + accion);
  }
}

function admin_(p, fn) {
  if (!p.pass || p.pass !== ADMIN_PASS) throw new Error('Contraseña de administrador incorrecta');
  return conLock_(() => fn(p));
}

function conLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ----------------------------- Lectura ----------------------------- */
function hoja_(nombre) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!sh) throw new Error('Falta la hoja "' + nombre + '"');
  return sh;
}

function leerCfg_() {
  const vals = hoja_(HOJA_CFG).getRange(2, 1, CLAVES_CFG.length, 2).getValues();
  const cfg = {};
  vals.forEach(r => { cfg[String(r[0]).trim()] = r[1]; });
  return {
    estado: cfg.estado || 'sin_iniciar',
    t0: cfg.t0_ms === '' || cfg.t0_ms == null ? null : Number(cfg.t0_ms),
    pausaInicio: cfg.pausa_inicio_ms === '' || cfg.pausa_inicio_ms == null ? null : Number(cfg.pausa_inicio_ms),
    pausaAcum: Number(cfg.pausa_acumulada_ms) || 0,
    version: Number(cfg.version) || 0
  };
}

function escribirCfg_(cambios) {
  const sh = hoja_(HOJA_CFG);
  const claves = sh.getRange(2, 1, CLAVES_CFG.length, 1).getValues().map(r => String(r[0]).trim());
  Object.keys(cambios).forEach(k => {
    const i = claves.indexOf(k);
    if (i < 0) throw new Error('Falta la clave "' + k + '" en Config');
    sh.getRange(2 + i, 2).setValue(cambios[k] === null ? '' : cambios[k]);
  });
}

function bool_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function numONull_(v) { return v === '' || v == null ? null : Number(v); }

let colsOk_ = false;
function asegurarColumnas_(sh) {
  if (colsOk_) return;
  const cache = CacheService.getScriptCache();
  if (!cache.get('cols_ok')) {
    const max = sh.getMaxColumns();
    if (max < N_COLS) sh.insertColumnsAfter(max, N_COLS - max);
    if (!sh.getRange(1, C.expulsado).getValue()) sh.getRange(1, C.expulsado).setValue('expulsado');
    if (!sh.getRange(1, C.mov).getValue()) sh.getRange(1, C.mov).setValue('mov');
    cache.put('cols_ok', '1', 21600);
  }
  colsOk_ = true;
}

function leerEquipos_() {
  const sh = hoja_(HOJA_EQ);
  asegurarColumnas_(sh);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, N_COLS).getValues();
  return vals.map((r, i) => {
    const pts = [], seg = [];
    for (let b = 0; b < N_BLOQUES; b++) {
      pts.push(numONull_(r[C.pts - 1 + b]));
      seg.push(numONull_(r[C.seg - 1 + b]));
    }
    return {
      fila: i + 2,
      id: Number(r[C.id - 1]),
      nombre: String(r[C.nombre - 1]),
      ocupado: bool_(r[C.ocupado - 1]),
      token: String(r[C.token - 1] || ''),
      bloque: Number(r[C.bloque - 1]) || 0,
      pts: pts,
      seg: seg,
      total: pts.reduce((a, v) => a + (v || 0), 0),
      tiempoTotal: numONull_(r[C.tiempo - 1]),
      terminado: bool_(r[C.terminado - 1]),
      finMs: numONull_(r[C.fin - 1]),
      clasificado: bool_(r[C.clasif - 1]),
      expulsado: bool_(r[C.expulsado - 1]),
      mov: Number(r[C.mov - 1]) || 0
    };
  });
}

function buscar_(equipos, id) {
  const eq = equipos.filter(e => e.id === Number(id))[0];
  if (!eq) throw new Error('Equipo inexistente');
  return eq;
}

/** Siguiente bloque pendiente (1..N) después de 'actual'; si no hay adelante, el primero pendiente; N+1 si ya hizo todos. */
function siguientePendiente_(pts, actual) {
  for (let i = actual; i < N_BLOQUES; i++) if (pts[i] === null) return i + 1;
  for (let i = 0; i < N_BLOQUES; i++) if (pts[i] === null) return i + 1;
  return N_BLOQUES + 1;
}

function efectivoMs_(cfg, ahora) {
  if (cfg.estado === 'sin_iniciar' || cfg.t0 == null) return 0;
  if (cfg.estado === 'pausada' || cfg.estado === 'finalizada') {
    return (cfg.pausaInicio != null ? cfg.pausaInicio : ahora) - cfg.t0 - cfg.pausaAcum;
  }
  return ahora - cfg.t0 - cfg.pausaAcum;
}

/* ----------------------------- Estado público ----------------------------- */
function estado_(p) {
  const ahora = Date.now();
  const cache = CacheService.getScriptCache();
  let base = cache.get('base');
  if (base) {
    base = JSON.parse(base);
  } else {
    const cfg = leerCfg_();
    const equipos = leerEquipos_();
    base = { cfg: cfg, equipos: equipos };
    cache.put('base', JSON.stringify(base), 3);
  }
  const cfg = base.cfg, equipos = base.equipos;

  // latido del equipo que consulta (solo si el token es válido)
  let miEquipo = null;
  if (p.token) {
    const mio = equipos.filter(e => e.ocupado && e.token === p.token)[0];
    if (mio) {
      miEquipo = mio.id;
      cache.put('hb_' + mio.id, String(ahora), 120);
    }
  }
  const expulsado = !!p.token && equipos.some(e => e.expulsado && e.token === p.token);
  const hbs = cache.getAll(equipos.map(e => 'hb_' + e.id));

  const salida = equipos.map(e => ({
    id: e.id,
    nombre: e.nombre,
    ocupado: e.ocupado,
    bloque: e.bloque,
    pts: e.pts,
    seg: e.seg,
    total: e.total,
    tiempoTotal: e.tiempoTotal,
    terminado: e.terminado,
    finMs: e.finMs,
    clasificado: e.clasificado,
    expulsado: e.expulsado,
    mov: e.mov,
    conectado: e.ocupado && hbs['hb_' + e.id] ? (ahora - Number(hbs['hb_' + e.id]) < HB_VIVO_MS) : false
  }));

  // Ranking: puntaje (desc), luego menor tiempo total; los que no terminaron van después en empate
  const ranking = salida.filter(e => e.ocupado).sort((a, b) => {
    if (b.total !== a.total) return b.total - a.total;
    const ta = a.terminado && a.tiempoTotal != null ? a.tiempoTotal : Infinity;
    const tb = b.terminado && b.tiempoTotal != null ? b.tiempoTotal : Infinity;
    if (ta !== tb) return ta - tb;
    return a.id - b.id;
  }).map(e => e.id);

  // Orden de llegada a la meta (primeros en terminar los 6 bloques)
  const llegada = salida.filter(e => e.terminado && e.finMs != null)
    .sort((a, b) => a.finMs - b.finMs).map(e => e.id);

  return {
    ok: true,
    serverNow: ahora,
    estado: cfg.estado,
    t0: cfg.t0,
    version: cfg.version,
    tiempoEfectivoMs: efectivoMs_(cfg, ahora),
    miEquipo: miEquipo,
    expulsado: expulsado,
    ranking: ranking,
    llegada: llegada,
    equipos: salida
  };
}

/* ----------------------------- Acciones de jugador ----------------------------- */
function elegir_(p) {
  const token = String(p.token || '');
  if (token.length < 8) throw new Error('Token inválido');
  const cfg = leerCfg_();
  if (cfg.estado === 'finalizada') throw new Error('La competencia ya finalizó');
  const equipos = leerEquipos_();
  const eq = buscar_(equipos, p.id);

  if (equipos.some(e => e.expulsado && e.token === token)) throw new Error('Tu dispositivo fue retirado de la competencia');
  if (eq.expulsado) throw new Error('Ese grupo fue retirado de la competencia');
  if (eq.ocupado) {
    if (eq.token === token) return { ok: true, reconectado: true, id: eq.id };
    throw new Error('Ese grupo ya fue elegido por otro dispositivo');
  }
  if (equipos.some(e => e.ocupado && e.token === token)) {
    throw new Error('Este dispositivo ya eligió un grupo');
  }
  const sh = hoja_(HOJA_EQ);
  sh.getRange(eq.fila, C.ocupado).setValue(true);
  sh.getRange(eq.fila, C.token).setValue(token);
  // si el grupo ya tenía avance (p. ej. el admin lo liberó), se conserva su bloque
  if (!(eq.bloque > 0)) sh.getRange(eq.fila, C.bloque).setValue(cfg.estado === 'sin_iniciar' ? 0 : 1);
  log_(eq.id, 'elegir', eq.nombre);
  return invalidar_({ ok: true, id: eq.id });
}

function puntaje_(p) {
  const cfg = leerCfg_();
  if (cfg.estado !== 'en_curso') throw new Error('La competencia no está en curso');
  const equipos = leerEquipos_();
  const eq = buscar_(equipos, p.id);
  if (!eq.ocupado || eq.token !== String(p.token || '')) throw new Error('Token no válido para este grupo');

  const b = parseInt(p.bloque, 10);
  const pts = Math.round(Number(p.pts));
  const seg = Math.round(Number(p.seg));
  if (!(b >= 1 && b <= N_BLOQUES)) throw new Error('Bloque inválido');
  if (!(pts >= 0 && pts <= MAX_PTS)) throw new Error('Puntaje fuera de rango (0 a ' + MAX_PTS + ')');
  if (!(seg >= 0)) throw new Error('Tiempo inválido');

  if (eq.pts[b - 1] !== null) throw new Error('Ese bloque ya fue registrado');
  if (b !== eq.bloque) throw new Error('El grupo debe jugar el bloque ' + eq.bloque + ' (no el ' + b + ')');

  const ahora = Date.now();
  const efectivoSeg = efectivoMs_(cfg, ahora) / 1000;
  if (seg > efectivoSeg + 5) throw new Error('El tiempo reportado supera el tiempo transcurrido');

  const sh = hoja_(HOJA_EQ);
  sh.getRange(eq.fila, C.pts + b - 1).setValue(pts);
  sh.getRange(eq.fila, C.seg + b - 1).setValue(seg);
  const nuevosPts = eq.pts.slice(); nuevosPts[b - 1] = pts;
  const sig = siguientePendiente_(nuevosPts, b);
  sh.getRange(eq.fila, C.bloque).setValue(sig);

  let terminado = false;
  if (sig > N_BLOQUES) {
    terminado = true;
    sh.getRange(eq.fila, C.terminado).setValue(true);
    sh.getRange(eq.fila, C.fin).setValue(ahora);
    sh.getRange(eq.fila, C.tiempo).setValue(Math.round(efectivoSeg));
  }
  log_(eq.id, 'puntaje', 'bloque ' + b + ', ' + pts + ' pts, ' + seg + ' s');
  return invalidar_({ ok: true, bloque: b, pts: pts, terminado: terminado });
}

/* ----------------------------- Acciones de admin ----------------------------- */
function iniciar_() {
  const cfg = leerCfg_();
  if (cfg.estado !== 'sin_iniciar') throw new Error('La competencia ya fue iniciada');
  const ahora = Date.now();
  escribirCfg_({ estado: 'en_curso', t0_ms: ahora, pausa_inicio_ms: null, pausa_acumulada_ms: 0 });
  // todos los equipos que ya eligieron pasan al bloque 1
  const sh = hoja_(HOJA_EQ);
  leerEquipos_().forEach(e => { if (e.ocupado) sh.getRange(e.fila, C.bloque).setValue(1); });
  log_('', 'iniciar', 't0=' + ahora);
  return invalidar_({ ok: true, t0: ahora });
}

function pausar_() {
  const cfg = leerCfg_();
  if (cfg.estado !== 'en_curso') throw new Error('Solo se puede pausar una competencia en curso');
  const ahora = Date.now();
  escribirCfg_({ estado: 'pausada', pausa_inicio_ms: ahora });
  log_('', 'pausar', '');
  return invalidar_({ ok: true });
}

function reanudar_() {
  const cfg = leerCfg_();
  if (cfg.estado !== 'pausada') throw new Error('La competencia no está en pausa');
  const ahora = Date.now();
  const acum = cfg.pausaAcum + (ahora - (cfg.pausaInicio || ahora));
  escribirCfg_({ estado: 'en_curso', pausa_inicio_ms: null, pausa_acumulada_ms: acum });
  log_('', 'reanudar', '');
  return invalidar_({ ok: true });
}

function finalizar_() {
  const cfg = leerCfg_();
  if (cfg.estado !== 'en_curso' && cfg.estado !== 'pausada') throw new Error('Nada que finalizar');
  const ahora = Date.now();
  escribirCfg_({ estado: 'finalizada', pausa_inicio_ms: cfg.estado === 'pausada' ? cfg.pausaInicio : ahora });
  log_('', 'finalizar', '');
  return invalidar_({ ok: true });
}

function reiniciar_() {
  escribirCfg_({ estado: 'sin_iniciar', t0_ms: null, pausa_inicio_ms: null, pausa_acumulada_ms: 0 });
  const sh = hoja_(HOJA_EQ);
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    sh.getRange(2, C.ocupado, n, 1).setValue(false);
    sh.getRange(2, C.token, n, 1).clearContent();
    sh.getRange(2, C.bloque, n, 1).setValue(0);
    sh.getRange(2, C.pts, n, N_BLOQUES * 2).clearContent();   // pts_b1..b6 y seg_b1..b6
    sh.getRange(2, C.tiempo, n, 1).clearContent();
    sh.getRange(2, C.terminado, n, 1).setValue(false);
    sh.getRange(2, C.fin, n, 2).clearContent();               // fin_ms y ultima_conexion_ms
    sh.getRange(2, C.clasif, n, 1).setValue(false);
    sh.getRange(2, C.expulsado, n, 1).setValue(false);
    sh.getRange(2, C.mov, n, 1).setValue(0);
  }
  const cache = CacheService.getScriptCache();
  const equipos = leerEquipos_();
  equipos.forEach(e => {
    cache.remove('hb_' + e.id);
    // restaurar el nombre original (por si el admin lo renombró)
    if (NOMBRES_ORIGINALES[e.id - 1]) sh.getRange(e.fila, C.nombre).setValue(NOMBRES_ORIGINALES[e.id - 1]);
  });
  // dejar el registro (Log) en blanco; queda solo el encabezado
  const shLog = hoja_(HOJA_LOG);
  const nLog = shLog.getLastRow() - 1;
  if (nLog > 0) shLog.getRange(2, 1, nLog, 4).clearContent();
  log_('', 'reiniciar', 'juego reiniciado desde el panel');
  return invalidar_({ ok: true });
}

function renombrar_(p) {
  const nombre = String(p.nombre || '').trim();
  if (!nombre) throw new Error('Nombre vacío');
  if (nombre.length > 60) throw new Error('Nombre demasiado largo (máx. 60)');
  const eq = buscar_(leerEquipos_(), p.id);
  hoja_(HOJA_EQ).getRange(eq.fila, C.nombre).setValue(nombre);
  log_(eq.id, 'renombrar', eq.nombre + ' -> ' + nombre);
  return invalidar_({ ok: true });
}

function liberar_(p) {
  const eq = buscar_(leerEquipos_(), p.id);
  const sh = hoja_(HOJA_EQ);
  sh.getRange(eq.fila, C.ocupado).setValue(false);
  sh.getRange(eq.fila, C.token).clearContent();
  CacheService.getScriptCache().remove('hb_' + eq.id);
  log_(eq.id, 'liberar', eq.nombre);
  return invalidar_({ ok: true });
}

function corregir_(p) {
  const cfg = leerCfg_();
  const eq = buscar_(leerEquipos_(), p.id);
  const b = parseInt(p.bloque, 10);
  const pts = Math.round(Number(p.pts));
  const seg = Math.round(Number(p.seg));
  if (!(b >= 1 && b <= N_BLOQUES)) throw new Error('Bloque inválido');
  if (!(pts >= 0 && pts <= MAX_PTS)) throw new Error('Puntaje fuera de rango');
  if (!(seg >= 0)) throw new Error('Tiempo inválido');

  const sh = hoja_(HOJA_EQ);
  sh.getRange(eq.fila, C.pts + b - 1).setValue(pts);
  sh.getRange(eq.fila, C.seg + b - 1).setValue(seg);

  // recalcular avance: bloques consecutivos completados
  const nuevos = eq.pts.slice(); nuevos[b - 1] = pts;
  const sig = siguientePendiente_(nuevos, 0);
  sh.getRange(eq.fila, C.bloque).setValue(sig);
  if (sig > N_BLOQUES && !eq.terminado) {
    const ahora = Date.now();
    sh.getRange(eq.fila, C.terminado).setValue(true);
    sh.getRange(eq.fila, C.fin).setValue(ahora);
    sh.getRange(eq.fila, C.tiempo).setValue(Math.round(efectivoMs_(cfg, ahora) / 1000));
  }
  log_(eq.id, 'corregir', 'bloque ' + b + ', ' + pts + ' pts, ' + seg + ' s');
  return invalidar_({ ok: true });
}

/** Saca al grupo de la competencia: su dispositivo queda fuera, no cuenta en el ranking y no se puede volver a elegir. */
function expulsar_(p) {
  const eq = buscar_(leerEquipos_(), p.id);
  if (eq.expulsado) throw new Error('Ese grupo ya fue expulsado');
  if (!eq.ocupado) throw new Error('Ese grupo aún no ha sido elegido por ningún dispositivo');
  const sh = hoja_(HOJA_EQ);
  sh.getRange(eq.fila, C.ocupado).setValue(false);
  sh.getRange(eq.fila, C.expulsado).setValue(true);   // el token se conserva para reconocer al dispositivo expulsado
  sh.getRange(eq.fila, C.clasif).setValue(false);
  CacheService.getScriptCache().remove('hb_' + eq.id);
  log_(eq.id, 'expulsar', eq.nombre);
  return invalidar_({ ok: true });
}

/** Revierte una expulsión: el grupo queda libre (conserva su avance) para ser elegido otra vez. */
function readmitir_(p) {
  const eq = buscar_(leerEquipos_(), p.id);
  if (!eq.expulsado) throw new Error('Ese grupo no está expulsado');
  const sh = hoja_(HOJA_EQ);
  sh.getRange(eq.fila, C.expulsado).setValue(false);
  sh.getRange(eq.fila, C.token).clearContent();
  log_(eq.id, 'readmitir', eq.nombre);
  return invalidar_({ ok: true });
}

/** Envía al grupo a un reto: si ya lo había hecho, se borra ese resultado para que lo repita. */
function mover_(p) {
  const cfg = leerCfg_();
  if (cfg.estado !== 'en_curso' && cfg.estado !== 'pausada') throw new Error('Solo se puede mover un grupo con la competencia en curso o en pausa');
  const eq = buscar_(leerEquipos_(), p.id);
  if (!eq.ocupado) throw new Error('El grupo no está activo');
  const b = parseInt(p.bloque, 10);
  if (!(b >= 1 && b <= N_BLOQUES)) throw new Error('Bloque inválido');
  if (eq.bloque === b && eq.pts[b - 1] === null) throw new Error('El grupo ya está en ese reto');

  const sh = hoja_(HOJA_EQ);
  const repite = eq.pts[b - 1] !== null;
  if (repite) {
    sh.getRange(eq.fila, C.pts + b - 1).clearContent();
    sh.getRange(eq.fila, C.seg + b - 1).clearContent();
  }
  if (eq.terminado) {                     // si ya había terminado, deja de estarlo
    sh.getRange(eq.fila, C.terminado).setValue(false);
    sh.getRange(eq.fila, C.fin).clearContent();
    sh.getRange(eq.fila, C.tiempo).clearContent();
  }
  sh.getRange(eq.fila, C.bloque).setValue(b);
  sh.getRange(eq.fila, C.mov).setValue(eq.mov + 1);   // avisa al celular del grupo para que se recoloque
  log_(eq.id, 'mover', (repite ? 'repetir' : 'ir a') + ' bloque ' + b);
  return invalidar_({ ok: true });
}

function clasificar_() {
  const equipos = leerEquipos_().filter(e => e.ocupado).sort((a, b) => {
    if (b.total !== a.total) return b.total - a.total;
    const ta = a.terminado && a.tiempoTotal != null ? a.tiempoTotal : Infinity;
    const tb = b.terminado && b.tiempoTotal != null ? b.tiempoTotal : Infinity;
    if (ta !== tb) return ta - tb;
    return a.id - b.id;
  });
  const ganadores = equipos.slice(0, CLASIFICAN).map(e => e.id);
  const sh = hoja_(HOJA_EQ);
  leerEquipos_().forEach(e => sh.getRange(e.fila, C.clasif).setValue(ganadores.indexOf(e.id) >= 0));
  log_('', 'clasificar', 'ids: ' + ganadores.join(','));
  return invalidar_({ ok: true, clasificados: ganadores });
}

/* ----------------------------- Utilidades ----------------------------- */
function log_(equipoId, accion, detalle) {
  hoja_(HOJA_LOG).appendRow([Date.now(), equipoId, accion, detalle]);
}

function invalidar_(resp) {
  const cfg = leerCfg_();
  escribirCfg_({ version: cfg.version + 1 });
  CacheService.getScriptCache().remove('base');
  return resp;
}

/** Ejecuta esta función una vez desde el editor para autorizar permisos y verificar las hojas. */
function verificar() {
  const cfg = leerCfg_();
  const eq = leerEquipos_();
  Logger.log('Estado: ' + cfg.estado + ' | Equipos: ' + eq.length);
}
