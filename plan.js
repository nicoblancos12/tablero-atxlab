/* ══════════════════════════════════════════════════════════════════════
   plan.js — lector del plan de trabajo de ATX (export de MS Project)
   Requiere SheetJS (XLSX) como global.

   Estructura que espera en la hoja "Plan de trabajo":
     A1  CLIENTE | Nombre del proyecto
     A3  Inicio: ... | Fin: ... | N horas
     A5  ID | WBS | Nombre de tarea | % Avance | Horas | Duración |
         Inicio | Fin | Predecesoras | Grupo de recurso
   Si el archivo cambia de forma, el parser avisa en vez de inventar.
   ══════════════════════════════════════════════════════════════════════ */
window.PlanATX = (function () {

const iso = (d) => {
  if (d == null || d === '') return '';
  if (d instanceof Date) {
    if (isNaN(d)) return '';
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  // Un número suelto es el serial de Excel (días desde 1899-12-30).
  if (typeof d === 'number' && isFinite(d) && d > 1 && d < 80000) {
    const base = new Date(1899, 11, 30);
    base.setDate(base.getDate() + Math.round(d));
    return iso(base);
  }
  const t = String(d).trim();
  if (!t) return '';
  // Texto: se arma a mano. new Date('2026-11-30') lo toma como UTC y en
  // zonas al oeste devuelve el día anterior — así se corrían TODAS las
  // fechas de los planes que traen la fecha como texto.
  let m = t.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
  // dd/mm/aaaa (y dd-mm-aaaa). Con día <= 12 es ambiguo: en los planes de
  // ATX el orden es día primero, así que se respeta ese.
  m = t.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
  if (m) {
    const anio = m[3].length === 2 ? '20' + m[3] : m[3];
    return anio + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  }
  const f = new Date(t);
  if (isNaN(f)) return '';
  return f.getFullYear() + '-' + String(f.getMonth() + 1).padStart(2, '0') + '-' + String(f.getDate()).padStart(2, '0');
};
const masDias = (isoStr, n) => {
  const f = new Date(isoStr + 'T12:00:00');
  f.setDate(f.getDate() + n);
  return iso(f);
};
const dias = (a, b) => Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 86400000);

/** El % de avance llega como 0-100 o como fracción 0-1 según el export. */
function normAvance(v) {
  // Algunos planes ponen el estado con palabras en vez de un porcentaje.
  if (typeof v === 'string') {
    const t = v.toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (/^(terminad|completad|cerrad|listo|hecho|done|complete|finalizad)/.test(t)) return 100;
    if (/^(en progreso|en curso|iniciad|empezad|avanzand|wip|in progress)/.test(t)) return 50;
    if (/^(pendiente|no iniciad|sin empezar|por hacer|todo|not started|n\/?a)/.test(t)) return 0;
  }
  const n = parseFloat(String(v == null ? '' : v).replace('%', '').replace(',', '.'));
  if (!isFinite(n) || n <= 0) return 0;
  return Math.min(100, Math.round(n <= 1 ? n * 100 : n));
}

function leer(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const hoja = wb.SheetNames.find((n) => /plan/i.test(n)) || wb.SheetNames[0];
  // raw:true + cellDates:true — con raw:false las fechas vuelven como texto
  // formateado y se malinterpretan (7/27/26 se leía como 10 de enero).
  const filas = XLSX.utils.sheet_to_json(wb.Sheets[hoja], { header: 1, raw: true, cellDates: true, defval: null });

  // Encabezado del proyecto
  // El encabezado viene como "CLIENTE | Proyecto" o como "Cliente · Proyecto"
  const a1 = String((filas[0] || [])[0] || '');
  const sep = a1.indexOf('|') >= 0 ? '|' : (a1.indexOf('·') >= 0 ? '·' : null);
  const trozos = sep ? a1.split(sep).map((s) => s.trim()) : [a1.trim()];
  const cliente = trozos[0] || '';
  const nombre = trozos.slice(1).join(' · ') || '';

  // Horas totales y fechas del bloque de resumen
  const a3 = String((filas[2] || [])[0] || '');
  const mHoras = a3.match(/([\d,]+)\s*horas/i);
  const horasEncabezado = mHoras ? parseInt(mHoras[1].replace(/,/g, ''), 10) : null;

  // Fila de encabezados de la tabla
  let h = filas.findIndex((f) => f && String(f[0]).trim() === 'ID' && /WBS/i.test(String(f[1] || '')));
  if (h < 0) throw new Error('No encontré la fila de encabezados (ID | WBS | Nombre de tarea). ¿Es el formato de plan de ATX?');

  const crudas = [];
  for (let i = h + 1; i < filas.length; i++) {
    const f = filas[i];
    if (!f) continue;
    const id = f[0], wbs = f[1] == null ? '' : String(f[1]).trim(), nom = (f[2] == null ? '' : String(f[2])).trim();
    if (!nom || /^TOTAL/i.test(nom)) continue;
    if (id == null || wbs === '') continue;
    crudas.push({
      id: String(id).trim(),
      wbs,
      nivel: wbs.split('.').length,
      nombre: nom,
      avance: normAvance(f[3]),
      horas: parseFloat(f[4]) || 0,
      ini: iso(f[6]),
      fin: iso(f[7]),
      preds: String(f[8] || '').split(',').map((s) => s.trim()).filter(Boolean),
      grupo: String(f[9] || '').trim(),
    });
  }
  if (!crudas.length) throw new Error('El archivo no trae tareas debajo de los encabezados.');

  // Códigos de recurso presentes (para preguntar cuál es el cliente)
  const codigos = new Set();
  crudas.forEach((t) => t.grupo.split(',').forEach((g) => { if (g.trim()) codigos.add(g.trim()); }));

  const fechas = crudas.map((t) => t.ini).filter(Boolean).sort();
  const finales = crudas.map((t) => t.fin).filter(Boolean).sort();

  return {
    cliente, nombre, horasEncabezado,
    inicioPlan: fechas[0] || '', finPlan: finales[finales.length - 1] || '',
    codigos: Array.from(codigos).sort(),
    tareas: crudas,
  };
}

/**
 * Clasifica el plan leído. Devuelve fases, actividades, requisitos e hitos,
 * con las fechas ya recorridas al kickoff real.
 *   codigoCliente: p.ej. 'ARA' — las tareas cuyo único recurso es ese código
 *                  se vuelven requisitos del cliente, no actividades nuestras.
 *   kickoffReal:   fecha en la que de verdad arrancó el proyecto.
 */
function clasificar(plan, { codigoCliente, kickoffReal }) {
  const desfase = (kickoffReal && plan.inicioPlan) ? dias(plan.inicioPlan, kickoffReal) : 0;
  const mover = (d) => (d ? masDias(d, desfase) : '');

  // La columna "Grupo de recurso" se usa de dos maneras según quién arme el
  // plan: unos ponen códigos (ATX, ARA) y otros ponen nombres de personas
  // ("Ana Carol · Rodrigo"). Soportamos las dos: lo que no sea un código se
  // toma como dueño de la tarea.
  const partes = (g) => String(g || '').split(/[,·;/]| y /).map((s) => s.trim()).filter(Boolean);
  const esCodigo = (g) => g.length <= 5 && g === g.toUpperCase();

  const esCliente = (t) => {
    const gs = partes(t.grupo);
    return !!codigoCliente && gs.length > 0 && gs.every((g) => g === codigoCliente);
  };
  const duenoDe = (t) => partes(t.grupo)
    .filter((g) => g !== codigoCliente && !esCodigo(g))
    .join(', ');
  const esEntregable = (t) => /^entregable\s*:/i.test(t.nombre);

  // Fases = nivel 1. Cada tarea cuelga de la última fase vista.
  const fases = [];
  const porFase = {};
  let faseActual = null;
  const items = [];

  plan.tareas.forEach((t) => {
    if (t.nivel === 1) {
      faseActual = { clave: 'f' + t.wbs, nombre: t.nombre, icono: iconoPara(t.nombre), wbs: t.wbs };
      fases.push(faseActual);
      porFase[t.id] = faseActual.clave;
      return;
    }
    if (!faseActual) return;
    const tipo = esEntregable(t) ? 'hito' : (esCliente(t) ? 'requisito' : 'actividad');
    items.push({
      idPlan: t.id, wbs: t.wbs, nombre: t.nombre.replace(/^entregable\s*:\s*/i, ''),
      frente: faseActual.clave, horas: t.horas, avance: t.avance,
      ini: mover(t.ini), fin: mover(t.fin),
      preds: t.preds, grupo: t.grupo, dueno: duenoDe(t), tipo,
    });
  });

  return { fases: fases.filter((f) => items.some((i) => i.frente === f.clave)), items, desfase };
}

/** Iconos por nombre de fase — puro apoyo visual, editable después. */
function iconoPara(nombre) {
  const n = nombre.toLowerCase();
  if (/previa|arranque|kick/.test(n)) return '🚀';
  if (/descubr|análisis|analisis|diseño|diseno/.test(n)) return '🔍';
  if (/piloto|whatsapp|canal/.test(n)) return '💬';
  if (/dato|conocimiento|analít|analit|report/.test(n)) return '📊';
  if (/nube|azure|infra|plataforma/.test(n)) return '☁️';
  if (/motor|modelo|ia|agente/.test(n)) return '🧠';
  if (/cierre|capacit|document/.test(n)) return '🏁';
  return '⚙️';
}

/**
 * Construye los bloqueos: un requisito bloquea a toda actividad que lo
 * declare como predecesora. Es la liga que el plan ya trae y que de otro
 * modo habría que capturar a mano.
 */
function bloqueos(items) {
  const porIdPlan = {};
  items.forEach((i) => { porIdPlan[i.idPlan] = i; });
  const mapa = {};
  items.filter((i) => i.tipo === 'requisito').forEach((r) => { mapa[r.idPlan] = []; });
  items.filter((i) => i.tipo === 'actividad').forEach((a) => {
    a.preds.forEach((p) => { if (mapa[p]) mapa[p].push(a.idPlan); });
  });
  return mapa;
}


/* ══════════════════════════════════════════════════════════════════════
   LECTURA FLEXIBLE — para planes que no vienen en la plantilla de ATX
   El lector estricto de arriba exige la forma exacta del export de MS
   Project. Cuando un plan viene de otra fuente (otro cliente, otra
   herramienta, un Excel hecho a mano) se rechazaba entero. Aquí la
   columna se busca por lo que DICE su encabezado y por lo que TRAE
   adentro, no por su posición.

   Nunca adivina en silencio: devuelve qué entendió y con qué confianza,
   para que la persona lo confirme antes de crear nada.
   ══════════════════════════════════════════════════════════════════════ */

/** Sinónimos por campo. El orden importa: gana el de mayor puntaje. */
const SINONIMOS = {
  nombre:  ['nombre de tarea', 'nombre de la tarea', 'tarea / actividad', 'tarea', 'actividad',
            'actividades', 'descripcion', 'descripción', 'concepto', 'entregable', 'tema',
            'task name', 'task', 'nombre', 'detalle', 'trabajo a realizar', 'pendiente',
            'pendientes', 'que hacer', 'qué hacer', 'to do', 'todo', 'item', 'asunto'],
  wbs:     ['wbs', 'edt', 'eps', 'clave', 'codigo', 'código', 'nivel', 'no.', 'num', 'número'],
  id:      ['id', 'identificador', 'folio', '#'],
  avance:  ['% avance', '%avance', 'avance', '% completado', 'completado', 'progreso',
            'porcentaje', '% complete', 'estatus', 'estado'],
  horas:   ['horas', 'hrs', 'hr', 'esfuerzo', 'trabajo', 'work', 'tiempo est.', 'tiempo estimado',
            'tiempo', 'dias', 'días', 'duracion', 'duración', 'duration', 'jornadas'],
  ini:     ['inicio', 'fecha inicio', 'fecha de inicio', 'comienzo', 'desde', 'start',
            'start date', 'arranque', 'inicia'],
  fin:     ['fin', 'termino', 'término', 'fecha fin', 'fecha de fin', 'fecha final', 'hasta',
            'finish', 'end', 'vencimiento', 'vence', 'entrega', 'fecha compromiso',
            'fecha limite', 'fecha límite', 'deadline', 'due', 'due date', 'cierre', 'fecha'],
  dueno:   ['grupo de recurso', 'recurso', 'recursos', 'responsable', 'responsables', 'dueño',
            'dueno', 'asignado a', 'asignado', 'owner', 'quien', 'quién', 'encargado', 'equipo',
            'assigned to', 'assignee', 'persona', 'a cargo', 'ejecuta'],
  preds:   ['predecesoras', 'predecesora', 'depende de', 'dependencias', 'predecessors', 'bloquea'],
  agrupa:  ['semana', 'sprint', 'modulo', 'módulo', 'fase', 'etapa', 'track', 'bloque', 'mes'],
};

const norm = (s) => String(s == null ? '' : s).toLowerCase().trim()
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ').replace(/[:*]+$/, '');

/** ¿Qué campo nombra este encabezado? Devuelve [campo, puntaje]. */
function campoDeEncabezado(txt) {
  const t = norm(txt);
  if (!t) return null;
  let mejor = null;
  Object.keys(SINONIMOS).forEach((campo) => {
    SINONIMOS[campo].forEach((syn) => {
      const s = norm(syn);
      let p = 0;
      if (t === s) p = 100;
      else if (t.indexOf(s) === 0) p = 80 - (t.length - s.length);
      else if (t.indexOf(s) > 0) p = 55 - (t.length - s.length);
      if (p > 0 && (!mejor || p > mejor[1])) mejor = [campo, p];
    });
  });
  return mejor;
}

const esFechaVal = (v) => (v instanceof Date && !isNaN(v)) ||
  (typeof v === 'string' && /^\s*\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(v)) ||
  // Serial de Excel: un entero en el rango de fechas razonables (1990-2070).
  (typeof v === 'number' && Number.isInteger(v) && v >= 32874 && v <= 62000);
const esNumVal = (v) => typeof v === 'number' ||
  (typeof v === 'string' && /^\s*[~<>]?\s*\d+([.,]\d+)?\s*(h|hr|hrs|horas|d|dias|días)?\s*$/i.test(v));

/**
 * Horas escritas a mano: "2-3 h", "~1.5 h", "1 h 45", "3 días".
 * Un rango toma el valor mayor (planear por el peor caso).
 * Los días se convierten con jornada de 8 h.
 */
/** Tope de cordura: más de 2000 h (un año-persona) en UNA tarea es un error
    de captura, no una estimación. Mejor 0 que descuadrar todo el avance. */
const HORAS_MAX = 2000;
function acotarHoras(n) {
  if (!isFinite(n) || n <= 0) return 0;
  return n > HORAS_MAX ? 0 : Math.round(n * 100) / 100;
}

function normHoras(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return acotarHoras(v);
  let t = String(v).toLowerCase().replace(/[~≈]/g, '').trim();
  if (!t || t === '-' || t === '—') return 0;
  const esDias = /d[ií]as?\b|\bd\b|jornada/.test(t);
  const hm = t.match(/(\d+)\s*h\w*\s*(\d{1,2})\s*(?:m|min)?\b/);
  if (hm) return acotarHoras(+hm[1] + (+hm[2]) / 60);
  // "45 min" son 0.75 h, no 45; sin esto una tarea corta pesa como una semana.
  const mm = t.match(/^(\d+(?:[.,]\d+)?)\s*(?:m|min|mins|minutos)\b/);
  if (mm) return acotarHoras(parseFloat(mm[1].replace(',', '.')) / 60);
  const nums = (t.match(/\d+(?:[.,]\d+)?/g) || []).map((n) => parseFloat(n.replace(',', '.')));
  if (!nums.length) return 0;
  const val = Math.max.apply(null, nums);
  return acotarHoras(esDias ? val * 8 : val);
}

/** Cuánta sangría trae el texto (para inferir jerarquía sin WBS). */
const sangria = (s) => {
  const m = String(s == null ? '' : s).match(/^[\s \t]*/);
  return m ? m[0].replace(/\t/g, '    ').length : 0;
};

/**
 * Encuentra la fila de encabezados: la que reconoce más campos distintos,
 * exigiendo al menos un nombre de actividad. Se revisan las primeras filas
 * porque arriba suele haber título, cliente y fechas del proyecto.
 */
function hallarEncabezados(filas, maxBuscar) {
  let mejor = null;
  const tope = Math.min(filas.length, maxBuscar || 60);
  for (let i = 0; i < tope; i++) {
    const f = filas[i];
    if (!f) continue;
    const mapa = {}, puntos = {};
    let distintos = 0, suma = 0;
    f.forEach((celda, col) => {
      const r = campoDeEncabezado(celda);
      if (!r) return;
      const [campo, p] = r;
      if (puntos[campo] == null || p > puntos[campo]) {
        if (puntos[campo] == null) distintos++;
        puntos[campo] = p; mapa[campo] = col;
      }
    });
    if (mapa.nombre == null) continue;          // sin actividad no hay plan
                                                // (== null, no !: la columna 0 es válida)
    // Un encabezado de verdad tiene filas de datos abajo. Sin esto, una
    // celda suelta que casualmente dice "Pendiente:" o "Concepto" se toma
    // por tabla y el archivo entero se lee mal.
    const abajo = contarDatos(filas, i, mapa.nombre);
    if (!abajo) continue;
    // Con varios campos reconocidos (Actividad + Responsable + Fin) una sola
    // fila ya es un plan válido. Con uno solo la señal es débil y hace falta
    // ver varias filas, que es la lista de pendientes a secas.
    if (distintos < 2 && abajo < 2) continue;

    suma = Object.keys(puntos).reduce((s, k) => s + puntos[k], 0);
    const score = distintos * 1000 + suma + Math.min(abajo, 50);
    if (!mejor || score > mejor.score) mejor = { fila: i, mapa, puntos, distintos, score, abajo };
  }
  return mejor;
}

/** ¿La hoja es una tabla de datos (muchos números en varias columnas)? */
function pareceReporte(filas, ancho) {
  let colsNum = 0;
  for (let c = 0; c < Math.min(ancho, 12); c++) {
    const vals = filas.map((f) => (f ? f[c] : null)).filter((v) => v != null && v !== '');
    if (vals.length < 5) continue;
    const nums = vals.filter((v) => typeof v === 'number').length;
    if (nums >= vals.length * 0.7) colsNum++;
  }
  return colsNum >= 2;
}

/** Cuántas filas con texto hay bajo un encabezado, parando en el primer hueco grande. */
function contarDatos(filas, hFila, colNombre) {
  let n = 0, huecos = 0;
  for (let i = hFila + 1; i < filas.length && huecos < 3; i++) {
    const f = filas[i];
    const v = f ? f[colNombre] : null;
    const t = String(v == null ? '' : v).trim();
    if (!t) { huecos++; continue; }
    huecos = 0;
    if (/^(total|subtotal|suma)\b/i.test(t)) break;
    n++;
  }
  return n;
}

/**
 * Verifica contra el contenido: si la columna marcada como fecha no trae
 * fechas, se descarta. Evita el error silencioso de leer la columna que no es.
 */
function verificarMapa(filas, hFila, mapa) {
  const muestra = filas.slice(hFila + 1, hFila + 31).filter((f) => f && f.some((c) => c != null && c !== ''));
  const avisos = [];
  const prueba = (campo, test, minimo) => {
    if (mapa[campo] == null) return;
    const vals = muestra.map((f) => f[mapa[campo]]).filter((v) => v != null && v !== '' && v !== '—' && v !== '-');
    if (!vals.length) { avisos.push({ campo, motivo: 'la columna viene vacía' }); delete mapa[campo]; return; }
    const ok = vals.filter(test).length / vals.length;
    if (ok < minimo) { avisos.push({ campo, motivo: 'el contenido no parece ' + campo }); delete mapa[campo]; }
  };
  prueba('ini', esFechaVal, 0.6);
  prueba('fin', esFechaVal, 0.6);
  prueba('horas', esNumVal, 0.5);
  return { avisos, llenas: contarLlenas(filas, hFila, mapa, muestra.length) };
}

function contarLlenas(filas, hFila, mapa, total) {
  const muestra = filas.slice(hFila + 1).filter((f) => f && f.some((c) => c != null && c !== ''));
  const out = {};
  Object.keys(mapa).forEach((campo) => {
    const c = mapa[campo];
    out[campo] = muestra.filter((f) => f[c] != null && f[c] !== '' && f[c] !== '—').length;
  });
  out.__filas = muestra.length;
  return out;
}

/**
 * Lee cualquier hoja con el mapa dado. Si no hay WBS, arma la jerarquía
 * con la sangría del nombre; si tampoco hay, las filas sin horas que
 * encabezan un bloque se toman como fase.
 */
function leerConMapa(filas, hFila, mapa, opciones) {
  const op = opciones || {};
  const col = (f, campo) => (mapa[campo] == null ? null : f[mapa[campo]]);
  const crudas = [];
  const niveles = [];

  for (let i = hFila + 1; i < filas.length; i++) {
    const f = filas[i];
    if (!f) continue;
    const nomCrudo = col(f, 'nombre');
    const nom = String(nomCrudo == null ? '' : nomCrudo).trim();
    if (!nom || /^(total|subtotal|suma)\b/i.test(nom)) continue;
    // Los exports paginados repiten la fila de encabezados cada N filas.
    if (encabezadoRepetido(f, mapa)) continue;

    const wbs = mapa.wbs != null ? String(col(f, 'wbs') == null ? '' : col(f, 'wbs')).trim() : '';
    const horas = normHoras(col(f, 'horas'));
    const t = {
      id: String((mapa.id != null ? col(f, 'id') : null) || wbs || (crudas.length + 1)).trim(),
      wbs,
      nombre: nom,
      sangria: sangria(nomCrudo),
      avance: normAvance(col(f, 'avance')),
      horas,
      ini: iso(col(f, 'ini')),
      fin: iso(col(f, 'fin')),
      preds: String(col(f, 'preds') || '').split(/[,;]/).map((s) => s.trim()).filter(Boolean),
      grupo: String(col(f, 'dueno') == null ? '' : col(f, 'dueno')).replace(/^[—-]$/, '').trim(),
    };
    crudas.push(t);
    niveles.push(t);
  }

  // Nivel: por WBS si lo hay; si no, por sangría; si no, por "fila sin horas".
  const conWbs = crudas.filter((t) => t.wbs).length;
  if (conWbs > crudas.length * 0.6) {
    crudas.forEach((t) => { t.nivel = t.wbs ? t.wbs.split('.').length : 2; });
  } else {
    const sangrias = Array.from(new Set(crudas.map((t) => t.sangria))).sort((a, b) => a - b);
    if (sangrias.length > 1) {
      crudas.forEach((t) => { t.nivel = sangrias.indexOf(t.sangria) + 1; });
    } else if (op.agrupa && op.agrupa.length === crudas.length) {
      // Una columna repetida (Semana, Sprint, Módulo, Track) dice a qué
      // bloque pertenece cada fila: es una jerarquía real que el archivo ya
      // trae, no una inventada. Se vuelve fase y se inserta como encabezado.
      const vistos = {};
      const conFases = [];
      crudas.forEach((t, i) => {
        const g = String(op.agrupa[i] == null ? '' : op.agrupa[i]).trim();
        if (g && !vistos[g]) {
          vistos[g] = true;
          conFases.push({ id: 'g' + conFases.length, wbs: '', nivel: 1,
                          nombre: (op.etiquetaAgrupa ? op.etiquetaAgrupa + ' ' : '') + g,
                          sangria: 0, avance: 0, horas: 0, ini: '', fin: '', preds: [], grupo: '' });
        }
        t.nivel = 2;
        conFases.push(t);
      });
      crudas.length = 0;
      conFases.forEach((t) => crudas.push(t));
    } else {
      // Sin WBS, sin sangría y sin agrupador no hay jerarquía que recuperar.
      // Antes se tomaba "fila sin horas" como fase, pero eso ascendía
      // renglones sueltos (un feriado, una nota) a fase. Es preferible una
      // sola fase que una jerarquía inventada.
      crudas.forEach((t) => { t.nivel = 2; });
    }
  }
  // Si quedó todo en un nivel, se inventa una fase única para no perder nada.
  if (!crudas.some((t) => t.nivel === 1)) {
    crudas.unshift({ id: 'f0', wbs: '1', nivel: 1, nombre: op.faseUnica || 'Actividades',
                     sangria: 0, avance: 0, horas: 0, ini: '', fin: '', preds: [], grupo: '' });
    crudas.forEach((t, i) => { if (i) t.nivel = Math.max(2, t.nivel); });
  }

  const codigos = new Set();
  crudas.forEach((t) => t.grupo.split(/[,·;/]| y /).forEach((g) => { if (g.trim()) codigos.add(g.trim()); }));
  const fechas = crudas.map((t) => t.ini).filter(Boolean).sort();
  const finales = crudas.map((t) => t.fin).filter(Boolean).sort();
  return {
    inicioPlan: fechas[0] || '', finPlan: finales[finales.length - 1] || '',
    codigos: Array.from(codigos).sort(), tareas: crudas,
  };
}

/** Título y cliente de las filas de arriba, cuando existen. */
function cabecera(filas, hFila) {
  let cliente = '', nombre = '';
  for (let i = 0; i < Math.min(hFila, 6); i++) {
    const f = filas[i]; if (!f) continue;
    const txt = f.map((c) => (c == null ? '' : String(c).trim())).filter(Boolean).join(' | ');
    if (!txt || /^(inicio|fin)\b/i.test(txt)) continue;
    const sep = txt.indexOf('|') >= 0 ? '|' : (txt.indexOf('·') >= 0 ? '·' : null);
    const tr = sep ? txt.split(sep).map((s) => s.trim()) : [txt];
    if (!cliente) { cliente = tr[0] || ''; nombre = tr.slice(1).join(' · '); }
    if (cliente && nombre) break;
  }
  return { cliente, nombre };
}

/**
 * Entrada principal del modo flexible. Revisa TODAS las hojas y se queda
 * con la que mejor se entiende; devuelve el plan y el informe de lectura
 * para que la persona lo confirme o corrija.
 */
function leerFlexible(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const candidatas = [];

  wb.SheetNames.forEach((nombreHoja) => {
    const filas = XLSX.utils.sheet_to_json(wb.Sheets[nombreHoja],
      { header: 1, raw: true, cellDates: true, defval: null, blankrows: true });
    if (!filas.length) return;
    const h = hallarEncabezados(filas);
    if (!h) return;
    const mapa = Object.assign({}, h.mapa);
    const chequeo = verificarMapa(filas, h.fila, mapa);
    const plan = leerConMapa(filas, h.fila, mapa, {
      faseUnica: nombreHoja,
      agrupa: columnaDe(filas, h.fila, mapa, 'agrupa'),
      etiquetaAgrupa: etiquetaDe(filas, h.fila, mapa.agrupa),
    });
    if (!plan.tareas.length) return;
    candidatas.push({
      hoja: nombreHoja, filas, filaEnc: h.fila, mapa, chequeo, plan,
      // Preferimos la hoja con más campos reconocidos y más tareas.
      score: h.distintos * 1000 + Math.min(plan.tareas.length, 200) +
             (/plan|cronog|activ|tarea/i.test(nombreHoja) ? 500 : 0),
    });
  });

  // Último intento: una lista a secas, sin encabezado reconocible. Si la
  // primera columna trae puro texto en varias filas, se toma como la lista
  // de actividades. Es preferible leerla y enseñar lo entendido que rechazar.
  if (!candidatas.length) {
    const rescate = rescatarLista(wb);
    if (rescate) candidatas.push(rescate);
  }

  if (!candidatas.length) {
    throw new Error('No encontré ninguna tabla de actividades. Revisa que el archivo tenga una fila de ' +
      'encabezados con algo como "Actividad", "Tarea" o "Descripción", o una columna con la lista de pendientes.');
  }
  candidatas.sort((a, b) => b.score - a.score);
  const g = candidatas[0];
  const cab = cabecera(g.filas, g.filaEnc);

  return {
    cliente: cab.cliente, nombre: cab.nombre, horasEncabezado: null,
    inicioPlan: g.plan.inicioPlan, finPlan: g.plan.finPlan,
    codigos: g.plan.codigos, tareas: g.plan.tareas,
    lectura: {
      hoja: g.hoja, hojas: wb.SheetNames, filaEncabezados: g.filaEnc + 1,
      mapa: g.mapa, encabezados: (g.filas[g.filaEnc] || []).map((c) => (c == null ? '' : String(c))),
      avisos: g.chequeo.avisos, llenas: g.chequeo.llenas, columnas: anchoDe(g.filas),
      candidatas: candidatas.map((c) => ({ hoja: c.hoja, tareas: c.plan.tareas.length })),
    },
  };
}

/**
 * Rescate: ninguna hoja tenía encabezados reconocibles. Si alguna trae una
 * columna con varias celdas de texto (y no números, que serían un reporte),
 * se lee como lista simple de actividades.
 */
function rescatarLista(wb) {
  let mejor = null;
  wb.SheetNames.forEach((nombreHoja) => {
    const filas = XLSX.utils.sheet_to_json(wb.Sheets[nombreHoja],
      { header: 1, raw: true, cellDates: true, defval: null, blankrows: true });
    const ancho = anchoDe(filas);
    // Un reporte (ventas, inventario) tiene varias columnas de números en
    // paralelo. Un plan no. Sin este corte, cualquier reporte se leería
    // como si su primera columna de texto fuera una lista de tareas.
    if (pareceReporte(filas, ancho)) return;
    for (let c = 0; c < Math.min(ancho, 4); c++) {
      const vals = filas.map((f) => (f ? f[c] : null))
        .filter((v) => v != null && String(v).trim() !== '');
      if (vals.length < 3) continue;
      const textos = vals.filter((v) => typeof v === 'string' && String(v).trim().length > 3 &&
                                        !/^\d+([.,]\d+)?$/.test(String(v).trim()));
      // Casi todo texto y suficientemente largo: parece lista de tareas.
      // Exigir varias filas de texto evita tomar por lista los títulos
      // sueltos de un reporte (un nombre de planta, un encabezado de sección).
      if (textos.length < vals.length * 0.8 || textos.length < 4) continue;
      // Una tarea se escribe como frase. Las etiquetas de un formulario
      // ("FECHA DE VENTA:") y los encabezados en mayúsculas no lo son: si
      // dominan, esto es la carátula de un reporte, no una lista de trabajo.
      const frases = textos.filter((v) => {
        const t = String(v).trim();
        if (/:$/.test(t)) return false;
        if (t === t.toUpperCase() && /[A-ZÁÉÍÓÚÑ]/.test(t)) return false;
        return /\s/.test(t);
      });
      if (frases.length < textos.length * 0.6 || frases.length < 3) continue;
      const score = frases.length;
      if (!mejor || score > mejor.score) {
        // La primera celda se toma como encabezado de la lista.
        const mapa = { nombre: c };
        const plan = leerConMapa(filas, 0, mapa, { faseUnica: nombreHoja });
        // Menos de 3 tareas no es un plan: es texto suelto malinterpretado.
        if (plan.tareas.filter((t) => t.nivel !== 1).length < 3) continue;
        mejor = {
          hoja: nombreHoja, filas, filaEnc: 0, mapa, score,
          chequeo: { avisos: [{ campo: 'nombre', motivo: 'no había encabezados: leí la columna como lista' }],
                     llenas: contarLlenas(filas, 0, mapa) },
          plan,
        };
      }
    }
  });
  return mejor;
}

/** ¿Esta fila es otra copia de los encabezados? (exports paginados) */
function encabezadoRepetido(f, mapa) {
  let iguales = 0, campos = 0;
  Object.keys(mapa).forEach((campo) => {
    const v = f[mapa[campo]];
    if (v == null || v === '') return;
    campos++;
    const r = campoDeEncabezado(v);
    if (r && r[0] === campo) iguales++;
  });
  return campos >= 2 && iguales === campos;
}

const anchoDe = (filas) => filas.reduce((m, f) => Math.max(m, f ? f.length : 0), 0);

/** Valores de una columna, alineados 1 a 1 con las filas que sí traen tarea. */
function columnaDe(filas, hFila, mapa, campo) {
  if (mapa[campo] == null || mapa.nombre == null) return null;
  const out = [];
  for (let i = hFila + 1; i < filas.length; i++) {
    const f = filas[i];
    if (!f) continue;
    const nom = String(f[mapa.nombre] == null ? '' : f[mapa.nombre]).trim();
    if (!nom || /^(total|subtotal|suma)\b/i.test(nom)) continue;
    out.push(f[mapa[campo]]);
  }
  const distintos = new Set(out.map((v) => String(v == null ? '' : v).trim()).filter(Boolean));
  // Ni una sola categoría ni una por fila: eso no agrupa nada.
  if (distintos.size < 2 || distintos.size > out.length * 0.6) return null;
  return out;
}

const etiquetaDe = (filas, hFila, col) => {
  if (col == null) return '';
  const v = (filas[hFila] || [])[col];
  return v == null ? '' : String(v).trim();
};

/** Relectura con el mapa corregido a mano desde el panel de confirmación. */
function releerCon(buffer, hoja, filaEnc, mapa) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const filas = XLSX.utils.sheet_to_json(wb.Sheets[hoja],
    { header: 1, raw: true, cellDates: true, defval: null, blankrows: true });
  const limpio = {};
  Object.keys(mapa).forEach((k) => { if (mapa[k] != null && mapa[k] !== '') limpio[k] = +mapa[k]; });
  const plan = leerConMapa(filas, filaEnc - 1, limpio, { faseUnica: hoja });
  const cab = cabecera(filas, filaEnc - 1);
  return {
    cliente: cab.cliente, nombre: cab.nombre, horasEncabezado: null,
    inicioPlan: plan.inicioPlan, finPlan: plan.finPlan,
    codigos: plan.codigos, tareas: plan.tareas,
    lectura: { hoja, hojas: wb.SheetNames, filaEncabezados: filaEnc, mapa: limpio,
               encabezados: (filas[filaEnc - 1] || []).map((c) => (c == null ? '' : String(c))),
               avisos: [], llenas: contarLlenas(filas, filaEnc - 1, limpio), columnas: anchoDe(filas),
               candidatas: [] },
  };
}

/**
 * Puerta de entrada: intenta la plantilla exacta y, si falla, el modo
 * flexible. Así un plan con el formato de siempre sigue el camino probado.
 */
function leerAuto(buffer) {
  try {
    const p = leer(buffer);
    p.lectura = { estricto: true, hoja: 'Plan de trabajo', avisos: [] };
    return p;
  } catch (e) {
    const p = leerFlexible(buffer);
    p.lectura.estricto = false;
    p.lectura.porQue = e.message;
    return p;
  }
}

return { leer, leerAuto, leerFlexible, releerCon, clasificar, bloqueos, iso, masDias, dias, iconoPara,
         SINONIMOS, normHoras, hallarEncabezados };
})();
