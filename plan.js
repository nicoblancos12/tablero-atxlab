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
  if (!d) return '';
  const f = (d instanceof Date) ? d : new Date(d);
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
  const n = parseFloat(v);
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

return { leer, clasificar, bloqueos, iso, masDias, dias, iconoPara };
})();
