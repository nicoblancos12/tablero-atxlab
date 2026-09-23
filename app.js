/* ══════════════════════════════════════════════════════════════════════
   Tablero de proyectos atxlab
   Capa de datos: una lista de SharePoint con 3 columnas (Tipo, Ref, Datos).
   Un ítem por registro, para que dos personas editando no se pisen.
   ══════════════════════════════════════════════════════════════════════ */

const CONFIG = {
  backend: 'sharepoint',       // 'local' (pruebas) | 'sharepoint' (compartido)
  clientId: 'e9fdfbd4-c437-466c-a4d7-057c9b2b79e7',
  tenantId: '3a5ef2b1-86c6-41bc-a30c-94709a8d0590',
  siteHost: 'atx1.sharepoint.com',
  sitePath: '/sites/atxlabProyectos',
  listName: 'Tablero atxlab',
  // SharePoint reserva el nombre "Tipo" para una columna del sistema, así que
  // la nuestra se llama TipoReg. Al leer se aceptan variantes por si la lista
  // se creó con otro nombre.
  campoTipo: 'TipoReg',
  pollSeconds: 30,
  ventanaDias: 14,             // la quincena que reporta el deck
};

const CONTACTO_ATX = 'Nicolas Blanco  ·  AI Solutions Architect  ·  nicolas@atx.mx  ·  atxlab.ai';
const CONTACTO_CIERRE = 'Nicolas Blanco  ·  nicolas@atx.mx  ·  +52 241 135 3897  ·  atxlab.ai';

let DB = { proyecto: [], actividad: [], requisito: [], comentario: [], persona: [] };
const SPID = {};
let usuario = null;
let vista = { pantalla: 'general', proyecto: null, tab: 'actividades', filtro: 'todas' };
let borrador = null;           // plan cargado, esperando confirmación

const uid = () => Math.random().toString(36).slice(2, 10);
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const hoyISO = () => PlanATX.iso(new Date());
const dias = (a, b) => PlanATX.dias(a, b);

const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
  'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
function fechaLarga(isoStr) {
  if (!isoStr) return 'por definir';
  const f = new Date(isoStr + 'T12:00:00');
  return f.getDate() + ' de ' + MESES[f.getMonth()];
}
function fechaCorta(isoStr) {
  if (!isoStr) return '—';
  const f = new Date(isoStr + 'T12:00:00');
  return f.getDate() + ' ' + MESES[f.getMonth()].slice(0, 3);
}

/* ══════════════════════════════════════════════════════════════════════
   1) PERSISTENCIA
   ══════════════════════════════════════════════════════════════════════ */
const LKEY = 'atxlab-tablero-v3';
let pca = null, siteId = null, listId = null;

async function token() {
  const cuenta = pca.getAllAccounts()[0];
  const req = { scopes: ['Sites.ReadWrite.All'], account: cuenta };
  try { return (await pca.acquireTokenSilent(req)).accessToken; }
  catch (e) { return (await pca.acquireTokenPopup(req)).accessToken; }
}

async function graph(url, opts = {}) {
  const t = await token();
  const r = await fetch(url.startsWith('http') ? url : 'https://graph.microsoft.com/v1.0' + url, {
    ...opts,
    headers: { Authorization: 'Bearer ' + t, 'Content-Type': 'application/json', ...(opts.headers || {}) },
  });
  if (!r.ok) throw new Error('Graph ' + r.status + ' · ' + (await r.text()).slice(0, 200));
  return r.status === 204 ? null : r.json();
}

async function resolverLista() {
  if (listId) return;
  const site = await graph('/sites/' + CONFIG.siteHost + ':' + CONFIG.sitePath);
  siteId = site.id;
  const listas = await graph('/sites/' + siteId + '/lists?$select=id,name,displayName');
  const l = listas.value.find((x) => x.displayName === CONFIG.listName || x.name === CONFIG.listName);
  if (!l) throw new Error('No encontré la lista "' + CONFIG.listName + '" en el sitio.');
  listId = l.id;
}

async function cargar() {
  if (CONFIG.backend === 'local') {
    const raw = localStorage.getItem(LKEY);
    DB = raw ? JSON.parse(raw) : semilla();
    if (!raw) localStorage.setItem(LKEY, JSON.stringify(DB));
    return;
  }
  await resolverLista();
  const nuevo = { proyecto: [], actividad: [], requisito: [], comentario: [], persona: [] };
  let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=500';
  while (url) {
    const r = await graph(url);
    r.value.forEach((it) => {
      const f = it.fields || {};
      const tipo = f[CONFIG.campoTipo] || f.TipoReg || f.Tipo || f.Tipo0 || f.TipoRegistro;
      if (!tipo || !f.Datos) return;
      try {
        const d = JSON.parse(f.Datos);
        if (nuevo[tipo]) { nuevo[tipo].push(d); SPID[d.id] = it.id; }
      } catch (e) { console.warn('registro ilegible', it.id); }
    });
    url = r['@odata.nextLink'] || null;
  }
  DB = nuevo;
}

async function guardar(tipo, obj) {
  const lista = DB[tipo];
  const i = lista.findIndex((x) => x.id === obj.id);
  if (i >= 0) lista[i] = obj; else lista.push(obj);
  if (CONFIG.backend === 'local') { localStorage.setItem(LKEY, JSON.stringify(DB)); return; }
  await resolverLista();
  const fields = { Title: (obj.titulo || obj.nombre || tipo).slice(0, 200), Ref: obj.id, Datos: JSON.stringify(obj) };
  fields[CONFIG.campoTipo] = tipo;
  if (SPID[obj.id]) {
    await graph('/sites/' + siteId + '/lists/' + listId + '/items/' + SPID[obj.id] + '/fields',
      { method: 'PATCH', body: JSON.stringify(fields) });
  } else {
    const r = await graph('/sites/' + siteId + '/lists/' + listId + '/items',
      { method: 'POST', body: JSON.stringify({ fields }) });
    SPID[obj.id] = r.id;
  }
}

async function guardarVarios(pares) { for (const [t, o] of pares) await guardar(t, o); }

async function borrar(tipo, id) {
  DB[tipo] = DB[tipo].filter((x) => x.id !== id);
  if (CONFIG.backend === 'local') { localStorage.setItem(LKEY, JSON.stringify(DB)); return; }
  if (SPID[id]) {
    await graph('/sites/' + siteId + '/lists/' + listId + '/items/' + SPID[id], { method: 'DELETE' });
    delete SPID[id];
  }
}

/* ══════════════════════════════════════════════════════════════════════
   2) CÁLCULOS — misma regla que la skill: horas del plan consumidas
   ══════════════════════════════════════════════════════════════════════ */
const actsDe = (pid) => DB.actividad.filter((a) => a.proyecto === pid);
const reqsDe = (pid) => DB.requisito.filter((r) => r.proyecto === pid);
const comsDe = (pid, aid) => DB.comentario
  .filter((c) => c.proyecto === pid && (aid === undefined ? !c.actividad : c.actividad === aid))
  .sort((a, b) => (a.fecha < b.fecha ? 1 : -1));

const semanaDe = (p, isoStr) => Math.max(0, Math.floor(dias(p.kickoff, isoStr) / 7));
const semanaActual = (p) => semanaDe(p, hoyISO());
function semanasTotales(p) {
  let fin = p.fin;
  if (!fin || fin <= p.kickoff) {
    const fines = actsDe(p.id).map((a) => a.fin).filter(Boolean).sort();
    fin = fines[fines.length - 1] || fin;
  }
  return fin && fin > p.kickoff ? Math.max(1, Math.ceil(dias(p.kickoff, fin) / 7)) : 0;
}

/** Comentarios marcados como riesgo y todavía no atendidos. */
function riesgosDe(pid, aid) {
  return DB.comentario.filter((c) => c.proyecto === pid && c.actividad === aid && c.riesgo && !c.atendido);
}
const enRiesgo = (a) => riesgosDe(a.proyecto, a.id).length > 0;

function bloqueosDe(a) {
  return reqsDe(a.proyecto).filter((r) => r.estado !== 'recibido' && (r.bloquea || []).includes(a.id));
}
function estaDetenida(a) {
  return a.avance < 100 && a.ini && a.ini <= hoyISO() && bloqueosDe(a).length > 0;
}

function metricas(p) {
  const acts = actsDe(p.id), hoy = hoyISO();
  const horas = acts.reduce((s, a) => s + (+a.horas || 0), 0);

  // Si el plan no trae horas, todas las actividades pesan igual. Sin este
  // respaldo la división queda 0/0 y el tablero reporta 0% teniendo avance.
  const sinHoras = horas === 0 && acts.length > 0;
  const peso = (a) => (sinHoras ? 1 : (+a.horas || 0));
  const total = sinHoras ? acts.length : (horas || 1);

  const conFechas = acts.filter((a) => a.ini && a.fin);
  const sinFechas = acts.length > 0 && conFechas.length === 0;

  const real = acts.reduce((s, a) => s + peso(a) * (+a.avance || 0) / 100, 0);
  const plan = acts.reduce((s, a) => {
    if (!a.ini || !a.fin) return s;
    const dur = Math.max(1, dias(a.ini, a.fin) + 1);
    const frac = Math.max(0, Math.min(1, (dias(a.ini, hoy) + 1) / dur));
    return s + peso(a) * frac;
  }, 0);

  const reqAbiertos = reqsDe(p.id).filter((r) => r.estado !== 'recibido');
  const vencidos = reqAbiertos.filter((r) => r.fecha && r.fecha < hoy);
  const detenidas = acts.filter(estaDetenida);
  const riesgos = acts.filter(enRiesgo);
  // Atrasada: su fecha de fin ya pasó y no está al 100%
  const atrasadas = acts.filter((a) => a.avance < 100 && a.fin && a.fin < hoy);

  const pReal = Math.round(real / total * 100), pPlan = Math.round(plan / total * 100);
  const delta = pReal - pPlan;

  // Sin fechas no hay plan contra el cual comparar. Pintar verde ahí sería
  // decir "todo bien" precisamente cuando el tablero no sabe nada.
  const sinPlan = sinFechas || !p.fin || p.fin <= p.kickoff;

  let semaforo;
  if (sinPlan) {
    semaforo = 'sin';
  } else {
    semaforo = 'verde';
    const vencidoViejo = vencidos.some((r) => dias(r.fecha, hoy) > 14);
    if (delta < -10 || vencidoViejo) semaforo = 'rojo';
    else if (delta <= -4 || vencidos.length || detenidas.length || riesgos.length) semaforo = 'ambar';
  }

  const m = { sem: semanaActual(p), semTot: semanasTotales(p), total, pReal, pPlan, delta,
              semaforo, sinHoras, sinFechas, sinPlan,
              reqAbiertos, vencidos, detenidas, riesgos, atrasadas, horasReales: real };
  m.retraso = retrasoDias(p, m);
  m.estimado = m.retraso === null ? null : PlanATX.masDias(p.fin, m.retraso);
  m.desvioDias = m.retraso;
  return m;
}

/**
 * Retraso medido en TIEMPO, no en proporción.
 *
 * Busca en qué fecha el plan alcanzaba el avance que tenemos hoy. Si el plan
 * llegaba a este punto hace 20 días, vamos 20 días atrás y el cierre se
 * recorre 20 días. Antes esto era una regla de tres sobre el ritmo
 * (avance/plan), que en las primeras semanas divide entre un número diminuto
 * y proyectaba disparates: 18 puntos abajo en la semana 2 daban 587 días.
 */
function curvaPlan(p, f) {
  const acts = actsDe(p.id);
  const horas = acts.reduce((s, a) => s + (+a.horas || 0), 0);
  const sinHoras = horas === 0 && acts.length > 0;
  const peso = (a) => (sinHoras ? 1 : (+a.horas || 0));
  const total = sinHoras ? acts.length : (horas || 1);
  const acum = acts.reduce((s, a) => {
    if (!a.ini || !a.fin) return s;
    const dur = Math.max(1, dias(a.ini, a.fin) + 1);
    return s + peso(a) * Math.max(0, Math.min(1, (dias(a.ini, f) + 1) / dur));
  }, 0);
  return acum / total * 100;
}

/** Días de retraso: cuánto tiempo atrás quedó el avance real respecto al plan. */
function retrasoDias(p, m) {
  if (m.sinPlan || !p.fin || p.fin <= p.kickoff) return null;
  // Con muy poco avance capturado no hay evidencia para proyectar nada.
  if (m.pReal < 5) return null;
  if (dias(p.kickoff, hoyISO()) < 10) return null;

  const hoy = hoyISO();
  // La curva del plan no decrece, así que se puede buscar por bisección.
  let lo = 0, hi = dias(p.kickoff, p.fin);
  if (curvaPlan(p, p.fin) < m.pReal) return dias(p.fin, hoy);   // vamos adelante del plan completo
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2);
    if (curvaPlan(p, PlanATX.masDias(p.kickoff, mid)) >= m.pReal) hi = mid; else lo = mid + 1;
  }
  return dias(PlanATX.masDias(p.kickoff, lo), hoy);
}

/** Cierre estimado = fecha acordada recorrida por el retraso observado. */
function fechaEstimada(p, m) {
  const r = retrasoDias(p, m);
  if (r === null) return null;
  return PlanATX.masDias(p.fin, r);
}

/* ══════════════════════════════════════════════════════════════════════
   3) EL CORTE — filtrado a la ventana de la quincena
   ══════════════════════════════════════════════════════════════════════ */
function construirCorte(p) {
  const m = metricas(p), acts = actsDe(p.id), hoy = hoyISO();
  const desde = PlanATX.masDias(hoy, -CONFIG.ventanaDias);
  const periodo = fechaLarga(desde) + ' – ' + fechaLarga(hoy) + ' de ' + new Date(hoy + 'T12:00:00').getFullYear();
  const corta = (t, max) => {                    // corta en palabra, no a la mitad
    max = max || 96;
    if (t.length <= max) return t;
    const c = t.slice(0, max);
    return c.slice(0, Math.max(c.lastIndexOf(' '), max - 18)).trim() + '…';
  };

  // Hitos: máximo 6 alrededor de hoy, para que la línea no se sature
  const todos = (p.hitos || []).slice().sort((a, b) => (a.fecha < b.fecha ? -1 : 1));
  let hitos = todos;
  if (todos.length > 6) {
    let i = todos.findIndex((h) => h.fecha >= hoy);
    if (i < 0) i = todos.length - 1;
    const ini = Math.max(0, Math.min(i - 2, todos.length - 6));
    hitos = todos.slice(ini, ini + 6);
  }
  hitos = hitos.map((h) => {
    const previas = acts.filter((a) => a.fin && a.fin <= h.fecha);
    const listo = previas.length ? previas.every((a) => a.avance === 100) : h.fecha < hoy;
    return { nombre: h.nombre, semana: 'S' + semanaDe(p, h.fecha), estado: listo ? 'cumplido' : 'pendiente' };
  });
  const iCurso = hitos.findIndex((h) => h.estado !== 'cumplido');
  if (iCurso >= 0) hitos[iCurso].estado = 'encurso';

  // Frentes con movimiento en la ventana, máximo 3 y 5 renglones cada uno
  const porHoras = (x, y) => y.horas - x.horas;
  // El título del frente se acorta: los nombres del plan traen paréntesis
  // largos que en una tarjeta de tres columnas se comen media tarjeta.
  const tituloFrente = (n) => {
    let s = String(n).replace(/\s*\([^)]*\)\s*$/, '').trim();
    return s.length > 38 ? s.slice(0, 36).trim() + '…' : s;
  };
  let frentes = (p.frentes || []).map((f) => {
    const propias = acts.filter((a) => a.frente === f.clave);
    return {
      icono: f.icono, titulo: tituloFrente(f.nombre), clave: f.clave,
      cerradas: propias.filter((a) => a.avance === 100 && a.fin && a.fin >= desde).sort(porHoras)
        .map((a) => ({ texto: corta(a.nombre) })),
      enCurso: propias.filter((a) => a.avance > 0 && a.avance < 100 && !estaDetenida(a)).sort(porHoras)
        .map((a) => ({ texto: corta(a.nombre), avance: a.avance })),
      detenidas: propias.filter(estaDetenida).sort(porHoras).map((a) => {
        const r = bloqueosDe(a)[0];
        return { texto: corta(a.nombre, 58) + (r ? ': requiere ' + corta(r.titulo, 46) : ''), avance: a.avance };
      }),
    };
  }).filter((f) => f.cerradas.length + f.enCurso.length + f.detenidas.length > 0);

  frentes.sort((a, b) => (b.detenidas.length * 10 + b.enCurso.length) - (a.detenidas.length * 10 + a.enCurso.length));
  frentes = frentes.slice(0, 3);
  frentes.forEach((f) => {   // lo detenido y lo en curso mandan; lo cerrado rellena
    let cupo = 5;
    f.detenidas = f.detenidas.slice(0, Math.min(3, cupo)); cupo -= f.detenidas.length;
    f.enCurso = f.enCurso.slice(0, Math.max(0, cupo)); cupo -= f.enCurso.length;
    f.cerradas = f.cerradas.slice(0, Math.max(0, cupo));
  });

  // Pendientes del cliente: SOLO lo que se necesita de aquí a dos semanas
  // (o lo que ya venció). Lo que hace falta en noviembre no es material para
  // la sesión de hoy y solo diluye lo urgente.
  const limite = PlanATX.masDias(hoy, 14);
  const pendientes = m.reqAbiertos
    .filter((r) => r.fecha && r.fecha <= limite)
    .sort((a, b) => (a.fecha < b.fecha ? -1 : 1))
    .map((r) => {
      const bloq = acts.filter((a) => (r.bloquea || []).includes(a.id));
      const frenando = bloq.filter((a) => a.ini && a.ini <= hoy && a.avance < 100);
      const vencido = r.fecha < hoy;
      let impacto, tono;
      if (vencido && frenando.length) {
        const d = dias(r.fecha, hoy);
        impacto = 'Vencido hace ' + d + ' días · ' + frenando.length + ' actividad' +
                  (frenando.length > 1 ? 'es detenidas' : ' detenida') + ' esperándolo';
        tono = d > 7 ? 'rojo' : 'ambar';
      } else if (vencido) {
        impacto = 'Vencido hace ' + dias(r.fecha, hoy) + ' días';
        tono = 'ambar';
      } else if (frenando.length) {
        impacto = 'Detiene ' + frenando.length + ' actividad' +
                  (frenando.length > 1 ? 'es' : '') + ' si no llega el ' + fechaLarga(r.fecha);
        tono = 'ambar';
      } else {
        impacto = 'Sin impacto si se recibe antes del ' + fechaLarga(r.fecha);
        tono = 'verde';
      }
      return { lado: 'cliente', orden: vencido ? dias(r.fecha, hoy) : -1,
               titulo: r.titulo,
               etiqueta: p.cliente,
               meta: (r.responsable || 'Responsable por definir') + ' · comprometido para el ' + fechaLarga(r.fecha),
               impacto, tono };
    });

  // Lo atrasado de nuestro lado. Se excluye lo que está detenido esperando al
  // cliente: eso ya lo representa su pendiente, y listarlo aquí sería contar
  // el mismo atraso dos veces y cargárselo a quien no le toca.
  const nuestros = m.atrasadas
    .filter((a) => bloqueosDe(a).length === 0)
    .sort((a, b) => dias(b.fin, hoy) - dias(a.fin, hoy))
    .map((a) => {
      const d = dias(a.fin, hoy);
      return { lado: 'atx', orden: d,
               titulo: corta(a.nombre, 88),
               etiqueta: 'atxlab',
               meta: (a.dueno || 'Sin dueño asignado') + ' · debió cerrar el ' + fechaLarga(a.fin),
               impacto: 'Lleva ' + d + (d === 1 ? ' día' : ' días') + ' de retraso · va al ' + a.avance + '%',
               tono: d > 7 ? 'rojo' : 'ambar' };
    });

  // Ordena por gravedad primero: lo que ya detiene trabajo va arriba, y solo
  // dentro del mismo tono manda la antigüedad. Ordenar solo por días dejaba
  // fuera del corte lo que más impacto tiene.
  const grav = { rojo: 0, ambar: 1, verde: 2 };
  const puntos = pendientes.concat(nuestros)
    .sort((a, b) => (grav[a.tono] - grav[b.tono]) || (b.orden - a.orden))
    .slice(0, 4);   // caben cuatro en la slide; más se recortarían en silencio

  // Lo que se trabaja en las próximas dos semanas, con dueño y fecha límite
  const pasos = acts
    .filter((a) => a.avance < 100 && a.ini && a.ini <= limite)
    .sort((a, b) => ((a.ini < b.ini ? -1 : 1) || (b.horas - a.horas)))
    .slice(0, 4)
    .map((a) => ({
      texto: corta(a.nombre, 72),
      meta: [a.dueno || 'Sin dueño asignado',
             a.fin ? 'para el ' + fechaLarga(a.fin) : null].filter(Boolean).join('  ·  '),
    }));

  // Gantt por frente: barra planeada, avance real y dónde debería ir el plan
  const semTot = m.semTot;
  const filasGantt = (p.frentes || []).map((f) => {
    const propias = acts.filter((a) => a.frente === f.clave && a.ini && a.fin);
    if (!propias.length) return null;
    const hrs = propias.reduce((s2, a) => s2 + (+a.horas || 0), 0);
    const peso = (a) => (hrs ? (+a.horas || 0) / hrs : 1 / propias.length);
    const avance = Math.round(propias.reduce((s2, a) => s2 + peso(a) * a.avance, 0));
    const plan = Math.round(propias.reduce((s2, a) => {
      const dur = Math.max(1, dias(a.ini, a.fin) + 1);
      return s2 + peso(a) * 100 * Math.max(0, Math.min(1, (dias(a.ini, hoy) + 1) / dur));
    }, 0));
    const ini = Math.min.apply(null, propias.map((a) => semanaDe(p, a.ini)));
    const fin = Math.max.apply(null, propias.map((a) => semanaDe(p, a.fin) + 1));
    return { nombre: tituloFrente(f.nombre), ini, fin, avance, plan };
  }).filter(Boolean).sort((a, b) => a.ini - b.ini).slice(0, 6);

  const ganttData = filasGantt.length ? {
    semanas: semTot, semanaActual: m.sem, filas: filasGantt,
    hitos: (p.hitos || []).map((h) => ({
      semana: semanaDe(p, h.fecha),
      cumplido: acts.filter((a) => a.fin && a.fin <= h.fecha).every((a) => a.avance === 100),
    })),
  } : null;

  const T = {
    verde: 'El proyecto avanza\nconforme al plan',
    ambar: 'El proyecto avanza,\ncon puntos por destrabar',
    rojo: 'El proyecto va retrasado\nfrente al plan',
  };
  const limpio = (s) => String(s).replace(/[^A-Za-z0-9]/g, '');

  return {
    proyecto: p.nombre, cliente: p.cliente, contactoCliente: p.contactoCliente || '',
    objetivo: p.objetivo || '',
    semana: m.sem, semanasTotales: m.semTot, periodo,
    semaforo: m.semaforo === 'sin' ? 'ambar' : m.semaforo,
    sinPlan: !!m.sinPlan, sinHoras: !!m.sinHoras,
    avanceReal: m.pReal, avancePlan: m.pPlan,
    hitos, gantt: ganttData, frentes,
    pendientes: puntos, pasos,
    proximaSesion: fechaLarga(p.proximaSesion),
    contactoAtx: CONTACTO_ATX, contactoCierre: CONTACTO_CIERRE,
    tituloAvance: T[m.semaforo],
    tituloActividades: 'Lo trabajado en estas dos semanas',
    tituloPendientes: 'Puntos abiertos',
    tituloSigue: 'Lo que haremos del ' + fechaLarga(hoy) + '\nal ' + fechaLarga(limite),
    archivo: limpio(p.cliente) + '_' + limpio(p.nombre) + '_Avance_S' + String(m.sem).padStart(2, '0') + '.pptx',
  };
}

/* ══════════════════════════════════════════════════════════════════════
   4) VISTAS
   ══════════════════════════════════════════════════════════════════════ */
const $ = (s) => document.querySelector(s);
const app = () => $('#app');

const acts0 = (p) => actsDe(p.id).length;
const barra = (pct, plan) => '<div class="barra"><div class="barra-fill" style="width:' + Math.min(pct, 100) + '%"></div>' +
  (plan != null ? '<div class="barra-plan" style="left:' + Math.min(plan, 100) + '%"></div>' : '') + '</div>';
/** Fechas de cierre al final de la barra: la acordada y la proyectada. */
function cierres(p, m) {
  const valida = p.fin && p.fin > p.kickoff;
  const acordada = '<span class="cierre acordada">Cierre acordado<b>' +
    (valida ? fechaFull(p.fin) : 'sin definir') + '</b></span>';
  if (!m.estimado) {
    const motivo = !valida ? 'falta la fecha de cierre'
      : (m.sinFechas ? 'las actividades no tienen fechas'
      : (m.pReal < 5 ? 'muy poco avance capturado para proyectar'
      : 'aún no hay historial suficiente'));
    return '<div class="cierres">' + acordada +
      '<span class="cierre sin">Cierre estimado<b>' + motivo + '</b></span></div>';
  }
  // En semanas: "23 días después" sugiere una precisión que el dato no tiene.
  const d = m.desvioDias;
  const cls = d <= 3 ? 'ok' : (d <= 14 ? 'alerta' : 'mal');
  const sem = Math.round(Math.abs(d) / 7);
  const txt = Math.abs(d) <= 3 ? 'en la fecha acordada'
    : (sem <= 1 ? (d > 0 ? '≈1 semana después' : '≈1 semana antes')
               : '≈' + sem + ' semanas ' + (d > 0 ? 'después' : 'antes'));
  return '<div class="cierres">' + acordada +
    '<span class="cierre estimada ' + cls + '">Cierre estimado con el retraso actual' +
    '<b>' + fechaFull(m.estimado) + '</b><i>' + txt + '</i></span></div>';
}

function fechaFull(isoStr) {
  if (!isoStr) return 'sin fecha';
  const f = new Date(isoStr + 'T12:00:00');
  return f.getDate() + ' de ' + MESES[f.getMonth()] + ' de ' + f.getFullYear();
}

const ETIQUETA = { verde: 'En tiempo', ambar: 'Con atención', rojo: 'En riesgo',
                   sin: 'Sin datos de plan' };
const chipSem = (s) => '<span class="chip chip-' + s + '"><i></i>' + ETIQUETA[s] + '</span>';

/**
 * Por qué el proyecto está en el color que está.
 * `corto` es lo que se ve directo en el tablero; los nombres concretos ya
 * viven en el panel de atención de abajo, así que aquí no se repiten.
 */
function razonesSemaforo(p, m) {
  const hoy = hoyISO(), r = [];
  const pl = (n, s, pl2) => n + ' ' + (n === 1 ? s : (pl2 || s + 's'));

  if (m.sinPlan) {
    if (m.sinFechas) r.push({ peso: 'sin', corto: 'Las actividades no tienen fechas' });
    if (!p.fin || p.fin <= p.kickoff) r.push({ peso: 'sin', corto: 'El proyecto no tiene fecha de cierre' });
    if (m.sinHoras) r.push({ peso: 'sin', corto: 'El plan no trae horas: todas pesan igual' });
    r.push({ peso: 'sin', corto: 'No hay plan contra el cual comparar el avance' });
    return r;
  }
  if (m.sinHoras) {
    r.push({ peso: 'sin', corto: 'El plan no trae horas: todas las actividades pesan igual' });
  }
  if (m.delta <= -4) {
    r.push({ peso: m.delta < -10 ? 'rojo' : 'ambar',
             corto: 'Avance ' + Math.abs(m.delta) + ' puntos porcentuales abajo del plan' });
  }
  if (m.atrasadas.length) {
    r.push({ peso: 'ambar', corto: pl(m.atrasadas.length, 'actividad con fecha vencida', 'actividades con fecha vencida') });
  }
  const viejos = m.vencidos.filter((x) => dias(x.fecha, hoy) > 14);
  if (viejos.length) {
    r.push({ peso: 'rojo', corto: pl(viejos.length, 'entregable vencido', 'entregables vencidos') + ' hace más de 14 días' });
  } else if (m.vencidos.length) {
    r.push({ peso: 'ambar', corto: pl(m.vencidos.length, 'entregable vencido', 'entregables vencidos') + ' del cliente' });
  }
  if (m.detenidas.length) {
    r.push({ peso: 'ambar', corto: pl(m.detenidas.length, 'actividad detenida', 'actividades detenidas') });
  }
  if (m.riesgos.length) {
    r.push({ peso: 'ambar', corto: pl(m.riesgos.length, 'riesgo abierto', 'riesgos abiertos') });
  }
  if (!r.length) r.push({ peso: 'verde', corto: 'Sin desviación, nada detenido y sin riesgos' });
  return r;
}

/** Tira de razones, visible en el tablero bajo la barra de avance. */
function tiraRazones(p, m) {
  const items = razonesSemaforo(p, m).map((x) =>
    '<span class="razon r-' + x.peso + '"><i></i>' + esc(x.corto) + '</span>').join('');
  return '<div class="razones"><span class="razones-etq">Estatus por</span>' + items + '</div>';
}

/** Chip con la leyenda de los tres colores al pasar el cursor. */
function chipSemDetalle(p, m) {
  const regla = [
    ['verde', 'Hasta 3 puntos abajo del plan, sin entregables vencidos, nada detenido y sin riesgos abiertos.'],
    ['ambar', 'De 4 a 10 puntos abajo del plan, o hay un entregable vencido, una actividad detenida o un riesgo abierto.'],
    ['rojo', 'Más de 10 puntos abajo del plan, o un entregable del cliente lleva más de 14 días vencido.'],
    ['sin', 'Faltan fechas u horas en el plan. Sin eso no hay contra qué comparar, así que no se emite un estatus.'],
  ].map(([c, d]) => '<div class="sem-regla' + (c === m.semaforo ? ' actual' : '') + '">' +
    '<span class="chip chip-' + c + '"><i></i>' + ETIQUETA[c] + '</span>' +
    '<span class="sem-desc">' + d + '</span></div>').join('');

  return '<span class="sem-wrap" tabindex="0">' + chipSem(m.semaforo) +
    '<div class="sem-pop"><h5>Cómo se define el estatus</h5>' + regla + '</div></span>';
}

function render() {
  if (vista.pantalla === 'general') return renderGeneral();
  if (vista.pantalla === 'equipo') return renderEquipo();
  if (vista.pantalla === 'confirmar') return renderConfirmar();
  return renderProyecto();
}

function renderGeneral() {
  const cards = DB.proyecto.map((p) => {
    const m = metricas(p);
    return '<div class="card proj" onclick="abrir(\'' + p.id + '\')">' +
      '<div class="card-top"><div><div class="cliente">' + esc(p.cliente) + '</div>' +
      '<h3>' + esc(p.nombre) + '</h3></div>' + chipSem(m.semaforo) + '</div>' +
      '<div class="pct">' + m.pReal + '%<span> de avance · plan ' + m.pPlan + '%</span></div>' +
      barra(m.pReal, m.pPlan) +
      '<div class="meta"><span>Semana ' + m.sem + (m.semTot ? ' de ' + m.semTot : '') + '</span>' +
      (m.vencidos.length ? '<span class="alerta">' + m.vencidos.length + ' entregable(s) del cliente vencido(s)</span>' :
        (m.reqAbiertos.length ? '<span>' + m.reqAbiertos.length + ' pendiente(s) del cliente</span>' :
          '<span class="ok">Sin pendientes del cliente</span>')) +
      '</div></div>';
  }).join('');

  app().innerHTML = '<div class="head"><h1>Proyectos en curso</h1>' +
    '<div class="acciones">' +
    (CONFIG.backend === 'local' ? '<button class="btn ghost" onclick="reiniciarLocal()">Vaciar datos locales</button>' : '') +
    '<button class="btn" onclick="formProyecto()">Nuevo proyecto</button></div></div>' +
    '<div class="grid">' + (cards || '<p class="vacio">Todavía no hay proyectos dados de alta.</p>') + '</div>';
}

/**
 * Panel de atención de UN proyecto. Cada bloque ocupa el ancho completo y
 * reparte sus elementos en horizontal, para no empujar la vista hacia abajo.
 * Un bloque sin contenido no se dibuja: el espacio vacío no informa nada.
 */
function panelAtencion(p) {
  const hoy = hoyISO(), pronto = PlanATX.masDias(hoy, 7), desde = PlanATX.masDias(hoy, -7);
  const m = metricas(p);

  // Actividades con problema: detenidas por el cliente o con fecha vencida
  const problema = [];
  m.detenidas.forEach((a) => {
    const r = bloqueosDe(a)[0];
    problema.push({ a, tipo: 'detenida', por: r ? 'espera: ' + r.titulo : 'espera algo del cliente' });
  });
  m.atrasadas.forEach((a) => {
    if (problema.some((x) => x.a.id === a.id)) return;
    problema.push({ a, tipo: 'atrasada', por: 'debió cerrar el ' + fechaCorta(a.fin) + ' · lleva ' + dias(a.fin, hoy) + ' días' });
  });

  const riesgos = DB.comentario.filter((c) => c.proyecto === p.id && c.riesgo && !c.atendido)
    .map((c) => {
      const a = DB.actividad.find((x) => x.id === c.actividad);
      return { aid: a ? a.id : null, donde: a ? a.nombre : 'Comentario general',
               texto: c.texto, autor: c.autor, fecha: c.fecha, cid: c.id };
    }).sort((a, b) => (a.fecha < b.fecha ? 1 : -1));

  const entregables = reqsDe(p.id)
    .filter((r) => r.estado !== 'recibido' && r.fecha && r.fecha <= pronto)
    .sort((a, b) => (a.fecha < b.fecha ? -1 : 1));

  const recientes = DB.comentario
    .filter((c) => c.proyecto === p.id && c.fecha >= desde && !(c.riesgo && !c.atendido))
    .sort((a, b) => (a.fecha < b.fecha ? 1 : -1)).slice(0, 6)
    .map((c) => {
      const a = DB.actividad.find((x) => x.id === c.actividad);
      return { aid: a ? a.id : null, donde: a ? a.nombre : 'General', texto: c.texto, autor: c.autor, fecha: c.fecha };
    });

  const fila = (titulo, n, clase, items, extra) => !items ? '' :
    '<div class="pa-fila"><div class="pa-cab"><h4 class="' + clase + '">' + titulo +
    '<span class="pa-n">' + n + '</span></h4>' + (extra || '') + '</div>' +
    '<div class="pa-items">' + items + '</div></div>';

  const fProblema = problema.length ? fila('Actividades con problema', problema.length, 'mal',
    problema.slice(0, 4).map((x) => '<div class="pa-item' + (x.tipo === 'detenida' ? ' alerta-item' : ' mal-item') +
      '" onclick="irActividad(\'' + x.a.id + '\')">' +
      '<div class="pa-tit">' + esc(x.a.nombre) +
      '<span class="pa-tag ' + x.tipo + '">' + (x.tipo === 'detenida' ? 'Detenida' : 'Atrasada') + '</span></div>' +
      '<div class="pa-pie">' + (x.a.dueno ? esc(x.a.dueno) : 'sin dueño') + ' · ' + esc(x.por) + '</div></div>').join(''),
    problema.length > 4 ? '<button class="link" onclick="irTabFiltro(\'problema\')">Ver las ' +
      problema.length + ' completas</button>' : '') : '';

  const fRiesgos = riesgos.length ? fila('Riesgos abiertos', riesgos.length, 'mal',
    riesgos.slice(0, 4).map((r) => '<div class="pa-item mal-item"' +
      (r.aid ? ' onclick="irActividad(\'' + r.aid + '\')"' : '') + '>' +
      '<div class="pa-tit">' + esc(r.donde) + '</div>' +
      '<div class="pa-txt">' + esc(r.texto) + '</div>' +
      '<div class="pa-pie">' + esc(r.autor) + ' · ' + fechaCorta(r.fecha) +
      ' · <button class="link" onclick="event.stopPropagation();atenderRiesgo(\'' + r.cid + '\')">atendido</button>' +
      '</div></div>').join('')) : '';

  const fEnt = entregables.length ? fila('Entregables del cliente', entregables.length, 'alerta',
    entregables.slice(0, 4).map((e) => {
      const venc = e.fecha < hoy;
      return '<div class="pa-item' + (venc ? ' mal-item' : ' alerta-item') +
        '" onclick="irRequisito(\'' + e.id + '\')">' +
        '<div class="pa-tit">' + esc(e.titulo) + '</div>' +
        '<div class="pa-pie">' + (e.responsable ? esc(e.responsable) + ' · ' : '') +
        (venc ? 'vencido hace ' + dias(e.fecha, hoy) + ' días' : 'vence ' + fechaCorta(e.fecha)) +
        '</div></div>';
    }).join(''),
    entregables.length > 4 ? '<button class="link" onclick="irTab(\'requisitos\')">Ver todos</button>' : '') : '';

  const fCom = recientes.length ? fila('Comentarios de la semana', recientes.length, '',
    recientes.slice(0, 4).map((c) => '<div class="pa-item"' +
      (c.aid ? ' onclick="irActividad(\'' + c.aid + '\')"' : '') + '>' +
      '<div class="pa-txt">' + esc(c.texto) + '</div>' +
      '<div class="pa-pie">' + esc(c.donde) + ' · ' + esc(c.autor) + ' · ' + fechaCorta(c.fecha) + '</div></div>').join('')) : '';

  const todo = fProblema + fRiesgos + fEnt + fCom;
  return todo ? '<div class="panel-atencion">' + todo + '</div>' : '';
}

/** Lleva a la actividad dentro de la lista cronológica y la resalta. */
function irActividad(id) {
  vista.tab = 'actividades';
  vista.filtro = 'todas';
  render();
  resaltar('act-' + id);
}

function irRequisito(id) {
  vista.tab = 'requisitos';
  render();
  resaltar('req-' + id);
}

function resaltar(domId) {
  setTimeout(() => {
    const el = document.getElementById(domId);
    if (!el) return;
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    el.classList.add('resaltada');
    setTimeout(() => el.classList.remove('resaltada'), 2600);
  }, 60);
}

function renderProyecto() {
  const p = DB.proyecto.find((x) => x.id === vista.proyecto);
  if (!p) { vista.pantalla = 'general'; return render(); }
  const m = metricas(p);
  const tabs = ['actividades', 'requisitos', 'comentarios']
    .map((t) => '<button class="tab' + (vista.tab === t ? ' on' : '') + '" onclick="irTab(\'' + t + '\')">' +
      (t === 'requisitos' ? 'Entregables del cliente' : t[0].toUpperCase() + t.slice(1)) + '</button>').join('');

  app().innerHTML =
    '<div class="head"><div><button class="link" onclick="volver()">← Proyectos</button>' +
    '<h1>' + esc(p.nombre) + ' <span class="sub">' + esc(p.cliente) + '</span></h1></div>' +
    '<div class="acciones">' +
    '<button class="btn ghost" onclick="copiarCorte()">Copiar corte para Claude</button>' +
    '<button class="btn" onclick="descargarDeck()">Descargar deck</button></div></div>' +

    '<div class="card resumen"><div class="res-izq"><div class="pct grande">' + m.pReal + '%</div>' +
    '<div class="cap">de avance real</div></div><div class="res-der">' +
    '<div class="res-top">' + chipSemDetalle(p, m) + '<span class="delta">' +
    (m.delta === 0 ? 'En línea con el plan' : (m.delta > 0 ? m.delta + ' puntos adelante' : Math.abs(m.delta) + ' puntos abajo del plan')) +
    '</span></div>' + barra(m.pReal, m.pPlan) + cierres(p, m) + tiraRazones(p, m) +
    '<div class="meta"><span>Plan a la fecha ' + m.pPlan + '%</span>' +
    '<span>Semana ' + m.sem + (m.semTot ? ' de ' + m.semTot : '') + '</span>' +
    '<span>' + (m.sinHoras ? acts0(p) + ' actividades (plan sin horas)'
      : Math.round(m.horasReales) + ' de ' + Math.round(m.total) + ' h del plan') + '</span>' +
    '<span>Kickoff ' + fechaCorta(p.kickoff) + '</span>' +
    '<span><button class="link" onclick="formProyectoEditar()">Editar proyecto</button></span></div></div></div>' +

    panelAtencion(p) +
    '<div class="tabs">' + tabs + '</div><div id="tabbody"></div>';
  renderTab(p);
}

function renderTab(p) {
  if (vista.tab === 'actividades') return renderActividades(p);
  if (vista.tab === 'requisitos') return renderRequisitos(p);
  return renderComentarios(p);
}

const FILTROS = {
  todas:     { etq: 'Todas',      test: () => true },
  problema:  { etq: 'Con problema', test: (a) => estaDetenida(a) || (a.avance < 100 && a.fin && a.fin < hoyISO()) },
  detenidas: { etq: 'Detenidas',  test: (a) => estaDetenida(a) },
  atrasadas: { etq: 'Atrasadas',  test: (a) => a.avance < 100 && a.fin && a.fin < hoyISO() },
  riesgo:    { etq: 'Con riesgo', test: (a) => enRiesgo(a) },
  abiertas:  { etq: 'Sin cerrar', test: (a) => a.avance < 100 },
};

function irTabFiltro(f) {
  vista.tab = 'actividades'; vista.filtro = f; render();
  setTimeout(() => {
    const el = document.getElementById('tabbody');
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 60);
}
function ponerFiltro(f) { vista.filtro = f; render(); }

function renderActividades(p) {
  const hoy = hoyISO();
  const filtro = FILTROS[vista.filtro] || FILTROS.todas;
  const todas = actsDe(p.id);

  const barraFiltros = '<div class="filtros">' + Object.keys(FILTROS).map((k) => {
    const n = todas.filter(FILTROS[k].test).length;
    if (k !== 'todas' && !n) return '';
    return '<button class="filtro' + (vista.filtro === k ? ' on' : '') + '" onclick="ponerFiltro(\'' + k + '\')">' +
      FILTROS[k].etq + '<span>' + n + '</span></button>';
  }).join('') + '</div>';

  const bloques = (p.frentes || []).map((f) => {
    const acts = actsDe(p.id).filter((a) => a.frente === f.clave).filter(filtro.test)
      .sort((a, b) => ((a.ini || '') < (b.ini || '') ? -1 : 1));
    if (!acts.length) return '';
    const hrs = acts.reduce((s, a) => s + (+a.horas || 0), 0);
    const av = hrs ? Math.round(acts.reduce((s, a) => s + a.horas * a.avance / 100, 0) / hrs * 100) : 0;
    return '<div class="frente-bloque"><h4 class="frente">' + f.icono + ' ' + esc(f.nombre) +
      '<span class="frente-av">' + av + '% · ' + hrs + ' h</span></h4>' +
      acts.map((a) => {
        const det = estaDetenida(a), bloq = bloqueosDe(a), ries = riesgosDe(p.id, a.id);
        const nc = comsDe(p.id, a.id).length;
        const tarde = a.avance < 100 && a.fin && a.fin < hoy;
        return '<div class="act' + (ries.length ? ' riesgo' : (det ? ' detenida' : (tarde ? ' atrasada' : ''))) +
          '" id="act-' + a.id + '">' +
          '<button class="palomear av-' + a.avance + '" onclick="ciclar(\'' + a.id + '\')">' +
          (a.avance === 100 ? '✓' : a.avance + '%') + '</button>' +
          '<div class="act-cuerpo"><div class="act-tit">' + esc(a.nombre) +
          (ries.length ? '<span class="badge-riesgo">Riesgo</span>' : '') + '</div>' +
          '<div class="act-meta">' +
          '<span class="' + (a.dueno ? 'dueno' : 'dueno sin') + '" onclick="asignarDueno(\'' + a.id + '\')">' +
          (a.dueno ? '👤 ' + esc(a.dueno) : '＋ asignar dueño') + '</span>' +
          '<span' + (tarde ? ' class="tarde"' : '') + '>' + fechaCorta(a.ini) + ' → ' + fechaCorta(a.fin) +
          (tarde ? ' · ' + dias(a.fin, hoy) + ' días tarde' : '') + '</span>' +
          '<span>' + a.horas + ' h</span>' +
          (det ? '<span class="badge-stop">Detenida: ' + esc(bloq.map((r) => r.titulo).join(', ')) + '</span>' : '') +
          '</div>' +
          (ries.length ? '<div class="riesgo-txt">⚠ ' + esc(ries[0].texto) +
            ' <i>— ' + esc(ries[0].autor) + '</i></div>' : '') +
          '</div>' +
          '<button class="coment" onclick="verComentarios(\'' + a.id + '\')">💬 ' + (nc || '') + '</button>' +
          '<button class="coment" onclick="formActividad(\'' + a.id + '\')">✎</button>' +
          '<button class="borrar" onclick="quitar(\'actividad\',\'' + a.id + '\')">×</button></div>';
      }).join('') + '</div>';
  }).join('');

  $('#tabbody').innerHTML = '<div class="card">' + barraFiltros +
    (bloques || '<p class="vacio">Ninguna actividad en este filtro.</p>') +
    '<button class="btn ghost full" onclick="formActividad()">Agregar actividad</button></div>';
}

function renderRequisitos(p) {
  const hoy = hoyISO();
  const filas = reqsDe(p.id)
    .sort((a, b) => ((a.fecha || '9999') < (b.fecha || '9999') ? -1 : 1))
    .map((r) => {
      const bloq = actsDe(p.id).filter((a) => (r.bloquea || []).includes(a.id));
      const vencido = r.estado !== 'recibido' && r.fecha && r.fecha < hoy;
      const frenando = bloq.filter((a) => a.ini && a.ini <= hoy && a.avance < 100);
      return '<div class="req' + (r.estado === 'recibido' ? ' ok' : (vencido ? ' mal' : '')) +
        '" id="req-' + r.id + '">' +
        '<button class="palomear' + (r.estado === 'recibido' ? ' on' : '') + '" onclick="toggleReq(\'' + r.id + '\')">' +
        (r.estado === 'recibido' ? '✓' : '') + '</button>' +
        '<div class="act-cuerpo"><div class="act-tit">' + esc(r.titulo) + '</div>' +
        '<div class="act-meta"><span>' + esc(r.responsable || 'Sin responsable') + '</span>' +
        '<span>' + fechaCorta(r.fecha) + '</span>' +
        (vencido ? '<span class="badge-stop">Vencido hace ' + dias(r.fecha, hoy) + ' días</span>' : '') + '</div>' +
        (bloq.length
          ? '<div class="bloquea">Bloquea: ' + bloq.map((a) => '<span class="' + (frenando.indexOf(a) >= 0 ? 'frenando' : '') + '">' + esc(a.nombre) + '</span>').join(' · ') + '</div>'
          : '<div class="bloquea suelto">No bloquea ninguna actividad</div>') +
        '</div><button class="coment" onclick="formRequisito(\'' + r.id + '\')">✎</button>' +
        '<button class="borrar" onclick="quitar(\'requisito\',\'' + r.id + '\')">×</button></div>';
    }).join('');

  $('#tabbody').innerHTML = '<div class="card">' +
    '<p class="nota">Los entregables marcados con el recurso del cliente en el plan ya están aquí, ' +
    'con lo que bloquean según las predecesoras. Agrega los que no vienen en el archivo: accesos, ' +
    'datos, validaciones, cualquier cosa que necesites de su lado.</p>' +
    (filas || '<p class="vacio">Sin entregables del cliente.</p>') +
    '<button class="btn ghost full" onclick="formRequisito()">Agregar entregable del cliente</button></div>';
}

function renderComentarios(p) {
  const lista = comsDe(p.id).map((c) =>
    '<div class="com' + (c.riesgo && !c.atendido ? ' con-riesgo' : '') + '">' +
    '<div class="com-top"><b>' + esc(c.autor) + '</b><span>' + fechaCorta(c.fecha) + '</span></div>' +
    '<p>' + (c.riesgo ? '<span class="badge-riesgo">Riesgo</span> ' : '') + esc(c.texto) + '</p>' +
    (c.riesgo && !c.atendido ? '<button class="link" onclick="atenderRiesgo(\'' + c.id + '\')">Marcar como atendido</button>' : '') +
    '</div>').join('');
  $('#tabbody').innerHTML = '<div class="card">' +
    '<textarea id="nuevoCom" placeholder="Comentario general del proyecto..."></textarea>' +
    '<label class="check-riesgo"><input type="checkbox" id="nuevoRiesgo"> Marcar como riesgo del proyecto</label>' +
    '<button class="btn ghost" onclick="agregarComentario()">Comentar</button>' +
    (lista || '<p class="vacio">Sin comentarios.</p>') + '</div>';
}

function renderEquipo() {
  const filas = DB.persona.map((q) =>
    '<div class="act"><div class="act-cuerpo"><div class="act-tit">' + esc(q.nombre) + '</div>' +
    '<div class="act-meta"><span>' + esc(q.rol || '') + '</span>' + (q.correo ? '<span>' + esc(q.correo) + '</span>' : '') + '</div></div>' +
    '<button class="borrar" onclick="quitar(\'persona\',\'' + q.id + '\')">×</button></div>').join('');
  app().innerHTML = '<div class="head"><div><button class="link" onclick="volver()">← Proyectos</button>' +
    '<h1>Equipo atxlab</h1></div></div><div class="card">' +
    (filas || '<p class="vacio">Sin personas dadas de alta.</p>') +
    '<button class="btn ghost full" onclick="formPersona()">Agregar persona</button></div>';
}

/* ══════════════════════════════════════════════════════════════════════
   5) ALTA DE PROYECTO DESDE EL PLAN DE TRABAJO
   ══════════════════════════════════════════════════════════════════════ */
const campo = (etq, ctrl) => '<label class="campo"><span>' + etq + '</span>' + ctrl + '</label>';

function formProyecto() {
  borrador = null;
  modal('Nuevo proyecto',
    '<p class="nota">Sube el plan de trabajo en el formato de ATX. De ahí leo el cliente, las fases, ' +
    'las tareas con horas y fechas, los entregables del cliente y las dependencias.</p>' +
    campo('Plan de trabajo (.xlsx)', '<input type="file" id="fArchivo" accept=".xlsx,.xlsm" onchange="leerArchivo()">') +
    '<div id="resumenPlan"></div>' +
    '<div id="camposProyecto" style="display:none">' +
      campo('Nombre del proyecto', '<input id="fNombre">') +
      campo('Cliente', '<input id="fCliente">') +
      campo('Objetivo del proyecto', '<textarea id="fObjetivo" placeholder="Para qué existe el proyecto. Sirve de contexto al redactar los updates."></textarea>') +
      campo('Contacto del cliente', '<input id="fContacto" placeholder="Quién responde por los entregables de su lado">') +
      campo('Recurso que identifica al cliente en el plan', '<select id="fCodigo"></select>') +
      campo('Fecha real de kickoff', '<input type="date" id="fKickoff" onchange="avisarDesfase()"><span class="ayuda" id="avisoDesfase"></span>') +
      campo('Próxima sesión de avance', '<input type="date" id="fSesion">') +
    '</div>',
    'Revisar el plan →', prepararConfirmacion);
}

function leerArchivo() {
  const f = $('#fArchivo').files[0];
  if (!f) return;
  const lector = new FileReader();
  lector.onload = (e) => {
    try {
      const plan = PlanATX.leer(new Uint8Array(e.target.result));
      borrador = { plan };
      $('#resumenPlan').innerHTML = '<div class="ok-caja">Leí <b>' + plan.tareas.length + ' tareas</b>. ' +
        'El plan corre del ' + fechaCorta(plan.inicioPlan) + ' al ' + fechaCorta(plan.finPlan) +
        (plan.horasEncabezado ? ' · ' + plan.horasEncabezado + ' horas' : '') + '.</div>';
      $('#camposProyecto').style.display = 'block';
      $('#fNombre').value = plan.nombre;
      $('#fCliente').value = plan.cliente;
      $('#fKickoff').value = plan.inicioPlan;
      $('#fCodigo').innerHTML = plan.codigos.map((c) =>
        '<option value="' + esc(c) + '"' + (/^atx/i.test(c) ? '' : ' selected') + '>' + esc(c) + '</option>').join('');
      avisarDesfase();
    } catch (err) {
      $('#resumenPlan').innerHTML = '<div class="mal-caja">' + esc(err.message) + '</div>';
      $('#camposProyecto').style.display = 'none';
      borrador = null;
    }
  };
  lector.readAsArrayBuffer(f);
}

function avisarDesfase() {
  if (!borrador) return;
  const k = $('#fKickoff').value;
  const d = k ? dias(borrador.plan.inicioPlan, k) : 0;
  $('#avisoDesfase').textContent = !k ? '' : (d === 0 ? 'Igual que el plan.'
    : (d > 0 ? 'Recorro el plan completo ' + d + ' días hacia adelante.'
             : 'Recorro el plan completo ' + Math.abs(d) + ' días hacia atrás.'));
}

function prepararConfirmacion() {
  if (!borrador) return aviso('Falta el archivo del plan.');
  const kickoff = $('#fKickoff').value;
  if (!kickoff) return aviso('Falta la fecha real de kickoff.');
  borrador.meta = {
    nombre: $('#fNombre').value.trim(), cliente: $('#fCliente').value.trim(),
    objetivo: $('#fObjetivo').value.trim(), contactoCliente: $('#fContacto').value.trim(),
    codigoCliente: $('#fCodigo').value, kickoff, proximaSesion: $('#fSesion').value || '',
  };
  const c = PlanATX.clasificar(borrador.plan, { codigoCliente: borrador.meta.codigoCliente, kickoffReal: kickoff });
  borrador.fases = c.fases; borrador.items = c.items; borrador.desfase = c.desfase;
  cerrarModal();
  vista.pantalla = 'confirmar';
  render();
}

function renderConfirmar() {
  const b = borrador;
  const n = (t) => b.items.filter((i) => i.tipo === t).length;
  const mapa = PlanATX.bloqueos(b.items);
  const porIdPlan = {}; b.items.forEach((i) => { porIdPlan[i.idPlan] = i; });

  const filas = b.items.map((i, k) => {
    const bloq = (mapa[i.idPlan] || []).map((id) => porIdPlan[id] && porIdPlan[id].nombre).filter(Boolean);
    return '<tr class="t-' + i.tipo + '"><td class="wbs">' + esc(i.wbs) + '</td>' +
      '<td>' + esc(i.nombre) +
      (i.tipo === 'requisito' ? (bloq.length
        ? '<div class="mini">Bloquea: ' + esc(bloq.join(' · ')) + '</div>'
        : '<div class="mini suelto">Las predecesoras no dicen qué bloquea</div>') : '') +
      '</td><td class="nowrap">' + fechaCorta(i.ini) + ' → ' + fechaCorta(i.fin) + '</td>' +
      '<td class="num">' + (i.horas || '') + '</td>' +
      '<td class="num">' + esc(i.grupo || '—') + '</td>' +
      '<td><select onchange="cambiarTipo(' + k + ', this.value)">' +
      ['actividad', 'requisito', 'hito', 'ignorar'].map((t) =>
        '<option value="' + t + '"' + (i.tipo === t ? ' selected' : '') + '>' +
        ({ actividad: 'Actividad', requisito: 'Entregable del cliente', hito: 'Hito', ignorar: 'Ignorar' }[t]) +
        '</option>').join('') + '</select></td></tr>';
  }).join('');

  app().innerHTML = '<div class="head"><div><button class="link" onclick="cancelarAlta()">← Cancelar</button>' +
    '<h1>Revisa lo que entendí del plan</h1></div>' +
    '<button class="btn" onclick="crearProyecto()">Crear proyecto</button></div>' +
    '<div class="card"><div class="resumen-alta">' +
    '<div><b>' + esc(b.meta.cliente) + '</b> · ' + esc(b.meta.nombre) + '</div>' +
    '<div>' + b.fases.length + ' fases · ' + n('actividad') + ' actividades · ' +
    n('requisito') + ' entregables del cliente · ' + n('hito') + ' hitos</div>' +
    '<div>Kickoff ' + fechaCorta(b.meta.kickoff) +
    (b.desfase ? ' · plan recorrido ' + b.desfase + ' días' : ' · sin desfase respecto al plan') + '</div></div>' +
    avisoPlan(b) +
    '<p class="nota">Los entregables del cliente salieron de las tareas cuyo único recurso es <b>' +
    esc(b.meta.codigoCliente) + '</b>, y lo que bloquean salió de la columna de predecesoras. ' +
    'Corrige aquí lo que no cuadre; una vez creado el proyecto puedes agregar más entregables a mano.</p>' +
    '<div class="tabla-scroll"><table class="tabla"><thead><tr>' +
    '<th>WBS</th><th>Tarea</th><th>Fechas</th><th>Horas</th><th>Recurso</th><th>Se carga como</th>' +
    '</tr></thead><tbody>' + filas + '</tbody></table></div></div>';
}

/**
 * Aviso antes de crear el proyecto. Un plan sin horas o sin fechas se carga
 * igual, pero el avance y el semáforo quedan a medias — mejor decirlo aquí
 * que dejar que el tablero reporte 0% con la mitad de las tareas palomeadas.
 */
function avisoPlan(b) {
  const usados = b.items.filter((i) => i.tipo === 'actividad');
  if (!usados.length) return '';
  const sinH = usados.filter((i) => !i.horas).length;
  const sinF = usados.filter((i) => !i.ini || !i.fin).length;
  const faltas = [];
  if (sinH === usados.length) faltas.push('<b>ninguna actividad trae horas</b>, así que el avance se calculará dando el mismo peso a todas');
  else if (sinH) faltas.push('<b>' + sinH + ' de ' + usados.length + ' actividades no traen horas</b> y pesarán cero en el avance');
  if (sinF === usados.length) faltas.push('<b>ninguna actividad trae fechas</b>, así que no habrá plan a la fecha ni semáforo');
  else if (sinF) faltas.push('<b>' + sinF + ' de ' + usados.length + ' actividades no traen fechas</b> y no contarán para el plan a la fecha');
  if (!faltas.length) return '';
  return '<div class="mal-caja">Revisa la columna de horas y las de fechas en el archivo: ' +
    faltas.join('; y ') + '. Puedes continuar, pero conviene corregir el plan antes.</div>';
}

function cambiarTipo(k, tipo) { borrador.items[k].tipo = tipo; render(); }
function cancelarAlta() { borrador = null; vista.pantalla = 'general'; render(); }

async function crearProyecto() {
  const b = borrador, P = uid();
  const mapa = PlanATX.bloqueos(b.items);
  const usados = b.items.filter((i) => i.tipo !== 'ignorar');
  const finPlan = usados.map((i) => i.fin).filter(Boolean).sort().pop() || b.meta.kickoff;

  const idsNuevos = {};
  const actividades = b.items.filter((i) => i.tipo === 'actividad').map((i) => {
    const a = { id: uid(), proyecto: P, nombre: i.nombre, frente: i.frente, dueno: '',
                horas: i.horas, avance: i.avance || 0, ini: i.ini, fin: i.fin };
    idsNuevos[i.idPlan] = a.id;
    return a;
  });
  const requisitos = b.items.filter((i) => i.tipo === 'requisito').map((i) => ({
    id: uid(), proyecto: P, titulo: i.nombre, detalle: '',
    responsable: b.meta.contactoCliente, fecha: i.fin, estado: 'pendiente',
    bloquea: (mapa[i.idPlan] || []).map((x) => idsNuevos[x]).filter(Boolean),
  }));
  const hitos = b.items.filter((i) => i.tipo === 'hito')
    .map((i) => ({ nombre: i.nombre, fecha: i.fin }))
    .sort((x, y) => (x.fecha < y.fecha ? -1 : 1));

  const proyecto = {
    id: P, nombre: b.meta.nombre, cliente: b.meta.cliente, objetivo: b.meta.objetivo,
    contactoCliente: b.meta.contactoCliente, kickoff: b.meta.kickoff, fin: finPlan,
    proximaSesion: b.meta.proximaSesion,
    frentes: b.fases.map((f) => ({ clave: f.clave, nombre: f.nombre, icono: f.icono })),
    hitos,
  };

  aviso('Guardando ' + (actividades.length + requisitos.length + 1) + ' registros...');
  try {
    await guardarVarios([['proyecto', proyecto]]
      .concat(actividades.map((a) => ['actividad', a]))
      .concat(requisitos.map((r) => ['requisito', r])));
    borrador = null;
    aviso('Proyecto creado.');
    abrir(P);
  } catch (e) { aviso('No se pudo guardar: ' + e.message); }
}

/* ══════════════════════════════════════════════════════════════════════
   6) FORMULARIOS Y ACCIONES
   ══════════════════════════════════════════════════════════════════════ */
let modalOk = null;
function modal(titulo, cuerpo, textoOk, alAceptar) {
  modalOk = alAceptar;
  $('#modal').innerHTML = '<div class="modal-caja"><h2>' + esc(titulo) + '</h2>' +
    '<div class="modal-cuerpo">' + cuerpo + '</div>' +
    '<div class="modal-pie"><button class="btn ghost" onclick="cerrarModal()">Cancelar</button>' +
    '<button class="btn" onclick="modalOk()">' + esc(textoOk) + '</button></div></div>';
  $('#modal').style.display = 'flex';
}
function cerrarModal() { $('#modal').style.display = 'none'; $('#modal').innerHTML = ''; modalOk = null; }

/** Solo la lista de personas. Para lo demás está la edición completa (✎). */
function asignarDueno(id) {
  const a = DB.actividad.find((x) => x.id === id);
  if (!DB.persona.length) {
    return modal('Asignar dueño',
      '<p class="vacio">Todavía no hay nadie dado de alta. Agrega a tu equipo en la pantalla de Equipo ' +
      'y vuelve aquí.</p>', 'Entendido', cerrarModal);
  }
  const opciones = DB.persona.map((q) =>
    '<label class="opcion-persona"><input type="radio" name="dueno" value="' + esc(q.nombre) + '"' +
    (a.dueno === q.nombre ? ' checked' : '') + '>' +
    '<span class="op-nombre">' + esc(q.nombre) + '</span>' +
    (q.rol ? '<span class="op-rol">' + esc(q.rol) + '</span>' : '') + '</label>').join('');

  modal('Asignar dueño',
    '<p class="mini">' + esc(a.nombre) + '</p>' +
    '<div class="lista-personas">' + opciones +
    '<label class="opcion-persona"><input type="radio" name="dueno" value=""' +
    (!a.dueno ? ' checked' : '') + '><span class="op-nombre op-vacio">Sin dueño</span></label></div>',
    'Asignar', async () => {
      const sel = document.querySelector('input[name=dueno]:checked');
      a.dueno = sel ? sel.value : '';
      await guardar('actividad', a);
      cerrarModal(); render();
    });
}

function formActividad(id) {
  const p = DB.proyecto.find((x) => x.id === vista.proyecto);
  const a = id ? DB.actividad.find((x) => x.id === id) : null;
  const opcion = (n) => '<option' + (a && a.dueno === n ? ' selected' : '') + '>' + esc(n) + '</option>';

  modal(a ? 'Editar actividad' : 'Nueva actividad',
    campo('Actividad', '<input id="aNombre" value="' + esc(a ? a.nombre : '') + '">') +
    campo('Frente', '<select id="aFrente">' + p.frentes.map((f) =>
      '<option value="' + esc(f.clave) + '"' + (a && a.frente === f.clave ? ' selected' : '') + '>' +
      esc(f.nombre) + '</option>').join('') + '</select>') +
    campo('Dueño', '<select id="aDueno"><option value="">Sin dueño</option>' +
      DB.persona.map((q) => opcion(q.nombre)).join('') +
      (a && a.dueno && !DB.persona.some((q) => q.nombre === a.dueno) ? opcion(a.dueno) : '') +
      '</select><span class="ayuda">¿Falta alguien? Agrégalo en Equipo.</span>') +
    campo('Horas planeadas', '<input type="number" id="aHoras" min="0" value="' + (a ? a.horas : 8) + '">') +
    campo('Avance', '<select id="aAvance">' + [0, 25, 50, 75, 100].map((v) =>
      '<option value="' + v + '"' + (a && a.avance === v ? ' selected' : '') + '>' + v + '%</option>').join('') + '</select>') +
    campo('Inicio', '<input type="date" id="aIni" value="' + (a ? a.ini : hoyISO()) + '">') +
    campo('Fin', '<input type="date" id="aFin" value="' + (a ? a.fin : PlanATX.masDias(hoyISO(), 7)) + '">'),
    a ? 'Guardar' : 'Agregar', async () => {
      const nombre = $('#aNombre').value.trim();
      if (!nombre) return aviso('Falta el nombre.');
      await guardar('actividad', {
        id: a ? a.id : uid(), proyecto: p.id, nombre, frente: $('#aFrente').value,
        dueno: $('#aDueno').value, horas: parseFloat($('#aHoras').value) || 0,
        avance: parseInt($('#aAvance').value, 10), ini: $('#aIni').value, fin: $('#aFin').value,
      });
      cerrarModal(); render();
    });
}

/** Edición de los datos del proyecto, incluida la fecha de cierre acordada. */
function formProyectoEditar() {
  const p = DB.proyecto.find((x) => x.id === vista.proyecto);
  modal('Editar proyecto',
    campo('Nombre', '<input id="eNombre" value="' + esc(p.nombre) + '">') +
    campo('Cliente', '<input id="eCliente" value="' + esc(p.cliente) + '">') +
    campo('Objetivo', '<textarea id="eObjetivo">' + esc(p.objetivo || '') + '</textarea>') +
    campo('Contacto del cliente', '<input id="eContacto" value="' + esc(p.contactoCliente || '') + '">') +
    campo('Fecha de kickoff', '<input type="date" id="eKickoff" value="' + (p.kickoff || '') + '">') +
    campo('Fecha de cierre acordada', '<input type="date" id="eFin" value="' + (p.fin || '') + '">' +
      '<span class="ayuda">Es la que se compara contra la fecha estimada por el ritmo actual.</span>') +
    campo('Próxima sesión de avance', '<input type="date" id="eSesion" value="' + (p.proximaSesion || '') + '">'),
    'Guardar', async () => {
      p.nombre = $('#eNombre').value.trim() || p.nombre;
      p.cliente = $('#eCliente').value.trim() || p.cliente;
      p.objetivo = $('#eObjetivo').value.trim();
      p.contactoCliente = $('#eContacto').value.trim();
      p.kickoff = $('#eKickoff').value || p.kickoff;
      p.fin = $('#eFin').value || p.fin;
      p.proximaSesion = $('#eSesion').value;
      await guardar('proyecto', p);
      cerrarModal(); render();
    });
}

function formRequisito(id) {
  const p = DB.proyecto.find((x) => x.id === vista.proyecto);
  const r = id ? DB.requisito.find((x) => x.id === id) : null;
  const acts = actsDe(p.id).sort((a, b) => ((a.ini || '') < (b.ini || '') ? -1 : 1));
  const marcadas = r ? (r.bloquea || []) : [];

  modal(r ? 'Editar entregable del cliente' : 'Nuevo entregable del cliente',
    campo('¿Qué necesitamos del cliente?', '<input id="rTitulo" value="' + esc(r ? r.titulo : '') + '">') +
    campo('Detalle (la línea que verá en el deck)', '<textarea id="rDetalle">' + esc(r ? r.detalle : '') + '</textarea>') +
    campo('Responsable del lado del cliente', '<input id="rResp" value="' + esc(r ? r.responsable : (p.contactoCliente || '')) + '">') +
    campo('Fecha comprometida', '<input type="date" id="rFecha" value="' + (r && r.fecha ? r.fecha : '') + '">') +
    '<div class="campo"><span>¿Qué actividades no pueden cerrar sin esto?</span>' +
    '<div class="checks">' + (acts.length ? acts.map((a) =>
      '<label><input type="checkbox" class="rBloq" value="' + a.id + '"' +
      (marcadas.indexOf(a.id) >= 0 ? ' checked' : '') + '> ' + esc(a.nombre) +
      ' <i>' + fechaCorta(a.ini) + '</i></label>').join('') : '<p class="vacio">Sin actividades.</p>') +
    '</div></div>',
    r ? 'Guardar' : 'Agregar', async () => {
      const titulo = $('#rTitulo').value.trim();
      if (!titulo) return aviso('Falta el título.');
      const bloquea = Array.prototype.slice.call(document.querySelectorAll('.rBloq:checked')).map((c) => c.value);
      await guardar('requisito', { id: r ? r.id : uid(), proyecto: p.id, titulo,
        detalle: $('#rDetalle').value.trim(), responsable: $('#rResp').value.trim(),
        fecha: $('#rFecha').value, estado: r ? r.estado : 'pendiente', bloquea });
      cerrarModal(); render();
    });
}

function formPersona() {
  modal('Agregar persona',
    campo('Nombre', '<input id="pNombre">') + campo('Rol', '<input id="pRol">') +
    campo('Correo', '<input id="pCorreo">'),
    'Agregar', async () => {
      const nombre = $('#pNombre').value.trim();
      if (!nombre) return aviso('Falta el nombre.');
      await guardar('persona', { id: uid(), nombre, rol: $('#pRol').value.trim(), correo: $('#pCorreo').value.trim() });
      cerrarModal(); render();
    });
}

/** Modo local: borra lo guardado en este navegador y vuelve a sembrar. */
async function reiniciarLocal() {
  if (!confirm('Esto borra los proyectos guardados en este navegador y vuelve a cargar los de ejemplo. ¿Continuar?')) return;
  localStorage.removeItem(LKEY);
  await cargar();
  vista = { pantalla: 'general', proyecto: null, tab: 'actividades' };
  render();
  aviso('Datos de ejemplo recargados.');
}

function abrir(id) { vista = { pantalla: 'proyecto', proyecto: id, tab: 'actividades' }; render(); }
function volver() { vista.pantalla = 'general'; render(); }
function irEquipo() { vista.pantalla = 'equipo'; render(); }
function irTab(t) { vista.tab = t; render(); }

async function ciclar(id) {
  const a = DB.actividad.find((x) => x.id === id);
  const paso = { 0: 25, 25: 50, 50: 75, 75: 100, 100: 0 };
  a.avance = paso[a.avance] != null ? paso[a.avance] : 25;
  await guardar('actividad', a); render();
}
async function toggleReq(id) {
  const r = DB.requisito.find((x) => x.id === id);
  r.estado = r.estado === 'recibido' ? 'pendiente' : 'recibido';
  await guardar('requisito', r); render();
}
async function quitar(tipo, id) {
  if (!confirm('¿Eliminar este registro?')) return;
  await borrar(tipo, id); render();
}
async function agregarComentario() {
  const t = $('#nuevoCom').value.trim(); if (!t) return;
  await guardar('comentario', { id: uid(), proyecto: vista.proyecto, actividad: null,
    autor: usuario ? usuario.nombre : 'Anónimo', texto: t, fecha: hoyISO(),
    riesgo: $('#nuevoRiesgo').checked, atendido: false });
  render();
}
function verComentarios(aid) {
  const a = DB.actividad.find((x) => x.id === aid);
  const prev = comsDe(vista.proyecto, aid);
  modal('Comentarios · ' + a.nombre,
    (prev.length ? prev.map((c) =>
      '<div class="com' + (c.riesgo && !c.atendido ? ' con-riesgo' : '') + '">' +
      '<div class="com-top"><b>' + esc(c.autor) + '</b><span>' + fechaCorta(c.fecha) + '</span></div>' +
      '<p>' + (c.riesgo ? '<span class="badge-riesgo">Riesgo</span> ' : '') + esc(c.texto) + '</p>' +
      (c.riesgo && !c.atendido
        ? '<button class="link" onclick="atenderRiesgo(\'' + c.id + '\')">Marcar riesgo como atendido</button>'
        : (c.riesgo ? '<span class="mini suelto">Riesgo atendido</span>' : '')) +
      '</div>').join('')
      : '<p class="vacio">Sin comentarios.</p>') +
    '<textarea id="cTexto" placeholder="Escribe un comentario..."></textarea>' +
    '<label class="check-riesgo"><input type="checkbox" id="cRiesgo"> ' +
    'Marcar como riesgo — la actividad se pinta en rojo hasta que alguien lo marque atendido</label>',
    'Comentar', async () => {
      const txt = $('#cTexto').value.trim(); if (!txt) return cerrarModal();
      await guardar('comentario', { id: uid(), proyecto: vista.proyecto, actividad: aid,
        autor: usuario ? usuario.nombre : 'Anónimo', texto: txt, fecha: hoyISO(),
        riesgo: $('#cRiesgo').checked, atendido: false });
      cerrarModal(); render();
    });
}

async function atenderRiesgo(cid) {
  const c = DB.comentario.find((x) => x.id === cid);
  c.atendido = true;
  await guardar('comentario', c);
  cerrarModal(); render();
}

/* ── Los dos botones del deck ───────────────────────────────────────── */
function copiarCorte() {
  const p = DB.proyecto.find((x) => x.id === vista.proyecto);
  const corte = construirCorte(p);
  const texto = 'Genera el reporte de avance quincenal con la skill atxlab-avance a partir de este corte. ' +
    'Redacta tú los títulos y la redacción de trabas y pendientes; los que vienen abajo son de respaldo.\n\n' +
    '```json\n' + JSON.stringify(corte, null, 2) + '\n```';
  navigator.clipboard.writeText(texto)
    .then(() => aviso('Corte copiado. Pégalo en un chat con Claude.'))
    .catch(() => prompt('Copia esto y pégalo en Claude:', texto));
}

function descargarDeck() {
  const p = DB.proyecto.find((x) => x.id === vista.proyecto);
  const corte = construirCorte(p);
  if (!corte.frentes.length) return aviso('No hay actividades con movimiento en la quincena.');
  DeckAvance.descargar(corte)
    .then(() => aviso('Deck generado: ' + corte.archivo))
    .catch((e) => aviso('No se pudo generar: ' + e.message));
}

function aviso(t) {
  const d = $('#aviso'); d.textContent = t; d.classList.add('on');
  clearTimeout(aviso._t); aviso._t = setTimeout(() => d.classList.remove('on'), 4000);
}

/* ══════════════════════════════════════════════════════════════════════
   7) ARRANQUE
   ══════════════════════════════════════════════════════════════════════ */
async function iniciar() {
  if (CONFIG.backend === 'local') {
    usuario = { nombre: 'Nicolas Blanco', correo: '' };
    $('#quien').textContent = usuario.nombre + ' · modo local';
    await cargar(); render();
    return;
  }
  pca = new msal.PublicClientApplication({
    auth: { clientId: CONFIG.clientId, authority: 'https://login.microsoftonline.com/' + CONFIG.tenantId,
            redirectUri: location.origin + location.pathname },
    cache: { cacheLocation: 'localStorage' },
  });
  await pca.initialize();
  const cuentas = pca.getAllAccounts();
  if (!cuentas.length) return mostrarGate('sesion');
  usuario = { nombre: cuentas[0].name, correo: cuentas[0].username };
  try {
    await cargar();
    $('#gate').style.display = 'none';
    $('#quien').textContent = usuario.nombre;
    render();
    setInterval(async () => {
      if (vista.pantalla === 'general') { await cargar(); render(); }
    }, CONFIG.pollSeconds * 1000);
  } catch (e) { mostrarGate('permisos', e.message); }
}

function mostrarGate(modo, detalle) {
  const g = $('#gate'); g.style.display = 'flex';
  g.innerHTML = modo === 'sesion'
    ? '<div class="gate-caja"><h2>Inicia sesión para ver el tablero</h2>' +
      '<p>Sin sesión el tablero se ve vacío: los proyectos, el avance y los pendientes viven en SharePoint.</p>' +
      '<button class="btn" onclick="entrar()">Iniciar sesión con Microsoft</button></div>'
    : '<div class="gate-caja"><h2>Tu cuenta no tiene acceso al sitio</h2>' +
      '<p>Iniciaste sesión, pero no puedo leer la lista del tablero. Pide acceso al sitio atxlab Proyectos.</p>' +
      '<p class="detalle">' + esc(detalle || '') + '</p>' +
      '<button class="btn ghost" onclick="salir()">Cambiar de cuenta</button></div>';
}
async function entrar() { await pca.loginPopup({ scopes: ['Sites.ReadWrite.All'] }); location.reload(); }
async function salir() { await pca.logoutPopup(); }

/* ══════════════════════════════════════════════════════════════════════
   8) ARRANQUE EN BLANCO — en modo 'sharepoint' los datos vienen de la lista.
   ══════════════════════════════════════════════════════════════════════ */
function semilla() {
  return { proyecto: [], actividad: [], requisito: [], comentario: [], persona: [] };
}

window.addEventListener('DOMContentLoaded', iniciar);
