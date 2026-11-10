/* ══════════════════════════════════════════════════════════════════════
   PRODUCTOS — backlog de productos en mejora continua (p. ej. Jelpi)
   No son proyectos de cliente: no tienen semáforo ni entran a los cortes,
   al deck ni a las métricas. Viven en su propia lista de SharePoint
   (listProducts) y en sus propias colecciones: DB.producto y DB.tarea.

   Una tarea comparte la forma de una actividad de proyecto (responsable,
   fechas, horas, avance, descripción) y suma prioridad, área, tipo y fase.
   ══════════════════════════════════════════════════════════════════════ */

const PT_FASES = [
  { k: 'libre',    n: 'Por tomar',   a: 'Se asigna y se planea' },
  { k: 'curso',    n: 'En curso',    a: 'Ya tiene responsable' },
  { k: 'revision', n: 'En revisión', a: 'Esperando una segunda mirada' },
  { k: 'hecho',    n: 'Hecho',       a: 'Últimos 30 días' },
];
const PT_NOMBRE_FASE = { libre: 'Por tomar', curso: 'En curso', revision: 'En revisión', hecho: 'Hecho' };
const PT_PRI = { alta: 'Alta', media: 'Media', nice: 'Nice to have' };
const PT_PRI_ORDEN = { alta: 0, media: 1, nice: 2 };
const PT_TIPOS = { bug: 'Bug', feature: 'Feature', mejora: 'Mejora', mant: 'Mantenimiento' };
const PT_AREAS_BASE = ['Front-end', 'Backend'];

let ptFiltro = { area: 'Todas', tipo: 'Todos', pri: 'Todas', dueno: 'Todos' };
let ptOrden = 'prioridad';          // prioridad | antiguas | recientes
let ptArrastrada = null;            // id de la tarea que se está arrastrando
let ptRapida = null;                // columna que tiene abierta la captura rápida de tarea
const ptTimers = {};

/* ── acceso a datos ──────────────────────────────────────────────────── */
const ptProducto = (id) => DB.producto.find((p) => p.id === id);
const ptTareas = (pid) => DB.tarea.filter((t) => t.producto === pid);
const ptTarea = (id) => DB.tarea.find((t) => t.id === id);
const ptActivos = () => DB.producto.filter((p) => !p.archivado);
const ptYo = () => (usuario ? canon(usuario.nombre) : '');
const ptEnCurso = (t) => t.estado === 'curso' || t.estado === 'revision';
const ptAreasDe = (p) => (p.areas && p.areas.length ? p.areas : PT_AREAS_BASE);

/** El equipo para los desplegables: las personas dadas de alta, tú y quien ya esté asignado. */
function ptEquipo(extra) {
  const v = [];
  DB.persona.forEach((q) => v.push(q.nombre));
  [ptYo(), extra].forEach((n) => { if (n) v.push(n); });
  return v.filter((n, i) => n && v.indexOf(n) === i);
}

function ptGuardar(t, demora) {
  clearTimeout(ptTimers[t.id]);
  const hacer = () => guardar('tarea', t).catch((e) => aviso('No se pudo guardar: ' + e.message));
  if (demora) ptTimers[t.id] = setTimeout(hacer, demora); else hacer();
}

/** Orden manual dentro de una prioridad: posición propia, o la fecha de creación si nunca se movió. */
const ptClave = (t) => (t.pos != null ? t.pos : Date.parse((t.creada || hoyISO()) + 'T12:00:00') / 864e5);
function ptOrdenar(a, b) {
  const pri = PT_PRI_ORDEN[a.prioridad] - PT_PRI_ORDEN[b.prioridad];
  if (ptOrden === 'antiguas') return ptClave2(a) - ptClave2(b) || pri;
  if (ptOrden === 'recientes') return ptClave2(b) - ptClave2(a) || pri;
  return pri || ptClave(a) - ptClave(b);
}
// Para ordenar por antigüedad se usa solo la fecha de creación, ignorando el orden manual.
const ptClave2 = (t) => Date.parse((t.creada || hoyISO()) + 'T12:00:00') / 864e5;

function ptGrupo(t) {
  return DB.tarea.filter((x) => x.producto === t.producto && x.estado === 'libre' && x.prioridad === t.prioridad)
    .sort((a, b) => ptClave(a) - ptClave(b));
}
/** Coloca t junto a ref. Escribe solo la posición de t: el promedio entre sus dos vecinas. */
function ptReubicar(t, ref, antes) {
  const g = ptGrupo(t).filter((x) => x !== t);
  const i = g.indexOf(ref); if (i < 0) return;
  const ant = antes ? g[i - 1] : g[i], sig = antes ? g[i] : g[i + 1];
  if (ant && sig && Math.abs(ptClave(sig) - ptClave(ant)) < 0.0001) {   // muy juntas: se renumera el grupo una vez
    const base = ptClave(g[0]);
    g.forEach((x, n) => { x.pos = base + n; ptGuardar(x, 300); });
  }
  if (ant && sig) t.pos = (ptClave(ant) + ptClave(sig)) / 2;
  else if (sig) t.pos = ptClave(sig) - 1;
  else if (ant) t.pos = ptClave(ant) + 1;
  ptGuardar(t, 300);
}
function ptPaso(t, delta) {
  const g = ptGrupo(t), j = g.indexOf(t) + delta;
  if (j < 0 || j >= g.length) return false;
  ptReubicar(t, g[j], delta < 0);
  return true;
}

/** Reglas al cambiar de fase: en curso, en revisión y hecho siempre tienen responsable. */
function ptMover(t, k) {
  if (t.estado === k) return '';
  let msg = '"' + t.titulo + '" pasó a ' + PT_NOMBRE_FASE[k] + '.';
  if (t.estado === 'hecho') delete t.cerrada;
  if (k !== 'libre' && !t.dueno) { t.dueno = ptYo(); msg += ' Quedó a tu nombre.'; }
  if (k === 'hecho') { t.cerrada = hoyISO(); t.avance = 100; }
  else if (k === 'curso' || k === 'revision') t.avance = t.avance || 0;
  t.estado = k;
  return msg;
}

/* ══════════════════════════════════════════════════════════════════════
   BARRA LATERAL
   ══════════════════════════════════════════════════════════════════════ */
function sidenavProductos() {
  if (PROD_ERROR) {
    return '<div class="grp-label">Productos</div>' +
      '<div class="pt-nav-aviso">No pude leer la lista “' + esc(CONFIG.listaProductos) + '”. Los proyectos siguen igual.</div>';
  }
  const item = (p) => '<button class="nav-item' + (vista.pantalla === 'producto' && vista.producto === p.id ? ' on' : '') +
    '" onclick="irProducto(\'' + p.id + '\')"><span class="ic">◆</span>' + esc(p.nombre) +
    '<span class="count">' + ptTareas(p.id).filter((t) => t.estado === 'libre').length + '</span></button>';
  const arch = DB.producto.filter((p) => p.archivado);
  return '<div class="grp-label">Productos</div>' +
    ptActivos().map(item).join('') +
    '<button class="nav-item" onclick="formProducto()"><span class="ic">＋</span>Nuevo producto</button>' +
    (CONFIG.backend === 'local' && !DB.producto.length
      ? '<button class="nav-item" onclick="ptCargarEjemplo()" title="Solo en modo local: crea un producto de muestra"><span class="ic">✦</span>Cargar ejemplo</button>' : '') +
    (arch.length ? '<div class="grp-label">Archivados</div>' + arch.map(item).join('') : '');
}

function irProducto(id) {
  ptFiltro = { area: 'Todas', tipo: 'Todos', pri: 'Todas', dueno: 'Todos' }; ptRapida = null;
  Object.assign(vista, { pantalla: 'producto', producto: id });
  render();
}

/** El refresco automático no debe interrumpir un arrastre ni un formulario abierto. */
function productoLibreParaRefrescar() {
  return !ptArrastrada && !ptRapida && $('#modal').style.display !== 'flex';
}

/* ══════════════════════════════════════════════════════════════════════
   PANTALLA DE UN PRODUCTO
   ══════════════════════════════════════════════════════════════════════ */
function ptSelect(id, etq, opciones, actual, alCambiar) {
  return '<label class="pt-filtro"><span>' + etq + '</span><select class="pt-sel" id="' + id + '" onchange="' + alCambiar + '">' +
    opciones.map((o) => '<option value="' + esc(o[0]) + '"' + (String(o[0]) === String(actual) ? ' selected' : '') + '>' + esc(o[1]) + '</option>').join('') +
    '</select></label>';
}

function ptVisible(t) {
  if (ptFiltro.area !== 'Todas' && t.area !== ptFiltro.area) return false;
  if (ptFiltro.tipo !== 'Todos' && (t.tipo || 'mejora') !== ptFiltro.tipo) return false;
  if (ptFiltro.pri !== 'Todas' && t.prioridad !== ptFiltro.pri) return false;
  if (ptFiltro.dueno === '__nadie') return !t.dueno;
  if (ptFiltro.dueno !== 'Todos' && t.dueno !== ptFiltro.dueno) return false;
  return true;
}
function ptPonerFiltro(k, v) { ptFiltro[k] = v; render(); }
function ptPonerOrden(v) { ptOrden = v; render(); }

function renderProducto() {
  const p = ptProducto(vista.producto);
  if (!p) { vista.pantalla = 'general'; return render(); }
  const ro = !!p.archivado;
  const todas = ptTareas(p.id).filter((t) => t.estado !== 'descartada');
  const descartadas = ptTareas(p.id).filter((t) => t.estado === 'descartada');
  const abiertas = todas.filter((t) => t.estado !== 'hecho').length;
  const hoy = hoyISO();

  const duenos = ['Todos', '__nadie'].concat(ptEquipo()).filter((n, i, a) => a.indexOf(n) === i);
  const filtros =
    '<div class="pt-barra">' +
    ptSelect('ptfArea', 'Área', [['Todas', 'Todas']].concat(ptAreasDe(p).map((a) => [a, a])), ptFiltro.area, "ptPonerFiltro('area',this.value)") +
    ptSelect('ptfTipo', 'Tipo', [['Todos', 'Todos']].concat(Object.keys(PT_TIPOS).map((k) => [k, PT_TIPOS[k]])), ptFiltro.tipo, "ptPonerFiltro('tipo',this.value)") +
    ptSelect('ptfPri', 'Prioridad', [['Todas', 'Todas']].concat(Object.keys(PT_PRI).map((k) => [k, PT_PRI[k]])), ptFiltro.pri, "ptPonerFiltro('pri',this.value)") +
    ptSelect('ptfDueno', 'Responsable', duenos.map((n) => [n, n === 'Todos' ? 'Todos' : (n === '__nadie' ? 'Sin responsable' : n)]), ptFiltro.dueno, "ptPonerFiltro('dueno',this.value)") +
    '<span class="pt-espacio"></span>' +
    ptSelect('ptOrden', 'Ordenar', [['prioridad', 'Por prioridad'], ['antiguas', 'Más antiguas primero'], ['recientes', 'Más recientes primero']], ptOrden, 'ptPonerOrden(this.value)') +
    '</div>';

  let tablero;
  if (!todas.length) {
    tablero = '<div class="card pt-vacio"><b>' + esc(p.nombre) + ' todavía no tiene tareas</b>' +
      '<p>Agrega ideas, mejoras o pendientes aunque sean chicos; quien tenga tiempo las toma.</p>' +
      (ro ? '' : '<button class="btn" onclick="formTarea()">＋ Agregar la primera tarea</button>') + '</div>';
  } else {
    tablero = '<div class="pt-tablero">' + PT_FASES.map((col) => {
      const todasCol = todas.filter((t) => t.estado === col.k && (col.k !== 'hecho' || (t.cerrada || '0') >= PlanATX.masDias(hoy, -30)));
      const lista = todasCol.filter(ptVisible);
      if (col.k === 'hecho') lista.sort((a, b) => ((a.cerrada || '') < (b.cerrada || '') ? 1 : -1));
      else lista.sort(ptOrdenar);
      const ayuda = col.k === 'libre' && !ro
        ? (ptOrden === 'prioridad' ? 'Arrastra sobre otra del mismo nivel para reordenar' : 'Para reordenar a mano, ordena por prioridad')
        : col.a;
      return '<div class="pt-col" data-col="' + col.k + '"' + (ro ? '' :
        ' ondragover="ptSobreCol(event,this)" ondragleave="ptFueraCol(event,this)" ondrop="ptSoltarCol(event,this,\'' + col.k + '\')"') + '>' +
        '<div class="pt-col-cab"><i class="pt-dot pt-dot-' + col.k + '"></i>' + col.n + '<span class="pt-n">' + lista.length + '</span></div>' +
        '<div class="pt-col-ayuda">' + ayuda + '</div>' +
        (lista.length ? lista.map((t) => ptTarjeta(t, ro, hoy)).join('')
          : '<div class="pt-col-vacia">' + (todasCol.length ? 'Nada con estos filtros' : 'Sin tareas') + '</div>') +
        (ro ? '' : ptRapidaHtml(p, col.k)) +
        '</div>';
    }).join('') + '</div>';
  }

  const arch = ro
    ? '<button class="btn ghost" onclick="ptRestaurarProducto()">Restaurar producto</button>'
    : '<button class="btn ghost" onclick="formProducto(\'' + p.id + '\')">Editar nombre y áreas</button>' +
      '<button class="btn ghost peligro" onclick="ptConfirmarArchivar()">Archivar producto</button>';

  app().innerHTML =
    '<div class="head pt-head"><div><button class="link" onclick="volver()">← Proyectos</button>' +
    '<h1>' + esc(p.nombre) + '<span class="sub">Producto · ' + abiertas + ' abiertas</span></h1>' +
    (p.descripcion ? '<p class="sub-h1">' + esc(p.descripcion) + '</p>' : '') + '</div>' +
    '<div class="acciones">' + (ro ? '' : '<button class="btn" onclick="formTarea()">＋ Tarea</button>') + '</div></div>' +
    (ro ? '<div class="nota">Este producto está archivado: se ve en modo de solo lectura. Restaurarlo lo devuelve al sidebar.</div>' : '') +
    (todas.length ? filtros : '') + tablero +
    '<p class="pt-pie">Sin semáforo ni reportes: este backlog no entra a los cortes de proyectos. Cada tarea lleva su avance, ' +
    'y desde “Por tomar” ya se puede asignar y planear. Las horas son una estimación opcional.' +
    (descartadas.length ? ' <button class="link" onclick="ptVerDescartadas()">Ver descartadas (' + descartadas.length + ')</button>' : '') + '</p>' +
    '<div class="pt-zona"><div><b>Administrar producto</b><span>' +
    (ro ? 'Archivado. No se borró nada.' : 'Archivarlo lo saca del sidebar y deja su tablero de solo lectura. No se borra nada y se puede restaurar.') +
    '</span></div><div class="pt-zona-btns">' + arch + '</div></div>';
  if (ptRapida) { const c = $('#ptRapidaTxt'); if (c) c.focus(); }
}

/* ── captura rápida: agregar una tarea directo en una columna ────────── */
function ptRapidaHtml(p, fase) {
  if (ptRapida !== fase) {
    return '<button class="pt-agregar" onclick="ptAbrirRapida(\'' + fase + '\')">＋ Agregar tarea</button>';
  }
  const pri = ptFiltro.pri !== 'Todas' ? ptFiltro.pri : 'media';
  return '<div class="pt-rapida">' +
    '<input id="ptRapidaTxt" maxlength="90" placeholder="Título de la tarea" aria-label="Título de la nueva tarea" onkeydown="ptRapidaTecla(event,\'' + fase + '\')">' +
    '<div class="pt-rapida-fila"><select class="pt-sel pt-sel-mini" id="ptRapidaPri" aria-label="Prioridad">' + ptOpts(PT_OPC_PRI, pri) + '</select>' +
    '<button class="btn pt-tomar" onclick="ptCrearRapida(\'' + fase + '\')">Agregar</button>' +
    '<button class="btn ghost pt-tomar" onclick="ptCerrarRapida()">Cancelar</button></div>' +
    '<span class="mini">Enter agrega y deja abierta la captura. Esc cancela. El resto se completa con el lápiz.</span></div>';
}
function ptAbrirRapida(fase) { ptRapida = fase; render(); }
function ptCerrarRapida() { ptRapida = null; render(); }
function ptRapidaTecla(e, fase) {
  if (e.key === 'Enter') { e.preventDefault(); ptCrearRapida(fase); }
  else if (e.key === 'Escape') ptCerrarRapida();
}
function ptCrearRapida(fase) {
  const p = ptProducto(vista.producto), inp = $('#ptRapidaTxt'); if (!p || !inp) return;
  const titulo = inp.value.trim();
  if (!titulo) return inp.focus();
  const t = { id: uid(), producto: p.id, titulo, descripcion: '', estado: 'libre', avance: 0,
    // Si hay filtros puestos, la tarea nace ya dentro de ellos para que no desaparezca del tablero.
    tipo: ptFiltro.tipo !== 'Todos' ? ptFiltro.tipo : 'mejora',
    area: ptFiltro.area !== 'Todas' ? ptFiltro.area : ptAreasDe(p)[0],
    prioridad: $('#ptRapidaPri').value, creadaPor: usuario ? usuario.nombre : 'Anónimo',
    creada: hoyISO(), pos: Date.now() / 864e5 };
  if (ptFiltro.dueno !== 'Todos' && ptFiltro.dueno !== '__nadie') t.dueno = ptFiltro.dueno;
  ptMover(t, fase);                      // en curso, revisión y hecho quedan a nombre de quien la agrega
  DB.tarea.push(t); ptGuardar(t);
  render(); aviso('Tarea agregada a ' + PT_NOMBRE_FASE[fase] + '.');
}

/* ── tarjeta ─────────────────────────────────────────────────────────── */
function ptTarjeta(t, ro, hoy) {
  const ti = t.tipo || 'mejora';
  const dueno = t.dueno || '';
  const tarde = t.fin && t.fin < hoy && t.estado !== 'hecho';
  const fechas = t.ini || t.fin
    ? fechaCorta(t.ini || t.fin) + (t.ini && t.fin ? ' → ' + fechaCorta(t.fin) : '')
    : '＋ Fechas';
  const opciones = ptEquipo(dueno);
  const pie = ro
    ? '<span class="pt-resp">' + esc(dueno || 'Sin responsable') + '</span>'
    : '<select class="pt-sel pt-sel-mini" aria-label="Responsable" onchange="ptAsignar(\'' + t.id + '\',this.value)">' +
      '<option value="">Sin responsable</option>' +
      opciones.map((n) => '<option' + (n === dueno ? ' selected' : '') + '>' + esc(n) + '</option>').join('') + '</select>' +
      (t.estado === 'libre' ? '<button class="btn pt-tomar" onclick="ptTomar(\'' + t.id + '\')">' + (dueno ? 'Iniciar' : 'Tomar') + '</button>' : '');
  const reubicable = !ro && t.estado === 'libre';
  return '<div class="pt' + (dueno && dueno === ptYo() ? ' mia' : '') + (t.estado === 'hecho' ? ' hecha' : '') + '"' +
    (ro ? '' : ' draggable="true" ondragstart="ptInicioArrastre(event,\'' + t.id + '\')" ondragend="ptFinArrastre()"') +
    (reubicable ? ' ondragover="ptSobreTarjeta(event,this,\'' + t.id + '\')" ondragleave="ptFueraTarjeta(this)" ondrop="ptSoltarTarjeta(event,this,\'' + t.id + '\')"' : '') + '>' +
    '<div class="pt-top"><div class="pt-tit"' + (t.descripcion ? ' title="' + esc(t.descripcion) + '"' : '') + '>' + esc(t.titulo) + '</div>' +
    (ro ? '' : '<button class="editar" onclick="formEditarTarea(\'' + t.id + '\')" aria-label="Editar tarea" title="Editar">✎</button>') + '</div>' +
    '<div class="pt-etqs">' +
    (t.prioridad === 'alta' ? '<span class="pt-etq alta">Alta</span>' : '') +
    (t.prioridad === 'nice' ? '<span class="pt-etq dim">Nice to have</span>' : '') +
    '<span class="pt-etq' + (ti === 'bug' ? ' bug' : '') + '">' + esc(PT_TIPOS[ti]) + '</span>' +
    '<span class="pt-etq">' + esc(t.area) + '</span>' +
    (t.horas ? '<span class="pt-horas">~' + t.horas + ' h</span>' : '') + '</div>' +
    (ro ? (t.ini || t.fin ? '<div class="pt-fechas">' + fechas + '</div>' : '')
        : '<button class="pt-fechas' + (tarde ? ' tarde' : '') + (t.ini || t.fin ? '' : ' vacia') + '" onclick="formFechasTarea(\'' + t.id + '\')" title="Planear fechas">' + fechas +
          (tarde ? ' · vencida' : '') + '</button>') +
    '<button class="pt-prog" ' + (ro || t.estado === 'hecho' ? 'disabled ' : '') + 'onclick="ptAvance(\'' + t.id + '\')" title="Clic para subir 25%; después de 100% vuelve a 0%">' +
    '<span class="barra"><span class="barra-fill" style="width:' + (t.avance || 0) + '%"></span></span><b>' + (t.avance || 0) + '%</b></button>' +
    '<div class="pt-pie-card">' + pie + '</div></div>';
}

/* ── acciones directas en la tarjeta ─────────────────────────────────── */
function ptAvance(id) {
  const t = ptTarea(id); if (!t || t.estado === 'hecho') return;
  t.avance = (t.avance || 0) >= 100 ? 0 : Math.min(100, (t.avance || 0) + 25);
  ptGuardar(t, 400); render();
}
function ptAsignar(id, nombre) {
  const t = ptTarea(id); if (!t) return;
  if (!nombre && ptEnCurso(t)) { aviso('Una tarea en curso o en revisión necesita responsable.'); return render(); }
  if (nombre) t.dueno = nombre; else delete t.dueno;
  ptGuardar(t); render();
}
function ptTomar(id) {
  const t = ptTarea(id); if (!t) return;
  const msg = ptMover(t, 'curso'); ptGuardar(t); render(); if (msg) aviso(msg);
}

/* ── arrastrar y soltar ──────────────────────────────────────────────── */
function ptInicioArrastre(e, id) {
  ptArrastrada = id; e.currentTarget.classList.add('arrastrando');
  e.dataTransfer.effectAllowed = 'move';
  try { e.dataTransfer.setData('text/plain', id); } catch (x) { /* algunos navegadores lo exigen, otros no */ }
}
function ptFinArrastre() {
  ptArrastrada = null;
  document.querySelectorAll('.pt.arrastrando').forEach((x) => x.classList.remove('arrastrando'));
  document.querySelectorAll('.pt-col.sobre').forEach((x) => x.classList.remove('sobre'));
  document.querySelectorAll('.pt.ins-arriba,.pt.ins-abajo').forEach((x) => x.classList.remove('ins-arriba', 'ins-abajo'));
}
function ptSobreCol(e, col) {
  if (!ptArrastrada) return;
  e.preventDefault(); e.dataTransfer.dropEffect = 'move'; col.classList.add('sobre');
}
function ptFueraCol(e, col) { if (!col.contains(e.relatedTarget)) col.classList.remove('sobre'); }
function ptSoltarCol(e, col, fase) {
  e.preventDefault(); col.classList.remove('sobre');
  const t = ptArrastrada && ptTarea(ptArrastrada); ptArrastrada = null;
  if (!t) return;
  let msg = ptMover(t, fase);
  if (!msg && fase === 'libre' && ptOrden === 'prioridad' && e.target.closest && e.target.closest('.pt')) {
    msg = 'Solo se reordena entre tareas del mismo nivel de prioridad. Para cambiarla, usa el lápiz.';
  }
  if (t.estado === fase) ptGuardar(t);
  render(); if (msg) aviso(msg);
}
// Reordenar: arrastrar sobre otra tarjeta de Por tomar con la misma prioridad (mitad superior: antes; inferior: después).
function ptPuedeReubicar(id) {
  const a = ptArrastrada && ptTarea(ptArrastrada), t = ptTarea(id);
  return ptOrden === 'prioridad' && a && t && a !== t && a.estado === 'libre' && t.estado === 'libre' && a.prioridad === t.prioridad;
}
const ptMitadSuperior = (e, el) => { const r = el.getBoundingClientRect(); return e.clientY < r.top + r.height / 2; };
function ptSobreTarjeta(e, el, id) {
  if (!ptPuedeReubicar(id)) return;
  e.preventDefault(); e.stopPropagation(); e.dataTransfer.dropEffect = 'move';
  const arriba = ptMitadSuperior(e, el);
  el.classList.toggle('ins-arriba', arriba); el.classList.toggle('ins-abajo', !arriba);
}
function ptFueraTarjeta(el) { el.classList.remove('ins-arriba', 'ins-abajo'); }
function ptSoltarTarjeta(e, el, id) {
  if (!ptPuedeReubicar(id)) return;
  e.preventDefault(); e.stopPropagation();
  const arriba = ptMitadSuperior(e, el), mov = ptTarea(ptArrastrada), ref = ptTarea(id); ptArrastrada = null;
  ptReubicar(mov, ref, arriba);
  render(); aviso('"' + mov.titulo + '" quedó ' + (arriba ? 'antes' : 'después') + ' de "' + ref.titulo + '".');
}

/* ══════════════════════════════════════════════════════════════════════
   FORMULARIOS
   ══════════════════════════════════════════════════════════════════════ */
const ptOpts = (lista, actual) => lista.map((o) =>
  '<option value="' + esc(o[0]) + '"' + (String(o[0]) === String(actual) ? ' selected' : '') + '>' + esc(o[1]) + '</option>').join('');
const PT_OPC_TIPO = Object.keys(PT_TIPOS).map((k) => [k, PT_TIPOS[k]]);
const PT_OPC_PRI = Object.keys(PT_PRI).map((k) => [k, PT_PRI[k]]);

function ptCamposToma(t) {
  const equipo = ptEquipo(t && t.dueno);
  return campo('Responsable (opcional)', '<select id="ptResp"><option value="">Sin responsable</option>' +
      equipo.map((n) => '<option' + (t && t.dueno === n ? ' selected' : '') + '>' + esc(n) + '</option>').join('') + '</select>') +
    '<div class="pt-fila2">' + campo('Inicio', '<input type="date" id="ptIni" value="' + esc(t && t.ini || '') + '">') +
    campo('Fin', '<input type="date" id="ptFin" value="' + esc(t && t.fin || '') + '">') + '</div>' +
    campo('Horas estimadas (opcional)', '<input type="number" id="ptHoras" min="0" step="0.5" value="' + (t && t.horas != null ? t.horas : '') + '">');
}
/** Lee responsable, fechas y horas. Devuelve un texto de error o null. */
function ptLeerToma(dest) {
  const ini = $('#ptIni').value, fin = $('#ptFin').value, h = $('#ptHoras').value.trim();
  if (ini && fin && fin < ini) return 'La fecha de fin no puede ser antes que la de inicio.';
  if (h !== '' && !(parseFloat(h) >= 0)) return 'Las horas deben ser un número.';
  dest.ini = ini || ''; dest.fin = fin || ''; dest.horas = h === '' ? null : parseFloat(h);
  const r = $('#ptResp').value; if (r) dest.dueno = r; else delete dest.dueno;
  return null;
}

function formTarea() {
  const p = ptProducto(vista.producto); if (!p) return;
  modal('Nueva tarea · ' + p.nombre,
    campo('Título', '<input id="ptTitulo" maxlength="90" placeholder="Qué hay que hacer">') +
    campo('Descripción (opcional)', '<textarea id="ptDesc" maxlength="400" placeholder="Contexto, criterio de listo, enlaces…"></textarea>') +
    '<div class="pt-fila2">' + campo('Tipo', '<select id="ptTipo">' + ptOpts(PT_OPC_TIPO, 'feature') + '</select>') +
    campo('Área', '<select id="ptArea">' + ptOpts(ptAreasDe(p).map((a) => [a, a]), '') + '</select>') + '</div>' +
    campo('Prioridad', '<select id="ptPri">' + ptOpts(PT_OPC_PRI, 'media') + '</select>') +
    ptCamposToma(null) +
    '<p class="mini">Entra a “Por tomar”. El responsable y las fechas se pueden poner ahora o después.</p>',
    'Agregar tarea', () => {
      const titulo = $('#ptTitulo').value.trim();
      if (!titulo) return aviso('Falta el título.');
      const t = { id: uid(), producto: p.id, titulo, descripcion: $('#ptDesc').value.trim(), tipo: $('#ptTipo').value,
        area: $('#ptArea').value, prioridad: $('#ptPri').value, estado: 'libre', avance: 0,
        creadaPor: usuario ? usuario.nombre : 'Anónimo', creada: hoyISO(), pos: Date.now() / 864e5 };
      const err = ptLeerToma(t); if (err) return aviso(err);
      DB.tarea.push(t); ptGuardar(t); cerrarModal(); render();
    });
}

function formEditarTarea(id) {
  const t = ptTarea(id), p = t && ptProducto(t.producto); if (!t || !p) return;
  const hecha = t.estado === 'hecho';
  const fases = PT_FASES.map((c) => [c.k, c.n]);
  modal('Editar tarea',
    (t.creadaPor ? '<p class="mini">Agregada por ' + esc(t.creadaPor) + (t.creada ? ' · ' + fechaCorta(t.creada) : '') + '</p>' : '') +
    campo('Título', '<input id="ptTitulo" maxlength="90" value="' + esc(t.titulo) + '">') +
    campo('Descripción', '<textarea id="ptDesc" maxlength="400">' + esc(t.descripcion || '') + '</textarea>') +
    '<div class="pt-fila2">' + campo('Tipo', '<select id="ptTipo">' + ptOpts(PT_OPC_TIPO, t.tipo || 'mejora') + '</select>') +
    campo('Área', '<select id="ptArea">' + ptOpts(ptAreasDe(p).map((a) => [a, a]).concat(ptAreasDe(p).indexOf(t.area) < 0 ? [[t.area, t.area]] : []), t.area) + '</select>') + '</div>' +
    '<div class="pt-fila2">' + campo('Prioridad', '<select id="ptPri">' + ptOpts(PT_OPC_PRI, t.prioridad) + '</select>') +
    campo('Fase', '<select id="ptFase">' + ptOpts(fases, t.estado) + '</select>') + '</div>' +
    ptCamposToma(t) +
    (hecha ? '' : campo('Avance', '<select id="ptAv">' + ptOpts([0, 25, 50, 75, 100].map((v) => [v, v + '%']), t.avance || 0) + '</select>')) +
    (t.estado === 'libre'
      ? '<div class="campo"><span>Lugar entre las tareas de su prioridad</span><div class="pt-mover">' +
        '<button type="button" class="btn ghost" id="ptSube" onclick="ptMoverLugar(\'' + t.id + '\',-1)">↑ Subir</button>' +
        '<button type="button" class="btn ghost" id="ptBaja" onclick="ptMoverLugar(\'' + t.id + '\',1)">↓ Bajar</button>' +
        '<span class="mini" id="ptLugar"></span></div></div>' : '') +
    '<p class="mini">Renombrarla no rompe su historia: la tarea conserva su identidad y guarda el nombre anterior.</p>' +
    '<button type="button" class="btn ghost peligro" onclick="cerrarModal();ptDescartar(\'' + t.id + '\')">Descartar tarea…</button>',
    'Guardar cambios', () => {
      const titulo = $('#ptTitulo').value.trim();
      if (!titulo) return aviso('El título no puede quedar vacío.');
      const copia = {}, err = ptLeerToma(copia); if (err) return aviso(err);
      const fase = $('#ptFase').value;
      if (fase !== 'libre' && !copia.dueno) copia.dueno = t.dueno || ptYo();   // en curso/revisión/hecho siempre tiene responsable
      if (titulo !== t.titulo) { t.antes = (t.antes || []).concat(t.titulo); t.titulo = titulo; }
      t.ini = copia.ini; t.fin = copia.fin; t.horas = copia.horas;
      if (copia.dueno) t.dueno = copia.dueno; else delete t.dueno;
      t.descripcion = $('#ptDesc').value.trim(); t.tipo = $('#ptTipo').value; t.area = $('#ptArea').value;
      t.prioridad = $('#ptPri').value;
      if (!hecha) t.avance = parseInt($('#ptAv').value, 10);
      const msg = ptMover(t, fase);
      ptGuardar(t); cerrarModal(); render(); if (msg) aviso(msg);
    });
  if (t.estado === 'libre') ptActualizarLugar(t);
}
function ptActualizarLugar(t) {
  const g = ptGrupo(t), i = g.indexOf(t);
  $('#ptLugar').textContent = 'Lugar ' + (i + 1) + ' de ' + g.length +
    (ptOrden === 'prioridad' ? '' : ' · el tablero está ordenado por antigüedad; cámbialo a “Por prioridad” para verlo');
  $('#ptSube').disabled = i === 0; $('#ptBaja').disabled = i === g.length - 1;
}
function ptMoverLugar(id, delta) {
  const t = ptTarea(id); if (!t) return;
  if (ptPaso(t, delta)) { render(); ptActualizarLugar(t); }
}

function formFechasTarea(id) {
  const t = ptTarea(id); if (!t) return;
  modal('Fechas de la tarea',
    '<p class="mini">' + esc(t.titulo) + '</p>' +
    '<div class="pt-fila2">' + campo('Inicio', '<input type="date" id="ptIni" value="' + esc(t.ini || '') + '">') +
    campo('Fin', '<input type="date" id="ptFin" value="' + esc(t.fin || '') + '">') + '</div>' +
    '<p class="mini">Son opcionales y se pueden cambiar cuando quieras.</p>',
    'Guardar fechas', () => {
      const ini = $('#ptIni').value, fin = $('#ptFin').value;
      if (ini && fin && fin < ini) return aviso('La fecha de fin no puede ser antes que la de inicio.');
      t.ini = ini; t.fin = fin; ptGuardar(t); cerrarModal(); render();
    });
}

function ptDescartar(id) {
  const t = ptTarea(id); if (!t) return;
  modal('Descartar tarea',
    '<p class="mini">' + esc(t.titulo) + '</p>' +
    campo('Motivo (opcional)', '<textarea id="ptMotivo" maxlength="300" placeholder="Por qué ya no se hará"></textarea>') +
    '<p class="mini">Sale del tablero pero no se borra: queda en “Ver descartadas” y se puede restaurar.</p>',
    'Descartar', () => {
      t.estado = 'descartada'; t.descartada = hoyISO(); t.motivo = $('#ptMotivo').value.trim();
      ptGuardar(t); cerrarModal(); render(); aviso('"' + t.titulo + '" se descartó.');
    });
}
function ptVerDescartadas() {
  const p = ptProducto(vista.producto); if (!p) return;
  const lista = ptTareas(p.id).filter((t) => t.estado === 'descartada');
  modal('Tareas descartadas',
    '<div class="pt-desc">' + (lista.map((t) => '<div class="act"><div class="act-cuerpo"><div class="act-tit">' + esc(t.titulo) + '</div>' +
      '<div class="act-meta"><span>' + esc(t.motivo || 'Sin motivo') + '</span>' +
      (t.descartada ? '<span>' + fechaCorta(t.descartada) + '</span>' : '') + '</div></div>' +
      (p.archivado ? '' : '<button class="btn ghost" onclick="ptRestaurarTarea(\'' + t.id + '\')">Restaurar</button>') + '</div>').join('') ||
    '<p class="vacio">No hay tareas descartadas.</p>') + '</div>',
    'Cerrar', cerrarModal);
}
function ptRestaurarTarea(id) {
  const t = ptTarea(id); if (!t) return;
  t.estado = 'libre'; delete t.descartada; delete t.motivo;
  ptGuardar(t); cerrarModal(); render(); aviso('"' + t.titulo + '" volvió a Por tomar.');
}

/* ── productos ───────────────────────────────────────────────────────── */
function formProducto(id) {
  const p = id ? ptProducto(id) : null;
  modal(p ? 'Editar producto' : 'Nuevo producto',
    '<p class="nota">Un producto es algo que mejoramos de forma continua, sin cliente ni fecha de cierre. ' +
    'Su backlog es aparte de los proyectos y no entra a sus reportes.</p>' +
    campo('Nombre', '<input id="ppNombre" maxlength="60" value="' + esc(p ? p.nombre : '') + '" placeholder="Ej. Jelpi">') +
    campo('Descripción (opcional)', '<textarea id="ppDesc" maxlength="300">' + esc(p ? p.descripcion || '' : '') + '</textarea>') +
    campo('Áreas', '<input id="ppAreas" value="' + esc(p ? ptAreasDe(p).join(', ') : PT_AREAS_BASE.join(', ')) + '">' +
      '<span class="mini">Separadas por coma. Cada tarea se clasifica en una.</span>'),
    p ? 'Guardar' : 'Crear producto', () => {
      const nombre = $('#ppNombre').value.trim();
      if (!nombre) return aviso('Falta el nombre.');
      if (DB.producto.some((x) => x.id !== (p && p.id) && x.nombre.toLowerCase() === nombre.toLowerCase())) return aviso('Ya existe un producto con ese nombre.');
      const areas = $('#ppAreas').value.split(',').map((a) => a.trim()).filter((a, i, v) => a && v.indexOf(a) === i);
      const nuevo = p || { id: uid(), creado: hoyISO(), archivado: false };
      nuevo.nombre = nombre; nuevo.descripcion = $('#ppDesc').value.trim(); nuevo.areas = areas.length ? areas : PT_AREAS_BASE.slice();
      guardar('producto', nuevo).catch((e) => aviso('No se pudo guardar: ' + e.message));
      cerrarModal(); if (p) render(); else irProducto(nuevo.id);
    });
}
function ptConfirmarArchivar() {
  const p = ptProducto(vista.producto); if (!p) return;
  const n = ptTareas(p.id).filter((t) => t.estado !== 'descartada' && t.estado !== 'hecho').length;
  modal('¿Archivar “' + p.nombre + '”?',
    '<p class="nota">Sale del sidebar y su tablero queda de solo lectura. Sus ' + n + ' tareas abiertas y todo el historial se conservan, ' +
    'y puedes restaurarlo cuando quieras desde “Archivados”.</p>',
    'Archivar producto', () => {
      p.archivado = true; guardar('producto', p).catch((e) => aviso('No se pudo guardar: ' + e.message));
      cerrarModal(); const otro = ptActivos()[0];
      if (otro) irProducto(otro.id); else volver();
      aviso('"' + p.nombre + '" se archivó. Está en Archivados, en el sidebar.');
    });
}
function ptRestaurarProducto() {
  const p = ptProducto(vista.producto); if (!p) return;
  p.archivado = false; guardar('producto', p).catch((e) => aviso('No se pudo guardar: ' + e.message));
  render(); aviso('"' + p.nombre + '" volvió al sidebar.');
}

/* Solo modo local: crea un producto de muestra (Jelpi) para ver el tablero con datos. */
function ptCargarEjemplo() {
  const hoy = hoyISO(), d = (n) => PlanATX.masDias(hoy, n);
  const p = { id: uid(), nombre: 'Jelpi', descripcion: 'Dashboard de lo que hacen nuestros agentes. Crece de forma continua, sin cliente.',
    areas: PT_AREAS_BASE.slice(), archivado: false, creado: d(-120) };
  const T = (titulo, area, pri, estado, o) => DB.tarea.push(Object.assign({
    id: uid(), producto: p.id, titulo, descripcion: '', tipo: 'mejora', area, prioridad: pri, estado, avance: 0,
    creadaPor: 'Nicolas Blanco', creada: d(-30), pos: null }, o || {}));
  const yo = ptYo() || 'Nicolas Blanco';
  T('Exportar la actividad a PDF', 'Front-end', 'media', 'hecho', { avance: 100, dueno: 'Nicolas Blanco', cerrada: d(-4), creada: d(-40), tipo: 'feature' });
  T('Caché de la lista de agentes', 'Backend', 'alta', 'hecho', { avance: 100, dueno: yo, cerrada: d(-11), creada: d(-50), tipo: 'mejora' });
  T('Filtro por rango de fechas', 'Front-end', 'media', 'hecho', { avance: 100, dueno: 'Nicolas Blanco', cerrada: d(-19), creada: d(-60), tipo: 'feature' });
  T('Gráfica de costos por agente', 'Front-end', 'alta', 'curso', { avance: 50, dueno: 'Nicolas Blanco', ini: d(-6), fin: d(8), horas: 6, tipo: 'feature', creada: d(-25) });
  T('Migrar el historial de runs a paginación', 'Backend', 'alta', 'revision', { avance: 75, dueno: yo, ini: d(-15), fin: d(2), horas: 4, tipo: 'mejora', creada: d(-31) });
  T('Filtro por agente en la vista de actividad', 'Front-end', 'alta', 'libre', { horas: 3, dueno: 'Nicolas Blanco', ini: d(1), fin: d(4), tipo: 'feature', creada: d(-12) });
  T('Alerta cuando un agente falla 3 veces seguidas', 'Backend', 'alta', 'libre', { horas: 8, tipo: 'feature', creada: d(-20), descripcion: 'Aviso por correo y marca roja en el tablero.' });
  T('Las fechas se muestran un día antes en algunos navegadores', 'Front-end', 'alta', 'libre', { tipo: 'bug', creada: d(-3) });
  T('Costos por agente, desglosados por día', 'Backend', 'media', 'libre', { tipo: 'feature', creada: d(-47) });
  T('Exportar conversaciones a CSV', 'Backend', 'media', 'libre', { horas: 5, tipo: 'feature', creada: d(-33) });
  T('Documentar cómo conectar un agente nuevo', 'Backend', 'media', 'libre', { horas: 2, tipo: 'mant', creada: d(-9) });
  T('El historial de runs tarda en cargar con más de 500 filas', 'Backend', 'media', 'libre', { horas: 4, tipo: 'bug', creada: d(-16) });
  T('Vista móvil del dashboard', 'Front-end', 'nice', 'libre', { tipo: 'mejora', creada: d(-58) });
  T('Compartir un enlace a una vista filtrada', 'Front-end', 'nice', 'libre', { tipo: 'feature', creada: d(-52) });
  T('Soporte para varios idiomas', 'Front-end', 'nice', 'libre', { tipo: 'feature', creada: d(-141) });
  T('Resumen semanal por correo', 'Backend', 'nice', 'libre', { tipo: 'feature', creada: d(-28) });
  T('Integración con Teams', 'Backend', 'media', 'descartada', { tipo: 'feature', creada: d(-55), descartada: d(-20), motivo: 'La piden pocas personas y mantenerla cuesta más de lo que aporta.' });
  DB.producto.push(p);
  localStorage.setItem(LKEY, JSON.stringify(DB));
  irProducto(p.id);
}

/* ══════════════════════════════════════════════════════════════════════
   MI DÍA — las tareas de producto asignadas a una persona
   Se presentan con la forma de una actividad (esProducto la distingue) para
   que Mi día las agrupe por fecha igual que las de proyectos.
   ══════════════════════════════════════════════════════════════════════ */
function tareasAsignadasA() {
  const vivos = {};
  DB.producto.forEach((p) => { if (!p.archivado) vivos[p.id] = p; });
  return DB.tarea
    .filter((t) => t.dueno && vivos[t.producto] && (t.estado === 'libre' || t.estado === 'curso' || t.estado === 'revision'))
    .map((t) => ({ id: t.id, esProducto: true, producto: t.producto, proyecto: null, nombre: t.titulo,
      dueno: t.dueno, ini: t.ini || '', fin: t.fin || '', horas: t.horas || 0, avance: t.avance || 0,
      estado: t.estado, tipo: t.tipo }));
}

function lineaMiDiaProducto(a, hoy) {
  const p = ptProducto(a.producto);
  const tarde = a.fin && a.fin < hoy;
 return '<div class="act pt-linea' + (tarde ? ' atrasada' : '') + '">' +
    '<button class="palomear av-' + a.avance + '" onclick="ptAvance(\'' + a.id + '\')" ' +
    'title="Avanzar esta tarea: +25%, y después de 100% vuelve a 0%">' + (a.avance === 100 ? '✓' : a.avance + '%') + '</button>' +
    '<div class="act-cuerpo"><div class="act-tit">' + esc(a.nombre) + '</div>' +
    '<div class="act-meta">' +
    '<button class="chip-proy" onclick="irProducto(\'' + a.producto + '\')" title="Abrir el producto">◆ ' + esc(p ? p.nombre : 'producto') + '</button>' +
    '<span>' + esc(PT_NOMBRE_FASE[a.estado]) + '</span>' +
    '<span' + (tarde ? ' class="tarde"' : '') + '>' +
    (a.fin ? 'cierra ' + fechaCorta(a.fin) : 'sin fecha de cierre') +
    (tarde ? ' · ' + dias(a.fin, hoy) + ' días tarde' : '') + '</span>' +
    (a.horas ? '<span>' + a.horas + ' h</span>' : '') +
    '</div></div>' +
    '<button class="coment" onclick="irProducto(\'' + a.producto + '\')" title="Ver en su producto">→</button></div>';
}
