// Reportes › Café en fruta CAPRIS (solo administradores): control informativo de lo entregado a Capris.
// Recibo = Fecha, Tipo, Recibo, Finca, Sociedad, Cajuelas; las fanegas se calculan (cajuelas ÷ 20). Sin dinero.
import { h, clear, toast, confirmar, diaHoy } from './util.js';
import { cafeFruta } from './sync.js';
import { puede } from './users.js';
import { leerArchivoTabla } from './leer_tabla.js';
import { descargar } from './files.js';

const CAJ_FANEGA = 20;
const norm = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const num = (v, d = 1) => Number(v || 0).toLocaleString('es-CR', { maximumFractionDigits: d }).replace(/[  ]/g, ' ');
const fan = (caj) => Number(caj || 0) / CAJ_FANEGA;
const plural = (n, a, b) => `${n} ${n === 1 ? a : b}`;
const corta = (s) => String(s || '').replace(/\s*S\.?A\.?$/i, '').replace(/^BENEFICIADORA\s+/i, 'BEN. ');
const dma = (iso) => (iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}` : '');
const p2 = (n) => String(n).padStart(2, '0');

/** Fecha de Excel o texto (número de serie, dd/mm/aaaa, aaaa-mm-dd) → 'aaaa-mm-dd' ('' si no se entiende). */
export function fechaISO(v) {
  const t = String(v ?? '').trim();
  if (!t) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(t)) return t.slice(0, 10);
  if (/^\d{5}(\.\d+)?$/.test(t)) { const d = new Date(Math.round((Number(t) - 25569) * 86400000)); return `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}`; }
  const m = t.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/);
  if (m) { const y = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]); return `${y}-${p2(m[2])}-${p2(m[1])}`; }
  return '';
}
const esNum = (s) => /^\d+([.,]\d+)?$/.test(String(s ?? '').trim());

/** Texto pegado desde Excel → filas {fecha, tipo, recibo, finca, sociedad, cajuelas}. Orden: Fecha, Tipo, Recibo, Finca, Sociedad, Cajuelas (si trae Rec. interno antes de las cajuelas, se ignora). */
export function parsearPegado(texto) {
  const out = [];
  String(texto || '').split(/\r?\n/).forEach((l) => {
    if (!l.trim()) return;
    const c = l.split('\t').map((x) => x.trim());
    if (/fecha/i.test(c[0]) && /recibo/i.test(l)) return;
    if (c.length < 4) { out.push({ fecha: fechaISO(c[0]), tipo: c[1] || '', recibo: c[2] || '', finca: '', sociedad: '', cajuelas: '' }); return; }
    const caj = esNum(c[5]) ? c[5] : (esNum(c[6]) ? c[6] : c[5] || '');
    out.push({ fecha: fechaISO(c[0]) || c[0], tipo: c[1], recibo: c[2], finca: c[3] || '', sociedad: c[4] || '', cajuelas: caj });
  });
  return out;
}

/** Filas de un archivo (con encabezados FECHA, TIPO, RECIBO, FINCA, SOCIEDAD, CAJUELAS) → filas. */
export function filasDeArchivo(filas) {
  const ini = filas.findIndex((r) => r.some((c) => /^recibo$/i.test(String(c).trim())));
  if (ini < 0) throw new Error('No encontré la fila de encabezados (FECHA, TIPO, RECIBO, FINCA, SOCIEDAD, CAJUELAS). Use la plantilla.');
  const enc = filas[ini].map(norm), f = (re) => enc.findIndex((x) => re.test(x));
  const col = { fecha: f(/^fecha/), tipo: f(/^tipo/), recibo: f(/^recibo/), finca: f(/^finca/), sociedad: f(/^sociedad/), cajuelas: f(/^cajuela/) };
  if (col.fecha < 0 || col.tipo < 0 || col.recibo < 0 || col.cajuelas < 0) throw new Error('Al archivo le faltan columnas: necesita FECHA, TIPO, RECIBO y CAJUELAS (y FINCA, SOCIEDAD). Use la plantilla.');
  const g = (r, i) => (i >= 0 ? String(r[i] ?? '').trim() : '');
  return filas.slice(ini + 1).filter((r) => r.some((c) => String(c).trim()) && g(r, col.recibo))
    .map((r) => ({ fecha: fechaISO(g(r, col.fecha)) || g(r, col.fecha), tipo: g(r, col.tipo), recibo: g(r, col.recibo), finca: g(r, col.finca), sociedad: g(r, col.sociedad), cajuelas: g(r, col.cajuelas) }))
    .filter((x) => !/^ejemplo/i.test(x.recibo));
}

export function crearCafeFruta({ abrir, volver, alCambiar }) {
  let res = null, error = '', cargando = false;
  let pantalla = 'inicio', param = null;
  const el = h('div', { class: 'cafe-fruta' });
  const tituloCab = h('h1', null, 'Café en fruta · CAPRIS');
  const cab = h('header', { class: 'cab-sub' }, h('button', { type: 'button', class: 'atras', 'aria-label': 'Volver', onclick: () => atras() }, '‹'), tituloCab);
  const chip = (t, c) => h('span', { class: 'chip-est ' + c }, t);
  const boton = (t, fn, cls = '') => h('button', { type: 'button', class: `btn-texto ${cls} ancho-total`.replace('  ', ' '), onclick: fn }, t);

  function ir(p, prm = null) { pantalla = p; param = prm; pintar(); window.scrollTo(0, 0); }
  function atras() {
    if (pantalla === 'inicio') volver();
    else if (pantalla === 'finca') ir('fincas');
    else if (pantalla === 'previa') ir(param.origen);
    else if (pantalla === 'recibo') ir(param.origen || 'recibos', param.origenParam || null);
    else ir('inicio');
  }

  async function cargar(forzar) {
    if (!puede('todos')) return;
    if (cargando) return;
    cargando = true; error = ''; alCambiar && alCambiar(); if (pantalla === 'inicio') pintar();
    const r = await cafeFruta('resumen', {});
    cargando = false;
    if (r.ok && r.json) res = r.json; else error = r.message || 'No se pudo cargar el café en fruta.';
    alCambiar && alCambiar(); pintar();
  }

  // ---------- Tarjeta del inicio de Reportes ----------
  function tarjetaHub(opc = {}) {
    const refBtn = h('button', { type: 'button', class: 'hub-ref' + (cargando ? ' gira' : ''), 'aria-label': 'Actualizar todo', onclick: (e) => { e.stopPropagation(); opc.alRefrescar && opc.alRefrescar(); } }, '↻');
    const cab3 = h('div', { class: 'fuente-cab' }, h('h3', null, '🍒 Café en fruta · CAPRIS'));
    if (!res) return h('div', { class: 'fuente vacia hub-fuente' }, refBtn, cab3, h('p', { class: 'sub' }, cargando ? 'Descargando datos…' : error ? '⚠ ' + error : 'Toque ↻ para traer los recibos.'));
    const tipos = res.tipos.filter((t) => t.recibos > 0 || t.maximo);
    const cuerpo = tipos.length ? tipos.map((t) => {
      const fa = fan(t.cajuelas), pct = t.maximo ? Math.min(100, Math.round(fa / t.maximo * 100)) : 0;
      return h('div', null, h('p', { class: 'sub' }, h('b', null, t.tipo), t.maximo ? ` · ${num(fa)} de ${num(t.maximo, 0)} fanegas` : ` · ${num(fa)} fanegas`),
        t.maximo ? h('div', { class: 'progreso', role: 'img', 'aria-label': `${pct}% entregado` }, h('i', { style: `width:${pct}%` })) : null,
        t.maximo ? h('p', { class: 'sub' }, `Faltan ${num(Math.max(0, t.maximo - fa))} fanegas`) : null);
    }) : [h('p', { class: 'sub' }, 'Todavía no hay recibos.')];
    return h('div', { class: 'fuente toque hub-fuente' + (cargando ? ' carga' : ''), role: 'button', tabindex: '0', onclick: () => { ir('inicio'); abrir(); } }, refBtn, cab3, ...cuerpo,
      h('p', { class: 'sub' }, plural(res.tipos.reduce((t, x) => t + x.recibos, 0), 'recibo', 'recibos')), h('span', { class: 'hub-flecha', 'aria-hidden': 'true' }, '›'));
  }

  // ---------- Pantallas ----------
  function pintar() {
    clear(el);
    const t = { inicio: 'Café en fruta · CAPRIS', fincas: 'Por finca', finca: param && param.finca || 'Finca', recibos: 'Recibos', nuevo: 'Nuevo recibo', pegar: 'Pegar desde Excel', archivo: 'Subir archivo', previa: 'Vista previa', recibo: param && param.r ? `Recibo ${param.r.recibo}` : 'Recibo', maximo: 'Máximo del contrato', listo: 'Café en fruta' }[pantalla] || 'Café en fruta';
    tituloCab.textContent = t;
    if (error) { el.append(h('p', { class: 'estado aviso' }, '⚠ ' + error), boton('Reintentar', () => cargar(true))); return; }
    if (!res) { el.append(h('p', { class: 'nota' }, cargando ? 'Cargando…' : 'Sin datos todavía.')); return; }
    ({ inicio, fincas, finca, recibos, nuevo, pegar, archivo, previa, recibo, maximo, listo })[pantalla]();
  }

  function inicio() {
    res.tipos.filter((t) => t.recibos > 0 || t.maximo).forEach((t) => {
      const fa = fan(t.cajuelas), pct = t.maximo ? Math.min(100, Math.round(fa / t.maximo * 100)) : 0;
      el.append(h('div', { class: 'tarjeta' }, h('b', null, t.tipo), h('p', { class: 'sub' }, t.maximo ? `${num(fa)} de ${num(t.maximo, 0)} fanegas` : `${num(fa)} fanegas entregadas`),
        t.maximo ? h('div', { class: 'progreso' }, h('i', { style: `width:${pct}%` })) : null,
        h('p', { class: 'nota' }, t.maximo ? (fa > t.maximo ? `Se pasó del máximo por ${num(fa - t.maximo)} fanegas` : `Faltan ${num(t.maximo - fa)} fanegas`) : 'Sin tope definido', ` · ${plural(t.recibos, 'recibo', 'recibos')}`)));
    });
    el.append(boton('＋ Recibo', () => ir('nuevo'), 'primario'), boton('📋 Pegar desde Excel', () => ir('pegar')), boton('📁 Subir archivo Excel / CSV', () => ir('archivo')),
      boton('Ver por finca', () => ir('fincas')), boton('Ver todos los recibos', () => ir('recibos')), boton('Máximo del contrato', () => ir('maximo')));
  }

  function fincas() {
    if (!res.fincas.length) { el.append(h('p', { class: 'nota' }, 'Todavía no hay recibos.')); return; }
    res.fincas.forEach((x) => el.append(h('div', { class: 'fila-aj fila-usuario', role: 'button', tabindex: '0', onclick: () => ir('finca', { finca: x.finca }) },
      h('span', { class: 'tx' }, x.finca || 'Sin finca', h('small', null, `${corta(x.sociedad)}${x.sociedad ? ' · ' : ''}${plural(x.recibos, 'recibo', 'recibos')}`)), h('b', null, num(fan(x.cajuelas))))));
    el.append(h('p', { class: 'nota' }, 'Fanegas entregadas por finca. Toque una para ver sus recibos.'));
  }

  const filaRecibo = (r, origen, origenParam) => h('div', { class: 'fila-aj fila-usuario', role: 'button', tabindex: '0', onclick: () => ir('recibo', { r, origen, origenParam }) },
    h('span', { class: 'tx' }, `${r.recibo}${r.finca ? ' · ' + r.finca : ''}`, h('small', null, `${dma(r.fecha)} · ${r.tipo} · ${num(r.cajuelas, 2)} cajuelas`)), h('b', null, num(fan(r.cajuelas), 2)));
  function finca() {
    const rs = res.recibos.filter((r) => r.finca === param.finca);
    const tot = rs.reduce((t, r) => t + Number(r.cajuelas), 0);
    el.append(h('div', { class: 'tarjeta' }, h('b', null, `${num(fan(tot), 2)} fanegas`), h('p', { class: 'nota' }, `${plural(rs.length, 'recibo', 'recibos')} · ${num(tot, 2)} cajuelas`)));
    rs.forEach((r) => el.append(filaRecibo(r, 'finca', param)));
  }
  function recibos() {
    const q = h('input', { class: 'campo', type: 'search', placeholder: '🔍 Buscar recibo o finca', autocomplete: 'off' });
    const zona = h('div');
    const dib = () => {
      clear(zona);
      const t = norm(q.value).trim();
      const rs = res.recibos.filter((r) => !t || norm(r.recibo).includes(t) || norm(r.finca).includes(t) || norm(r.tipo).includes(t));
      if (!rs.length) zona.append(h('p', { class: 'nota' }, 'Ningún recibo coincide.'));
      rs.slice(0, 200).forEach((r) => zona.append(filaRecibo(r, 'recibos')));
      if (rs.length > 200) zona.append(h('p', { class: 'nota' }, `Mostrando 200 de ${rs.length}. Escriba para filtrar.`));
    };
    q.addEventListener('input', dib);
    el.append(q, zona); dib();
  }

  // ---------- Formulario de un recibo (nuevo o editar) ----------
  function formulario(ini, textoBtn, enviar) {
    const NUEVO = '__nuevo';
    const fecha = h('input', { class: 'campo', type: 'date', value: ini.fecha || diaHoy() });
    const selTipo = h('select', { class: 'campo selec' }, ...res.tipos.map((t) => h('option', { value: t.tipo }, t.tipo)), h('option', { value: NUEVO }, '＋ Tipo nuevo…'));
    const tipoNuevo = h('input', { class: 'campo', type: 'text', placeholder: 'Nombre del tipo', autocomplete: 'off', hidden: true });
    const recibo = h('input', { class: 'campo', type: 'text', inputmode: 'numeric', value: ini.recibo || '', autocomplete: 'off' });
    const selFinca = h('select', { class: 'campo selec' }, h('option', { value: '' }, '(sin finca)'), ...res.fincasLista.map((f) => h('option', { value: f.finca }, f.finca)), h('option', { value: NUEVO }, '＋ Finca nueva…'));
    const fincaNueva = h('input', { class: 'campo', type: 'text', placeholder: 'Nombre de la finca', autocomplete: 'off', hidden: true });
    const soc = h('input', { class: 'campo', type: 'text', value: ini.sociedad || '', autocomplete: 'off', list: 'cf-soc' });
    const dl = h('datalist', { id: 'cf-soc' }, ...res.sociedades.map((s) => h('option', { value: s })));
    const caj = h('input', { class: 'campo', type: 'text', inputmode: 'decimal', value: ini.cajuelas != null ? String(ini.cajuelas) : '', autocomplete: 'off' });
    const calc = h('p', { class: 'calc-nota' });
    const defSoc = (f) => { const x = res.fincasLista.find((y) => y.finca === f); return x ? x.sociedad : ''; };
    const masUsado = [...res.tipos].sort((x, y) => y.recibos - x.recibos)[0];
    if (!ini.tipo && masUsado) selTipo.value = masUsado.tipo;
    if (ini.tipo) selTipo.value = res.tipos.some((t) => t.tipo === ini.tipo) ? ini.tipo : NUEVO;
    if (ini.tipo && selTipo.value === NUEVO) tipoNuevo.value = ini.tipo;
    if (ini.finca) selFinca.value = res.fincasLista.some((f) => f.finca === ini.finca) ? ini.finca : NUEVO;
    if (ini.finca && selFinca.value === NUEVO) fincaNueva.value = ini.finca;
    const act = () => {
      tipoNuevo.hidden = selTipo.value !== NUEVO; fincaNueva.hidden = selFinca.value !== NUEVO;
      const c = Number(String(caj.value).replace(',', '.'));
      calc.textContent = c > 0 ? `= ${num(fan(c), 2)} fanegas` : 'Las fanegas se calculan solas (cajuelas ÷ 20).';
    };
    selTipo.addEventListener('change', act); selFinca.addEventListener('change', () => { act(); if (selFinca.value && selFinca.value !== NUEVO) soc.value = defSoc(selFinca.value) || soc.value; }); caj.addEventListener('input', act);
    act();
    const btn = h('button', { type: 'button', class: 'btn-texto primario ancho-total', onclick: async () => {
      const datos = { fecha: fecha.value, tipo: selTipo.value === NUEVO ? tipoNuevo.value : selTipo.value, recibo: recibo.value, finca: selFinca.value === NUEVO ? fincaNueva.value : selFinca.value, sociedad: soc.value, cajuelas: caj.value };
      btn.disabled = true; await enviar(datos); btn.disabled = false;
    } }, textoBtn);
    return [h('label', { class: 'etiqueta' }, 'Fecha'), fecha, h('label', { class: 'etiqueta' }, 'Tipo'), selTipo, tipoNuevo, h('label', { class: 'etiqueta' }, 'Recibo'), recibo,
      h('label', { class: 'etiqueta' }, 'Finca'), selFinca, fincaNueva, h('label', { class: 'etiqueta' }, 'Sociedad (sale de la finca; cámbiela si hace falta)'), soc, dl,
      h('label', { class: 'etiqueta' }, 'Cajuelas'), caj, calc, btn];
  }
  function nuevo() {
    el.append(...formulario({}, 'Guardar recibo', async (d) => {
      const r = await cafeFruta('guardar', { filas: [d] });
      if (!r.ok) return toast(r.message, 5000);
      const f = r.json.filas[0];
      if (f.estado !== 'nueva') return toast(f.mensaje || 'No se guardó.', 5000);
      toast('Recibo guardado.'); await cargar(true); ir('inicio');
    }), boton('Cancelar', () => ir('inicio')));
  }

  // ---------- Pegar y archivo ----------
  function pegar() {
    const area = h('textarea', { class: 'campo', rows: '8', placeholder: 'Pegue aquí las filas copiadas de Excel:\nFecha ⇥ Tipo ⇥ Recibo ⇥ Finca ⇥ Sociedad ⇥ Cajuelas', autocomplete: 'off', spellcheck: 'false' });
    el.append(h('p', { class: 'nota' }, 'Columnas en este orden: Fecha, Tipo, Recibo, Finca, Sociedad, Cajuelas.'), area,
      boton('Ver vista previa', () => { const f = parsearPegado(area.value); if (!f.length) return toast('Pegue primero las filas de Excel.'); revisar(f, '', 'pegar'); }, 'primario'), boton('Cancelar', () => ir('inicio')));
  }
  function archivo() {
    const arch = h('input', { type: 'file', accept: '.xlsx,.csv,.txt,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', hidden: true });
    arch.addEventListener('change', async () => {
      const f = arch.files && arch.files[0]; if (!f) return;
      try {
        const filas = filasDeArchivo(await leerArchivoTabla(f, { hoja: /recibos/i, clave: /^recibo$/i }));
        if (!filas.length) return toast('No encontré recibos en el archivo.', 5000);
        await revisar(filas, f.name, 'archivo');
      } catch (e) { toast(e.message || 'No se pudo leer el archivo.', 6000); }
      arch.value = '';
    });
    el.append(h('p', { class: 'nota' }, 'Un recibo por fila: FECHA, TIPO, RECIBO, FINCA, SOCIEDAD y CAJUELAS. Use la plantilla para que quede igual.'), arch,
      boton('📁 Elegir archivo', () => arch.click(), 'primario'),
      boton('⬇ Descargar la plantilla', async () => { try { const r = await fetch('./Plantilla_Cafe_en_Fruta.xlsx', { cache: 'no-store' }); if (!r.ok) throw new Error(); descargar(await r.blob(), 'Plantilla_Cafe_en_Fruta.xlsx'); } catch { toast('No se pudo bajar la plantilla (¿sin internet?).', 5000); } }),
      boton('Cancelar', () => ir('inicio')));
  }
  async function revisar(filas, nombre, origen) {
    if (filas.length > 1500) return toast('Máximo 1500 filas por carga. Divida el archivo.', 6000);
    toast('Revisando…', 1500);
    const r = await cafeFruta('previa', { filas });
    if (!r.ok) return toast(r.message, 6000);
    ir('previa', { filas, j: r.json, nombre, origen });
  }
  function previa() {
    const { filas, j, nombre } = param;
    const nuevas = j.filas.filter((f) => f.estado === 'nueva');
    const caj = nuevas.reduce((t, f) => t + Number(f.cajuelas || 0), 0);
    el.append(h('div', { class: 'tarjeta' }, h('b', null, nombre ? `📁 ${nombre}` : 'Revisión'), h('p', { class: 'nota' }, `${plural(j.nuevas, 'nuevo', 'nuevos')} · ${j.ya} ya ${j.ya === 1 ? 'estaba' : 'estaban'} · ${j.errores} con error`),
      h('p', { class: 'nota' }, `Los nuevos suman ${num(caj, 2)} cajuelas = ${num(fan(caj), 2)} fanegas.`)));
    const orden = { error: 0, nueva: 1, ya: 2 };
    [...j.filas].sort((a, b) => orden[a.estado] - orden[b.estado]).slice(0, 300).forEach((f) => {
      const o = filas[f.i - 1] || {};
      const det = f.estado === 'error' ? f.mensaje : f.estado === 'ya' ? f.mensaje : `${dma(f.fecha)} · ${num(f.cajuelas, 2)} cajuelas · ${num(fan(f.cajuelas), 2)} fanegas` + (f.fincaNueva ? ` · finca nueva (${f.sociedad || 'sin sociedad'})` : '') + (f.tipoNuevo ? ' · tipo nuevo' : '');
      el.append(h('div', { class: 'fila-aj fila-usuario' }, h('span', { class: 'tx' }, `${f.recibo || o.recibo || '(sin recibo)'}${f.finca ? ' · ' + f.finca : ''}`, h('small', null, f.tipo || o.tipo || ''), h('small', null, det)),
        f.estado === 'error' ? chip('Revisar', 'ambar') : f.estado === 'ya' ? chip('Ya está', 'gris') : chip(f.fincaNueva || f.tipoNuevo ? 'Nueva finca/tipo' : 'Nuevo', 'verde')));
    });
    if (j.filas.length > 300) el.append(h('p', { class: 'nota' }, `Se muestran 300 de ${j.filas.length}; al guardar entran todas las nuevas.`));
    if (j.errores) el.append(h('p', { class: 'nota' }, `Las ${j.errores} con error no se guardan: corríjalas y cárguelas de nuevo.`));
    const btn = h('button', { type: 'button', class: 'btn-texto primario ancho-total', disabled: !j.nuevas, onclick: async () => {
      btn.disabled = true;
      const r = await cafeFruta('guardar', { filas });
      if (!r.ok) { btn.disabled = false; return toast(r.message, 6000); }
      await cargar(true); ir('listo', { j: r.json });
    } }, j.nuevas ? `Guardar ${plural(j.nuevas, 'recibo', 'recibos')}` : 'No hay nada nuevo para guardar');
    el.append(btn, boton('Volver', () => ir(param.origen)));
  }
  function listo() {
    const j = param.j;
    el.append(h('div', { class: 'tarjeta' }, h('b', null, `✓ ${plural(j.nuevas, 'recibo guardado', 'recibos guardados')}`), h('p', { class: 'nota' }, j.ya || j.errores ? `${j.ya} ya estaban y ${j.errores} traían error; esos no se guardaron.` : '')), boton('Listo', () => ir('inicio'), 'primario'));
  }

  // ---------- Un recibo ----------
  function recibo() {
    const r = param.r;
    el.append(h('div', { class: 'tarjeta' }, h('b', null, `${r.finca || 'Sin finca'}${r.sociedad ? ' · ' + corta(r.sociedad) : ''}`), h('p', { class: 'nota' }, `${dma(r.fecha)} · ${r.tipo} · ${num(r.cajuelas, 2)} cajuelas = ${num(fan(r.cajuelas), 2)} fanegas`)),
      ...formulario(r, 'Guardar cambios', async (d) => {
        const x = await cafeFruta('editar', { id: r.id, ...d }); toast(x.message, 4000);
        if (x.ok) { await cargar(true); ir(param.origen || 'recibos', param.origenParam || null); }
      }),
      h('button', { type: 'button', class: 'btn-texto peligro ancho-total', onclick: async () => {
        if (!(await confirmar(`Anular el recibo ${r.recibo}`, 'Deja de contar en las fanegas entregadas. Queda guardado en el historial.', { si: 'Anular', peligro: true }))) return;
        const x = await cafeFruta('anular', { id: r.id }); toast(x.message, 4000);
        if (x.ok) { await cargar(true); ir(param.origen || 'recibos', param.origenParam || null); }
      } }, 'Anular recibo'));
  }

  // ---------- Máximo por tipo ----------
  function maximo() {
    el.append(h('p', { class: 'nota' }, 'Fanegas máximas del contrato con Capris para cada tipo. Si lo amplían, cámbielo aquí y «Faltan» se recalcula. Vacío = sin tope.'));
    res.tipos.forEach((t) => {
      const inp = h('input', { class: 'campo', type: 'text', inputmode: 'decimal', value: t.maximo != null ? String(t.maximo) : '', placeholder: 'Sin tope', autocomplete: 'off' });
      el.append(h('div', { class: 'tarjeta' }, h('b', null, t.tipo), inp, h('button', { type: 'button', class: 'btn-texto primario ancho-total', onclick: async () => {
        const x = await cafeFruta('maximo', { tipo: t.tipo, maximo: inp.value }); toast(x.message, 4000); if (x.ok) await cargar(true);
      } }, 'Guardar máximo')));
    });
    el.append(boton('Volver', () => ir('inicio')));
  }

  return { cab, el, tarjetaHub, cargar, ocupado: () => cargando, reiniciar() { pantalla = 'inicio'; param = null; pintar(); cargar(false); } };
}
