// Tablas dinámicas: vuelve a calcular sus celdas desde la caché, como hace Excel al filtrar con una segmentación.
// Cubre el diseño compacto de Excel (el de siempre): campos de fila anidados con subtotales arriba, un campo de
// columna o varios valores en columnas, filtros de informe y totales generales. Si la tabla usa algo más
// (diseño tabular, campos calculados, «mostrar valores como»…), devuelve null y la tabla se deja como está.
(function (g) {
  const O = (g.Oficina = g.Oficina || {});
  const on = (v) => v === '1' || v === 'true';
  const EPOCH = Date.UTC(1899, 11, 30);
  const valueOf = (el) => {
    const v = O.attr(el, 'v');
    switch (el.localName) {
      case 'n': return +v;
      case 's': return v;
      case 'b': return on(v);
      case 'd': return (Date.parse(v + (/Z|[+-]\d\d:?\d\d$/.test(v) ? '' : 'Z')) - EPOCH) / 86400000;
      case 'e': return { err: v };
      default: return null;
    }
  };
  const keyOf = (v) => (v === null || v === undefined ? '\u0000' : v && v.err ? '#' + v.err : typeof v + ':' + (typeof v === 'string' ? v.toLowerCase() : v));

  O.pivotModel = (xml, cache) => {
    if (!cache || !cache.records) return null;
    const root = O.parseXml(xml).documentElement;
    const def = O.parseXml(cache.def).documentElement;
    const cacheFields = O.kids(O.kid(def, 'cacheFields'), 'cacheField');
    if (cacheFields.some((f) => O.attr(f, 'formula') !== null || O.attr(f, 'databaseField') === '0' || O.kid(f, 'fieldGroup'))) return null;
    const fields = cacheFields.map((f) => ({ name: O.attr(f, 'name'), shared: O.kids(O.kid(f, 'sharedItems')).map(valueOf) }));
    const records = O.kids(O.parseXml(cache.records).documentElement, 'r').map((r) =>
      O.kids(r).map((el, i) => (el.localName === 'x' ? (fields[i] ? fields[i].shared[+O.attr(el, 'v')] : null) : valueOf(el))));
    const pivotFields = O.kids(O.kid(root, 'pivotFields'), 'pivotField').map((pf) => ({
      axis: O.attr(pf, 'axis'),
      compact: O.attr(pf, 'compact') !== '0',
      outline: O.attr(pf, 'outline') !== '0',
      subtotal: O.attr(pf, 'defaultSubtotal') !== '0',
      showAll: O.attr(pf, 'showAll') === '1',
      other: ['sumSubtotal', 'countASubtotal', 'avgSubtotal', 'maxSubtotal', 'minSubtotal', 'productSubtotal', 'countSubtotal', 'stdDevSubtotal', 'stdDevPSubtotal', 'varSubtotal', 'varPSubtotal'].some((a) => on(O.attr(pf, a))),
      items: O.kids(O.kid(pf, 'items'), 'item').map((it) => ({ x: O.attr(it, 'x') === null ? null : +O.attr(it, 'x'), t: O.attr(it, 't'), h: on(O.attr(it, 'h')) })),
    }));
    const fieldList = (tag) => O.kids(O.kid(root, tag), 'field').map((f) => +O.attr(f, 'x'));
    const loc = O.kid(root, 'location');
    const m = {
      ref: O.attr(loc, 'ref'),
      fields, records, pivotFields,
      rowFields: fieldList('rowFields'),
      colFields: fieldList('colFields'),
      pageFields: O.kids(O.kid(root, 'pageFields'), 'pageField').map((p) => ({ fld: +O.attr(p, 'fld'), item: O.attr(p, 'item') === null ? null : +O.attr(p, 'item') })),
      dataFields: O.kids(O.kid(root, 'dataFields'), 'dataField').map((d) => ({ name: O.attr(d, 'name'), fld: +O.attr(d, 'fld'), fn: O.attr(d, 'subtotal') || 'sum', as: O.attr(d, 'showDataAs') })),
      rowGrand: O.attr(root, 'rowGrandTotals') !== '0',
      colGrand: O.attr(root, 'colGrandTotals') !== '0',
      rowCaption: O.attr(root, 'rowHeaderCaption') || 'Etiquetas de fila',
      colCaption: O.attr(root, 'colHeaderCaption') || 'Etiquetas de columna',
      grandCaption: O.attr(root, 'grandTotalCaption') || 'Total general',
      dataCaption: O.attr(root, 'dataCaption') || 'Valores',
      dataOnRows: on(O.attr(root, 'dataOnRows')),
    };
    // Lo que Oficina sabe dibujar igual que Excel.
    const real = m.colFields.filter((x) => x >= 0);
    const ok = m.rowFields.length && m.dataFields.length && !m.dataOnRows && m.rowFields.every((x) => x >= 0 && pivotFields[x] && pivotFields[x].compact && pivotFields[x].outline && !pivotFields[x].other)
      && real.length <= 1 && !(real.length && m.dataFields.length > 1) && m.dataFields.every((d) => !d.as && AGG[d.fn])
      && records.every((r) => r.length >= fields.length);
    return ok ? m : null;
  };

  const nums = (list) => list.filter((v) => typeof v === 'number');
  const mean = (a) => a.reduce((s, v) => s + v, 0) / a.length;
  const sq = (a) => { const mu = mean(a); return a.reduce((s, v) => s + (v - mu) * (v - mu), 0); };
  const AGG = {
    sum: (v) => nums(v).reduce((s, x) => s + x, 0),
    count: (v) => v.filter((x) => x !== null && x !== '').length,
    countNums: (v) => nums(v).length,
    average: (v) => (nums(v).length ? mean(nums(v)) : { err: '#DIV/0!' }),
    max: (v) => (nums(v).length ? Math.max(...nums(v)) : 0),
    min: (v) => (nums(v).length ? Math.min(...nums(v)) : 0),
    product: (v) => nums(v).reduce((s, x) => s * x, 1),
    stdDev: (v) => (nums(v).length > 1 ? Math.sqrt(sq(nums(v)) / (nums(v).length - 1)) : { err: '#DIV/0!' }),
    stdDevp: (v) => (nums(v).length ? Math.sqrt(sq(nums(v)) / nums(v).length) : { err: '#DIV/0!' }),
    var: (v) => (nums(v).length > 1 ? sq(nums(v)) / (nums(v).length - 1) : { err: '#DIV/0!' }),
    varp: (v) => (nums(v).length ? sq(nums(v)) / nums(v).length : { err: '#DIV/0!' }),
  };

  // Valores de un campo en el orden en que los muestra la tabla, sin los que están ocultos.
  const shown = (m, f) => {
    const pf = m.pivotFields[f], list = [], hidden = new Set(), seen = new Set();
    for (const it of pf.items) {
      if (it.t || it.x === null) continue;
      const v = m.fields[f].shared[it.x], k = keyOf(v);
      seen.add(k);
      if (it.h) hidden.add(k);
      else list.push(v);
    }
    for (const r of m.records) if (!seen.has(keyOf(r[f]))) { seen.add(keyOf(r[f])); list.push(r[f]); }
    return { list, hidden };
  };

  // Filas y columnas de valores, empezando en la esquina de la tabla.
  O.pivotLayout = (m) => {
    const order = m.pivotFields.map((_, f) => shown(m, f));
    const page = m.pageFields.filter((p) => p.item !== null).map((p) => {
      const it = m.pivotFields[p.fld].items[p.item];
      return { f: p.fld, k: keyOf(it && it.x !== null ? m.fields[p.fld].shared[it.x] : null) };
    });
    const recs = m.records.filter((r) => order.every((o, f) => !o.hidden.has(keyOf(r[f]))) && page.every((p) => keyOf(r[p.f]) === p.k));
    const label = (v) => (v === null || v === '' ? '(en blanco)' : v);
    const agg = (d, list) => (list.length ? AGG[d.fn](list.map((r) => r[d.fld])) : null);
    const col = m.colFields.find((x) => x >= 0);
    const colItems = col === undefined ? [] : order[col].list.filter((v) => order[col].showAll || recs.some((r) => keyOf(r[col]) === keyOf(v)));
    const vals = (list) => {
      if (col === undefined) return m.dataFields.map((d) => agg(d, list));
      const d = m.dataFields[0];
      const out = colItems.map((v) => agg(d, list.filter((r) => keyOf(r[col]) === keyOf(v))));
      return m.colGrand ? [...out, agg(d, list)] : out;
    };
    const rows = [];
    if (col !== undefined) {
      rows.push([m.dataFields[0].name, m.colCaption]);
      rows.push([m.rowCaption, ...colItems.map(label), ...(m.colGrand ? [m.grandCaption] : [])]);
    } else if (m.dataFields.length > 1) {
      rows.push([null, m.dataCaption]);
      rows.push([m.rowCaption, ...m.dataFields.map((d) => d.name)]);
    } else rows.push([m.rowCaption, m.dataFields[0].name]);
    const walk = (level, list) => {
      const f = m.rowFields[level], pf = m.pivotFields[f];
      for (const v of order[f].list) {
        const sub = list.filter((r) => keyOf(r[f]) === keyOf(v));
        if (!sub.length && !pf.showAll) continue;
        const leaf = level === m.rowFields.length - 1;
        rows.push([label(v), ...(leaf || pf.subtotal ? vals(sub) : vals([]).map(() => null))]);
        if (!leaf) walk(level + 1, sub);
      }
    };
    walk(0, recs);
    if (m.rowGrand) rows.push([m.grandCaption, ...vals(recs)]);
    const at = O.parseAddr(m.ref.split(':')[0]);
    const width = Math.max(...rows.map((r) => r.length));
    return { r: at.r, c: at.c, rows, ref: `${O.addr(at.r, at.c)}:${O.addr(at.r + rows.length - 1, at.c + width - 1)}` };
  };

  // Vuelve a escribir las celdas de la tabla en la hoja. Devuelve false si no sabe calcularla.
  O.pivotRefresh = (sheet, pivot, cache) => {
    const m = O.pivotModel(pivot.xml, cache);
    if (!m) return false;
    const out = O.pivotLayout(m);
    const [a, b] = m.ref.split(':').map(O.parseAddr);
    for (let r = a.r; r <= (b || a).r; r++) for (let c = a.c; c <= (b || a).c; c++) {
      const key = O.addr(r, c), cell = sheet.cells[key];
      if (!cell) continue;
      delete cell.v;
      delete cell.f;
      if (!Object.keys(cell).length) delete sheet.cells[key];
    }
    out.rows.forEach((row, i) => row.forEach((v, j) => {
      if (v === null) return;
      const key = O.addr(out.r + i, out.c + j);
      sheet.cells[key] = { ...(sheet.cells[key] || {}), v };
    }));
    pivot.xml = pivot.xml.replace(/(<(?:\w+:)?location\b[^>]*\bref=")[^"]*"/, `$1${out.ref}"`);
    return true;
  };

  // Oculta en la definición de la tabla los elementos de un campo que no están elegidos (los x de la caché).
  O.pivotSetHidden = (xml, field, hidden, count) => {
    let n = -1;
    return xml.replace(/<pivotField\b[^>]*?(?:\/>|>[\s\S]*?<\/pivotField>)/g, (pf) => {
      if (++n !== field) return pf;
      const item = (x) => `<item x="${x}"${hidden.has(x) ? ' h="1"' : ''}/>`;
      if (/<items\b/.test(pf)) {
        return pf.replace(/<item\b([^>]*?)\/>/g, (all, attrs) => {
          const x = /\bx="(\d+)"/.exec(attrs);
          if (!x) return all;
          const rest = attrs.replace(/\s+h="[^"]*"/, '');
          return `<item${rest}${hidden.has(+x[1]) ? ' h="1"' : ''}/>`;
        });
      }
      // Un campo sin elementos los necesita para poder ocultar alguno.
      const items = `<items count="${count + 1}">${Array.from({ length: count }, (_, x) => item(x)).join('')}<item t="default"/></items>`;
      return pf.endsWith('/>') ? pf.replace(/\s*\/>$/, `>${items}</pivotField>`) : pf.replace(/^(<pivotField\b[^>]*>)/, `$1${items}`);
    });
  };

  // Aplica lo elegido en una segmentación de tabla dinámica a su caché y a cada tabla conectada.
  // Devuelve cuántas tablas ha podido recalcular y cuántas quedan para Excel.
  O.pivotSlicer = (state, slicer) => {
    const part = slicer.cachePart && state.parts && state.parts[slicer.cachePart];
    if (part && typeof part.data === 'string') {
      const pick = new Map(slicer.items.map((it) => [it.x, it.on]));
      part.data = part.data.replace(/<i\b([^>]*?)\/>/g, (all, attrs) => {
        const x = /\bx="(\d+)"/.exec(attrs);
        if (!x || !pick.has(+x[1])) return all;
        return `<i${attrs.replace(/\s+s="[^"]*"/, '')}${pick.get(+x[1]) ? ' s="1"' : ''}/>`;
      });
    }
    const hidden = new Set(slicer.items.filter((it) => !it.on).map((it) => it.x));
    let done = 0, left = 0;
    for (const link of slicer.pivots || []) {
      const sheet = state.sheets.find((s) => String(s.sheetId) === String(link.tabId));
      const pivot = sheet && (sheet.pivots || []).find((p) => new RegExp(`<(?:\\w+:)?pivotTableDefinition\\b[^>]*\\bname="${link.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`).test(p.xml));
      const cache = pivot && (state.pivotCaches || []).find((c) => c.id === pivot.cache);
      if (!cache) { left++; continue; }
      const fields = O.kids(O.kid(O.parseXml(cache.def).documentElement, 'cacheFields'), 'cacheField').map((x) => O.attr(x, 'name'));
      const f = fields.indexOf(slicer.source);
      if (f < 0) { left++; continue; }
      pivot.xml = O.pivotSetHidden(pivot.xml, f, hidden, slicer.shared || slicer.items.length);
      if (O.pivotRefresh(sheet, pivot, cache)) done++;
      else left++;
    }
    return { done, left };
  };
})(globalThis);
