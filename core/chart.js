// Gráficos de Cálculo: lee el XML de un gráfico de Excel, lo dibuja en SVG y crea gráficos nuevos.
(function (g) {
  const O = g.Oficina;
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  // Tema por defecto de Office (lt1, dk1, lt2, dk2, acentos 1-6, vínculos).
  O.DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
  const KINDS = { barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie',
    doughnutChart: 'doughnut', areaChart: 'area', area3DChart: 'area', scatterChart: 'scatter', radarChart: 'radar', bubbleChart: 'scatter' };

  // Luminosidad de los colores del tema (lumMod, lumOff, tint y shade), como en DrawingML.
  const hsl = (hex) => {
    const [r, gr, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
    const max = Math.max(r, gr, b), min = Math.min(r, gr, b), l = (max + min) / 2;
    if (max === min) return [0, 0, l];
    const d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    const h = (max === r ? (gr - b) / d + (gr < b ? 6 : 0) : max === gr ? (b - r) / d + 2 : (r - gr) / d + 4) / 6;
    return [h, s, l];
  };
  const fromHsl = (h, s, l) => {
    const f = (p, q, t) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p;
    };
    let rgb = [l, l, l];
    if (s) {
      const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
      rgb = [f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)];
    }
    return rgb.map((x) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  };
  const SCHEME = { lt1: 0, bg1: 0, dk1: 1, tx1: 1, lt2: 2, bg2: 2, dk2: 3, tx2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10, folHlink: 11 };
  O.drawingColor = (el, theme) => {
    if (!el) return null;
    const clr = O.kids(el).find((k) => /Clr$/.test(k.localName));
    if (!clr) return null;
    let hex = null;
    if (clr.localName === 'srgbClr') hex = (O.attr(clr, 'val') || '').toUpperCase();
    else if (clr.localName === 'schemeClr') hex = theme[SCHEME[O.attr(clr, 'val')]] || null;
    else if (clr.localName === 'sysClr') hex = (O.attr(clr, 'lastClr') || '000000').toUpperCase();
    else if (clr.localName === 'prstClr') hex = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '00FF00', blue: '0000FF' }[O.attr(clr, 'val')] || '000000';
    if (!hex) return null;
    let [h, s, l] = hsl(hex);
    for (const m of O.kids(clr)) {
      const v = +O.attr(m, 'val') / 100000;
      if (m.localName === 'lumMod') l *= v;
      else if (m.localName === 'lumOff') l += v;
      else if (m.localName === 'tint') l = l * v + (1 - v);
      else if (m.localName === 'shade') l *= v;
    }
    return fromHsl(h, s, Math.max(0, Math.min(1, l)));
  };
  // Colores que Excel da a las series sin color propio: los seis acentos y luego sus variantes más oscuras y más claras.
  const autoColor = (i, theme) => {
    const base = theme[4 + (i % 6)] || '4472C4';
    const round = Math.floor(i / 6);
    if (!round) return base;
    const [h, s, l] = hsl(base);
    const k = [1, 0.6, 0.8, 0.5, 0.7][round % 5];
    return fromHsl(h, s, round % 2 ? l * k : l * k + (1 - k));
  };

  const textOf = (el) => O.all(el, 't').map((t) => t.textContent).join('');
  const cache = (ref) => {
    const pts = O.all(ref, 'pt');
    const n = +O.attr(O.all(ref, 'ptCount')[0], 'val') || pts.length;
    const out = Array(n).fill(null);
    for (const p of pts) {
      const t = O.kid(p, 'v') ? O.kid(p, 'v').textContent : '';
      out[+O.attr(p, 'idx')] = t !== '' && !isNaN(t) && O.all(ref, 'numCache').length ? +t : t;
    }
    return out;
  };

  // Modelo para dibujar. Las referencias se guardan por su posición entre los <c:f> del XML.
  const models = new Map();
  O.chartModel = (xml, theme) => {
    const key = xml + '|' + theme.join();
    if (models.has(key)) return models.get(key);
    const doc = O.parseXml(xml);
    const root = doc.documentElement;
    const fs = O.all(root, 'f');
    const refOf = (el) => {
      const f = el && O.all(el, 'f')[0];
      return f ? fs.indexOf(f) : -1;
    };
    const chart = O.kid(root, 'chart');
    const plot = O.kid(chart, 'plotArea');
    const groups = O.kids(plot).filter((k) => KINDS[k.localName]);
    const first = groups[0];
    const val = (el, name) => O.attr(O.kid(el, name), 'val');
    const m = {
      kind: first ? KINDS[first.localName] : 'bar',
      horizontal: val(first, 'barDir') === 'bar',
      grouping: val(first, 'grouping') || 'clustered',
      gap: +(val(first, 'gapWidth') || 150),
      overlap: +(val(first, 'overlap') || 0),
      hole: +(val(first, 'holeSize') || 50),
      firstAngle: +(val(first, 'firstSliceAng') || 0),
      vary: val(first, 'varyColors') === '1',
      series: [],
      legend: null,
      title: null,
      threeD: !!first && /3D/.test(first.localName),
    };
    let n = 0;
    for (const grp of groups) {
      const kind = KINDS[grp.localName];
      for (const ser of O.kids(grp, 'ser')) {
        const tx = O.kid(ser, 'tx');
        const sp = O.kid(ser, 'spPr');
        const ln = O.kid(sp, 'ln');
        const lineOff = !!O.kid(ln, 'noFill');
        const color = O.drawingColor(O.kid(sp, 'solidFill'), theme) || O.drawingColor(O.kid(ln, 'solidFill'), theme);
        const points = {};
        for (const d of O.kids(ser, 'dPt')) {
          const c = O.drawingColor(O.path(d, 'spPr', 'solidFill'), theme);
          if (c) points[+val(d, 'idx')] = c;
        }
        const cat = O.kid(ser, 'cat') || O.kid(ser, 'xVal');
        const vals = O.kid(ser, 'val') || O.kid(ser, 'yVal');
        const symbol = O.attr(O.path(ser, 'marker', 'symbol'), 'val');
        const labels = O.kid(ser, 'dLbls') || O.kid(grp, 'dLbls');
        m.series.push({
          kind,
          name: { ref: refOf(O.kid(tx, 'strRef')), v: tx ? textOf(tx) || (O.kid(tx, 'v') || {}).textContent || '' : '' },
          cat: { ref: refOf(cat), cache: cat ? cache(cat) : [] },
          val: { ref: refOf(vals), cache: vals ? cache(vals) : [] },
          color: color || (kind === 'line' || kind === 'scatter' || kind === 'radar' ? null : null),
          auto: autoColor(n, theme),
          points,
          line: !lineOff,
          marker: kind === 'scatter' ? symbol !== 'none' : kind === 'line' ? !!symbol && symbol !== 'none' : false,
          smooth: val(ser, 'smooth') === '1',
          labels: labels ? { val: val(labels, 'showVal') === '1', pct: val(labels, 'showPercent') === '1', cat: val(labels, 'showCatName') === '1' } : null,
        });
        n++;
      }
    }
    // Marcadores automáticos: Excel va cambiando la forma en cada serie.
    const AUTO = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus', 'dash', 'dot'];
    const groupMarker = first && val(first, 'marker') === '1';
    let k = 0;
    for (const grp of groups) for (const ser of O.kids(grp, 'ser')) {
      const s = m.series[k];
      const symbol = O.attr(O.path(ser, 'marker', 'symbol'), 'val');
      if ((s.kind === 'line' || s.kind === 'scatter') && (symbol === null || symbol === 'auto') && (groupMarker || s.kind === 'scatter')) s.marker = true;
      if (s.marker) s.symbol = symbol && symbol !== 'auto' ? symbol : AUTO[k % AUTO.length];
      const mk = O.kid(ser, 'marker');
      s.markerFill = O.drawingColor(O.path(mk, 'spPr', 'solidFill'), theme);
      k++;
    }
    // Como en Excel, los datos de filas o columnas ocultas no se dibujan, salvo que el gráfico diga lo contrario.
    m.visibleOnly = val(chart, 'plotVisOnly') !== '0';
    const t = O.kid(chart, 'title');
    const deleted = val(chart, 'autoTitleDeleted') === '1';
    if (t) {
      const run = O.all(t, 'rPr')[0] || O.all(t, 'defRPr')[0];
      m.titleBold = run ? O.attr(run, 'b') === '1' : false;
      if (run && O.attr(run, 'sz')) m.titleSize = +O.attr(run, 'sz') / 100;
    }
    if (t) {
      const strRef = O.path(t, 'tx', 'strRef');
      m.title = strRef ? { ref: refOf(strRef), v: textOf(strRef) || (O.all(strRef, 'v')[0] || {}).textContent || '' } : { v: textOf(O.kid(t, 'tx')) || '' };
      if (!m.title.v && m.title.ref < 0) m.title.auto = true;
    } else if (!deleted && m.series.length === 1) m.title = { auto: true };
    const legend = O.kid(chart, 'legend');
    if (legend) m.legend = val(legend, 'legendPos') || 'r';
    const valAx = O.kids(plot, 'valAx');
    const ax = valAx[valAx.length > 1 && m.kind === 'scatter' ? 1 : 0];
    if (ax) {
      m.valFmt = O.attr(O.kid(ax, 'numFmt'), 'formatCode');
      if (O.attr(O.kid(ax, 'numFmt'), 'sourceLinked') === '1') m.valFmt = null;
      const sc = O.kid(ax, 'scaling');
      if (val(sc, 'min') !== null) m.min = +val(sc, 'min');
      if (val(sc, 'max') !== null) m.max = +val(sc, 'max');
      if (val(ax, 'majorUnit') !== null) m.unit = +val(ax, 'majorUnit');
      m.grid = !!O.kid(ax, 'majorGridlines');
      m.valHidden = val(ax, 'delete') === '1';
    }
    const catAx = O.kid(plot, 'catAx') || O.kid(plot, 'dateAx');
    if (catAx && m.kind !== 'scatter') m.catHidden = val(catAx, 'delete') === '1';
    if (m.kind === 'scatter' && valAx[0]) m.xHidden = val(valAx[0], 'delete') === '1';
    // Un gráfico sin formato propio (como los de Excel 2007) usa texto negro y título de 18 puntos en negrita.
    m.legacy = !O.kid(root, 'spPr') && !O.kid(root, 'txPr');
    if (m.legacy && m.title) {
      const run = O.all(t, 'rPr')[0] || O.all(t, 'defRPr')[0];
      if (!run || O.attr(run, 'b') === null) m.titleBold = true;
      if (!m.titleSize) m.titleSize = 18;
    }
    m.refs = fs.length;
    models.set(key, m);
    return m;
  };

  // Números redondos para el eje de valores, como los que elige Excel.
  const niceStep = (span, slots) => {
    const raw = span / Math.max(1, slots);
    const p = Math.pow(10, Math.floor(Math.log10(raw)));
    const f = raw / p;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p;
  };
  const scale = (lo, hi, px, m, zero = true) => {
    if (zero) { lo = Math.min(0, lo); hi = Math.max(0, hi); }
    if (lo === hi) hi = lo + 1;
    // Excel deja un 5 % de aire por encima de los datos antes de redondear, y pone hasta 10 divisiones si caben.
    const top = m.max === undefined && hi > 0 ? hi + 0.05 * (hi - lo) : hi;
    const bottom = m.min === undefined && lo < 0 ? lo - 0.05 * (hi - lo) : lo;
    const step = m.unit || niceStep(top - bottom, Math.max(2, Math.min(10, Math.floor(px / 16))));
    const min = m.min !== undefined ? m.min : Math.floor(bottom / step + 1e-9) * step;
    let max = m.max !== undefined ? m.max : Math.ceil(top / step - 1e-9) * step;
    if (max === min) max = min + step;
    const ticks = [];
    for (let v = min; v <= max + step / 1e6; v += step) ticks.push(Math.round(v / step) * step);
    return { min, max, step, ticks };
  };
  const esc = (s) => O.esc(String(s == null ? '' : s));
  // Marcador de una serie de líneas o dispersión, con las formas de Excel.
  const marker = (x, y, symbol, color) => {
    const r = 3.5, f = `fill="#${color}"`, st = `stroke="#${color}" stroke-width="1.5"`;
    switch (symbol) {
      case 'diamond': return `<path d="M${x} ${y - r - 0.5}L${x + r + 0.5} ${y}L${x} ${y + r + 0.5}L${x - r - 0.5} ${y}Z" ${f}/>`;
      case 'square': return `<rect x="${x - r}" y="${y - r}" width="${2 * r}" height="${2 * r}" ${f}/>`;
      case 'triangle': return `<path d="M${x} ${y - r - 0.5}L${x + r + 0.5} ${y + r}L${x - r - 0.5} ${y + r}Z" ${f}/>`;
      case 'x': return `<path d="M${x - r} ${y - r}L${x + r} ${y + r}M${x + r} ${y - r}L${x - r} ${y + r}" ${st}/>`;
      case 'star': return `<path d="M${x - r} ${y - r}L${x + r} ${y + r}M${x + r} ${y - r}L${x - r} ${y + r}M${x} ${y - r}V${y + r}" ${st}/>`;
      case 'plus': return `<path d="M${x - r} ${y}H${x + r}M${x} ${y - r}V${y + r}" ${st}/>`;
      case 'dash': return `<path d="M${x - r} ${y}H${x + r}" ${st}/>`;
      case 'dot': return `<circle cx="${x}" cy="${y}" r="1.5" ${f}/>`;
      case 'none': return '';
      default: return `<circle cx="${x}" cy="${y}" r="${r}" ${f}/>`;
    }
  };
  const fmt = (v, nf) => (typeof v === 'number' ? (nf && nf !== 'General' ? O.formatValue(v, nf) : O.formatValue(Math.round(v * 1e9) / 1e9, 'General')) : String(v == null ? '' : v));

  // SVG del gráfico. data: series ya resueltas {name, cats, vals}.
  O.chartSvg = (m, data, w, h, theme) => {
    const INK = m.legacy ? '000000' : '595959', LINE = m.legacy ? 'BFBFBF' : 'D9D9D9', FONT = 'font-family="Calibri, Carlito, sans-serif"';
    const parts = [`<rect x="0.5" y="0.5" width="${w - 1}" height="${h - 1}" fill="#fff" stroke="#${m.legacy ? '868686' : LINE}"/>`];
    const colorOf = (s, i, j) => (j !== undefined && s.points[j]) || (m.vary || m.kind === 'pie' || m.kind === 'doughnut' ? (j !== undefined ? s.points[j] || autoColor(j, theme) : null) : null) || s.color || s.auto;
    let top = 8, bottom = h - 8, left = 8, right = w - 8;
    const title = m.title && (m.title.text || m.title.v);
    if (title) {
      const tsz = ((m.titleSize || 14) * 4) / 3;
      parts.push(`<text x="${w / 2}" y="${top + tsz}" text-anchor="middle" font-size="${tsz}"${m.titleBold ? ' font-weight="700"' : ''} fill="#${INK}" ${FONT}>${esc(title)}</text>`);
      top += tsz + 12;
    }
    const pie = m.kind === 'pie' || m.kind === 'doughnut';
    const entries = pie ? (data[0] ? data[0].cats.map((c, j) => ({ label: c, color: colorOf(m.series[0], 0, j) })) : [])
      : data.map((s, i) => ({ label: s.name, color: colorOf(m.series[i], i), line: m.series[i].kind === 'line' || (m.series[i].kind === 'scatter' && m.series[i].line) }));
    // En barras horizontales Excel pone la primera serie abajo, y la leyenda en ese mismo orden.
    if (!pie && m.kind === 'bar' && m.horizontal) entries.reverse();
    if (m.legend && entries.length) {
      const pos = m.legend;
      const size = 12, item = (e) => 16 + Math.min(160, String(e.label).length * 6.2) + 10;
      const swatch = (e, x, y) => (e.line ? `<line x1="${x}" y1="${y - 4}" x2="${x + 12}" y2="${y - 4}" stroke="#${e.color}" stroke-width="2.5"/>` : `<rect x="${x + 2}" y="${y - 8}" width="8" height="8" fill="#${e.color}"/>`)
        + `<text x="${x + 16}" y="${y}" font-size="${size}" fill="#${INK}" ${FONT}>${esc(String(e.label).slice(0, 26))}</text>`;
      if (pos === 'b' || pos === 't') {
        const total = entries.reduce((a, e) => a + item(e), 0);
        let x = Math.max(left, (w - total) / 2);
        const y = pos === 'b' ? bottom - 2 : top + 12;
        for (const e of entries) { parts.push(swatch(e, x, y)); x += item(e); }
        if (pos === 'b') bottom -= 22; else top += 22;
      } else {
        const wide = Math.min(w / 3, Math.max(...entries.map(item)));
        const x = pos === 'l' ? left : right - wide;
        let y = (top + bottom) / 2 - (entries.length * 18) / 2 + 12;
        for (const e of entries) { parts.push(swatch(e, x, y)); y += 18; }
        if (pos === 'l') left += wide + 6; else right -= wide + 6;
      }
    }
    if (pie) {
      const s = data[0];
      const vals = s ? s.vals.map((v) => (typeof v === 'number' && v > 0 ? v : 0)) : [];
      const sum = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = (left + right) / 2, cy = (top + bottom) / 2, r = Math.max(10, Math.min(right - left, bottom - top) / 2 - 6);
      const r0 = m.kind === 'doughnut' ? (r * m.hole) / 100 : 0;
      let a = (m.firstAngle * Math.PI) / 180;
      const pt = (ang, rad) => [cx + rad * Math.sin(ang), cy - rad * Math.cos(ang)];
      vals.forEach((v, j) => {
        if (!v) return;
        const b = a + (2 * Math.PI * v) / sum, big = b - a > Math.PI ? 1 : 0;
        const [x1, y1] = pt(a, r), [x2, y2] = pt(b, r);
        let d;
        if (v === sum) d = r0 ? `M${cx} ${cy - r}A${r} ${r} 0 1 1 ${cx - 0.01} ${cy - r}ZM${cx} ${cy - r0}A${r0} ${r0} 0 1 0 ${cx + 0.01} ${cy - r0}Z` : `M${cx} ${cy - r}A${r} ${r} 0 1 1 ${cx - 0.01} ${cy - r}Z`;
        else if (r0) {
          const [x3, y3] = pt(b, r0), [x4, y4] = pt(a, r0);
          d = `M${x1} ${y1}A${r} ${r} 0 ${big} 1 ${x2} ${y2}L${x3} ${y3}A${r0} ${r0} 0 ${big} 0 ${x4} ${y4}Z`;
        } else d = `M${cx} ${cy}L${x1} ${y1}A${r} ${r} 0 ${big} 1 ${x2} ${y2}Z`;
        parts.push(`<path d="${d}" fill="#${colorOf(m.series[0], 0, j)}" stroke="#fff" stroke-width="1"/>`);
        const lab = m.series[0].labels;
        if (lab && (lab.val || lab.pct || lab.cat)) {
          const [lx, ly] = pt((a + b) / 2, r0 ? (r + r0) / 2 : r * 0.65);
          const text = [lab.cat ? s.cats[j] : '', lab.val ? fmt(v) : '', lab.pct ? Math.round((v / sum) * 100) + '%' : ''].filter((x) => x !== '').join('; ');
          parts.push(`<text x="${lx}" y="${ly + 4}" text-anchor="middle" font-size="12" fill="#fff" ${FONT}>${esc(text)}</text>`);
        }
        a = b;
      });
      return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${parts.join('')}</svg>`;
    }

    const scatter = m.kind === 'scatter';
    const nCat = Math.max(0, ...data.map((s) => s.vals.length));
    const stacked = m.grouping === 'stacked' || m.grouping === 'percentStacked';
    const pct = m.grouping === 'percentStacked';
    const num = (v) => (typeof v === 'number' && isFinite(v) ? v : null);
    // Valores que se dibujan: apilados o en porcentaje si el gráfico lo pide.
    const shown = data.map((s) => s.vals.map(num));
    const base = data.map(() => Array(nCat).fill(0));
    if (stacked) {
      for (let j = 0; j < nCat; j++) {
        const tot = pct ? data.reduce((a, s, i) => a + Math.abs(shown[i][j] || 0), 0) || 1 : 1;
        let pos = 0, neg = 0;
        data.forEach((s, i) => {
          let v = shown[i][j] || 0;
          if (pct) v = (v / tot) * 100;
          if (v >= 0) { base[i][j] = pos; pos += v; } else { base[i][j] = neg; neg += v; }
          shown[i][j] = v;
        });
      }
    }
    const tops = [];
    data.forEach((s, i) => shown[i].forEach((v, j) => { if (v !== null) tops.push(base[i][j] + v, base[i][j]); }));
    if (!tops.length) tops.push(0, 1);
    const catLabels = (data[0] ? data[0].cats : []).map((c) => fmt(c));
    const vm = { ...m };
    if (pct) { vm.max = vm.max !== undefined ? vm.max : 100; vm.valFmt = vm.valFmt || '0%'; }
    const valFmt = (v) => (pct ? Math.round(v) + '%' : fmt(v, m.valFmt));
    const labelW = Math.max(...[Math.min(...tops), Math.max(...tops)].map((v) => valFmt(v).length)) * 7 + 10;
    // Ejes: valores a la izquierda (o abajo en barras horizontales) y categorías abajo.
    const horiz = m.kind === 'bar' && m.horizontal;
    const catW = horiz ? Math.min((right - left) / 3, Math.max(20, ...catLabels.map((c) => c.length * 6.5)) + 8) : 0;
    if (!m.valHidden && !horiz) left += labelW;
    if (horiz) left += catW;
    bottom -= m.catHidden && !horiz ? 4 : 20;
    if (horiz && m.valHidden) bottom += 16;
    right -= 8;
    const pw = right - left, ph = bottom - top;
    if (pw < 20 || ph < 20) return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">${parts.join('')}</svg>`;
    const sv = scale(Math.min(...tops), Math.max(...tops), horiz ? pw : ph, vm, !scatter || true);
    const vpos = (v) => (horiz ? left + ((v - sv.min) / (sv.max - sv.min)) * pw : bottom - ((v - sv.min) / (sv.max - sv.min)) * ph);
    for (const t of sv.ticks) {
      const p = vpos(t);
      if (m.grid !== false) parts.push(horiz ? `<line x1="${p}" y1="${top}" x2="${p}" y2="${bottom}" stroke="#${LINE}"/>` : `<line x1="${left}" y1="${p}" x2="${right}" y2="${p}" stroke="#${LINE}"/>`);
      if (!m.valHidden) parts.push(horiz ? `<text x="${p}" y="${bottom + 14}" text-anchor="middle" font-size="12" fill="#${INK}" ${FONT}>${esc(valFmt(t))}</text>`
        : `<text x="${left - 6}" y="${p + 4}" text-anchor="end" font-size="12" fill="#${INK}" ${FONT}>${esc(valFmt(t))}</text>`);
    }
    let xs = null;
    if (scatter) {
      const allX = data.flatMap((s) => s.cats.map(num)).filter((v) => v !== null);
      const useIndex = !allX.length;
      xs = data.map((s) => s.vals.map((_, j) => (useIndex ? j + 1 : num(s.cats[j]))));
      const flat = xs.flat().filter((v) => v !== null);
      const sx = scale(Math.min(...flat), Math.max(...flat), pw, {}, true);
      const xpos = (v) => left + ((v - sx.min) / (sx.max - sx.min)) * pw;
      for (const t of sx.ticks) {
        parts.push(`<line x1="${xpos(t)}" y1="${top}" x2="${xpos(t)}" y2="${bottom}" stroke="#${LINE}"/>`);
        if (!m.xHidden) parts.push(`<text x="${xpos(t)}" y="${bottom + 14}" text-anchor="middle" font-size="12" fill="#${INK}" ${FONT}>${esc(fmt(t))}</text>`);
      }
      data.forEach((s, i) => {
        const color = colorOf(m.series[i], i);
        const pts = s.vals.map((v, j) => (num(v) !== null && xs[i][j] !== null ? [xpos(xs[i][j]), vpos(v)] : null)).filter(Boolean);
        if (m.series[i].line && pts.length > 1) parts.push(`<polyline points="${pts.map((p) => p.join(',')).join(' ')}" fill="none" stroke="#${color}" stroke-width="2.5"/>`);
        if (m.series[i].marker || !m.series[i].line) for (const [x, y] of pts) parts.push(marker(x, y, m.series[i].symbol || 'circle', m.series[i].markerFill || color));
      });
      return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${parts.join('')}</svg>`;
    }
    // Eje de categorías.
    const zero = vpos(Math.max(sv.min, Math.min(sv.max, 0)));
    const band = (horiz ? ph : pw) / Math.max(1, nCat);
    const cpos = (j) => (horiz ? top + band * j : left + band * j);
    parts.push(horiz ? `<line x1="${zero}" y1="${top}" x2="${zero}" y2="${bottom}" stroke="#BFBFBF"/>` : `<line x1="${left}" y1="${zero}" x2="${right}" y2="${zero}" stroke="#BFBFBF"/>`);
    if (!m.catHidden) {
      const every = horiz ? 1 : Math.max(1, Math.ceil((Math.max(...catLabels.map((c) => c.length), 1) * 6.5) / band));
      catLabels.forEach((c, j) => {
        if (j % every) return;
        // En barras horizontales Excel pone la primera categoría abajo.
        parts.push(horiz ? `<text x="${left - 6}" y="${bottom - band * j - band / 2 + 4}" text-anchor="end" font-size="12" fill="#${INK}" ${FONT}>${esc(c)}</text>`
          : `<text x="${cpos(j) + band / 2}" y="${bottom + 15}" text-anchor="middle" font-size="12" fill="#${INK}" ${FONT}>${esc(c.length > 14 && every === 1 ? c.slice(0, 13) + '…' : c)}</text>`);
      });
    }
    const bars = data.map((s, i) => i).filter((i) => m.series[i].kind === 'bar');
    const others = data.map((s, i) => i).filter((i) => m.series[i].kind !== 'bar');
    if (bars.length) {
      const gap = m.gap / 100;
      const slots = stacked ? 1 : bars.length;
      const overlap = stacked ? 1 : m.overlap / 100;
      const barW = band / (gap + slots - overlap * (slots - 1));
      bars.forEach((i, k) => {
        const color = colorOf(m.series[i], i);
        shown[i].forEach((v, j) => {
          if (v === null) return;
          const off = (band - (barW * (slots - overlap * (slots - 1)))) / 2 + (stacked ? 0 : (horiz ? bars.length - 1 - k : k) * barW * (1 - overlap));
          const a = vpos(base[i][j]), b = vpos(base[i][j] + v);
          const fill = m.vary && data.length === 1 ? colorOf(m.series[i], i, j) : (m.series[i].points[j] || color);
          if (horiz) {
            const y = bottom - band * (j + 1) + off;
            parts.push(`<rect x="${Math.min(a, b)}" y="${y}" width="${Math.abs(b - a)}" height="${barW}" fill="#${fill}"/>`);
          } else parts.push(`<rect x="${cpos(j) + off}" y="${Math.min(a, b)}" width="${barW}" height="${Math.abs(b - a)}" fill="#${fill}"/>`);
          const lab = m.series[i].labels;
          if (lab && lab.val) parts.push(horiz ? `<text x="${b + 4}" y="${bottom - band * (j + 1) + off + barW / 2 + 4}" font-size="12" fill="#${INK}" ${FONT}>${esc(fmt(data[i].vals[j]))}</text>`
            : `<text x="${cpos(j) + off + barW / 2}" y="${Math.min(a, b) - 4}" text-anchor="middle" font-size="12" fill="#${INK}" ${FONT}>${esc(fmt(data[i].vals[j]))}</text>`);
        });
      });
    }
    for (const i of others) {
      const s = m.series[i], color = colorOf(s, i);
      const pts = shown[i].map((v, j) => (v === null ? null : [cpos(j) + band / 2, vpos(base[i][j] + v)]));
      if (s.kind === 'area') {
        const lower = shown[i].map((v, j) => [cpos(j) + band / 2, vpos(base[i][j])]).reverse();
        const upper = pts.map((p, j) => p || [cpos(j) + band / 2, vpos(base[i][j])]);
        parts.push(`<polygon points="${[...upper, ...lower].map((p) => p.join(',')).join(' ')}" fill="#${color}"/>`);
        continue;
      }
      let run = [];
      const flush = () => { if (run.length > 1) parts.push(`<polyline points="${run.map((p) => p.join(',')).join(' ')}" fill="none" stroke="#${color}" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"/>`); run = []; };
      for (const p of pts) { if (p) run.push(p); else flush(); }
      flush();
      if (s.marker) for (const p of pts) if (p) parts.push(marker(p[0], p[1], s.symbol, s.markerFill || color));
      if (s.labels && s.labels.val) pts.forEach((p, j) => { if (p) parts.push(`<text x="${p[0]}" y="${p[1] - 7}" text-anchor="middle" font-size="12" fill="#${INK}" ${FONT}>${esc(fmt(data[i].vals[j]))}</text>`); });
    }
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${parts.join('')}</svg>`;
  };

  // Datos de cada serie leídos de las celdas; si una referencia falla se usa lo que guardó Excel.
  // Celdas de una referencia de serie ('Hoja'!$A$1:$A$5), con la marca de si están ocultas.
  const cellsOf = (book, sheets, f) => {
    const m = /^=?(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(f.trim());
    if (!m) return null;
    const name = (m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]).toLowerCase();
    const sh = sheets.find((s) => s.name.toLowerCase() === name);
    const box = sh && O.parseRanges(m[3].replace(/\$/g, ''))[0];
    if (!box) return null;
    const out = [];
    for (let r = box.r1; r <= box.r2; r++) for (let c = box.c1; c <= box.c2; c++) {
      const v = book.value(sh, r, c);
      out.push({ v: v && v.err ? null : v, hidden: !!((sh.hiddenRows && sh.hiddenRows[r]) || (sh.hiddenCols && sh.hiddenCols[c])), text: book.display(sh, r, c) });
    }
    return out;
  };
  O.chartData = (book, sheets, m, refs) => {
    const read = (part) => {
      const f = part.ref >= 0 ? refs[part.ref] : null;
      const got = f && cellsOf(book, sheets, f);
      return got || (part.cache || []).map((v) => ({ v, hidden: false, text: v == null ? '' : String(v) }));
    };
    const all = m.series.map((s) => ({ s, vals: read(s.val), cats: read(s.cat) }));
    // Un punto se quita si su celda de valores o de categoría está oculta.
    const hide = (list, j) => m.visibleOnly && list.some((x) => (x.vals[j] && x.vals[j].hidden) || (x.cats[j] && x.cats[j].hidden));
    return all.map((x) => {
      const keep = x.vals.map((_, j) => !hide([x], j));
      const vals = x.vals.filter((_, j) => keep[j]).map((p) => (typeof p.v === 'number' ? p.v : typeof p.v === 'string' && p.v !== '' && !isNaN(p.v) ? +p.v : null));
      let cats = x.cats.filter((_, j) => keep[j]).map((p) => (p.text !== undefined && p.text !== '' ? p.text : p.v));
      if (!x.cats.length) cats = vals.map((_, j) => j + 1);
      if (m.kind === 'scatter') cats = x.cats.filter((_, j) => keep[j]).map((p) => p.v);
      const nameCell = x.s.name.ref >= 0 ? read(x.s.name)[0] : null;
      const name = nameCell ? nameCell.text || O.valueText(nameCell.v ?? '') : x.s.name.v;
      return { name: name || `Serie${m.series.indexOf(x.s) + 1}`, cats, vals };
    });
  };
  O.chartTitle = (book, sheets, m, refs, data) => {
    if (!m.title) return null;
    if (m.title.ref >= 0 && refs[m.title.ref]) {
      const got = cellsOf(book, sheets, refs[m.title.ref]);
      return got && got[0] ? got[0].text : m.title.v;
    }
    if (m.title.auto) return m.series.length === 1 && data[0] ? data[0].name : (m.kind === 'pie' ? '' : 'Título del gráfico');
    return m.title.v;
  };

  // XML de un gráfico nuevo, con el estilo por defecto de Excel. kind: col, bar, line, pie, area, scatter, doughnut.
  O.chartXml = (kind, series, title) => {
    const q = (f) => O.esc(f);
    const tag = { col: 'barChart', bar: 'barChart', line: 'lineChart', pie: 'pieChart', doughnut: 'doughnutChart', area: 'areaChart', scatter: 'scatterChart' }[kind];
    const txt = (t, sz) => `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}" b="0"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="es-ES" sz="${sz}" b="0"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:rPr><a:t>${O.esc(t)}</a:t></a:r></a:p></c:rich></c:tx>`;
    const sers = series.map((s, i) => {
      const color = `<a:schemeClr val="accent${(i % 6) + 1}"/>`;
      const sp = kind === 'line' || kind === 'scatter'
        ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill>${color}</a:solidFill><a:round/></a:ln></c:spPr>`
        : `<c:spPr><a:solidFill>${color}</a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>`;
      const name = s.name ? `<c:tx><c:strRef><c:f>${q(s.name)}</c:f></c:strRef></c:tx>` : '';
      const marker = kind === 'line' ? '<c:marker><c:symbol val="none"/></c:marker>' : kind === 'scatter' ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>' : '';
      const cat = s.cat ? (kind === 'scatter' ? `<c:xVal><c:numRef><c:f>${q(s.cat)}</c:f></c:numRef></c:xVal>` : `<c:cat><c:strRef><c:f>${q(s.cat)}</c:f></c:strRef></c:cat>`) : '';
      const val = kind === 'scatter' ? `<c:yVal><c:numRef><c:f>${q(s.val)}</c:f></c:numRef></c:yVal>` : `<c:val><c:numRef><c:f>${q(s.val)}</c:f></c:numRef></c:val>`;
      const sp2 = kind === 'scatter' ? '<c:spPr><a:ln w="25400" cap="rnd"><a:noFill/><a:round/></a:ln></c:spPr>' : sp;
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${name}${sp2}${marker}${cat}${val}${kind === 'line' || kind === 'scatter' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
    }).join('');
    const axes = '<c:axId val="500000001"/><c:axId val="500000002"/>';
    const gridline = '<c:majorGridlines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill><a:round/></a:ln></c:spPr></c:majorGridlines>';
    const axText = '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="es-ES"/></a:p></c:txPr>';
    const axLine = (on) => `<c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${on ? '<a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill>' : '<a:noFill/>'}<a:round/></a:ln></c:spPr>`;
    const horizontal = kind === 'bar';
    const catAx = kind === 'scatter'
      ? `<c:valAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${gridline}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axLine(false)}${axText}<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
      : `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axLine(true)}${axText}<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const valAx = `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'b' : 'l'}"/>${gridline}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axLine(false)}${axText}<c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="${kind === 'scatter' ? 'midCat' : 'between'}"/></c:valAx>`;
    let group;
    if (kind === 'pie') group = `<c:pieChart><c:varyColors val="1"/>${sers.replace(/<c:spPr>[\s\S]*?<\/c:spPr>/g, '<c:spPr><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></c:spPr>')}<c:firstSliceAng val="0"/></c:pieChart>`;
    else if (kind === 'doughnut') group = `<c:doughnutChart><c:varyColors val="1"/>${sers.replace(/<c:spPr>[\s\S]*?<\/c:spPr>/g, '<c:spPr><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></c:spPr>')}<c:firstSliceAng val="0"/><c:holeSize val="75"/></c:doughnutChart>`;
    else if (tag === 'barChart') group = `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="219"/><c:overlap val="-27"/>${axes}</c:barChart>`;
    else if (tag === 'lineChart') group = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${axes}</c:lineChart>`;
    else if (tag === 'areaChart') group = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${axes}</c:areaChart>`;
    else group = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}${axes}</c:scatterChart>`;
    const round = kind === 'pie' || kind === 'doughnut';
    const head = title !== null && title !== undefined
      ? `<c:title>${title.ref ? `<c:tx><c:strRef><c:f>${q(title.ref)}</c:f></c:strRef></c:tx>` : txt(title.text || 'Título del gráfico', 1400)}<c:overlay val="0"/>${title.ref ? '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="es-ES"/></a:p></c:txPr>' : ''}</c:title><c:autoTitleDeleted val="0"/>`
      : '<c:autoTitleDeleted val="1"/>';
    return `${O.XML}<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>${head}<c:plotArea><c:layout/>${group}${round ? '' : catAx + valAx}</c:plotArea><c:legend><c:legendPos val="${round || series.length > 1 ? 'b' : 'b'}"/><c:overlay val="0"/>${axText}</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill><a:round/></a:ln></c:spPr></c:chartSpace>`;
  };

  // Cambia los <c:f> del XML por las referencias actuales, en el mismo orden en que se leyeron.
  O.chartWithRefs = (xml, refs) => {
    let i = 0;
    return xml.replace(/<((?:[\w]+:)?f)>([\s\S]*?)<\/\1>/g, (all, tag, body) => {
      const k = i++;
      return refs[k] !== undefined ? `<${tag}>${O.esc(refs[k])}</${tag}>` : all;
    });
  };
  O.chartRefs = (xml) => {
    const doc = O.parseXml(xml);
    return O.all(doc.documentElement, 'f').map((f) => f.textContent);
  };
})(globalThis);
