// Formato condicional: decide qué formato extra lleva cada celda según las reglas de su hoja.
// Una regla es { ref, type, ... , style: { fill, color, b, i, u, strike } } y el orden de sheet.cf es su prioridad.
(function (g) {
  const O = g.Oficina;

  // "A1:B5 D2" -> [{ r1, c1, r2, c2 }]
  O.parseRanges = (ref) => String(ref || '').trim().split(/\s+/).map((part) => {
    const [a, b] = part.split(':').map((x) => O.parseAddr(x));
    if (!a) return null;
    const e = b || a;
    return { r1: Math.min(a.r, e.r), c1: Math.min(a.c, e.c), r2: Math.max(a.r, e.r), c2: Math.max(a.c, e.c) };
  }).filter(Boolean);
  O.rangesText = (ranges) => ranges.map((x) => (x.r1 === x.r2 && x.c1 === x.c2 ? O.addr(x.r1, x.c1) : `${O.addr(x.r1, x.c1)}:${O.addr(x.r2, x.c2)}`)).join(' ');

  const isBlank = (v) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  const isErr = (v) => !!(v && v.err);
  const truthy = (v) => v === true || (typeof v === 'number' && v !== 0);
  const key = (v) => (typeof v === 'number' ? 'n' + v : typeof v === 'boolean' ? 'b' + v : 's' + String(v).toLowerCase());

  const hexToRgb = (h) => [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
  const mix = (a, b, t) => {
    const x = hexToRgb(a), y = hexToRgb(b);
    return x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('').toUpperCase();
  };
  // Percentil inclusivo, como PERCENTIL de Excel.
  const percentile = (sorted, p) => {
    if (!sorted.length) return 0;
    const pos = (sorted.length - 1) * p, lo = Math.floor(pos);
    return sorted[lo] + (sorted[Math.min(lo + 1, sorted.length - 1)] - sorted[lo]) * (pos - lo);
  };

  function makeTest(rule, book, sheet, cells, anchor) {
    const val = (r, c) => book.value(sheet, r, c);
    // Las formulas de la regla son relativas a la primera celda del rango, como en Excel.
    const formula = (src, r, c) => book.evaluate(sheet, O.shiftFormula(src, r - anchor.r1, c - anchor.c1));
    const style = rule.style || {};
    const numbers = () => cells.map(([r, c]) => val(r, c)).filter((v) => typeof v === 'number');
    switch (rule.type) {
      case 'cellIs': {
        const f = rule.f || [];
        return (r, c) => {
          const v = val(r, c);
          if (isErr(v)) return null;
          const a = formula(f[0] || '0', r, c);
          if (isErr(a)) return null;
          const cmp = (x) => O.compareValues(v, x);
          let hit;
          if (rule.op === 'between' || rule.op === 'notBetween') {
            const b = formula(f[1] || '0', r, c);
            if (isErr(b)) return null;
            const lo = O.compareValues(a, b) <= 0 ? a : b, hi = lo === a ? b : a;
            const inside = cmp(lo) >= 0 && cmp(hi) <= 0;
            hit = rule.op === 'between' ? inside : !inside;
          } else {
            const k = cmp(a);
            hit = { equal: k === 0, notEqual: k !== 0, greaterThan: k > 0, lessThan: k < 0, greaterThanOrEqual: k >= 0, lessThanOrEqual: k <= 0 }[rule.op];
          }
          return hit ? style : null;
        };
      }
      case 'expression':
        return (r, c) => (truthy(formula((rule.f || [])[0] || 'FALSE', r, c)) ? style : null);
      case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
        const t = String(rule.text || '').toLowerCase();
        return (r, c) => {
          const v = val(r, c);
          if (isErr(v)) return null;
          const s = O.valueText(v).toLowerCase();
          const hit = rule.type === 'containsText' ? s.includes(t) : rule.type === 'notContainsText' ? !s.includes(t)
            : rule.type === 'beginsWith' ? s.startsWith(t) : s.endsWith(t);
          return hit ? style : null;
        };
      }
      case 'containsBlanks': return (r, c) => (isBlank(val(r, c)) ? style : null);
      case 'notContainsBlanks': return (r, c) => (!isBlank(val(r, c)) ? style : null);
      case 'containsErrors': return (r, c) => (isErr(val(r, c)) ? style : null);
      case 'notContainsErrors': return (r, c) => (!isErr(val(r, c)) ? style : null);
      case 'duplicateValues': case 'uniqueValues': {
        const counts = new Map();
        for (const [r, c] of cells) {
          const v = val(r, c);
          if (!isBlank(v) && !isErr(v)) counts.set(key(v), (counts.get(key(v)) || 0) + 1);
        }
        return (r, c) => {
          const v = val(r, c);
          if (isBlank(v) || isErr(v)) return null;
          const dup = counts.get(key(v)) > 1;
          return (rule.type === 'duplicateValues' ? dup : !dup) ? style : null;
        };
      }
      case 'top10': {
        const sorted = numbers().sort((x, y) => (rule.bottom ? x - y : y - x));
        if (!sorted.length) return () => null;
        const n = rule.percent ? Math.max(1, Math.floor((sorted.length * (rule.rank || 10)) / 100)) : rule.rank || 10;
        const edge = sorted[Math.min(n, sorted.length) - 1];
        return (r, c) => {
          const v = val(r, c);
          if (typeof v !== 'number') return null;
          return (rule.bottom ? v <= edge : v >= edge) ? style : null;
        };
      }
      case 'aboveAverage': {
        const list = numbers();
        if (!list.length) return () => null;
        const avg = list.reduce((x, y) => x + y, 0) / list.length;
        const sd = Math.sqrt(list.reduce((s, x) => s + (x - avg) ** 2, 0) / list.length);
        const above = rule.above !== false;
        const edge = rule.std ? avg + (above ? 1 : -1) * rule.std * sd : avg;
        return (r, c) => {
          const v = val(r, c);
          if (typeof v !== 'number') return null;
          const hit = above ? (rule.equal ? v >= edge : v > edge) : rule.equal ? v <= edge : v < edge;
          return hit ? style : null;
        };
      }
      case 'timePeriod': {
        const t = new Date();
        const today = O.dateToSerial(t.getFullYear(), t.getMonth() + 1, t.getDate());
        const wd = t.getDay();
        const month = (k) => {
          const first = new Date(Date.UTC(t.getFullYear(), t.getMonth() + k, 1));
          const next = new Date(Date.UTC(t.getFullYear(), t.getMonth() + k + 1, 1));
          return [O.dateToSerial(first.getUTCFullYear(), first.getUTCMonth() + 1, 1), O.dateToSerial(next.getUTCFullYear(), next.getUTCMonth() + 1, 1) - 1];
        };
        const span = {
          today: [today, today], yesterday: [today - 1, today - 1], tomorrow: [today + 1, today + 1],
          last7Days: [today - 6, today], thisWeek: [today - wd, today - wd + 6], lastWeek: [today - wd - 7, today - wd - 1],
          nextWeek: [today - wd + 7, today - wd + 13], thisMonth: month(0), lastMonth: month(-1), nextMonth: month(1),
        }[rule.period];
        if (!span) return () => null;
        return (r, c) => {
          const v = val(r, c);
          return typeof v === 'number' && Math.floor(v) >= span[0] && Math.floor(v) <= span[1] ? style : null;
        };
      }
      case 'colorScale': case 'dataBar': case 'iconSet': {
        const sorted = numbers().sort((x, y) => x - y);
        if (!sorted.length) return () => null;
        const lo = sorted[0], hi = sorted[sorted.length - 1];
        const point = (p, fallback) => {
          if (!p) return fallback;
          const n = parseFloat(p.val);
          if (p.type === 'min') return lo;
          if (p.type === 'max') return hi;
          if (p.type === 'percent') return lo + ((hi - lo) * n) / 100;
          if (p.type === 'percentile') return percentile(sorted, n / 100);
          if (p.type === 'formula' || p.type === 'num') {
            const v = Number.isNaN(n) || p.type === 'formula' ? book.evaluate(sheet, String(p.val)) : n;
            return typeof v === 'number' ? v : fallback;
          }
          return fallback;
        };
        if (rule.type === 'iconSet') {
          const set = rule.icons || {};
          const name = set.set || '3TrafficLights1';
          const n = +name[0] || 3;
          const pts = (set.points && set.points.length === n ? set.points : O.iconDefaults(n)).map((p) => ({ at: point(p, lo), gte: p.gte !== false }));
          return (r, c) => {
            const v = val(r, c);
            if (typeof v !== 'number') return null;
            // Gana el umbral más alto que se alcanza; el primero es siempre el mínimo.
            let i = 0;
            for (let k = n - 1; k >= 1; k--) if (pts[k].gte ? v >= pts[k].at : v > pts[k].at) { i = k; break; }
            if (set.reverse) i = n - 1 - i;
            return { icon: { set: name, index: i }, hideValue: set.showValue === false || undefined };
          };
        }
        if (rule.type === 'dataBar') {
          const bar = rule.bar || {};
          const a = point(bar.min, lo), b = point(bar.max, hi);
          return (r, c) => {
            const v = val(r, c);
            if (typeof v !== 'number') return null;
            const pct = b > a ? Math.max(0, Math.min(1, (v - a) / (b - a))) : 1;
            return { bar: { pct, color: bar.color || '638EC6' } };
          };
        }
        const stops = (rule.scale || []).map((s) => ({ at: point(s, lo), color: s.color || 'FFFFFF' }));
        if (stops.length < 2) return () => null;
        return (r, c) => {
          const v = val(r, c);
          if (typeof v !== 'number') return null;
          if (v <= stops[0].at) return { fill: stops[0].color };
          for (let i = 1; i < stops.length; i++) {
            if (v <= stops[i].at) {
              const span = stops[i].at - stops[i - 1].at;
              return { fill: mix(stops[i - 1].color, stops[i].color, span > 0 ? (v - stops[i - 1].at) / span : 1) };
            }
          }
          return { fill: stops[stops.length - 1].color };
        };
      }
      default:
        return () => null;
    }
  }

  // Devuelve un Map "r,c" -> { fill, color, b, i, u, strike, bar }. maxR y maxC limitan los rangos enormes.
  O.conditionalFormats = (book, sheet, maxR, maxC) => {
    const out = new Map();
    const stopped = new Set();
    for (const rule of sheet.cf || []) {
      const ranges = O.parseRanges(rule.ref);
      if (!ranges.length) continue;
      const cells = [];
      for (const rg of ranges) {
        for (let r = rg.r1; r <= Math.min(rg.r2, maxR); r++) for (let c = rg.c1; c <= Math.min(rg.c2, maxC); c++) cells.push([r, c]);
      }
      const test = makeTest(rule, book, sheet, cells, ranges[0]);
      for (const [r, c] of cells) {
        const k = r + ',' + c;
        if (stopped.has(k)) continue;
        const res = test(r, c);
        if (!res) continue;
        const cur = out.get(k) || {};
        // Si dos reglas se cumplen, en cada propiedad manda la de más prioridad.
        for (const [p, v] of Object.entries(res)) if (v !== undefined && cur[p] === undefined) cur[p] = v;
        out.set(k, cur);
        if (rule.stop) stopped.add(k);
      }
    }
    return out;
  };

  // Umbrales por defecto de Excel para 3, 4 y 5 iconos: porcentajes repartidos a partes iguales.
  O.iconDefaults = (n) => Array.from({ length: n }, (_, i) => ({ type: 'percent', val: String(Math.round((100 * i) / n)) }));
  // Iconos de cada conjunto, del valor más bajo al más alto: [forma, color]. Los colores son los de Excel.
  const R = 'D9512C', Y = 'EAC27C', G = '5BA08A', K = '505050', A = '8C8C8C', B = '3E6DB5';
  const many = (shape, n, color) => Array.from({ length: n }, (_, i) => [shape + i, color]);
  O.ICON_SETS = {
    '3Arrows': [['down', R], ['right', Y], ['up', G]], '3ArrowsGray': [['down', A], ['right', A], ['up', A]],
    '3Flags': [['flag', R], ['flag', Y], ['flag', G]], '3TrafficLights1': [['dot', R], ['dot', Y], ['dot', G]],
    '3TrafficLights2': [['light', R], ['light', Y], ['light', G]], '3Signs': [['diamond', R], ['sign', Y], ['dot', G]],
    '3Symbols': [['xc', R], ['!c', Y], ['okc', G]], '3Symbols2': [['x', R], ['!', Y], ['ok', G]],
    '3Stars': [['star0', Y], ['star1', Y], ['star2', Y]], '3Triangles': [['tdown', R], ['dash', Y], ['tup', G]],
    '4Arrows': [['down', R], ['dr', Y], ['ur', Y], ['up', G]], '4ArrowsGray': [['down', A], ['dr', A], ['ur', A], ['up', A]],
    '4RedToBlack': [['dot', K], ['dot', A], ['dot', 'E9A3A3'], ['dot', R]], '4Rating': many('bars', 5, B).slice(1),
    '4TrafficLights': [['dot', K], ['dot', R], ['dot', Y], ['dot', G]],
    '5Arrows': [['down', R], ['dr', Y], ['right', Y], ['ur', Y], ['up', G]], '5ArrowsGray': [['down', A], ['dr', A], ['right', A], ['ur', A], ['up', A]],
    '5Rating': many('bars', 5, B), '5Quarters': many('pie', 5, K), '5Boxes': many('box', 5, B),
  };
  // Dibujo SVG de 16x16 de cada forma.
  O.iconSvg = (shape, color) => {
    const c = '#' + color, w = '#fff', grey = '#c8c8c8';
    const arrow = (deg) => `<path d="M8 1 14.5 8H10.5V15H5.5V8H1.5Z" fill="${c}" transform="rotate(${deg} 8 8)"/>`;
    const star = 'M8 1 10.1 5.6 15 6.1 11.3 9.4 12.4 14.3 8 11.8 3.6 14.3 4.7 9.4 1 6.1 5.9 5.6Z';
    const n = +shape.slice(-1);
    const body = {
      up: arrow(0), ur: arrow(45), right: arrow(90), dr: arrow(135), down: arrow(180),
      dot: `<circle cx="8" cy="8" r="6.5" fill="${c}"/>`,
      light: `<rect x="1" y="1" width="14" height="14" rx="3" fill="#333"/><circle cx="8" cy="8" r="5" fill="${c}"/>`,
      diamond: `<path d="M8 1 15 8 8 15 1 8Z" fill="${c}"/>`, sign: `<path d="M8 1.5 15 14.5H1Z" fill="${c}"/>`,
      xc: `<circle cx="8" cy="8" r="7" fill="${c}"/><path d="M5 5 11 11M11 5 5 11" stroke="${w}" stroke-width="2"/>`,
      '!c': `<circle cx="8" cy="8" r="7" fill="${c}"/><path d="M8 3.5V9.5M8 11v1.8" stroke="${w}" stroke-width="2"/>`,
      okc: `<circle cx="8" cy="8" r="7" fill="${c}"/><path d="M4.5 8.2 7 10.7 11.5 5.5" fill="none" stroke="${w}" stroke-width="2"/>`,
      x: `<path d="M3 3 13 13M13 3 3 13" stroke="${c}" stroke-width="3"/>`, '!': `<path d="M8 1.5V10M8 12v2.5" stroke="${c}" stroke-width="3"/>`,
      ok: `<path d="M2 8.5 6 12.5 14 3.5" fill="none" stroke="${c}" stroke-width="3"/>`,
      flag: `<path d="M3 1V15" stroke="#555" stroke-width="1.5"/><path d="M3.8 1.5H14L11.5 5 14 8.5H3.8Z" fill="${c}"/>`,
      tup: `<path d="M8 3 15 13H1Z" fill="${c}"/>`, tdown: `<path d="M8 13 15 3H1Z" fill="${c}"/>`, dash: `<rect x="2" y="6" width="12" height="4" fill="${c}"/>`,
    }[shape];
    let svg = body;
    if (!svg && shape.startsWith('star')) {
      svg = `<path d="${star}" fill="${grey}" stroke="${c}"/>`;
      if (n === 2) svg = `<path d="${star}" fill="${c}"/>`;
      if (n === 1) svg += `<path d="${star}" fill="${c}" style="clip-path:inset(0 50% 0 0)"/>`;
    }
    if (!svg && shape.startsWith('bars')) svg = [0, 1, 2, 3].map((i) => `<rect x="${1 + i * 3.6}" y="${11 - i * 3}" width="2.8" height="${4 + i * 3}" fill="${i < n ? c : grey}"/>`).join('');
    if (!svg && shape.startsWith('pie')) {
      const a = (n / 4) * 2 * Math.PI, x = 8 + 6.5 * Math.sin(a), y = 8 - 6.5 * Math.cos(a);
      const fill = n === 4 ? `<circle cx="8" cy="8" r="6.5" fill="${c}"/>` : n ? `<path d="M8 8V1.5A6.5 6.5 0 0 1 ${x.toFixed(2)} ${y.toFixed(2)}Z" fill="${c}"/>` : '';
      svg = `<circle cx="8" cy="8" r="6.5" fill="#fff" stroke="${c}"/>${fill}`;
    }
    if (!svg && shape.startsWith('box')) svg = [[1, 1], [8.5, 1], [8.5, 8.5], [1, 8.5]].map(([x, y], i) => `<rect x="${x}" y="${y}" width="6.5" height="6.5" fill="${i < n ? c : grey}"/>`).join('');
    return `<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true">${svg || ''}</svg>`;
  };
  const OPS = {
    equal: 'igual a', notEqual: 'distinto de', greaterThan: 'mayor que', lessThan: 'menor que',
    greaterThanOrEqual: 'mayor o igual que', lessThanOrEqual: 'menor o igual que', between: 'entre', notBetween: 'no entre',
  };
  const PERIODS = {
    today: 'hoy', yesterday: 'ayer', tomorrow: 'mañana', last7Days: 'en los últimos 7 días', thisWeek: 'esta semana',
    lastWeek: 'la semana pasada', nextWeek: 'la semana próxima', thisMonth: 'este mes', lastMonth: 'el mes pasado', nextMonth: 'el mes próximo',
  };
  O.cfOperators = OPS;
  O.cfPeriods = PERIODS;
  const shown = (f) => {
    const m = /^"((?:[^"]|"")*)"$/.exec(f || '');
    return m ? m[1].replace(/""/g, '"') : '=' + O.toLocalFormula(f || '');
  };
  const plainNumber = (f) => (/^-?\d+(\.\d+)?$/.test(f || '') ? String(f).replace('.', ',') : shown(f));
  // Texto que describe la regla, para el cuadro de administrar reglas.
  O.describeRule = (rule) => {
    const f = rule.f || [];
    switch (rule.type) {
      case 'cellIs':
        return `Valor de celda ${OPS[rule.op] || rule.op} ${plainNumber(f[0])}${rule.op === 'between' || rule.op === 'notBetween' ? ' y ' + plainNumber(f[1]) : ''}`;
      case 'expression': return `Fórmula: =${O.toLocalFormula(f[0] || '')}`;
      case 'containsText': return `Texto que contiene "${rule.text}"`;
      case 'notContainsText': return `Texto que no contiene "${rule.text}"`;
      case 'beginsWith': return `Texto que empieza por "${rule.text}"`;
      case 'endsWith': return `Texto que termina en "${rule.text}"`;
      case 'containsBlanks': return 'Celdas en blanco';
      case 'notContainsBlanks': return 'Celdas que no están en blanco';
      case 'containsErrors': return 'Celdas con errores';
      case 'notContainsErrors': return 'Celdas sin errores';
      case 'duplicateValues': return 'Valores duplicados';
      case 'uniqueValues': return 'Valores únicos';
      case 'top10': return `${rule.bottom ? 'Inferiores' : 'Superiores'}: ${rule.rank}${rule.percent ? ' %' : ''}`;
      case 'aboveAverage': return `${rule.above === false ? 'Por debajo' : 'Por encima'} del promedio${rule.equal ? ' o igual' : ''}`;
      case 'timePeriod': return `Fecha ${PERIODS[rule.period] || rule.period}`;
      case 'colorScale': return `Escala de ${(rule.scale || []).length} colores`;
      case 'dataBar': return 'Barra de datos';
      case 'iconSet': return `Conjunto de iconos (${(rule.icons || {}).set || '3TrafficLights1'})`;
      default: return 'Regla de Excel que Oficina no aplica';
    }
  };

  // Quita un rectángulo de celdas de las reglas, como "Borrar reglas de las celdas seleccionadas" de Excel.
  // Si cambia la primera celda del rango, las fórmulas se desplazan para seguir apuntando a lo mismo.
  const subtract = (a, b) => {
    if (b.r1 > a.r2 || b.r2 < a.r1 || b.c1 > a.c2 || b.c2 < a.c1) return [a];
    const out = [];
    if (a.r1 < b.r1) out.push({ ...a, r2: b.r1 - 1 });
    const r1 = Math.max(a.r1, b.r1), r2 = Math.min(a.r2, b.r2);
    if (a.c1 < b.c1) out.push({ r1, r2, c1: a.c1, c2: b.c1 - 1 });
    if (a.c2 > b.c2) out.push({ r1, r2, c1: b.c2 + 1, c2: a.c2 });
    if (a.r2 > b.r2) out.push({ ...a, r1: b.r2 + 1 });
    return out;
  };
  O.removeRuleCells = (rules, box) => (rules || []).map((rule) => {
    const before = O.parseRanges(rule.ref);
    const after = before.flatMap((x) => subtract(x, box));
    if (!after.length) return null;
    const next = { ...rule, ref: O.rangesText(after) };
    if (rule.f && before.length) {
      const dr = after[0].r1 - before[0].r1, dc = after[0].c1 - before[0].c1;
      if (dr || dc) next.f = rule.f.map((x) => O.shiftFormula(x, dr, dc));
    }
    return next;
  }).filter(Boolean);

  // ---------- Validación de datos ----------
  // Regla de validación que afecta a una celda, con la esquina de su rango para mover las referencias relativas.
  O.validationAt = (sheet, r, c) => {
    for (const rule of sheet.dv || []) {
      const ranges = O.parseRanges(rule.ref);
      if (ranges.some((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2)) return { rule, origin: ranges[0] };
    }
    return null;
  };
  const dvFormula = (hit, i, r, c) => O.shiftFormula((hit.rule.f || [])[i] || '', r - hit.origin.r1, c - hit.origin.c1);
  // Opciones de una lista: escritas en la regla ("a,b,c") o tomadas de un rango.
  O.validationList = (book, sheet, hit, r, c) => {
    const f = (hit.rule.f || [])[0] || '';
    const lit = /^"((?:[^"]|"")*)"$/.exec(f.trim());
    if (lit) return lit[1].replace(/""/g, '"').split(',').map((x) => x.trim()).filter((x) => x !== '');
    return book.evaluateAll(sheet, dvFormula(hit, 0, r, c))
      .filter((v) => v !== null && v !== undefined && v !== '' && !(v && v.err))
      .map((v) => O.valueText(v));
  };
  // ¿Cumple el valor la regla? Las celdas vacías siempre valen, como al borrar en Excel.
  O.validates = (book, sheet, hit, value, r, c) => {
    if (value === null || value === undefined || value === '') return true;
    const a = hit.rule.attrs || {};
    const type = a.type || 'none';
    if (type === 'none') return true;
    if (type === 'list') {
      const text = O.valueText(value).toLowerCase();
      return O.validationList(book, sheet, hit, r, c).some((x) => x.toLowerCase() === text || (typeof value === 'number' && (O.readNumber(x) || {}).v === value));
    }
    if (type === 'custom') {
      // Excel evalúa la fórmula como si el valor nuevo ya estuviera en la celda.
      const key = O.addr(r, c), prev = sheet.cells[key];
      sheet.cells[key] = { ...(prev || {}), v: value };
      delete sheet.cells[key].f;
      book.clearCache();
      try { return book.evaluate(sheet, dvFormula(hit, 0, r, c)) === true; } finally {
        if (prev) sheet.cells[key] = prev; else delete sheet.cells[key];
        book.clearCache();
      }
    }
    let n = value;
    if (type === 'textLength') n = O.valueText(value).length;
    else if (typeof n !== 'number') return false;
    if (type === 'whole' && !Number.isInteger(n)) return false;
    const bound = (i) => {
      const v = book.evaluate(sheet, dvFormula(hit, i, r, c));
      return typeof v === 'number' ? v : NaN;
    };
    const x = bound(0);
    switch (a.operator || 'between') {
      case 'between': return n >= x && n <= bound(1);
      case 'notBetween': return n < x || n > bound(1);
      case 'equal': return n === x;
      case 'notEqual': return n !== x;
      case 'greaterThan': return n > x;
      case 'lessThan': return n < x;
      case 'greaterThanOrEqual': return n >= x;
      case 'lessThanOrEqual': return n <= x;
      default: return true;
    }
  };

  // Ajusta rangos y formulas de las reglas al insertar o eliminar filas o columnas. Quita las que se quedan sin celdas.
  O.adjustRules = (rules, opts) => (rules || []).map((rule) => {
    const parts = String(rule.ref).trim().split(/\s+/).map((p) => O.adjustRefs(p, { ...opts, local: true })).filter((p) => !p.includes('#REF!'));
    if (!parts.length) return null;
    const next = { ...rule, ref: parts.join(' ') };
    if (rule.f) next.f = rule.f.map((x) => O.adjustRefs(x, { ...opts, local: true }));
    return next;
  }).filter(Boolean);
})(globalThis);
