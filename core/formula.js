// Motor de formulas y modelo de libro de la hoja de calculo. No depende del navegador.
(function (g) {
  const O = (g.Oficina = g.Oficina || {});

  class CellError extends Error {
    constructor(code) { super(code); this.code = code; }
  }
  const fail = (code) => { throw new CellError(code); };

  O.colName = (c) => {
    let s = '';
    for (c += 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
    return s;
  };
  O.colIndex = (name) => [...name.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  O.addr = (r, c) => O.colName(c) + (r + 1);
  O.parseAddr = (a) => {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a);
    return m ? { c: O.colIndex(m[1]), r: +m[2] - 1 } : null;
  };

  // ---------- Tablas ----------
  // Referencias estructuradas: Tabla1[Ventas], Tabla1[[#Totals],[Ventas]], [@Ventas]...
  // En el archivo van en inglés y con el nombre de la tabla; en pantalla, como en un Excel en español.
  const BRACKETS = "\\[(?:[^[\\]']|'.|\\[(?:[^[\\]']|'.)*\\])*\\]";
  const SREF = `(?:[A-Za-z_À-ɏ][\\w.À-ɏ]*)?${BRACKETS}`;
  const SPECIAL = {
    '#all': 'All', '#todo': 'All', '#data': 'Data', '#datos': 'Data', '#headers': 'Headers', '#encabezados': 'Headers',
    '#totals': 'Totals', '#totales': 'Totals', '#this row': 'This Row', '#esta fila': 'This Row',
  };
  const SPECIAL_ES = { All: '#Todo', Data: '#Datos', Headers: '#Encabezados', Totals: '#Totales', 'This Row': '#Esta fila' };
  const escapeCol = (s) => s.replace(/['#[\]]/g, "'$&");
  O.splitTableRef = (text) => {
    const i = text.indexOf('[');
    const table = text.slice(0, i) || null;
    let inner = text.slice(i + 1, -1).trim();
    const areas = [], cols = [];
    if (inner.startsWith('@')) { areas.push('This Row'); inner = inner.slice(1).trim(); }
    const items = [];
    if (!inner.includes('[')) { if (inner) items.push(inner); } else for (const m of inner.matchAll(/\[((?:[^[\]']|'.)*)\]|:/g)) items.push(m[0] === ':' ? ':' : m[1]);
    for (let k = 0; k < items.length; k++) {
      const it = items[k];
      if (it === ':') continue;
      const sp = it.trim().startsWith('#') && SPECIAL[it.trim().toLowerCase()];
      if (sp) areas.push(sp);
      else if (items[k + 1] === ':' && items[k + 2] !== undefined) { cols.push(it.replace(/'(.)/g, '$1'), items[k + 2].replace(/'(.)/g, '$1')); k += 2; } else cols.push(it.replace(/'(.)/g, '$1'));
    }
    return { table, areas, cols };
  };
  const joinTableRef = (table, areas, cols, local) => {
    const colPart = cols.length === 2 ? `[${escapeCol(cols[0])}]:[${escapeCol(cols[1])}]` : cols.length ? `[${escapeCol(cols[0])}]` : '';
    const name = table || '';
    if (local && areas.length === 1 && areas[0] === 'This Row') {
      if (!cols.length) return name + '[@]';
      return cols.length === 1 && /^[\wÀ-ɏ.]+$/.test(cols[0]) ? `${name}[@${cols[0]}]` : `${name}[@${colPart}]`;
    }
    const sp = areas.map((a) => `[${local ? SPECIAL_ES[a] : '#' + a}]`);
    if (!sp.length && cols.length === 1) return name + colPart;
    if (sp.length === 1 && !cols.length) return name + sp[0];
    return `${name}[${[...sp, colPart].filter(Boolean).join(local ? ';' : ',')}]`;
  };
  // Como Excel, dentro de su propia tabla la referencia se muestra sin el nombre: [@Ventas].
  O.tableRefToLocal = (text, own) => {
    const p = O.splitTableRef(text);
    const table = p.table && own && p.table.toLowerCase() === own.toLowerCase() ? null : p.table;
    return joinTableRef(table, p.areas, p.cols, true);
  };
  O.tableRefToFile = (text, own) => {
    const p = O.splitTableRef(text);
    return joinTableRef(p.table || own || null, p.areas, p.cols, false);
  };
  const tableBox = (t) => {
    const [a, b] = t.ref.split(':').map(O.parseAddr);
    return { r1: a.r, c1: a.c, r2: (b || a).r, c2: (b || a).c };
  };
  O.tableBox = tableBox;
  O.tableAt = (sheet, r, c) => (sheet.tables || []).find((t) => {
    const b = tableBox(t);
    return r >= b.r1 && r <= b.r2 && c >= b.c1 && c <= b.c2;
  }) || null;
  O.findTable = (state, name) => {
    for (const s of state.sheets) for (const t of s.tables || []) if (t.name.toLowerCase() === String(name).toLowerCase()) return { sheet: s, table: t };
    return null;
  };
  // Filas de datos de una tabla, sin encabezado ni totales.
  O.tableData = (t) => {
    const b = tableBox(t);
    return { ...b, r1: b.r1 + (t.header === false ? 0 : 1), r2: b.r2 - (t.totals || 0) };
  };
  // Rango al que apunta una referencia estructurada, como un nodo de referencia normal.
  O.resolveTableRef = (state, sheet, r, c, text) => {
    const p = O.splitTableRef(text);
    const hit = p.table ? O.findTable(state, p.table) : sheet && r !== undefined ? { sheet, table: O.tableAt(sheet, r, c) } : null;
    if (!hit || !hit.table) fail('#REF!');
    const t = hit.table, b = tableBox(t), d = O.tableData(t);
    let r1 = Infinity, r2 = -Infinity;
    const take = (lo, hi) => { r1 = Math.min(r1, lo); r2 = Math.max(r2, hi); };
    const areas = p.areas.length ? p.areas : ['Data'];
    for (const a of areas) {
      if (a === 'All') take(b.r1, b.r2);
      else if (a === 'Data') take(d.r1, d.r2);
      else if (a === 'Headers') { if (t.header === false) fail('#REF!'); take(b.r1, b.r1); } else if (a === 'Totals') { if (!t.totals) fail('#REF!'); take(b.r2, b.r2); } else {
        if (hit.sheet !== sheet || r === undefined || r < d.r1 || r > d.r2) fail('#VALUE!');
        take(r, r);
      }
    }
    let c1 = b.c1, c2 = b.c2;
    if (p.cols.length) {
      const idx = p.cols.map((name) => (t.columns || []).findIndex((x) => x.name.toLowerCase() === name.toLowerCase()));
      if (idx.some((i) => i < 0)) fail('#REF!');
      c1 = b.c1 + Math.min(...idx);
      c2 = b.c1 + Math.max(...idx);
    }
    const single = r1 === r2 && c1 === c2 && areas.includes('This Row');
    return { k: 'ref', sheet: hit.sheet === sheet ? null : hit.sheet.name, a: { r: r1, c: c1 }, b: single ? null : { r: r2, c: c2 } };
  };
  // Cambia el nombre de una columna en las referencias a su tabla.
  O.renameTableColumn = (src, table, oldName, newName, own) => src.replace(new RegExp(`"(?:[^"]|"")*"|(${SREF})`, 'g'), (m, ref) => {
    if (!ref) return m;
    const p = O.splitTableRef(ref);
    if ((p.table || own || '').toLowerCase() !== table.toLowerCase() || !p.cols.some((x) => x.toLowerCase() === oldName.toLowerCase())) return m;
    return joinTableRef(p.table, p.areas, p.cols.map((x) => (x.toLowerCase() === oldName.toLowerCase() ? newName : x)), false);
  });

  // ---------- Analisis lexico ----------
  const TOKEN = new RegExp(`\\s*(?:(\\d+\\.?\\d*(?:[eE][+-]?\\d+)?|\\.\\d+)|"((?:[^"]|"")*)"|(${SREF})|((?:'[^']+'|[A-Za-z_À-ɏ][\\w.À-ɏ]*)!)?(\\$?[A-Za-z]{1,3}\\$?\\d+)(?::(\\$?[A-Za-z]{1,3}\\$?\\d+))?(?![\\w(])|([A-Za-z_À-ɏ][\\w.À-ɏ]*)|(<>|<=|>=|[-+*/^&=<>%(),;]))`, 'y');

  function tokenize(src) {
    const out = [];
    TOKEN.lastIndex = 0;
    while (TOKEN.lastIndex < src.length) {
      if (/^\s*$/.test(src.slice(TOKEN.lastIndex))) break;
      if (/^\s*#REF!/i.test(src.slice(TOKEN.lastIndex))) fail('#REF!');
      const m = TOKEN.exec(src);
      if (!m) fail('#ERROR!');
      if (m[1] !== undefined) out.push({ t: 'num', v: parseFloat(m[1]) });
      else if (m[2] !== undefined) out.push({ t: 'str', v: m[2].replace(/""/g, '"') });
      else if (m[3] !== undefined) out.push({ t: 'sref', text: m[3] });
      else if (m[5] !== undefined) {
        const sheet = m[4] ? m[4].slice(0, -1).replace(/^'|'$/g, '') : null;
        out.push({ t: 'ref', sheet, a: O.parseAddr(m[5]), b: m[6] ? O.parseAddr(m[6]) : null });
      } else if (m[7] !== undefined) out.push({ t: 'name', v: m[7].toUpperCase() });
      else out.push({ t: 'op', v: m[8] === ';' ? ',' : m[8] });
    }
    return out;
  }

  // ---------- Analisis sintactico ----------
  function parse(src) {
    const toks = tokenize(src);
    let i = 0;
    const peek = (v) => toks[i] && toks[i].t === 'op' && toks[i].v === v;
    const eat = (v) => (peek(v) ? (i++, true) : false);
    const binary = (next, ops) => () => {
      let left = next();
      for (;;) {
        const op = toks[i] && toks[i].t === 'op' && ops.includes(toks[i].v) ? toks[i].v : null;
        if (!op) return left;
        i++;
        left = { k: 'bin', op, l: left, r: next() };
      }
    };
    const primary = () => {
      const t = toks[i++];
      if (!t) fail('#ERROR!');
      if (t.t === 'num' || t.t === 'str') return { k: 'lit', v: t.v };
      if (t.t === 'ref') return { k: 'ref', ...t };
      if (t.t === 'sref') return { k: 'sref', text: t.text };
      if (t.t === 'name') {
        if (eat('(')) {
          const args = [];
          if (!eat(')')) {
            do {
              args.push(peek(',') || peek(')') ? { k: 'lit', v: null } : expr());
            } while (eat(','));
            if (!eat(')')) fail('#ERROR!');
          }
          return { k: 'call', name: t.v, args };
        }
        if (t.v === 'TRUE' || t.v === 'VERDADERO') return { k: 'lit', v: true };
        if (t.v === 'FALSE' || t.v === 'FALSO') return { k: 'lit', v: false };
        return { k: 'named', name: t.v };
      }
      if (t.t === 'op' && t.v === '(') {
        const e = expr();
        if (!eat(')')) fail('#ERROR!');
        return e;
      }
      fail('#ERROR!');
    };
    const postfix = () => {
      let e = primary();
      while (eat('%')) e = { k: 'pct', e };
      return e;
    };
    const unary = () => {
      if (eat('-')) return { k: 'neg', e: unary() };
      if (eat('+')) return unary();
      return postfix();
    };
    const pow = binary(unary, ['^']);
    const mul = binary(pow, ['*', '/']);
    const add = binary(mul, ['+', '-']);
    const concat = binary(add, ['&']);
    const expr = binary(concat, ['=', '<>', '<', '>', '<=', '>=']);
    const tree = expr();
    if (i < toks.length) fail('#ERROR!');
    return tree;
  }

  // ---------- Conversiones con las reglas habituales de una hoja ----------
  const num = (v) => {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (Array.isArray(v)) return num(v[0] && v[0][0]);
    // Como Excel, opera con texto que parece número, porcentaje, hora o fecha.
    const n = O.textToNumber(v);
    if (n === null) fail('#VALUE!');
    return n;
  };
  const str = (v) => {
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'VERDADERO' : 'FALSO';
    if (Array.isArray(v)) return str(v[0] && v[0][0]);
    // Igual que un Excel en español: los números pasan a texto con coma decimal.
    if (typeof v === 'number') return String(+v.toPrecision(15)).replace('.', ',');
    return String(v);
  };
  const bool = (v) => (typeof v === 'string' ? fail('#VALUE!') : !!num(v));
  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  const compare = (a, b) => {
    if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    if (typeof a === 'string') {
      const x = a.toLowerCase(), y = b.toLowerCase();
      return x < y ? -1 : x > y ? 1 : 0;
    }
    return a < b ? -1 : a > b ? 1 : 0;
  };
  const flat = (args) => args.flatMap((a) => (Array.isArray(a) ? a.flat() : [a]));
  // Solo cuenta numeros dentro de rangos; los valores directos se convierten.
  const numbers = (args) =>
    args.flatMap((a) => (Array.isArray(a) ? a.flat().filter((v) => typeof v === 'number') : a === null ? [] : [num(a)]));

  // Comodines de Excel: * y ?, con ~ como escape. Devuelve null si el texto no tiene ninguno.
  const wildcard = (text) => {
    if (!/[*?]/.test(text)) return null;
    let re = '';
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (ch === '~' && /[*?~]/.test(text[i + 1] || '')) re += '\\' + text[++i];
      else if (ch === '*') re += '[\\s\\S]*';
      else if (ch === '?') re += '[\\s\\S]';
      else re += ch.replace(/[.+^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp('^' + re + '$', 'i');
  };
  const criterion = (raw) => {
    if (typeof raw !== 'string') return (v) => compare(v, raw) === 0 && v !== null && v !== undefined;
    const m = /^(<>|<=|>=|=|<|>)?(.*)$/.exec(raw);
    const op = m[1] || '=';
    const target = m[2] !== '' && !Number.isNaN(Number(m[2])) ? Number(m[2]) : m[2];
    const pattern = (op === '=' || op === '<>') && typeof target === 'string' ? wildcard(target) : null;
    return (v) => {
      if (pattern) {
        const hit = typeof v === 'string' && pattern.test(v);
        return op === '=' ? hit : !hit;
      }
      if (typeof target === 'number' && typeof v !== 'number') return op === '<>';
      if (typeof target === 'string' && op !== '=' && op !== '<>' && typeof v !== 'string') return false;
      const c = compare(v === undefined ? null : v, target);
      return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op];
    };
  };
  const round = (x, d, fn) => {
    const f = 10 ** d, y = x * f, sign = y < 0 ? -1 : 1;
    return (sign * fn(+(Math.abs(y)).toPrecision(15))) / f;
  };
  const EPOCH = Date.UTC(1899, 11, 30);
  O.dateToSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - EPOCH) / 86400000;
  O.serialToDate = (n) => new Date(EPOCH + Math.floor(n) * 86400000);
  const todaySerial = () => {
    const d = new Date();
    return O.dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate());
  };

  const FUNCS = {
    SUM: (a) => numbers(a).reduce((x, y) => x + y, 0),
    AVERAGE: (a) => { const n = numbers(a); return n.length ? n.reduce((x, y) => x + y, 0) / n.length : fail('#DIV/0!'); },
    MIN: (a) => { const n = numbers(a); return n.length ? Math.min(...n) : 0; },
    MAX: (a) => { const n = numbers(a); return n.length ? Math.max(...n) : 0; },
    // En un rango solo cuentan numeros; como argumento directo tambien lo que se pueda convertir.
    COUNT: (a) => a.reduce((n, x) => n + (Array.isArray(x)
      ? x.flat().filter((v) => typeof v === 'number').length
      : typeof x === 'number' || typeof x === 'boolean' || (typeof x === 'string' && x.trim() !== '' && !Number.isNaN(Number(x))) ? 1 : 0), 0),
    COUNTA: (a) => flat(a).filter((v) => v !== null && v !== undefined).length,
    AND: (a) => flat(a).filter((v) => v !== null).every(bool),
    OR: (a) => flat(a).filter((v) => v !== null).some(bool),
    NOT: (a) => !bool(a[0]),
    ROUND: (a) => round(num(a[0]), num(a[1]), Math.round),
    ROUNDUP: (a) => round(num(a[0]), num(a[1]), Math.ceil),
    ROUNDDOWN: (a) => round(num(a[0]), num(a[1]), Math.floor),
    ABS: (a) => Math.abs(num(a[0])),
    INT: (a) => Math.floor(num(a[0])),
    MOD: (a) => { const d = num(a[1]); if (!d) fail('#DIV/0!'); const n = num(a[0]); return n - d * Math.floor(n / d); },
    POWER: (a) => num(a[0]) ** num(a[1]),
    SQRT: (a) => { const n = num(a[0]); return n < 0 ? fail('#NUM!') : Math.sqrt(n); },
    PI: () => Math.PI,
    CONCAT: (a) => flat(a).map(str).join(''),
    CONCATENATE: (a) => flat(a).map(str).join(''),
    LEN: (a) => str(a[0]).length,
    UPPER: (a) => str(a[0]).toUpperCase(),
    LOWER: (a) => str(a[0]).toLowerCase(),
    TRIM: (a) => str(a[0]).trim().replace(/ +/g, ' '),
    LEFT: (a) => str(a[0]).slice(0, a[1] === undefined ? 1 : num(a[1])),
    RIGHT: (a) => { const s = str(a[0]), n = a[1] === undefined ? 1 : num(a[1]); return n ? s.slice(-n) : ''; },
    MID: (a) => str(a[0]).substr(num(a[1]) - 1, num(a[2])),
    TODAY: todaySerial,
    NOW: () => { const d = new Date(); return todaySerial() + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; },
    DATE: (a) => O.dateToSerial(num(a[0]), num(a[1]), num(a[2])),
    YEAR: (a) => O.serialToDate(num(a[0])).getUTCFullYear(),
    MONTH: (a) => O.serialToDate(num(a[0])).getUTCMonth() + 1,
    DAY: (a) => O.serialToDate(num(a[0])).getUTCDate(),
    SUMIF: (a) => {
      const test = criterion(a[1]), keys = flat([a[0]]), vals = flat([a[2] === undefined ? a[0] : a[2]]);
      return keys.reduce((s, k, i) => (test(k) && typeof vals[i] === 'number' ? s + vals[i] : s), 0);
    },
    COUNTIF: (a) => { const test = criterion(a[1]); return flat([a[0]]).filter(test).length; },
    AVERAGEIF: (a) => {
      const test = criterion(a[1]), keys = flat([a[0]]), vals = flat([a[2] === undefined ? a[0] : a[2]]);
      const hit = keys.map((k, i) => (test(k) ? vals[i] : null)).filter((v) => typeof v === 'number');
      return hit.length ? hit.reduce((x, y) => x + y, 0) / hit.length : fail('#DIV/0!');
    },
    VLOOKUP: (a) => {
      const table = Array.isArray(a[1]) ? a[1] : fail('#REF!'), col = num(a[2]);
      if (col < 1 || col > (table[0] || []).length) fail('#REF!');
      const exact = a[3] !== undefined && !bool(a[3] === null ? 0 : a[3]);
      let found = null;
      for (const row of table) {
        const c = compare(row[0], a[0]);
        if (c === 0) { found = row; break; }
        if (!exact && c < 0 && typeof row[0] === typeof a[0]) found = row;
      }
      return found ? found[col - 1] : fail('#N/A');
    },
    MATCH: (a) => {
      const list = flat([a[1]]), mode = a[2] === undefined ? 1 : num(a[2]);
      if (mode === 0) {
        const i = list.findIndex((v) => compare(v, a[0]) === 0 && v !== null);
        return i < 0 ? fail('#N/A') : i + 1;
      }
      let best = -1;
      list.forEach((v, i) => { if (v !== null && (mode > 0 ? compare(v, a[0]) <= 0 : compare(v, a[0]) >= 0)) best = i; });
      return best < 0 ? fail('#N/A') : best + 1;
    },
    INDEX: (a) => {
      const t = Array.isArray(a[0]) ? a[0] : [[a[0]]];
      let r = num(a[1]), c = a[2] === undefined ? 0 : num(a[2]);
      if (t.length === 1 && a[2] === undefined) { c = r; r = 1; }
      if (r < 1 || r > t.length) fail('#REF!');
      c = c || 1;
      if (c < 1 || c > t[0].length) fail('#REF!');
      return t[r - 1][c - 1];
    },
    SUMIFS: (a) => { const vals = flat([a[0]]); return multiHits(a.slice(1), vals.length).reduce((s, i) => (typeof vals[i] === 'number' ? s + vals[i] : s), 0); },
    COUNTIFS: (a) => multiHits(a, flat([a[0]]).length).length,
    AVERAGEIFS: (a) => {
      const vals = flat([a[0]]), hit = multiHits(a.slice(1), vals.length).map((i) => vals[i]).filter((v) => typeof v === 'number');
      return hit.length ? hit.reduce((x, y) => x + y, 0) / hit.length : fail('#DIV/0!');
    },
    MAXIFS: (a) => { const vals = flat([a[0]]), hit = multiHits(a.slice(1), vals.length).map((i) => vals[i]).filter((v) => typeof v === 'number'); return hit.length ? Math.max(...hit) : 0; },
    MINIFS: (a) => { const vals = flat([a[0]]), hit = multiHits(a.slice(1), vals.length).map((i) => vals[i]).filter((v) => typeof v === 'number'); return hit.length ? Math.min(...hit) : 0; },
    COUNTBLANK: (a) => flat(a).filter((v) => v === null || v === undefined || v === '').length,
    ISBLANK: (a) => a[0] === null || a[0] === undefined,
    ISNUMBER: (a) => typeof a[0] === 'number',
    ISTEXT: (a) => typeof a[0] === 'string',
    SEARCH: (a) => {
      const find = str(a[0]), text = str(a[1]), start = a[2] === undefined ? 1 : num(a[2]);
      if (start < 1 || start > text.length + 1) fail('#VALUE!');
      const re = wildcard(find);
      if (re) {
        const inner = new RegExp(re.source.slice(1, -1), 'i');
        const m = inner.exec(text.slice(start - 1));
        return m ? m.index + start : fail('#VALUE!');
      }
      const i = text.toLowerCase().indexOf(find.toLowerCase(), start - 1);
      return i < 0 ? fail('#VALUE!') : i + 1;
    },
    FIND: (a) => {
      const text = str(a[1]), start = a[2] === undefined ? 1 : num(a[2]);
      if (start < 1 || start > text.length + 1) fail('#VALUE!');
      const i = text.indexOf(str(a[0]), start - 1);
      return i < 0 ? fail('#VALUE!') : i + 1;
    },
    SUBSTITUTE: (a) => {
      const text = str(a[0]), old = str(a[1]), rep = str(a[2]);
      if (!old) return text;
      if (a[3] === undefined) return text.split(old).join(rep);
      const n = num(a[3]);
      let i = -1;
      for (let k = 0; k < n; k++) { i = text.indexOf(old, i + 1); if (i < 0) return text; }
      return text.slice(0, i) + rep + text.slice(i + old.length);
    },
    REPLACE: (a) => { const t = str(a[0]), s = num(a[1]) - 1; return t.slice(0, s) + str(a[3]) + t.slice(s + num(a[2])); },
    // En los archivos de Excel en español el formato de TEXTO va con letras españolas: aaaa es el año.
    TEXT: (a) => {
      const f = str(a[1]).replace(/"[^"]*"|a+/gi, (m) => (m[0] === '"' ? m : 'y'.repeat(m.length)));
      return typeof a[0] === 'number' || O.textToNumber(a[0]) !== null ? O.formatValue(num(a[0]), f) : str(a[0]);
    },
    VALUE: (a) => (typeof a[0] === 'number' ? a[0] : num(str(a[0]))),
    PRODUCT: (a) => numbers(a).reduce((x, y) => x * y, 1),
    SUMPRODUCT: (a) => {
      const lists = a.map((x) => flat([x]));
      if (lists.some((l) => l.length !== lists[0].length)) fail('#VALUE!');
      return lists[0].reduce((s, _, i) => s + lists.reduce((p, l) => p * (typeof l[i] === 'number' ? l[i] : 0), 1), 0);
    },
    MEDIAN: (a) => {
      const n = numbers(a).sort((x, y) => x - y);
      if (!n.length) fail('#NUM!');
      const h = n.length >> 1;
      return n.length % 2 ? n[h] : (n[h - 1] + n[h]) / 2;
    },
    LARGE: (a) => { const n = numbers([a[0]]).sort((x, y) => y - x), k = num(a[1]); return k >= 1 && k <= n.length ? n[Math.ceil(k) - 1] : fail('#NUM!'); },
    SMALL: (a) => { const n = numbers([a[0]]).sort((x, y) => x - y), k = num(a[1]); return k >= 1 && k <= n.length ? n[Math.ceil(k) - 1] : fail('#NUM!'); },
    RANK: (a) => {
      const v = num(a[0]), n = numbers([a[1]]), asc = a[2] !== undefined && a[2] !== null && num(a[2]) !== 0;
      if (!n.includes(v)) fail('#N/A');
      return n.filter((x) => (asc ? x < v : x > v)).length + 1;
    },
    WEEKDAY: (a) => {
      const d = O.serialToDate(num(a[0])).getUTCDay(), t = a[1] === undefined ? 1 : num(a[1]);
      if (t === 1) return d + 1;
      if (t === 2) return ((d + 6) % 7) + 1;
      if (t === 3) return (d + 6) % 7;
      return fail('#NUM!');
    },
    HOUR: (a) => Math.floor((((num(a[0]) % 1) * 86400 + 0.5) / 3600)) % 24,
    MINUTE: (a) => Math.floor((((num(a[0]) % 1) * 86400 + 0.5) % 3600) / 60),
    SECOND: (a) => Math.floor(((num(a[0]) % 1) * 86400 + 0.5) % 60),
    TIME: (a) => {
      const t = (num(a[0]) * 3600 + num(a[1]) * 60 + num(a[2])) / 86400;
      return t < 0 ? fail('#NUM!') : t % 1;
    },
    EDATE: (a) => {
      const d = O.serialToDate(num(a[0])), m = Math.trunc(num(a[1]));
      const y = d.getUTCFullYear(), mo = d.getUTCMonth() + m;
      const last = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate();
      return O.dateToSerial(y, mo + 1, Math.min(d.getUTCDate(), last));
    },
    EOMONTH: (a) => {
      const d = O.serialToDate(num(a[0])), m = Math.trunc(num(a[1]));
      return (Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + m + 1, 0) - EPOCH) / 86400000;
    },
    DATEDIF: (a) => {
      const s = Math.floor(num(a[0])), e = Math.floor(num(a[1])), unit = str(a[2]).toUpperCase();
      if (s > e) fail('#NUM!');
      const ds = O.serialToDate(s), de = O.serialToDate(e);
      let months = (de.getUTCFullYear() - ds.getUTCFullYear()) * 12 + de.getUTCMonth() - ds.getUTCMonth();
      if (de.getUTCDate() < ds.getUTCDate()) months--;
      if (unit === 'D') return e - s;
      if (unit === 'M') return months;
      if (unit === 'Y') return Math.floor(months / 12);
      if (unit === 'YM') return months % 12;
      if (unit === 'MD') {
        const prev = de.getUTCDate() >= ds.getUTCDate() ? 0 : new Date(Date.UTC(de.getUTCFullYear(), de.getUTCMonth(), 0)).getUTCDate();
        return de.getUTCDate() - ds.getUTCDate() + prev;
      }
      return fail('#NUM!');
    },
    NETWORKDAYS: (a) => {
      let s = Math.floor(num(a[0])), e = Math.floor(num(a[1])), sign = 1;
      if (s > e) { [s, e] = [e, s]; sign = -1; }
      const off = new Set(a[2] === undefined ? [] : numbers([a[2]]).map(Math.floor));
      let n = 0;
      for (let d = s; d <= e; d++) {
        const w = O.serialToDate(d).getUTCDay();
        if (w !== 0 && w !== 6 && !off.has(d)) n++;
      }
      return sign * n;
    },
    HLOOKUP: (a) => {
      const table = Array.isArray(a[1]) ? a[1] : fail('#REF!'), row = num(a[2]);
      if (row < 1 || row > table.length) fail('#REF!');
      const exact = a[3] !== undefined && !bool(a[3] === null ? 0 : a[3]);
      let found = -1;
      for (let c = 0; c < table[0].length; c++) {
        const cmp = compare(table[0][c], a[0]);
        if (cmp === 0) { found = c; break; }
        if (!exact && cmp < 0 && typeof table[0][c] === typeof a[0]) found = c;
      }
      return found < 0 ? fail('#N/A') : table[row - 1][found];
    },
    XLOOKUP: (a) => {
      const keys = flat([a[1]]), vals = Array.isArray(a[2]) ? a[2] : [[a[2]]];
      const i = keys.findIndex((k) => k !== null && compare(k, a[0]) === 0);
      if (i < 0) return a[3] !== undefined ? a[3] : fail('#N/A');
      // Devuelve la fila o columna correspondiente; aqui solo el primer valor.
      return vals.length === 1 ? vals[0][i] : (vals[i] || [])[0];
    },
    CHOOSE: (a) => { const i = Math.floor(num(a[0])); return i >= 1 && i < a.length ? a[i] : fail('#VALUE!'); },
    TEXTJOIN: (a) => {
      const skip = bool(a[1]);
      return flat(a.slice(2)).filter((v) => !(skip && (v === null || v === ''))).map(str).join(str(a[0]));
    },
    PROPER: (a) => str(a[0]).toLowerCase().replace(/(^|[^a-zà-ÿ])([a-zà-ÿ])/gi, (m, p, ch) => p + ch.toUpperCase()),
    REPT: (a) => str(a[0]).repeat(Math.max(0, Math.floor(num(a[1])))),
    EXACT: (a) => str(a[0]) === str(a[1]),
    TRUNC: (a) => { const d = a[1] === undefined ? 0 : num(a[1]), f = 10 ** d; return Math.trunc(num(a[0]) * f) / f; },
    SIGN: (a) => Math.sign(num(a[0])),
    NA: () => fail('#N/A'),
    RAND: () => Math.random(),
    RANDBETWEEN: (a) => { const lo = Math.ceil(num(a[0])), hi = Math.floor(num(a[1])); return lo > hi ? fail('#NUM!') : lo + Math.floor(Math.random() * (hi - lo + 1)); },
  };
  FUNCS['RANK.EQ'] = FUNCS.RANK;
  // Indices que cumplen todos los pares rango, criterio de los *.SI.CONJUNTO.
  function multiHits(pairs, size) {
    if (pairs.length % 2) fail('#VALUE!');
    let hits = [...Array(size).keys()];
    for (let p = 0; p < pairs.length; p += 2) {
      const list = flat([pairs[p]]), test = criterion(pairs[p + 1]);
      if (list.length !== size) fail('#VALUE!');
      hits = hits.filter((i) => test(list[i]));
    }
    return hits;
  }
  // Funciones que reciben los errores de sus argumentos en vez de propagarlos.
  const ERROR_AWARE = new Set(['IFERROR', 'IFNA', 'ISERROR', 'ISERR', 'ISNA', 'IFS']);
  O.formulaNames = Object.keys(FUNCS).concat([...ERROR_AWARE], 'IF', 'SUBTOTAL').sort();

  // Nombres de Excel en español. En el archivo siempre se guardan en inglés.
  const ES = {
    SUM: 'SUMA', AVERAGE: 'PROMEDIO', COUNT: 'CONTAR', COUNTA: 'CONTARA', AND: 'Y', OR: 'O', NOT: 'NO',
    ROUND: 'REDONDEAR', ROUNDUP: 'REDONDEAR.MAS', ROUNDDOWN: 'REDONDEAR.MENOS', INT: 'ENTERO', MOD: 'RESIDUO',
    POWER: 'POTENCIA', SQRT: 'RAIZ', CONCATENATE: 'CONCATENAR', LEN: 'LARGO', UPPER: 'MAYUSC', LOWER: 'MINUSC',
    TRIM: 'ESPACIOS', LEFT: 'IZQUIERDA', RIGHT: 'DERECHA', MID: 'EXTRAE', TODAY: 'HOY', NOW: 'AHORA', DATE: 'FECHA',
    YEAR: 'AÑO', MONTH: 'MES', DAY: 'DIA', SUMIF: 'SUMAR.SI', COUNTIF: 'CONTAR.SI', AVERAGEIF: 'PROMEDIO.SI',
    VLOOKUP: 'BUSCARV', MATCH: 'COINCIDIR', INDEX: 'INDICE', IF: 'SI', IFERROR: 'SI.ERROR',
    SUMIFS: 'SUMAR.SI.CONJUNTO', COUNTIFS: 'CONTAR.SI.CONJUNTO', AVERAGEIFS: 'PROMEDIO.SI.CONJUNTO',
    MAXIFS: 'MAX.SI.CONJUNTO', MINIFS: 'MIN.SI.CONJUNTO', COUNTBLANK: 'CONTAR.BLANCO', ISBLANK: 'ESBLANCO',
    ISNUMBER: 'ESNUMERO', ISTEXT: 'ESTEXTO', ISERROR: 'ESERROR', ISERR: 'ESERR', ISNA: 'ESNOD', IFNA: 'SI.ND',
    SEARCH: 'HALLAR', FIND: 'ENCONTRAR', SUBSTITUTE: 'SUSTITUIR', REPLACE: 'REEMPLAZAR', TEXT: 'TEXTO', VALUE: 'VALOR',
    PRODUCT: 'PRODUCTO', SUMPRODUCT: 'SUMAPRODUCTO', MEDIAN: 'MEDIANA', LARGE: 'K.ESIMO.MAYOR', SMALL: 'K.ESIMO.MENOR',
    RANK: 'JERARQUIA', 'RANK.EQ': 'JERARQUIA.EQV', WEEKDAY: 'DIASEM', HOUR: 'HORA', MINUTE: 'MINUTO', SECOND: 'SEGUNDO',
    TIME: 'NSHORA', EDATE: 'FECHA.MES', EOMONTH: 'FIN.MES', DATEDIF: 'SIFECHA', NETWORKDAYS: 'DIAS.LAB',
    HLOOKUP: 'BUSCARH', XLOOKUP: 'BUSCARX', CHOOSE: 'ELEGIR', IFS: 'SI.CONJUNTO', TEXTJOIN: 'UNIRCADENAS',
    PROPER: 'NOMPROPIO', REPT: 'REPETIR', EXACT: 'IGUAL', TRUNC: 'TRUNCAR', SIGN: 'SIGNO', RAND: 'ALEATORIO',
    RANDBETWEEN: 'ALEATORIO.ENTRE', NA: 'NOD', SUBTOTAL: 'SUBTOTALES',
  };
  const EN = {};
  for (const [en, es] of Object.entries(ES)) EN[es] = en;
  EN['DÍA'] = 'DAY';
  O.localName = (en) => ES[en] || en;
  O.englishName = (name) => {
    const up = name.toUpperCase();
    return EN[up] || up;
  };
  const NAME_CALL = /"(?:[^"]|"")*"|'[^']+'!|([A-Za-z_À-ɏ][\w.À-ɏ]*)(?=\s*\()/g;
  // De la forma guardada (inglés, comas) a la que se ve y se escribe en un Excel en español.
  const LOCAL = new RegExp(`"(?:[^"]|"")*"|'[^']+'!|(${SREF})|([A-Za-z_À-ɏ][\\w.À-ɏ]*)(\\s*\\()?|(\\d+\\.\\d*(?:[eE][+-]?\\d+)?|\\.\\d+)|,`, 'g');
  const SREF_ONLY = new RegExp(`"(?:[^"]|"")*"|(${SREF})`, 'g');
  // own es la tabla en la que está la fórmula, si está dentro de una.
  O.toLocalFormula = (src, own) =>
    src.replace(LOCAL, (m, sref, name, call, number) => {
      if (sref !== undefined) return O.tableRefToLocal(sref, own);
      if (name !== undefined) {
        if (call) return (ES[name.toUpperCase()] || name.toUpperCase()) + call;
        if (/^(TRUE|FALSE)$/i.test(name)) return /^t/i.test(name) ? 'VERDADERO' : 'FALSO';
        return m;
      }
      if (number !== undefined) return number.replace('.', ',');
      return m === ',' ? ';' : m;
    });
  // De lo que escribe el usuario a la forma guardada. Con ; o nombres en español la coma es decimal,
  // igual que en un Excel en español; si no, la coma separa argumentos, como en uno en inglés.
  O.fromLocalFormula = (src, own) => {
    // Las referencias de tabla pasan a la forma del archivo, con el nombre de la tabla.
    src = src.replace(SREF_ONLY, (m, ref) => (ref ? O.tableRefToFile(ref, own) : m));
    const outside = src.replace(/"(?:[^"]|"")*"/g, '""').replace(new RegExp(BRACKETS, 'g'), '[]');
    let spanish = outside.includes(';') || !outside.includes('(');
    if (!spanish) {
      for (const m of outside.matchAll(NAME_CALL)) {
        const up = m[1] && m[1].toUpperCase();
        if (up && EN[up] && !FUNCS[up]) spanish = true;
      }
    }
    let out = src;
    if (spanish) {
      out = out.replace(new RegExp(`"(?:[^"]|"")*"|'[^']+'!|${BRACKETS}|[A-Za-z_À-ɏ$][\\w.À-ɏ$]*|(\\d+),(\\d+)|;`, 'g'), (m, a, b) =>
        (a !== undefined ? `${a}.${b}` : m === ';' ? ',' : m));
    }
    return out.replace(NAME_CALL, (m, name) => (name === undefined ? m : O.englishName(name)));
  };

  // SUBTOTALES: como Excel, no cuenta las filas que oculta un filtro ni otros SUBTOTALES;
  // con 101 a 111 tampoco las ocultas a mano.
  const SUBTOTAL_FN = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' };
  const variance = (a, sample) => {
    const n = numbers(a);
    if (n.length < (sample ? 2 : 1)) fail('#DIV/0!');
    const mean = n.reduce((x, y) => x + y, 0) / n.length;
    return n.reduce((x, y) => x + (y - mean) ** 2, 0) / (n.length - (sample ? 1 : 0));
  };
  const STATS = { STDEV: (a) => Math.sqrt(variance(a, true)), STDEVP: (a) => Math.sqrt(variance(a, false)), VAR: (a) => variance(a, true), VARP: (a) => variance(a, false) };
  function subtotal(n, ctx) {
    const code = num(evalNode(n.args[0], ctx));
    const fn = SUBTOTAL_FN[code % 100];
    if (!fn || code > 111 || (code > 11 && code < 101)) fail('#VALUE!');
    const ranges = n.args.slice(1).map((a) => {
      const node = a.k === 'sref' ? ctx.table(a.text) : a;
      if (node.k !== 'ref') fail('#VALUE!');
      const b = node.b || node.a;
      const rows = [];
      for (let r = Math.min(node.a.r, b.r); r <= Math.max(node.a.r, b.r); r++) {
        const state = ctx.rowState(node.sheet, r);
        if (state.filtered || (code > 100 && state.hidden)) continue;
        const row = [];
        for (let c = Math.min(node.a.c, b.c); c <= Math.max(node.a.c, b.c); c++) row.push(ctx.isSubtotal(node.sheet, r, c) ? null : ctx.cell(node.sheet, r, c));
        rows.push(row);
      }
      return rows;
    });
    return (FUNCS[fn] || STATS[fn])(ranges);
  }

  function evalNode(n, ctx) {
    switch (n.k) {
      case 'lit': return n.v;
      case 'sref': return evalNode(ctx.table(n.text), ctx);
      case 'named': return evalNode(ctx.named(n.name), ctx);
      case 'ref': {
        if (!n.b) return ctx.cell(n.sheet, n.a.r, n.a.c);
        const r1 = Math.min(n.a.r, n.b.r), r2 = Math.max(n.a.r, n.b.r), c1 = Math.min(n.a.c, n.b.c), c2 = Math.max(n.a.c, n.b.c);
        const rows = [];
        for (let r = r1; r <= r2; r++) {
          const row = [];
          for (let c = c1; c <= c2; c++) row.push(ctx.cell(n.sheet, r, c));
          rows.push(row);
        }
        return rows;
      }
      case 'neg': return -num(evalNode(n.e, ctx));
      case 'pct': return num(evalNode(n.e, ctx)) / 100;
      case 'bin': {
        const l = evalNode(n.l, ctx), r = evalNode(n.r, ctx);
        switch (n.op) {
          case '+': return num(l) + num(r);
          case '-': return num(l) - num(r);
          case '*': return num(l) * num(r);
          case '/': { const d = num(r); if (!d) fail('#DIV/0!'); return num(l) / d; }
          case '^': return num(l) ** num(r);
          case '&': return str(l) + str(r);
          case '=': return compare(l, r) === 0;
          case '<>': return compare(l, r) !== 0;
          case '<': return compare(l, r) < 0;
          case '>': return compare(l, r) > 0;
          case '<=': return compare(l, r) <= 0;
          default: return compare(l, r) >= 0;
        }
      }
      case 'call': {
        const name = FUNCS[n.name] || n.name === 'IF' || n.name === 'SUBTOTAL' || ERROR_AWARE.has(n.name) ? n.name : O.englishName(n.name);
        if (name === 'SUBTOTAL') return subtotal(n, ctx);
        if (name === 'IF') {
          if (bool(evalNode(n.args[0], ctx))) return n.args[1] ? evalNode(n.args[1], ctx) : true;
          return n.args[2] ? evalNode(n.args[2], ctx) : false;
        }
        if (name === 'IFS') {
          for (let k = 0; k + 1 < n.args.length; k += 2) if (bool(evalNode(n.args[k], ctx))) return evalNode(n.args[k + 1], ctx);
          return fail('#N/A');
        }
        if (ERROR_AWARE.has(name)) {
          // Para ESERR y SI.ND solo cuentan algunos errores; el resto se propaga o da FALSO.
          const counts = (code) => (name === 'ISNA' || name === 'IFNA' ? code === '#N/A' : name === 'ISERR' ? code !== '#N/A' : true);
          let v;
          try { v = evalNode(n.args[0], ctx); } catch (e) {
            if (!(e instanceof CellError)) throw e;
            if (!counts(e.code)) { if (name === 'IFNA') throw e; return false; }
            if (name.startsWith('IS')) return true;
            return n.args[1] ? evalNode(n.args[1], ctx) : '';
          }
          return name.startsWith('IS') ? false : v;
        }
        const fn = FUNCS[name];
        if (!fn) fail('#NAME?');
        return fn(n.args.map((a) => evalNode(a, ctx)));
      }
    }
    return fail('#ERROR!');
  }

  // Desplaza las referencias relativas de una formula al copiarla, respetando $ y cadenas.
  O.shiftFormula = (src, dr, dc) =>
    src.replace(new RegExp(`"(?:[^"]|"")*"|'[^']+'!|${BRACKETS}|(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+)(?![\\w(])`, 'g'), (m, ac, col, ar, row, offset, all) => {
      if (col === undefined) return m;
      if (offset > 0 && /[\w.]/.test(all[offset - 1])) return m;
      const c = ac ? O.colIndex(col) : O.colIndex(col) + dc;
      const r = ar ? +row - 1 : +row - 1 + dr;
      if (c < 0 || r < 0) return '#REF!';
      return ac + O.colName(c) + ar + (r + 1);
    });

  const REFS = new RegExp(`"(?:[^"]|"")*"|${BRACKETS}|((?:'[^']+'|[A-Za-z_À-ɏ][\\w.À-ɏ]*)!)?(\\$?[A-Za-z]{1,3}\\$?\\d+)(?::(\\$?[A-Za-z]{1,3}\\$?\\d+))?(?![\\w(])`, 'g');
  const prefixName = (prefix) => prefix.slice(0, -1).replace(/^'|'$/g, '').toLowerCase();

  // Ajusta las referencias al insertar (n > 0) o eliminar (n < 0) filas o columnas en una hoja.
  // Las referencias sin hoja solo cuentan si la formula esta en la hoja modificada.
  // band limita el cambio a una franja: al desplazar celdas sueltas, las de fuera no se mueven.
  O.adjustRefs = (src, { axis, at, n, sheet, local, band }) =>
    src.replace(REFS, (m, prefix, first, second, offset, all) => {
      if (first === undefined) return m;
      if (offset > 0 && /[\w.À-ɏ]/.test(all[offset - 1])) return m;
      if (prefix ? prefixName(prefix) !== sheet.toLowerCase() : !local) return m;
      const parts = (a) => /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(a);
      const a = parts(first), b = parts(second || first);
      const get = (p) => (axis === 'r' ? +p[4] - 1 : O.colIndex(p[2]));
      if (band) {
        const across = (p) => (axis === 'r' ? O.colIndex(p[2]) : +p[4] - 1);
        if (Math.min(across(a), across(b)) < band[0] || Math.max(across(a), across(b)) > band[1]) return m;
      }
      let lo = Math.min(get(a), get(b)), hi = Math.max(get(a), get(b));
      const swap = get(a) > get(b);
      if (n > 0) {
        if (lo >= at) lo += n;
        if (hi >= at) hi += n;
      } else {
        const end = at - n;
        if (lo >= at && hi < end) return '#REF!';
        lo = lo >= end ? lo + n : Math.min(lo, at);
        hi = hi >= end ? hi + n : Math.min(hi, at - 1);
      }
      const put = (p, i) => (axis === 'r' ? p[1] + p[2] + p[3] + (i + 1) : p[1] + O.colName(i) + p[3] + p[4]);
      const out = second ? put(a, swap ? hi : lo) + ':' + put(b, swap ? lo : hi) : put(a, lo);
      return (prefix || '') + out;
    });

  // Cambia el nombre de una hoja en las referencias de una formula.
  O.renameSheetRefs = (src, oldName, newName) => {
    const prefix = /^[A-Za-z_À-ɏ][\w.À-ɏ]*$/.test(newName) ? newName + '!' : `'${newName}'!`;
    return src.replace(REFS, (m, p, first) => {
      if (first === undefined || !p || prefixName(p) !== oldName.toLowerCase()) return m;
      return prefix + m.slice(p.length);
    });
  };

  // Pasa ; a , fuera de cadenas para que el archivo guardado siga la sintaxis estandar.
  O.normalizeFormula = (src) => src.replace(/"(?:[^"]|"")*"|;/g, (m) => (m === ';' ? ',' : m));

  // ---------- Formato de numeros ----------
  const bareFormat = (nf) => nf.replace(/"[^"]*"|\\.|\[(?![hms]+\])[^\]]*\]/gi, '');
  O.isDateFormat = (nf) => !!nf && /[dy]|m{1,4}\/|\/m|[hs]|am\/pm/i.test(bareFormat(nf));
  const MONTHS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const DAYS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
  // Formatos de fecha y hora: d, dd, ddd, dddd, m..mmmm, yy, yyyy, h, hh, mm (minutos), ss, [h], AM/PM.
  const formatDate = (v, f) => {
    const toks = [];
    const re = /"([^"]*)"|\\(.)|\[(h+|m+|s+)\]|(am\/pm|a\/p)|(y+|m+|d+|h+|s+)|(\.0+)|\[[^\]]*\]|(.)/gi;
    let m;
    while ((m = re.exec(f))) {
      if (m[1] !== undefined || m[2] !== undefined) toks.push({ lit: m[1] !== undefined ? m[1] : m[2] });
      else if (m[3]) toks.push({ elapsed: m[3].toLowerCase() });
      else if (m[4]) toks.push({ ampm: m[4] });
      else if (m[5]) toks.push({ code: m[5].toLowerCase() });
      else if (m[6]) toks.push({ frac: m[6].length - 1 });
      else if (m[7] !== undefined) toks.push({ lit: m[7] });
    }
    // m es minutos si va detrás de una hora o delante de segundos.
    toks.forEach((t, i) => {
      if (!t.code || t.code[0] !== 'm' || t.code.length > 2) return;
      const prev = toks.slice(0, i).reverse().find((x) => x.code || x.elapsed);
      const next = toks.slice(i + 1).find((x) => x.code || x.elapsed);
      if ((prev && /^h/.test(prev.code || prev.elapsed)) || (next && /^s/.test(next.code || next.elapsed))) t.min = true;
    });
    const hasAmPm = toks.some((t) => t.ampm);
    const fracDigits = (toks.find((t) => t.frac) || {}).frac || 0;
    const totalSec = Math.round(v * 86400 * 10 ** fracDigits) / 10 ** fracDigits;
    const day = Math.floor(totalSec / 86400);
    const d = O.serialToDate(day);
    const secOfDay = totalSec - day * 86400;
    const H = Math.floor(secOfDay / 3600), Mi = Math.floor((secOfDay % 3600) / 60), S = Math.floor(secOfDay % 60);
    const pad = (n, w) => String(n).padStart(w, '0');
    return toks.map((t) => {
      if (t.lit !== undefined) return t.lit;
      if (t.ampm) return H < 12 ? (t.ampm.length > 3 ? 'a. m.' : 'a') : t.ampm.length > 3 ? 'p. m.' : 'p';
      if (t.frac) return ',' + pad(Math.round((secOfDay % 1) * 10 ** t.frac), t.frac);
      if (t.elapsed) {
        const n = t.elapsed[0] === 'h' ? Math.floor(totalSec / 3600) : t.elapsed[0] === 'm' ? Math.floor(totalSec / 60) : Math.floor(totalSec);
        return (totalSec < 0 ? '-' : '') + pad(Math.abs(n), t.elapsed.length);
      }
      const c = t.code, w = c.length;
      if (c[0] === 'y') return w <= 2 ? pad(d.getUTCFullYear() % 100, 2) : String(d.getUTCFullYear());
      if (c[0] === 'd') return w === 1 ? String(d.getUTCDate()) : w === 2 ? pad(d.getUTCDate(), 2) : w === 3 ? DAYS[d.getUTCDay()].slice(0, 3) : DAYS[d.getUTCDay()];
      if (c[0] === 'h') {
        const h = hasAmPm ? H % 12 || 12 : H;
        return w === 1 ? String(h) : pad(h, 2);
      }
      if (c[0] === 's') return w === 1 ? String(S) : pad(S, 2);
      if (t.min) return w === 1 ? String(Mi) : pad(Mi, 2);
      const mo = d.getUTCMonth();
      return w === 1 ? String(mo + 1) : w === 2 ? pad(mo + 1, 2) : w === 3 ? MONTHS[mo].slice(0, 3) : w === 5 ? MONTHS[mo][0] : MONTHS[mo];
    }).join('');
  };
  // Formato General de Excel: como mucho 11 caracteres, con signo y coma. Lo que no cabe, o los
  // números muy pequeños, pasan a notación científica (1,23457E+11, 1E-05).
  const generalText = (v) => {
    if (v === 0) return '0';
    const abs = Math.abs(v), sign = v < 0 ? 1 : 0;
    const intLen = String(Math.trunc(abs)).length;
    if (intLen + sign > 11 || abs < 1e-4) {
      const [m, e] = abs.toExponential(5).split('e');
      const exp = +e;
      return (sign ? '-' : '') + String(+m).replace('.', ',') + 'E' + (exp < 0 ? '-' : '+') + String(Math.abs(exp)).padStart(2, '0');
    }
    const dec = Math.max(0, 11 - sign - intLen - 1);
    return (+v.toFixed(dec)).toLocaleString('es-ES', { maximumFractionDigits: dec, useGrouping: false });
  };
  O.formatValue = (v, nf) => {
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'VERDADERO' : 'FALSO';
    if (typeof v !== 'number') return String(v);
    const f = nf && nf !== 'General' ? nf.split(';')[0] : null;
    if (!f) return generalText(v);
    // [$-F400] y [$-F800] son la hora y la fecha larga del sistema; se usan las de Windows en español.
    if (/\[\$-F400\]/i.test(f)) return O.formatValue(v, 'h:mm:ss');
    if (/\[\$-F800\]/i.test(f)) return O.formatValue(v, 'dddd, d "de" mmmm "de" yyyy');
    if (O.isDateFormat(f)) return v < 0 && !/\[[hms]+\]/i.test(f) ? '#'.repeat(8) : formatDate(v, f);
    const clean = f.replace(/\[[^\]]*\]/g, '');
    const quoted = /(\s*)"([^"]*)"/.exec(clean);
    const literal = quoted ? quoted[1] + quoted[2] : clean.includes('€') ? ' €' : clean.includes('$') ? ' $' : '';
    const bare = clean.replace(/"[^"]*"/g, '');
    const dec = (/\.(0+)/.exec(bare) || [, ''])[1].length;
    const pct = bare.includes('%');
    const text = (pct ? v * 100 : v).toLocaleString('es-ES', {
      minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: bare.includes(','),
    });
    return text + (pct ? ' %' : '') + literal;
  };

  // Reconoce un número escrito como texto: coma o punto decimal, %, hora h:mm[:ss] [AM/PM]
  // y fecha dd/mm/aaaa con hora opcional. Devuelve {v, nf} o null.
  const TIME = '(\\d{1,2}):(\\d{2})(?::(\\d{2}(?:[.,]\\d+)?))?(?:\\s*([ap])\\.?\\s?m\\.?)?';
  const timeOf = (m, i) => {
    let h = +m[i];
    const min = +m[i + 1], sec = m[i + 2] ? parseFloat(m[i + 2].replace(',', '.')) : 0, ap = m[i + 3];
    if (min > 59 || sec >= 60 || (ap && (h < 1 || h > 12))) return null;
    if (ap) h = (h % 12) + (/p/i.test(ap) ? 12 : 0);
    return (h * 3600 + min * 60 + sec) / 86400;
  };
  O.readNumber = (raw) => {
    const s = String(raw).trim();
    if (/^[+-]?\d+([.,]\d+)?$/.test(s) || /^[+-]?[.,]\d+$/.test(s)) return { v: parseFloat(s.replace(',', '.')) };
    if (/^[+-]?\d+([.,]\d+)?\s*%$/.test(s)) return { v: parseFloat(s.replace(',', '.')) / 100, nf: '0%' };
    let m = new RegExp(`^${TIME}$`, 'i').exec(s);
    if (m) {
      const t = timeOf(m, 1);
      return t === null ? null : { v: t, nf: m[3] ? 'hh:mm:ss' : 'hh:mm' };
    }
    m = new RegExp(`^(\\d{1,2})[/-](\\d{1,2})[/-](\\d{4})(?:\\s+${TIME})?$`, 'i').exec(s);
    if (m && +m[2] >= 1 && +m[2] <= 12 && +m[1] >= 1 && +m[1] <= 31) {
      const day = O.dateToSerial(+m[3], +m[2], +m[1]);
      if (m[4] === undefined) return { v: day, nf: 'dd/mm/yyyy' };
      const t = timeOf(m, 4);
      return t === null ? null : { v: day + t, nf: 'dd/mm/yyyy hh:mm' };
    }
    return null;
  };
  O.textToNumber = (v) => {
    const r = O.readNumber(v);
    return r ? r.v : null;
  };

  // Errores con el texto de Excel en español. Por dentro y en el archivo se usan los códigos en inglés.
  const ERR_ES = {
    '#N/A': '#N/D', '#VALUE!': '#¡VALOR!', '#DIV/0!': '#¡DIV/0!', '#NAME?': '#¿NOMBRE?', '#REF!': '#¡REF!',
    '#NUM!': '#¡NUM!', '#NULL!': '#¡NULO!', '#SPILL!': '#¡DESBORDAMIENTO!', '#CALC!': '#¡CALC!', '#ERROR!': '#¿NOMBRE?',
  };
  const ERR_EN = {};
  for (const [en, es] of Object.entries(ERR_ES)) if (en !== '#ERROR!') { ERR_EN[es.toUpperCase()] = en; ERR_EN[en] = en; }
  O.localError = (code) => ERR_ES[code] || code;
  O.errorCode = (text) => ERR_EN[String(text).trim().toUpperCase()] || null;

  // Comprueba una fórmula escrita antes de aceptarla. Como Excel, cierra los paréntesis que falten.
  // Devuelve el texto que se acepta o null si la fórmula no es válida.
  O.fixFormulaInput = (text) => {
    const t = String(text).trim();
    const ok = (s) => {
      try { parse(O.fromLocalFormula(s.slice(1))); return true; } catch (e) { return e instanceof CellError && e.code === '#REF!'; }
    };
    if (ok(t)) return t;
    const bare = t.replace(/"(?:[^"]|"")*"/g, '');
    const open = (bare.match(/\(/g) || []).length - (bare.match(/\)/g) || []).length;
    if (open > 0 && ok(t + ')'.repeat(open))) return t + ')'.repeat(open);
    return null;
  };

  // Interpreta lo que el usuario escribe: admite coma decimal, %, horas, fechas dd/mm/aaaa y errores.
  // own es la tabla de la celda: [@Ventas] se guarda como Tabla1[[#This Row],[Ventas]].
  O.parseInput = (raw, own) => {
    const s = String(raw).trim();
    if (s === '') return { empty: true };
    // Un apóstrofo delante obliga a guardar el texto tal cual, como en Excel.
    if (s.startsWith("'")) return { v: String(raw).replace(/^\s*'/, '') };
    if (s.startsWith('=') && s.length > 1) return { f: O.fromLocalFormula(s.slice(1), own) };
    const n = O.readNumber(s);
    if (n) return n;
    const err = O.errorCode(s);
    if (err) return { v: { err } };
    if (/^(verdadero|true)$/i.test(s)) return { v: true };
    if (/^(falso|false)$/i.test(s)) return { v: false };
    return { v: s };
  };

  // ---------- Libro ----------
  const MAX_DEPTH = 120;
  class Deferred {
    constructor(sheet, r, c) { this.sheet = sheet; this.r = r; this.c = c; }
  }
  class Workbook {
    constructor(state) {
      this.state = state;
      this.asts = new Map();
      this.depth = 0;
      this.clearCache();
    }
    clearCache() { this.cache = new Map(); }
    sheet(name) {
      if (name === null || name === undefined) return null;
      return this.state.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase()) || null;
    }
    // Devuelve el valor calculado, o un objeto {err} si la celda tiene un error.
    // Las cadenas muy largas se resuelven por tramos desde arriba para no agotar la pila.
    value(sheet, r, c) {
      if (this.depth > 0) return this.compute(sheet, r, c);
      const pending = [[sheet, r, c]];
      const seen = new Set();
      for (;;) {
        const [s, rr, cc] = pending[pending.length - 1];
        try {
          const res = this.compute(s, rr, cc);
          pending.pop();
          if (!pending.length) return res;
        } catch (e) {
          if (!(e instanceof Deferred)) throw e;
          this.depth = 0;
          const key = e.sheet.name + '!' + e.r + ',' + e.c;
          if (seen.has(key)) this.cache.set(key, { err: '#REF!', circular: true });
          else { seen.add(key); pending.push([e.sheet, e.r, e.c]); }
        }
      }
    }
    compute(sheet, r, c) {
      const cell = sheet.cells[O.addr(r, c)];
      if (!cell) return null;
      if (cell.f === undefined) return cell.v === undefined ? null : cell.v;
      const key = sheet.name + '!' + r + ',' + c;
      if (this.cache.has(key)) {
        const hit = this.cache.get(key);
        return hit === Workbook.BUSY ? { err: '#REF!', circular: true } : hit;
      }
      if (this.depth >= MAX_DEPTH) throw new Deferred(sheet, r, c);
      this.cache.set(key, Workbook.BUSY);
      this.depth++;
      let out;
      try {
        const res = evalNode(this.ast(cell.f), this.context(sheet, r, c));
        out = Array.isArray(res) ? (res[0] ? res[0][0] : null) : res;
        if (out === null || out === undefined) out = 0;
        if (typeof out === 'number' && !Number.isFinite(out)) out = { err: '#NUM!' };
      } catch (e) {
        if (!(e instanceof CellError)) {
          this.cache.delete(key);
          throw e;
        }
        out = { err: e.code };
      } finally {
        this.depth--;
      }
      this.cache.set(key, out);
      return out;
    }
    ast(src) {
      let ast = this.asts.get(src);
      if (!ast) { ast = parse(src); this.asts.set(src, ast); }
      return ast;
    }
    context(sheet, r, c) {
      const target = (name) => {
        const s = name === null ? sheet : this.sheet(name);
        if (!s) fail('#REF!');
        return s;
      };
      return {
        cell: (name, rr, cc) => {
          const val = this.value(target(name), rr, cc);
          if (val && val.err) fail(val.err);
          return val;
        },
        named: (name) => {
          const names = this.state.names || {};
          const key = Object.keys(names).find((k) => k.toUpperCase() === name);
          if (key) return this.ast(names[key]);
          // El nombre de una tabla sola son sus datos.
          if (O.findTable(this.state, name)) return O.resolveTableRef(this.state, sheet, r, c, name + '[#Data]');
          return fail('#NAME?');
        },
        table: (text) => O.resolveTableRef(this.state, sheet, r, c, text),
        rowState: (name, rr) => {
          const s = target(name);
          return { filtered: !!(s.filtered && s.filtered[rr]), hidden: !!(s.hiddenRows && s.hiddenRows[rr]) };
        },
        isSubtotal: (name, rr, cc) => {
          const cell = target(name).cells[O.addr(rr, cc)];
          return !!(cell && cell.f && /\bSUBTOTAL\s*\(/i.test(cell.f));
        },
      };
    }
    // Calcula una formula suelta en una hoja, como las de las reglas de formato condicional.
    evaluate(sheet, src) {
      try {
        const res = evalNode(this.ast(src), this.context(sheet));
        return Array.isArray(res) ? (res[0] ? res[0][0] : null) : res;
      } catch (e) {
        if (e instanceof CellError) return { err: e.code };
        throw e;
      }
    }
    // Como evaluate, pero devuelve todos los valores de un rango, fila a fila.
    evaluateAll(sheet, src) {
      try {
        const res = evalNode(this.ast(src), this.context(sheet));
        return Array.isArray(res) ? res.flat() : [res];
      } catch (e) {
        if (e instanceof CellError) return [{ err: e.code }];
        throw e;
      }
    }
    display(sheet, r, c) {
      const cell = sheet.cells[O.addr(r, c)];
      if (!cell) return '';
      const v = this.value(sheet, r, c);
      if (v && v.err) return O.localError(v.err);
      return O.formatValue(v, cell.s && cell.s.nf);
    }
    source(sheet, r, c) {
      const cell = sheet.cells[O.addr(r, c)];
      if (!cell) return '';
      if (cell.f !== undefined) return '=' + O.toLocalFormula(cell.f, (O.tableAt(sheet, r, c) || {}).name);
      if (cell.v === undefined || cell.v === null) return '';
      if (cell.v.err) return O.localError(cell.v.err);
      if (typeof cell.v === 'number' && !(cell.s && O.isDateFormat(cell.s.nf))) return String(cell.v).replace('.', ',');
      if (typeof cell.v === 'number') {
        // Para editar, fechas y horas se muestran en una forma que se vuelve a leer igual.
        const time = Math.abs(cell.v % 1) > 1e-9, date = Math.abs(cell.v) >= 1;
        const sec = Math.round((cell.v % 1) * 86400) % 60 !== 0;
        return O.formatValue(cell.v, [date || !time ? 'dd/mm/yyyy' : '', time ? (sec ? 'hh:mm:ss' : 'hh:mm') : ''].filter(Boolean).join(' '));
      }
      // Un texto que se leería como otra cosa lleva el apóstrofo, para que al confirmar siga siendo texto.
      if (typeof cell.v === 'string' && cell.v !== '' && (O.parseInput(cell.v).f !== undefined || typeof O.parseInput(cell.v).v !== 'string')) return "'" + cell.v;
      return O.formatValue(cell.v, cell.s && cell.s.nf);
    }
  }
  Workbook.BUSY = Symbol('busy');
  O.Workbook = Workbook;
  O.parseFormula = parse;
  O.compareValues = compare;
  O.valueText = str;
})(globalThis);
