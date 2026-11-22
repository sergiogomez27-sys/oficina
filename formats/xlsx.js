// Lectura y escritura de .xlsx sobre el modelo de libro de core/formula.js.
(function (g) {
  const O = g.Oficina;
  const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
  const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const BUILTIN = { 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 14: 'dd/mm/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'dd/mm/yyyy h:mm', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0' };

  const isTrue = (v) => v === '1' || v === 'true';
  // Paleta clásica de Excel para los colores indexados 0-63.
  const INDEXED = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
    '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
    '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
    '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333').split(' ');
  // Tema por defecto de Office (lt1, dk1, lt2, dk2, acentos 1-6, vínculos), en el orden en que lo numeran los estilos.
  const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

  // El tono ajusta la luminosidad del color en HSL, como hace Excel.
  const applyTint = (hex, tint) => {
    if (!tint) return hex;
    let [r, gr, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
    const max = Math.max(r, gr, b), min = Math.min(r, gr, b);
    let h = 0, s = 0, l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      h = max === r ? (gr - b) / d + (gr < b ? 6 : 0) : max === gr ? (b - r) / d + 2 : (r - gr) / d + 4;
      h /= 6;
    }
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    const hue = (p, q, t) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p;
    };
    if (s === 0) r = gr = b = l;
    else {
      const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
      [r, gr, b] = [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
    }
    return [r, gr, b].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  };

  O.applyTint = applyTint;

  // Color de un elemento <color>, <fgColor> o <tabColor>: rgb directo, del tema o indexado.
  const colorOf = (el, palette) => {
    if (!el) return null;
    const tint = +O.attr(el, 'tint') || 0;
    const v = O.attr(el, 'rgb');
    if (v) return applyTint(v.slice(-6).toUpperCase(), tint);
    const th = O.attr(el, 'theme');
    if (th !== null) return palette.theme[+th] ? applyTint(palette.theme[+th], tint) : null;
    const ix = O.attr(el, 'indexed');
    if (ix !== null) return palette.indexed[+ix] ? applyTint(palette.indexed[+ix], tint) : null;
    return null;
  };
  // Color del tema con su tono, para volver a guardarlo como color del tema y que siga al tema en Excel.
  const themeRef = (el) => {
    const th = O.attr(el, 'theme');
    return el && th !== null && !O.attr(el, 'rgb') ? [+th, +O.attr(el, 'tint') || 0] : null;
  };
  const colorXml = (tag, hex, th, palette) => (th && palette && applyTint(palette[th[0]] || '', th[1]) === hex
    ? `<${tag} theme="${th[0]}"${th[1] ? ` tint="${th[1]}"` : ''}/>` : `<${tag} rgb="FF${hex}"/>`);

  async function readPalette(zip, themePart, stylesRoot) {
    const theme = THEME.slice();
    const xml = themePart && (await O.readXml(zip, themePart));
    const scheme = xml && O.all(xml.documentElement, 'clrScheme')[0];
    if (scheme) {
      const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      order.forEach((name, i) => {
        const slot = O.kid(scheme, name);
        const v = O.attr(O.kid(slot, 'srgbClr'), 'val') || O.attr(O.kid(slot, 'sysClr'), 'lastClr');
        if (v) theme[i] = v.toUpperCase();
      });
    }
    let indexed = INDEXED;
    const custom = stylesRoot && O.kids(O.path(stylesRoot, 'colors', 'indexedColors'), 'rgbColor');
    if (custom && custom.length) indexed = custom.map((c) => (O.attr(c, 'rgb') || '').slice(-6).toUpperCase());
    return { theme, indexed };
  }

  // En el archivo, las funciones modernas llevan un prefijo; sin el, Excel muestra #NAME?.
  const fileFormula = (src) => src.replace(/"(?:[^"]|"")*"|(?<![\w.])(CONCAT|IFS|IFNA|MAXIFS|MINIFS|TEXTJOIN|XLOOKUP|RANK\.EQ)(?=\()/gi, (m) => (m[0] === '"' ? m : '_xlfn.' + m.toUpperCase()));

  // Formato que aplica una regla condicional (dxf): fuente y relleno.
  const readDxf = (dxf, rgb) => {
    const s = {};
    const font = O.kid(dxf, 'font');
    if (font) {
      if (O.isOn(O.kid(font, 'b'))) s.b = true;
      if (O.isOn(O.kid(font, 'i'))) s.i = true;
      if (O.isOn(O.kid(font, 'u'))) s.u = true;
      if (O.isOn(O.kid(font, 'strike'))) s.strike = true;
      const color = rgb(O.kid(font, 'color'));
      if (color) s.color = color;
    }
    const pattern = O.path(dxf, 'fill', 'patternFill');
    if (pattern && O.attr(pattern, 'patternType') !== 'none') {
      // En los dxf Excel pinta el relleno sólido con bgColor; si solo hay fgColor, lo muestra en blanco.
      const fill = rgb(O.kid(pattern, 'bgColor')) || (O.kid(pattern, 'fgColor') ? 'FFFFFF' : null);
      if (fill) s.fill = fill;
    }
    return s;
  };
  const DV_ATTRS = ['type', 'errorStyle', 'imeMode', 'operator', 'allowBlank', 'showDropDown', 'showInputMessage', 'showErrorMessage', 'errorTitle', 'error', 'promptTitle', 'prompt'];
  const CF_TYPES = new Set(['cellIs', 'expression', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks',
    'containsErrors', 'notContainsErrors', 'duplicateValues', 'uniqueValues', 'top10', 'aboveAverage', 'timePeriod', 'colorScale', 'dataBar', 'iconSet']);
  function readRules(root, dxfs, rgb, warnings) {
    const rules = [];
    for (const block of O.kids(root, 'conditionalFormatting')) {
      const ref = O.attr(block, 'sqref');
      for (const el of O.kids(block, 'cfRule')) {
        const type = O.attr(el, 'type');
        if (!CF_TYPES.has(type)) { warnings.add('algunas reglas de formato condicional'); continue; }
        const rule = { ref, type, prio: +O.attr(el, 'priority') || 0 };
        const dxfId = O.attr(el, 'dxfId');
        if (dxfId !== null && dxfs[+dxfId]) rule.style = { ...dxfs[+dxfId] };
        const f = O.kids(el, 'formula').map((x) => x.textContent.replace(/_xl(fn|ws)\./g, ''));
        if (f.length) rule.f = f;
        const at = (n) => O.attr(el, n);
        if (at('operator')) rule.op = at('operator');
        if (at('text') !== null) rule.text = at('text');
        if (at('rank') !== null) rule.rank = +at('rank');
        if (isTrue(at('percent'))) rule.percent = true;
        if (isTrue(at('bottom'))) rule.bottom = true;
        if (at('aboveAverage') === '0' || at('aboveAverage') === 'false') rule.above = false;
        if (isTrue(at('equalAverage'))) rule.equal = true;
        if (at('stdDev') !== null) rule.std = +at('stdDev');
        if (isTrue(at('stopIfTrue'))) rule.stop = true;
        if (at('timePeriod')) rule.period = at('timePeriod');
        const point = (v) => ({ type: O.attr(v, 'type'), val: O.attr(v, 'val') });
        const scale = O.kid(el, 'colorScale');
        if (scale) {
          const colors = O.kids(scale, 'color');
          rule.scale = O.kids(scale, 'cfvo').map((v, i) => ({ ...point(v), color: rgb(colors[i]) || 'FFFFFF' }));
        }
        const bar = O.kid(el, 'dataBar');
        if (bar) {
          const [min, max] = O.kids(bar, 'cfvo').map(point);
          rule.bar = { min, max, color: rgb(O.kid(bar, 'color')) || '638EC6' };
        }
        const icons = O.kid(el, 'iconSet');
        if (icons) {
          const set = O.attr(icons, 'iconSet') || '3TrafficLights1';
          // Los conjuntos que Excel no tiene en la lista conocida se muestran con semáforos.
          rule.icons = { set: O.ICON_SETS && O.ICON_SETS[set] ? set : '3TrafficLights1' };
          if (set !== rule.icons.set) warnings.add('algunos conjuntos de iconos');
          rule.icons.points = O.kids(icons, 'cfvo').map((v) => ({ ...point(v), ...(O.attr(v, 'gte') === '0' || O.attr(v, 'gte') === 'false' ? { gte: false } : {}) }));
          if (O.attr(icons, 'showValue') === '0' || O.attr(icons, 'showValue') === 'false') rule.icons.showValue = false;
          if (isTrue(O.attr(icons, 'reverse'))) rule.icons.reverse = true;
        }
        rules.push(rule);
      }
    }
    rules.sort((a, b) => a.prio - b.prio);
    for (const r of rules) delete r.prio;
    return rules;
  }

  // Formulas que Excel espera en las reglas de texto, blancos, errores y fechas, relativas a la primera celda.
  const ruleFormula = (rule, cell) => {
    const t = `"${String(rule.text || '').replace(/"/g, '""')}"`;
    const d = `ROUNDDOWN(${cell},0)`;
    const month = (k) => `AND(MONTH(${cell})=MONTH(EDATE(TODAY(),0${k})),YEAR(${cell})=YEAR(EDATE(TODAY(),0${k})))`;
    return {
      containsText: `NOT(ISERROR(SEARCH(${t},${cell})))`, notContainsText: `ISERROR(SEARCH(${t},${cell}))`,
      beginsWith: `LEFT(${cell},LEN(${t}))=${t}`, endsWith: `RIGHT(${cell},LEN(${t}))=${t}`,
      containsBlanks: `LEN(TRIM(${cell}))=0`, notContainsBlanks: `LEN(TRIM(${cell}))>0`,
      containsErrors: `ISERROR(${cell})`, notContainsErrors: `NOT(ISERROR(${cell}))`,
    }[rule.type] || (rule.type === 'timePeriod' ? {
      today: `FLOOR(${cell},1)=TODAY()`, yesterday: `FLOOR(${cell},1)=TODAY()-1`, tomorrow: `FLOOR(${cell},1)=TODAY()+1`,
      last7Days: `AND(TODAY()-FLOOR(${cell},1)<=6,FLOOR(${cell},1)<=TODAY())`,
      thisWeek: `AND(TODAY()-${d}<=WEEKDAY(TODAY())-1,${d}-TODAY()<=7-WEEKDAY(TODAY()))`,
      lastWeek: `AND(TODAY()-${d}>=(WEEKDAY(TODAY())),TODAY()-${d}<(WEEKDAY(TODAY())+7))`,
      nextWeek: `AND(${d}-TODAY()>(7-WEEKDAY(TODAY())),${d}-TODAY()<(15-WEEKDAY(TODAY())))`,
      thisMonth: `AND(MONTH(${cell})=MONTH(TODAY()),YEAR(${cell})=YEAR(TODAY()))`, lastMonth: month('-1'), nextMonth: month('+1'),
    }[rule.period] : null);
  };

  async function readStyles(zip, part, palette) {
    const xml = part && (await O.readXml(zip, part));
    if (!xml) return { xfs: [], dxfs: [] };
    const root = xml.documentElement;
    const rgb = (el) => colorOf(el, palette);
    const dxfs = O.kids(O.kid(root, 'dxfs'), 'dxf').map((d) => readDxf(d, rgb));
    const custom = {};
    for (const n of O.kids(O.kid(root, 'numFmts'), 'numFmt')) custom[O.attr(n, 'numFmtId')] = O.attr(n, 'formatCode');
    const fonts = O.kids(O.kid(root, 'fonts'), 'font');
    const fills = O.kids(O.kid(root, 'fills'), 'fill');
    const borders = O.kids(O.kid(root, 'borders'), 'border');
    const xfs = O.kids(O.kid(root, 'cellXfs'), 'xf').map((xf) => {
      const s = {};
      const font = fonts[+O.attr(xf, 'fontId')];
      if (font) {
        if (O.isOn(O.kid(font, 'b'))) s.b = true;
        if (O.isOn(O.kid(font, 'i'))) s.i = true;
        const color = rgb(O.kid(font, 'color'));
        if (color && color !== '000000') {
          s.color = color;
          if (themeRef(O.kid(font, 'color'))) s.thColor = themeRef(O.kid(font, 'color'));
        }
        if (O.isOn(O.kid(font, 'u'))) s.u = true;
        if (O.isOn(O.kid(font, 'strike'))) s.strike = true;
        const sz = +O.attr(O.kid(font, 'sz'), 'val');
        if (sz && sz !== 11) s.sz = sz;
        const name = O.attr(O.kid(font, 'name'), 'val');
        if (name && name !== 'Calibri') s.font = name;
      }
      const border = borders[+O.attr(xf, 'borderId')];
      if (border) {
        const sides = [['left', 'l'], ['right', 'r'], ['top', 't'], ['bottom', 'b']].filter(([side]) => O.attr(O.kid(border, side), 'style') && O.attr(O.kid(border, side), 'style') !== 'none');
        if (sides.length) s.bd = sides.map(([, k]) => k).join('');
        // Grosor y color de cada lado cuando no es la línea fina negra de siempre.
        const bdx = {};
        for (const [side, k] of sides) {
          const el = O.kid(border, side), style = O.attr(el, 'style'), color = rgb(O.kid(el, 'color')) || '000000';
          const th = themeRef(O.kid(el, 'color'));
          if (style !== 'thin' || color !== '000000') bdx[k] = th ? [style, color, th] : [style, color];
        }
        if (Object.keys(bdx).length) s.bdx = bdx;
      }
      const pattern = O.kid(fills[+O.attr(xf, 'fillId')], 'patternFill');
      const fill = O.attr(pattern, 'patternType') !== 'none' && rgb(O.kid(pattern, 'fgColor'));
      if (fill) {
        s.fill = fill;
        if (themeRef(O.kid(pattern, 'fgColor'))) s.thFill = themeRef(O.kid(pattern, 'fgColor'));
      }
      // Bloqueo y fórmula oculta: solo cuentan con la hoja protegida.
      const prot = O.kid(xf, 'protection');
      if (O.attr(prot, 'locked') === '0' || O.attr(prot, 'locked') === 'false') s.unlocked = true;
      if (isTrue(O.attr(prot, 'hidden'))) s.hideF = true;
      const id = O.attr(xf, 'numFmtId');
      const nf = custom[id] || BUILTIN[id];
      if (nf) s.nf = nf;
      const align = O.attr(O.kid(xf, 'alignment'), 'horizontal');
      if (['left', 'center', 'right'].includes(align)) s.align = align;
      const valign = O.attr(O.kid(xf, 'alignment'), 'vertical');
      if (['top', 'center'].includes(valign)) s.valign = valign;
      const wrap = O.attr(O.kid(xf, 'alignment'), 'wrapText');
      if (wrap === '1' || wrap === 'true') s.wrap = true;
      const rot = +O.attr(O.kid(xf, 'alignment'), 'textRotation');
      if (rot) s.rot = rot;
      return s;
    });
    return { xfs, dxfs, tableStyles: readTableStyles(root, rgb) };
  }

  // Estilos de tabla personalizados: cada parte guarda su dxf tal cual, para volver a escribirlo igual,
  // y su versión con colores fijos para dibujarla como los integrados.
  const STYLE_PARTS = {
    wholeTable: 'w', headerRow: 'h', totalRow: 't', firstColumn: 'f', lastColumn: 'l', firstRowStripe: 'r1', secondRowStripe: 'r2',
    firstColumnStripe: 'c1', secondColumnStripe: 'c2', firstHeaderCell: 'fh', lastHeaderCell: 'lh', firstTotalCell: 'ft', lastTotalCell: 'lt',
  };
  function tablePart(dxf, rgb) {
    const s = readDxf(dxf, rgb), e = {};
    if (s.fill) e.fill = s.fill;
    if (s.color) e.color = s.color;
    if (s.b) e.b = 1;
    const bd = {};
    for (const [side, k] of [['left', 'l'], ['right', 'r'], ['top', 't'], ['bottom', 'b'], ['vertical', 'v'], ['horizontal', 'h']]) {
      const el = O.kid(O.kid(dxf, 'border'), side), style = O.attr(el, 'style');
      if (style && style !== 'none') bd[k] = [style, rgb(O.kid(el, 'color')) || '000000'];
    }
    if (Object.keys(bd).length) e.bd = bd;
    return e;
  }
  function readTableStyles(root, rgb) {
    const box = O.kid(root, 'tableStyles');
    if (!box) return null;
    const dxfEls = O.kids(O.kid(root, 'dxfs'), 'dxf');
    const out = { def: O.attr(box, 'defaultTableStyle'), pivot: O.attr(box, 'defaultPivotStyle'), list: [] };
    for (const st of O.kids(box, 'tableStyle')) {
      const item = { name: O.attr(st, 'name'), pivotOnly: O.attr(st, 'table') === '0', tableOnly: O.attr(st, 'pivot') === '0', parts: [], def: {} };
      for (const el of O.kids(st, 'tableStyleElement')) {
        const type = O.attr(el, 'type'), dxf = dxfEls[+O.attr(el, 'dxfId')];
        const part = { type, raw: dxf ? serialize(dxf).replace(/ xmlns="[^"]*"/, '') : '<dxf/>' };
        if (O.attr(el, 'size')) part.size = +O.attr(el, 'size');
        item.parts.push(part);
        const k = STYLE_PARTS[type];
        if (k && dxf) item.def[k] = { ...tablePart(dxf, rgb), ...(part.size ? { size: part.size } : {}) };
      }
      out.list.push(item);
    }
    return out.list.length || out.def ? out : null;
  }

  async function readXlsx(buffer) {
    const zip = await g.JSZip.loadAsync(buffer);
    const main = await O.mainPart(zip);
    const wb = await O.readXml(zip, main);
    const rels = await O.readRels(zip, main);
    const relOf = (suffix) => Object.values(rels).find((r) => r.type.endsWith(suffix));
    const warnings = new Set();

    let shared = [];
    const ssRel = relOf('/sharedStrings');
    if (ssRel) {
      const xml = await O.readXml(zip, ssRel.target);
      shared = O.kids(xml && xml.documentElement, 'si').map((si) =>
        O.all(si, 't').filter((t) => t.parentNode.localName !== 'rPh').map((t) => t.textContent).join(''));
    }
    const stylesRel = relOf('/styles');
    const themeRel = relOf('/theme');
    const stylesXml = stylesRel && (await O.readXml(zip, stylesRel.target));
    const palette = await readPalette(zip, themeRel && themeRel.target, stylesXml && stylesXml.documentElement);
    const rgb = (el) => colorOf(el, palette);
    const { xfs, dxfs, tableStyles } = await readStyles(zip, stylesRel && stylesRel.target, palette);
    const extra = { themeColors: palette.theme };
    if (tableStyles) extra.tableStyles = tableStyles;
    if (themeRel && zip.file(themeRel.target)) extra.theme = await zip.file(themeRel.target).async('string');
    // Macros: el proyecto VBA se guarda tal cual. Oficina no las ejecuta, como Excel con las macros deshabilitadas.
    const vbaRel = relOf('/vbaProject');
    if (vbaRel && zip.file(vbaRel.target)) extra.vba = await zip.file(vbaRel.target).async('base64');
    // Un libro .xlsm se sigue guardando como .xlsm aunque no tenga macros, como en Excel.
    const types = zip.file('[Content_Types].xml') ? await zip.file('[Content_Types].xml').async('string') : '';
    if (extra.vba || types.includes('macroEnabled.main+xml')) extra.macro = true;
    const wbPr = O.kid(wb.documentElement, 'workbookPr');
    if (O.attr(wbPr, 'codeName')) extra.codeName = O.attr(wbPr, 'codeName');
    const wbProt = O.kid(wb.documentElement, 'workbookProtection');
    if (wbProt) extra.wbProtect = attrsOf(wbProt);
    // Tablas dinámicas: la caché y la tabla se guardan tal cual y Excel las actualiza al abrir.
    const pivotCaches = [];
    for (const pc of O.kids(O.kid(wb.documentElement, 'pivotCaches'), 'pivotCache')) {
      const rel = rels[O.attr(pc, 'id')];
      if (!rel || !zip.file(rel.target)) continue;
      const def = await zip.file(rel.target).async('string');
      const defRels = await O.readRels(zip, rel.target);
      const recId = Object.keys(defRels).find((k) => defRels[k].type.endsWith('/pivotCacheRecords'));
      const records = recId && zip.file(defRels[recId].target) ? await zip.file(defRels[recId].target).async('string') : null;
      pivotCaches.push({ id: O.attr(pc, 'cacheId'), part: rel.target, def, records, recId: records ? recId : null });
    }
    const media = {};
    const mediaKey = new Map();
    const ctypes = { over: {}, def: {} };
    for (const el of O.kids(types ? O.parseXml(types).documentElement : null)) {
      if (el.localName === 'Override') ctypes.over[(O.attr(el, 'PartName') || '').replace(/^\//, '')] = O.attr(el, 'ContentType');
      if (el.localName === 'Default') ctypes.def[(O.attr(el, 'Extension') || '').toLowerCase()] = O.attr(el, 'ContentType');
    }
    const parts = {};
    const ctx = { zip, warnings, media, mediaKey, theme: palette.theme, parts, ctypes };
    // Cachés de segmentaciones y escalas de tiempo del libro.
    const wbExts = [];
    for (const ext of O.kids(O.kid(wb.documentElement, 'extLst'), 'ext')) {
      if (O.kids(ext).some((k) => ['slicerCaches', 'timelineCacheRefs'].includes(k.localName))) wbExts.push(await keepFragment(ctx, ext, rels));
    }
    if (wbExts.length) extra.wbExts = wbExts;

    const sheets = [];
    for (const sh of O.kids(O.kid(wb.documentElement, 'sheets'), 'sheet')) {
      const rel = rels[O.attr(sh, 'id')];
      const xml = rel && (await O.readXml(zip, rel.target));
      if (!xml) { warnings.add('hojas de gráfico'); continue; }
      const root = xml.documentElement;
      const sheet = { name: O.attr(sh, 'name'), cells: {}, cols: {} };
      // Las segmentaciones de tablas dinámicas nombran la hoja por su número: se conserva.
      if (+O.attr(sh, 'sheetId')) sheet.sheetId = +O.attr(sh, 'sheetId');
      // Ancho y alto por defecto de la hoja, si no son los de Excel.
      const fmtPr = O.kid(root, 'sheetFormatPr');
      if (O.attr(fmtPr, 'defaultColWidth')) sheet.defW = Math.round(+O.attr(fmtPr, 'defaultColWidth') * 7 + 5);
      if (O.attr(fmtPr, 'defaultRowHeight') && Math.abs(+O.attr(fmtPr, 'defaultRowHeight') - 15) > 0.01) sheet.defH = Math.round(+O.attr(fmtPr, 'defaultRowHeight') / 0.75);
      if (sheet.defW === 64) delete sheet.defW;
      for (const col of O.kids(O.kid(root, 'cols'), 'col')) {
        const width = +O.attr(col, 'width');
        const last = Math.min(+O.attr(col, 'max'), 200);
        const off = isTrue(O.attr(col, 'hidden'));
        for (let i = +O.attr(col, 'min'); i <= last; i++) {
          if (width && !off) sheet.cols[i - 1] = Math.round(width * 7 + 5);
          if (off) (sheet.hiddenCols = sheet.hiddenCols || {})[i - 1] = true;
        }
      }
      for (const row of O.kids(O.kid(root, 'sheetData'), 'row')) {
        const r = +O.attr(row, 'r') - 1;
        if (isTrue(O.attr(row, 'hidden'))) (sheet.hiddenRows = sheet.hiddenRows || {})[r] = true;
        if (isTrue(O.attr(row, 'customHeight')) && +O.attr(row, 'ht')) (sheet.rowH = sheet.rowH || {})[r] = Math.round(+O.attr(row, 'ht') / 0.75);
      }
      if (O.attr(sh, 'state') === 'hidden' || O.attr(sh, 'state') === 'veryHidden') sheet.hidden = true;
      if (O.attr(sh, 'state') === 'veryHidden') sheet.veryHidden = true;
      const prot = O.kid(root, 'sheetProtection');
      if (prot) {
        sheet.protect = attrsOf(prot);
        if (isTrue(O.attr(prot, 'sheet'))) sheet.protected = true;
      }
      const tabEl = O.path(root, 'sheetPr', 'tabColor');
      const tab = rgb(tabEl);
      if (tab) sheet.tabColor = tab;
      if (tab && themeRef(tabEl)) sheet.thTab = themeRef(tabEl);
      if (O.attr(O.kid(root, 'sheetPr'), 'codeName')) sheet.codeName = O.attr(O.kid(root, 'sheetPr'), 'codeName');
      const merged = O.kids(O.kid(root, 'mergeCells'), 'mergeCell').map((m) => O.attr(m, 'ref')).filter((ref) => ref && ref.includes(':'));
      if (merged.length) sheet.merges = merged;
      const rules = readRules(root, dxfs, rgb, warnings);
      if (rules.length) sheet.cf = rules;
      const readDv = (el) => {
        const attrs = {};
        for (const n of DV_ATTRS) if (O.attr(el, n) !== null) attrs[n] = O.attr(el, n);
        // Las de Excel 2010 que miran otra hoja llevan la fórmula en <xm:f> y el rango en <xm:sqref>.
        const f = ['formula1', 'formula2'].map((n) => O.kid(el, n)).filter(Boolean).map((x) => (O.kid(x, 'f') || x).textContent.replace(/_xl(fn|ws)\./g, ''));
        const ref = O.attr(el, 'sqref') || (O.kid(el, 'sqref') || {}).textContent;
        return { ref, attrs, f };
      };
      const dv = [...O.kids(O.kid(root, 'dataValidations'), 'dataValidation'),
        ...O.all(O.kid(root, 'extLst'), 'dataValidations').flatMap((x) => O.kids(x, 'dataValidation'))].map(readDv).filter((d) => d.ref);
      if (dv.length) sheet.dv = dv;
      const af = O.kid(root, 'autoFilter');
      if (af && O.attr(af, 'ref')) {
        sheet.autoFilter = readAutoFilter(af, dxfs);
        // Las filas ocultas dentro del rango las ha ocultado el filtro.
        const box = O.parseRanges(sheet.autoFilter.ref)[0];
        for (const r of Object.keys(sheet.hiddenRows || {})) if (box && +r > box.r1 && +r <= box.r2) (sheet.filtered = sheet.filtered || {})[r] = true;
      }
      const sheetRels = await O.readRels(zip, rel.target);
      const drawingEl = O.kid(root, 'drawing');
      if (drawingEl && sheetRels[O.attr(drawingEl, 'id')]) {
        const items = await readDrawing(sheetRels[O.attr(drawingEl, 'id')].target, ctx);
        if (items.length) sheet.drawings = items;
      }
      const pivots = [];
      for (const r of Object.values(sheetRels).filter((x) => x.type.endsWith('/pivotTable'))) {
        if (!zip.file(r.target)) continue;
        const xmlText = await zip.file(r.target).async('string');
        const pRels = await O.readRels(zip, r.target);
        const relId = Object.keys(pRels).find((k) => pRels[k].type.endsWith('/pivotCacheDefinition'));
        const cache = relId && pivotCaches.findIndex((p) => p.part === pRels[relId].target);
        if (cache === undefined || cache < 0) { warnings.add('algunas tablas dinámicas'); continue; }
        pivots.push({ xml: xmlText, relId, cache: pivotCaches[cache].id });
      }
      if (pivots.length) sheet.pivots = pivots;
      const tables = [];
      for (const tp of O.kids(O.kid(root, 'tableParts'), 'tablePart')) {
        const r = sheetRels[O.attr(tp, 'id')];
        const doc = r && (await O.readXml(zip, r.target));
        if (doc) tables.push(readTable(doc.documentElement, dxfs, warnings));
      }
      if (tables.length) {
        sheet.tables = tables;
        // Las filas ocultas dentro de una tabla filtrada las ha ocultado su filtro.
        for (const t of tables) {
          if (!t.autoFilter || !Object.keys(t.autoFilter.cols).length) continue;
          const d = O.tableData(t);
          for (const r of Object.keys(sheet.hiddenRows || {})) if (+r >= d.r1 && +r <= d.r2) (sheet.filtered = sheet.filtered || {})[r] = true;
        }
      }
      // Controles de formulario y ActiveX, objetos OLE y segmentaciones: se guardan tal cual.
      const raw = [];
      for (const el of O.kids(root)) {
        const inner = el.localName === 'AlternateContent' ? O.kids(O.kid(el, 'Choice'))[0] : el;
        if (inner && ['controls', 'oleObjects'].includes(inner.localName)) raw.push(await keepFragment(ctx, el, sheetRels));
      }
      if (raw.length) sheet.ctrlXml = raw;
      const controls = await readControls(root, sheetRels, ctx);
      const exts = [];
      for (const ext of O.kids(O.kid(root, 'extLst'), 'ext')) {
        if (O.kids(ext).some((k) => ['slicerList', 'timelineRefs'].includes(k.localName))) exts.push(await keepFragment(ctx, ext, sheetRels));
      }
      if (exts.length) sheet.exts = exts;
      // Dibujo heredado (VML): las formas de los controles se guardan; las notas se vuelven a crear.
      const vmlEl = O.kid(root, 'legacyDrawing');
      const vmlRel = vmlEl && sheetRels[O.attr(vmlEl, 'id')];
      if (vmlRel && zip.file(vmlRel.target)) {
        const text = await zip.file(vmlRel.target).async('string');
        const shapes = [...text.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)].map((m) => m[0]).filter((s) => !/ObjectType="Note"/.test(s));
        if (shapes.length) {
          const typesXml = [...text.matchAll(/<v:shapetype\b[\s\S]*?<\/v:shapetype>/g)].map((m) => m[0]).join('');
          sheet.vml = { data: (/<o:idmap\b[^>]*\bdata="([^"]+)"/.exec(text) || [])[1] || '', ...(await keepFragment(ctx, typesXml + shapes.join(''), await O.readRels(zip, vmlRel.target))) };
          for (const ctl of controls) {
            const shape = shapes.find((s) => new RegExp(`id="_x0000_s${ctl.id}"`).test(s));
            const box = shape && /<v:textbox\b[^>]*>([\s\S]*?)<\/v:textbox>/.exec(shape);
            if (box && textOf(box[1])) ctl.label = textOf(box[1]);
          }
        }
      }
      if (controls.length) sheet.controls = controls;
      const sharedF = {};
      for (const c of O.all(O.kid(root, 'sheetData'), 'c')) {
        const ref = O.attr(c, 'r');
        const t = O.attr(c, 't');
        const vEl = O.kid(c, 'v');
        const v = vEl ? vEl.textContent : null;
        const fEl = O.kid(c, 'f');
        const cell = {};
        const style = xfs[+O.attr(c, 's')];
        if (style && Object.keys(style).length) cell.s = { ...style };
        if (fEl) {
          const si = O.attr(fEl, 'si');
          const here = O.parseAddr(ref);
          if (O.attr(fEl, 't') === 'shared' && si !== null) {
            if (fEl.textContent) { sharedF[si] = { f: fEl.textContent, at: here }; cell.f = fEl.textContent; }
            else if (sharedF[si]) cell.f = O.shiftFormula(sharedF[si].f, here.r - sharedF[si].at.r, here.c - sharedF[si].at.c);
          } else cell.f = fEl.textContent;
          if (O.attr(fEl, 't') === 'array') warnings.add('fórmulas matriciales');
          if (cell.f !== undefined) cell.f = cell.f.replace(/_xl(fn|ws)\./g, '');
        }
        if (cell.f === undefined) {
          if (t === 's') cell.v = shared[+v];
          else if (t === 'inlineStr') cell.v = O.all(c, 't').map((x) => x.textContent).join('');
          else if (t === 'str') cell.v = v;
          else if (t === 'e') cell.v = { err: v };
          else if (t === 'b') cell.v = v === '1';
          else if (v !== null) cell.v = parseFloat(v);
          else if (!cell.s) continue;
        }
        sheet.cells[ref] = cell;
      }
      const addOn = (ref, prop, value) => {
        (sheet.cells[ref] = sheet.cells[ref] || {})[prop] = value;
      };
      for (const link of O.kids(O.kid(root, 'hyperlinks'), 'hyperlink')) {
        const target = sheetRels[O.attr(link, 'id')];
        if (target && target.external) addOn(O.attr(link, 'ref').split(':')[0], 'link', target.target);
      }
      const notesRel = Object.values(sheetRels).find((r) => r.type.endsWith('/comments'));
      const notes = notesRel && (await O.readXml(zip, notesRel.target));
      for (const note of O.all(notes && notes.documentElement, 'comment')) {
        addOn(O.attr(note, 'ref'), 'note', O.all(note, 't').filter((t) => t.parentNode.localName !== 'rPh').map((t) => t.textContent).join(''));
      }
      sheets.push(sheet);
    }
    if (!sheets.length) throw new Error('El libro no contiene hojas de datos.');
    const names = {};
    for (const dn of O.kids(O.kid(wb.documentElement, 'definedNames'), 'definedName')) {
      const name = O.attr(dn, 'name');
      if (!name.startsWith('_xlnm.') && O.attr(dn, 'localSheetId') === null) names[name] = dn.textContent;
    }
    const firstVisible = sheets.findIndex((s) => !s.hidden);
    if (firstVisible < 0) delete sheets[0].hidden;
    // Solo se guardan las cachés que usa alguna tabla dinámica.
    const usedCaches = new Set(sheets.flatMap((s) => (s.pivots || []).map((p) => p.cache)));
    if (usedCaches.size) extra.pivotCaches = pivotCaches.filter((p) => usedCaches.has(p.id)).map(({ part, ...p }) => p);
    if (Object.keys(media).length) extra.media = media;
    // Qué muestra cada segmentación: una columna de una tabla o un campo de una tabla dinámica.
    const xmlOf = (path) => (parts[path] && typeof parts[path].data === 'string' ? O.parseXml(parts[path].data).documentElement : null);
    const slicerCaches = {};
    for (const frag of wbExts) {
      for (const l of Object.values(frag.links)) {
        const el = l.part && xmlOf(l.part);
        if (el && el.localName === 'slicerCacheDefinition') { el.partPath = l.part; slicerCaches[O.attr(el, 'name')] = el; }
      }
    }
    for (const sheet of sheets) {
      const slicers = {};
      for (const frag of sheet.exts || []) for (const l of Object.values(frag.links)) for (const s of O.kids(l.part && xmlOf(l.part), 'slicer')) slicers[O.attr(s, 'name')] = s;
      for (const d of sheet.drawings || []) {
        const s = d.kind === 'slicer' && slicers[d.slicer.name];
        if (!s) continue;
        d.slicer.caption = O.attr(s, 'caption') || d.slicer.name;
        if (+O.attr(s, 'columnCount')) d.slicer.columns = +O.attr(s, 'columnCount');
        const cache = slicerCaches[O.attr(s, 'cache')];
        const tsc = cache && O.all(cache, 'tableSlicerCache')[0];
        const tab = cache && O.all(cache, 'tabular')[0];
        if (tsc) Object.assign(d.slicer, { table: +O.attr(tsc, 'tableId'), column: +O.attr(tsc, 'column') });
        else if (tab) {
          const pc = pivotCaches.find((p) => p.def.includes(`pivotCacheId="${O.attr(tab, 'pivotCacheId')}"`));
          const field = pc && O.kids(O.kid(O.parseXml(pc.def).documentElement, 'cacheFields'), 'cacheField').find((f) => O.attr(f, 'name') === O.attr(cache, 'sourceName'));
          const shared = O.kids(O.kid(field, 'sharedItems')).map((x) => O.attr(x, 'v') || '');
          d.slicer.items = O.all(tab, 'i').map((i) => ({ name: shared[+O.attr(i, 'x')] || '', x: +O.attr(i, 'x'), on: isTrue(O.attr(i, 's')) }));
          // Para filtrar en Oficina: la caché de la segmentación y las tablas dinámicas que conecta.
          Object.assign(d.slicer, {
            cachePart: cache.partPath, source: O.attr(cache, 'sourceName'), shared: shared.length,
            pivots: O.all(cache, 'pivotTable').map((p) => ({ tabId: O.attr(p, 'tabId'), name: O.attr(p, 'name') })),
          });
        }
      }
    }
    if (Object.keys(parts).length) extra.parts = parts;
    return { state: { sheets, active: Math.max(0, firstVisible), names, ...extra }, warnings: [...warnings] };
  }

  // Tema mínimo de Office con los colores dados (lt1, dk1, lt2, dk2, acentos y vínculos) y Calibri.
  const themePart = (colors) => {
    const srgb = (i) => `<a:srgbClr val="${colors[i]}"/>`;
    const slot = (name, inner) => `<a:${name}>${inner}</a:${name}>`;
    const three = (x) => x + x + x;
    const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    return `${O.XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office">`
      + slot('dk1', `<a:sysClr val="windowText" lastClr="${colors[1]}"/>`) + slot('lt1', `<a:sysClr val="window" lastClr="${colors[0]}"/>`)
      + slot('dk2', srgb(3)) + slot('lt2', srgb(2)) + [4, 5, 6, 7, 8, 9].map((i) => slot('accent' + (i - 3), srgb(i))).join('') + slot('hlink', srgb(10)) + slot('folHlink', srgb(11))
      + '</a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
      + `<a:fmtScheme name="Office"><a:fillStyleLst>${three(fill)}</a:fillStyleLst><a:lnStyleLst>${[6350, 12700, 19050].map((w) => `<a:ln w="${w}">${fill}</a:ln>`).join('')}</a:lnStyleLst>`
      + `<a:effectStyleLst>${three('<a:effectStyle><a:effectLst/></a:effectStyle>')}</a:effectStyleLst><a:bgFillStyleLst>${three(fill)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
  };
  const attrsOf = (el) => {
    const out = {};
    for (const a of el.attributes) if (!a.name.startsWith('xmlns')) out[a.localName] = a.value;
    return out;
  };
  const serialize = (el) => new XMLSerializer().serializeToString(el);
  const EMU = 9525;
  const anchorPos = (el) => el && {
    c: +O.kid(el, 'col').textContent, r: +O.kid(el, 'row').textContent,
    dx: Math.round(+O.kid(el, 'colOff').textContent / EMU), dy: Math.round(+O.kid(el, 'rowOff').textContent / EMU),
  };

  // Partes que Oficina no interpreta pero guarda tal cual, junto con las que enlazan a su vez.
  async function grabPart(ctx, path) {
    if (ctx.parts[path]) return path;
    const f = ctx.zip.file(path);
    if (!f) return null;
    const type = ctx.ctypes.over[path] || ctx.ctypes.def[path.split('.').pop().toLowerCase()] || 'application/octet-stream';
    const node = (ctx.parts[path] = { type, rels: [] });
    node.data = /xml/.test(type) || /\.(xml|vml)$/i.test(path) ? await f.async('string') : { b64: await f.async('base64') };
    for (const [id, r] of Object.entries(await O.readRels(ctx.zip, path))) {
      node.rels.push(r.external ? { id, type: r.type, url: r.target } : { id, type: r.type, part: await grabPart(ctx, r.target) });
    }
    return path;
  }
  // Un trozo de XML que se guarda tal cual, con las partes a las que apunta (r:id, r:embed, o:relid...).
  // Al copiar un trozo, el prefijo de las relaciones puede cambiar de nombre (r:, ns1:...).
  const LINK = /(\b\w+:(?:id|embed|link|pict|href)="|\bo:relid=")([^"]+)"/g;
  async function keepFragment(ctx, el, rels) {
    const xml = typeof el === 'string' ? el : serialize(el);
    const links = {};
    for (const m of xml.matchAll(LINK)) {
      const r = rels[m[2]];
      if (!r || links[m[2]]) continue;
      links[m[2]] = r.external ? { type: r.type, url: r.target } : { type: r.type, part: await grabPart(ctx, r.target) };
    }
    return { xml, links };
  }
  const textOf = (html) => html.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/\s*\n\s*/g, ' ').replace(/\s+/g, ' ').trim();

  // Tabla de Excel: rango, encabezado, totales, columnas (con su fórmula calculada), filtro y estilo.
  function readTable(el, dxfs, warnings) {
    const t = { id: +O.attr(el, 'id'), name: O.attr(el, 'displayName') || O.attr(el, 'name'), ref: O.attr(el, 'ref') };
    if (O.attr(el, 'headerRowCount') === '0') t.header = false;
    if (+O.attr(el, 'totalsRowCount')) t.totals = +O.attr(el, 'totalsRowCount');
    if (O.attr(el, 'totalsRowShown') === '0') t.totalsShown = false;
    if (O.attr(el, 'tableType') && O.attr(el, 'tableType') !== 'worksheet') warnings.add('conexiones de tablas con datos externos');
    t.columns = O.kids(O.kid(el, 'tableColumns'), 'tableColumn').map((c) => {
      const col = { id: +O.attr(c, 'id'), name: O.attr(c, 'name') || '' };
      if (O.attr(c, 'totalsRowFunction')) col.fn = O.attr(c, 'totalsRowFunction');
      if (O.attr(c, 'totalsRowLabel') !== null) col.label = O.attr(c, 'totalsRowLabel');
      const calc = O.kid(c, 'calculatedColumnFormula');
      if (calc) col.calc = calc.textContent.replace(/_xl(fn|ws)\./g, '');
      const tf = O.kid(c, 'totalsRowFormula');
      if (tf) col.totalF = tf.textContent.replace(/_xl(fn|ws)\./g, '');
      return col;
    });
    const af = O.kid(el, 'autoFilter');
    if (af) t.autoFilter = readAutoFilter(af, dxfs);
    const st = O.kid(el, 'tableStyleInfo');
    if (st) {
      t.style = { name: O.attr(st, 'name') || '' };
      for (const [k, a] of [['first', 'showFirstColumn'], ['last', 'showLastColumn'], ['rows', 'showRowStripes'], ['cols', 'showColumnStripes']]) if (isTrue(O.attr(st, a))) t.style[k] = true;
    }
    return t;
  }

  // Controles de formulario: tipo, posición, celda vinculada y estado, leídos de sus propiedades.
  async function readControls(root, sheetRels, ctx) {
    const out = [], seen = new Set();
    for (const el of O.all(root, 'control')) {
      const rel = sheetRels[O.attr(el, 'id')];
      if (rel && rel.type.endsWith('/control') && !seen.has(O.attr(el, 'shapeId'))) {
        const ctl = await readActiveX(el, rel, ctx);
        if (ctl) { seen.add(ctl.id); out.push(ctl); }
        continue;
      }
      if (!rel || !rel.type.endsWith('/ctrlProp')) continue;
      const prop = await O.readXml(ctx.zip, rel.target);
      const p = prop && prop.documentElement;
      if (!p) continue;
      const anchor = O.all(el, 'anchor')[0];
      const ctl = { id: O.attr(el, 'shapeId'), name: O.attr(el, 'name') || '', type: O.attr(p, 'objectType') || '', prop: await grabPart(ctx, rel.target) };
      if (anchor) { ctl.from = anchorPos(O.kid(anchor, 'from')); ctl.to = anchorPos(O.kid(anchor, 'to')); }
      if (O.attr(p, 'checked') === 'Checked') ctl.checked = true;
      for (const [k, a] of [['link', 'fmlaLink'], ['range', 'fmlaRange'], ['group', 'fmlaGroup']]) if (O.attr(p, a)) ctl[k] = O.attr(p, a);
      for (const [k, a] of [['sel', 'sel'], ['val', 'val'], ['min', 'min'], ['max', 'max'], ['inc', 'inc'], ['page', 'page']]) if (O.attr(p, a) !== null) ctl[k] = +O.attr(p, a);
      if (O.attr(p, 'firstButton') === '1') ctl.first = true;
      out.push(ctl);
    }
    return out;
  }

  // Controles ActiveX de Microsoft Forms. Su celda vinculada y su lista van en la hoja; el texto y el valor,
  // en el control: Oficina los lee cuando van en XML (persistPropertyBag). Si van en binario, muestra su nombre.
  const AX_TYPES = {
    '{8BD21D40-EC42-11CE-9E0D-00AA006002A5}': 'CheckBox', '{8BD21D50-EC42-11CE-9E0D-00AA006002A5}': 'Radio',
    '{8BD21D60-EC42-11CE-9E0D-00AA006002A5}': 'Toggle', '{8BD21D10-EC42-11CE-9E0D-00AA006002A5}': 'Edit',
    '{8BD21D30-EC42-11CE-9E0D-00AA006002A5}': 'Drop', '{8BD21D20-EC42-11CE-9E0D-00AA006002A5}': 'List',
    '{D7053240-CE69-11CD-A777-00DD01143C57}': 'Button', '{978C9E23-D4B0-11CE-BF2D-00AA003F40D0}': 'Label',
    '{DFD181E0-5E2F-11CE-A449-00AA004A803D}': 'Scroll', '{79176FB0-B7F2-11CE-97EF-00AA006D2776}': 'Spin',
    '{4C599241-6926-101B-9992-00000B65C6F9}': 'Image',
  };
  async function readActiveX(el, rel, ctx) {
    const doc = await O.readXml(ctx.zip, rel.target);
    const ocx = doc && doc.documentElement;
    if (!ocx) return null;
    const pr = O.all(el, 'controlPr')[0];
    const anchor = pr && O.kid(pr, 'anchor');
    const props = {};
    for (const p of O.all(ocx, 'ocxPr')) props[O.attr(p, 'name')] = O.attr(p, 'value');
    const ctl = {
      id: O.attr(el, 'shapeId'), name: O.attr(el, 'name') || '', activeX: true,
      type: AX_TYPES[String(O.attr(ocx, 'classid') || '').toUpperCase()] || 'ActiveX',
      prop: await grabPart(ctx, rel.target), bag: O.attr(ocx, 'persistence') === 'persistPropertyBag',
    };
    if (anchor) { ctl.from = anchorPos(O.kid(anchor, 'from')); ctl.to = anchorPos(O.kid(anchor, 'to')); }
    if (O.attr(pr, 'linkedCell')) ctl.link = O.attr(pr, 'linkedCell');
    if (O.attr(pr, 'listFillRange')) ctl.range = O.attr(pr, 'listFillRange');
    ctl.label = props.Caption !== undefined ? props.Caption : ctl.name;
    if (props.Value !== undefined) ctl.value = props.Value;
    if (props.GroupName) ctl.group = props.GroupName;
    for (const [k, a] of [['min', 'Min'], ['max', 'Max'], ['inc', 'SmallChange']]) if (props[a] !== undefined && !isNaN(+props[a])) ctl[k] = +props[a];
    return ctl;
  }

  // Autofiltro: valores elegidos, filtros personalizados, diez mejores y colores. Lo demás se guarda tal cual.
  function readAutoFilter(af, dxfs) {
    const out = { ref: O.attr(af, 'ref'), cols: {} };
    for (const fc of O.kids(af, 'filterColumn')) {
      const id = +O.attr(fc, 'colId');
      const filters = O.kid(fc, 'filters'), custom = O.kid(fc, 'customFilters'), top = O.kid(fc, 'top10'), color = O.kid(fc, 'colorFilter');
      let spec;
      if (filters && !O.kid(filters, 'dateGroupItem')) spec = { vals: O.kids(filters, 'filter').map((x) => O.attr(x, 'val')), ...(isTrue(O.attr(filters, 'blank')) ? { blank: true } : {}) };
      else if (custom) spec = { custom: { and: isTrue(O.attr(custom, 'and')), items: O.kids(custom, 'customFilter').map((x) => ({ op: O.attr(x, 'operator') || 'equal', val: O.attr(x, 'val') })) } };
      else if (top) spec = { top: { n: +O.attr(top, 'val') || 10, percent: isTrue(O.attr(top, 'percent')), bottom: O.attr(top, 'top') === '0' || O.attr(top, 'top') === 'false' } };
      else if (color && dxfs[+O.attr(color, 'dxfId')]) {
        const d = dxfs[+O.attr(color, 'dxfId')];
        spec = O.attr(color, 'cellColor') === '0' || O.attr(color, 'cellColor') === 'false' ? { font: d.color || '000000' } : { fill: d.fill || null };
      } else spec = { raw: serialize(fc) };
      if (O.attr(fc, 'hiddenButton') === '1') spec.noButton = true;
      out.cols[id] = spec;
    }
    return out;
  }

  // Dibujos de la hoja: gráficos, imágenes y formas sencillas, con su posición en celdas.
  async function readDrawing(part, ctx) {
    const { zip, warnings } = ctx;
    const xml = await O.readXml(zip, part);
    if (!xml) return [];
    const rels = await O.readRels(zip, part);
    const pos = anchorPos;
    const out = [];
    for (const top of O.kids(xml.documentElement)) {
      // Lo que va dentro de mc:AlternateContent (controles de formulario, objetos de Excel 2010) se guarda tal cual.
      if (top.localName === 'AlternateContent') {
        const anchor = O.all(top, 'twoCellAnchor')[0] || O.all(top, 'oneCellAnchor')[0];
        const item = { kind: 'raw', outer: true, ...(await keepFragment(ctx, top, rels)) };
        if (anchor) { item.from = pos(O.kid(anchor, 'from')); if (O.kid(anchor, 'to')) item.to = pos(O.kid(anchor, 'to')); }
        const spid = O.all(top, 'compatExt')[0];
        if (spid) item.spid = (O.attr(spid, 'spid') || '').replace(/^_x0000_s/, '');
        const cNvPr = O.all(top, 'cNvPr')[0];
        if (cNvPr && O.attr(cNvPr, 'name')) item.name = O.attr(cNvPr, 'name');
        if (!(cNvPr && isTrue(O.attr(cNvPr, 'hidden'))) && !spid) item.box = true;
        out.push(item);
        continue;
      }
      const anchor = top;
      if (!/Anchor$/.test(anchor.localName)) continue;
      if (anchor.localName === 'absoluteAnchor') { warnings.add('algunos objetos dibujados'); continue; }
      const item = { from: pos(O.kid(anchor, 'from')) };
      if (O.kid(anchor, 'to')) item.to = pos(O.kid(anchor, 'to'));
      else {
        const ext = O.kid(anchor, 'ext');
        item.size = { w: Math.round(+O.attr(ext, 'cx') / EMU), h: Math.round(+O.attr(ext, 'cy') / EMU) };
      }
      if (O.attr(anchor, 'editAs')) item.editAs = O.attr(anchor, 'editAs');
      const body = O.kids(anchor).find((k) => !['from', 'to', 'ext', 'clientData'].includes(k.localName));
      if (!body) continue;
      const cNvPr = O.all(body, 'cNvPr')[0];
      if (cNvPr && O.attr(cNvPr, 'name')) item.name = O.attr(cNvPr, 'name');
      if (body.localName === 'AlternateContent') {
        // Segmentaciones y escalas de tiempo: Excel las dibuja; aquí se guardan tal cual y se muestran.
        const sl = O.all(body, 'slicer')[0] || O.all(body, 'timeslicer')[0];
        Object.assign(item, { kind: sl ? 'slicer' : 'raw', ...(await keepFragment(ctx, body, rels)) });
        if (sl) item.slicer = { name: O.attr(sl, 'name'), timeline: sl.localName === 'timeslicer' };
        else item.box = true;
        out.push(item);
        continue;
      }
      if (body.localName === 'graphicFrame' && O.all(body, 'chart').length) {
        const rel = rels[O.attr(O.all(body, 'chart')[0], 'id')];
        const doc = rel && (await O.readXml(zip, rel.target));
        if (!doc) { warnings.add('algunos gráficos'); continue; }
        // Lo que el gráfico enlaza fuera (formas dentro del gráfico, datos incrustados) no se conserva.
        for (const tag of ['externalData', 'userShapes']) for (const el of O.all(doc.documentElement, tag)) el.parentNode.removeChild(el);
        item.kind = 'chart';
        item.xml = serialize(doc.documentElement);
        item.refs = O.chartRefs(item.xml);
      } else if (body.localName === 'pic') {
        const blip = O.all(body, 'blip')[0];
        const rel = blip && rels[O.attr(blip, 'embed')];
        if (!rel || rel.external) { warnings.add('algunas imágenes'); continue; }
        let key = ctx.mediaKey.get(rel.target);
        if (!key) {
          const url = await O.partToDataUrl(zip, rel.target);
          if (!url) { warnings.add('algunas imágenes'); continue; }
          key = 'm' + (ctx.mediaKey.size + 1);
          ctx.mediaKey.set(rel.target, key);
          ctx.media[key] = url;
        }
        item.kind = 'img';
        item.media = key;
      } else if (['sp', 'cxnSp', 'grpSp'].includes(body.localName)) {
        const raw = serialize(body);
        if (/\br:(embed|id|link)=/.test(raw)) { warnings.add('algunas formas'); continue; }
        item.kind = 'shape';
        item.xml = raw;
        item.text = O.all(body, 'p').map((p) => O.all(p, 't').map((t) => t.textContent).join('')).join('\n').trim();
        const sp = O.all(body, 'spPr')[0];
        item.fill = O.kid(sp, 'noFill') ? null : O.drawingColor(O.kid(sp, 'solidFill'), ctx.theme) || (O.all(body, 'fillRef')[0] ? O.drawingColor(O.all(body, 'fillRef')[0], ctx.theme) : null);
        const ln = O.kid(sp, 'ln');
        item.line = ln && O.kid(ln, 'noFill') ? null : O.drawingColor(O.kid(ln, 'solidFill'), ctx.theme) || (O.all(body, 'lnRef')[0] ? O.drawingColor(O.all(body, 'lnRef')[0], ctx.theme) : null);
        item.geom = O.attr(O.all(body, 'prstGeom')[0], 'prst') || 'rect';
      } else { warnings.add('algunos objetos dibujados'); continue; }
      out.push(item);
    }
    return out;
  }

  async function writeXlsx(state) {
    const book = new O.Workbook(state);
    const theme = state.themeColors || THEME;
    const fonts = [{}], fills = [null, 'gray'], borders = [''], numFmts = [], xfs = [{ font: 0, fill: 0, num: 0, border: 0 }];
    const keys = new Map([['', 0]]);
    const idOf = (list, val) => {
      const k = JSON.stringify(val);
      let i = list.findIndex((x) => JSON.stringify(x) === k);
      if (i < 0) i = list.push(val) - 1;
      return i;
    };
    const styleIndex = (s) => {
      if (!s) return 0;
      const key = JSON.stringify([s.b, s.i, s.color, s.fill, s.nf, s.align, s.u, s.sz, s.valign, s.wrap, s.bd, s.bdx, s.rot, s.font, s.strike, s.thColor, s.thFill, s.unlocked, s.hideF]);
      if (keys.has(key)) return keys.get(key);
      const font = idOf(fonts, { b: !!s.b, i: !!s.i, u: !!s.u, strike: !!s.strike, sz: s.sz || 11, color: s.color || null, th: s.color && s.thColor ? s.thColor : null, name: s.font || null });
      const border = s.bd ? idOf(borders, { bd: s.bd, bdx: s.bdx || {} }) : 0;
      const fill = s.fill ? idOf(fills, { hex: s.fill, th: s.thFill || null }) : 0;
      let num = 0;
      if (s.nf && s.nf !== 'General') {
        let j = numFmts.indexOf(s.nf);
        if (j < 0) j = numFmts.push(s.nf) - 1;
        num = 164 + j;
      }
      const idx = xfs.push({ font, fill, num, border, align: s.align, valign: s.valign, wrap: s.wrap, rot: s.rot, unlocked: s.unlocked, hideF: s.hideF }) - 1;
      keys.set(key, idx);
      return idx;
    };

    const sheetXml = (sheet) => {
      const headers = {};
      const tableNums = (sheet.tables || []).map((t) => tableXml(t, sheet, headers));
      const rows = new Map();
      for (const ref of new Set([...Object.keys(sheet.cells), ...Object.keys(headers)])) {
        const a = O.parseAddr(ref);
        if (!rows.has(a.r)) rows.set(a.r, []);
        rows.get(a.r).push(a);
      }
      // Una fila oculta o con alto propio se escribe aunque no tenga celdas.
      for (const r of [...Object.keys(sheet.hiddenRows || {}), ...Object.keys(sheet.rowH || {})]) {
        if (!rows.has(+r)) rows.set(+r, []);
      }
      let data = '';
      for (const r of [...rows.keys()].sort((x, y) => x - y)) {
        const ht = sheet.rowH && sheet.rowH[r];
        data += `<row r="${r + 1}"${ht ? ` ht="${(ht * 0.75).toFixed(2)}" customHeight="1"` : ''}${sheet.hiddenRows && sheet.hiddenRows[r] ? ' hidden="1"' : ''}>`;
        for (const a of rows.get(r).sort((x, y) => x.c - y.c)) {
          const ref = O.addr(a.r, a.c);
          const cell = sheet.cells[ref] || {};
          const s = styleIndex(cell.s);
          const sAttr = s ? ` s="${s}"` : '';
          // El encabezado de una tabla siempre es texto con el nombre de su columna.
          if (headers[ref] !== undefined) data += `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${O.esc(headers[ref])}</t></is></c>`;
          else if (cell.f !== undefined) {
            const val = book.value(sheet, a.r, a.c);
            const f = `<f>${O.esc(fileFormula(cell.f))}</f>`;
            if (val && val.err) data += `<c r="${ref}"${sAttr} t="e">${f}<v>${O.esc(val.err)}</v></c>`;
            else if (typeof val === 'string') data += `<c r="${ref}"${sAttr} t="str">${f}<v>${O.esc(val)}</v></c>`;
            else if (typeof val === 'boolean') data += `<c r="${ref}"${sAttr} t="b">${f}<v>${val ? 1 : 0}</v></c>`;
            else data += `<c r="${ref}"${sAttr}>${f}<v>${val === null ? 0 : val}</v></c>`;
          } else if (typeof cell.v === 'number') data += `<c r="${ref}"${sAttr}><v>${cell.v}</v></c>`;
          else if (typeof cell.v === 'boolean') data += `<c r="${ref}"${sAttr} t="b"><v>${cell.v ? 1 : 0}</v></c>`;
          else if (cell.v && cell.v.err) data += `<c r="${ref}"${sAttr} t="e"><v>${O.esc(cell.v.err)}</v></c>`;
          else if (cell.v !== undefined && cell.v !== null && cell.v !== '') data += `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${O.esc(cell.v)}</t></is></c>`;
          else data += `<c r="${ref}"${sAttr}/>`;
        }
        data += '</row>';
      }
      const hiddenCols = sheet.hiddenCols || {};
      const colKeys = [...new Set([...Object.keys(sheet.cols || {}), ...Object.keys(hiddenCols)].map(Number))].sort((x, y) => x - y);
      const cols = colKeys.length
        ? `<cols>${colKeys.map((c) => `<col min="${c + 1}" max="${c + 1}" width="${Math.max(1, (((sheet.cols || {})[c] || sheet.defW || 64) - 5) / 7).toFixed(2)}" customWidth="1"${hiddenCols[c] ? ' hidden="1"' : ''}/>`).join('')}</cols>`
        : '';
      const refs = Object.keys(sheet.cells);
      const links = refs.filter((k) => sheet.cells[k].link);
      const notes = refs.filter((k) => sheet.cells[k].note);
      const rels = links.map((k, n) => `<Relationship Id="rIdL${n + 1}" Type="${R}/hyperlink" Target="${O.esc(sheet.cells[k].link)}" TargetMode="External"/>`);
      if (notes.length) {
        rels.push(`<Relationship Id="rIdC" Type="${R}/comments" Target="../comments${index + 1}.xml"/>`);
        zip.file(`xl/comments${index + 1}.xml`, `${O.XML}<comments xmlns="${MAIN}"><authors><author>Oficina</author></authors><commentList>${notes.map((k) => `<comment ref="${k}" authorId="0"><text><r><t xml:space="preserve">${O.esc(sheet.cells[k].note)}</t></r></text></comment>`).join('')}</commentList></comments>`);
      }
      const extraRels = [];
      const relate = relater(extraRels, 'rIdX');
      if (notes.length || sheet.vml) {
        rels.push(`<Relationship Id="rIdV" Type="${R}/vmlDrawing" Target="../drawings/vmlDrawing${index + 1}.vml"/>`);
        const vmlRels = [];
        zip.file(`xl/drawings/vmlDrawing${index + 1}.vml`, vmlXml(sheet, notes, relater(vmlRels, 'rId')));
        if (vmlRels.length) zip.file(`xl/drawings/_rels/vmlDrawing${index + 1}.vml.rels`, relsDoc(vmlRels.join('')));
      }
      patchControls(sheet);
      const ctrl = (sheet.ctrlXml || []).map((f) => patchCtrlXml(sheet, fragmentXml(f, relate))).join('');
      tableNums.forEach((n) => rels.push(`<Relationship Id="rIdT${n}" Type="${R}/table" Target="../tables/table${n}.xml"/>`));
      const tableParts = tableNums.length ? `<tableParts count="${tableNums.length}">${tableNums.map((n) => `<tablePart r:id="rIdT${n}"/>`).join('')}</tableParts>` : '';
      const exts = (sheet.exts || []).map((f) => fragmentXml(f, relate)).join('');
      const drawing = drawingXml(sheet, index);
      if (drawing) rels.push(`<Relationship Id="rIdD" Type="${R}/drawing" Target="../drawings/drawing${index + 1}.xml"/>`);
      (sheet.pivots || []).forEach((p) => {
        const n = ++pivotCount;
        const cache = cacheIndex.get(p.cache);
        if (cache === undefined) return;
        zip.file(`xl/pivotTables/pivotTable${n}.xml`, p.xml);
        zip.file(`xl/pivotTables/_rels/pivotTable${n}.xml.rels`, `${O.XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${O.esc(p.relId)}" Type="${R}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${cache + 1}.xml"/></Relationships>`);
        rels.push(`<Relationship Id="rIdP${n}" Type="${R}/pivotTable" Target="../pivotTables/pivotTable${n}.xml"/>`);
        types.push(`<Override PartName="/xl/pivotTables/pivotTable${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>`);
      });
      rels.push(...extraRels);
      if (rels.length) zip.file(`xl/worksheets/_rels/sheet${index + 1}.xml.rels`, relsDoc(rels.join('')));
      const tabXml = sheet.tabColor ? colorXml('tabColor', sheet.tabColor, sheet.thTab, theme) : '';
      const sheetPr = sheet.tabColor || sheet.codeName ? `<sheetPr${sheet.codeName ? ` codeName="${O.esc(sheet.codeName)}"` : ''}>${tabXml}</sheetPr>` : '';
      const protection = sheet.protected ? protectionXml(sheet.protect) : '';
      const hyperlinks = links.length ? `<hyperlinks>${links.map((k, n) => `<hyperlink ref="${k}" r:id="rIdL${n + 1}"/>`).join('')}</hyperlinks>` : '';
      const legacy = notes.length || sheet.vml ? '<legacyDrawing r:id="rIdV"/>' : '';
      const formatPr = sheet.defW || sheet.defH
        ? `<sheetFormatPr${sheet.defW ? ` defaultColWidth="${((sheet.defW - 5) / 7).toFixed(2)}"` : ''} defaultRowHeight="${((sheet.defH || 20) * 0.75).toFixed(2)}"${sheet.defH ? ' customHeight="1"' : ''}/>` : '';
      // Los trozos guardados tal cual usan estos prefijos sin declararlos (Requires="x14").
      const ns = ctrl || exts ? ` xmlns:mc="${MC}" xmlns:x14="${X14}" xmlns:xdr="${XDR}"` : '';
      return `${O.XML}<worksheet xmlns="${MAIN}" xmlns:r="${R}"${ns}>${sheetPr}${formatPr}${cols}<sheetData>${data}</sheetData>${protection}${autoFilterXml(sheet)}${mergeXml(sheet)}${cfXml(sheet)}${dvXml(sheet)}${hyperlinks}${drawing ? '<drawing r:id="rIdD"/>' : ''}${legacy}${ctrl}${tableParts}${exts ? `<extLst>${exts}</extLst>` : ''}</worksheet>`;
    };
    // Orden de los atributos según el esquema; los que no se conocen se descartan.
    const PROT = ['password', 'algorithmName', 'hashValue', 'saltValue', 'spinCount', 'sheet', 'objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows',
      'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'];
    const protectionXml = (attrs) => {
      const a = { ...(attrs || {}), sheet: '1' };
      if (!attrs) Object.assign(a, { objects: '1', scenarios: '1' });
      return `<sheetProtection${PROT.filter((k) => a[k] !== undefined).map((k) => ` ${k}="${O.esc(a[k])}"`).join('')}/>`;
    };
    let pivotCount = 0;
    const cacheIndex = new Map((state.pivotCaches || []).map((p, i) => [p.id, i]));
    const types = [];
    const autoFilterXml = (sheet) => {
      const af = sheet.autoFilter;
      if (!af || !af.ref) return '';
      return `<autoFilter ref="${O.esc(af.ref)}">${filterColumnsXml(af)}</autoFilter>`;
    };
    const filterColumnsXml = (af) => Object.keys(af.cols || {}).map(Number).sort((a, b) => a - b).map((id) => {
        const spec = af.cols[id];
        const hb = spec.noButton ? ' hiddenButton="1"' : '';
        if (spec.raw) return spec.raw.replace(/\s+xmlns(:\w+)?="[^"]*"/g, '').replace(/colId="\d+"/, `colId="${id}"`);
        if (spec.vals) return `<filterColumn colId="${id}"${hb}><filters${spec.blank ? ' blank="1"' : ''}>${spec.vals.map((v) => `<filter val="${O.esc(v)}"/>`).join('')}</filters></filterColumn>`;
        if (spec.custom) return `<filterColumn colId="${id}"${hb}><customFilters${spec.custom.and ? ' and="1"' : ''}>${spec.custom.items.map((x) => `<customFilter${x.op && x.op !== 'equal' ? ` operator="${x.op}"` : ''} val="${O.esc(x.val)}"/>`).join('')}</customFilters></filterColumn>`;
        if (spec.top) return `<filterColumn colId="${id}"${hb}><top10${spec.top.bottom ? ' top="0"' : ''}${spec.top.percent ? ' percent="1"' : ''} val="${spec.top.n}"/></filterColumn>`;
        if (spec.fill !== undefined || spec.font !== undefined) {
          const style = spec.font !== undefined ? { color: spec.font } : { fill: spec.fill || 'FFFFFF' };
          return `<filterColumn colId="${id}"${hb}><colorFilter dxfId="${idOf(dxfs, style)}"${spec.font !== undefined ? ' cellColor="0"' : ''}/></filterColumn>`;
        }
        return hb ? `<filterColumn colId="${id}"${hb}/>` : '';
      }).join('');

    // Partes guardadas tal cual: se escriben una sola vez, con sus relaciones.
    const relsDoc = (inner) => `${O.XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
    const relsPath = (p) => p.replace(/([^/]+)$/, '_rels/$1.rels');
    const written = new Map();
    const patched = {};
    const putPart = (path) => {
      if (written.has(path)) return written.get(path);
      const p = (state.parts || {})[path];
      if (!p) return null;
      // Con un nombre propio si coincide con las partes que crea Oficina.
      const out = /^xl\/(worksheets|drawings|charts|media|pivotTables|pivotCache|tables|comments|theme|styles|workbook|vbaProject)/.test(path) ? path.replace(/([^/]+)$/, 'o_$1') : path;
      written.set(path, out);
      const data = patched[path] !== undefined ? patched[path] : p.data;
      if (data && data.b64 !== undefined) zip.file(out, data.b64, { base64: true });
      else zip.file(out, data);
      types.push(`<Override PartName="/${O.esc(out)}" ContentType="${O.esc(p.type)}"/>`);
      const rels = (p.rels || []).map((r) => {
        if (r.url) return `<Relationship Id="${O.esc(r.id)}" Type="${O.esc(r.type)}" Target="${O.esc(r.url)}" TargetMode="External"/>`;
        const target = r.part && putPart(r.part);
        return target ? `<Relationship Id="${O.esc(r.id)}" Type="${O.esc(r.type)}" Target="/${O.esc(target)}"/>` : '';
      }).join('');
      if (rels) zip.file(relsPath(out), relsDoc(rels));
      return out;
    };
    // Un trozo guardado tal cual, con sus enlaces renumerados en la parte que lo contiene.
    const fragmentXml = (frag, relate) => {
      const ids = {};
      for (const [id, l] of Object.entries(frag.links || {})) {
        const target = l.url || (l.part && putPart(l.part));
        if (target) ids[id] = relate(l.type, l.url ? target : '/' + target, !!l.url);
      }
      return frag.xml.replace(LINK, (m, pre, id) => (ids[id] ? `${pre}${ids[id]}"` : m));
    };
    const relater = (list, prefix) => (type, target, external) => {
      const id = `${prefix}${list.length + 1}`;
      list.push(`<Relationship Id="${id}" Type="${O.esc(type)}" Target="${O.esc(target)}"${external ? ' TargetMode="External"' : ''}/>`);
      return id;
    };

    // Tablas: los nombres de las columnas son los textos de su encabezado, como exige Excel.
    let tableCount = 0, maxTable = 0;
    for (const s of state.sheets) for (const t of s.tables || []) maxTable = Math.max(maxTable, t.id || 0);
    const tableIds = new Set();
    const tableXml = (t, sheet, headers) => {
      const b = O.tableBox(t), hdr = t.header !== false;
      let id = t.id;
      if (!id || tableIds.has(id)) id = ++maxTable;
      tableIds.add(id);
      let nextCol = Math.max(0, ...(t.columns || []).map((c) => c.id || 0));
      const cols = [];
      for (let k = 0; k <= b.c2 - b.c1; k++) {
        const col = { ...((t.columns || [])[k] || { id: ++nextCol }) };
        const base = (hdr ? book.display(sheet, b.r1, b.c1 + k) : '') || col.name || `Columna${k + 1}`;
        let name = base;
        for (let n = 2; cols.some((x) => x.name.toLowerCase() === name.toLowerCase()); n++) name = base + n;
        col.name = name;
        cols.push(col);
        if (hdr) headers[O.addr(b.r1, b.c1 + k)] = name;
      }
      const af = hdr && t.autoFilter ? `<autoFilter ref="${O.addr(b.r1, b.c1)}:${O.addr(Math.max(b.r1, b.r2 - (t.totals || 0)), b.c2)}">${filterColumnsXml(t.autoFilter)}</autoFilter>` : '';
      const colXml = (c) => {
        const inner = (c.calc ? `<calculatedColumnFormula>${O.esc(fileFormula(c.calc))}</calculatedColumnFormula>` : '') + (c.totalF ? `<totalsRowFormula>${O.esc(fileFormula(c.totalF))}</totalsRowFormula>` : '');
        return `<tableColumn id="${c.id}" name="${O.esc(c.name)}"${c.fn ? ` totalsRowFunction="${O.esc(c.fn)}"` : ''}${c.label !== undefined ? ` totalsRowLabel="${O.esc(c.label)}"` : ''}${inner ? `>${inner}</tableColumn>` : '/>'}`;
      };
      const st = t.style;
      const style = st ? `<tableStyleInfo name="${O.esc(st.name)}" showFirstColumn="${st.first ? 1 : 0}" showLastColumn="${st.last ? 1 : 0}" showRowStripes="${st.rows ? 1 : 0}" showColumnStripes="${st.cols ? 1 : 0}"/>` : '';
      const n = ++tableCount;
      zip.file(`xl/tables/table${n}.xml`, `${O.XML}<table xmlns="${MAIN}" id="${id}" name="${O.esc(t.name)}" displayName="${O.esc(t.name)}" ref="${O.esc(t.ref)}"${hdr ? '' : ' headerRowCount="0"'}${t.totals ? ` totalsRowCount="${t.totals}"` : t.totalsShown === false ? ' totalsRowShown="0"' : ''}>${af}<tableColumns count="${cols.length}">${cols.map(colXml).join('')}</tableColumns>${style}</table>`);
      types.push(`<Override PartName="/xl/tables/table${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
      return n;
    };

    // Controles de formulario: su posición y su estado (marcado, elegido, valor) pueden haber cambiado.
    const setAttr = (xml, tag, name, v) => {
      let out = xml.replace(new RegExp(`(<(?:\\w+:)?${tag}\\b[^>]*?)\\s${name}="[^"]*"`), '$1');
      if (v !== null && v !== undefined) out = out.replace(new RegExp(`<(?:\\w+:)?${tag}\\b`), (m) => `${m} ${name}="${v}"`);
      return out;
    };
    const ctrlState = (ctl) => ({
      checked: /^(CheckBox|Radio)$/.test(ctl.type) ? !!ctl.checked : undefined,
      sel: /^(Drop|List)$/.test(ctl.type) ? ctl.sel : undefined,
      val: /^(Spin|Scroll)$/.test(ctl.type) ? ctl.val : undefined,
    });
    const patchControls = (sheet) => {
      for (const ctl of sheet.controls || []) {
        const p = ctl.prop && (state.parts || {})[ctl.prop];
        if (!p || typeof p.data !== 'string') continue;
        if (ctl.activeX) {
          // El valor de un ActiveX solo se puede cambiar si sus propiedades van en XML.
          if (!ctl.bag || ctl.value === undefined) continue;
          const v = O.esc(String(ctl.value));
          const m = /<((?:\w+:)?)ocx\b/.exec(p.data), px = m ? m[1] : '';
          const re = new RegExp(`<(?:\\w+:)?ocxPr\\b[^>]*\\bname="Value"[^>]*/>`);
          const tag = `<${px}ocxPr ${px}name="Value" ${px}value="${v}"/>`;
          patched[ctl.prop] = re.test(p.data) ? p.data.replace(re, tag) : p.data.replace(new RegExp(`</${px}ocx>`), `${tag}</${px}ocx>`);
          continue;
        }
        const st = ctrlState(ctl);
        let d = p.data;
        if (st.checked !== undefined) d = setAttr(d, 'formControlPr', 'checked', st.checked ? 'Checked' : null);
        if (st.sel !== undefined) d = setAttr(d, 'formControlPr', 'sel', st.sel);
        if (st.val !== undefined) d = setAttr(d, 'formControlPr', 'val', st.val);
        patched[ctl.prop] = d;
      }
    };
    const vmlAnchor = (f, t) => [f.c, f.dx || 0, f.r, f.dy || 0, t.c, t.dx || 0, t.r, t.dy || 0].join(', ');
    const patchVml = (sheet, xml) => {
      for (const ctl of sheet.controls || []) {
        xml = xml.replace(new RegExp(`<v:shape\\b[^>]*\\b(?:id|o:spid)="_x0000_s${ctl.id}"[\\s\\S]*?<\\/v:shape>`), (s) => {
          if (ctl.from && ctl.to) s = s.replace(/<x:Anchor>[\s\S]*?<\/x:Anchor>/, `<x:Anchor>${vmlAnchor(ctl.from, ctl.to)}</x:Anchor>`);
          if (ctl.activeX) return s;
          const st = ctrlState(ctl);
          const tag = (name, v) => {
            s = s.replace(new RegExp(`\\s*<x:${name}>[^<]*</x:${name}>`), '');
            if (v !== null && v !== undefined) s = s.replace('</x:ClientData>', `<x:${name}>${v}</x:${name}></x:ClientData>`);
          };
          if (st.checked !== undefined) tag('Checked', st.checked ? 1 : null);
          if (st.sel !== undefined) tag('Sel', st.sel);
          if (st.val !== undefined) tag('Val', st.val);
          return s;
        });
      }
      return xml;
    };
    const patchCtrlXml = (sheet, xml) => {
      for (const ctl of sheet.controls || []) {
        if (!ctl.from || !ctl.to) continue;
        xml = xml.replace(new RegExp(`(<(?:\\w+:)?control\\b[^>]*\\bshapeId="${ctl.id}"[\\s\\S]*?<(?:\\w+:)?anchor\\b[^>]*>)[\\s\\S]*?(</(?:\\w+:)?anchor>)`),
          (m, a, b) => `${a}<from>${anchorXml(ctl.from)}</from><to>${anchorXml(ctl.to)}</to>${b}`);
      }
      return xml;
    };
    // Bloques de números de forma del dibujo heredado: los de los controles se mantienen.
    const vmlBlocks = new Set();
    for (const s of state.sheets) for (const b of String((s.vml && s.vml.data) || '').split(',')) if (+b) vmlBlocks.add(+b);
    let nextBlock = 1;
    const blockFor = (sheet) => {
      if (sheet.vml && sheet.vml.data) return sheet.vml.data;
      while (vmlBlocks.has(nextBlock)) nextBlock++;
      vmlBlocks.add(nextBlock);
      return String(nextBlock);
    };
    // Dibujos: gráficos, imágenes y formas, cada uno anclado a sus celdas.
    let chartCount = 0, imageCount = 0;
    const mediaPart = new Map();
    const exts = new Set();
    const anchorXml = (p) => `<xdr:col>${p.c}</xdr:col><xdr:colOff>${Math.round((p.dx || 0) * EMU)}</xdr:colOff><xdr:row>${p.r}</xdr:row><xdr:rowOff>${Math.round((p.dy || 0) * EMU)}</xdr:rowOff>`;
    const drawingXml = (sheet, index) => {
      const items = (sheet.drawings || []).filter((d) => d.from && (d.kind !== 'img' || (state.media || {})[d.media]));
      if (!items.length) return false;
      const rels = [];
      const relate = relater(rels, 'rIdX');
      // Los números de los objetos guardados tal cual se respetan; los demás van detrás.
      const kept = items.filter((d) => d.kind === 'raw' || d.kind === 'slicer');
      const base = Math.max(0, ...kept.flatMap((d) => [...d.xml.matchAll(/cNvPr\b[^>]*?\bid="(\d+)"/g)].map((m) => +m[1])));
      const body = items.map((d, n) => {
        const id = base + n + 2;
        const name = O.esc(d.name || (d.kind === 'chart' ? `Gráfico ${n + 1}` : d.kind === 'img' ? `Imagen ${n + 1}` : `Forma ${n + 1}`));
        let inner;
        if (d.kind === 'raw' || d.kind === 'slicer') {
          // Un control de formulario se mueve con su control.
          const ctl = d.spid && (sheet.controls || []).find((x) => x.id === d.spid);
          const from = ctl && ctl.from ? ctl.from : d.from, to = ctl && ctl.to ? ctl.to : d.to;
          let xml = fragmentXml(d, relate);
          if (d.outer) {
            xml = xml.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, `<xdr:from>${anchorXml(from)}</xdr:from>`);
            if (to) xml = xml.replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, `<xdr:to>${anchorXml(to)}</xdr:to>`);
            return xml;
          }
          inner = xml;
        } else if (d.kind === 'chart') {
          const k = ++chartCount;
          zip.file(`xl/charts/chart${k}.xml`, O.chartWithRefs(d.xml.startsWith('<?xml') ? d.xml : O.XML + d.xml, d.refs || []));
          types.push(`<Override PartName="/xl/charts/chart${k}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
          rels.push(`<Relationship Id="rId${n + 1}" Type="${R}/chart" Target="../charts/chart${k}.xml"/>`);
          inner = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId${n + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
        } else if (d.kind === 'img') {
          let target = mediaPart.get(d.media);
          if (!target) {
            const img = O.splitDataUrl(state.media[d.media]);
            target = `image${++imageCount}.${img.ext}`;
            zip.file(`xl/media/${target}`, img.b64, { base64: true });
            exts.add(img.ext);
            mediaPart.set(d.media, target);
          }
          rels.push(`<Relationship Id="rId${n + 1}" Type="${R}/image" Target="../media/${target}"/>`);
          const w = d.size ? d.size.w : 100, h = d.size ? d.size.h : 100;
          inner = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId${n + 1}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * EMU}" cy="${h * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
        } else inner = d.xml.replace(/(<(?:\w+:)?cNvPr\b[^>]*?\bid=")\d+"/, `$1${id}"`);
        const edit = d.editAs ? ` editAs="${d.editAs}"` : '';
        if (d.to) return `<xdr:twoCellAnchor${edit}><xdr:from>${anchorXml(d.from)}</xdr:from><xdr:to>${anchorXml(d.to)}</xdr:to>${inner}<xdr:clientData/></xdr:twoCellAnchor>`;
        return `<xdr:oneCellAnchor><xdr:from>${anchorXml(d.from)}</xdr:from><xdr:ext cx="${(d.size ? d.size.w : 100) * EMU}" cy="${(d.size ? d.size.h : 100) * EMU}"/>${inner}<xdr:clientData/></xdr:oneCellAnchor>`;
      }).join('');
      const ns = kept.length ? ` xmlns:mc="${MC}" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"` : '';
      zip.file(`xl/drawings/drawing${index + 1}.xml`, `${O.XML}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}"${ns}>${body}</xdr:wsDr>`);
      if (rels.length) zip.file(`xl/drawings/_rels/drawing${index + 1}.xml.rels`, `${O.XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
      types.push(`<Override PartName="/xl/drawings/drawing${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
      return true;
    };
    const mergeXml = (sheet) => (sheet.merges && sheet.merges.length
      ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${O.esc(m)}"/>`).join('')}</mergeCells>` : '');
    const dxfs = [];
    let priority = 0;
    const cfvoXml = (p) => `<cfvo type="${O.esc(p.type)}"${p.val !== null && p.val !== undefined && p.val !== '' ? ` val="${O.esc(String(p.val))}"` : ''}/>`;
    const cfXml = (sheet) => (sheet.cf || []).map((rule) => {
      const ranges = O.parseRanges(rule.ref);
      if (!ranges.length) return '';
      let attrs = ` type="${rule.type}" priority="${++priority}"`;
      if (rule.style && Object.keys(rule.style).length) attrs += ` dxfId="${idOf(dxfs, rule.style)}"`;
      if (rule.stop) attrs += ' stopIfTrue="1"';
      if (rule.type === 'aboveAverage') attrs += `${rule.above === false ? ' aboveAverage="0"' : ''}${rule.equal ? ' equalAverage="1"' : ''}${rule.std ? ` stdDev="${rule.std}"` : ''}`;
      if (rule.type === 'top10') attrs += `${rule.percent ? ' percent="1"' : ''}${rule.bottom ? ' bottom="1"' : ''} rank="${rule.rank || 10}"`;
      if (rule.op && rule.type === 'cellIs') attrs += ` operator="${rule.op}"`;
      if (/Text|With$/.test(rule.type)) attrs += ` operator="${{ containsText: 'containsText', notContainsText: 'notContains', beginsWith: 'beginsWith', endsWith: 'endsWith' }[rule.type]}" text="${O.esc(rule.text || '')}"`;
      if (rule.type === 'timePeriod') attrs += ` timePeriod="${rule.period}"`;
      const generated = ruleFormula(rule, O.addr(ranges[0].r1, ranges[0].c1));
      const formulas = generated ? [generated] : rule.f || [];
      let inner = formulas.map((f) => `<formula>${O.esc(fileFormula(f))}</formula>`).join('');
      if (rule.type === 'colorScale') inner = `<colorScale>${rule.scale.map(cfvoXml).join('')}${rule.scale.map((s) => `<color rgb="FF${s.color}"/>`).join('')}</colorScale>`;
      if (rule.type === 'dataBar') {
        const bar = rule.bar || {};
        inner = `<dataBar>${cfvoXml(bar.min || { type: 'min' })}${cfvoXml(bar.max || { type: 'max' })}<color rgb="FF${bar.color || '638EC6'}"/></dataBar>`;
      }
      if (rule.type === 'iconSet') {
        const set = rule.icons || {};
        const name = set.set || '3TrafficLights1';
        const n = +name[0] || 3;
        const pts = set.points && set.points.length === n ? set.points : O.iconDefaults(n);
        inner = `<iconSet${name !== '3TrafficLights1' ? ` iconSet="${O.esc(name)}"` : ''}${set.showValue === false ? ' showValue="0"' : ''}${set.reverse ? ' reverse="1"' : ''}>`
          + pts.map((p) => cfvoXml(p).replace('/>', p.gte === false ? ' gte="0"/>' : '/>')).join('') + '</iconSet>';
      }
      return `<conditionalFormatting sqref="${O.esc(O.rangesText(ranges))}"><cfRule${attrs}>${inner}</cfRule></conditionalFormatting>`;
    }).join('');
    const dvXml = (sheet) => (sheet.dv && sheet.dv.length
      ? `<dataValidations count="${sheet.dv.length}">${sheet.dv.map((d) => `<dataValidation${DV_ATTRS.filter((n) => d.attrs && d.attrs[n] !== undefined).map((n) => ` ${n}="${O.esc(String(d.attrs[n]))}"`).join('')} sqref="${O.esc(d.ref)}">${(d.f || []).map((f, i) => `<formula${i + 1}>${O.esc(fileFormula(f))}</formula${i + 1}>`).join('')}</dataValidation>`).join('')}</dataValidations>` : '');
    const dxfXml = (s) => {
      if (s.raw) return s.raw;
      const font = `${s.b ? '<b/>' : ''}${s.i ? '<i/>' : ''}${s.strike ? '<strike/>' : ''}${s.u ? '<u/>' : ''}${s.color ? `<color rgb="FF${s.color}"/>` : ''}`;
      return `<dxf>${font ? `<font>${font}</font>` : ''}${s.fill ? `<fill><patternFill patternType="solid"><fgColor rgb="FF${s.fill}"/><bgColor rgb="FF${s.fill}"/></patternFill></fill>` : ''}</dxf>`;
    };
    // Excel necesita este dibujo heredado para mostrar las notas.
    // Las formas de los controles se guardan como venían; las notas van detrás, con números libres.
    const vmlXml = (sheet, notes, relate) => {
      const data = blockFor(sheet);
      const carried = sheet.vml ? patchVml(sheet, fragmentXml(sheet.vml, relate)) : '';
      const used = [...carried.matchAll(/id="_x0000_s(\d+)"/g)].map((m) => +m[1]);
      const first = used.length ? Math.max(...used) + 1 : 1024 * (+String(data).split(',')[0] || 1) + 1;
      const noteType = notes.length && !/id="_x0000_t202"/.test(carried) ? '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' : '';
      return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${data}"/></o:shapelayout>${carried}${noteType}${notes.map((k, n) => {
      const a = O.parseAddr(k);
      return `<v:shape id="_x0000_s${first + n}" type="#_x0000_t202" style="position:absolute;margin-left:60pt;margin-top:2pt;width:130pt;height:70pt;z-index:${n + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox/><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${a.c + 1}, 15, ${a.r}, 2, ${a.c + 3}, 15, ${a.r + 4}, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${a.r}</x:Row><x:Column>${a.c}</x:Column></x:ClientData></v:shape>`;
    }).join('')}</xml>`;
    };

    const zip = new g.JSZip();
    let index = 0;
    const noteSheets = [];
    let hasVml = false;
    state.sheets.forEach((sheet, i) => {
      index = i;
      if (Object.values(sheet.cells).some((c) => c.note)) noteSheets.push(i);
      if (sheet.vml || Object.values(sheet.cells).some((c) => c.note)) hasVml = true;
      zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(sheet));
    });

    const fontXml = (f) => `<font>${f.b ? '<b/>' : ''}${f.i ? '<i/>' : ''}${f.strike ? '<strike/>' : ''}${f.u ? '<u/>' : ''}<sz val="${f.sz || 11}"/>${f.color ? colorXml('color', f.color, f.th, theme) : ''}<name val="${O.esc(f.name || 'Calibri')}"/></font>`;
    const borderXml = (b) => `<border>${[['left', 'l'], ['right', 'r'], ['top', 't'], ['bottom', 'b']].map(([side, k]) => {
      if (!b || !b.bd || !b.bd.includes(k)) return `<${side}/>`;
      const [style, color, th] = (b.bdx && b.bdx[k]) || ['thin', null];
      return `<${side} style="${O.esc(style)}">${color ? colorXml('color', color, th, theme) : '<color auto="1"/>'}</${side}>`;
    }).join('')}<diagonal/></border>`;
    const alignXml = (x) => (x.align || x.valign || x.wrap || x.rot ? `<alignment${x.align ? ` horizontal="${x.align}"` : ''}${x.valign ? ` vertical="${x.valign}"` : ''}${x.rot ? ` textRotation="${x.rot}"` : ''}${x.wrap ? ' wrapText="1"' : ''}/>` : '');
    const protXml = (x) => (x.unlocked || x.hideF ? `<protection${x.unlocked ? ' locked="0"' : ''}${x.hideF ? ' hidden="1"' : ''}/>` : '');
    const fillXml = (f, i) => (i === 0 ? '<fill><patternFill patternType="none"/></fill>' : i === 1 ? '<fill><patternFill patternType="gray125"/></fill>' : `<fill><patternFill patternType="solid">${colorXml('fgColor', f.hex, f.th, theme)}<bgColor indexed="64"/></patternFill></fill>`);
    const xfXml = (x) => {
      const inner = alignXml(x) + protXml(x);
      return `<xf numFmtId="${x.num}" fontId="${x.font}" fillId="${x.fill}" borderId="${x.border}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"${alignXml(x) ? ' applyAlignment="1"' : ''}${protXml(x) ? ' applyProtection="1"' : ''}${inner ? `>${inner}</xf>` : '/>'}`;
    };
    zip.file('xl/styles.xml', `${O.XML}<styleSheet xmlns="${MAIN}">${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.map((n, i) => `<numFmt numFmtId="${164 + i}" formatCode="${O.esc(n)}"/>`).join('')}</numFmts>` : ''}<fonts count="${fonts.length}">${fonts.map(fontXml).join('')}</fonts><fills count="${fills.length}">${fills.map(fillXml).join('')}</fills><borders count="${borders.length}">${borders.map(borderXml).join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.map(xfXml).join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>${(() => {
      // Los estilos de tabla personalizados van detrás de los dxf, que tienen que existir antes.
      const ts = state.tableStyles;
      const list = ts ? ts.list.map((st) => `<tableStyle name="${O.esc(st.name)}"${st.pivotOnly ? ' table="0"' : ''}${st.tableOnly ? ' pivot="0"' : ''} count="${st.parts.length}">${st.parts.map((p) => `<tableStyleElement type="${p.type}"${p.size ? ` size="${p.size}"` : ''} dxfId="${idOf(dxfs, { raw: p.raw })}"/>`).join('')}</tableStyle>`).join('') : '';
      const tables = ts ? `<tableStyles count="${ts.list.length}"${ts.def ? ` defaultTableStyle="${O.esc(ts.def)}"` : ''}${ts.pivot ? ` defaultPivotStyle="${O.esc(ts.pivot)}"` : ''}>${list}</tableStyles>` : '';
      return (dxfs.length ? `<dxfs count="${dxfs.length}">${dxfs.map(dxfXml).join('')}</dxfs>` : '') + tables;
    })()}</styleSheet>`);

    const nameKeys = Object.keys(state.names || {});
    // El rango de cada autofiltro va en un nombre oculto, como lo deja Excel.
    const filterNames = state.sheets.map((s, i) => (s.autoFilter && s.autoFilter.ref ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${O.esc(`'${s.name.replace(/'/g, "''")}'!${O.parseRanges(s.autoFilter.ref).map((x) => `$${O.colName(x.c1)}$${x.r1 + 1}:$${O.colName(x.c2)}$${x.r2 + 1}`)[0]}`)}</definedName>` : '')).join('');
    const definedNames = nameKeys.length || filterNames ? `<definedNames>${filterNames}${nameKeys.map((k) => `<definedName name="${O.esc(k)}">${O.esc(state.names[k])}</definedName>`).join('')}</definedNames>` : '';
    // La hoja activa debe ser visible: Excel no abre bien un libro con la activa oculta.
    const activeTab = state.sheets[state.active || 0] && !state.sheets[state.active || 0].hidden ? state.active || 0 : Math.max(0, state.sheets.findIndex((s) => !s.hidden));
    const wbRels = [];
    const caches = (state.pivotCaches || []).map((p, i) => {
      // Excel vuelve a calcular la tabla dinámica al abrir, así recoge los cambios hechos en Oficina.
      let def = p.def.replace(/(<(?:\w+:)?pivotCacheDefinition\b)([^>]*)>/, (all, tag, attrs) => `${tag}${attrs.replace(/\s+refreshOnLoad="[^"]*"/, '')} refreshOnLoad="1">`);
      zip.file(`xl/pivotCache/pivotCacheDefinition${i + 1}.xml`, def);
      types.push(`<Override PartName="/xl/pivotCache/pivotCacheDefinition${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>`);
      if (p.records && p.recId) {
        zip.file(`xl/pivotCache/pivotCacheRecords${i + 1}.xml`, p.records);
        zip.file(`xl/pivotCache/_rels/pivotCacheDefinition${i + 1}.xml.rels`, `${O.XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${O.esc(p.recId)}" Type="${R}/pivotCacheRecords" Target="pivotCacheRecords${i + 1}.xml"/></Relationships>`);
        types.push(`<Override PartName="/xl/pivotCache/pivotCacheRecords${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>`);
      }
      wbRels.push(`<Relationship Id="rIdPC${i + 1}" Type="${R}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition${i + 1}.xml"/>`);
      return `<pivotCache cacheId="${O.esc(p.id)}" r:id="rIdPC${i + 1}"/>`;
    });
    // Sin tema propio se guarda el de Office con el que dibuja Oficina; si no, Excel usaría su tema nuevo y cambiarían los colores.
    const themeXml = state.theme || themePart(theme);
    if (themeXml) {
      zip.file('xl/theme/theme1.xml', themeXml);
      wbRels.push(`<Relationship Id="rIdTheme" Type="${R}/theme" Target="theme/theme1.xml"/>`);
      types.push('<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>');
    }
    if (state.vba) {
      zip.file('xl/vbaProject.bin', state.vba, { base64: true });
      wbRels.push('<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>');
      types.push('<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>');
    }
    const WBPROT = ['workbookPassword', 'workbookAlgorithmName', 'workbookHashValue', 'workbookSaltValue', 'workbookSpinCount', 'lockStructure', 'lockWindows'];
    const wbProt = state.wbProtect ? `<workbookProtection${WBPROT.filter((k) => state.wbProtect[k] !== undefined).map((k) => ` ${k}="${O.esc(state.wbProtect[k])}"`).join('')}/>` : '';
    const wbPr = state.codeName ? `<workbookPr codeName="${O.esc(state.codeName)}"/>` : '';
    // Cada hoja conserva su número: las segmentaciones de tablas dinámicas lo usan.
    const sheetIds = new Set();
    let maxSheetId = Math.max(0, ...state.sheets.map((s) => s.sheetId || 0));
    const idOfSheet = state.sheets.map((s) => {
      let id = s.sheetId;
      if (!id || sheetIds.has(id)) id = ++maxSheetId;
      sheetIds.add(id);
      return id;
    });
    const wbExts = (state.wbExts || []).map((f) => fragmentXml(f, relater(wbRels, 'rIdW'))).join('');
    const wbNs = wbExts ? ` xmlns:mc="${MC}" xmlns:x14="${X14}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"` : '';
    zip.file('xl/workbook.xml', `${O.XML}<workbook xmlns="${MAIN}" xmlns:r="${R}"${wbNs}>${wbPr}${wbProt}<bookViews><workbookView activeTab="${activeTab}"/></bookViews><sheets>${state.sheets.map((s, i) => `<sheet name="${O.esc(s.name)}" sheetId="${idOfSheet[i]}"${s.hidden ? ` state="${s.veryHidden ? 'veryHidden' : 'hidden'}"` : ''} r:id="rId${i + 1}"/>`).join('')}</sheets>${definedNames}<calcPr fullCalcOnLoad="1"/>${caches.length ? `<pivotCaches>${caches.join('')}</pivotCaches>` : ''}${wbExts ? `<extLst>${wbExts}</extLst>` : ''}</workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `${O.XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${state.sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>${wbRels.join('')}</Relationships>`);
    zip.file('_rels/.rels', `${O.XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    const IMG = { png: 'image/png', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };
    const defaults = [...exts].map((e) => `<Default Extension="${e}" ContentType="${IMG[e] || 'image/' + e}"/>`).join('');
    const mainType = state.vba || state.macro ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
    zip.file('[Content_Types].xml', `${O.XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${defaults}${hasVml ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : ''}${noteSheets.map((i) => `<Override PartName="/xl/comments${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`).join('')}<Override PartName="/xl/workbook.xml" ContentType="${mainType}"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${state.sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}${types.join('')}</Types>`);
    return O.zipOut(zip, state.vba || state.macro ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }

  // Con macros, el libro se guarda como .xlsm.
  O.xlsx = { read: readXlsx, write: writeXlsx, ext: (state) => (state.vba || state.macro ? 'xlsm' : 'xlsx') };
})(globalThis);
