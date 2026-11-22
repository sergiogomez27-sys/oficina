// Hoja de calculo: cuadricula, edicion, seleccion, portapapeles y deshacer sobre core/formula.js.
(function () {
  const O = window.Oficina;
  const NF = [['General', 'General'], ['0', 'Entero'], ['0.00', '2 decimales'], ['#,##0.00', 'Miles'], ['#,##0.00 "€"', 'Euros'], ['0%', 'Porcentaje'], ['0.00%', '% con decimales'], ['dd/mm/yyyy', 'Fecha']];
  // Tamaños por defecto de Excel: 8,43 caracteres (64 px) de ancho y 15 puntos (20 px) de alto.
  const DEFAULT_W = 64, DEFAULT_H = 20;

  function mount(root, state, onChange) {
    root.innerHTML = `
      <div class="toolbar" role="toolbar" aria-label="Formato de celdas">
        <button class="tb" id="s-undo" title="Deshacer (Ctrl+Z)">↶</button>
        <button class="tb" id="s-redo" title="Rehacer (Ctrl+Y)">↷</button>
        <span class="sep"></span>
        <button class="tb" id="s-b" title="Negrita (Ctrl+B)"><b>N</b></button>
        <button class="tb" id="s-i" title="Cursiva (Ctrl+I)"><i>K</i></button>
        <label class="tb tb-color" title="Color del texto"><i>A</i><input type="color" id="s-color" value="#c0392b" aria-label="Color del texto"></label>
        <label class="tb tb-color" title="Color de relleno"><i>▒</i><input type="color" id="s-fill" value="#fff2b3" aria-label="Color de relleno"></label>
        <button class="tb" id="s-nofill" title="Quitar colores">Sin color</button>
        <span class="sep"></span>
        <button class="tb" data-align="left" title="Alinear a la izquierda">⇤</button>
        <button class="tb" data-align="center" title="Centrar">↔</button>
        <button class="tb" data-align="right" title="Alinear a la derecha">⇥</button>
        <button class="tb" id="s-merge" title="Combinar y centrar">⊞ Combinar</button>
        <button class="tb tb-more" id="s-merge-more" title="Más opciones de combinar" aria-haspopup="menu">▾</button>
        <span class="sep"></span>
        <select id="s-nf" aria-label="Formato de número">${NF.map(([v, l]) => `<option value='${v}'>${l}</option>`).join('')}</select>
        <button class="tb" id="s-sum" title="Sumar lo que hay encima o a la izquierda">Σ Suma</button>
        <button class="tb" id="s-fx" title="Insertar función (Mayús+F3)">fx Función</button>
        <span class="sep"></span>
        <button class="tb" id="s-cf" title="Formato condicional" aria-haspopup="menu">Formato condicional ▾</button>
        <button class="tb" id="s-filter" title="Filtro (Ctrl+Mayús+L)">⏷ Filtro</button>
        <button class="tb" id="s-chart" title="Insertar gráfico con los datos seleccionados" aria-haspopup="menu">Gráfico ▾</button>
      </div>
      <div class="fbar"><span class="ref" id="s-ref">A1</span><button type="button" class="fx" id="s-fxbar" title="Insertar función">fx</button><input id="s-formula" aria-label="Contenido de la celda" spellcheck="false" autocomplete="off"></div>
      <div class="grid-wrap" tabindex="0" aria-label="Hoja de cálculo"></div>
      <div class="sheet-tabs"></div>`;
    const q = (s) => root.querySelector(s);
    const wrap = q('.grid-wrap');
    const fbar = q('#s-formula');
    let book = new O.Workbook(state);
    let tds = [], rowHeads = [], colHeads = [], rows = 0, cols = 0;
    let sel = { r1: 0, c1: 0, r2: 0, c2: 0, r: 0, c: 0 };
    let editor = null, dragging = false, marked = [], clip = null, cfMap = null;
    const undo = [], redo = [];

    const sheet = () => state.sheets[state.active || 0];
    // Las imágenes, el tema y las macros no cambian al deshacer: se dejan fuera para que cada paso pese poco.
    const snapshot = () => JSON.stringify({ sheets: state.sheets, names: state.names, active: state.active, wbProtect: state.wbProtect, pivotDefs: (state.pivotCaches || []).map((p) => p.def) });
    // Con la hoja protegida, como en Excel, solo se pueden cambiar las celdas desbloqueadas y lo que permita la protección.
    const allows = (perm) => !!(sheet().protect && sheet().protect[perm] === '0');
    const unlockedBox = (box) => {
      for (let r = box.r1; r <= box.r2; r++) for (let c = box.c1; c <= box.c2; c++) {
        if (!(((sheet().cells[O.addr(r, c)] || {}).s || {}).unlocked)) return false;
      }
      return true;
    };
    const locked = (box, perm) => {
      if (!sheet().protected) return false;
      if (perm && allows(perm)) return false;
      if (box && (box.r2 - box.r1 + 1) * (box.c2 - box.c1 + 1) <= 20000 && unlockedBox(box)) return false;
      O.notice('La celda o el gráfico que intenta cambiar están en una hoja protegida. Para cambiarlos, desproteja la hoja con clic derecho en su pestaña.');
      return true;
    };
    const hiddenFormula = (r, c) => sheet().protected && (((sheet().cells[O.addr(r, c)] || {}).s || {}).hideF);
    // Con el libro protegido no se pueden añadir, quitar, mover ni cambiar de nombre las hojas.
    const bookLocked = () => {
      if (!(state.wbProtect && ['1', 'true'].includes(state.wbProtect.lockStructure))) return false;
      O.notice('El libro está protegido y no se puede cambiar su estructura. Desprotéjalo con clic derecho en una pestaña.');
      return true;
    };
    const beforeChange = () => {
      undo.push(snapshot());
      if (undo.length > 100) undo.shift();
      redo.length = 0;
    };
    const restore = (json) => {
      const parsed = JSON.parse(json);
      state.sheets = parsed.sheets;
      state.names = parsed.names;
      state.active = Math.min(parsed.active || 0, state.sheets.length - 1);
      if (parsed.wbProtect) state.wbProtect = parsed.wbProtect;
      else delete state.wbProtect;
      (parsed.pivotDefs || []).forEach((def, i) => { if (state.pivotCaches && state.pivotCaches[i]) state.pivotCaches[i].def = def; });
      book = new O.Workbook(state);
      build();
      onChange();
    };

    const used = () => {
      let mr = 0, mc = 0;
      for (const k of Object.keys(sheet().cells)) {
        const a = O.parseAddr(k);
        if (a.r > mr) mr = a.r;
        if (a.c > mc) mc = a.c;
      }
      for (const m of merges()) {
        if (m.r2 > mr && m.r2 < 5000) mr = m.r2;
        if (m.c2 > mc && m.c2 < 500) mc = m.c2;
      }
      return { mr, mc };
    };
    // La cuadrícula también llega hasta los dibujos, para que se vean enteros.
    const drawnEdge = () => {
      let mr = 0, mc = 0;
      for (const d of sheet().drawings || []) {
        const p = d.to || d.from;
        mr = Math.max(mr, Math.min(p.r + (d.to ? 0 : Math.ceil(((d.size || {}).h || 100) / DEFAULT_H)), 5000));
        mc = Math.max(mc, Math.min(p.c + (d.to ? 0 : Math.ceil(((d.size || {}).w || 100) / DEFAULT_W)), 500));
      }
      return { mr, mc };
    };

    // ---------- Celdas combinadas ----------
    // sheet.merges guarda rangos como "A1:B2". mergeIndex da, para cada celda cubierta, su rango.
    const merges = () => (sheet().merges || []).map((m) => O.parseRanges(m)[0]).filter(Boolean);
    let mergeIndex = new Map();
    const mergeAt = (r, c) => mergeIndex.get(r + ',' + c) || null;
    function indexMerges() {
      mergeIndex = new Map();
      for (const m of merges()) {
        for (let r = m.r1; r <= Math.min(m.r2, rows - 1); r++) for (let c = m.c1; c <= Math.min(m.c2, cols - 1); c++) mergeIndex.set(r + ',' + c, m);
      }
    }
    // Amplía un rectángulo hasta que incluya enteras las celdas combinadas que toca, como hace Excel al seleccionar.
    function growToMerges(box) {
      const list = merges();
      for (let changed = true; changed;) {
        changed = false;
        for (const m of list) {
          if (m.r1 > box.r2 || m.r2 < box.r1 || m.c1 > box.c2 || m.c2 < box.c1) continue;
          if (m.r1 < box.r1) { box.r1 = m.r1; changed = true; }
          if (m.r2 > box.r2) { box.r2 = Math.min(m.r2, rows - 1); changed = true; }
          if (m.c1 < box.c1) { box.c1 = m.c1; changed = true; }
          if (m.c2 > box.c2) { box.c2 = Math.min(m.c2, cols - 1); changed = true; }
        }
      }
      return box;
    }
    // Siguiente celda al moverse con flechas, Intro o Tab: se salta el resto de una celda combinada.
    function step(r, c, dr, dc) {
      const m = mergeAt(r, c);
      let nr = dr > 0 ? (m ? m.r2 : r) + dr : dr < 0 ? (m ? m.r1 : r) + dr : r;
      let nc = dc > 0 ? (m ? m.c2 : c) + dc : dc < 0 ? (m ? m.c1 : c) + dc : c;
      const s = sheet();
      while ((dr && s.hiddenRows && s.hiddenRows[nr]) || (dc && s.hiddenCols && s.hiddenCols[nc])) { nr += dr; nc += dc; }
      return [nr, nc];
    }
    function applyMerges() {
      for (const m of merges()) {
        if (m.r1 >= rows || m.c1 >= cols) continue;
        const anchor = tds[m.r1][m.c1];
        anchor.rowSpan = Math.min(m.r2, rows - 1) - m.r1 + 1;
        anchor.colSpan = Math.min(m.c2, cols - 1) - m.c1 + 1;
        anchor.classList.add('merged');
        for (let r = m.r1; r <= Math.min(m.r2, rows - 1); r++) for (let c = m.c1; c <= Math.min(m.c2, cols - 1); c++) {
          if (r !== m.r1 || c !== m.c1) tds[r][c].style.display = 'none';
        }
      }
    }
    function mergeSelection(mode) {
      if (locked()) return;
      const s = sheet();
      const box = { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 };
      const inside = (m) => !(m.r1 > box.r2 || m.r2 < box.r1 || m.c1 > box.c2 || m.c2 < box.c1);
      if (mode === 'unmerge' || (mode === 'center' && merges().some((m) => m.r1 === box.r1 && m.c1 === box.c1 && m.r2 === box.r2 && m.c2 === box.c2))) {
        // Separar, o Combinar y centrar sobre una celda ya combinada, que la separa como en Excel.
        if (!merges().some(inside)) return;
        beforeChange();
        s.merges = (s.merges || []).filter((x) => !inside(O.parseRanges(x)[0]));
        if (!s.merges.length) delete s.merges;
        build();
        onChange();
        return;
      }
      const areas = mode === 'across' ? Array.from({ length: box.r2 - box.r1 + 1 }, (_, i) => ({ r1: box.r1 + i, r2: box.r1 + i, c1: box.c1, c2: box.c2 })) : [box];
      if (areas.every((a) => a.r1 === a.r2 && a.c1 === a.c2)) return;
      const lost = areas.some((a) => {
        for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) {
          if (r === a.r1 && c === a.c1) continue;
          const cell = s.cells[O.addr(r, c)];
          if (cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== ''))) return true;
        }
        return false;
      });
      if (lost && !window.confirm('Al combinar celdas, solo se mantiene el valor de la celda superior izquierda y se descartan los demás valores.')) return;
      beforeChange();
      s.merges = (s.merges || []).filter((x) => !inside(O.parseRanges(x)[0]));
      for (const a of areas) {
        if (a.r1 === a.r2 && a.c1 === a.c2) continue;
        for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) {
          if (r === a.r1 && c === a.c1) continue;
          delete s.cells[O.addr(r, c)];
        }
        s.merges.push(O.rangesText([a]));
        if (mode === 'center') {
          const key = O.addr(a.r1, a.c1);
          const cell = s.cells[key] || (s.cells[key] = {});
          cell.s = { ...(cell.s || {}), align: 'center' };
        }
      }
      sel = { r: box.r1, c: box.c1, ...box };
      build();
      onChange();
    }
    const mergeItems = () => [
      { label: 'Combinar y centrar', action: () => mergeSelection('center') },
      { label: 'Combinar horizontalmente', action: () => mergeSelection('across') },
      { label: 'Combinar celdas', action: () => mergeSelection('merge') },
      { label: 'Separar celdas', action: () => mergeSelection('unmerge') },
    ];

    function build() {
      const { mr, mc } = used();
      const edge = drawnEdge();
      rows = Math.max(100, mr + 30, edge.mr + 5);
      cols = Math.max(26, mc + 6, edge.mc + 3);
      const s = sheet();
      const table = document.createElement('table');
      table.className = 'grid';
      const colgroup = document.createElement('colgroup');
      colgroup.innerHTML = '<col style="width:46px">';
      let total = 46;
      const thead = document.createElement('thead');
      const hr = document.createElement('tr');
      hr.append(document.createElement('th'));
      colHeads = [];
      for (let c = 0; c < cols; c++) {
        const w = (s.cols && s.cols[c]) || s.defW || DEFAULT_W;
        const off = s.hiddenCols && s.hiddenCols[c];
        if (!off) total += w;
        const col = document.createElement('col');
        col.style.width = w + 'px';
        if (off) col.style.visibility = 'collapse';
        colgroup.append(col);
        const th = document.createElement('th');
        th.textContent = O.colName(c);
        th.dataset.col = c;
        const h = document.createElement('span');
        h.className = 'resize';
        th.append(h);
        hr.append(th);
        colHeads.push(th);
      }
      thead.append(hr);
      const tbody = document.createElement('tbody');
      tds = [];
      rowHeads = [];
      for (let r = 0; r < rows; r++) {
        const tr = document.createElement('tr');
        if (s.hiddenRows && s.hiddenRows[r]) tr.style.display = 'none';
        const height = (s.rowH && s.rowH[r]) || s.defH;
        const th = document.createElement('th');
        th.textContent = r + 1;
        th.dataset.row = r;
        if (height) th.style.height = height + 'px';
        rowHeads.push(th);
        tr.append(th);
        const line = [];
        for (let c = 0; c < cols; c++) {
          const td = document.createElement('td');
          td.dataset.r = r;
          td.dataset.c = c;
          if (height) td.style.height = height + 'px';
          tr.append(td);
          line.push(td);
        }
        tds.push(line);
        tbody.append(tr);
      }
      indexMerges();
      applyMerges();
      table.style.width = total + 'px';
      table.append(colgroup, thead, tbody);
      wrap.replaceChildren(table);
      paintAll();
      clampSel();
      paintSel();
      renderTabs();
      placeFilterButtons();
    }

    // Estilos de borde de Excel: los continuos con CSS y los de trazos con su patrón en píxeles (trazo, hueco, trazo…).
    const BORDER_CSS = { thin: '1px solid', medium: '2px solid', thick: '3px solid', double: '3px double' };
    const LINE_STYLES = [['thin', 'Fina'], ['medium', 'Media'], ['thick', 'Gruesa'], ['double', 'Doble'], ['hair', 'Muy fina'], ['dotted', 'Punteada'],
      ['dashed', 'Discontinua'], ['dashDot', 'Raya y punto'], ['dashDotDot', 'Raya y dos puntos'], ['mediumDashed', 'Discontinua media'],
      ['mediumDashDot', 'Raya y punto media'], ['mediumDashDotDot', 'Raya y dos puntos media'], ['slantDashDot', 'Raya y punto inclinada']];
    const BORDER_DASH = {
      hair: [1, [1, 1]], dotted: [1, [2, 2]], dashed: [1, [3, 1]], dashDot: [1, [9, 3, 3, 3]], dashDotDot: [1, [9, 3, 3, 3, 3, 3]],
      mediumDashed: [2, [9, 3]], mediumDashDot: [2, [9, 3, 3, 3]], mediumDashDotDot: [2, [9, 3, 3, 3, 3, 3]], slantDashDot: [2, [11, 1, 5, 1]],
    };
    const dashLayer = (side, width, pattern, color) => {
      let at = 0;
      const stops = pattern.map((len, i) => {
        const from = at;
        at += len;
        return `${i % 2 ? 'transparent' : '#' + color} ${from}px ${at}px`;
      }).join(', ');
      const across = side === 't' || side === 'b';
      return {
        image: `repeating-linear-gradient(${across ? 90 : 180}deg, ${stops})`,
        size: across ? `100% ${width}px` : `${width}px 100%`,
        pos: { t: '0 0', b: '0 100%', l: '0 0', r: '100% 0' }[side],
      };
    };
    function paint(r, c) {
      const td = tds[r] && tds[r][c];
      if (!td) return;
      const s = sheet();
      const cell = s.cells[O.addr(r, c)];
      td.textContent = book.display(s, r, c);
      // El estilo de la tabla va debajo del formato propio de la celda y el formato condicional, encima de los dos.
      const fx = (cfMap && cfMap.get(r + ',' + c)) || {};
      const own = (cell && cell.s) || {};
      const tsx = s.tables ? O.tableStyleAt(s, r, c, state.themeColors, state.tableStyles) : null;
      const st = { ...(tsx || {}), ...own, ...fx };
      if (tsx && tsx.bd) {
        const bdx = {};
        for (const k of tsx.bd) bdx[k] = tsx.bdx[k];
        for (const k of own.bd || '') bdx[k] = (own.bdx && own.bdx[k]) || ['thin', '000000'];
        st.bd = Object.keys(bdx).join('');
        st.bdx = bdx;
      }
      const v = cell ? book.value(s, r, c) : null;
      const isErr = !!(v && v.err);
      const note = cell && cell.note, link = cell && cell.link;
      td.className = (!st.align && (typeof v === 'number') ? 'num ' : '') + (isErr ? 'err ' : '') + (note ? 'has-note' : '');
      td.style.fontWeight = st.b ? '700' : '';
      td.style.fontStyle = st.i ? 'italic' : '';
      td.style.textDecoration = [st.u || link ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ');
      td.style.color = isErr ? '' : st.color ? '#' + st.color : link ? '#0563c1' : '';
      td.style.backgroundColor = st.fill ? '#' + st.fill : '';
      // Barra con degradado, como las de Excel, para que el número se lea encima.
      const pct = st.bar ? Math.round(st.bar.pct * 100) : 0;
      const layers = st.bar ? [{ image: `linear-gradient(90deg, #${st.bar.color} 0%, #${st.bar.color}40 ${pct}%, transparent ${pct}%)`, size: '100% 100%', pos: '0 0' }] : [];
      // Texto girado: 90 sube de abajo arriba, 91-180 baja de arriba abajo, 255 es vertical con letras apiladas.
      if (st.rot && td.textContent) {
        const turned = document.createElement('span');
        turned.className = 'rot';
        turned.textContent = td.textContent;
        if (st.rot === 90) turned.style.cssText = 'writing-mode:vertical-rl;transform:rotate(180deg)';
        else if (st.rot === 180) turned.style.cssText = 'writing-mode:vertical-rl';
        else if (st.rot === 255) turned.style.cssText = 'writing-mode:vertical-rl;text-orientation:upright';
        else turned.style.transform = `rotate(${st.rot <= 90 ? -st.rot : st.rot - 90}deg)`;
        td.replaceChildren(turned);
      }
      // Como en Excel, el icono va pegado a la izquierda y el valor sigue en su sitio.
      const icon = st.icon && O.ICON_SETS[st.icon.set] && O.ICON_SETS[st.icon.set][st.icon.index];
      if (icon) {
        if (st.hideValue) td.textContent = '';
        const mark = document.createElement('span');
        mark.className = 'cf-icon';
        mark.dataset.icon = icon[0];
        mark.innerHTML = O.iconSvg(icon[0], icon[1]);
        td.prepend(mark);
      }
      td.style.textAlign = st.align || '';
      // Como en Excel, sin alineación propia el texto va abajo.
      td.style.verticalAlign = st.valign === 'center' ? 'middle' : st.valign || 'bottom';
      td.style.fontSize = st.sz ? st.sz + 'pt' : '';
      td.style.fontFamily = st.font ? `"${st.font.replace(/"/g, '')}", Calibri, Carlito, sans-serif` : '';
      td.style.whiteSpace = st.wrap ? 'normal' : '';
      // Como en Excel, un texto que no cabe sigue por encima de la celda de la derecha si está vacía.
      const right = s.cells[O.addr(r, c + 1)];
      const spill = typeof v === 'string' && v !== '' && !st.wrap && !st.rot && (!st.align || st.align === 'left') && !mergeAt(r, c)
        && !(right && (right.f !== undefined || (right.v !== undefined && right.v !== null && right.v !== ''))) && !mergeAt(r, c + 1);
      td.style.overflow = spill ? 'visible' : '';
      td.style.textOverflow = spill ? 'clip' : '';
      const bd = st.bd || '';
      const line = (k) => {
        if (!bd.includes(k)) return '';
        const [style, color] = (st.bdx && st.bdx[k]) || ['thin', '000000'];
        const dash = BORDER_DASH[style];
        if (!dash) return (BORDER_CSS[style] || BORDER_CSS.thin) + ' #' + color;
        // Los trazos se pintan encima de un borde transparente del mismo grosor; en los huecos se ve el relleno.
        layers.unshift(dashLayer(k, dash[0], dash[1], color));
        return `${dash[0]}px solid transparent`;
      };
      td.style.borderLeft = line('l');
      td.style.borderRight = line('r');
      td.style.borderTop = line('t');
      td.style.borderBottom = line('b');
      td.style.backgroundImage = layers.map((x) => x.image).join(', ');
      td.style.backgroundSize = layers.length ? layers.map((x) => x.size).join(', ') : '';
      td.style.backgroundPosition = layers.length ? layers.map((x) => x.pos).join(', ') : '';
      td.style.backgroundRepeat = layers.length ? 'no-repeat' : '';
      td.style.backgroundOrigin = layers.length ? 'border-box' : '';
      const tips = [];
      if (note) tips.push(note);
      if (link) tips.push(`${link} (Ctrl+clic para abrir)`);
      td.title = tips.join('\n');
      // Como en Excel, un número que no cabe se muestra como ###. Se mide después, todas las celdas a la vez.
      const fits = typeof v === 'number' && !st.wrap && !st.rot && !st.icon && !mergeAt(r, c);
      if (fits) { td.dataset.num = '1'; toFit.add(td); } else delete td.dataset.num;
      if (fits && toFit.size === 1) queueMicrotask(fitNumbers);
    }
    const toFit = new Set();
    function fitNumbers() {
      const list = [...toFit];
      toFit.clear();
      const over = list.filter((td) => td.isConnected && td.dataset.num && td.scrollWidth > td.clientWidth + 1);
      for (const td of over) td.textContent = '#'.repeat(Math.max(1, Math.floor((td.clientWidth - 6) / 7.5)));
    }
    function paintAll() {
      book.clearCache();
      const s = sheet();
      cfMap = s.cf && s.cf.length ? O.conditionalFormats(book, s, rows - 1, cols - 1) : null;
      for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
        if (tds[r][c].textContent || tds[r][c].style.cssText || s.cells[O.addr(r, c)] || (cfMap && cfMap.has(r + ',' + c))) paint(r, c);
      }
      // Las celdas vacías de una tabla también llevan su estilo.
      for (const t of s.tables || []) {
        const b = O.tableBox(t);
        for (let r = b.r1; r <= Math.min(b.r2, rows - 1); r++) for (let c = b.c1; c <= Math.min(b.c2, cols - 1); c++) if (!s.cells[O.addr(r, c)]) paint(r, c);
      }
      // Los gráficos se vuelven a dibujar con los valores nuevos.
      renderDrawings();
    }
    // Solo repinta las celdas con formula, que son las que pueden cambiar con un valor nuevo.
    // Con formato condicional cualquier celda puede cambiar, así que se repinta todo.
    function recalc() {
      if (sheet().cf && sheet().cf.length) { paintAll(); return; }
      book.clearCache();
      for (const sh of Object.keys(sheet().cells)) {
        const a = O.parseAddr(sh);
        if (sheet().cells[sh].f !== undefined) paint(a.r, a.c);
      }
      if ((sheet().drawings || []).some((d) => d.kind === 'chart' || d.kind === 'slicer') || (sheet().controls || []).length) renderDrawings();
    }

    function clampSel() {
      sel.r = Math.min(sel.r, rows - 1);
      sel.c = Math.min(sel.c, cols - 1);
      sel.r1 = Math.min(sel.r1, rows - 1); sel.r2 = Math.min(sel.r2, rows - 1);
      sel.c1 = Math.min(sel.c1, cols - 1); sel.c2 = Math.min(sel.c2, cols - 1);
    }
    function paintSel() {
      for (const td of marked) td.classList.remove('sel', 'cur');
      root.querySelectorAll('th.hl').forEach((t) => t.classList.remove('hl'));
      marked = [];
      for (let r = sel.r1; r <= sel.r2; r++) {
        rowHeads[r].classList.add('hl');
        for (let c = sel.c1; c <= sel.c2; c++) {
          tds[r][c].classList.add('sel');
          marked.push(tds[r][c]);
        }
      }
      for (let c = sel.c1; c <= sel.c2; c++) colHeads[c].classList.add('hl');
      tds[sel.r][sel.c].classList.add('cur');
      if (!marked.includes(tds[sel.r][sel.c])) marked.push(tds[sel.r][sel.c]);
      const multi = sel.r1 !== sel.r2 || sel.c1 !== sel.c2;
      q('#s-ref').textContent = multi ? `${O.addr(sel.r1, sel.c1)}:${O.addr(sel.r2, sel.c2)}` : O.addr(sel.r, sel.c);
      if (!editor) fbar.value = hiddenFormula(sel.r, sel.c) ? '' : book.source(sheet(), sel.r, sel.c);
      const st = (sheet().cells[O.addr(sel.r, sel.c)] || {}).s || {};
      q('#s-b').classList.toggle('on', !!st.b);
      q('#s-i').classList.toggle('on', !!st.i);
      q('#s-nf').value = NF.some(([v]) => v === (st.nf || 'General')) ? st.nf || 'General' : 'General';
      root.querySelectorAll('[data-align]').forEach((b) => b.classList.toggle('on', st.align === b.dataset.align));
      placeListButton();
    }
    // ---------- Validación de datos ----------
    const listBtn = document.createElement('button');
    listBtn.type = 'button';
    listBtn.className = 'dv-btn';
    listBtn.textContent = '▾';
    listBtn.title = 'Elegir de la lista (Alt+Abajo)';
    listBtn.hidden = true;
    listBtn.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); openValidationList(); });
    // Botón de lista junto a la celda activa, solo si su validación es una lista con desplegable.
    function listHit() {
      const hit = O.validationAt(sheet(), sel.r, sel.c);
      // En el archivo, showDropDown="1" significa ocultar el desplegable.
      return hit && hit.rule.attrs && hit.rule.attrs.type === 'list' && !isTrue(hit.rule.attrs.showDropDown) ? hit : null;
    }
    function placeListButton() {
      const td = tds[sel.r] && tds[sel.r][sel.c];
      const show = !!(td && listHit());
      listBtn.hidden = !show;
      if (!show) return;
      if (listBtn.parentNode !== wrap) wrap.append(listBtn);
      listBtn.style.left = td.offsetLeft + td.offsetWidth + 'px';
      listBtn.style.top = td.offsetTop + td.offsetHeight - 18 + 'px';
    }
    function openValidationList() {
      const hit = listHit();
      if (!hit || locked({ r1: sel.r, r2: sel.r, c1: sel.c, c2: sel.c })) return;
      const r = sel.r, c = sel.c;
      const current = book.display(sheet(), r, c);
      const items = O.validationList(book, sheet(), hit, r, c).map((x) => ({
        label: x, checked: x === current,
        action: () => {
          beforeChange();
          setCell(r, c, x);
          if (!growIfNeeded(r, c)) { for (let k = 0; k < cols; k++) paint(r, k); recalc(); }
          onChange();
          paintSel();
          wrap.focus();
        },
      }));
      if (!items.length) return;
      const box = tds[r][c].getBoundingClientRect();
      O.contextMenu(box.left, box.bottom, items);
    }
    const isTrue = (v) => v === '1' || v === 'true';
    // Comprueba lo escrito contra la validación. Devuelve false si hay que seguir editando.
    function passesValidation(r, c, raw) {
      const hit = O.validationAt(sheet(), r, c);
      if (!hit) return true;
      const a = hit.rule.attrs || {};
      const parsed = O.parseInput(raw);
      const value = parsed.empty ? null : parsed.f !== undefined ? book.evaluate(sheet(), parsed.f) : parsed.v;
      if (!isTrue(a.showErrorMessage) || O.validates(book, sheet(), hit, value, r, c)) return true;
      const msg = a.error || 'Este valor no coincide con las restricciones de validación de datos definidas para esta celda.';
      const title = a.errorTitle ? a.errorTitle + '\n\n' : '';
      if (a.errorStyle === 'warning') return window.confirm(`${title}${msg}\n\n¿Desea continuar?`);
      if (a.errorStyle === 'information') { window.alert(title + msg); return true; }
      O.notice((a.errorTitle ? a.errorTitle + ': ' : '') + msg);
      return false;
    }
    function select(r, c, extend) {
      r = Math.max(0, Math.min(rows - 1, r));
      c = Math.max(0, Math.min(cols - 1, c));
      if (extend) {
        sel = { r: sel.r, c: sel.c, ...growToMerges({ r1: Math.min(sel.r, r), r2: Math.max(sel.r, r), c1: Math.min(sel.c, c), c2: Math.max(sel.c, c) }) };
      } else {
        // Dentro de una celda combinada se selecciona entera y se trabaja con su primera celda.
        const m = mergeAt(r, c);
        sel = m ? { r: m.r1, c: m.c1, r1: m.r1, c1: m.c1, r2: Math.min(m.r2, rows - 1), c2: Math.min(m.c2, cols - 1) } : { r, c, r1: r, r2: r, c1: c, c2: c };
      }
      paintSel();
      const target = extend ? mergeAt(r, c) : null;
      const td = tds[target ? target.r1 : extend ? r : sel.r][target ? target.c1 : extend ? c : sel.c];
      if (td) scrollIntoView(td);
    }
    function scrollIntoView(td) {
      const t = td.offsetTop, l = td.offsetLeft, headH = 23, headW = 46;
      if (t - headH < wrap.scrollTop) wrap.scrollTop = t - headH;
      else if (t + td.offsetHeight > wrap.scrollTop + wrap.clientHeight) wrap.scrollTop = t + td.offsetHeight - wrap.clientHeight;
      if (l - headW < wrap.scrollLeft) wrap.scrollLeft = l - headW;
      else if (l + td.offsetWidth > wrap.scrollLeft + wrap.clientWidth) wrap.scrollLeft = l + td.offsetWidth - wrap.clientWidth;
    }

    function setCell(r, c, raw) {
      const key = O.addr(r, c);
      // Dentro de una tabla, [@Ventas] se guarda con el nombre de la tabla, como en el archivo.
      const parsed = O.parseInput(raw, (O.tableAt(sheet(), r, c) || {}).name);
      const prev = sheet().cells[key];
      const extras = {};
      if (prev && prev.note) extras.note = prev.note;
      if (prev && prev.link) extras.link = prev.link;
      if (parsed.empty) {
        if (prev && prev.s) extras.s = prev.s;
        if (Object.keys(extras).length) sheet().cells[key] = extras;
        else delete sheet().cells[key];
        return;
      }
      const cell = extras;
      const style = prev && prev.s ? { ...prev.s } : {};
      if (parsed.nf && !style.nf) style.nf = parsed.nf;
      if (Object.keys(style).length) cell.s = style;
      if (parsed.f !== undefined) cell.f = parsed.f;
      else cell.v = parsed.v;
      sheet().cells[key] = cell;
    }
    const growIfNeeded = (r, c) => {
      if (r >= rows - 5 || c >= cols - 2) { build(); return true; }
      return false;
    };

    // ---------- Edicion ----------
    function startEdit(initial) {
      if (locked({ r1: sel.r, r2: sel.r, c1: sel.c, c2: sel.c })) return;
      const td = tds[sel.r][sel.c];
      scrollIntoView(td);
      const input = document.createElement('input');
      input.className = 'cell-editor';
      input.spellcheck = false;
      input.autocomplete = 'off';
      input.setAttribute('aria-label', 'Editar celda ' + O.addr(sel.r, sel.c));
      input.value = initial !== undefined ? initial : book.source(sheet(), sel.r, sel.c);
      input.style.left = td.offsetLeft - 1 + 'px';
      input.style.top = td.offsetTop - 1 + 'px';
      input.style.minWidth = td.offsetWidth + 3 + 'px';
      input.style.width = Math.max(td.offsetWidth + 3, 8 * input.value.length) + 'px';
      input.style.height = td.offsetHeight + 3 + 'px';
      editor = { input, r: sel.r, c: sel.c };
      wrap.append(input);
      input.focus();
      input.setSelectionRange(input.value.length, input.value.length);
      fbar.value = input.value;
      input.oninput = () => {
        fbar.value = input.value;
        input.style.width = Math.max(td.offsetWidth + 3, 8 * input.value.length) + 'px';
      };
      watchFormula(input);
      input.onkeydown = (e) => {
        e.stopPropagation();
        if (suggestKey(e, input)) { e.preventDefault(); return; }
        if (e.key === 'F3' && e.shiftKey) { e.preventDefault(); functionDialog(); return; }
        if (e.key === 'Enter') { e.preventDefault(); commit(e.shiftKey ? -1 : 1, 0); }
        else if (e.key === 'Tab') { e.preventDefault(); commit(0, e.shiftKey ? -1 : 1); }
        else if (e.key === 'Escape') { e.preventDefault(); cancel(); }
      };
      input.onblur = (e) => { if (editor && editor.input === input && e.relatedTarget !== fbar) commit(0, 0, true); };
    }
    function commit(dr, dc, keepFocus) {
      if (!editor) return;
      const { input, r, c } = editor;
      let val = input.value;
      // Como Excel, una fórmula mal escrita no se acepta: se avisa y se sigue editando.
      if (val.trim().startsWith('=') && val.trim().length > 1 && val !== book.source(sheet(), r, c)) {
        const fixed = O.fixFormulaInput(val);
        if (fixed === null) {
          O.notice('Hay un problema con esta fórmula. Revisa los paréntesis, las comillas y los separadores (;). Si no querías escribir una fórmula, empieza con un apóstrofo (\').');
          setTimeout(() => { if (editor && editor.input === input) input.focus(); });
          return;
        }
        val = fixed;
      }
      if (val !== book.source(sheet(), r, c) && !passesValidation(r, c, val)) {
        setTimeout(() => { if (editor && editor.input === input) { input.focus(); input.select(); } });
        return;
      }
      editor = null;
      hideSuggest();
      input.onblur = null;
      input.remove();
      if (val !== book.source(sheet(), r, c)) {
        beforeChange();
        const before = (sheet().cells[O.addr(r, c)] || {}).f;
        const tablesBefore = JSON.stringify(sheet().tables || null);
        setCell(r, c, val);
        tableAfterEdit(r, c, before);
        // Si la tabla ha cambiado (más filas, columna calculada, otro nombre), se repinta entera.
        if (JSON.stringify(sheet().tables || null) !== tablesBefore) build();
        // Se repinta la fila: el texto de las celdas de la izquierda puede dejar de pasar por encima de esta.
        else if (!growIfNeeded(r, c)) { for (let k = 0; k < cols; k++) paint(r, k); recalc(); }
        onChange();
      }
      select(...(dr || dc ? step(r, c, dr, dc) : [r, c]));
      if (!keepFocus) wrap.focus();
    }
    function cancel() {
      if (!editor) return;
      const { input } = editor;
      editor = null;
      hideSuggest();
      input.onblur = null;
      input.remove();
      paintSel();
      wrap.focus();
    }

    fbar.addEventListener('focus', () => { if (!editor) startFromBar(); });
    function startFromBar() {
      const value = fbar.value;
      startEdit(value);
      fbar.focus();
    }
    fbar.addEventListener('input', () => {
      if (editor) editor.input.value = fbar.value;
    });
    fbar.addEventListener('keydown', (e) => {
      if (suggestKey(e, fbar)) { e.preventDefault(); return; }
      if (e.key === 'Enter') { e.preventDefault(); commit(1, 0); }
      else if (e.key === 'Escape') { e.preventDefault(); cancel(); }
      else if (e.key === 'Tab') { e.preventDefault(); commit(0, 1); }
    });
    fbar.addEventListener('blur', () => {
      setTimeout(() => {
        if (editor && document.activeElement !== editor.input && document.activeElement !== fbar) commit(0, 0, true);
      });
    });

    // ---------- Raton ----------
    // El texto que pasa por encima de las celdas vecinas recibe el clic; la celda buena es la que está debajo del puntero.
    function cellAt(e) {
      const td = e.target.closest && e.target.closest('td');
      if (!td || !td.dataset.c) return td;
      const box = td.getBoundingClientRect();
      if (e.clientX <= box.right) return td;
      const line = tds[+td.dataset.r] || [];
      for (let c = +td.dataset.c + 1; c < line.length; c++) {
        const b = line[c].getBoundingClientRect();
        if (b.width && e.clientX >= b.left && e.clientX <= b.right) return line[c];
      }
      return td;
    }
    wrap.addEventListener('mousedown', (e) => {
      const rz = e.target.closest('.resize');
      if (rz && e.button === 0) { startResize(e, +rz.parentElement.dataset.col); return; }
      const td = cellAt(e);
      const th = e.target.closest('th');
      if (editor && e.target === editor.input) return;
      // Un clic en la cuadrícula suelta el gráfico o la imagen elegidos.
      if (picked >= 0 && (td || th)) { picked = -1; renderDrawings(); }
      if (e.button !== 0) {
        // El clic derecho conserva la seleccion si cae dentro de ella.
        if (!td && !th) return;
        // Si la fórmula no es válida el editor sigue abierto y el clic no hace nada más.
        if (editor) { commit(0, 0, true); if (editor) { e.preventDefault(); return; } }
        e.preventDefault();
        if (td) {
          const r = +td.dataset.r, c = +td.dataset.c;
          if (r < sel.r1 || r > sel.r2 || c < sel.c1 || c > sel.c2) select(r, c);
        } else if (th.dataset.col !== undefined) {
          const c = +th.dataset.col;
          if (!fullCol() || c < sel.c1 || c > sel.c2) { sel = { r: 0, c, r1: 0, r2: rows - 1, c1: c, c2: c }; paintSel(); }
        } else if (th.dataset.row !== undefined) {
          const r = +th.dataset.row;
          if (!fullRow() || r < sel.r1 || r > sel.r2) { sel = { r, c: 0, r1: r, r2: r, c1: 0, c2: cols - 1 }; paintSel(); }
        }
        wrap.focus();
        return;
      }
      if (td) {
        const r = +td.dataset.r, c = +td.dataset.c;
        // Con una formula a medio escribir, el clic inserta la referencia en vez de salir.
        if (editor && editor.input.value.startsWith('=') && /[=(,;+\-*/^&<>:]$/.test(editor.input.value)) {
          e.preventDefault();
          editor.input.value += O.addr(r, c);
          fbar.value = editor.input.value;
          editor.input.focus();
          return;
        }
        // Si la fórmula no es válida el editor sigue abierto y el clic no hace nada más.
        if (editor) { commit(0, 0, true); if (editor) { e.preventDefault(); return; } }
        e.preventDefault();
        const here = sheet().cells[O.addr(r, c)];
        if ((e.ctrlKey || e.metaKey) && here && here.link) openLink(here.link);
        select(r, c, e.shiftKey);
        dragging = true;
        wrap.focus();
      } else if (th) {
        // Si la fórmula no es válida el editor sigue abierto y el clic no hace nada más.
        if (editor) { commit(0, 0, true); if (editor) { e.preventDefault(); return; } }
        e.preventDefault();
        if (th.dataset.col !== undefined) {
          const c = +th.dataset.col;
          sel = { r: 0, c, r1: 0, r2: rows - 1, c1: c, c2: c };
          paintSel();
        } else if (th.dataset.row !== undefined) {
          const r = +th.dataset.row;
          sel = { r, c: 0, r1: r, r2: r, c1: 0, c2: cols - 1 };
          paintSel();
        }
        wrap.focus();
      }
    });
    const onMove = (e) => {
      if (!dragging) return;
      const td = cellAt(e);
      if (td && wrap.contains(td)) select(+td.dataset.r, +td.dataset.c, true);
    };
    const onUp = () => { dragging = false; };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
    wrap.addEventListener('dblclick', (e) => { if (e.target.closest('td')) startEdit(); });

    function startResize(e, c) {
      e.preventDefault();
      e.stopPropagation();
      const col = wrap.querySelectorAll('col')[c + 1];
      const table = wrap.querySelector('table');
      const startX = e.clientX, startW = parseInt(col.style.width, 10), startT = parseInt(table.style.width, 10);
      const move = (m) => {
        const w = Math.max(24, startW + m.clientX - startX);
        col.style.width = w + 'px';
        table.style.width = startT + w - startW + 'px';
      };
      const up = (m) => {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
        beforeChange();
        sheet().cols = sheet().cols || {};
        sheet().cols[c] = Math.max(24, startW + m.clientX - startX);
        for (let r = 0; r < rows; r++) if (tds[r][c].dataset.num) paint(r, c);
        placeFilterButtons();
        renderDrawings();
        onChange();
      };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
    }

    // ---------- Teclado ----------
    wrap.addEventListener('keydown', (e) => {
      if (editor) return;
      const mod = e.ctrlKey || e.metaKey;
      // Con un gráfico o una imagen elegidos, Supr lo borra y Esc lo suelta.
      if (picked >= 0) {
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deletePicked(); return; }
        picked = -1;
        renderDrawings();
        if (e.key === 'Escape') { e.preventDefault(); return; }
      }
      // Ctrl+Mayús+L pone o quita el filtro, como en Excel.
      if (mod && e.shiftKey && e.key.toLowerCase() === 'l') { e.preventDefault(); toggleAutoFilter(); return; }
      const key = e.key;
      const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      // Alt+Abajo abre la lista de la validación de datos, como en Excel.
      if (key === 'ArrowDown' && e.altKey) { e.preventDefault(); openValidationList(); return; }
      if (moves[key]) {
        e.preventDefault();
        const [dr, dc] = moves[key];
        if (e.shiftKey) {
          const er = sel.r1 === sel.r ? sel.r2 : sel.r1;
          const ec = sel.c1 === sel.c ? sel.c2 : sel.c1;
          select(er + dr, ec + dc, true);
        } else {
          // Las filas y columnas ocultas y el resto de una celda combinada se saltan al moverse con las flechas.
          const [nr, nc] = step(sel.r, sel.c, dr, dc);
          if (nr >= 0 && nc >= 0) select(nr, nc);
        }
      } else if (mod && key === '1') { e.preventDefault(); formatDialog(); }
      else if (key === 'F3' && e.shiftKey) { e.preventDefault(); functionDialog(); }
      else if (key === 'Enter') { e.preventDefault(); select(...step(sel.r, sel.c, e.shiftKey ? -1 : 1, 0)); }
      else if (key === 'Tab') { e.preventDefault(); select(...step(sel.r, sel.c, 0, e.shiftKey ? -1 : 1)); }
      else if (key === 'F2') { e.preventDefault(); startEdit(); }
      else if (key === 'Delete' || key === 'Backspace') { e.preventDefault(); clearRange(); }
      else if (mod && key.toLowerCase() === 'z') { e.preventDefault(); doUndo(); }
      else if (mod && key.toLowerCase() === 'y') { e.preventDefault(); doRedo(); }
      else if (mod && key.toLowerCase() === 'b') { e.preventDefault(); toggle('b'); }
      else if (mod && key.toLowerCase() === 'i') { e.preventDefault(); toggle('i'); }
      else if (mod && key.toLowerCase() === 'a') { e.preventDefault(); sel = { r: sel.r, c: sel.c, r1: 0, c1: 0, r2: rows - 1, c2: cols - 1 }; paintSel(); }
      // Las teclas muertas (tildes) abren el editor sin impedir que el sistema componga la letra.
      else if (key === 'Dead' || key === 'Process') startEdit('');
      else if (!mod && !e.altKey && key.length === 1) { e.preventDefault(); startEdit(key); }
    });

    function eachSel(fn) {
      const area = (sel.r2 - sel.r1 + 1) * (sel.c2 - sel.c1 + 1);
      const s = sheet();
      if (area > 2000) {
        for (const k of Object.keys(s.cells)) {
          const a = O.parseAddr(k);
          if (a.r >= sel.r1 && a.r <= sel.r2 && a.c >= sel.c1 && a.c <= sel.c2) fn(a.r, a.c);
        }
        return;
      }
      for (let r = sel.r1; r <= sel.r2; r++) for (let c = sel.c1; c <= sel.c2; c++) fn(r, c);
    }
    function clearRange() {
      if (locked(sel)) return;
      beforeChange();
      eachSel((r, c) => { setCell(r, c, ''); paint(r, c); });
      for (let r = sel.r1; r <= Math.min(sel.r2, rows - 1); r++) for (let k = 0; k < sel.c1; k++) paint(r, k);
      recalc();
      paintSel();
      onChange();
    }
    // Borra propiedades concretas de cada celda: formato (s), nota, vinculo o todo.
    function clearProps(props) {
      if (locked()) return;
      beforeChange();
      eachSel((r, c) => {
        const key = O.addr(r, c);
        const cell = sheet().cells[key];
        if (!cell) return;
        if (!props) delete sheet().cells[key];
        else {
          for (const p of props) delete cell[p];
          if (!Object.keys(cell).length) delete sheet().cells[key];
        }
        paint(r, c);
      });
      recalc();
      paintSel();
      onChange();
    }
    function restyle(fn) {
      if (locked(null, 'formatCells')) return;
      beforeChange();
      eachSel((r, c) => {
        const key = O.addr(r, c);
        const cell = sheet().cells[key] || (sheet().cells[key] = {});
        cell.s = cell.s || {};
        fn(cell.s, r, c);
        for (const k of Object.keys(cell.s)) if (cell.s[k] === undefined) delete cell.s[k];
        if (!Object.keys(cell.s).length) delete cell.s;
        if (!Object.keys(cell).length) delete sheet().cells[key];
        paint(r, c);
      });
      book.clearCache();
      paintSel();
      onChange();
    }
    function toggle(prop) {
      const cur = ((sheet().cells[O.addr(sel.r, sel.c)] || {}).s || {})[prop];
      restyle((s) => { s[prop] = cur ? undefined : true; });
    }
    function doUndo() {
      if (!undo.length) return;
      redo.push(snapshot());
      restore(undo.pop());
    }
    function doRedo() {
      if (!redo.length) return;
      undo.push(snapshot());
      restore(redo.pop());
    }

    // ---------- Portapapeles ----------
    wrap.addEventListener('copy', (e) => {
      if (editor) return;
      e.preventDefault();
      copyTo(e);
    });
    function copyTo(e) {
      e.clipboardData.setData('text/plain', takeClip());
    }
    function takeClip() {
      const lines = [], data = [], vals = [];
      for (let r = sel.r1; r <= Math.min(sel.r2, used().mr); r++) {
        const row = [], dRow = [], vRow = [];
        for (let c = sel.c1; c <= Math.min(sel.c2, used().mc); c++) {
          row.push(book.display(sheet(), r, c));
          dRow.push(sheet().cells[O.addr(r, c)] || null);
          const v = book.value(sheet(), r, c);
          vRow.push(v && v.err ? v.err : v);
        }
        lines.push(row.join('\t'));
        data.push(dRow);
        vals.push(vRow);
      }
      const text = lines.join('\n');
      clip = { text, data: JSON.parse(JSON.stringify(data)), vals, r: sel.r1, c: sel.c1, sheet: sheet().name };
      return text;
    }
    wrap.addEventListener('cut', (e) => {
      if (editor) return;
      e.preventDefault();
      copyTo(e);
      clearRange();
    });
    wrap.addEventListener('paste', (e) => {
      if (editor) return;
      e.preventDefault();
      pasteText(e.clipboardData.getData('text/plain'));
    });
    function pasteText(text) {
      const lines = clip && clip.text === text ? clip.data : text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
      if (locked({ r1: sel.r, c1: sel.c, r2: sel.r + Math.max(1, lines.length) - 1, c2: sel.c + Math.max(1, ...lines.map((l) => l.length)) - 1 })) return;
      beforeChange();
      if (clip && clip.text === text) {
        clip.data.forEach((row, i) => row.forEach((cell, j) => {
          const r = sel.r + i, c = sel.c + j, key = O.addr(r, c);
          if (!cell) { delete sheet().cells[key]; return; }
          const copy = JSON.parse(JSON.stringify(cell));
          if (copy.f !== undefined) copy.f = O.shiftFormula(copy.f, sel.r - clip.r, sel.c - clip.c);
          sheet().cells[key] = copy;
        }));
      } else {
        text.replace(/\r/g, '').replace(/\n$/, '').split('\n').forEach((line, i) => line.split('\t').forEach((v, j) => setCell(sel.r + i, sel.c + j, v)));
      }
      build();
      onChange();
    }

    // ---------- Filas y columnas ----------
    const fullCol = () => sel.r1 === 0 && sel.r2 === rows - 1;
    const fullRow = () => sel.c1 === 0 && sel.c2 === cols - 1;
    // n > 0 inserta y n < 0 elimina, desde la posicion at. Con band solo se desplaza esa franja de celdas.
    function structural(axis, at, n, band) {
      if (locked(null, band ? null : (n > 0 ? 'insert' : 'delete') + (axis === 'r' ? 'Rows' : 'Columns'))) return;
      beforeChange();
      const target = sheet();
      const opts = { axis, at, n, sheet: target.name, band };
      for (const sh of state.sheets) {
        for (const cell of Object.values(sh.cells)) {
          if (cell.f !== undefined) cell.f = O.adjustRefs(cell.f, { ...opts, local: sh === target });
        }
        // Los gráficos siguen a sus datos, y las validaciones de otras hojas que miran esta, también.
        for (const d of sh.drawings || []) if (d.refs) d.refs = d.refs.map((f) => O.adjustRefs(f, { ...opts, local: sh === target }));
        if (sh !== target && sh.dv) for (const d of sh.dv) if (d.f) d.f = d.f.map((f) => O.adjustRefs(f, { ...opts, local: false }));
      }
      // Como en Excel, los dibujos se mueven con las celdas que tienen debajo.
      if (!band && target.drawings) {
        const key = axis === 'r' ? 'r' : 'c';
        for (const d of target.drawings) {
          for (const p of [d.from, d.to].filter(Boolean)) {
            if (p[key] >= at) p[key] = n < 0 && p[key] < at - n ? at : p[key] + n;
          }
        }
      }
      if (!band && target.autoFilter) {
        const ref = O.adjustRefs(target.autoFilter.ref, { ...opts, local: true });
        if (ref.includes('#REF!')) delete target.autoFilter;
        else target.autoFilter.ref = ref;
      }
      // Los controles de formulario se mueven igual, y sus celdas vinculadas también.
      const newHeads = [];
      if (!band) {
        const key = axis === 'r' ? 'r' : 'c';
        for (const ctl of target.controls || []) {
          for (const p of [ctl.from, ctl.to].filter(Boolean)) if (p[key] >= at) p[key] = n < 0 && p[key] < at - n ? at : p[key] + n;
          for (const k of ['link', 'range']) {
            if (!ctl[k]) continue;
            const f = O.adjustRefs(ctl[k], { ...opts, local: true });
            if (!f.includes('#REF!')) ctl[k] = f;
          }
        }
        if (target.tables) {
          target.tables = target.tables.filter((t) => adjustTable(t, axis, at, n, opts, newHeads));
          if (!target.tables.length) delete target.tables;
        }
      }
      for (const k of Object.keys(state.names || {})) state.names[k] = O.adjustRefs(state.names[k], { ...opts, local: false });
      if (target.cf) target.cf = O.adjustRules(target.cf, opts);
      if (target.dv) target.dv = O.adjustRules(target.dv, opts);
      if (target.merges) {
        target.merges = target.merges.map((m) => O.adjustRefs(m, { ...opts, local: true })).filter((m) => {
          const x = !m.includes('#REF!') && O.parseRanges(m)[0];
          return x && (x.r1 !== x.r2 || x.c1 !== x.c2);
        });
        if (!target.merges.length) delete target.merges;
      }
      const remap =(i) => (i < at ? i : n < 0 && i < at - n ? null : i + n);
      const cells = {};
      for (const k of Object.keys(target.cells)) {
        const a = O.parseAddr(k);
        const across = axis === 'r' ? a.c : a.r;
        const i = band && (across < band[0] || across > band[1]) ? (axis === 'r' ? a.r : a.c) : remap(axis === 'r' ? a.r : a.c);
        if (i !== null) cells[axis === 'r' ? O.addr(i, a.c) : O.addr(a.r, i)] = target.cells[k];
      }
      target.cells = cells;
      // Una columna nueva dentro de una tabla se llama Columna1, Columna2..., como en Excel.
      for (const [addr, name] of newHeads) target.cells[addr] = { ...(target.cells[addr] || {}), v: name };
      if (!band) {
        const props = axis === 'r' ? ['rowH', 'hiddenRows', 'filtered'] : ['cols', 'hiddenCols'];
        for (const p of props) {
          if (!target[p]) continue;
          const moved = {};
          for (const k of Object.keys(target[p])) {
            const i = remap(+k);
            if (i !== null) moved[i] = target[p][k];
          }
          target[p] = moved;
        }
      }
      book.clearCache();
      build();
      onChange();
    }
    // Ajusta una tabla al insertar o eliminar filas o columnas. Devuelve false si se ha eliminado entera.
    function adjustTable(t, axis, at, n, opts, newHeads) {
      const ref = O.adjustRefs(t.ref, { ...opts, local: true });
      if (ref.includes('#REF!')) return false;
      const b = O.tableBox(t);
      const remapCols = (fn) => {
        if (!t.autoFilter) return;
        const out = {};
        for (const [k, v] of Object.entries(t.autoFilter.cols || {})) { const i = fn(+k); if (i !== null) out[i] = v; }
        t.autoFilter.cols = out;
      };
      if (axis === 'c') {
        const cols = t.columns || (t.columns = []);
        if (n > 0 && at > b.c1 && at <= b.c2) {
          const base = Math.max(0, ...cols.map((x) => x.id || 0));
          const taken = (name) => cols.some((x) => x.name.toLowerCase() === name.toLowerCase());
          const add = [];
          for (let k = 0, m = 1; k < n; k++) {
            while (taken('Columna' + m) || add.some((x) => x.name === 'Columna' + m)) m++;
            add.push({ id: base + k + 1, name: 'Columna' + m });
          }
          cols.splice(at - b.c1, 0, ...add);
          if (t.header !== false) add.forEach((x, k) => newHeads.push([O.addr(b.r1, at + k), x.name]));
          remapCols((i) => (i >= at - b.c1 ? i + n : i));
        } else if (n < 0) {
          const from = Math.max(at, b.c1), to = Math.min(at - n - 1, b.c2);
          if (from <= to) {
            cols.splice(from - b.c1, to - from + 1);
            remapCols((i) => (i < from - b.c1 ? i : i <= to - b.c1 ? null : i - (to - from + 1)));
          }
        }
      }
      t.ref = ref;
      return true;
    }
    // Lo que hace Excel al escribir en una tabla o junto a ella: cambiar el nombre de una columna desde el encabezado,
    // ampliar la tabla con la fila o columna que se escribe al lado y rellenar la columna con la fórmula escrita.
    function tableAfterEdit(r, c, before) {
      const s = sheet();
      const cell = s.cells[O.addr(r, c)];
      const filled = !!cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== ''));
      let t = O.tableAt(s, r, c);
      if (!t && filled) {
        t = (s.tables || []).find((x) => {
          const b = O.tableBox(x);
          return (!x.totals && r === b.r2 + 1 && c >= b.c1 && c <= b.c2) || (c === b.c2 + 1 && r >= b.r1 && r <= b.r2);
        });
        if (!t) return;
        const b = O.tableBox(t);
        if (c === b.c2 + 1) {
          const cols = t.columns || (t.columns = []);
          let m = cols.length + 1;
          while (cols.some((x) => x.name.toLowerCase() === ('Columna' + m).toLowerCase())) m++;
          const name = t.header !== false && r === b.r1 ? book.display(s, r, c) : 'Columna' + m;
          cols.push({ id: Math.max(0, ...cols.map((x) => x.id || 0)) + 1, name });
          if (t.header !== false && r !== b.r1) s.cells[O.addr(b.r1, c)] = { ...(s.cells[O.addr(b.r1, c)] || {}), v: name };
          t.ref = `${O.addr(b.r1, b.c1)}:${O.addr(b.r2, c)}`;
        } else {
          t.ref = `${O.addr(b.r1, b.c1)}:${O.addr(r, b.c2)}`;
          // La fila nueva recibe las fórmulas de las columnas calculadas.
          (t.columns || []).forEach((col, k) => {
            const key = O.addr(r, b.c1 + k);
            if (col.calc && !s.cells[key]) s.cells[key] = { f: col.calc };
          });
        }
        return;
      }
      if (!t) return;
      const b = O.tableBox(t), d = O.tableData(t), k = c - b.c1;
      const col = (t.columns || [])[k];
      if (!col) return;
      if (t.header !== false && r === b.r1) {
        const name = book.display(s, r, c);
        if (!name || name === col.name) return;
        // Las fórmulas que nombran la columna pasan a usar el nombre nuevo.
        for (const sh of state.sheets) {
          for (const [addr, x] of Object.entries(sh.cells)) {
            if (x.f === undefined) continue;
            const a = O.parseAddr(addr), own = O.tableAt(sh, a.r, a.c);
            x.f = O.renameTableColumn(x.f, t.name, col.name, name, own && own.name);
          }
        }
        for (const x of t.columns) if (x.calc) x.calc = O.renameTableColumn(x.calc, t.name, col.name, name, t.name);
        col.name = name;
        return;
      }
      if (r < d.r1 || r > d.r2 || !cell || cell.f === undefined || cell.f === before) return;
      let others = true;
      for (let rr = d.r1; rr <= d.r2; rr++) {
        const x = rr === r ? null : s.cells[O.addr(rr, c)];
        if (x && (x.f !== undefined || (x.v !== undefined && x.v !== null && x.v !== '')) && !(col.calc && x.f === col.calc)) others = false;
      }
      if (!others && !col.calc) return;
      // Columna calculada: la misma fórmula en todas las filas de datos.
      for (let rr = d.r1; rr <= d.r2; rr++) {
        if (rr === r) continue;
        const key = O.addr(rr, c);
        s.cells[key] = { ...(s.cells[key] || {}), f: O.shiftFormula(cell.f, rr - r, 0) };
        delete s.cells[key].v;
      }
      col.calc = cell.f;
    }
    function setHidden(axis, on) {
      const prop = axis === 'r' ? 'hiddenRows' : 'hiddenCols';
      const s = sheet();
      const lo = axis === 'r' ? sel.r1 : sel.c1, hi = axis === 'r' ? sel.r2 : sel.c2;
      beforeChange();
      s[prop] = s[prop] || {};
      if (on) for (let i = lo; i <= hi; i++) s[prop][i] = true;
      else {
        // Mostrar tambien alcanza las ocultas pegadas a la seleccion, que no se pueden seleccionar.
        let from = lo, to = hi;
        while (from > 0 && s[prop][from - 1]) from--;
        while (s[prop][to + 1]) to++;
        for (let i = from; i <= to; i++) {
          delete s[prop][i];
          if (axis === 'r' && s.filtered) delete s.filtered[i];
        }
      }
      build();
      onChange();
    }
    function askHeight() {
      const cur = (sheet().rowH && sheet().rowH[sel.r]) || sheet().defH || DEFAULT_H;
      const raw = window.prompt('Alto de fila en píxeles', cur);
      if (raw === null) return;
      const h = parseInt(raw, 10);
      if (!(h >= 12 && h <= 500)) { O.notice('El alto debe estar entre 12 y 500 píxeles.'); return; }
      if (locked(null, 'formatRows')) return;
      beforeChange();
      const s = sheet();
      s.rowH = s.rowH || {};
      for (let r = sel.r1; r <= sel.r2; r++) s.rowH[r] = h;
      build();
      onChange();
    }
    function setWidths(fn) {
      if (locked(null, 'formatColumns')) return;
      beforeChange();
      const s = sheet();
      s.cols = s.cols || {};
      for (let c = sel.c1; c <= sel.c2; c++) {
        const w = fn(c);
        if (w) s.cols[c] = w;
      }
      build();
      onChange();
    }
    function askWidth() {
      const cur = (sheet().cols && sheet().cols[sel.c]) || sheet().defW || DEFAULT_W;
      const raw = window.prompt('Ancho de columna en píxeles', cur);
      if (raw === null) return;
      const w = parseInt(raw, 10);
      if (!(w >= 24 && w <= 1000)) { O.notice('El ancho debe estar entre 24 y 1000 píxeles.'); return; }
      setWidths(() => w);
    }
    function fitWidth() {
      const s = sheet();
      const ctx = document.createElement('canvas').getContext('2d');
      const base = getComputedStyle(tds[sel.r][sel.c]);
      const widest = {};
      for (const k of Object.keys(s.cells)) {
        const a = O.parseAddr(k);
        if (a.c < sel.c1 || a.c > sel.c2) continue;
        const st = s.cells[k].s || {};
        ctx.font = `${st.i ? 'italic ' : ''}${st.b ? '700 ' : ''}${base.fontSize} ${base.fontFamily}`;
        widest[a.c] = Math.max(widest[a.c] || 0, ctx.measureText(book.display(s, a.r, a.c)).width);
      }
      setWidths((c) => (widest[c] ? Math.min(1000, Math.max(24, Math.ceil(widest[c]) + 14)) : DEFAULT_W));
    }
    const plural = (n, one, many) => (n === 1 ? one : `${n} ${many}`);
    const rowN = () => sel.r2 - sel.r1 + 1;
    const colN = () => sel.c2 - sel.c1 + 1;
    const insertRowItems = () => [
      { label: `${plural(rowN(), 'Fila', 'filas')} encima`, action: () => structural('r', sel.r1, rowN()) },
      { label: `${plural(rowN(), 'Fila', 'filas')} debajo`, action: () => structural('r', sel.r2 + 1, rowN()) },
    ];
    const insertColItems = () => [
      { label: `${plural(colN(), 'Columna', 'columnas')} a la izquierda`, action: () => structural('c', sel.c1, colN()) },
      { label: `${plural(colN(), 'Columna', 'columnas')} a la derecha`, action: () => structural('c', sel.c2 + 1, colN()) },
    ];
    const deleteRowItem = () => ({ label: plural(rowN(), 'Toda la fila', 'filas'), action: () => structural('r', sel.r1, -rowN()) });
    const deleteColItem = () => ({ label: plural(colN(), 'Toda la columna', 'columnas'), action: () => structural('c', sel.c1, -colN()) });
    const widthItems = () => [
      { label: 'Ajustar ancho al contenido', action: fitWidth },
      { label: 'Ancho de columna…', action: askWidth },
    ];
    const hideItems = (axis) => [
      { label: 'Ocultar', action: () => setHidden(axis, true) },
      { label: 'Mostrar', action: () => setHidden(axis, false) },
    ];

    // ---------- Pegado especial ----------
    function pasteSpecial(mode) {
      if (!clip || locked()) return;
      beforeChange();
      const put = (i, j, r, c) => {
        const src = clip.data[i][j];
        const key = O.addr(r, c);
        const prev = sheet().cells[key] || {};
        const next = {};
        if (prev.note) next.note = prev.note;
        if (prev.link) next.link = prev.link;
        const moved = () => (src && src.f !== undefined ? { f: O.shiftFormula(src.f, r - (clip.r + i), c - (clip.c + j)) } : src && src.v !== undefined ? { v: src.v } : {});
        const value = () => {
          const v = clip.vals[i][j];
          return v === null || v === undefined ? {} : { v };
        };
        const srcStyle = src && src.s ? { s: { ...src.s } } : {};
        const keepStyle = prev.s ? { s: prev.s } : {};
        if (mode === 'values') Object.assign(next, keepStyle, value());
        else if (mode === 'valuesFormat') Object.assign(next, srcStyle, value());
        else if (mode === 'formulas') Object.assign(next, keepStyle, moved());
        else if (mode === 'format') {
          if (prev.f !== undefined) next.f = prev.f;
          else if (prev.v !== undefined) next.v = prev.v;
          Object.assign(next, srcStyle);
        } else if (mode === 'transpose') Object.assign(next, srcStyle, moved());
        else if (mode === 'link') {
          const same = clip.sheet === sheet().name;
          next.f = (same ? '' : qname(clip.sheet) + '!') + '$' + O.colName(clip.c + j) + '$' + (clip.r + i + 1);
          Object.assign(next, keepStyle);
        }
        if (Object.keys(next).length) sheet().cells[key] = next;
        else delete sheet().cells[key];
      };
      clip.data.forEach((row, i) => row.forEach((_, j) => {
        if (mode === 'transpose') put(i, j, sel.r + j, sel.c + i);
        else put(i, j, sel.r + i, sel.c + j);
      }));
      book.clearCache();
      build();
      onChange();
    }
    const qname = (name) => (/^[A-Za-z_À-ɏ][\w.À-ɏ]*$/.test(name) ? name : `'${name}'`);
    const pasteSpecialItems = () => [
      ['values', 'Valores'], ['valuesFormat', 'Valores y formato'], ['formulas', 'Fórmulas'],
      ['format', 'Formato'], ['transpose', 'Transponer'], ['link', 'Pegar vínculo'],
    ].map(([mode, label]) => ({ label, action: () => pasteSpecial(mode) }));

    // ---------- Ordenar y filtrar ----------
    // Con una sola celda seleccionada se usa el bloque continuo de datos que la rodea, como en Excel.
    function dataRegion() {
      const s = sheet();
      const { mr, mc } = used();
      if (sel.r1 !== sel.r2 || sel.c1 !== sel.c2) return { r1: sel.r1, c1: sel.c1, r2: Math.min(sel.r2, mr), c2: Math.min(sel.c2, mc) };
      const reg = { r1: sel.r, r2: sel.r, c1: sel.c, c2: sel.c };
      const filled = (r, c) => {
        const cell = s.cells[O.addr(r, c)];
        return !!cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== ''));
      };
      const rowHas = (r) => r >= 0 && r <= mr && Array.from({ length: reg.c2 - reg.c1 + 3 }, (_, i) => reg.c1 - 1 + i).some((c) => c >= 0 && filled(r, c));
      const colHas = (c) => c >= 0 && c <= mc && Array.from({ length: reg.r2 - reg.r1 + 3 }, (_, i) => reg.r1 - 1 + i).some((r) => r >= 0 && filled(r, c));
      for (let grew = true; grew;) {
        grew = false;
        if (rowHas(reg.r1 - 1)) { reg.r1--; grew = true; }
        if (rowHas(reg.r2 + 1)) { reg.r2++; grew = true; }
        if (colHas(reg.c1 - 1)) { reg.c1--; grew = true; }
        if (colHas(reg.c2 + 1)) { reg.c2++; grew = true; }
      }
      return reg;
    }
    // Hay encabezado si la primera fila es solo texto y la siguiente no lo es, o si va en negrita.
    function hasHeader(reg) {
      if (reg.r1 === reg.r2) return false;
      const s = sheet();
      let bold = false, otherType = false;
      for (let c = reg.c1; c <= reg.c2; c++) {
        const first = book.value(s, reg.r1, c), next = book.value(s, reg.r1 + 1, c);
        if (first !== null && typeof first !== 'string') return false;
        if (((s.cells[O.addr(reg.r1, c)] || {}).s || {}).b) bold = true;
        if (typeof first === 'string' && next !== null && typeof next !== 'string') otherType = true;
      }
      return bold || otherType;
    }
    // opts: { reg, start, col } para ordenar un rango concreto, como el del autofiltro, por una columna dada.
    function sortRegion(compare, opts = {}) {
      if (locked(null, 'sort')) return;
      const reg = opts.reg || dataRegion();
      const key = opts.col !== undefined ? opts.col : sel.c;
      const s = sheet();
      // Excel no ordena un rango con celdas combinadas.
      if (merges().some((m) => !(m.r1 > reg.r2 || m.r2 < reg.r1 || m.c1 > reg.c2 || m.c2 < reg.c1))) {
        O.notice('Para hacer esto, todas las celdas combinadas deben tener el mismo tamaño. Separa las celdas combinadas antes de ordenar.');
        return;
      }
      const start = opts.start !== undefined ? opts.start : reg.r1 + (hasHeader(reg) ? 1 : 0);
      const lines = [];
      for (let r = start; r <= reg.r2; r++) {
        // Con un filtro puesto, Excel ordena solo las filas visibles y deja las ocultas en su sitio.
        if (s.filtered && s.filtered[r]) continue;
        const cells = {};
        for (let c = reg.c1; c <= reg.c2; c++) cells[c] = s.cells[O.addr(r, c)];
        lines.push({ r, cells, key: book.value(s, r, key), raw: s.cells[O.addr(r, key)] });
      }
      const slots = lines.map((l) => l.r);
      const order = lines.map((l, i) => [l, i]).sort((a, b) => compare(a[0], b[0]) || a[1] - b[1]).map(([l]) => l);
      beforeChange();
      order.forEach((line, i) => {
        const r = slots[i];
        for (let c = reg.c1; c <= reg.c2; c++) {
          const cell = line.cells[c];
          if (!cell) { delete s.cells[O.addr(r, c)]; continue; }
          s.cells[O.addr(r, c)] = cell.f !== undefined ? { ...cell, f: O.shiftFormula(cell.f, r - line.r, 0) } : cell;
        }
      });
      book.clearCache();
      build();
      onChange();
    }
    const rank = (v) => (v === null || (v && v.err) ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
    const byValue = (dir) => (a, b) => {
      if (rank(a.key) === 3 || rank(b.key) === 3) return rank(a.key) - rank(b.key);
      if (rank(a.key) !== rank(b.key)) return (rank(a.key) - rank(b.key)) * dir;
      if (typeof a.key === 'string') return a.key.localeCompare(b.key, 'es', { sensitivity: 'base' }) * dir;
      return (a.key - b.key) * dir;
    };
    const activeStyle = () => (sheet().cells[O.addr(sel.r, sel.c)] || {}).s || {};
    const byStyle = (prop) => {
      const wanted = activeStyle()[prop];
      const score = (l) => (((l.raw || {}).s || {})[prop] === wanted ? 0 : 1);
      return (a, b) => score(a) - score(b);
    };
    const sortItems = () => [
      { label: 'De A a Z / de menor a mayor', action: () => sortRegion(byValue(1)) },
      { label: 'De Z a A / de mayor a menor', action: () => sortRegion(byValue(-1)) },
      { label: 'Poner este color de celda primero', action: () => sortRegion(byStyle('fill')) },
      { label: 'Poner este color de fuente primero', action: () => sortRegion(byStyle('color')) },
    ];
    // ---------- Gráficos, imágenes y formas ----------
    // Cada dibujo va anclado a celdas (from/to con desplazamiento en píxeles) o a una celda con tamaño fijo (size).
    const drawLayer = document.createElement('div');
    drawLayer.className = 'dr-layer';
    let picked = -1;
    const themeColors = () => state.themeColors || O.DEFAULT_THEME;
    const colLeft = (c) => {
      if (c < cols) return colHeads[c].offsetLeft;
      return colHeads[cols - 1].offsetLeft + colHeads[cols - 1].offsetWidth + (c - cols) * DEFAULT_W;
    };
    // Una fila oculta no ocupa sitio: empieza donde empieza la siguiente visible.
    const rowTop = (r) => {
      const last = rowHeads[rows - 1];
      if (r < rows) for (let k = r; k < rows; k++) if (!(sheet().hiddenRows && sheet().hiddenRows[k])) return rowHeads[k].offsetTop;
      return last.offsetTop + last.offsetHeight + Math.max(0, r - rows) * DEFAULT_H;
    };
    const pointOf = (p) => ({ x: colLeft(p.c) + (p.dx || 0), y: rowTop(p.r) + (p.dy || 0) });
    // Punto en píxeles de la tabla a ancla de celda.
    function anchorAt(x, y) {
      let c = 0;
      while (c < cols - 1 && colLeft(c + 1) <= x) c++;
      let r = 0;
      while (r < rows - 1 && rowTop(r + 1) <= y) r++;
      return { c, r, dx: Math.max(0, Math.round(x - colLeft(c))), dy: Math.max(0, Math.round(y - rowTop(r))) };
    }
    function boxOf(d) {
      const a = pointOf(d.from);
      if (d.to) {
        const b = pointOf(d.to);
        return { x: a.x, y: a.y, w: Math.max(4, b.x - a.x), h: Math.max(4, b.y - a.y) };
      }
      return { x: a.x, y: a.y, w: (d.size || {}).w || 100, h: (d.size || {}).h || 100 };
    }
    const chartOf = (d) => {
      const m = O.chartModel(d.xml, themeColors());
      const data = O.chartData(book, state.sheets, m, d.refs || []);
      const title = O.chartTitle(book, state.sheets, m, d.refs || [], data);
      return { m: { ...m, title: title ? { text: title } : null }, data };
    };
    function renderDrawings() {
      const list = sheet().drawings || [];
      const ctls = sheet().controls || [];
      if (!list.length && !ctls.length) { drawLayer.remove(); picked = -1; return; }
      if (drawLayer.parentNode !== wrap) wrap.append(drawLayer);
      drawLayer.replaceChildren();
      if (picked >= list.length) picked = -1;
      list.forEach((d, i) => {
        // Las copias ocultas de los controles de formulario no se ven: el control se dibuja aparte.
        if (d.kind === 'raw' && !d.box) return;
        const b = boxOf(d);
        const el = document.createElement('div');
        el.className = 'dr-item dr-' + d.kind + (i === picked ? ' picked' : '');
        el.dataset.index = i;
        el.style.cssText = `left:${b.x}px;top:${b.y}px;width:${b.w}px;height:${b.h}px`;
        el.title = d.name || '';
        if (d.kind === 'slicer') el.append(slicerView(d));
        else if (d.kind === 'raw') el.textContent = d.name || 'Objeto';
        else if (d.kind === 'chart') {
          try {
            const { m, data } = chartOf(d);
            el.innerHTML = O.chartSvg(m, data, Math.round(b.w), Math.round(b.h), themeColors());
            el.setAttribute('role', 'img');
            el.setAttribute('aria-label', 'Gráfico' + (m.title ? ': ' + m.title.text : ''));
          } catch (e) {
            console.error(e);
            el.textContent = 'Gráfico';
          }
        } else if (d.kind === 'img') {
          const img = document.createElement('img');
          img.src = (state.media || {})[d.media] || '';
          img.alt = d.name || 'Imagen';
          img.draggable = false;
          el.append(img);
        } else {
          el.textContent = d.text || '';
          el.style.background = d.fill ? '#' + d.fill : 'transparent';
          el.style.border = d.line ? '1px solid #' + d.line : 'none';
          if (d.geom === 'ellipse') el.style.borderRadius = '50%';
          else if (/round/i.test(d.geom || '')) el.style.borderRadius = '10px';
        }
        if (i === picked) for (const k of ['nw', 'ne', 'sw', 'se']) {
          const h = document.createElement('span');
          h.className = 'dr-handle ' + k;
          h.dataset.handle = k;
          el.append(h);
        }
        drawLayer.append(el);
      });
      ctls.forEach((ctl) => { if (ctl.from && ctl.to) drawLayer.append(controlView(ctl)); });
    }

    // ---------- Segmentaciones ----------
    // La de una tabla muestra los valores de su columna y filtra la tabla al pulsarlos, como en Excel.
    const tableById = (id) => {
      for (const s of state.sheets) for (const t of s.tables || []) if (t.id === id) return { sheet: s, table: t };
      return null;
    };
    function slicerItems(d) {
      const sl = d.slicer || {};
      if (sl.items) return { items: sl.items.map((x) => ({ text: x.name, on: x.on })) };
      const hit = sl.table !== undefined && tableById(sl.table);
      if (!hit) return { items: [] };
      const t = hit.table, k = (t.columns || []).findIndex((x) => x.id === sl.column);
      if (k < 0) return { items: [] };
      const b = O.tableBox(t), dd = O.tableData(t), c = b.c1 + k;
      const spec = (t.autoFilter && t.autoFilter.cols[k]) || null;
      // Un valor cuyas filas oculta el filtro de otra columna queda "sin datos", al final y más claro.
      const af = t.autoFilter, others = af ? Object.entries(af.cols || {}).filter(([id]) => +id !== k) : [];
      const fbox = { ...b, r2: b.r2 - (t.totals || 0) };
      const seen = new Map();
      for (let r = dd.r1; r <= dd.r2; r++) {
        const text = book.display(hit.sheet, r, c);
        const key = text.toLowerCase();
        if (!seen.has(key)) seen.set(key, { text: text || '(en blanco)', raw: text, v: book.value(hit.sheet, r, c), data: false });
        if (!seen.get(key).data && others.every(([id, sp]) => matches(sp, r, b.c1 + +id, fbox))) seen.get(key).data = true;
      }
      const items = [...seen.values()].sort((a, x) => (a.data !== x.data ? (a.data ? -1 : 1) : slicerOrder(a, x)));
      for (const it of items) it.on = !spec || !spec.vals ? true : it.raw === '' ? !!spec.blank : spec.vals.some((v) => String(v).toLowerCase() === it.raw.toLowerCase());
      return { items, table: hit, col: c, spec };
    }
    // Excel ordena de la A a la Z usando las listas personalizadas: los meses y los días van en su orden.
    const CUSTOM_LISTS = [
      ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'],
      ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'],
      ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo'],
      ['lun', 'mar', 'mié', 'jue', 'vie', 'sáb', 'dom'],
    ];
    function slicerOrder(a, x) {
      if (typeof a.v === 'number' && typeof x.v === 'number') return a.v - x.v;
      for (const list of CUSTOM_LISTS) {
        const i = list.indexOf(a.raw.toLowerCase()), j = list.indexOf(x.raw.toLowerCase());
        if (i >= 0 && j >= 0) return i - j;
      }
      if (a.raw === '' || x.raw === '') return a.raw === '' ? 1 : -1;
      return a.text.localeCompare(x.text, 'es', { sensitivity: 'base' });
    }
    function slicerView(d) {
      const box = document.createElement('div');
      box.className = 'sl-box';
      const accent = themeColors()[4] || '4472C4';
      box.style.setProperty('--sl-on', '#' + O.applyTint(accent, 0.6));
      box.style.setProperty('--sl-line', '#' + O.applyTint(accent, 0.2));
      box.style.setProperty('--sl-pale', '#' + O.applyTint(accent, 0.85));
      const info = slicerItems(d);
      const head = document.createElement('div');
      head.className = 'sl-head';
      head.textContent = (d.slicer && d.slicer.caption) || d.name || 'Segmentación';
      const clear = document.createElement('button');
      clear.type = 'button';
      clear.className = 'sl-clear';
      clear.title = 'Borrar filtro';
      clear.innerHTML = FUNNEL + '<span aria-hidden="true">×</span>';
      clear.disabled = !info.items.some((x) => !x.on);
      clear.dataset.slicer = 'clear';
      head.append(clear);
      const list = document.createElement('div');
      list.className = 'sl-list';
      if (d.slicer && d.slicer.columns > 1) list.style.gridTemplateColumns = `repeat(${d.slicer.columns}, 1fr)`;
      info.items.forEach((it, n) => {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'sl-item' + (it.on ? ' on' : '') + (it.data === false ? ' nodata' : '');
        b.textContent = it.text;
        b.dataset.slicer = String(n);
        b.setAttribute('aria-pressed', it.on ? 'true' : 'false');
        list.append(b);
      });
      box.append(head, list);
      return box;
    }
    // Pulsar un valor deja solo ese; con Ctrl se añade o se quita, como en Excel.
    function slicerClick(d, which, toggle) {
      const info = slicerItems(d);
      if (!info.table && d.slicer && d.slicer.items && d.slicer.items.length) {
        // Segmentación de tabla dinámica: se recalculan las tablas que conecta desde su caché.
        let picks = d.slicer.items.map((x) => x.on);
        if (which === 'clear') picks = picks.map(() => true);
        else if (toggle) picks[which] = !picks[which];
        else picks = picks.map((_, k) => k === which);
        if (!picks.some(Boolean)) return;
        beforeChange();
        d.slicer.items.forEach((x, k) => { x.on = picks[k]; });
        const res = O.pivotSlicer(state, d.slicer);
        if (res.left) O.notice('Oficina no sabe recalcular esta tabla dinámica: el filtro se guarda y Excel la actualizará al abrir el archivo.');
        build();
        onChange();
        return;
      }
      if (!info.table) return;
      if (info.table.sheet !== sheet()) return;
      const t = info.table.table;
      if (!t.autoFilter) t.autoFilter = { cols: {} };
      const owner = filterOwners().find((o) => o.table === t);
      if (!owner) return;
      let on = info.items.map((x) => x.on);
      if (which === 'clear') on = on.map(() => true);
      else if (toggle) on[which] = !on[which];
      else on = on.map((_, k) => k === which);
      if (!on.some(Boolean)) return;
      const chosen = info.items.filter((_, k) => on[k]);
      if (chosen.length === info.items.length) setFilter(owner, info.col, null);
      else setFilter(owner, info.col, { vals: chosen.filter((x) => x.raw !== '').map((x) => x.raw), ...(chosen.some((x) => x.raw === '') ? { blank: true } : {}) });
    }

    // ---------- Controles de formulario ----------
    // Casillas, botones de opción, listas desplegables y controles de número: cambian su celda vinculada, como en Excel.
    const linkTarget = (ref) => {
      const m = /^(?:'?(.+?)'?!)?\$?([A-Z]+)\$?(\d+)$/i.exec(String(ref || '').trim());
      if (!m) return null;
      const s = m[1] ? state.sheets.find((x) => x.name.toLowerCase() === m[1].toLowerCase()) : sheet();
      return s ? { sheet: s, r: +m[3] - 1, c: O.colIndex(m[2]) } : null;
    };
    const linkValue = (ctl) => {
      const t = linkTarget(ctl.link);
      return t ? book.value(t.sheet, t.r, t.c) : undefined;
    };
    const radios = () => (sheet().controls || []).filter((x) => x.type === 'Radio' && !x.activeX);
    const radioLink = () => { const r = radios().find((x) => x.link); return r && r.link; };
    // Un ActiveX escribe en su celda su propio valor: VERDADERO o FALSO, el texto elegido o el número.
    function activeXState(ctl) {
      const v = ctl.link ? linkValue(ctl) : undefined;
      const has = v !== undefined && v !== null && !(v && v.err);
      if (/^(CheckBox|Radio|Toggle)$/.test(ctl.type)) return has ? v === true || (typeof v === 'number' && v !== 0) : /^(1|-1|true|verdadero)$/i.test(ctl.value || '');
      if (/^(Spin|Scroll)$/.test(ctl.type)) return has && typeof v === 'number' ? v : +ctl.value || 0;
      return has ? O.formatValue(v) : ctl.value || '';
    }
    function setActiveX(ctl, value) {
      beforeChange();
      const writes = [[ctl, value]];
      // Los botones de opción del mismo grupo se desmarcan, como en Excel.
      if (ctl.type === 'Radio' && value) {
        for (const x of sheet().controls || []) if (x !== ctl && x.activeX && x.type === 'Radio' && (x.group || '') === (ctl.group || '')) writes.push([x, false]);
      }
      for (const [x, v] of writes) {
        x.value = typeof v === 'boolean' ? (v ? '1' : '0') : String(v);
        const t = linkTarget(x.link);
        if (!t) continue;
        const key = O.addr(t.r, t.c);
        t.sheet.cells[key] = { ...(t.sheet.cells[key] || {}), v };
        delete t.sheet.cells[key].f;
      }
      paintAll();
      onChange();
    }
    function activeXView(ctl, el) {
      const st = activeXState(ctl);
      const label = ctl.label || ctl.name;
      el.classList.add('fc-ax');
      el.addEventListener('keydown', (e) => e.stopPropagation());
      if (ctl.type === 'CheckBox' || ctl.type === 'Radio') {
        const l = document.createElement('label');
        const input = document.createElement('input');
        input.type = ctl.type === 'CheckBox' ? 'checkbox' : 'radio';
        input.checked = !!st;
        input.onchange = () => setActiveX(ctl, ctl.type === 'CheckBox' ? input.checked : true);
        l.append(input, document.createTextNode(' ' + label));
        el.append(l);
      } else if (ctl.type === 'Toggle' || ctl.type === 'Button') {
        const b = document.createElement('button');
        b.type = 'button';
        b.textContent = label;
        if (ctl.type === 'Toggle') {
          b.setAttribute('aria-pressed', st ? 'true' : 'false');
          b.onclick = () => setActiveX(ctl, !st);
        } else b.onclick = () => O.notice('Este botón ejecuta una macro. Oficina conserva las macros pero no las ejecuta.');
        el.append(b);
      } else if (ctl.type === 'Drop' || ctl.type === 'List') {
        const s2 = document.createElement('select');
        s2.setAttribute('aria-label', ctl.name);
        if (ctl.type === 'List') s2.size = 3;
        const opts = ctl.range ? book.evaluateAll(sheet(), ctl.range.replace(/^=/, '')).map((v) => (v && v.err ? O.localError(v.err) : O.formatValue(v))) : [];
        if (!opts.includes(st)) s2.append(new Option(st, st));
        opts.forEach((text) => s2.append(new Option(text, text)));
        s2.value = st;
        s2.onchange = () => setActiveX(ctl, s2.value);
        el.append(s2);
      } else if (ctl.type === 'Edit') {
        const input = document.createElement('input');
        input.type = 'text';
        input.value = st;
        input.setAttribute('aria-label', ctl.name);
        input.onchange = () => setActiveX(ctl, input.value);
        el.append(input);
      } else if (ctl.type === 'Spin' || ctl.type === 'Scroll') {
        const lo = ctl.min || 0, hi = ctl.max !== undefined ? ctl.max : 100, inc = ctl.inc || 1;
        if (ctl.type === 'Spin') {
          for (const [dir, txt] of [[1, '▲'], [-1, '▼']]) {
            const b = document.createElement('button');
            b.type = 'button';
            b.textContent = txt;
            b.setAttribute('aria-label', dir > 0 ? 'Subir' : 'Bajar');
            b.onclick = () => setActiveX(ctl, Math.max(lo, Math.min(hi, st + dir * inc)));
            el.append(b);
          }
        } else {
          const r2 = document.createElement('input');
          Object.assign(r2, { type: 'range', min: lo, max: hi, step: inc, value: st });
          r2.setAttribute('aria-label', ctl.name);
          r2.onchange = () => setActiveX(ctl, +r2.value);
          el.append(r2);
        }
      } else {
        el.textContent = ctl.type === 'Label' ? label : ctl.name;
        if (ctl.type !== 'Label') el.classList.add('fc-group');
      }
      return el;
    }
    function controlState(ctl) {
      if (ctl.activeX) return activeXState(ctl);
      const v = ctl.type === 'Radio' ? linkValue({ link: radioLink() }) : linkValue(ctl);
      const has = v !== undefined && v !== null && !(v && v.err);
      if (ctl.type === 'CheckBox') return has ? v === true || (typeof v === 'number' && v !== 0) : !!ctl.checked;
      if (ctl.type === 'Radio') return has ? radios().indexOf(ctl) + 1 === v : !!ctl.checked;
      if (/^(Drop|List)$/.test(ctl.type)) return has && typeof v === 'number' ? v : ctl.sel || 0;
      if (/^(Spin|Scroll)$/.test(ctl.type)) return has && typeof v === 'number' ? v : ctl.val || 0;
      return null;
    }
    function setControl(ctl, value) {
      beforeChange();
      if (ctl.type === 'Radio') {
        for (const x of radios()) x.checked = x === ctl;
        value = radios().indexOf(ctl) + 1;
      } else if (ctl.type === 'CheckBox') ctl.checked = value;
      else if (/^(Drop|List)$/.test(ctl.type)) ctl.sel = value;
      else ctl.val = value;
      const t = linkTarget(ctl.type === 'Radio' ? radioLink() : ctl.link);
      if (t) {
        const key = O.addr(t.r, t.c);
        const prev = t.sheet.cells[key] || {};
        t.sheet.cells[key] = { ...prev, v: value };
        delete t.sheet.cells[key].f;
      }
      paintAll();
      onChange();
    }
    function controlView(ctl) {
      const a = pointOf(ctl.from), z = pointOf(ctl.to);
      const el = document.createElement('div');
      el.className = 'fc-item fc-' + ctl.type.toLowerCase();
      el.style.cssText = `left:${a.x}px;top:${a.y}px;width:${Math.max(8, z.x - a.x)}px;height:${Math.max(8, z.y - a.y)}px`;
      el.title = ctl.name;
      el.addEventListener('mousedown', (e) => e.stopPropagation());
      if (ctl.activeX) return activeXView(ctl, el);
      const st = controlState(ctl);
      const label = ctl.label || '';
      if (ctl.type === 'CheckBox' || ctl.type === 'Radio') {
        const l = document.createElement('label');
        const input = document.createElement('input');
        input.type = ctl.type === 'CheckBox' ? 'checkbox' : 'radio';
        input.checked = !!st;
        input.onchange = () => setControl(ctl, ctl.type === 'CheckBox' ? input.checked : true);
        l.append(input, document.createTextNode(' ' + label));
        el.append(l);
      } else if (ctl.type === 'Drop' || ctl.type === 'List') {
        const s2 = document.createElement('select');
        s2.setAttribute('aria-label', ctl.name);
        if (ctl.type === 'List') s2.size = 3;
        const opts = ctl.range ? book.evaluateAll(sheet(), ctl.range.replace(/^=/, '')).map((v) => (v && v.err ? O.localError(v.err) : O.formatValue(v))) : [];
        s2.append(new Option('', '0'));
        opts.forEach((text, k) => s2.append(new Option(text, String(k + 1))));
        s2.value = String(st || 0);
        s2.onchange = () => setControl(ctl, +s2.value);
        el.append(s2);
      } else if (ctl.type === 'Spin') {
        const step = (dir) => {
          const lo = ctl.min || 0, hi = ctl.max !== undefined ? ctl.max : 100, inc = ctl.inc || 1;
          setControl(ctl, Math.max(lo, Math.min(hi, (controlState(ctl) || 0) + dir * inc)));
        };
        for (const [dir, txt] of [[1, '▲'], [-1, '▼']]) {
          const b = document.createElement('button');
          b.type = 'button';
          b.textContent = txt;
          b.setAttribute('aria-label', dir > 0 ? 'Subir' : 'Bajar');
          b.onclick = () => step(dir);
          el.append(b);
        }
      } else if (ctl.type === 'Scroll') {
        const r2 = document.createElement('input');
        r2.type = 'range';
        r2.min = ctl.min || 0;
        r2.max = ctl.max !== undefined ? ctl.max : 100;
        r2.step = ctl.inc || 1;
        r2.value = st || 0;
        r2.setAttribute('aria-label', ctl.name);
        r2.onchange = () => setControl(ctl, +r2.value);
        el.append(r2);
      } else if (ctl.type === 'Button') {
        const b = document.createElement('button');
        b.type = 'button';
        b.textContent = label || ctl.name;
        b.onclick = () => O.notice('Este botón ejecuta una macro. Oficina conserva las macros pero no las ejecuta.');
        el.append(b);
      } else {
        el.textContent = label;
        if (ctl.type === 'GBox') el.classList.add('fc-group');
      }
      return el;
    }
    const objectsLocked = () => {
      if (!sheet().protected || (sheet().protect && sheet().protect.objects === '0')) return false;
      O.notice('Los gráficos y los objetos de esta hoja están protegidos. Desprotege la hoja para cambiarlos.');
      return true;
    };
    function pickDrawing(i) {
      picked = i;
      renderDrawings();
    }
    drawLayer.addEventListener('mousedown', (e) => {
      const el = e.target.closest('.dr-item');
      if (!el || e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      if (editor) { commit(0, 0, true); if (editor) return; }
      const i = +el.dataset.index;
      const slot = e.target.closest('[data-slicer]');
      if (slot && !slot.disabled) {
        slicerClick(sheet().drawings[i], slot.dataset.slicer === 'clear' ? 'clear' : +slot.dataset.slicer, e.ctrlKey || e.metaKey);
        return;
      }
      if (picked !== i) pickDrawing(i);
      wrap.focus({ preventScroll: true });
      const handle = e.target.dataset.handle;
      const d = sheet().drawings[i];
      const start = boxOf(d), x0 = e.clientX, y0 = e.clientY;
      const node = drawLayer.querySelector(`[data-index="${i}"]`);
      let moved = false;
      const geometry = (m) => {
        const dx = m.clientX - x0, dy = m.clientY - y0;
        const g = { ...start };
        if (!handle) { g.x += dx; g.y += dy; }
        else {
          if (handle.includes('w')) { g.x += dx; g.w -= dx; }
          if (handle.includes('e')) g.w += dx;
          if (handle.includes('n')) { g.y += dy; g.h -= dy; }
          if (handle.includes('s')) g.h += dy;
        }
        g.w = Math.max(12, g.w);
        g.h = Math.max(12, g.h);
        g.x = Math.max(46, g.x);
        g.y = Math.max(23, g.y);
        return g;
      };
      const move = (m) => {
        if (!moved && Math.abs(m.clientX - x0) + Math.abs(m.clientY - y0) < 3) return;
        moved = true;
        const g = geometry(m);
        node.style.cssText = `left:${g.x}px;top:${g.y}px;width:${g.w}px;height:${g.h}px`;
      };
      const up = (m) => {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
        if (!moved) return;
        if (objectsLocked()) { renderDrawings(); return; }
        const g = geometry(m);
        beforeChange();
        d.from = anchorAt(g.x, g.y);
        if (d.to) d.to = anchorAt(g.x + g.w, g.y + g.h);
        else d.size = { w: Math.round(g.w), h: Math.round(g.h) };
        renderDrawings();
        onChange();
      };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
    });
    function deletePicked() {
      if (picked < 0 || objectsLocked()) return;
      beforeChange();
      sheet().drawings.splice(picked, 1);
      if (!sheet().drawings.length) delete sheet().drawings;
      picked = -1;
      renderDrawings();
      onChange();
    }
    // Referencia absoluta con hoja, como la escribe Excel en las series: 'Mi hoja'!$B$2:$B$6.
    const sheetRef = (r1, c1, r2, c2) => {
      const n = sheet().name;
      const name = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`;
      const a = `$${O.colName(c1)}$${r1 + 1}`, b = `$${O.colName(c2)}$${r2 + 1}`;
      return `${name}!${a === b ? a : a + ':' + b}`;
    };
    // Insertar gráfico con la selección: como Excel, la fila y la columna de texto son nombres y categorías.
    function insertChart(kind) {
      if (objectsLocked()) return;
      const s = sheet();
      const reg = dataRegion();
      const isText = (r, c) => { const v = book.value(s, r, c); return v === null || typeof v === 'string'; };
      const headRow = Array.from({ length: reg.c2 - reg.c1 + 1 }, (_, i) => reg.c1 + i).every((c) => isText(reg.r1, c)) && reg.r2 > reg.r1;
      const headCol = Array.from({ length: reg.r2 - reg.r1 + 1 }, (_, i) => reg.r1 + i).every((r) => isText(r, reg.c1)) && reg.c2 > reg.c1;
      const r0 = reg.r1 + (headRow ? 1 : 0), c0 = reg.c1 + (headCol ? 1 : 0);
      if (r0 > reg.r2 || c0 > reg.c2) { O.notice('Selecciona los datos del gráfico: una tabla con números.'); return; }
      const byCols = kind === 'scatter' || reg.r2 - r0 >= reg.c2 - c0;
      const series = [];
      if (byCols) {
        for (let c = c0; c <= reg.c2; c++) series.push({ name: headRow ? sheetRef(reg.r1, c, reg.r1, c) : null, cat: headCol ? sheetRef(r0, reg.c1, reg.r2, reg.c1) : null, val: sheetRef(r0, c, reg.r2, c) });
      } else {
        for (let r = r0; r <= reg.r2; r++) series.push({ name: headCol ? sheetRef(r, reg.c1, r, reg.c1) : null, cat: headRow ? sheetRef(reg.r1, c0, reg.r1, reg.c2) : null, val: sheetRef(r, c0, r, reg.c2) });
      }
      const round = kind === 'pie' || kind === 'doughnut';
      const used2 = round ? series.slice(0, 1) : series;
      const title = used2.length === 1 ? (used2[0].name ? { ref: used2[0].name } : null) : { text: 'Título del gráfico' };
      const xml = O.chartXml(kind, used2, title);
      beforeChange();
      const at = pointOf({ c: reg.c2 + 2, r: reg.r1, dx: 0, dy: 0 });
      const n = (s.drawings || []).filter((d) => d.kind === 'chart').length + 1;
      const d = { kind: 'chart', name: `Gráfico ${n}`, from: anchorAt(at.x, at.y), xml, refs: O.chartRefs(xml) };
      s.drawings = s.drawings || [];
      s.drawings.push(d);
      // Tamaño de un gráfico nuevo de Excel: 12,7 × 7,6 cm.
      d.to = anchorAt(at.x + 480, at.y + 288);
      picked = s.drawings.length - 1;
      build();
      onChange();
    }
    const CHART_KINDS = [['col', 'Columnas agrupadas'], ['bar', 'Barras agrupadas'], ['line', 'Líneas'], ['pie', 'Circular'], ['doughnut', 'Anillo'], ['area', 'Área'], ['scatter', 'Dispersión (X Y)']];

    // ---------- Autofiltro ----------
    // sheet.autoFilter = { ref, cols: { columna relativa: criterio } }. La primera fila del rango es el encabezado con las flechas.
    const filterBox = () => (sheet().autoFilter ? O.parseRanges(sheet().autoFilter.ref)[0] : null);
    const filledRow = (r, box) => {
      for (let c = box.c1; c <= box.c2; c++) {
        const cell = sheet().cells[O.addr(r, c)];
        if (cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== ''))) return true;
      }
      return false;
    };
    // Como en Excel, el filtro alcanza las filas que se escriban justo debajo.
    function growFilter() {
      const s = sheet(), box = filterBox();
      if (!box) return null;
      while (box.r2 < 1048575 && filledRow(box.r2 + 1, box)) box.r2++;
      s.autoFilter.ref = O.rangesText([box]);
      return box;
    }
    const wild = (pattern) => new RegExp('^' + pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002')
      .replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'i');
    const cellColors = (r, c) => {
      const fx = (cfMap && cfMap.get(r + ',' + c)) || {};
      const st = { ...((sheet().cells[O.addr(r, c)] || {}).s || {}), ...fx };
      return { fill: st.fill || null, font: st.color || '000000' };
    };
    function matches(spec, r, c, box) {
      const s = sheet();
      const v = book.value(s, r, c);
      const text = book.display(s, r, c);
      if (spec.vals) return text === '' ? !!spec.blank : spec.vals.some((x) => String(x).toLowerCase() === text.toLowerCase());
      if (spec.custom) {
        const one = ({ op, val }) => {
          const num = O.readNumber(String(val));
          if (num && typeof v === 'number' && !/[*?]/.test(val)) {
            const d = v - num.v;
            return { equal: d === 0, notEqual: d !== 0, greaterThan: d > 0, greaterThanOrEqual: d >= 0, lessThan: d < 0, lessThanOrEqual: d <= 0 }[op];
          }
          if (op === 'equal' || op === 'notEqual') return wild(String(val)).test(text) === (op === 'equal');
          const cmp = text.localeCompare(String(val), 'es', { sensitivity: 'base' });
          return { greaterThan: cmp > 0, greaterThanOrEqual: cmp >= 0, lessThan: cmp < 0, lessThanOrEqual: cmp <= 0 }[op];
        };
        const res = spec.custom.items.map(one);
        return spec.custom.and ? res.every(Boolean) : res.some(Boolean);
      }
      if (spec.top) {
        if (typeof v !== 'number') return false;
        const nums = [];
        for (let k = box.r1 + 1; k <= box.r2; k++) { const x = book.value(s, k, c); if (typeof x === 'number') nums.push(x); }
        nums.sort((a, b) => (spec.top.bottom ? a - b : b - a));
        const n = spec.top.percent ? Math.max(1, Math.floor((nums.length * spec.top.n) / 100)) : spec.top.n;
        const edge = nums[Math.min(n, nums.length) - 1];
        return spec.top.bottom ? v <= edge : v >= edge;
      }
      if (spec.fill !== undefined) return cellColors(r, c).fill === spec.fill;
      if (spec.font !== undefined) return cellColors(r, c).font === spec.font;
      // Un criterio que Oficina no sabe evaluar mantiene lo que dejó Excel.
      if (spec.raw) return !(s.filtered && s.filtered[r]);
      return true;
    }
    // Cada tabla de Excel tiene su propio filtro, además del autofiltro de la hoja.
    // Un filtro es { af, box, table }: box va del encabezado a la última fila de datos.
    const filterOwners = () => {
      const s = sheet(), out = [];
      const box = filterBox();
      if (box) out.push({ af: s.autoFilter, box });
      for (const t of s.tables || []) {
        if (!t.autoFilter || t.header === false) continue;
        const b = O.tableBox(t);
        out.push({ af: t.autoFilter, box: { ...b, r2: b.r2 - (t.totals || 0) }, table: t });
      }
      return out;
    };
    const ownerAt = (r, c) => filterOwners().find((o) => r >= o.box.r1 && r <= o.box.r2 && c >= o.box.c1 && c <= o.box.c2) || null;
    function applyFilter() {
      const s = sheet();
      growFilter();
      s.hiddenRows = s.hiddenRows || {};
      const before = s.filtered || {};
      const now = {};
      for (const o of filterOwners()) {
        const specs = Object.entries(o.af.cols || {}).filter(([, spec]) => !spec.noButton || Object.keys(spec).length > 1);
        for (let r = o.box.r1 + 1; r <= o.box.r2; r++) {
          if (specs.some(([id, spec]) => !matches(spec, r, o.box.c1 + +id, o.box))) now[r] = true;
        }
      }
      for (const r of Object.keys(before)) if (!now[r]) delete s.hiddenRows[r];
      for (const r of Object.keys(now)) s.hiddenRows[r] = true;
      if (Object.keys(now).length) s.filtered = now;
      else delete s.filtered;
    }
    function setFilter(owner, c, spec) {
      if (locked(null, 'autoFilter') || !owner) return;
      beforeChange();
      const id = c - owner.box.c1;
      if (spec) owner.af.cols[id] = spec;
      else delete owner.af.cols[id];
      applyFilter();
      build();
      onChange();
    }
    // Filtro (Ctrl+Mayús+L): pone o quita las flechas en la tabla de la celda activa.
    function toggleAutoFilter() {
      if (locked(null, 'autoFilter')) return;
      const s = sheet();
      const t = O.tableAt(s, sel.r, sel.c);
      beforeChange();
      if (t) {
        if (t.autoFilter) delete t.autoFilter;
        else if (t.header !== false) t.autoFilter = { cols: {} };
        applyFilter();
      } else if (s.autoFilter) {
        delete s.autoFilter;
        applyFilter();
      } else {
        const reg = dataRegion();
        s.autoFilter = { ref: O.rangesText([reg]), cols: {} };
      }
      build();
      onChange();
    }
    function clearAllFilters() {
      const o = ownerAt(sel.r, sel.c) || filterOwners()[0];
      if (!o || !Object.keys(o.af.cols).length) return;
      beforeChange();
      o.af.cols = {};
      applyFilter();
      build();
      onChange();
    }
    // Filtrar por la celda activa, como en el menú contextual de Excel: crea el autofiltro si no lo hay.
    function filterByCell(kind) {
      if (locked(null, 'autoFilter')) return;
      const s = sheet();
      let owner = ownerAt(sel.r, sel.c);
      const t = O.tableAt(s, sel.r, sel.c);
      if (!owner && t && t.header !== false) { t.autoFilter = { cols: {} }; owner = ownerAt(sel.r, sel.c); }
      if (!owner) {
        const reg = dataRegion();
        s.autoFilter = { ref: O.rangesText([reg]), cols: {} };
        owner = ownerAt(sel.r, sel.c);
      }
      if (!owner || sel.r <= owner.box.r1) { O.notice('Elige una celda con datos, no el encabezado.'); return; }
      const colors = cellColors(sel.r, sel.c);
      const text = book.display(s, sel.r, sel.c);
      setFilter(owner, sel.c, kind === 'fill' ? { fill: colors.fill } : kind === 'font' ? { font: colors.font } : text === '' ? { vals: [], blank: true } : { vals: [text] });
    }
    const filterItems = () => {
      const o = ownerAt(sel.r, sel.c) || (O.tableAt(sheet(), sel.r, sel.c) ? null : filterOwners().find((x) => !x.table));
      const t = O.tableAt(sheet(), sel.r, sel.c);
      const on = t ? !!t.autoFilter : !!sheet().autoFilter;
      return [
        { label: 'Filtrar por el valor de la celda seleccionada', action: () => filterByCell('value') },
        { label: 'Filtrar por el color de la celda seleccionada', action: () => filterByCell('fill') },
        { label: 'Filtrar por el color de fuente de la celda seleccionada', action: () => filterByCell('font') },
        { label: 'Borrar filtro', disabled: !o || !Object.keys(o.af.cols).length, action: clearAllFilters },
        { label: on ? 'Quitar las flechas de filtro' : 'Poner flechas de filtro', action: toggleAutoFilter },
      ];
    };
    const filterLayer = document.createElement('div');
    filterLayer.className = 'af-layer';
    const FUNNEL = '<svg viewBox="0 0 12 12" width="11" height="11" aria-hidden="true"><path d="M1 2h10L7 6.5V11L5 10V6.5z" fill="currentColor"/></svg>';
    function placeFilterButtons() {
      filterLayer.replaceChildren();
      const owners = filterOwners().filter((o) => o.box.r1 < rows);
      if (!owners.length) { filterLayer.remove(); return; }
      if (filterLayer.parentNode !== wrap) wrap.append(filterLayer);
      for (const owner of owners) filterButtons(owner);
    }
    function filterButtons(owner) {
      const { af, box } = owner;
      for (let c = box.c1; c <= Math.min(box.c2, cols - 1); c++) {
        const spec = af.cols[c - box.c1];
        if (spec && spec.noButton) continue;
        const m = mergeAt(box.r1, c);
        const td = m ? tds[m.r1][m.c1] : tds[box.r1][c];
        if (sheet().hiddenCols && sheet().hiddenCols[c]) continue;
        const on = !!spec && Object.keys(spec).some((k) => k !== 'noButton');
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'af-btn' + (on ? ' on' : '');
        b.dataset.col = c;
        b.innerHTML = on ? FUNNEL : '▾';
        b.title = on ? `${book.display(sheet(), box.r1, c) || O.colName(c)}: filtrado` : `${book.display(sheet(), box.r1, c) || O.colName(c)}: (Todo)`;
        b.setAttribute('aria-label', 'Filtro de ' + (book.display(sheet(), box.r1, c) || O.colName(c)));
        const right = m ? td.offsetLeft + td.offsetWidth : tds[box.r1][c].offsetLeft + tds[box.r1][c].offsetWidth;
        b.style.left = right - 17 + 'px';
        b.style.top = td.offsetTop + td.offsetHeight - 17 + 'px';
        b.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); openFilterPanel(owner, c, b); });
        filterLayer.append(b);
      }
    }
    const CONDITIONS = [['', '(Ninguna)'], ['equal', 'Es igual a'], ['notEqual', 'No es igual a'], ['greaterThan', 'Mayor que'], ['greaterThanOrEqual', 'Mayor o igual que'],
      ['lessThan', 'Menor que'], ['lessThanOrEqual', 'Menor o igual que'], ['begins', 'Comienza por'], ['ends', 'Termina con'], ['contains', 'Contiene'], ['notContains', 'No contiene'],
      ['top', 'Diez mejores (n)'], ['bottom', 'Diez inferiores (n)']];
    const conditionSpec = (op, val) => {
      if (!op || val === '') return null;
      if (op === 'top' || op === 'bottom') return { top: { n: Math.max(1, parseInt(val, 10) || 10), bottom: op === 'bottom', percent: false } };
      const pattern = { begins: `${val}*`, ends: `*${val}`, contains: `*${val}*`, notContains: `*${val}*` }[op];
      return { custom: { and: false, items: [{ op: op === 'notContains' ? 'notEqual' : pattern ? 'equal' : op, val: pattern || val }] } };
    };
    let filterPanel = null;
    function closeFilterPanel() {
      if (filterPanel) filterPanel.remove();
      filterPanel = null;
      document.removeEventListener('mousedown', outsidePanel, true);
    }
    const outsidePanel = (e) => { if (filterPanel && !filterPanel.contains(e.target)) closeFilterPanel(); };
    // Panel de la flecha: ordenar, borrar el filtro de la columna, buscar y marcar valores, o una condición.
    function openFilterPanel(owner, c, anchor) {
      closeFilterPanel();
      if (editor) commit(0, 0, true);
      const s = sheet();
      if (!owner.table) growFilter();
      owner = filterOwners().find((o) => o.af === owner.af);
      if (!owner) return;
      const box = owner.box;
      const id = c - box.c1, spec = owner.af.cols[id] || {};
      const head = book.display(s, box.r1, c) || O.colName(c);
      // Los valores salen de las filas que dejan ver los filtros de las otras columnas.
      const others = Object.entries(owner.af.cols).filter(([k]) => +k !== id);
      const seen = new Map();
      let blanks = false;
      for (let r = box.r1 + 1; r <= box.r2; r++) {
        if (others.some(([k, sp]) => !matches(sp, r, box.c1 + +k, box))) continue;
        const text = book.display(s, r, c);
        if (text === '') { blanks = true; continue; }
        const key = text.toLowerCase();
        if (!seen.has(key)) seen.set(key, { text, v: book.value(s, r, c) });
      }
      const values = [...seen.values()].sort((a, b) => (typeof a.v === 'number' && typeof b.v === 'number' ? a.v - b.v
        : typeof a.v === 'number' ? -1 : typeof b.v === 'number' ? 1 : a.text.localeCompare(b.text, 'es', { sensitivity: 'base' })));
      const checked = (text) => (spec.vals ? spec.vals.some((x) => String(x).toLowerCase() === text.toLowerCase()) : !Object.keys(spec).some((k) => k !== 'noButton'));
      const panel = document.createElement('div');
      panel.className = 'af-panel';
      panel.setAttribute('role', 'dialog');
      panel.setAttribute('aria-label', 'Filtro de ' + head);
      const action = (label, fn, disabled) => {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'af-action';
        b.textContent = label;
        b.disabled = !!disabled;
        b.onclick = () => { closeFilterPanel(); fn(); };
        panel.append(b);
      };
      const sortBy = (dir) => sortRegion(byValue(dir), { reg: { ...box }, start: box.r1 + 1, col: c });
      action('Ordenar de A a Z', () => sortBy(1));
      action('Ordenar de Z a A', () => sortBy(-1));
      action(`Borrar filtro de «${head}»`, () => setFilter(owner, c, null), !Object.keys(spec).some((k) => k !== 'noButton'));
      const cond = document.createElement('div');
      cond.className = 'af-cond';
      cond.innerHTML = `<label>Filtro de condición <select aria-label="Condición">${CONDITIONS.map(([v, l]) => `<option value="${v}">${l}</option>`).join('')}</select></label><input aria-label="Valor de la condición" spellcheck="false">`;
      const [opSel, opVal] = [cond.querySelector('select'), cond.querySelector('input')];
      if (spec.custom && spec.custom.items.length === 1) {
        const it = spec.custom.items[0], v = String(it.val);
        const m = /^\*(.*)\*$/.exec(v);
        if (m && it.op === 'notEqual') { opSel.value = 'notContains'; opVal.value = m[1]; }
        else if (m && it.op === 'equal') { opSel.value = 'contains'; opVal.value = m[1]; }
        else if (it.op === 'equal' && v.endsWith('*')) { opSel.value = 'begins'; opVal.value = v.slice(0, -1); }
        else if (it.op === 'equal' && v.startsWith('*')) { opSel.value = 'ends'; opVal.value = v.slice(1); }
        else { opSel.value = it.op; opVal.value = v; }
      } else if (spec.top && !spec.top.percent) { opSel.value = spec.top.bottom ? 'bottom' : 'top'; opVal.value = spec.top.n; }
      panel.append(cond);
      const search = document.createElement('input');
      search.className = 'af-search';
      search.placeholder = 'Buscar';
      search.setAttribute('aria-label', 'Buscar valores');
      panel.append(search);
      const list = document.createElement('div');
      list.className = 'af-list';
      const row = (label, on, value) => {
        const l = document.createElement('label');
        const box2 = document.createElement('input');
        box2.type = 'checkbox';
        box2.checked = on;
        if (value !== undefined) box2.dataset.value = value;
        l.append(box2, document.createTextNode(' ' + label));
        list.append(l);
        return box2;
      };
      const all = row('(Seleccionar todo)', false);
      const items = values.map((x) => row(x.text, checked(x.text), x.text));
      const blankBox = blanks ? row('(Vacías)', spec.vals ? !!spec.blank : checked(''), '') : null;
      const boxes = () => [...items, blankBox].filter(Boolean);
      const sync = () => { all.checked = boxes().every((b) => b.checked); all.indeterminate = !all.checked && boxes().some((b) => b.checked); };
      sync();
      all.onchange = () => { for (const b of boxes()) if (b.parentNode.style.display !== 'none') b.checked = all.checked; };
      list.addEventListener('change', (e) => { if (e.target !== all) sync(); });
      search.oninput = () => {
        const t = search.value.trim().toLowerCase();
        for (const b of boxes()) {
          const show = !t || String(b.dataset.value).toLowerCase().includes(t);
          b.parentNode.style.display = show ? '' : 'none';
          // Como en Excel, al buscar quedan marcados solo los que coinciden.
          b.checked = show;
        }
        sync();
      };
      panel.append(list);
      const foot = document.createElement('div');
      foot.className = 'af-foot';
      const ok = document.createElement('button');
      ok.type = 'button';
      ok.className = 'primary';
      ok.textContent = 'Aceptar';
      const cancel2 = document.createElement('button');
      cancel2.type = 'button';
      cancel2.textContent = 'Cancelar';
      foot.append(ok, cancel2);
      panel.append(foot);
      cancel2.onclick = closeFilterPanel;
      ok.onclick = () => {
        const byCondition = conditionSpec(opSel.value, opVal.value.trim());
        closeFilterPanel();
        if (byCondition) { setFilter(owner, c, byCondition); return; }
        const chosen = items.filter((b) => b.checked).map((b) => b.dataset.value);
        const blank = blankBox ? blankBox.checked : false;
        if (chosen.length === items.length && (!blankBox || blank)) setFilter(owner, c, null);
        else setFilter(owner, c, { vals: chosen, ...(blank ? { blank: true } : {}) });
      };
      panel.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') { e.preventDefault(); closeFilterPanel(); wrap.focus(); }
        if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok.click(); }
      });
      document.body.append(panel);
      const at = anchor.getBoundingClientRect();
      panel.style.left = Math.max(4, Math.min(window.innerWidth - panel.offsetWidth - 4, at.right - panel.offsetWidth)) + 'px';
      panel.style.top = Math.min(window.innerHeight - panel.offsetHeight - 4, at.bottom + 2) + 'px';
      filterPanel = panel;
      setTimeout(() => document.addEventListener('mousedown', outsidePanel, true));
      search.focus();
    }

    // ---------- Formato de celdas ----------
    function formatDialog() {
      if (locked(null, 'formatCells')) return;
      const st = activeStyle();
      const nfOptions = NF.map(([v, l]) => [v, l]);
      if (st.nf && !NF.some(([v]) => v === st.nf)) nfOptions.push([st.nf, st.nf]);
      O.formDialog({
        title: 'Formato de celdas',
        fields: [
          { name: 'nf', label: 'Número', type: 'select', options: nfOptions, value: st.nf || 'General' },
          { name: 'align', label: 'Alineación horizontal', type: 'select', options: [['', 'General'], ['left', 'Izquierda'], ['center', 'Centro'], ['right', 'Derecha']], value: st.align || '' },
          { name: 'valign', label: 'Alineación vertical', type: 'select', options: [['', 'Abajo'], ['center', 'Centro'], ['top', 'Arriba']], value: st.valign || '' },
          { name: 'wrap', label: 'Ajustar texto', type: 'checkbox', value: st.wrap },
          { name: 'b', label: 'Negrita', type: 'checkbox', value: st.b },
          { name: 'i', label: 'Cursiva', type: 'checkbox', value: st.i },
          { name: 'u', label: 'Subrayado', type: 'checkbox', value: st.u },
          { name: 'sz', label: 'Tamaño de fuente', type: 'number', value: st.sz || 11 },
          { name: 'useColor', label: 'Usar color de fuente', type: 'checkbox', value: !!st.color },
          { name: 'color', label: 'Color de fuente', type: 'color', value: '#' + (st.color || 'C0392B').toLowerCase() },
          { name: 'useFill', label: 'Usar relleno', type: 'checkbox', value: !!st.fill },
          { name: 'fill', label: 'Color de relleno', type: 'color', value: '#' + (st.fill || 'FFF2B3').toLowerCase() },
          { name: 'border', label: 'Bordes', type: 'select', options: [['keep', 'Sin cambios'], ['none', 'Ninguno'], ['all', 'Todos'], ['outline', 'Contorno'], ['bottom', 'Solo abajo']], value: 'keep' },
          { name: 'bstyle', label: 'Estilo de línea', type: 'select', options: LINE_STYLES, value: 'thin' },
          { name: 'bcolor', label: 'Color de línea', type: 'color', value: '#000000' },
          // Como en la pestaña Proteger de Excel: solo cuenta al proteger la hoja.
          { name: 'locked', label: 'Bloqueada (al proteger la hoja)', type: 'checkbox', value: !st.unlocked },
          { name: 'hideF', label: 'Oculta: no mostrar la fórmula al proteger la hoja', type: 'checkbox', value: !!st.hideF },
        ],
        onOk: (v) => {
          const sz = +v.sz;
          if (!(sz >= 6 && sz <= 72)) return 'El tamaño debe estar entre 6 y 72.';
          const box = { ...sel };
          restyle((s, r, c) => {
            s.nf = v.nf === 'General' ? undefined : v.nf;
            s.align = v.align || undefined;
            s.valign = v.valign || undefined;
            s.wrap = v.wrap || undefined;
            s.b = v.b || undefined;
            s.i = v.i || undefined;
            s.u = v.u || undefined;
            s.sz = sz === 11 ? undefined : sz;
            s.color = v.useColor ? O.hex(v.color) : undefined;
            s.fill = v.useFill ? O.hex(v.fill) : undefined;
            s.unlocked = v.locked ? undefined : true;
            s.hideF = v.hideF || undefined;
            // Los bordes elegidos aquí usan el estilo y color del diálogo; se olvida el que trajera el archivo.
            if (['none', 'all', 'bottom', 'outline'].includes(v.border)) s.bdx = undefined;
            if (v.border === 'none') s.bd = undefined;
            else if (v.border === 'all') s.bd = 'lrtb';
            else if (v.border === 'bottom') s.bd = 'b';
            else if (v.border === 'outline') {
              const bd = (c === box.c1 ? 'l' : '') + (c === box.c2 ? 'r' : '') + (r === box.r1 ? 't' : '') + (r === box.r2 ? 'b' : '');
              s.bd = bd || undefined;
            }
            const color = O.hex(v.bcolor) || '000000';
            if (s.bd && ['all', 'bottom', 'outline'].includes(v.border) && (v.bstyle !== 'thin' || color !== '000000')) {
              s.bdx = Object.fromEntries([...s.bd].map((k) => [k, [v.bstyle, color]]));
            }
          }, true);
          return null;
        },
      });
    }

    // ---------- Notas, vinculos, nombres y lista ----------
    function setCellProp(prop, value) {
      if (locked()) return;
      beforeChange();
      const key = O.addr(sel.r, sel.c);
      const cell = sheet().cells[key] || (sheet().cells[key] = {});
      if (value) cell[prop] = value;
      else delete cell[prop];
      if (!Object.keys(cell).length) delete sheet().cells[key];
      paint(sel.r, sel.c);
      paintSel();
      onChange();
    }
    function noteDialog() {
      if (locked()) return;
      const cur = (sheet().cells[O.addr(sel.r, sel.c)] || {}).note || '';
      O.formDialog({
        title: cur ? 'Editar nota' : 'Nueva nota',
        fields: [{ name: 'note', label: 'Nota', type: 'textarea', value: cur }],
        onOk: (v) => { setCellProp('note', v.note.trim()); return null; },
      });
    }
    function openLink(url) {
      if (/^(https?:|mailto:)/i.test(url)) window.open(url, '_blank', 'noopener');
      else O.notice('Oficina solo abre vínculos web y de correo.');
    }
    function linkDialog() {
      if (locked()) return;
      const cell = sheet().cells[O.addr(sel.r, sel.c)] || {};
      O.formDialog({
        title: cell.link ? 'Editar vínculo' : 'Insertar vínculo',
        fields: [
          { name: 'url', label: 'Dirección', value: cell.link || '' },
          { name: 'text', label: 'Texto que se muestra', value: cell.f === undefined && cell.v !== undefined ? cell.v : '' },
        ],
        onOk: (v) => {
          let url = v.url.trim();
          if (!url) return 'Escribe una dirección.';
          if (!/^[a-z][a-z0-9+.-]*:/i.test(url)) url = 'https://' + url;
          if (!/^(https?:|mailto:)/i.test(url)) return 'Solo se admiten direcciones web (http, https) y de correo (mailto).';
          const text = v.text.trim();
          beforeChange();
          if (cell.f === undefined) setCell(sel.r, sel.c, text || url);
          sheet().cells[O.addr(sel.r, sel.c)].link = url;
          book.clearCache();
          paint(sel.r, sel.c);
          recalc();
          paintSel();
          onChange();
          return null;
        },
      });
    }
    function nameDialog() {
      const multi = sel.r1 !== sel.r2 || sel.c1 !== sel.c2;
      const abs = (r, c) => `$${O.colName(c)}$${r + 1}`;
      const ref = `${qname(sheet().name)}!${abs(sel.r1, sel.c1)}${multi ? ':' + abs(sel.r2, sel.c2) : ''}`;
      O.formDialog({
        title: 'Definir nombre',
        fields: [{ name: 'name', label: 'Nombre' }, { name: 'ref', label: 'Se refiere a', value: ref }],
        onOk: (v) => {
          const name = v.name.trim();
          if (!/^[A-Za-z_À-ɏ][\w.À-ɏ]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || /^(true|false|verdadero|falso)$/i.test(name)) return 'El nombre debe empezar por una letra, sin espacios, y no puede parecer una celda como A1.';
          if (!v.ref.trim()) return 'Indica a qué celdas se refiere.';
          beforeChange();
          state.names = state.names || {};
          for (const k of Object.keys(state.names)) if (k.toLowerCase() === name.toLowerCase()) delete state.names[k];
          state.names[name] = v.ref.trim().replace(/^=/, '');
          book.clearCache();
          paintAll();
          paintSel();
          onChange();
          return null;
        },
      });
    }
    // Los valores de texto del bloque continuo de la columna, como la lista desplegable de Excel.
    function pickItems() {
      const s = sheet(), c = sel.c;
      const found = new Set();
      const text = (r) => {
        const cell = s.cells[O.addr(r, c)];
        return cell && typeof cell.v === 'string' && cell.v !== '' ? cell.v : null;
      };
      for (let r = sel.r - 1; r >= 0 && text(r) !== null; r--) found.add(text(r));
      for (let r = sel.r + 1; r < rows && text(r) !== null; r++) found.add(text(r));
      if (!found.size) return [{ label: 'Sin valores en esta columna', disabled: true, action: () => {} }];
      return [...found].sort((a, b) => a.localeCompare(b, 'es')).map((v) => ({
        label: v,
        action: () => {
          if (locked({ r1: sel.r, r2: sel.r, c1: sel.c, c2: sel.c })) return;
          beforeChange();
          setCell(sel.r, sel.c, v);
          paint(sel.r, sel.c);
          recalc();
          paintSel();
          onChange();
        },
      }));
    }
    function cellExtraItems() {
      const cell = sheet().cells[O.addr(sel.r, sel.c)] || {};
      const noteItems = [{ label: cell.note ? 'Editar nota…' : 'Nueva nota…', action: noteDialog }];
      if (cell.note) noteItems.push({ label: 'Eliminar nota', action: () => setCellProp('note', null) });
      const linkItems = [{ label: cell.link ? 'Editar vínculo…' : 'Vínculo…', action: linkDialog }];
      if (cell.link) linkItems.push({ label: 'Abrir vínculo', action: () => openLink(cell.link) }, { label: 'Quitar vínculo', action: () => setCellProp('link', null) });
      return { noteItems, linkItems };
    }
    const clearMoreItems = () => [
      { label: 'Borrar formato', action: () => clearProps(['s']) },
      { label: 'Borrar notas', action: () => clearProps(['note']) },
      { label: 'Borrar vínculos', action: () => clearProps(['link']) },
      { label: 'Borrar todo', action: () => clearProps(null) },
    ];
    function menuCopy(after) {
      const failed = () => O.notice('El navegador no ha dado acceso al portapapeles. Usa Ctrl+C o Ctrl+X.');
      if (!navigator.clipboard) { failed(); return; }
      navigator.clipboard.writeText(takeClip()).then(after, failed);
    }
    function menuPaste() {
      const failed = () => O.notice('El navegador no ha dado acceso al portapapeles. Usa Ctrl+V.');
      if (!navigator.clipboard) { failed(); return; }
      navigator.clipboard.readText().then(pasteText, failed);
    }
    function editItems() {
      return [
        { label: 'Cortar', hint: 'Ctrl+X', action: () => menuCopy(clearRange) },
        { label: 'Copiar', hint: 'Ctrl+C', action: () => menuCopy() },
        { label: 'Pegar', hint: 'Ctrl+V', action: menuPaste },
        { label: 'Pegado especial', disabled: !clip, items: pasteSpecialItems() },
      ];
    }
    const clearItems = () => [
      { label: 'Borrar contenido', hint: 'Supr', action: clearRange },
      { label: 'Borrar', items: clearMoreItems() },
    ];
    const formatItem = () => ({ label: 'Formato de celdas…', hint: 'Ctrl+1', action: formatDialog });
    wrap.addEventListener('contextmenu', (e) => {
      if (editor && e.target === editor.input) return;
      e.preventDefault();
      const th = e.target.closest('th');
      const td = cellAt(e);
      let x = e.clientX, y = e.clientY;
      // Con la tecla de menu no hay clic: se abre junto a la celda actual.
      if (!td && !th) {
        const box = tds[sel.r][sel.c].getBoundingClientRect();
        x = box.left + 10;
        y = box.bottom;
      }
      let items;
      if (th && th.dataset.col !== undefined) {
        items = [...editItems(), '-', { label: 'Insertar', items: insertColItems() }, { label: 'Eliminar', items: [deleteColItem()] }, ...clearItems(), '-', formatItem(), ...widthItems(), '-', ...hideItems('c')];
      } else if (th && th.dataset.row !== undefined) {
        items = [...editItems(), '-', { label: 'Insertar', items: insertRowItems() }, { label: 'Eliminar', items: [deleteRowItem()] }, ...clearItems(), '-', formatItem(), { label: 'Alto de fila…', action: askHeight }, '-', ...hideItems('r')];
      } else {
        const r1 = sel.r1, r2 = sel.r2, c1 = sel.c1, c2 = sel.c2;
        const insert = [], remove = [];
        if (!fullCol() && !fullRow()) {
          insert.push(
            { label: 'Desplazar celdas a la derecha', action: () => structural('c', c1, colN(), [r1, r2]) },
            { label: 'Desplazar celdas hacia abajo', action: () => structural('r', r1, rowN(), [c1, c2]) },
            '-',
          );
          remove.push(
            { label: 'Desplazar celdas a la izquierda', action: () => structural('c', c1, -colN(), [r1, r2]) },
            { label: 'Desplazar celdas hacia arriba', action: () => structural('r', r1, -rowN(), [c1, c2]) },
            '-',
          );
        }
        if (!fullCol()) { insert.push(...insertRowItems()); remove.push(deleteRowItem()); }
        if (!fullRow()) { insert.push(...insertColItems()); remove.push(deleteColItem()); }
        const { noteItems, linkItems } = cellExtraItems();
        items = [...editItems(), '-'];
        if (insert.length) items.push({ label: 'Insertar', items: insert }, { label: 'Eliminar', items: remove });
        items.push(...clearItems(), '-',
          { label: 'Filtrar', items: filterItems() }, { label: 'Ordenar', items: sortItems() }, '-',
          ...noteItems, '-', formatItem(),
          { label: 'Elegir de la lista', items: pickItems() },
          { label: 'Definir nombre…', action: nameDialog },
          ...linkItems, '-', ...widthItems());
      }
      O.contextMenu(x, y, items);
    });

    // ---------- Formato condicional ----------
    const PRESETS = [
      ['red', 'Relleno rojo claro con texto rojo oscuro', { fill: 'FFC7CE', color: '9C0006' }],
      ['yellow', 'Relleno amarillo con texto amarillo oscuro', { fill: 'FFEB9C', color: '9C5700' }],
      ['green', 'Relleno verde con texto verde oscuro', { fill: 'C6EFCE', color: '006100' }],
      ['lightred', 'Relleno rojo claro', { fill: 'FFC7CE' }],
      ['redtext', 'Texto rojo', { color: '9C0006' }],
      ['bold', 'Negrita', { b: true }],
      ['custom', 'Personalizado (con las opciones de abajo)', null],
    ];
    const styleFields = (preset) => [
      { name: 'preset', label: 'Con este formato', type: 'select', options: PRESETS.map(([v, l]) => [v, l]), value: preset || 'red' },
      { name: 'useFill', label: 'Personalizado: relleno', type: 'checkbox', value: true },
      { name: 'fill', label: 'Color de relleno', type: 'color', value: '#ffeb9c' },
      { name: 'useColor', label: 'Personalizado: color de fuente', type: 'checkbox', value: false },
      { name: 'color', label: 'Color de fuente', type: 'color', value: '#9c0006' },
      { name: 'b', label: 'Personalizado: negrita', type: 'checkbox', value: false },
      { name: 'i', label: 'Personalizado: cursiva', type: 'checkbox', value: false },
    ];
    const styleOf = (v) => {
      const p = PRESETS.find(([k]) => k === v.preset);
      if (p && p[2]) return { ...p[2] };
      const s = {};
      if (v.useFill) s.fill = O.hex(v.fill);
      if (v.useColor) s.color = O.hex(v.color);
      if (v.b) s.b = true;
      if (v.i) s.i = true;
      return Object.keys(s).length ? s : null;
    };
    // Lo escrito en el cuadro pasa a fórmula: =A1 es una referencia, 5 un número y lo demás texto.
    const valueFormula = (text) => {
      const t = String(text).trim();
      if (t.startsWith('=')) return O.fromLocalFormula(t.slice(1));
      const n = O.readNumber(t);
      if (n) return String(n.v);
      return `"${t.replace(/"/g, '""')}"`;
    };
    // Rango de la selección; una columna o fila entera llega hasta el final de la hoja, como en Excel.
    const selRef = () => O.rangesText([{ r1: sel.r1, c1: sel.c1, r2: fullCol() ? 1048575 : sel.r2, c2: fullRow() ? 16383 : sel.c2 }]);
    const checkFormula = (f) => {
      try { O.parseFormula(f); return null; } catch (e) { return 'La fórmula no es válida. Revisa paréntesis, comillas y separadores.'; }
    };
    function addRule(rule) {
      if (locked()) return;
      beforeChange();
      const s = sheet();
      // La regla nueva va primero, con la prioridad más alta, como en Excel.
      s.cf = [{ ref: selRef(), ...rule }, ...(s.cf || [])];
      paintAll();
      paintSel();
      onChange();
    }
    const curText = () => {
      const v = book.value(sheet(), sel.r, sel.c);
      return v === null || (v && v.err) ? '' : O.formatValue(v, null);
    };
    function cellIsDialog(op) {
      if (locked()) return;
      const two = op === 'between' || op === 'notBetween';
      const fields = [{ name: 'a', label: `Dar formato a las celdas con valor ${O.cfOperators[op]}`, value: curText() }];
      if (two) fields.push({ name: 'b', label: 'y', value: '' });
      O.formDialog({
        title: `Valor de celda ${O.cfOperators[op]}`,
        fields: [...fields, ...styleFields()],
        onOk: (v) => {
          if (!v.a.trim() || (two && !v.b.trim())) return 'Escribe el valor.';
          const f = [valueFormula(v.a)];
          if (two) f.push(valueFormula(v.b));
          for (const x of f) { const bad = checkFormula(x); if (bad) return bad; }
          const style = styleOf(v);
          if (!style) return 'Elige algún formato.';
          addRule({ type: 'cellIs', op, f, style });
          return null;
        },
      });
    }
    function textDialog() {
      if (locked()) return;
      O.formDialog({
        title: 'Texto que contiene',
        fields: [
          { name: 'kind', label: 'Celdas cuyo texto', type: 'select', options: [['containsText', 'Contiene'], ['notContainsText', 'No contiene'], ['beginsWith', 'Empieza por'], ['endsWith', 'Termina en']], value: 'containsText' },
          { name: 'text', label: 'Texto', value: curText() },
          ...styleFields(),
        ],
        onOk: (v) => {
          if (!v.text) return 'Escribe el texto.';
          const style = styleOf(v);
          if (!style) return 'Elige algún formato.';
          addRule({ type: v.kind, text: v.text, style });
          return null;
        },
      });
    }
    function dateDialog() {
      if (locked()) return;
      O.formDialog({
        title: 'Una fecha',
        fields: [{ name: 'period', label: 'Dar formato a las fechas de', type: 'select', options: Object.entries(O.cfPeriods), value: 'today' }, ...styleFields()],
        onOk: (v) => {
          const style = styleOf(v);
          if (!style) return 'Elige algún formato.';
          addRule({ type: 'timePeriod', period: v.period, style });
          return null;
        },
      });
    }
    function dupDialog() {
      if (locked()) return;
      O.formDialog({
        title: 'Duplicar valores',
        fields: [{ name: 'kind', label: 'Dar formato a los valores', type: 'select', options: [['duplicateValues', 'Duplicados'], ['uniqueValues', 'Únicos']], value: 'duplicateValues' }, ...styleFields()],
        onOk: (v) => {
          const style = styleOf(v);
          if (!style) return 'Elige algún formato.';
          addRule({ type: v.kind, style });
          return null;
        },
      });
    }
    function topDialog(bottom, percent) {
      if (locked()) return;
      O.formDialog({
        title: `${percent ? '10 %' : '10'} ${bottom ? 'inferiores' : 'superiores'}`,
        fields: [{ name: 'rank', label: `Dar formato a ${bottom ? 'los inferiores' : 'los superiores'}${percent ? ' (en %)' : ''}`, type: 'number', value: 10 }, ...styleFields()],
        onOk: (v) => {
          const rank = Math.floor(+v.rank);
          if (!(rank >= 1 && rank <= (percent ? 100 : 1000))) return percent ? 'Indica un porcentaje entre 1 y 100.' : 'Indica un número entre 1 y 1000.';
          const style = styleOf(v);
          if (!style) return 'Elige algún formato.';
          addRule({ type: 'top10', rank, percent: percent || undefined, bottom: bottom || undefined, style });
          return null;
        },
      });
    }
    function averageDialog(above) {
      if (locked()) return;
      O.formDialog({
        title: above ? 'Por encima del promedio' : 'Por debajo del promedio',
        fields: styleFields(),
        onOk: (v) => {
          const style = styleOf(v);
          if (!style) return 'Elige algún formato.';
          addRule({ type: 'aboveAverage', above: above ? undefined : false, style });
          return null;
        },
      });
    }
    const NEW_TYPES = [
      ['greaterThan', 'Valor de celda mayor que'], ['lessThan', 'Valor de celda menor que'], ['between', 'Valor de celda entre'],
      ['notBetween', 'Valor de celda no entre'], ['equal', 'Valor de celda igual a'], ['notEqual', 'Valor de celda distinto de'],
      ['greaterThanOrEqual', 'Valor de celda mayor o igual que'], ['lessThanOrEqual', 'Valor de celda menor o igual que'],
      ['expression', 'Utilice una fórmula que determine las celdas'], ['containsBlanks', 'Celdas en blanco'], ['notContainsBlanks', 'Celdas que no están en blanco'],
      ['containsErrors', 'Celdas con errores'],
    ];
    function newRuleDialog() {
      if (locked()) return;
      O.formDialog({
        title: 'Nueva regla de formato',
        fields: [
          { name: 'kind', label: 'Tipo de regla', type: 'select', options: NEW_TYPES, value: 'expression' },
          { name: 'a', label: `Valor o fórmula (las fórmulas se escriben con =, relativas a ${O.addr(sel.r1, sel.c1)})`, value: '' },
          { name: 'b', label: 'Segundo valor (solo para entre y no entre)', value: '' },
          { name: 'stop', label: 'Detener si es verdad', type: 'checkbox', value: false },
          ...styleFields(),
        ],
        onOk: (v) => {
          const style = styleOf(v);
          if (!style) return 'Elige algún formato.';
          const rule = { style };
          if (v.stop) rule.stop = true;
          if (v.kind === 'expression') {
            const t = v.a.trim().replace(/^=/, '');
            if (!t) return 'Escribe la fórmula, por ejemplo =A1>10.';
            rule.type = 'expression';
            rule.f = [O.fromLocalFormula(t)];
          } else if (/Blanks|Errors/.test(v.kind)) rule.type = v.kind;
          else {
            const two = v.kind === 'between' || v.kind === 'notBetween';
            if (!v.a.trim() || (two && !v.b.trim())) return 'Escribe el valor.';
            rule.type = 'cellIs';
            rule.op = v.kind;
            rule.f = [valueFormula(v.a)];
            if (two) rule.f.push(valueFormula(v.b));
          }
          for (const x of rule.f || []) { const bad = checkFormula(x); if (bad) return bad; }
          addRule(rule);
          return null;
        },
      });
    }
    function clearRules(all) {
      if (locked()) return;
      const s = sheet();
      if (!s.cf || !s.cf.length) { O.notice('Esta hoja no tiene reglas de formato condicional.'); return; }
      beforeChange();
      s.cf = all ? [] : O.removeRuleCells(s.cf, { r1: sel.r1, c1: sel.c1, r2: fullCol() ? 1048575 : sel.r2, c2: fullRow() ? 16383 : sel.c2 });
      if (!s.cf.length) delete s.cf;
      paintAll();
      paintSel();
      onChange();
    }
    function manageRules() {
      const s = sheet();
      if (!s.cf || !s.cf.length) { O.notice('Esta hoja no tiene reglas de formato condicional.'); return; }
      O.formDialog({
        title: 'Reglas de formato condicional de esta hoja',
        okLabel: 'Eliminar las marcadas',
        fields: s.cf.map((rule, i) => ({ name: 'r' + i, type: 'checkbox', label: `${i + 1}. ${O.describeRule(rule)} — se aplica a ${rule.ref}${rule.stop ? ' (detener si es verdad)' : ''}` })),
        onOk: (v) => {
          if (locked()) return null;
          const keep = s.cf.filter((_, i) => !v['r' + i]);
          if (keep.length === s.cf.length) return null;
          beforeChange();
          if (keep.length) s.cf = keep;
          else delete s.cf;
          paintAll();
          paintSel();
          onChange();
          return null;
        },
      });
    }
    const BARS = [['Azul', '638EC6'], ['Verde', '63C384'], ['Rojo', 'FF555A'], ['Naranja', 'FFB628'], ['Celeste', '008AEF'], ['Morado', 'D6007B']];
    const SCALES = [
      ['Verde, amarillo y rojo (verde lo más alto)', ['F8696B', 'FFEB84', '63BE7B']],
      ['Rojo, amarillo y verde (rojo lo más alto)', ['63BE7B', 'FFEB84', 'F8696B']],
      ['Verde, blanco y rojo', ['F8696B', 'FCFCFF', '63BE7B']],
      ['Azul, blanco y rojo', ['F8696B', 'FCFCFF', '5A8AC6']],
      ['Blanco y verde', ['FCFCFF', '63BE7B']],
      ['Blanco y rojo', ['FCFCFF', 'F8696B']],
      ['Verde y amarillo', ['FFEF9C', '63BE7B']],
    ];
    const ICONS = [
      ['3 flechas de colores', '3Arrows'], ['3 flechas grises', '3ArrowsGray'], ['3 triángulos', '3Triangles'], ['4 flechas de colores', '4Arrows'],
      ['5 flechas de colores', '5Arrows'], ['3 semáforos', '3TrafficLights1'], ['3 señales', '3Signs'], ['4 semáforos', '4TrafficLights'],
      ['3 símbolos', '3Symbols'], ['3 banderas', '3Flags'], ['3 estrellas', '3Stars'], ['4 clasificaciones', '4Rating'],
      ['5 cuartos', '5Quarters'], ['5 clasificaciones', '5Rating'], ['5 cuadros', '5Boxes'],
    ];
    const scaleRule = (colors) => ({
      type: 'colorScale',
      scale: colors.length === 3
        ? [{ type: 'min', color: colors[0] }, { type: 'percentile', val: '50', color: colors[1] }, { type: 'max', color: colors[2] }]
        : [{ type: 'min', color: colors[0] }, { type: 'max', color: colors[1] }],
    });
    const cfItems = () => [
      { label: 'Resaltar reglas de celdas', items: [
        { label: 'Es mayor que…', action: () => cellIsDialog('greaterThan') },
        { label: 'Es menor que…', action: () => cellIsDialog('lessThan') },
        { label: 'Entre…', action: () => cellIsDialog('between') },
        { label: 'Es igual a…', action: () => cellIsDialog('equal') },
        { label: 'Texto que contiene…', action: textDialog },
        { label: 'Una fecha…', action: dateDialog },
        { label: 'Duplicar valores…', action: dupDialog },
      ] },
      { label: 'Reglas superiores e inferiores', items: [
        { label: '10 superiores…', action: () => topDialog(false, false) },
        { label: '10 % superiores…', action: () => topDialog(false, true) },
        { label: '10 inferiores…', action: () => topDialog(true, false) },
        { label: '10 % inferiores…', action: () => topDialog(true, true) },
        { label: 'Por encima del promedio…', action: () => averageDialog(true) },
        { label: 'Por debajo del promedio…', action: () => averageDialog(false) },
      ] },
      { label: 'Barras de datos', items: BARS.map(([name, color]) => ({ label: name, action: () => addRule({ type: 'dataBar', bar: { min: { type: 'min' }, max: { type: 'max' }, color } }) })) },
      { label: 'Escalas de color', items: SCALES.map(([name, colors]) => ({ label: name, action: () => addRule(scaleRule(colors)) })) },
      { label: 'Conjuntos de iconos', items: ICONS.map(([name, set]) => ({
        label: name, html: O.ICON_SETS[set].map(([shape, color]) => O.iconSvg(shape, color)).join(''), action: () => addRule({ type: 'iconSet', icons: { set } }),
      })) },
      '-',
      { label: 'Nueva regla…', action: newRuleDialog },
      { label: 'Borrar reglas', items: [
        { label: 'Borrar reglas de las celdas seleccionadas', action: () => clearRules(false) },
        { label: 'Borrar reglas de toda la hoja', action: () => clearRules(true) },
      ] },
      { label: 'Administrar reglas…', action: manageRules },
    ];

    // ---------- Funciones: asistente y autocompletado ----------
    // Mete el texto de una función en la celda: si ya se está escribiendo una fórmula, se añade al final.
    let pendingFormula = null;
    function insertFunction(text) {
      if (locked()) return;
      const start = pendingFormula && pendingFormula.r === sel.r && pendingFormula.c === sel.c ? pendingFormula.text : '=';
      pendingFormula = null;
      // Después de que el cuadro se cierre y devuelva el foco, para que el editor se quede con él.
      setTimeout(() => startEdit(start + text));
    }
    function argsDialog(en) {
      const info = O.FUNC_INFO[en] || ['', '', ''];
      const args = info[2] ? info[2].split(';').map((x) => x.trim()).filter((x) => x !== '…') : [];
      O.formDialog({
        title: `Argumentos de ${O.localName(en)}`,
        fields: [
          ...args.map((a, i) => ({ name: 'a' + i, label: `${a.replace(/[[\]]/g, '')}${a.startsWith('[') ? ' (opcional)' : ''}`, value: i === 0 && /rango|número1|valor1|matriz/.test(a) ? selRange() : '' })),
        ],
        okLabel: 'Insertar',
        onOk: (v) => {
          const vals = args.map((_, i) => v['a' + i].trim());
          while (vals.length && !vals[vals.length - 1]) vals.pop();
          const missing = args.findIndex((a, i) => !a.startsWith('[') && !vals[i]);
          if (missing >= 0) return `Falta ${args[missing]}.`;
          insertFunction(`${O.localName(en)}(${vals.join(';')})`);
          return null;
        },
      });
    }
    const selRange = () => (sel.r1 !== sel.r2 || sel.c1 !== sel.c2 ? `${O.addr(sel.r1, sel.c1)}:${O.addr(sel.r2, sel.c2)}` : '');
    function functionDialog() {
      // Una fórmula a medio escribir se guarda aparte y la función se añade al final.
      pendingFormula = null;
      if (editor) {
        if (editor.input.value.startsWith('=')) {
          pendingFormula = { text: editor.input.value, r: editor.r, c: editor.c };
          cancel();
        } else commit(0, 0, true);
      }
      const cats = [...new Set(Object.values(O.FUNC_INFO).map((x) => x[0]))];
      const list = O.localFunctionNames();
      O.formDialog({
        title: 'Insertar función',
        fields: [{
          name: 'fn', label: 'Función', type: 'select', value: 'SUM',
          options: cats.flatMap((cat) => list.filter((x) => O.FUNC_INFO[x.en][0] === cat).map((x) => [x.en, `${cat} · ${x.name}: ${O.FUNC_INFO[x.en][1]}`])),
        }],
        okLabel: 'Siguiente',
        onOk: (v) => {
          setTimeout(() => argsDialog(v.fn));
          return null;
        },
      });
    }
    // Lista de funciones que empiezan por lo escrito y, dentro de una función, su sintaxis.
    const suggest = document.createElement('div');
    suggest.className = 'fx-suggest';
    suggest.setAttribute('role', 'listbox');
    suggest.hidden = true;
    document.body.append(suggest);
    let suggestFor = null, suggestPick = 0, suggestItems = [];
    const hideSuggest = () => { suggest.hidden = true; suggestFor = null; suggestItems = []; };
    function updateSuggest(input) {
      const text = input.value.slice(0, input.selectionStart);
      if (!text.startsWith('=') || (text.match(/"/g) || []).length % 2) { hideSuggest(); return; }
      const m = /(?:^=|[=(;,+\-*/^&<>:\s])([A-Za-zÁÉÍÓÚÑáéíóúñ][\w.ÁÉÍÓÚÑáéíóúñ]*)$/.exec(text);
      suggestItems = m ? O.localFunctionNames().filter((x) => x.name.startsWith(m[1].toUpperCase())).slice(0, 8) : [];
      suggest.replaceChildren();
      if (suggestItems.length) {
        suggestPick = Math.min(suggestPick, suggestItems.length - 1);
        suggestItems.forEach((x, i) => {
          const row = document.createElement('div');
          row.className = 'fx-item' + (i === suggestPick ? ' on' : '');
          row.setAttribute('role', 'option');
          const b = document.createElement('b');
          b.textContent = x.name;
          const d = document.createElement('span');
          d.textContent = O.FUNC_INFO[x.en][1];
          row.append(b, d);
          row.onmousedown = (e) => { e.preventDefault(); acceptSuggest(input, i); };
          suggest.append(row);
        });
      } else {
        // Muestra la sintaxis de la función en la que está el cursor.
        let depth = 0, name = null;
        for (let i = text.length - 1; i >= 0; i--) {
          if (text[i] === ')') depth++;
          else if (text[i] === '(') {
            if (depth === 0) { const n = /([A-Za-zÁÉÍÓÚÑáéíóúñ][\w.ÁÉÍÓÚÑáéíóúñ]*)$/.exec(text.slice(0, i)); name = n && n[1]; break; }
            depth--;
          }
        }
        const en = name && O.englishName(name);
        if (!en || !O.FUNC_INFO[en]) { hideSuggest(); return; }
        const hint = document.createElement('div');
        hint.className = 'fx-hint';
        hint.textContent = O.funcSignature(en);
        suggest.append(hint);
      }
      suggestFor = input;
      suggest.hidden = false;
      const box = input.getBoundingClientRect();
      suggest.style.left = box.left + 'px';
      suggest.style.top = box.bottom + 2 + 'px';
    }
    function acceptSuggest(input, i) {
      const x = suggestItems[i];
      if (!x) return;
      const pos = input.selectionStart;
      const before = input.value.slice(0, pos).replace(/[A-Za-zÁÉÍÓÚÑáéíóúñ][\w.ÁÉÍÓÚÑáéíóúñ]*$/, x.name + '(');
      input.value = before + input.value.slice(pos);
      input.setSelectionRange(before.length, before.length);
      input.dispatchEvent(new Event('input'));
      updateSuggest(input);
    }
    // Devuelve true si la tecla la ha usado la lista de sugerencias.
    function suggestKey(e, input) {
      if (suggest.hidden || suggestFor !== input || !suggestItems.length) return false;
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        suggestPick = (suggestPick + (e.key === 'ArrowDown' ? 1 : suggestItems.length - 1)) % suggestItems.length;
        updateSuggest(input);
        return true;
      }
      if (e.key === 'Tab') { acceptSuggest(input, suggestPick); return true; }
      if (e.key === 'Escape') { hideSuggest(); return true; }
      return false;
    }
    const watchFormula = (input) => {
      input.addEventListener('input', () => { suggestPick = 0; updateSuggest(input); });
      input.addEventListener('click', () => updateSuggest(input));
      input.addEventListener('blur', () => setTimeout(() => { if (document.activeElement !== input) hideSuggest(); }));
    };
    watchFormula(fbar);

    // ---------- Barra ----------
    q('#s-cf').onclick = () => {
      if (editor) commit(0, 0, true);
      const box = q('#s-cf').getBoundingClientRect();
      O.contextMenu(box.left, box.bottom + 2, cfItems());
    };
    q('#s-filter').onclick = toggleAutoFilter;
    q('#s-chart').onclick = () => {
      const box = q('#s-chart').getBoundingClientRect();
      O.contextMenu(box.left, box.bottom + 2, CHART_KINDS.map(([kind, label]) => ({ label, action: () => insertChart(kind) })));
    };
    q('#s-merge').onclick = () => mergeSelection('center');
    q('#s-merge-more').onclick = () => {
      const box = q('#s-merge-more').getBoundingClientRect();
      O.contextMenu(box.left, box.bottom + 2, mergeItems());
    };
    // Los botones de función no quitan el foco al editor, para poder añadir la función a la fórmula.
    for (const id of ['#s-fx', '#s-fxbar']) {
      q(id).onmousedown = (e) => e.preventDefault();
      q(id).onclick = functionDialog;
    }
    q('.toolbar').addEventListener('mousedown', (e) => { if (e.target.closest('button') && !e.target.closest('#s-fx') && editor) commit(0, 0, true); });
    q('#s-undo').onclick = doUndo;
    q('#s-redo').onclick = doRedo;
    q('#s-b').onclick = () => toggle('b');
    q('#s-i').onclick = () => toggle('i');
    q('#s-color').onchange = (e) => restyle((s) => { s.color = O.hex(e.target.value); });
    q('#s-fill').onchange = (e) => restyle((s) => { s.fill = O.hex(e.target.value); });
    q('#s-nofill').onclick = () => restyle((s) => { s.fill = undefined; s.color = undefined; });
    q('#s-nf').onchange = (e) => {
      restyle((s) => { s.nf = e.target.value === 'General' ? undefined : e.target.value; });
      wrap.focus();
    };
    root.querySelectorAll('[data-align]').forEach((b) => {
      b.onclick = () => {
        const a = b.dataset.align;
        const cur = ((sheet().cells[O.addr(sel.r, sel.c)] || {}).s || {}).align;
        restyle((s) => { s.align = cur === a ? undefined : a; });
      };
    });
    q('#s-sum').onclick = () => {
      const s = sheet();
      const isNum = (r, c) => { const v = r >= 0 && c >= 0 ? book.value(s, r, c) : null; return typeof v === 'number'; };
      let f;
      if (isNum(sel.r - 1, sel.c)) {
        let top = sel.r - 1;
        while (isNum(top - 1, sel.c)) top--;
        f = `=SUMA(${O.addr(top, sel.c)}:${O.addr(sel.r - 1, sel.c)})`;
      } else if (isNum(sel.r, sel.c - 1)) {
        let left = sel.c - 1;
        while (isNum(sel.r, left - 1)) left--;
        f = `=SUMA(${O.addr(sel.r, left)}:${O.addr(sel.r, sel.c - 1)})`;
      } else f = '=SUMA(';
      startEdit(f);
    };

    // ---------- Pestanas de hojas ----------
    function activate(i) {
      if (i === (state.active || 0)) return;
      if (editor) commit(0, 0, true);
      state.active = i;
      sel = { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
      build();
      wrap.scrollTop = wrap.scrollLeft = 0;
    }
    function renameSheet(s) {
      if (bookLocked()) return;
      const name = (window.prompt('Nombre de la hoja', s.name) || '').trim();
      if (!name || name === s.name) return;
      if (/[\\/?*[\]:]/.test(name) || name.length > 31) { O.notice('El nombre admite hasta 31 caracteres, sin \\ / ? * [ ] ni :.'); return; }
      if (state.sheets.some((o) => o !== s && o.name.toLowerCase() === name.toLowerCase())) { O.notice('Ya hay una hoja con ese nombre.'); return; }
      beforeChange();
      for (const sh of state.sheets) {
        for (const cell of Object.values(sh.cells)) {
          if (cell.f !== undefined) cell.f = O.renameSheetRefs(cell.f, s.name, name);
        }
        for (const rule of sh.cf || []) if (rule.f) rule.f = rule.f.map((x) => O.renameSheetRefs(x, s.name, name));
        for (const d of sh.drawings || []) if (d.refs) d.refs = d.refs.map((x) => O.renameSheetRefs(x, s.name, name));
        for (const d of sh.dv || []) if (d.f) d.f = d.f.map((x) => O.renameSheetRefs(x, s.name, name));
      }
      for (const k of Object.keys(state.names || {})) state.names[k] = O.renameSheetRefs(state.names[k], s.name, name);
      // La caché de una tabla dinámica nombra la hoja de origen.
      for (const p of state.pivotCaches || []) p.def = p.def.split(` sheet="${O.esc(s.name)}"`).join(` sheet="${O.esc(name)}"`);
      s.name = name;
      book.clearCache();
      renderTabs();
      paintAll();
      onChange();
    }
    const visibleSheets = () => state.sheets.filter((s) => !s.hidden);
    function setTabColor(color) {
      beforeChange();
      if (color) sheet().tabColor = color;
      else delete sheet().tabColor;
      renderTabs();
      onChange();
    }
    function hideSheet() {
      if (bookLocked()) return;
      if (visibleSheets().length < 2) { O.notice('Un libro necesita al menos una hoja visible.'); return; }
      beforeChange();
      sheet().hidden = true;
      state.active = state.sheets.findIndex((s) => !s.hidden);
      sel = { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
      build();
      onChange();
    }
    function showSheet(s) {
      if (bookLocked()) return;
      beforeChange();
      delete s.hidden;
      state.active = state.sheets.indexOf(s);
      sel = { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
      build();
      onChange();
    }
    // Proteger pide una contraseña opcional y lo que se permite; desproteger la comprueba, como en Excel.
    function askPassword(title, attrs, done) {
      if (!O.hasPassword(attrs)) { done(); return; }
      O.formDialog({
        title,
        fields: [{ name: 'pw', label: 'Contraseña', type: 'password', value: '' }],
        onOk: (v) => {
          const ok = O.checkProtection(attrs, v.pw);
          if (ok === null) return 'Oficina no puede comprobar esta contraseña: está cifrada con un método que no conoce.';
          if (!ok) return 'La contraseña no es correcta. Compruebe que la tecla BLOQ MAYÚS no está activada y que usa las mayúsculas correctas.';
          done();
          return null;
        },
      });
    }
    const PERMS = [['formatCells', 'Aplicar formato a celdas'], ['formatColumns', 'Aplicar formato a columnas'], ['formatRows', 'Aplicar formato a filas'],
      ['insertRows', 'Insertar filas'], ['insertColumns', 'Insertar columnas'], ['deleteRows', 'Eliminar filas'], ['deleteColumns', 'Eliminar columnas'],
      ['sort', 'Ordenar'], ['autoFilter', 'Usar Autofiltro']];
    function toggleProtect() {
      const s = sheet();
      if (s.protected) {
        askPassword('Desproteger hoja', s.protect, () => {
          beforeChange();
          delete s.protected;
          delete s.protect;
          renderTabs();
          paintSel();
          onChange();
        });
        return;
      }
      O.formDialog({
        title: 'Proteger hoja',
        fields: [
          { name: 'pw', label: 'Contraseña para desproteger la hoja (opcional)', type: 'password', value: '' },
          { name: 'pw2', label: 'Vuelva a escribir la contraseña', type: 'password', value: '' },
          ...PERMS.map(([name, label]) => ({ name, label, type: 'checkbox', value: false })),
        ],
        onOk: (v) => {
          if (v.pw !== v.pw2) return 'La contraseña de confirmación no es idéntica.';
          beforeChange();
          s.protect = { sheet: '1', objects: '1', scenarios: '1', ...O.newProtection(v.pw) };
          for (const [name] of PERMS) if (v[name]) s.protect[name] = '0';
          s.protected = true;
          renderTabs();
          paintSel();
          onChange();
          return null;
        },
      });
    }
    // El libro protegido usa los mismos datos con el prefijo workbook.
    const bookAttrs = () => {
      const p = state.wbProtect || {};
      return { algorithmName: p.workbookAlgorithmName, hashValue: p.workbookHashValue, saltValue: p.workbookSaltValue, spinCount: p.workbookSpinCount, password: p.workbookPassword };
    };
    function toggleBookProtect() {
      if (state.wbProtect && ['1', 'true'].includes(state.wbProtect.lockStructure)) {
        askPassword('Desproteger libro', bookAttrs(), () => {
          beforeChange();
          delete state.wbProtect;
          renderTabs();
          onChange();
        });
        return;
      }
      O.formDialog({
        title: 'Proteger estructura del libro',
        fields: [
          { name: 'pw', label: 'Contraseña (opcional)', type: 'password', value: '' },
          { name: 'pw2', label: 'Vuelva a escribir la contraseña', type: 'password', value: '' },
        ],
        onOk: (v) => {
          if (v.pw !== v.pw2) return 'La contraseña de confirmación no es idéntica.';
          beforeChange();
          const p = O.newProtection(v.pw);
          state.wbProtect = { lockStructure: '1' };
          if (p.hashValue) Object.assign(state.wbProtect, { workbookAlgorithmName: p.algorithmName, workbookHashValue: p.hashValue, workbookSaltValue: p.saltValue, workbookSpinCount: p.spinCount });
          renderTabs();
          onChange();
          return null;
        },
      });
    }
    function moveCopyDialog() {
      if (bookLocked()) return;
      const cur = sheet();
      O.formDialog({
        title: 'Mover o copiar',
        fields: [
          { name: 'before', label: 'Antes de la hoja', type: 'select', options: [...state.sheets.filter((s) => s !== cur).map((s) => [s.name, s.name]), ['', '(mover al final)']], value: '' },
          { name: 'copy', label: 'Crear una copia', type: 'checkbox', value: false },
        ],
        onOk: (v) => {
          beforeChange();
          let moving = cur;
          if (v.copy) {
            moving = JSON.parse(JSON.stringify(cur));
            moving.name = copyName(cur);
          } else state.sheets.splice(state.sheets.indexOf(cur), 1);
          const at = v.before ? state.sheets.findIndex((s) => s.name === v.before) : state.sheets.length;
          state.sheets.splice(at, 0, moving);
          state.active = state.sheets.indexOf(moving);
          book.clearCache();
          build();
          onChange();
          return null;
        },
      });
    }
    function copyName(s) {
      const base = s.name.slice(0, 26);
      let n = 2;
      while (state.sheets.some((o) => o.name.toLowerCase() === `${base} (${n})`.toLowerCase())) n++;
      return `${base} (${n})`;
    }
    const TAB_COLORS = [['Rojo', 'C0392B'], ['Naranja', 'E67E22'], ['Amarillo', 'F1C40F'], ['Verde', '2E7D5B'], ['Azul', '2B5F8A'], ['Morado', '7D3C98']];
    function addSheet() {
      if (bookLocked()) return;
      beforeChange();
      let n = state.sheets.length + 1;
      while (state.sheets.some((s) => s.name === `Hoja${n}`)) n++;
      state.sheets.splice((state.active || 0) + 1, 0, { name: `Hoja${n}`, cells: {}, cols: {} });
      state.active = (state.active || 0) + 1;
      book.clearCache();
      build();
      onChange();
    }
    function deleteSheet() {
      if (bookLocked()) return;
      if (visibleSheets().length < 2) { O.notice('Un libro necesita al menos una hoja visible.'); return; }
      if (!window.confirm(`¿Eliminar la hoja «${sheet().name}»?`)) return;
      beforeChange();
      state.sheets.splice(state.active || 0, 1);
      state.active = Math.max(0, (state.active || 0) - 1);
      if (state.sheets[state.active].hidden) state.active = Math.max(0, state.sheets.findIndex((s) => !s.hidden));
      book.clearCache();
      build();
      onChange();
    }
    function duplicateSheet() {
      if (bookLocked()) return;
      beforeChange();
      const s = sheet();
      const copy = JSON.parse(JSON.stringify(s));
      copy.name = copyName(s);
      state.sheets.splice((state.active || 0) + 1, 0, copy);
      state.active = (state.active || 0) + 1;
      book.clearCache();
      build();
      onChange();
    }
    function moveSheet(d) {
      if (bookLocked()) return;
      const i = state.active || 0;
      if (i + d < 0 || i + d >= state.sheets.length) return;
      beforeChange();
      const [s] = state.sheets.splice(i, 1);
      state.sheets.splice(i + d, 0, s);
      state.active = i + d;
      renderTabs();
      onChange();
    }
    function renderTabs() {
      const box = q('.sheet-tabs');
      box.replaceChildren();
      state.sheets.forEach((s, i) => {
        if (s.hidden) return;
        const b = document.createElement('button');
        b.className = 'sheet-tab' + (i === (state.active || 0) ? ' active' : '');
        b.textContent = (s.protected ? '🔒 ' : '') + s.name;
        if (s.tabColor) {
          b.dataset.color = s.tabColor;
          b.style.setProperty('--tab-mark', '#' + s.tabColor);
        }
        b.title = 'Doble clic para cambiar el nombre. Clic derecho para más opciones';
        b.onclick = () => activate(i);
        b.ondblclick = () => renameSheet(s);
        b.oncontextmenu = (e) => {
          e.preventDefault();
          activate(i);
          const last = state.sheets.length - 1;
          const hiddenOnes = state.sheets.filter((o) => o.hidden);
          O.contextMenu(e.clientX, e.clientY, [
            { label: 'Insertar hoja', action: addSheet },
            { label: 'Eliminar hoja', disabled: visibleSheets().length < 2, action: deleteSheet },
            { label: 'Cambiar nombre…', action: () => renameSheet(s) },
            { label: 'Mover o copiar…', action: moveCopyDialog },
            { label: 'Duplicar hoja', action: duplicateSheet },
            { label: 'Mover a la izquierda', disabled: i === 0, action: () => moveSheet(-1) },
            { label: 'Mover a la derecha', disabled: i === last, action: () => moveSheet(1) },
            '-',
            { label: s.protected ? 'Desproteger hoja…' : 'Proteger hoja…', action: toggleProtect },
            { label: state.wbProtect && ['1', 'true'].includes(state.wbProtect.lockStructure) ? 'Desproteger libro…' : 'Proteger libro…', action: toggleBookProtect },
            { label: 'Color de etiqueta', items: [...TAB_COLORS.map(([label, color]) => ({ label, action: () => setTabColor(color) })), '-', { label: 'Sin color', action: () => setTabColor(null) }] },
            '-',
            { label: 'Ocultar', disabled: visibleSheets().length < 2, action: hideSheet },
            hiddenOnes.length
              ? { label: 'Mostrar', items: hiddenOnes.map((o) => ({ label: o.name, action: () => showSheet(o) })) }
              : { label: 'Mostrar', disabled: true, action: () => {} },
          ]);
        };
        box.append(b);
      });
      const add = document.createElement('button');
      add.className = 'tb';
      add.textContent = '+ Hoja';
      add.onclick = addSheet;
      box.append(add);
      if (state.sheets.length > 1) {
        const del = document.createElement('button');
        del.className = 'tb';
        del.textContent = 'Eliminar hoja';
        del.onclick = deleteSheet;
        box.append(del);
      }
    }

    // Al imprimir: solo el area con datos de la hoja activa, sin encabezados ni seleccion.
    let printBox = null;
    const beforePrint = () => {
      if (root.hidden) return;
      if (editor) commit(0, 0, true);
      const { mr, mc } = used();
      const table = document.createElement('table');
      const cols = document.createElement('colgroup');
      for (let c = 0; c <= mc; c++) {
        const col = document.createElement('col');
        col.style.width = ((sheet().cols && sheet().cols[c]) || sheet().defW || DEFAULT_W) + 'px';
        cols.append(col);
      }
      table.append(cols);
      const body = document.createElement('tbody');
      const off = sheet().hiddenCols || {};
      for (let c = 0; c <= mc; c++) if (off[c]) cols.children[c].style.visibility = 'collapse';
      for (let r = 0; r <= mr; r++) {
        if (sheet().hiddenRows && sheet().hiddenRows[r]) continue;
        const tr = document.createElement('tr');
        for (let c = 0; c <= mc; c++) {
          const td = tds[r][c].cloneNode(true);
          td.classList.remove('sel', 'cur');
          tr.append(td);
        }
        body.append(tr);
      }
      table.append(body);
      printBox = document.createElement('div');
      printBox.className = 'print-sheet';
      printBox.append(table);
      root.append(printBox);
    };
    const afterPrint = () => {
      if (printBox) printBox.remove();
      printBox = null;
    };
    window.addEventListener('beforeprint', beforePrint);
    window.addEventListener('afterprint', afterPrint);

    state.active = state.active || 0;
    build();

    return {
      getState: () => state,
      flush: () => { if (editor) commit(0, 0, true); },
      onShow: () => {
        // Al montar, la pestaña aún está oculta y no hay medidas: los dibujos y las flechas se colocan al mostrarla.
        placeFilterButtons();
        renderDrawings();
        placeListButton();
        if (!editor) wrap.focus({ preventScroll: true });
      },
      destroy: () => {
        window.removeEventListener('beforeprint', beforePrint);
        window.removeEventListener('afterprint', afterPrint);
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        suggest.remove();
        closeFilterPanel();
      },
    };
  }

  O.apps.sheet = { mount };
})();
