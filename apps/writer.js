// Procesador de textos: edicion sobre HTML con formato que docx.js convierte al guardar. La interfaz imita a Word.
(function () {
  const O = window.Oficina;
  const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  const ALLOWED = new Set(['P', 'DIV', 'BR', 'H1', 'H2', 'H3', 'UL', 'OL', 'LI', 'TABLE', 'TBODY', 'TR', 'TD', 'TH', 'B', 'STRONG', 'I', 'EM', 'U', 'S', 'SUB', 'SUP', 'SPAN', 'A', 'IMG', 'HR']);
  const KEEP_CSS = ['font-family', 'font-weight', 'font-style', 'text-decoration', 'color', 'font-size', 'text-align', 'background-color', 'line-height', 'margin-left'];
  // Estilos de un documento nuevo de Word 2024: Aptos 11, 8 pt despues y 1,08 de interlineado; titulos en azul.
  const DEFAULTS = {
    base: { font: 'Aptos', size: 11, after: 8, line: 1.08 },
    heads: {
      H1: { font: 'Aptos Display', size: 20, color: '0F4761', b: false, before: 18, after: 4 },
      H2: { font: 'Aptos Display', size: 16, color: '0F4761', b: false, before: 8, after: 4 },
      H3: { font: 'Aptos', size: 14, color: '0F4761', b: false, before: 8, after: 4 },
    },
  };
  // Altura de una linea sencilla respecto al tamano de letra (Aptos y Calibri rondan 1,2).
  const LINE = 1.2;
  // A4 vertical con los margenes de Word 2024, en twips; hueco gris entre hojas en px.
  const PAGE_DEF = { w: 11906, h: 16838, top: 1417, right: 1701, bottom: 1417, left: 1701 };
  const PAGE_GAP = 16;
  // Margenes y tamanos de la pestana Disposicion, con los nombres de Word.
  const MARGINS = [
    ['Normal', 'Superior: 2,5 cm  Inferior: 2,5 cm  Izquierdo: 3 cm  Derecho: 3 cm', 1417, 1701],
    ['Estrecho', 'Superior: 1,27 cm  Inferior: 1,27 cm  Izquierdo: 1,27 cm  Derecho: 1,27 cm', 720, 720],
    ['Moderado', 'Superior: 2,54 cm  Inferior: 2,54 cm  Izquierdo: 1,91 cm  Derecho: 1,91 cm', 1440, 1080],
    ['Ancho', 'Superior: 2,54 cm  Inferior: 2,54 cm  Izquierdo: 5,08 cm  Derecho: 5,08 cm', 1440, 2880],
  ];
  const PAPERS = [
    ['Carta', '21,59 cm x 27,94 cm', 12240, 15840],
    ['A4', '21 cm x 29,7 cm', 11906, 16838],
    ['A5', '14,8 cm x 21 cm', 8391, 11906],
    ['Legal', '21,59 cm x 35,56 cm', 12240, 20160],
  ];
  O.WORD_DEFAULTS = DEFAULTS;
  O.WORD_LINE = LINE;

  // Limpia lo pegado: conserva el formato basico y elimina scripts, clases y estilos ajenos.
  function sanitize(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const walk = (node) => {
      for (const c of [...node.children]) {
        walk(c);
        if (!ALLOWED.has(c.tagName)) {
          if (['SCRIPT', 'STYLE', 'META', 'LINK', 'TITLE', 'IFRAME', 'OBJECT', 'EMBED'].includes(c.tagName)) c.remove();
          else c.replaceWith(...c.childNodes);
          continue;
        }
        const css = KEEP_CSS.map((p) => (c.style.getPropertyValue(p) ? `${p}:${c.style.getPropertyValue(p)}` : '')).filter(Boolean).join(';');
        for (const a of [...c.attributes]) {
          const keep = (c.tagName === 'A' && a.name === 'href' && /^(https?:|mailto:)/i.test(a.value))
            || (c.tagName === 'IMG' && ['width', 'height'].includes(a.name))
            || (c.tagName === 'IMG' && a.name === 'src' && a.value.startsWith('data:image/'))
            || (['TD', 'TH'].includes(c.tagName) && a.name === 'colspan');
          if (!keep) c.removeAttribute(a.name);
        }
        if (css) c.setAttribute('style', css);
      }
    };
    walk(doc.body);
    return doc.body.innerHTML;
  }

  function ribbon() {
    const rb = O.rb;
    const style = (tag, name) => `<button class="rs" data-style="${tag}" title="${name}"><span class="rs-sample rs-${tag}">AaBbCc</span><span class="rs-name">${tag === 'P' ? '¶ ' : ''}${name}</span></button>`;
    return O.ribbon([
      { name: 'Inicio', groups: [
        { label: 'Portapapeles', html: `${rb('w-paste', 'paste', 'Pegar', 'Pegar (Ctrl+V)', { big: true })}
          <div class="rcol">${rb('w-cut', 'cut', 'Cortar', 'Cortar (Ctrl+X)', { text: 'Cortar' })}${rb('w-copy', 'copy', 'Copiar', 'Copiar (Ctrl+C)', { text: 'Copiar' })}${rb('w-painter', 'brush', 'Copiar formato', 'Copiar formato: aplica el formato del texto actual al que selecciones después', { text: 'Copiar formato' })}</div>` },
        { label: 'Fuente', html: `<div class="rcol">
          <div class="rrow"><select id="w-font" class="r-font" aria-label="Fuente">${O.installedFonts().map(O.fontOption).join('')}</select><select id="w-size" class="r-size" aria-label="Tamaño de fuente">${SIZES.map((s) => `<option value="${s}"${s === 11 ? ' selected' : ''}>${String(s).replace('.', ',')}</option>`).join('')}</select>${rb('w-grow', 'grow', 'Aumentar tamaño de fuente', 'Aumentar tamaño de fuente (Ctrl+>)')}${rb('w-shrink', 'shrink', 'Disminuir tamaño de fuente', 'Disminuir tamaño de fuente (Ctrl+<)')}${rb('w-case', 'case', 'Cambiar mayúsculas y minúsculas', '', { menu: true })}${rb('w-clear', 'clear', 'Borrar todo el formato')}</div>
          <div class="rrow"><button class="rb" data-cmd="bold" title="Negrita (Ctrl+N)" aria-label="Negrita"><b class="glyph">N</b></button><button class="rb" data-cmd="italic" title="Cursiva (Ctrl+K)" aria-label="Cursiva"><i class="glyph serif">K</i></button><button class="rb" data-cmd="underline" title="Subrayado (Ctrl+S)" aria-label="Subrayado"><u class="glyph">S</u></button><button class="rb" data-cmd="strikeThrough" title="Tachado" aria-label="Tachado"><s class="glyph">abc</s></button><button class="rb" data-cmd="subscript" title="Subíndice (Ctrl+=)" aria-label="Subíndice"><span class="glyph">x<sub>2</sub></span></button><button class="rb" data-cmd="superscript" title="Superíndice (Ctrl+Mayús++)" aria-label="Superíndice"><span class="glyph">x<sup>2</sup></span></button><span class="rsep"></span>${O.rbColor('w-hl', '<span class="hl-pen">ab</span>', 'Color de resaltado del texto', '#FFFF00')}${O.rbColor('w-color', 'A', 'Color de fuente', '#C00000')}</div></div>` },
        { label: 'Párrafo', html: `<div class="rcol">
          <div class="rrow">${rb('', 'bullets', 'Viñetas', '', { attrs: 'data-cmd="insertUnorderedList"' })}${rb('', 'numbers', 'Numeración', '', { attrs: 'data-cmd="insertOrderedList"' })}${rb('w-outdent', 'outdent', 'Disminuir sangría')}${rb('w-indent', 'indent', 'Aumentar sangría')}<span class="rsep"></span>${rb('w-marks', 'marks', 'Mostrar todo', 'Mostrar todo: marcas de párrafo y otros símbolos de formato ocultos (Ctrl+*)')}</div>
          <div class="rrow">${rb('', 'left', 'Alinear a la izquierda', 'Alinear a la izquierda (Ctrl+Q)', { attrs: 'data-cmd="justifyLeft"' })}${rb('', 'center', 'Centrar', 'Centrar (Ctrl+T)', { attrs: 'data-cmd="justifyCenter"' })}${rb('', 'right', 'Alinear a la derecha', 'Alinear a la derecha (Ctrl+D)', { attrs: 'data-cmd="justifyRight"' })}${rb('', 'justify', 'Justificar', 'Justificar (Ctrl+J)', { attrs: 'data-cmd="justifyFull"' })}<span class="rsep"></span>${rb('w-spacing', 'spacing', 'Espaciado entre líneas y párrafos', '', { menu: true })}</div></div>` },
        { label: 'Estilos', html: `<div class="rstyles">${style('P', 'Normal')}${style('H1', 'Título 1')}${style('H2', 'Título 2')}${style('H3', 'Título 3')}</div>` },
        { label: 'Edición', html: `<div class="rcol">${rb('w-find', 'search', 'Buscar', 'Buscar (Ctrl+B)', { text: 'Buscar' })}${rb('w-replace', 'replace', 'Reemplazar', 'Reemplazar (Ctrl+L)', { text: 'Reemplazar' })}${rb('w-select', 'select', 'Seleccionar', '', { text: 'Seleccionar', menu: true })}</div>` },
      ] },
      { name: 'Insertar', groups: [
        { label: 'Páginas', html: rb('w-break', 'pagebreak', 'Salto de página', 'Salto de página (Ctrl+Entrar)', { big: true }) },
        { label: 'Tablas', html: rb('w-table', 'table', 'Tabla', 'Insertar tabla', { big: true, menu: true }) },
        { label: 'Ilustraciones', html: `${rb('w-image', 'image', 'Imágenes', 'Insertar imágenes de este equipo', { big: true })}<input type="file" id="w-image-file" accept="image/png,image/jpeg,image/gif" hidden>` },
        { label: 'Vínculos', html: rb('w-link', 'link', 'Vínculo', 'Insertar vínculo (Ctrl+Alt+K)', { big: true }) },
      ] },
      { name: 'Disposición', groups: [
        { label: 'Configurar página', html: `${rb('w-margins', 'margins', 'Márgenes', 'Márgenes de la página', { big: true, menu: true })}${rb('w-orient', 'orient', 'Orientación', 'Orientación de la página', { big: true, menu: true })}${rb('w-papersize', 'papersize', 'Tamaño', 'Tamaño del papel', { big: true, menu: true })}${rb('w-break2', 'pagebreak', 'Saltos', 'Salto de página (Ctrl+Entrar)', { big: true })}` },
      ] },
      { name: 'Vista', groups: [
        { label: 'Vistas', html: `${rb('w-printview', 'page', 'Diseño de impresión', 'Ver las hojas como se imprimen', { big: true })}${rb('w-mobileview', 'phone', 'Vista móvil', 'Ajustar el texto al ancho de la pantalla, a tamaño de lectura', { big: true })}` },
        { label: 'Mostrar', html: rb('w-marks2', 'marks', 'Marcas de formato', 'Mostrar marcas de párrafo y saltos', { big: true }) },
        { label: 'Zoom', html: `${rb('w-zoom100', 'zoom', '100 %', 'Zoom al 100 %', { big: true })}${rb('w-onepage', 'page', 'Una página', 'Ver una página entera', { big: true })}${rb('w-pagewidth', 'fit', 'Ancho de página', 'Ajustar al ancho de la ventana', { big: true })}` },
      ] },
    ]);
  }

  function mount(root, state, onChange) {
    const base = { ...DEFAULTS.base, ...(state.base || {}) };
    const heads = { ...DEFAULTS.heads, ...(state.heads || {}) };
    root.innerHTML = `
      ${ribbon()}
      <div class="writer-scroll"><div class="page" contenteditable="true" spellcheck="true" lang="es" role="textbox" aria-multiline="true" aria-label="Documento"></div></div>
      <div class="statusbar office">
        <span class="w-pages">Página 1 de 1</span><span class="w-count"></span><span class="w-lang" title="Idioma de revisión">Español (España)</span>
        <span class="st-fill"></span>${O.zoomControl('w-zoom')}
      </div>`;
    const q = (s) => root.querySelector(s);
    const page = q('.page');
    const scroller = q('.writer-scroll');
    page.innerHTML = state.html;
    O.ribbonInit(root);
    // Estilo base y de titulos del documento, para verlo con su fuente, como en Word.
    const css = (k, v) => page.style.setProperty(k, v);
    css('--w-font', `'${base.font}'`);
    css('--w-size', base.size + 'pt');
    css('--w-after', base.after + 'pt');
    css('--w-line', String(base.line * LINE));
    for (const [tag, h] of Object.entries(heads)) {
      const t = tag.toLowerCase();
      css(`--${t}-font`, `'${h.font || base.font}'`);
      css(`--${t}-size`, h.size + 'pt');
      css(`--${t}-color`, '#' + (h.color || '000000'));
      css(`--${t}-weight`, h.b ? '700' : '400');
      css(`--${t}-before`, (h.before || 0) + 'pt');
      css(`--${t}-after`, (h.after || 0) + 'pt');
    }
    let saved = null;
    let zoom = 100;

    // Pagina: tamano y margenes en twips (1 px = 15 twips), como en el sectPr de Word.
    const pg = { ...PAGE_DEF, ...(state.page || {}) };
    const px = (t) => t / 15;
    const applyPage = () => {
      css('--pg-w', px(pg.w) + 'px');
      css('--pg-h', px(pg.h) + 'px');
      css('--pg-t', px(pg.top) + 'px');
      css('--pg-r', px(pg.right) + 'px');
      css('--pg-b', px(pg.bottom) + 'px');
      css('--pg-l', px(pg.left) + 'px');
    };
    applyPage();

    // ---------- Paginas: como Word, cada parrafo que no cabe pasa a la hoja siguiente ----------
    let pageCount = 1;
    const relTop = (el) => (el.getBoundingClientRect().top - page.getBoundingClientRect().top) / (zoom / 100);
    const unitsOf = () => {
      const out = [];
      for (const b of page.children) {
        if (b.matches('ul, ol')) out.push(...b.children);
        else out.push(b);
      }
      return out;
    };
    const clearPush = () => {
      for (const el of page.querySelectorAll('[data-push]')) {
        el.style.marginTop = el.dataset.push;
        el.removeAttribute('data-push');
        if (!el.getAttribute('style')) el.removeAttribute('style');
      }
    };
    const pushTo = (el, target) => {
      if (!el.hasAttribute('data-push')) el.dataset.push = el.style.marginTop;
      for (let i = 0; i < 2; i++) {
        const d = target - relTop(el);
        if (Math.abs(d) < 0.5) break;
        el.style.marginTop = (parseFloat(getComputedStyle(el).marginTop) || 0) + d + 'px';
      }
    };
    const paginate = () => {
      if (!page.isConnected || !page.offsetWidth) return;
      if (page.classList.contains('mobile-view')) { clearPush(); page.style.minHeight = ''; return; }
      clearPush();
      const H = px(pg.h), T = px(pg.top), B = px(pg.bottom), S = H + PAGE_GAP;
      let forceAfter = -1;
      let bottom = 0;
      for (const el of unitsOf()) {
        const top = relTop(el);
        if (el.matches('hr.salto')) { forceAfter = Math.floor(top / S); continue; }
        const h = el.offsetHeight;
        const k = Math.floor(top / S);
        let target = null;
        if (forceAfter >= 0 && k <= forceAfter) target = (forceAfter + 1) * S + T;
        else if (top > k * S + H - B - 0.5) target = (k + 1) * S + T;
        else if (top + h > k * S + H - B + 0.5 && top > k * S + T + 0.5 && h <= H - T - B) target = (k + 1) * S + T;
        forceAfter = -1;
        if (target !== null && target > top) pushTo(el, target);
        bottom = Math.max(bottom, relTop(el) + el.offsetHeight);
      }
      pageCount = Math.max(1, Math.floor(Math.max(0, bottom - 1) / S) + 1, forceAfter >= 0 ? forceAfter + 2 : 1);
      page.style.minHeight = pageCount * S - PAGE_GAP + 'px';
    };
    let pending = 0;
    const repaginate = () => {
      cancelAnimationFrame(pending);
      pending = requestAnimationFrame(() => { paginate(); updateCount(); });
    };
    const cleanHtml = () => {
      const c = page.cloneNode(true);
      for (const el of c.querySelectorAll('[data-push]')) {
        el.style.marginTop = el.dataset.push;
        el.removeAttribute('data-push');
        if (!el.getAttribute('style')) el.removeAttribute('style');
      }
      return c.innerHTML;
    };

    const changed = () => { repaginate(); onChange(); };
    const updateCount = () => {
      const text = page.innerText.trim();
      const words = text ? text.split(/\s+/).length : 0;
      const sel = window.getSelection();
      const part = sel.rangeCount && !sel.isCollapsed && page.contains(sel.anchorNode) ? sel.toString().trim() : '';
      const label = (n) => `${n.toLocaleString('es-ES')} ${n === 1 ? 'palabra' : 'palabras'}`;
      q('.w-count').textContent = part ? `${part.split(/\s+/).length.toLocaleString('es-ES')} de ${label(words)}` : label(words);
      let at = 1;
      if (sel.rangeCount && page.contains(sel.anchorNode)) {
        const r = sel.getRangeAt(0).getBoundingClientRect();
        const el = sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement;
        const y = ((r.top || el.getBoundingClientRect().top) - page.getBoundingClientRect().top) / (zoom / 100);
        at = Math.min(pageCount, Math.max(1, Math.floor(y / (px(pg.h) + PAGE_GAP)) + 1));
      }
      q('.w-pages').textContent = page.classList.contains('mobile-view') ? 'Vista móvil' : `Página ${at} de ${pageCount}`;
    };
    const remember = () => {
      const s = window.getSelection();
      if (s.rangeCount && page.contains(s.anchorNode)) saved = s.getRangeAt(0).cloneRange();
    };
    const restore = () => {
      page.focus({ preventScroll: true });
      if (saved) {
        const s = window.getSelection();
        s.removeAllRanges();
        s.addRange(saved);
      }
    };
    const exec = (cmd, val) => {
      restore();
      document.execCommand('styleWithCSS', false, true);
      document.execCommand(cmd, false, val);
      remember();
      changed();
      refresh();
    };
    const insertHtml = (html) => exec('insertHTML', html);
    const anchorEl = () => {
      const s = window.getSelection();
      if (!s.anchorNode || !page.contains(s.anchorNode)) return null;
      return s.anchorNode.nodeType === 1 ? s.anchorNode : s.anchorNode.parentElement;
    };
    // Parrafos tocados por la seleccion.
    const selectedBlocks = () => {
      restore();
      const s = window.getSelection();
      if (!s.rangeCount) return [];
      const r = s.getRangeAt(0);
      const all = [...page.querySelectorAll('p,h1,h2,h3,li,div')].filter((b) => !b.querySelector('p,h1,h2,h3,li,div,table'));
      const hit = all.filter((b) => r.intersectsNode(b));
      if (hit.length) return hit;
      const el = anchorEl();
      const b = el && el.closest('p,h1,h2,h3,li,div');
      return b && page.contains(b) && b !== page ? [b] : [];
    };
    const currentSize = () => {
      const el = anchorEl();
      return el ? Math.round(parseFloat(getComputedStyle(el).fontSize) * 0.75 * 2) / 2 : base.size;
    };
    // Tamano de letra en la seleccion: execCommand solo sabe de 7 tamanos, asi que se marca y se cambia por puntos.
    const setSize = (pt) => {
      restore();
      document.execCommand('styleWithCSS', false, false);
      document.execCommand('fontSize', false, '7');
      const spans = [...page.querySelectorAll('font[size="7"]')].map((f) => {
        const span = document.createElement('span');
        span.style.fontSize = pt + 'pt';
        span.append(...f.childNodes);
        f.replaceWith(span);
        return span;
      });
      // Al cambiar las etiquetas se pierde la seleccion: se vuelve a marcar el texto cambiado, como hace Word.
      if (spans.length) {
        const r = document.createRange();
        r.setStartBefore(spans[0].firstChild || spans[0]);
        r.setEndAfter(spans[spans.length - 1].lastChild || spans[spans.length - 1]);
        const s = window.getSelection();
        s.removeAllRanges();
        s.addRange(r);
      }
      remember();
      changed();
      refresh();
    };

    const refresh = () => {
      for (const b of root.querySelectorAll('[data-cmd]')) b.classList.toggle('on', document.queryCommandState(b.dataset.cmd));
      const el = anchorEl();
      updateCount();
      if (!el) return;
      const block = el.closest('h1,h2,h3');
      const tag = block ? block.tagName : 'P';
      root.querySelectorAll('[data-style]').forEach((b) => b.classList.toggle('on', b.dataset.style === tag));
      const cs = getComputedStyle(el);
      const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
      const sizeSel = q('#w-size');
      if (![...sizeSel.options].some((o) => +o.value === pt)) {
        const o = document.createElement('option');
        o.value = pt;
        o.textContent = String(pt).replace('.', ',');
        sizeSel.insertBefore(o, [...sizeSel.options].find((x) => +x.value > pt) || null);
      }
      sizeSel.value = pt;
      const family = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (family) { O.addFontOption(q('#w-font'), family); q('#w-font').value = family; }
    };

    // ---------- Portapapeles ----------
    q('#w-copy').onclick = () => { restore(); document.execCommand('copy'); };
    q('#w-cut').onclick = () => { restore(); document.execCommand('cut'); changed(); };
    q('#w-paste').onclick = async () => {
      try {
        const items = await navigator.clipboard.read();
        for (const it of items) {
          if (it.types.includes('text/html')) { insertHtml(sanitize(await (await it.getType('text/html')).text())); return; }
          if (it.types.includes('text/plain')) { restore(); document.execCommand('insertText', false, await (await it.getType('text/plain')).text()); changed(); return; }
        }
      } catch {
        O.notice('El navegador no deja pegar desde este botón. Usa Ctrl+V.');
      }
    };
    // Copiar formato: guarda el formato del cursor y lo aplica a la siguiente seleccion.
    let painter = null;
    q('#w-painter').onclick = () => {
      const el = anchorEl();
      if (!el) return;
      const cs = getComputedStyle(el);
      painter = {
        font: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(), size: parseFloat(cs.fontSize) * 0.75, color: cs.color,
        b: parseInt(cs.fontWeight, 10) >= 600, i: cs.fontStyle === 'italic',
        u: !!el.closest('u') || cs.textDecorationLine.includes('underline'), s: !!el.closest('s') || cs.textDecorationLine.includes('line-through'),
        bg: (el.closest('[style*="background"]') || {}).style ? el.closest('[style*="background"]').style.backgroundColor : '',
      };
      q('#w-painter').classList.add('on');
      page.classList.add('painting');
    };
    const applyPainter = () => {
      const p = painter;
      painter = null;
      q('#w-painter').classList.remove('on');
      page.classList.remove('painting');
      const s = window.getSelection();
      if (s.isCollapsed) return;
      remember();
      exec('removeFormat');
      exec('fontName', p.font);
      exec('foreColor', p.color);
      if (p.bg) exec('hiliteColor', p.bg);
      for (const [k, cmd] of [['b', 'bold'], ['i', 'italic'], ['u', 'underline'], ['s', 'strikeThrough']]) if (p[k] !== document.queryCommandState(cmd)) exec(cmd);
      setSize(p.size);
    };

    // ---------- Fuente ----------
    root.querySelectorAll('[data-cmd]').forEach((b) => { b.onclick = () => exec(b.dataset.cmd); });
    q('#w-font').onchange = (e) => exec('fontName', e.target.value);
    q('#w-size').onchange = (e) => setSize(+e.target.value);
    const step = (up) => {
      const now = currentSize();
      const next = up ? SIZES.find((s) => s > now) || now + 10 : [...SIZES].reverse().find((s) => s < now) || Math.max(1, now - 1);
      setSize(next);
    };
    q('#w-grow').onclick = () => step(true);
    q('#w-shrink').onclick = () => step(false);
    q('#w-clear').onclick = () => {
      exec('removeFormat');
      // removeFormat no quita el tamano puesto con estilo: se limpian los span sin otro formato.
      const s = window.getSelection();
      if (!s.rangeCount) return;
      const r = s.getRangeAt(0);
      page.querySelectorAll('span[style]').forEach((sp) => { if (r.intersectsNode(sp)) sp.style.fontSize = sp.style.fontFamily = ''; });
      changed();
    };
    // Cambiar mayusculas y minusculas sobre los nodos de texto de la seleccion, conservando el formato.
    const transformSelection = (fn) => {
      restore();
      const s = window.getSelection();
      if (!s.rangeCount || s.isCollapsed) return;
      const r = s.getRangeAt(0);
      let sc = r.startContainer, so = r.startOffset, ec = r.endContainer, eo = r.endOffset;
      if (ec.nodeType === 3 && eo < ec.length) ec.splitText(eo);
      if (sc.nodeType === 3 && so > 0) {
        const tail = sc.splitText(so);
        if (ec === sc) { ec = tail; eo = tail.length; }
        sc = tail;
        so = 0;
      }
      const range = document.createRange();
      range.setStart(sc, so);
      range.setEnd(ec, ec.nodeType === 3 ? ec.length : eo);
      const walker = document.createTreeWalker(range.commonAncestorContainer.nodeType === 3 ? range.commonAncestorContainer.parentNode : range.commonAncestorContainer, NodeFilter.SHOW_TEXT);
      const nodes = [];
      while (walker.nextNode()) if (range.intersectsNode(walker.currentNode) && walker.currentNode.length) nodes.push(walker.currentNode);
      const all = nodes.map((n) => n.nodeValue).join('');
      const out = fn(all);
      let at = 0;
      for (const n of nodes) { n.nodeValue = out.slice(at, at + n.length); at += n.length; }
      s.removeAllRanges();
      s.addRange(range);
      remember();
      changed();
    };
    q('#w-case').onclick = (e) => O.menuAt(e.currentTarget, [
      { label: 'Tipo oración.', action: () => transformSelection((t) => t.toLocaleLowerCase('es').replace(/(^\s*|[.!?¡¿]\s+)(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase('es'))) },
      { label: 'minúsculas', action: () => transformSelection((t) => t.toLocaleLowerCase('es')) },
      { label: 'MAYÚSCULAS', action: () => transformSelection((t) => t.toLocaleUpperCase('es')) },
      { label: 'Poner En Mayúsculas Cada Palabra', action: () => transformSelection((t) => t.toLocaleLowerCase('es').replace(/(^|[^\p{L}])(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase('es'))) },
      { label: 'tIPO iNVERSO', action: () => transformSelection((t) => [...t].map((c) => (c === c.toLocaleUpperCase('es') ? c.toLocaleLowerCase('es') : c.toLocaleUpperCase('es'))).join('')) },
    ]);
    let textColor = 'C00000', hlColor = 'FFFF00';
    const applyColor = (v) => { exec('foreColor', v === 'auto' ? '#000000' : '#' + v); };
    const applyHl = (v) => { exec('hiliteColor', v === 'none' ? 'transparent' : '#' + v); };
    q('#w-color').onclick = () => applyColor(textColor);
    q('#w-color-menu').onclick = (e) => O.colorPalette(e.currentTarget, 'theme', (v) => {
      if (v !== 'auto') { textColor = v; O.setColorBar(q('#w-color'), v); }
      applyColor(v);
    }, { auto: 'Automático' });
    q('#w-hl').onclick = () => applyHl(hlColor);
    q('#w-hl-menu').onclick = (e) => O.colorPalette(e.currentTarget, 'highlight', (v) => {
      if (v !== 'none') { hlColor = v; O.setColorBar(q('#w-hl'), v); }
      applyHl(v);
    }, { none: 'Sin color' });

    // ---------- Parrafo ----------
    const INDENT = 35.4; // 1,25 cm, el paso de Word
    const indent = (dir) => {
      const el = anchorEl();
      if (el && el.closest('li')) { exec(dir > 0 ? 'indent' : 'outdent'); return; }
      for (const b of selectedBlocks()) {
        const now = parseFloat(b.style.marginLeft) || 0;
        const next = Math.max(0, Math.round((now / INDENT + dir)) * INDENT);
        b.style.marginLeft = next ? next.toFixed(1) + 'pt' : '';
      }
      remember();
      changed();
    };
    q('#w-indent').onclick = () => indent(1);
    q('#w-outdent').onclick = () => indent(-1);
    const toggleMarks = () => {
      const on = page.classList.toggle('marks');
      q('#w-marks').classList.toggle('on', on);
      q('#w-marks2').classList.toggle('on', on);
      repaginate();
    };
    q('#w-marks').onclick = toggleMarks;
    q('#w-marks2').onclick = toggleMarks;
    q('#w-spacing').onclick = (e) => {
      const blocks = selectedBlocks();
      const cur = blocks.length ? parseFloat(blocks[0].style.lineHeight) / LINE || base.line : base.line;
      const setLine = (m) => {
        for (const b of blocks) b.style.lineHeight = Math.abs(m - base.line) < 0.001 ? '' : (m * LINE).toFixed(3);
        remember();
        changed();
      };
      const setAfter = (pt) => {
        for (const b of blocks) b.style.marginBottom = pt === null ? '' : pt + 'pt';
        remember();
        changed();
      };
      const noSpace = blocks.length && blocks.every((b) => b.style.marginBottom === '0pt');
      O.menuAt(e.currentTarget, [
        ...[1, 1.08, 1.15, 1.5, 2, 2.5, 3].map((m) => ({ label: (Math.abs(cur - m) < 0.01 ? '✓ ' : ' ') + String(m.toFixed(m === 1.08 || m === 1.15 ? 2 : 1)).replace('.', ','), action: () => setLine(m) })),
        '-',
        noSpace ? { label: 'Agregar espacio después del párrafo', action: () => setAfter(null) } : { label: 'Quitar espacio después del párrafo', action: () => setAfter(0) },
      ]);
    };

    // ---------- Estilos ----------
    root.querySelectorAll('[data-style]').forEach((b) => { b.onclick = () => exec('formatBlock', b.dataset.style); });

    // ---------- Edicion: buscar y reemplazar ----------
    const findBox = document.createElement('div');
    findBox.className = 'find-box';
    findBox.hidden = true;
    findBox.innerHTML = `<div class="find-head"><b>Buscar y reemplazar</b><button class="linkish" data-close aria-label="Cerrar">✕</button></div>
      <label>Buscar:<input data-find></label><label class="rep">Reemplazar con:<input data-rep></label>
      <div class="find-row"><span class="find-info" role="status"></span><button class="btn" data-next>Buscar siguiente</button><button class="btn rep" data-one>Reemplazar</button><button class="btn rep" data-all>Reemplazar todos</button></div>`;
    root.append(findBox);
    const fq = (s) => findBox.querySelector(s);
    const openFind = (replace) => {
      findBox.hidden = false;
      findBox.classList.toggle('only-find', !replace);
      const s = window.getSelection();
      if (s.rangeCount && !s.isCollapsed && page.contains(s.anchorNode) && s.toString().length < 80) fq('[data-find]').value = s.toString();
      fq('[data-find]').focus();
      fq('[data-find]').select();
      fq('.find-info').textContent = '';
    };
    fq('[data-close]').onclick = () => { findBox.hidden = true; restore(); };
    const findNext = () => {
      const text = fq('[data-find]').value;
      if (!text) return false;
      restore();
      // window.find busca desde la seleccion y da la vuelta; se limita a lo que hay dentro de la pagina.
      let ok = window.find(text, false, false, true, false, false, false);
      if (ok && !page.contains(window.getSelection().anchorNode)) ok = window.find(text, false, false, true, false, false, false) && page.contains(window.getSelection().anchorNode);
      remember();
      fq('.find-info').textContent = ok ? '' : 'No se encontró el texto.';
      return ok;
    };
    fq('[data-next]').onclick = findNext;
    fq('[data-find]').onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); findNext(); } };
    fq('[data-one]').onclick = () => {
      const text = fq('[data-find]').value;
      restore();
      const s = window.getSelection();
      if (text && s.toString().toLocaleLowerCase('es') === text.toLocaleLowerCase('es')) {
        document.execCommand('insertText', false, fq('[data-rep]').value);
        remember();
        changed();
      }
      findNext();
    };
    fq('[data-all]').onclick = () => {
      const text = fq('[data-find]').value;
      if (!text) return;
      const rx = new RegExp(text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
      let n = 0;
      const walker = document.createTreeWalker(page, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const t = walker.currentNode;
        const v = t.nodeValue.replace(rx, () => { n++; return fq('[data-rep]').value; });
        if (v !== t.nodeValue) t.nodeValue = v;
      }
      fq('.find-info').textContent = n ? (n === 1 ? 'Se ha realizado 1 reemplazo.' : `Se han realizado ${n} reemplazos.`) : 'No se encontró el texto.';
      if (n) changed();
    };
    q('#w-find').onclick = () => openFind(false);
    q('#w-replace').onclick = () => openFind(true);
    const selectAll = () => {
      page.focus();
      const r = document.createRange();
      r.selectNodeContents(page);
      const s = window.getSelection();
      s.removeAllRanges();
      s.addRange(r);
      remember();
      refresh();
    };
    q('#w-select').onclick = (e) => O.menuAt(e.currentTarget, [{ label: 'Seleccionar todo', hint: 'Ctrl+E', action: selectAll }]);

    // ---------- Insertar ----------
    q('#w-table').onclick = (e) => {
      const btn = e.currentTarget;
      const box = document.createElement('div');
      box.className = 'table-pick';
      box.innerHTML = `<div class="tp-title">Tabla</div><div class="tp-grid">${Array.from({ length: 80 }, (_, i) => `<span data-c="${i % 10 + 1}" data-r="${Math.floor(i / 10) + 1}"></span>`).join('')}</div>`;
      document.body.append(box);
      const r = btn.getBoundingClientRect();
      box.style.left = r.left + 'px';
      box.style.top = r.bottom + 2 + 'px';
      const close = () => { box.remove(); document.removeEventListener('mousedown', out, true); };
      const out = (ev) => { if (!box.contains(ev.target)) close(); };
      box.addEventListener('mousedown', (ev) => ev.preventDefault());
      box.onmouseover = (ev) => {
        const c = ev.target.dataset.c, rr = ev.target.dataset.r;
        if (!c) return;
        box.querySelectorAll('span').forEach((s) => s.classList.toggle('on', +s.dataset.c <= +c && +s.dataset.r <= +rr));
        box.querySelector('.tp-title').textContent = `Tabla de ${c}x${rr}`;
      };
      box.onclick = (ev) => {
        const c = +ev.target.dataset.c, rr = +ev.target.dataset.r;
        if (!c) return;
        close();
        const row = '<tr>' + '<td><br></td>'.repeat(c) + '</tr>';
        insertHtml(`<table><tbody>${row.repeat(rr)}</tbody></table><p><br></p>`);
      };
      setTimeout(() => document.addEventListener('mousedown', out, true));
    };
    q('#w-break').onclick = () => insertHtml('<hr class="salto"><p><br></p>');
    q('#w-link').onclick = () => {
      const url = window.prompt('Dirección del vínculo (https://…)', 'https://');
      if (!url) return;
      if (!/^(https?:|mailto:)/i.test(url)) { O.notice('El vínculo debe empezar por https://, http:// o mailto:.'); return; }
      exec('createLink', url);
    };
    q('#w-image').onclick = () => q('#w-image-file').click();
    q('#w-image-file').onchange = (e) => {
      const file = e.target.files[0];
      e.target.value = '';
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        const img = new Image();
        img.onload = () => {
          const w = Math.min(img.naturalWidth, 600);
          const h = Math.round(img.naturalHeight * (w / img.naturalWidth));
          insertHtml(`<img src="${reader.result}" width="${w}" height="${h}" alt="">`);
        };
        img.onerror = () => O.notice('No se pudo leer esa imagen.');
        img.src = reader.result;
      };
      reader.readAsDataURL(file);
    };

    // ---------- Zoom ----------
    const zoomer = O.zoomInit(q('#w-zoom'), (z) => {
      zoom = z;
      page.style.zoom = z / 100;
      updateCount();
    });
    q('#w-zoom100').onclick = () => zoomer.set(100);
    // Vista movil (como Word para movil): el texto se ajusta al ancho de la pantalla a su tamano real, sin hojas.
    // Diseno de impresion: las hojas como se imprimen. En el movil se empieza en vista movil.
    const setView = (mobile) => {
      page.classList.toggle('mobile-view', mobile);
      q('#w-mobileview').classList.toggle('on', mobile);
      q('#w-printview').classList.toggle('on', !mobile);
      if (mobile) zoomer.set(100);
      else if (phone.matches) zoomer.set(Math.min(100, (scroller.clientWidth - 8) / px(pg.w) * 100));
      repaginate();
    };
    q('#w-mobileview').onclick = () => setView(true);
    q('#w-printview').onclick = () => setView(false);
    const phone = window.matchMedia('(max-width: 720px)');
    let fitted = false;
    const fitPhone = () => {
      if (fitted || !scroller.clientWidth) return;
      fitted = true;
      setView(phone.matches);
    };
    q('#w-pagewidth').onclick = () => zoomer.set((scroller.clientWidth - 40) / page.offsetWidth * 100);
    q('#w-onepage').onclick = () => zoomer.set(Math.min((scroller.clientHeight - 40) / px(pg.h), (scroller.clientWidth - 40) / page.offsetWidth) * 100);

    // ---------- Teclado: atajos de Word en espanol ----------
    page.addEventListener('keydown', (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey) {
        if (e.ctrlKey && e.altKey && e.key.toLowerCase() === 'k') { e.preventDefault(); q('#w-link').click(); }
        return;
      }
      const k = e.key.toLowerCase();
      const map = { n: 'bold', k: 'italic', s: 'underline', q: 'justifyLeft', t: 'justifyCenter', d: 'justifyRight', j: 'justifyFull' };
      if (map[k] && !e.shiftKey) { e.preventDefault(); remember(); exec(map[k]); }
      else if (k === 'e') { e.preventDefault(); selectAll(); }
      else if (k === 'b') { e.preventDefault(); openFind(false); }
      else if (k === 'l') { e.preventDefault(); openFind(true); }
      else if (k === 'enter') { e.preventDefault(); remember(); q('#w-break').click(); }
      else if (e.key === '>' || (e.shiftKey && e.code === 'IntlBackslash')) { e.preventDefault(); remember(); step(true); }
      else if (e.key === '<') { e.preventDefault(); remember(); step(false); }
      else if (e.key === '=' && !e.shiftKey) { e.preventDefault(); remember(); exec('subscript'); }
      else if (e.key === '+' && e.shiftKey) { e.preventDefault(); remember(); exec('superscript'); }
      else if (e.key === '*') { e.preventDefault(); toggleMarks(); }
    });

    page.addEventListener('input', () => { remember(); changed(); });
    page.addEventListener('keyup', () => { remember(); refresh(); });
    page.addEventListener('mouseup', () => { remember(); if (painter) applyPainter(); refresh(); });
    page.addEventListener('paste', (e) => {
      const html = e.clipboardData.getData('text/html');
      if (!html) return;
      e.preventDefault();
      document.execCommand('insertHTML', false, sanitize(html));
    });
    page.addEventListener('click', (e) => {
      const a = e.target.closest('a');
      if (a && e.ctrlKey) window.open(a.href, '_blank', 'noopener');
    });
    for (const el of page.querySelectorAll('[style*="font-family"]')) {
      const f = el.style.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (f) O.addFontOption(q('#w-font'), f);
    }
    O.addFontOption(q('#w-font'), base.font);
    q('#w-font').value = base.font;
    q('#w-size').value = base.size;
    root.querySelector('[data-style="P"]').classList.add('on');
    // ---------- Disposicion: margenes, orientacion y tamano ----------
    const setPage = (patch) => {
      Object.assign(pg, patch);
      applyPage();
      changed();
    };
    const near = (a, b) => Math.abs(a - b) < 20;
    q('#w-margins').onclick = (e) => O.menuAt(e.currentTarget, MARGINS.map(([name, hint, tb, lr]) => ({
      label: (near(pg.top, tb) && near(pg.left, lr) && near(pg.bottom, tb) && near(pg.right, lr) ? '✓ ' : ' ') + name, hint,
      action: () => setPage({ top: tb, bottom: tb, left: lr, right: lr }),
    })));
    q('#w-orient').onclick = (e) => {
      const land = pg.w > pg.h;
      const turn = (want) => { if (want !== land) setPage({ w: pg.h, h: pg.w }); };
      O.menuAt(e.currentTarget, [
        { label: (land ? ' ' : '✓ ') + 'Vertical', action: () => turn(false) },
        { label: (land ? '✓ ' : ' ') + 'Horizontal', action: () => turn(true) },
      ]);
    };
    q('#w-papersize').onclick = (e) => {
      const land = pg.w > pg.h;
      const short = Math.min(pg.w, pg.h), long = Math.max(pg.w, pg.h);
      O.menuAt(e.currentTarget, PAPERS.map(([name, hint, w, h]) => ({
        label: (near(short, w) && near(long, h) ? '✓ ' : ' ') + name, hint,
        action: () => setPage(land ? { w: h, h: w } : { w, h }),
      })));
    };
    q('#w-break2').onclick = () => q('#w-break').click();

    // Repaginar cuando cambian las letras o se carga una imagen; al imprimir manda el navegador.
    document.fonts.ready.then(repaginate);
    page.addEventListener('load', repaginate, true);
    const printStyle = document.createElement('style');
    const beforePrint = () => {
      if (!page.isConnected || !page.offsetParent) return;
      clearPush();
      const i = (t) => t / 1440 + "in";
      printStyle.textContent = `@page { size: ${i(pg.w)} ${i(pg.h)}; margin: ${i(pg.top)} ${i(pg.right)} ${i(pg.bottom)} ${i(pg.left)}; }`;
      document.head.append(printStyle);
    };
    const afterPrint = () => { printStyle.remove(); repaginate(); };
    window.addEventListener('beforeprint', beforePrint);
    window.addEventListener('afterprint', afterPrint);
    paginate();
    updateCount();
    requestAnimationFrame(fitPhone);

    return {
      getState: () => ({ html: cleanHtml(), base, heads, page: { ...pg } }),
      onShow: () => { fitPhone(); repaginate(); },
      undo: () => exec('undo'),
      redo: () => exec('redo'),
    };
  }

  O.apps.writer = { mount };
  O.sanitizeHtml = sanitize;
})();
