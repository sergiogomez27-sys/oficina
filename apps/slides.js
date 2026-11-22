// Presentaciones: lienzo con elementos en puntos, miniaturas y modo de presentacion.
(function () {
  const O = window.Oficina;
  const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
  // Texto de los marcadores vacios, como en PowerPoint.
  const HINT = { ctrTitle: 'Haga clic para agregar título', title: 'Haga clic para agregar título', subTitle: 'Haga clic para agregar subtítulo', body: 'Haga clic para agregar texto' };
  const isTitle = (it) => it.ph === 'title' || it.ph === 'ctrTitle';
  const fontsOf = (state) => ({ major: 'Aptos Display', minor: 'Aptos', ...(state.fonts || {}) });
  // Disenos de diapositiva nueva de PowerPoint (16:9, 960 x 540 pt), escalados al tamano de la presentacion.
  const LAYOUTS = {
    title: { name: 'Diapositiva de título', els: [
      { ph: 'ctrTitle', x: 120, y: 88.4, w: 720, h: 188, size: 60, anchor: 'b', align: 'center' },
      { ph: 'subTitle', x: 120, y: 283.6, w: 720, h: 130.4, size: 24, anchor: 't', align: 'center' },
    ] },
    content: { name: 'Título y objetos', els: [
      { ph: 'title', x: 66, y: 28.75, w: 828, h: 104.4, size: 44, anchor: 'ctr', align: 'left' },
      { ph: 'body', x: 66, y: 143.75, w: 828, h: 342.6, size: 28, anchor: 't', align: 'left', bullets: true },
    ] },
    titleOnly: { name: 'Solo el título', els: [
      { ph: 'title', x: 66, y: 28.75, w: 828, h: 104.4, size: 44, anchor: 'ctr', align: 'left' },
    ] },
    blank: { name: 'En blanco', els: [] },
  };

  // Transiciones de PowerPoint que Oficina reproduce, con su duracion por defecto en segundos.
  const TRANSITIONS = [['none', 'Ninguna', 0], ['fade', 'Desvanecer', 0.7], ['push', 'Empujar', 1], ['wipe', 'Barrido', 1], ['split', 'Dividir', 1.25], ['cover', 'Cubrir', 1]];
  // dir es hacia donde se mueve la diapositiva nueva, como en PowerPoint (l = entra por la derecha).
  const MOVE = { l: [100, 0], r: [-100, 0], u: [0, 100], d: [0, -100] };
  // Hace entrar la diapositiva nueva (el) sobre la anterior (old, puede ser null).
  function playTransition(el, old, t) {
    if (!t || t.type === 'none' || !el.animate) return null;
    const opt = { duration: (t.dur || 1) * 1000, easing: 'ease-in-out' };
    const [x, y] = MOVE[t.dir] || MOVE.l;
    const from = `translate(${x}%, ${y}%)`;
    if (t.type === 'fade') return el.animate([{ opacity: 0 }, { opacity: 1 }], opt);
    if (t.type === 'cover') return el.animate([{ transform: from }, { transform: 'none' }], opt);
    if (t.type === 'push') {
      if (old) old.animate([{ transform: 'none' }, { transform: `translate(${-x}%, ${-y}%)` }], { ...opt, fill: 'forwards' });
      return el.animate([{ transform: from }, { transform: 'none' }], opt);
    }
    if (t.type === 'wipe') {
      const start = { l: 'inset(0 0 0 100%)', r: 'inset(0 100% 0 0)', u: 'inset(100% 0 0 0)', d: 'inset(0 0 100% 0)' }[t.dir] || 'inset(0 0 0 100%)';
      return el.animate([{ clipPath: start }, { clipPath: 'inset(0 0 0 0)' }], opt);
    }
    if (t.type === 'split') {
      const start = t.orient === 'vert' ? 'inset(0 50% 0 50%)' : 'inset(50% 0 50% 0)';
      return el.animate([{ clipPath: start }, { clipPath: 'inset(0 0 0 0)' }], opt);
    }
    return null;
  }

  function buildSlide(slide, state, opts = {}) {
    opts = { fonts: fontsOf(state), ...opts };
    const el = document.createElement('div');
    el.className = 'slide';
    el.style.width = state.w + 'pt';
    el.style.height = state.h + 'pt';
    el.style.background = '#' + (slide.bg || 'FFFFFF');
    for (const item of slide.els) el.append(buildEl(item, opts));
    return el;
  }

  function buildEl(item, opts = {}) {
    const n = document.createElement('div');
    n.dataset.id = item.id;
    n.style.left = item.x + 'pt';
    n.style.top = item.y + 'pt';
    n.style.width = item.w + 'pt';
    n.style.height = item.h + 'pt';
    if (item.type === 'image') {
      n.className = 'el';
      const img = document.createElement('img');
      img.src = item.src;
      img.alt = '';
      img.draggable = false;
      n.append(img);
      return n;
    }
    n.className = `el anchor-${item.anchor || 't'} geom-${item.geom || 'rect'}`;
    if (item.fill) n.style.background = '#' + item.fill;
    if (item.line) n.style.border = `1pt solid #${item.line}`;
    const txt = document.createElement('div');
    txt.className = 'txt';
    txt.style.fontSize = (item.size || 18) + 'pt';
    txt.style.color = '#' + (item.color || '000000');
    txt.style.textAlign = item.align || 'left';
    txt.style.fontWeight = item.bold ? '700' : '400';
    txt.style.fontStyle = item.italic ? 'italic' : 'normal';
    const font = item.font || (opts.fonts && (isTitle(item) ? opts.fonts.major : opts.fonts.minor));
    if (font) txt.style.fontFamily = `'${font}'`;
    txt.innerHTML = item.html || '<div><br></div>';
    n.append(txt);
    if (!txt.textContent.trim() && !item.fill) {
      n.classList.add('empty');
      txt.dataset.hint = HINT[item.ph] || 'Haga clic para agregar texto';
    }
    return n;
  }

  function mount(root, state, onChange) {
    const rb = O.rb;
    const fonts = fontsOf(state);
    root.innerHTML = `
      ${O.ribbon([
        { name: 'Inicio', groups: [
          { label: 'Portapapeles', html: `${rb('p-paste', 'paste', 'Pegar', 'Pegar (Ctrl+V)', { big: true })}<div class="rcol">${rb('p-cut', 'cut', 'Cortar', 'Cortar (Ctrl+X)', { text: 'Cortar' })}${rb('p-copy', 'copy', 'Copiar', 'Copiar (Ctrl+C)', { text: 'Copiar' })}</div>` },
          { label: 'Diapositivas', html: `<span class="rb-bigsplit">${rb('p-new', 'newslide', 'Nueva diapositiva', 'Nueva diapositiva (Ctrl+M)', { big: true })}<button class="rb rb-drop" id="p-new-menu" aria-label="Elegir el diseño de la nueva diapositiva">▾</button></span><div class="rcol">${rb('p-layout', 'layout', 'Diseño', 'Cambiar el diseño de la diapositiva', { text: 'Diseño', menu: true })}${rb('p-dup', 'duplicate', 'Duplicar', 'Duplicar diapositiva', { text: 'Duplicar' })}${rb('p-del', 'remove', 'Eliminar', 'Eliminar diapositiva', { text: 'Eliminar' })}</div>` },
          { label: 'Fuente', html: `<div class="rcol">
            <div class="rrow"><select id="p-font" class="r-font" aria-label="Fuente">${O.installedFonts().map(O.fontOption).join('')}</select><select id="p-size" class="r-size" aria-label="Tamaño de fuente">${SIZES.map((x) => `<option value="${x}">${String(x).replace('.', ',')}</option>`).join('')}</select>${rb('p-grow', 'grow', 'Aumentar tamaño de fuente', 'Aumentar tamaño de fuente (Ctrl+>)')}${rb('p-shrink', 'shrink', 'Disminuir tamaño de fuente', 'Disminuir tamaño de fuente (Ctrl+<)')}</div>
            <div class="rrow"><button class="rb" id="p-b" title="Negrita (Ctrl+N)" aria-label="Negrita"><b class="glyph">N</b></button><button class="rb" id="p-i" title="Cursiva (Ctrl+K)" aria-label="Cursiva"><i class="glyph serif">K</i></button><button class="rb" id="p-u" title="Subrayado (Ctrl+S)" aria-label="Subrayado"><u class="glyph">S</u></button><button class="rb" id="p-s" title="Tachado" aria-label="Tachado"><s class="glyph">abc</s></button><span class="rsep"></span>${O.rbColor('p-color', 'A', 'Color de fuente', '#C00000')}</div></div>` },
          { label: 'Párrafo', html: `<div class="rcol">
            <div class="rrow">${rb('p-bullet', 'bullets', 'Viñetas')}${rb('p-anchor', 'vtop', 'Alinear texto', 'Alinear texto arriba, en el medio o abajo', { menu: true })}</div>
            <div class="rrow">${rb('', 'left', 'Alinear a la izquierda', 'Alinear a la izquierda (Ctrl+Q)', { attrs: 'data-al="left"' })}${rb('', 'center', 'Centrar', 'Centrar (Ctrl+T)', { attrs: 'data-al="center"' })}${rb('', 'right', 'Alinear a la derecha', 'Alinear a la derecha (Ctrl+D)', { attrs: 'data-al="right"' })}${rb('', 'justify', 'Justificar', 'Justificar (Ctrl+J)', { attrs: 'data-al="justify"' })}</div></div>` },
          { label: 'Dibujo', html: `<div class="rshapes">${rb('p-text', 'textbox', 'Cuadro de texto')}${rb('', 'rect', 'Rectángulo', '', { attrs: 'data-add="rect"' })}${rb('', 'roundRect', 'Rectángulo redondeado', '', { attrs: 'data-add="roundRect"' })}${rb('', 'ellipse', 'Elipse', '', { attrs: 'data-add="ellipse"' })}</div>
            <div class="rcol">${rb('p-front', 'front', 'Traer al frente', 'Traer al frente', { text: 'Traer al frente' })}${rb('p-back', 'back', 'Enviar al fondo', 'Enviar al fondo', { text: 'Enviar al fondo' })}</div>
            <div class="rcol"><span class="rb-split wide" title="Relleno de forma"><button class="rb rb-color" id="p-fill" aria-label="Relleno de forma">${O.icon('fill')}<span>Relleno de forma</span><i class="bar" style="background:#156082"></i></button><button class="rb rb-drop" id="p-fill-menu" aria-label="Relleno de forma: colores">▾</button></span><span class="rb-split wide" title="Contorno de forma"><button class="rb rb-color" id="p-line" aria-label="Contorno de forma">${O.icon('outline')}<span>Contorno de forma</span><i class="bar" style="background:#042433"></i></button><button class="rb rb-drop" id="p-line-menu" aria-label="Contorno de forma: colores">▾</button></span></div>` },
        ] },
        { name: 'Insertar', groups: [
          { label: 'Diapositivas', html: rb('p-new2', 'newslide', 'Nueva diapositiva', 'Nueva diapositiva (Ctrl+M)', { big: true }) },
          { label: 'Imágenes', html: `${rb('p-image', 'image', 'Imágenes', 'Insertar imágenes de este equipo', { big: true })}<input type="file" id="p-image-file" accept="image/png,image/jpeg,image/gif" hidden>` },
          { label: 'Ilustraciones', html: `<div class="rshapes">${rb('', 'rect', 'Rectángulo', '', { attrs: 'data-add="rect"' })}${rb('', 'roundRect', 'Rectángulo redondeado', '', { attrs: 'data-add="roundRect"' })}${rb('', 'ellipse', 'Elipse', '', { attrs: 'data-add="ellipse"' })}</div>` },
          { label: 'Texto', html: rb('p-text2', 'textbox', 'Cuadro de texto', 'Dibujar un cuadro de texto', { big: true }) },
        ] },
        { name: 'Diseño', groups: [
          { label: 'Personalizar', html: `<span class="rb-bigsplit">${rb('p-bg', 'background', 'Formato del fondo', 'Color de fondo de la diapositiva', { big: true })}<button class="rb rb-drop" id="p-bg-menu" aria-label="Formato del fondo: colores">▾</button></span>` },
        ] },
        { name: 'Transiciones', groups: [
          { label: 'Vista previa', html: rb('p-trans-preview', 'playcur', 'Vista previa', 'Ver la transición de esta diapositiva', { big: true }) },
          { label: 'Transición a esta diapositiva', html: `<div class="rtrans">${TRANSITIONS.map(([k, label]) => rb('', k, label, label, { big: true, attrs: `data-trans="${k}"` })).join('')}</div>` },
          { label: 'Intervalos', html: `<div class="rcol"><label class="rfield">Duración: <input id="p-dur" type="number" min="0.1" max="60" step="0.25" value="1" aria-label="Duración en segundos"> s</label>${rb('p-trans-all', 'duplicate', 'Aplicar a todo', 'Usar esta transición en todas las diapositivas', { text: 'Aplicar a todo' })}</div>` },
        ] },
        { name: 'Presentación con diapositivas', groups: [
          { label: 'Iniciar presentación con diapositivas', html: `${rb('p-show-start', 'play', 'Desde el principio', 'Desde el principio (F5)', { big: true })}${rb('p-show', 'playcur', 'Desde la diapositiva actual', 'Desde la diapositiva actual (Mayús+F5)', { big: true })}` },
        ] },
      ])}
      <div class="slides-body">
        <div class="thumbs" aria-label="Diapositivas"></div>
        <div class="stage-wrap"><div class="stage-size stage"></div></div>
      </div>
      <div class="statusbar office">
        <span class="p-count">Diapositiva 1 de 1</span><span class="w-lang" title="Idioma de revisión">Español (España)</span>
        <span class="st-fill"></span>
        <button class="st-view on" id="p-view-normal" title="Normal">${O.icon('normal')}</button>
        <button class="st-view" id="p-view-show" title="Presentación con diapositivas">${O.icon('play')}</button>
        ${O.zoomControl('p-zoom')}
        <button class="st-view" id="p-fit" title="Ajustar la diapositiva a la ventana">${O.icon('fit')}</button>
      </div>`;
    O.ribbonInit(root);
    const q = (s) => root.querySelector(s);
    const stage = q('.stage');
    const wrap = q('.stage-wrap');
    let current = 0, selected = null, editing = null, scale = 1, zoom = null;
    let clip = null;
    const undo = [], redo = [];
    const PT = 96 / 72;
    state.nextId = state.nextId || 1;

    const slide = () => state.slides[current];
    const item = (id) => slide().els.find((e) => e.id === id);
    const snap = () => JSON.stringify({ slides: state.slides, nextId: state.nextId, current });
    const before = () => {
      undo.push(snap());
      if (undo.length > 60) undo.shift();
      redo.length = 0;
    };
    const restore = (json) => {
      const p = JSON.parse(json);
      state.slides = p.slides;
      state.nextId = p.nextId;
      current = Math.min(p.current, state.slides.length - 1);
      selected = null;
      editing = null;
      render();
      onChange();
    };
    const changed = () => { renderThumbs(); onChange(); };
    // Pestana Transiciones: marca la de la diapositiva actual.
    const markTrans = () => {
      const t = slide().transition;
      const type = t ? t.type : 'none';
      root.querySelectorAll('[data-trans]').forEach((b) => b.classList.toggle('on', b.dataset.trans === type));
      const dur = root.querySelector('#p-dur');
      if (dur && document.activeElement !== dur) { dur.value = t ? t.dur : 1; dur.disabled = !t; }
    };

    function fit() {
      const availW = wrap.clientWidth - 32, availH = wrap.clientHeight - 32;
      if (availW <= 0 || availH <= 0) return;
      scale = zoom ? zoom / 100 : Math.min(availW / (state.w * PT), availH / (state.h * PT));
      wrap.classList.toggle('zoomed', !!zoom);
      if (!zoom && zoomer) zoomer.set(scale * 100, false);
      stage.style.width = state.w * PT * scale + 'px';
      stage.style.height = state.h * PT * scale + 'px';
      const s = stage.querySelector('.slide');
      if (s) s.style.transform = `scale(${scale})`;
    }

    function renderStage() {
      const s = buildSlide(slide(), state);
      stage.replaceChildren(s);
      fit();
      markSelection();
    }
    function renderThumbs() {
      markTrans();
      const box = q('.thumbs');
      box.style.setProperty('--ratio', `${state.w} / ${state.h}`);
      box.replaceChildren();
      state.slides.forEach((sl, i) => {
        const b = document.createElement('button');
        b.className = 'thumb' + (i === current ? ' active' : '');
        b.setAttribute('aria-label', `Diapositiva ${i + 1}`);
        const num = document.createElement('span');
        num.textContent = i + 1;
        if (sl.transition && sl.transition.type !== 'none') {
          const star = document.createElement('span');
          star.className = 'thumb-trans';
          star.title = 'Tiene transición';
          star.textContent = '★';
          num.append(star);
        }
        const frame = document.createElement('div');
        frame.className = 'thumb-frame';
        const mini = buildSlide(sl, state);
        frame.append(mini);
        b.append(num, frame);
        b.onclick = () => { stopEditing(); current = i; selected = null; render(); };
        b.oncontextmenu = (e) => {
          e.preventDefault();
          b.onclick();
          O.contextMenu(e.clientX, e.clientY, slideItems());
        };
        box.append(b);
        requestAnimationFrame(() => { mini.style.transform = `scale(${frame.clientWidth / (state.w * PT)})`; });
      });
    }
    function render() { renderStage(); renderThumbs(); syncBar(); }
    const count = () => { q('.p-count').textContent = `Diapositiva ${current + 1} de ${state.slides.length}`; };

    function markSelection() {
      stage.querySelectorAll('.selected').forEach((n) => n.classList.remove('selected'));
      stage.querySelectorAll('.handle').forEach((n) => n.remove());
      if (selected === null) return;
      const n = stage.querySelector(`.el[data-id="${selected}"]`);
      if (!n) return;
      n.classList.add('selected');
      if (editing === null) {
        for (const d of ['nw', 'ne', 'sw', 'se']) {
          const h = document.createElement('span');
          h.className = 'handle ' + d;
          h.dataset.handle = d;
          // Compensa la escala del lienzo para que el tirador conserve su tamano.
          h.style.transform = `scale(${1 / scale})`;
          n.append(h);
        }
      }
    }
    function syncBar() {
      const it = selected !== null ? item(selected) : null;
      const isText = it && it.type !== 'image';
      q('#p-b').classList.toggle('on', !!(isText && it.bold));
      q('#p-i').classList.toggle('on', !!(isText && it.italic));
      // Como en PowerPoint, sin nada seleccionado la fuente y el tamano quedan en blanco y desactivados.
      for (const id of ['#p-font', '#p-size']) q(id).disabled = !isText;
      if (!isText) { q('#p-font').selectedIndex = -1; q('#p-size').selectedIndex = -1; }
      if (isText) {
        const sizeSel = q('#p-size');
        if (![...sizeSel.options].some((o) => +o.value === it.size)) {
          const o = document.createElement('option');
          o.value = it.size;
          o.textContent = String(it.size).replace('.', ',');
          sizeSel.insertBefore(o, [...sizeSel.options].find((x) => +x.value > it.size) || null);
        }
        sizeSel.value = it.size;
        const f = it.font || (isTitle(it) ? fonts.major : fonts.minor);
        O.addFontOption(q('#p-font'), f);
        q('#p-font').value = f;
      }
      root.querySelectorAll('[data-al]').forEach((b) => b.classList.toggle('on', !!(isText && (it.align || 'left') === b.dataset.al)));
      q('#p-del').disabled = state.slides.length < 2;
      count();
    }

    // ---------- Edicion de texto ----------
    function startEditing(id, cx, cy) {
      const it = item(id);
      if (!it || it.type === 'image') return;
      editing = id;
      selected = id;
      markSelection();
      const n = stage.querySelector(`.el[data-id="${id}"]`);
      n.classList.add('editing');
      n.classList.remove('empty');
      const txt = n.querySelector('.txt');
      txt.contentEditable = 'true';
      txt.focus();
      let r = cx !== undefined && document.caretRangeFromPoint ? document.caretRangeFromPoint(cx, cy) : null;
      if (!r || !txt.contains(r.startContainer)) {
        r = document.createRange();
        r.selectNodeContents(txt);
        r.collapse(false);
      }
      const s = window.getSelection();
      s.removeAllRanges();
      s.addRange(r);
    }
    function stopEditing() {
      if (editing === null) return;
      const it = item(editing);
      const n = stage.querySelector(`.el[data-id="${editing}"]`);
      const id = editing;
      editing = null;
      if (it && n) {
        const txt = n.querySelector('.txt');
        const html = txt.innerHTML;
        if (html !== it.html) { before(); it.html = html; }
        txt.contentEditable = 'false';
        n.classList.remove('editing');
        n.classList.toggle('empty', !txt.textContent.trim() && !it.fill);
        changed();
      }
      selected = id;
      markSelection();
    }

    // ---------- Raton sobre el lienzo ----------
    stage.addEventListener('mousedown', (e) => {
      const handle = e.target.dataset && e.target.dataset.handle;
      const n = e.target.closest('.el');
      if (n && editing === +n.dataset.id) return;
      if (editing !== null) stopEditing();
      if (!n) { selected = null; markSelection(); syncBar(); return; }
      e.preventDefault();
      const id = +n.dataset.id;
      const wasSelected = selected === id;
      selected = id;
      markSelection();
      syncBar();
      if (e.button !== 0) return;
      const it = item(id);
      const start = { x: it.x, y: it.y, w: it.w, h: it.h, mx: e.clientX, my: e.clientY };
      const k = 1 / (scale * PT);
      let moved = false;
      const move = (m) => {
        const dx = (m.clientX - start.mx) * k, dy = (m.clientY - start.my) * k;
        if (!moved && Math.abs(dx) + Math.abs(dy) < 2) return;
        if (!moved) { before(); moved = true; }
        if (handle) {
          if (handle.includes('e')) it.w = Math.max(10, start.w + dx);
          if (handle.includes('s')) it.h = Math.max(10, start.h + dy);
          if (handle.includes('w')) { it.w = Math.max(10, start.w - dx); it.x = start.x + start.w - it.w; }
          if (handle.includes('n')) { it.h = Math.max(10, start.h - dy); it.y = start.y + start.h - it.h; }
        } else {
          it.x = Math.max(-it.w + 10, Math.min(state.w - 10, start.x + dx));
          it.y = Math.max(-it.h + 10, Math.min(state.h - 10, start.y + dy));
        }
        n.style.left = it.x + 'pt';
        n.style.top = it.y + 'pt';
        n.style.width = it.w + 'pt';
        n.style.height = it.h + 'pt';
      };
      const up = (m) => {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
        if (moved) { changed(); return; }
        if (!handle && it.type !== 'image' && (wasSelected || !it.fill)) startEditing(id, m.clientX, m.clientY);
      };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
    });
    stage.addEventListener('dblclick', (e) => {
      const n = e.target.closest('.el');
      if (n) startEditing(+n.dataset.id);
    });
    function slideItems() {
      return [
        { label: 'Nueva diapositiva', action: () => q('#p-new').click() },
        { label: 'Duplicar diapositiva', action: () => q('#p-dup').click() },
        { label: 'Eliminar diapositiva', disabled: state.slides.length < 2, action: () => q('#p-del').click() },
      ];
    }
    function deleteSelected() {
      before();
      slide().els = slide().els.filter((x) => x.id !== selected);
      selected = null;
      render();
      onChange();
    }
    stage.addEventListener('contextmenu', (e) => {
      const n = e.target.closest('.el');
      // Mientras se escribe se deja el menu del navegador, con su corrector.
      if (n && editing === +n.dataset.id) return;
      e.preventDefault();
      if (!n) { O.contextMenu(e.clientX, e.clientY, slideItems()); return; }
      const id = +n.dataset.id;
      const items = [];
      if (item(id).type !== 'image') items.push({ label: 'Editar texto', hint: 'F2', action: () => startEditing(id) }, '-');
      items.push(
        { label: 'Traer al frente', action: () => q('#p-front').click() },
        { label: 'Enviar al fondo', action: () => q('#p-back').click() },
        '-',
        { label: 'Eliminar', hint: 'Supr', action: deleteSelected },
      );
      O.contextMenu(e.clientX, e.clientY, items);
    });

    const onKey = (e) => {
      if (root.hidden || editing !== null || !root.isConnected) return;
      if (e.target.closest && e.target.closest('input,select,textarea,[contenteditable="true"]')) return;
      const mod = e.ctrlKey || e.metaKey;
      if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); doUndo(); return; }
      if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); doRedo(); return; }
      if (mod && e.key.toLowerCase() === 'm') { e.preventDefault(); newSlide('content'); return; }
      if (mod && e.key.toLowerCase() === 'c' && selected !== null) { e.preventDefault(); copyEl(false); return; }
      if (mod && e.key.toLowerCase() === 'x' && selected !== null) { e.preventDefault(); copyEl(true); return; }
      if (mod && e.key.toLowerCase() === 'v' && clip) { e.preventDefault(); pasteEl(); return; }
      if (e.key === 'F5') { e.preventDefault(); present(e.shiftKey ? current : 0); return; }
      if (e.key === 'PageDown' || (e.key === 'ArrowDown' && selected === null)) { e.preventDefault(); if (current < state.slides.length - 1) { current++; selected = null; render(); } return; }
      if (e.key === 'PageUp' || (e.key === 'ArrowUp' && selected === null)) { e.preventDefault(); if (current > 0) { current--; selected = null; render(); } return; }
      if (selected === null) return;
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        deleteSelected();
      } else if (e.key === 'F2' || e.key === 'Enter') { e.preventDefault(); startEditing(selected); }
      else if (e.key.startsWith('Arrow')) {
        e.preventDefault();
        const it = item(selected), d = e.shiftKey ? 10 : 2;
        before();
        if (e.key === 'ArrowLeft') it.x -= d;
        if (e.key === 'ArrowRight') it.x += d;
        if (e.key === 'ArrowUp') it.y -= d;
        if (e.key === 'ArrowDown') it.y += d;
        renderStage();
        changed();
      }
    };
    document.addEventListener('keydown', onKey);

    function doUndo() {
      stopEditing();
      if (!undo.length) return;
      redo.push(snap());
      restore(undo.pop());
    }
    function doRedo() {
      if (!redo.length) return;
      undo.push(snap());
      restore(redo.pop());
    }

    // ---------- Barra ----------

    const addEl = (props) => {
      stopEditing();
      before();
      const el = { id: state.nextId++, x: state.w / 2 - 150, y: state.h / 2 - 75, w: 300, h: 150, ...props };
      slide().els.push(el);
      selected = el.id;
      renderStage();
      changed();
      syncBar();
      return el;
    };
    const fromLayout = (key) => {
      const kx = state.w / 960, ky = state.h / 540;
      return LAYOUTS[key].els.map((l) => ({
        id: state.nextId++, type: 'shape', geom: 'rect', fill: null, line: null, color: '000000', bold: false, ph: l.ph,
        x: l.x * kx, y: l.y * ky, w: l.w * kx, h: l.h * ky, size: l.size, anchor: l.anchor, align: l.align,
        html: l.bullets ? '<div class="bu"><br></div>' : '<div><br></div>',
      }));
    };
    const newSlide = (key) => {
      stopEditing();
      before();
      state.slides.splice(current + 1, 0, { bg: 'FFFFFF', els: fromLayout(key) });
      current++;
      selected = null;
      render();
      onChange();
    };
    const layoutMenu = (fn) => Object.entries(LAYOUTS).map(([k, l]) => ({ label: l.name, action: () => fn(k) }));
    q('#p-new').onclick = () => newSlide('content');
    q('#p-new2').onclick = () => newSlide('content');
    q('#p-new-menu').onclick = (e) => O.menuAt(e.currentTarget, layoutMenu(newSlide));
    // Cambiar el diseno conserva el texto de los marcadores que tienen pareja en el nuevo.
    q('#p-layout').onclick = (e) => O.menuAt(e.currentTarget, layoutMenu((key) => {
      stopEditing();
      before();
      const old = slide().els;
      const fresh = fromLayout(key);
      const family = (ph) => (ph === 'ctrTitle' || ph === 'title' ? 'title' : ph === 'subTitle' || ph === 'body' ? 'body' : null);
      for (const f of fresh) {
        const match = old.find((o) => o.ph && family(o.ph) === family(f.ph));
        if (match && match.html && match.html.replace(/<[^>]+>/g, '').trim()) f.html = match.html;
      }
      slide().els = [...fresh, ...old.filter((o) => !o.ph)];
      selected = null;
      render();
      onChange();
    }));
    q('#p-dup').onclick = () => {
      stopEditing();
      before();
      const copy = JSON.parse(JSON.stringify(slide()));
      copy.els.forEach((e) => { e.id = state.nextId++; });
      state.slides.splice(current + 1, 0, copy);
      current++;
      selected = null;
      render();
      onChange();
    };
    q('#p-del').onclick = () => {
      if (state.slides.length < 2) return;
      stopEditing();
      before();
      state.slides.splice(current, 1);
      current = Math.min(current, state.slides.length - 1);
      selected = null;
      render();
      onChange();
    };
    const addText = () => {
      const el = addEl({ type: 'shape', geom: 'rect', fill: null, line: null, size: 18, color: '000000', align: 'left', anchor: 't', bold: false, html: '<div><br></div>', h: 40 });
      startEditing(el.id);
    };
    q('#p-text').onclick = addText;
    q('#p-text2').onclick = addText;
    // Formas con el estilo por defecto de PowerPoint 2024: relleno azul del tema y contorno mas oscuro.
    root.querySelectorAll('[data-add]').forEach((b) => {
      b.onclick = () => addEl({ type: 'shape', geom: b.dataset.add, fill: '156082', line: '042433', size: 18, color: 'FFFFFF', align: 'center', anchor: 'ctr', bold: false, html: '<div><br></div>' });
    });
    q('#p-image').onclick = () => q('#p-image-file').click();
    q('#p-image-file').onchange = (e) => {
      const file = e.target.files[0];
      e.target.value = '';
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        const img = new Image();
        img.onload = () => {
          const k = Math.min(1, (state.w * 0.6) / img.naturalWidth, (state.h * 0.6) / img.naturalHeight);
          const w = img.naturalWidth * k, h = img.naturalHeight * k;
          addEl({ type: 'image', src: reader.result, w, h, x: (state.w - w) / 2, y: (state.h - h) / 2 });
        };
        img.onerror = () => O.notice('No se pudo leer esa imagen.');
        img.src = reader.result;
      };
      reader.readAsDataURL(file);
    };
    // Dentro de un texto en edicion el formato se aplica a la seleccion; si no, a todo el cuadro.
    const setProp = (fn) => {
      const it = selected !== null ? item(selected) : null;
      if (!it || it.type === 'image') return;
      before();
      fn(it);
      const keep = selected;
      renderStage();
      selected = keep;
      markSelection();
      changed();
      syncBar();
    };
    const inEdit = (cmd, val) => {
      if (editing === null) return false;
      document.execCommand('styleWithCSS', false, true);
      document.execCommand(cmd, false, val);
      return true;
    };
    q('#p-b').onclick = () => { if (!inEdit('bold')) setProp((i) => { i.bold = !i.bold; }); };
    q('#p-i').onclick = () => { if (!inEdit('italic')) setProp((i) => { i.italic = !i.italic; }); };
    // Fuera de la edicion, el formato que el modelo no guarda por cuadro se aplica a todo su texto.
    const wholeBox = (cmd, val) => {
      if (inEdit(cmd, val)) return;
      const id = selected;
      const it = id !== null ? item(id) : null;
      if (!it || it.type === 'image') return;
      startEditing(id);
      const txt = stage.querySelector(`.el[data-id="${id}"] .txt`);
      const r = document.createRange();
      r.selectNodeContents(txt);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
      document.execCommand('styleWithCSS', false, true);
      document.execCommand(cmd, false, val);
      stopEditing();
    };
    q('#p-u').onclick = () => wholeBox('underline');
    q('#p-s').onclick = () => wholeBox('strikeThrough');
    const sizeOf = () => { const it = selected !== null ? item(selected) : null; return it && it.type !== 'image' ? it.size || 18 : null; };
    const setSize = (v) => {
      if (editing !== null) stopEditing();
      // El tamano del cuadro manda: se quitan los tamanos puestos dentro del texto.
      setProp((i) => { i.size = v; i.html = (i.html || '').replace(/font-size:\s*[^;"]+;?/g, ''); });
    };
    q('#p-size').onchange = (e) => setSize(+e.target.value);
    q('#p-grow').onclick = () => { const v = sizeOf(); if (v) setSize(SIZES.find((x) => x > v) || v + 10); };
    q('#p-shrink').onclick = () => { const v = sizeOf(); if (v) setSize([...SIZES].reverse().find((x) => x < v) || Math.max(1, v - 1)); };
    q('#p-font').onchange = (e) => {
      if (inEdit('fontName', e.target.value)) return;
      setProp((i) => { i.font = e.target.value; i.html = (i.html || '').replace(/font-family:\s*[^;"]+(;|(?="))/g, ''); });
    };
    q('#p-anchor').onclick = (e) => O.menuAt(e.currentTarget, [['t', 'Superior'], ['ctr', 'En el medio'], ['b', 'Inferior']].map(([v, label]) => ({
      label, action: () => { if (editing !== null) stopEditing(); setProp((i) => { i.anchor = v; }); },
    })));
    root.querySelectorAll('[data-al]').forEach((b) => {
      b.onclick = () => {
        if (editing !== null) stopEditing();
        // La alineacion de cada parrafo manda sobre la del cuadro, asi que se limpia.
        setProp((i) => {
          i.align = b.dataset.al;
          i.html = (i.html || '').replace(/text-align:\s*\w+;?/g, '');
        });
      };
    });
    q('#p-bullet').onclick = () => {
      if (editing !== null) stopEditing();
      setProp((i) => {
        const doc = new DOMParser().parseFromString(`<body>${i.html || ''}`, 'text/html').body;
        const paras = [...doc.children];
        const on = paras.length && paras.every((p) => p.classList.contains('bu'));
        paras.forEach((p) => p.classList.toggle('bu', !on));
        i.html = doc.innerHTML;
      });
    };
    let textColor = 'C00000', fillColor = '156082', lineColor = '042433', bgColor = 'FFFFFF';
    const applyText = (v) => {
      if (!inEdit('foreColor', '#' + v)) setProp((i) => { i.color = v; i.html = (i.html || '').replace(/(^|[;"\s])color:\s*[^;"]+;?/g, '$1'); });
    };
    q('#p-color').onclick = () => applyText(textColor);
    q('#p-color-menu').onclick = (e) => O.colorPalette(e.currentTarget, 'theme', (v) => {
      const c = v === 'auto' ? '000000' : v;
      if (v !== 'auto') { textColor = v; O.setColorBar(q('#p-color'), v); }
      applyText(c);
    }, { auto: 'Automático' });
    const applyFill = (v) => { if (editing !== null) stopEditing(); setProp((i) => { i.fill = v === 'none' ? null : v; }); };
    q('#p-fill').onclick = () => applyFill(fillColor);
    q('#p-fill-menu').onclick = (e) => O.colorPalette(e.currentTarget, 'theme', (v) => {
      if (v !== 'none') { fillColor = v; O.setColorBar(q('#p-fill'), v); }
      applyFill(v);
    }, { none: 'Sin relleno' });
    const applyLine = (v) => { if (editing !== null) stopEditing(); setProp((i) => { i.line = v === 'none' ? null : v; }); };
    q('#p-line').onclick = () => applyLine(lineColor);
    q('#p-line-menu').onclick = (e) => O.colorPalette(e.currentTarget, 'theme', (v) => {
      if (v !== 'none') { lineColor = v; O.setColorBar(q('#p-line'), v); }
      applyLine(v);
    }, { none: 'Sin contorno' });
    const applyBg = (v) => {
      before();
      slide().bg = v;
      renderStage();
      changed();
    };
    q('#p-bg').onclick = () => applyBg(bgColor);
    q('#p-bg-menu').onclick = (e) => O.colorPalette(e.currentTarget, 'theme', (v) => { bgColor = v; applyBg(v); });

    // Portapapeles propio para formas; el texto en edicion usa el del sistema.
    const copyEl = (cut) => {
      if (editing !== null) { document.execCommand(cut ? 'cut' : 'copy'); return; }
      if (selected === null) return;
      clip = JSON.parse(JSON.stringify(item(selected)));
      if (cut) deleteSelected();
    };
    const pasteEl = () => {
      if (!clip) { if (editing === null) O.notice('No hay ninguna forma copiada. Para pegar texto, haz clic en un cuadro y usa Ctrl+V.'); return; }
      before();
      const el = { ...JSON.parse(JSON.stringify(clip)), id: state.nextId++ };
      el.x += 10; el.y += 10;
      clip.x += 10; clip.y += 10;
      slide().els.push(el);
      selected = el.id;
      renderStage();
      changed();
      syncBar();
    };
    q('#p-copy').onclick = () => copyEl(false);
    q('#p-cut').onclick = () => copyEl(true);
    q('#p-paste').onclick = pasteEl;
    const reorder = (front) => {
      if (selected === null) return;
      stopEditing();
      before();
      const it = item(selected);
      slide().els = slide().els.filter((x) => x !== it);
      if (front) slide().els.push(it);
      else slide().els.unshift(it);
      renderStage();
      changed();
    };
    q('#p-front').onclick = () => reorder(true);
    q('#p-back').onclick = () => reorder(false);

    // ---------- Presentar ----------
    const present = (from) => {
      stopEditing();
      let i = from;
      const over = document.createElement('div');
      over.className = 'presenter';
      over.tabIndex = 0;
      const holder = document.createElement('div');
      holder.className = 'presenter-holder';
      holder.style.cssText = 'position:relative;overflow:hidden';
      const hint = document.createElement('div');
      hint.className = 'hint';
      over.append(holder, hint);
      // animate: reproducir la transicion de la diapositiva que entra.
      const draw = (animate) => {
        const s = buildSlide(state.slides[i], state);
        const k = Math.min(window.innerWidth / (state.w * PT), window.innerHeight / (state.h * PT));
        holder.style.width = state.w * PT * k + 'px';
        holder.style.height = state.h * PT * k + 'px';
        const box = document.createElement('div');
        box.style.cssText = 'position:absolute;inset:0;overflow:hidden';
        s.style.transform = `scale(${k})`;
        box.append(s);
        const old = [...holder.children];
        holder.append(box);
        hint.textContent = `${i + 1} / ${state.slides.length} · Esc para salir`;
        const anim = animate === true ? playTransition(box, old[old.length - 1] || null, state.slides[i].transition) : null;
        if (anim) anim.finished.then(() => old.forEach((o) => o.remove()), () => {});
        else old.forEach((o) => o.remove());
      };
      // Al cambiar el tamano de la ventana solo se reescala, sin cortar una transicion en marcha.
      const redraw = () => {
        const k = Math.min(window.innerWidth / (state.w * PT), window.innerHeight / (state.h * PT));
        holder.style.width = state.w * PT * k + 'px';
        holder.style.height = state.h * PT * k + 'px';
        for (const sl of holder.querySelectorAll('.slide')) sl.style.transform = `scale(${k})`;
      };
      const close = () => {
        document.removeEventListener('keydown', key, true);
        window.removeEventListener('resize', redraw);
        if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
        over.remove();
        current = i;
        selected = null;
        render();
      };
      const go = (d) => {
        if (i + d >= state.slides.length) { close(); return; }
        i = Math.max(0, i + d);
        draw(d > 0);
      };
      const key = (e) => {
        e.stopPropagation();
        if (e.key === 'Escape') { e.preventDefault(); close(); }
        else if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'].includes(e.key)) { e.preventDefault(); go(1); }
        else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(e.key)) { e.preventDefault(); go(-1); }
      };
      over.onclick = () => go(1);
      document.addEventListener('keydown', key, true);
      window.addEventListener('resize', redraw);
      document.body.append(over);
      draw(true);
      if (over.requestFullscreen) over.requestFullscreen().then(redraw).catch(() => {});
    };
    // ---------- Transiciones ----------
    const preview = () => {
      const s = stage.querySelector('.slide');
      if (s) playTransition(s, null, slide().transition);
    };
    const setTrans = (type, all) => {
      before();
      const def = TRANSITIONS.find((x) => x[0] === type);
      const cur = slide().transition;
      const t = type === 'none' ? null : { ...(cur && cur.type === type ? cur : {}), type, dur: cur && cur.type === type ? cur.dur : def[2] };
      for (const sl of all ? state.slides : [slide()]) {
        if (t) sl.transition = { ...t };
        else delete sl.transition;
      }
      changed();
      if (t && !all) preview();
    };
    root.querySelectorAll('[data-trans]').forEach((b) => { b.onclick = () => setTrans(b.dataset.trans, false); });
    q('#p-trans-preview').onclick = preview;
    q('#p-trans-all').onclick = () => setTrans(slide().transition ? slide().transition.type : 'none', true);
    q('#p-dur').onchange = (e) => {
      const v = parseFloat(String(e.target.value).replace(',', '.'));
      if (!slide().transition || !(v > 0)) { markTrans(); return; }
      before();
      slide().transition.dur = Math.min(60, Math.round(v * 100) / 100);
      changed();
    };

    q('#p-show').onclick = () => present(current);
    q('#p-show-start').onclick = () => present(0);
    q('#p-view-show').onclick = () => present(current);
    q('#p-view-normal').onclick = () => { zoom = null; fit(); markSelection(); };
    var zoomer = O.zoomInit(q('#p-zoom'), (z) => { zoom = z; fit(); markSelection(); });
    q('#p-fit').onclick = () => { zoom = null; fit(); markSelection(); };

    // Al imprimir: todas las diapositivas en A4 apaisado (1122 px de ancho a 96 ppp), una por pagina.
    let printBox = null, pageStyle = null;
    const beforePrint = () => {
      if (root.hidden) return;
      stopEditing();
      const k = Math.min(1122 / (state.w * PT), 793 / (state.h * PT));
      printBox = document.createElement('div');
      printBox.className = 'print-slides';
      for (const sl of state.slides) {
        const frame = document.createElement('div');
        frame.className = 'print-slide';
        const inner = document.createElement('div');
        inner.style.cssText = `position:relative;overflow:hidden;width:${state.w * PT * k}px;height:${state.h * PT * k}px`;
        const s = buildSlide(sl, state);
        s.style.transform = `scale(${k})`;
        inner.append(s);
        frame.append(inner);
        printBox.append(frame);
      }
      root.append(printBox);
      pageStyle = document.createElement('style');
      pageStyle.textContent = '@page { size: A4 landscape; margin: 0; }';
      document.head.append(pageStyle);
    };
    const afterPrint = () => {
      if (printBox) printBox.remove();
      if (pageStyle) pageStyle.remove();
      printBox = pageStyle = null;
    };
    window.addEventListener('beforeprint', beforePrint);
    window.addEventListener('afterprint', afterPrint);

    const ro = new ResizeObserver(() => { fit(); markSelection(); });
    ro.observe(wrap);
    render();

    return {
      getState: () => state,
      flush: stopEditing,
      undo: doUndo,
      redo: doRedo,
      onShow: () => { fit(); renderThumbs(); },
      destroy: () => {
        ro.disconnect();
        window.removeEventListener('beforeprint', beforePrint);
        window.removeEventListener('afterprint', afterPrint);
        document.removeEventListener('keydown', onKey);
      },
    };
  }

  O.apps.slides = { mount };
})();
