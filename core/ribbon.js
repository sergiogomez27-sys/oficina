// Cinta de opciones al estilo de Office: pestanas, grupos con su nombre, iconos propios, menus y paletas de color.
(function () {
  const O = window.Oficina;

  // Iconos dibujados para Oficina (16 x 16, trazo del color del texto).
  const P = {
    paste: '<rect x="3" y="3" width="10" height="12" rx="1"/><path d="M6 3V2h4v1M6 7h4M6 10h4"/>',
    cut: '<circle cx="4.5" cy="12" r="2"/><circle cx="11.5" cy="12" r="2"/><path d="M6 10.5 12 2M10 10.5 4 2"/>',
    copy: '<rect x="5" y="5" width="8" height="10" rx="1"/><path d="M3 11V2h7"/>',
    brush: '<path d="M3 2h9v4H3zM12 4h2v4H8v2M7 10h2v5H7z"/>',
    undo: '<path d="M5 3 2 6l3 3M2 6h7a4 4 0 0 1 0 8H6"/>',
    redo: '<path d="m11 3 3 3-3 3M14 6H7a4 4 0 0 0 0 8h3"/>',
    grow: '<path d="M2 14 6 3l4 11M3.5 10h5M11 6l2-3 2 3"/>',
    shrink: '<path d="M2 14 6 3l4 11M3.5 10h5M11 3l2 3 2-3"/>',
    clear: '<path d="M3 3h9M7.5 3 5 13M9 9l5 5M14 9l-5 5"/>',
    bullets: '<circle cx="2.5" cy="4" r="1" fill="currentColor"/><circle cx="2.5" cy="8" r="1" fill="currentColor"/><circle cx="2.5" cy="12" r="1" fill="currentColor"/><path d="M6 4h8M6 8h8M6 12h8"/>',
    numbers: '<path d="M2 2.5h1v3M1.8 9h1.8L2 11.5h2M6 4h8M6 8h8M6 12h8"/>',
    outdent: '<path d="M2 3h12M7 6.5h7M7 9.5h7M2 13h12M5 6 2 8l3 2"/>',
    indent: '<path d="M2 3h12M7 6.5h7M7 9.5h7M2 13h12M2 6l3 2-3 2"/>',
    marks: '<path d="M9 2v12M12 2v12M12 2H7.5a3 3 0 0 0 0 6H9"/>',
    left: '<path d="M2 3h12M2 6.5h8M2 10h12M2 13.5h8"/>',
    center: '<path d="M2 3h12M4 6.5h8M2 10h12M4 13.5h8"/>',
    right: '<path d="M2 3h12M6 6.5h8M2 10h12M6 13.5h8"/>',
    justify: '<path d="M2 3h12M2 6.5h12M2 10h12M2 13.5h12"/>',
    spacing: '<path d="M7 3h7M7 8h7M7 13h7M3 2v12M1.5 4 3 2l1.5 2M1.5 12 3 14l1.5-2"/>',
    search: '<circle cx="6.5" cy="6.5" r="4"/><path d="m9.5 9.5 4.5 4.5"/>',
    replace: '<path d="M2 5h8l-2-2M14 11H6l2 2"/>',
    select: '<path d="M4 2v11l3-3 2 5 2-1-2-5h4z"/>',
    table: '<rect x="2" y="3" width="12" height="10"/><path d="M2 6.5h12M2 10h12M6 3v10M10 3v10"/>',
    image: '<rect x="2" y="3" width="12" height="10" rx="1"/><circle cx="6" cy="6.5" r="1.2"/><path d="m3 12 4-4 3 3 2-2 2 2"/>',
    link: '<path d="M7 9a3 3 0 0 0 4 0l2-2a3 3 0 0 0-4-4l-1 1M9 7a3 3 0 0 0-4 0L3 9a3 3 0 0 0 4 4l1-1"/>',
    pagebreak: '<path d="M3 2v4h10V2M3 14v-4h10v4M1 8h2M5 8h2M9 8h2M13 8h2"/>',
    zoom: '<circle cx="6.5" cy="6.5" r="4"/><path d="m9.5 9.5 4.5 4.5M4.5 6.5h4M6.5 4.5v4"/>',
    newslide: '<rect x="1.5" y="3.5" width="11" height="8" rx=".5"/><path d="M12.5 11.5v3M11 13h3M4 6h6M4 8.5h4"/>',
    layout: '<rect x="2" y="3" width="12" height="10"/><path d="M2 6h12M8 6v7"/>',
    duplicate: '<rect x="2" y="2" width="9" height="7"/><rect x="5" y="7" width="9" height="7"/>',
    remove: '<rect x="2" y="4" width="12" height="9"/><path d="m6 6.5 4 4M10 6.5l-4 4"/>',
    textbox: '<rect x="2" y="2" width="12" height="12"/><path d="M5 5h6M8 5v7"/>',
    rect: '<rect x="2.5" y="4" width="11" height="8"/>',
    roundRect: '<rect x="2.5" y="4" width="11" height="8" rx="2.5"/>',
    ellipse: '<ellipse cx="8" cy="8" rx="5.5" ry="4.5"/>',
    front: '<rect x="6" y="6" width="8" height="8" fill="currentColor"/><path d="M2 10V2h8"/>',
    back: '<rect x="2" y="2" width="8" height="8" fill="currentColor"/><path d="M6 14h8V6"/>',
    fill: '<path d="m3 8 5-5 5 5-5 5zM13 10s1.5 1.7 1.5 2.6a1.5 1.5 0 0 1-3 0c0-.9 1.5-2.6 1.5-2.6"/>',
    outline: '<path d="M2 13 11 4l2 2-9 9H2zM9 6l2 2"/>',
    vtop: '<path d="M2 2h12M8 5v9M5 8l3-3 3 3"/>',
    play: '<path d="M4 2.5v11l9-5.5z" fill="currentColor"/>',
    playcur: '<rect x="2" y="3" width="12" height="10"/><path d="M6.5 5.5v5l4-2.5z" fill="currentColor"/>',
    background: '<rect x="2" y="3" width="12" height="10"/><path d="m2 13 5-5 3 3 4-4"/>',
    normal: '<rect x="2" y="3" width="12" height="10"/><path d="M5 3v10"/>',
    page: '<path d="M4 2h6l3 3v9H4zM10 2v3h3"/>',
    fit: '<path d="M2 6V2h4M10 2h4v4M14 10v4h-4M6 14H2v-4"/>',
    case: '<path d="M1.5 12 4.5 4l3 8M2.5 9.5h4M10 8.5a2 2 0 1 1 0 3.5M12.5 8v4"/>',
    margins: '<rect x="3" y="1.5" width="10" height="13"/><path d="M5.5 1.5v13M10.5 1.5v13M3 4h10M3 12h10" stroke-dasharray="1 1"/>',
    orient: '<rect x="2" y="5" width="8" height="10"/><rect x="6" y="1.5" width="8.5" height="6" fill="var(--rb-bg, #fff)"/>',
    papersize: '<path d="M3 2h7l3 3v9H3zM10 2v3h3M5.5 11.5l5-5M8 6.5h2.5V9"/>',
    transition: '<rect x="1.5" y="3" width="7" height="10"/><rect x="7.5" y="3" width="7" height="10" fill="currentColor" fill-opacity=".25"/><path d="M5 8h5M8.5 6.5 10 8l-1.5 1.5"/>',
    phone: '<rect x="4.5" y="1.5" width="7" height="13" rx="1.2"/><path d="M7 12.5h2M6.5 4.5h3M6.5 6.5h3M6.5 8.5h2"/>',
    none: '<rect x="2" y="3" width="12" height="10"/>',
    fade: '<rect x="2" y="3" width="12" height="10"/><rect x="5" y="5.5" width="6" height="5" fill="currentColor" fill-opacity=".3" stroke="none"/>',
    push: '<rect x="2" y="3" width="12" height="10"/><path d="M5 8h6M9 6l2 2-2 2"/>',
    wipe: '<rect x="2" y="3" width="12" height="10"/><rect x="2" y="3" width="6" height="10" fill="currentColor" fill-opacity=".3" stroke="none"/>',
    split: '<rect x="2" y="3" width="12" height="10"/><path d="M8 3v10M6 8H3.5M10 8h2.5"/>',
    cover: '<rect x="2" y="3" width="12" height="10"/><rect x="6" y="3" width="8" height="10" fill="currentColor" fill-opacity=".3"/>',
  };
  O.icon = (name, size = 16) => `<svg class="ico" width="${size}" height="${size}" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.15" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${P[name] || ''}</svg>`;

  // Fuentes que suelen venir con Windows y Office; las listas solo muestran las instaladas.
  const CANDIDATES = ['Agency FB', 'Aharoni', 'Algerian', 'Aptos', 'Aptos Display', 'Aptos Narrow', 'Aptos Serif', 'Arial', 'Arial Black', 'Arial Narrow',
    'Arial Rounded MT Bold', 'Bahnschrift', 'Baskerville Old Face', 'Bauhaus 93', 'Bell MT', 'Berlin Sans FB', 'Bernard MT Condensed',
    'Blackadder ITC', 'Bodoni MT', 'Book Antiqua', 'Bookman Old Style', 'Bradley Hand ITC', 'Britannic Bold', 'Broadway', 'Brush Script MT',
    'Calibri', 'Calibri Light', 'Californian FB', 'Calisto MT', 'Cambria', 'Cambria Math', 'Candara', 'Castellar', 'Centaur', 'Century',
    'Century Gothic', 'Century Schoolbook', 'Chiller', 'Colonna MT', 'Comic Sans MS', 'Consolas', 'Constantia', 'Cooper Black',
    'Copperplate Gothic Bold', 'Corbel', 'Courier New', 'Curlz MT', 'Dubai', 'Ebrima', 'Edwardian Script ITC', 'Elephant', 'Engravers MT',
    'Eras Medium ITC', 'Felix Titling', 'Footlight MT Light', 'Forte', 'Franklin Gothic Book', 'Franklin Gothic Medium', 'Freestyle Script',
    'French Script MT', 'Gabriola', 'Gadugi', 'Garamond', 'Georgia', 'Gigi', 'Gill Sans MT', 'Gloucester MT Extra Condensed',
    'Goudy Old Style', 'Goudy Stout', 'Grandview', 'Haettenschweiler', 'Harlow Solid Italic', 'Harrington', 'High Tower Text', 'Impact',
    'Imprint MT Shadow', 'Informal Roman', 'Ink Free', 'Javanese Text', 'Jokerman', 'Juice ITC', 'Kristen ITC', 'Kunstler Script',
    'Leelawadee UI', 'Lucida Bright', 'Lucida Calligraphy', 'Lucida Console', 'Lucida Fax', 'Lucida Handwriting', 'Lucida Sans',
    'Lucida Sans Typewriter', 'Lucida Sans Unicode', 'Magneto', 'Maiandra GD', 'Malgun Gothic', 'Matura MT Script Capitals', 'Microsoft Himalaya',
    'Microsoft JhengHei', 'Microsoft New Tai Lue', 'Microsoft PhagsPa', 'Microsoft Sans Serif', 'Microsoft Tai Le', 'Microsoft YaHei',
    'Microsoft Yi Baiti', 'MingLiU-ExtB', 'Mistral', 'Modern No. 20', 'Mongolian Baiti', 'Monotype Corsiva', 'MS Gothic', 'MS Reference Sans Serif',
    'MS Reference Specialty', 'MV Boli', 'Myanmar Text', 'Niagara Engraved', 'Niagara Solid', 'Nirmala UI', 'OCR A Extended', 'Old English Text MT',
    'Onyx', 'Palace Script MT', 'Palatino Linotype', 'Papyrus', 'Parchment', 'Perpetua', 'Perpetua Titling MT', 'Playbill', 'Poor Richard',
    'Pristina', 'Rage Italic', 'Ravie', 'Rockwell', 'Rockwell Condensed', 'Rockwell Extra Bold', 'Script MT Bold', 'Segoe Print', 'Segoe Script',
    'Segoe UI', 'Segoe UI Black', 'Segoe UI Emoji', 'Segoe UI Light', 'Segoe UI Semibold', 'Segoe UI Symbol', 'Showcard Gothic', 'SimSun',
    'Sitka Text', 'Snap ITC', 'Stencil', 'Sylfaen', 'Symbol', 'Tahoma', 'Tempus Sans ITC', 'Times New Roman', 'Trebuchet MS', 'Tw Cen MT',
    'Verdana', 'Viner Hand ITC', 'Vivaldi', 'Vladimir Script', 'Webdings', 'Wide Latin', 'Wingdings', 'Wingdings 2', 'Wingdings 3', 'Yu Gothic'];
  let fonts = null;
  // Compara el ancho de un texto con la fuente pedida y con las genericas: si cambia en alguna, esta instalada.
  O.installedFonts = () => {
    if (fonts) return fonts;
    const ctx = document.createElement('canvas').getContext('2d');
    const sample = 'mmmmmmmmmmlliWWQ@#&0123456789';
    const width = (family) => { ctx.font = `72px ${family}`; return ctx.measureText(sample).width; };
    const bases = ['monospace', 'serif', 'sans-serif'].map((b) => [b, width(b)]);
    fonts = CANDIDATES.filter((f) => bases.some(([b, w]) => width(`"${f}", ${b}`) !== w));
    for (const f of ['Aptos', 'Calibri', 'Arial', 'Times New Roman']) if (!fonts.includes(f)) fonts.push(f);
    return fonts.sort((a, b) => a.localeCompare(b, 'es'));
  };
  O.fontOption = (f) => `<option style="font-family:'${O.esc(f)}'">${O.esc(f)}</option>`;
  // Anade a un desplegable de fuentes una que no esta en la lista (por ejemplo, la de un documento abierto).
  O.addFontOption = (sel, f) => {
    if (!f || [...sel.options].some((o) => o.value === f)) return;
    const opt = document.createElement('option');
    opt.textContent = f;
    opt.style.fontFamily = `'${f}'`;
    sel.insertBefore(opt, [...sel.options].find((o) => o.value.localeCompare(f, 'es') > 0) || null);
  };

  // ---------- Construccion de la cinta ----------
  // Boton: grande (icono arriba y texto debajo) o pequeno (icono y texto opcional en linea).
  O.rb = (id, icon, label, title, { big = false, text = !big ? '' : label, menu = false, attrs = '' } = {}) =>
    `<button class="rb${big ? ' rb-big' : ''}" ${id ? `id="${id}"` : ''} title="${O.esc(title || label)}" aria-label="${O.esc(label)}" ${attrs}>${
      big ? O.icon(icon, 32) : O.icon(icon)}${text ? `<span>${O.esc(text)}</span>` : ''}${menu ? '<i class="rb-arrow">▾</i>' : ''}</button>`;
  // Boton de color partido: la parte izquierda aplica el ultimo color y la flecha abre la paleta.
  O.rbColor = (id, glyph, title, color) =>
    `<span class="rb-split" title="${O.esc(title)}"><button class="rb rb-color" id="${id}" aria-label="${O.esc(title)}"><i class="glyph">${glyph}</i><i class="bar" style="background:${color}"></i></button><button class="rb rb-drop" id="${id}-menu" aria-label="${O.esc(title)}: más colores">▾</button></span>`;
  const group = (g) => `<div class="rg"><div class="rg-body">${g.html}</div><div class="rg-label">${O.esc(g.label)}</div></div>`;

  // tabs: [{ name, groups: [{ label, html }] }]. La pestana Archivo abre el menu de archivo.
  O.ribbon = (tabs, active = 1) => `
    <div class="ribbon" role="toolbar">
      <div class="rtabs" role="tablist">
        <button class="rtab rtab-file" data-file="1">Archivo</button>
        ${tabs.map((t, i) => `<button class="rtab${i + 1 === active ? ' on' : ''}" role="tab" data-tab="${i}">${O.esc(t.name)}</button>`).join('')}
      </div>
      <div class="rpanels">
        ${tabs.map((t, i) => `<div class="rpanel" data-panel="${i}"${i + 1 === active ? '' : ' hidden'}>${t.groups.map(group).join('<span class="rg-sep"></span>')}</div>`).join('')}
      </div>
    </div>`;

  O.ribbonInit = (root) => {
    const rib = root.querySelector('.ribbon');
    rib.addEventListener('mousedown', (e) => { if (e.target.closest('button')) e.preventDefault(); });
    rib.querySelectorAll('[data-tab]').forEach((b) => {
      b.onclick = () => {
        rib.querySelectorAll('[data-tab]').forEach((t) => t.classList.toggle('on', t === b));
        rib.querySelectorAll('[data-panel]').forEach((p) => { p.hidden = p.dataset.panel !== b.dataset.tab; });
        rib.classList.remove('collapsed');
      };
      // Como en Office, doble clic en una pestana oculta o muestra la cinta.
      b.ondblclick = () => rib.classList.toggle('collapsed');
    });
    rib.querySelector('[data-file]').onclick = (e) => {
      const r = e.currentTarget.getBoundingClientRect();
      const S = O.shell;
      O.contextMenu(r.left, r.bottom, [
        { label: 'Nuevo', action: () => S.newDoc(S.active.kind) },
        { label: 'Abrir', hint: 'Ctrl+O', action: () => document.getElementById('file-open').click() },
        { label: 'Guardar', hint: 'Ctrl+G', action: () => S.saveActive() },
        { label: 'Imprimir', hint: 'Ctrl+P', action: () => window.print() },
        '-',
        { label: 'Cambiar nombre', action: () => S.rename(S.active) },
        { label: 'Cerrar', action: () => S.close(S.active) },
      ]);
    };
  };

  // Abre un menu debajo de un boton.
  O.menuAt = (btn, items) => {
    const r = btn.getBoundingClientRect();
    O.contextMenu(r.left, r.bottom + 2, items);
  };

  // ---------- Paletas de color ----------
  const THEME = ['FFFFFF', '000000', 'E8E8E8', '0E2841', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E'];
  const STANDARD = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'];
  const HIGHLIGHT = ['FFFF00', '00FF00', '00FFFF', 'FF00FF', '0000FF', 'FF0000', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0', '000000'];
  const mix = (hex, k) => {
    const c = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    return c.map((v) => Math.round(k >= 0 ? v + (255 - v) * k : v * (1 + k)).toString(16).padStart(2, '0')).join('').toUpperCase();
  };
  // Variaciones de cada color del tema como en Office: mas claro 80, 60 y 40 %, mas oscuro 25 y 50 %.
  const shades = (hex) => {
    if (hex === 'FFFFFF') return [-0.05, -0.15, -0.25, -0.35, -0.5].map((k) => mix(hex, k));
    if (hex === '000000') return [0.5, 0.35, 0.25, 0.15, 0.05].map((k) => mix(hex, k));
    return [0.8, 0.6, 0.4, -0.25, -0.5].map((k) => mix(hex, k));
  };

  let openPalette = null;
  // kind: 'theme' (texto, relleno, contorno) o 'highlight' (los 15 colores de resaltado de Word).
  O.colorPalette = (btn, kind, onPick, { none = null, auto = null } = {}) => {
    if (openPalette) openPalette();
    const box = document.createElement('div');
    box.className = 'palette';
    const sw = (hex) => `<button class="sw" data-c="${hex}" title="#${hex}" style="background:#${hex}"></button>`;
    let html = '';
    if (auto) html += `<button class="pal-row" data-c="auto"><span class="sw-mini" style="background:#000"></span>${O.esc(auto)}</button>`;
    if (kind === 'highlight') html += `<div class="pal-grid hl">${HIGHLIGHT.map(sw).join('')}</div>`;
    else {
      html += `<div class="pal-title">Colores del tema</div><div class="pal-grid">${THEME.map(sw).join('')}</div>`;
      html += `<div class="pal-grid shades">${[0, 1, 2, 3, 4].map((r) => THEME.map((t) => sw(shades(t)[r])).join('')).join('')}</div>`;
      html += `<div class="pal-title">Colores estándar</div><div class="pal-grid">${STANDARD.map(sw).join('')}</div>`;
    }
    if (none) html += `<button class="pal-row" data-c="none"><span class="sw-mini none"></span>${O.esc(none)}</button>`;
    if (kind !== 'highlight') html += '<label class="pal-row pal-more">Más colores…<input type="color"></label>';
    box.innerHTML = html;
    document.body.append(box);
    const r = btn.getBoundingClientRect();
    const w = box.getBoundingClientRect();
    box.style.left = Math.max(4, Math.min(r.left, window.innerWidth - w.width - 4)) + 'px';
    box.style.top = Math.min(r.bottom + 2, window.innerHeight - w.height - 4) + 'px';
    const close = () => {
      box.remove();
      document.removeEventListener('mousedown', outside, true);
      document.removeEventListener('keydown', esc, true);
      openPalette = null;
    };
    const outside = (e) => { if (!box.contains(e.target)) close(); };
    const esc = (e) => { if (e.key === 'Escape') { e.stopPropagation(); close(); } };
    box.addEventListener('mousedown', (e) => { if (!e.target.closest('input')) e.preventDefault(); });
    box.onclick = (e) => {
      const b = e.target.closest('[data-c]');
      if (!b) return;
      close();
      onPick(b.dataset.c === 'none' || b.dataset.c === 'auto' ? b.dataset.c : b.dataset.c);
    };
    const more = box.querySelector('input[type=color]');
    if (more) more.onchange = () => { const v = O.hex(more.value); close(); onPick(v); };
    setTimeout(() => {
      document.addEventListener('mousedown', outside, true);
      document.addEventListener('keydown', esc, true);
    });
    openPalette = close;
  };
  O.setColorBar = (btn, hex) => {
    const bar = btn.querySelector('.bar');
    if (bar) bar.style.background = hex === 'none' ? 'transparent' : '#' + hex;
  };

  // ---------- Barra de estado: zoom como en Office ----------
  O.zoomControl = (id, onZoom) => `
    <div class="zoom" id="${id}">
      <button class="zbtn" data-z="-1" aria-label="Alejar">−</button>
      <input type="range" min="10" max="500" value="100" aria-label="Zoom">
      <button class="zbtn" data-z="1" aria-label="Acercar">+</button>
      <button class="zpct" title="Zoom">100 %</button>
    </div>`;
  // El deslizador de Office es lineal de 10 a 100 en la mitad izquierda y de 100 a 500 en la derecha.
  const toSlider = (z) => (z <= 100 ? (z - 10) / 90 * 50 : 50 + (z - 100) / 400 * 50);
  const fromSlider = (s) => Math.round(s <= 50 ? 10 + s / 50 * 90 : 100 + (s - 50) / 50 * 400);
  O.zoomInit = (box, onZoom) => {
    const range = box.querySelector('input');
    range.min = 0; range.max = 100; range.step = 0.5;
    const pct = box.querySelector('.zpct');
    const set = (z, notify = true) => {
      z = Math.max(10, Math.min(500, Math.round(z)));
      range.value = toSlider(z);
      pct.textContent = z + ' %';
      if (notify) onZoom(z);
    };
    set(100, false);
    range.oninput = () => set(fromSlider(+range.value));
    box.querySelectorAll('[data-z]').forEach((b) => {
      b.onclick = () => {
        const z = parseInt(pct.textContent, 10);
        set(b.dataset.z > 0 ? Math.floor(z / 10) * 10 + 10 : Math.ceil(z / 10) * 10 - 10);
      };
    });
    pct.onclick = () => O.contextMenu(pct.getBoundingClientRect().left, pct.getBoundingClientRect().top - 220,
      [200, 150, 100, 75, 50].map((z) => ({ label: z + ' %', action: () => set(z) })));
    return { set };
  };
})();
