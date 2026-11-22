// Lanzador: inicio, pestanas, abrir y guardar, y borradores. Equivale a la capa compartida que usan las tres aplicaciones.
(function () {
  const O = window.Oficina;
  const $ = (id) => document.getElementById(id);
  const DRAFTS = 'oficina.borradores.v1';
  const KINDS = {
    docx: { kind: 'writer', io: O.docx }, xlsx: { kind: 'sheet', io: O.xlsx }, pptx: { kind: 'slides', io: O.pptx },
    xlsm: { kind: 'sheet', io: O.xlsx },
  };
  const extOf = (kind) => Object.keys(KINDS).find((e) => KINDS[e].kind === kind);
  const docs = [];
  let active = null;
  let counter = 0;

  const notice = (text) => {
    const n = $('notice');
    n.replaceChildren();
    const span = document.createElement('span');
    span.textContent = text;
    const b = document.createElement('button');
    b.textContent = 'Cerrar';
    b.onclick = () => { n.hidden = true; };
    n.append(span, b);
    n.hidden = false;
  };
  O.notice = notice;

  // ---------- Borradores: comodidad local, nunca sustituyen a guardar el archivo ----------
  const readDrafts = () => {
    try { return JSON.parse(localStorage.getItem(DRAFTS)) || {}; } catch { return {}; }
  };
  const writeDraft = (doc) => {
    const all = readDrafts();
    all[doc.id] = { id: doc.id, kind: doc.kind, title: doc.title, at: Date.now(), state: doc.view.getState() };
    try {
      localStorage.setItem(DRAFTS, JSON.stringify(all));
      return true;
    } catch { return false; }
  };
  const dropDraft = (id) => {
    const all = readDrafts();
    delete all[id];
    try { localStorage.setItem(DRAFTS, JSON.stringify(all)); } catch { /* sin almacenamiento */ }
  };

  const setStatus = (doc) => {
    if (doc !== active) return;
    $('save-state').textContent = doc.dirty ? (doc.draftOk === false ? 'Sin guardar (borrador demasiado grande)' : 'Cambios en borrador') : 'Guardado';
  };
  const touched = (doc) => {
    doc.dirty = true;
    setStatus(doc);
    clearTimeout(doc.timer);
    doc.timer = setTimeout(() => { doc.draftOk = writeDraft(doc); setStatus(doc); }, 800);
    renderTabs();
  };

  // ---------- Documentos ----------
  function openDoc({ kind, title, state, id, dirty = false }) {
    const pane = document.createElement('div');
    pane.className = `pane kind-${kind}`;
    pane.hidden = true;
    $('workspace').append(pane);
    const doc = { id: id || `d${Date.now()}-${++counter}`, kind, title, pane, dirty };
    doc.view = O.apps[kind].mount(pane, state, () => touched(doc));
    docs.push(doc);
    activate(doc);
    return doc;
  }

  function activate(doc) {
    active = doc;
    for (const d of docs) {
      d.pane.hidden = d !== doc;
      if (d === doc && d.view.onShow) d.view.onShow();
    }
    homePane.hidden = !!doc;
    $('doc-actions').hidden = !doc;
    // La barra de titulo toma el color del programa, como en Office.
    document.body.dataset.app = doc ? doc.kind : '';
    $('qat').hidden = !doc;
    if (doc) {
      $('qat-undo').hidden = !doc.view.undo;
      $('qat-redo').hidden = !doc.view.redo;
    }
    if (doc) setStatus(doc);
    else renderHome();
    renderTabs();
  }

  function closeDoc(doc) {
    if (doc.dirty && !window.confirm(`«${doc.title}» tiene cambios sin guardar en un archivo. ¿Cerrar y borrar su borrador?`)) return;
    clearTimeout(doc.timer);
    dropDraft(doc.id);
    if (doc.view.destroy) doc.view.destroy();
    doc.pane.remove();
    docs.splice(docs.indexOf(doc), 1);
    activate(active === doc ? docs[docs.length - 1] || null : active);
  }

  function renderTabs() {
    const box = $('tabs');
    box.replaceChildren();
    for (const d of docs) {
      const t = document.createElement('div');
      t.className = `tab kind-${d.kind}${d === active ? ' active' : ''}`;
      t.title = 'Doble clic para cambiar el nombre';
      const name = document.createElement('span');
      name.className = 'tab-name';
      name.textContent = (d.dirty ? '• ' : '') + d.title;
      const close = document.createElement('button');
      close.className = 'tab-close';
      close.textContent = '×';
      close.setAttribute('aria-label', `Cerrar ${d.title}`);
      close.onclick = (e) => { e.stopPropagation(); closeDoc(d); };
      t.onclick = () => activate(d);
      t.ondblclick = () => rename(d);
      t.append(name, close);
      box.append(t);
    }
  }

  function rename(d) {
    const n = window.prompt('Nombre del documento', d.title);
    if (n && n.trim()) { d.title = n.trim(); touched(d); }
  }
  const newDoc = (kind) => {
    const c = CARDS.find((x) => x.kind === kind);
    return openDoc({ kind, title: `${c.blank} ${++counter}`, state: O.samples[kind].blank() });
  };

  async function saveActive() {
    const doc = active;
    if (!doc) return;
    if (doc.view.flush) doc.view.flush();
    const io = KINDS[extOf(doc.kind)].io;
    const ext = io.ext ? io.ext(doc.view.getState()) : extOf(doc.kind);
    try {
      const blob = await io.write(doc.view.getState());
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `${doc.title.replace(/[\\/:*?"<>|]/g, '_')}.${ext}`;
      document.body.append(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 10000);
      doc.dirty = false;
      setStatus(doc);
      renderTabs();
    } catch (e) {
      notice(`No se pudo guardar el archivo: ${e.message}`);
    }
  }

  async function openFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    const type = KINDS[ext];
    if (!type) {
      notice(`«${file.name}» no se puede abrir. Oficina trabaja con archivos .docx, .xlsx, .xlsm y .pptx.`);
      return;
    }
    try {
      const { state, warnings } = await type.io.read(await file.arrayBuffer());
      openDoc({ kind: type.kind, title: file.name.replace(/\.[^.]+$/, ''), state });
      if (warnings.length) notice(`Abierto. Oficina no conserva: ${warnings.join(', ')}. Si guardas, se perderán en el archivo nuevo.`);
    } catch (e) {
      console.error(e);
      notice(`No se pudo abrir «${file.name}». Puede estar dañado o protegido con contraseña.`);
    }
  }
  const openFiles = (files) => { [...files].forEach(openFile); };

  // ---------- Inicio ----------
  const homePane = document.createElement('div');
  homePane.className = 'pane';
  $('workspace').append(homePane);

  const CARDS = [
    { kind: 'writer', name: 'Textos', ext: '.docx', blank: 'Documento', text: 'Cartas, informes y actas con tablas e imágenes.' },
    { kind: 'sheet', name: 'Cálculo', ext: '.xlsx', blank: 'Libro', text: 'Hojas con fórmulas, formatos y varias pestañas.' },
    { kind: 'slides', name: 'Presentaciones', ext: '.pptx', blank: 'Presentación', text: 'Diapositivas con texto, formas e imágenes.' },
  ];

  function renderHome() {
    const home = document.createElement('div');
    home.className = 'home';
    const inner = document.createElement('div');
    inner.className = 'home-inner';
    home.append(inner);
    inner.innerHTML = '<div><h1>¿Qué quieres hacer hoy?</h1><p>Tus archivos se quedan en este equipo. Oficina funciona sin conexión.</p></div>';

    const grid = document.createElement('section');
    grid.className = 'new-grid';
    for (const c of CARDS) {
      const card = document.createElement('article');
      card.className = `new-card kind-${c.kind}`;
      card.innerHTML = `<h3>${c.name}</h3><small>${c.text} Compatible con ${c.ext}.</small><div class="row"></div>`;
      const row = card.querySelector('.row');
      const mk = (label, cls, fn) => {
        const b = document.createElement('button');
        b.className = 'btn ' + cls;
        b.textContent = label;
        b.onclick = fn;
        row.append(b);
      };
      mk('Nuevo', '', () => newDoc(c.kind));
      mk('Ver ejemplo', '', () => openDoc({ kind: c.kind, title: `Ejemplo de ${c.name.toLowerCase()}`, state: O.samples[c.kind].sample() }));
      grid.append(card);
    }
    inner.append(grid);

    const drop = document.createElement('section');
    drop.className = 'drop';
    drop.innerHTML = '<span>Arrastra aquí un archivo de Word, Excel o PowerPoint</span>';
    const pick = document.createElement('button');
    pick.className = 'btn primary';
    pick.textContent = 'Abrir archivo';
    pick.onclick = () => $('file-open').click();
    drop.append(pick);
    inner.append(drop);

    const drafts = Object.values(readDrafts()).filter((d) => !docs.some((o) => o.id === d.id)).sort((a, b) => b.at - a.at);
    if (drafts.length) {
      const sec = document.createElement('section');
      sec.innerHTML = '<h2>Borradores sin cerrar</h2>';
      const ul = document.createElement('ul');
      ul.className = 'draft-list';
      for (const d of drafts) {
        const li = document.createElement('li');
        li.className = `kind-${d.kind}`;
        const dot = document.createElement('span');
        dot.className = 'dot';
        const name = document.createElement('span');
        name.className = 'name';
        name.textContent = d.title;
        const time = document.createElement('time');
        time.textContent = new Date(d.at).toLocaleString('es-ES', { dateStyle: 'short', timeStyle: 'short' });
        const open = document.createElement('button');
        open.className = 'linkish';
        open.textContent = 'Abrir';
        open.onclick = () => openDoc({ kind: d.kind, title: d.title, state: d.state, id: d.id, dirty: true });
        const del = document.createElement('button');
        del.className = 'linkish';
        del.textContent = 'Eliminar';
        del.onclick = () => { dropDraft(d.id); renderHome(); };
        li.append(dot, name, time, open, del);
        ul.append(li);
      }
      sec.append(ul);
      inner.append(sec);
    }
    homePane.replaceChildren(home);
  }

  // ---------- Eventos globales ----------
  $('go-home').onclick = () => activate(null);
  $('btn-save').onclick = saveActive;
  $('btn-print').onclick = () => window.print();
  $('qat-save').onclick = saveActive;
  $('qat-undo').onclick = () => active && active.view.undo && active.view.undo();
  $('qat-redo').onclick = () => active && active.view.redo && active.view.redo();
  $('qat').addEventListener('mousedown', (e) => e.preventDefault());
  $('file-open').onchange = (e) => { openFiles(e.target.files); e.target.value = ''; };
  document.addEventListener('keydown', (e) => {
    if (e.defaultPrevented) return;
    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && 'sg'.includes(e.key.toLowerCase()) && e.key.length === 1) { e.preventDefault(); saveActive(); }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') { e.preventDefault(); $('file-open').click(); }
  });
  document.addEventListener('dragover', (e) => {
    e.preventDefault();
    const d = document.querySelector('.drop');
    if (d) d.classList.add('over');
  });
  document.addEventListener('dragleave', () => {
    const d = document.querySelector('.drop');
    if (d) d.classList.remove('over');
  });
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    if (e.dataTransfer.files.length) openFiles(e.dataTransfer.files);
  });
  window.addEventListener('beforeunload', () => {
    for (const d of docs) if (d.dirty) writeDraft(d);
  });

  O.shell = { docs, get active() { return active; }, openDoc, openFile, saveActive, activate, newDoc, rename, close: closeDoc };
  activate(null);
})();
