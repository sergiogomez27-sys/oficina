// Menu contextual y cuadros de dialogo compartidos.
// items: { label, hint, disabled, action } o { label, items } para un submenu, o '-' como separador.
(function () {
  const O = window.Oficina;
  let closeCurrent = null;

  O.contextMenu = (x, y, items) => {
    if (closeCurrent) closeCurrent();
    const previous = document.activeElement;
    let child = null;

    const place = (panel, px, py) => {
      const box = panel.getBoundingClientRect();
      panel.style.left = Math.max(4, Math.min(px, window.innerWidth - box.width - 4)) + 'px';
      panel.style.top = Math.max(4, Math.min(py, window.innerHeight - box.height - 4)) + 'px';
    };
    const closeChild = () => {
      if (child) child.remove();
      child = null;
    };
    const buttonsOf = (panel) => [...panel.querySelectorAll('.ctx-item:not(:disabled)')];

    function buildPanel(list, isChild) {
      const panel = document.createElement('div');
      panel.className = 'ctx-menu';
      panel.setAttribute('role', 'menu');
      for (const it of list) {
        if (it === '-') {
          const sep = document.createElement('div');
          sep.className = 'ctx-sep';
          sep.setAttribute('role', 'separator');
          panel.append(sep);
          continue;
        }
        const b = document.createElement('button');
        b.className = 'ctx-item';
        b.setAttribute('role', 'menuitem');
        b.disabled = !!it.disabled;
        const label = document.createElement('span');
        label.textContent = it.label;
        // Dibujo propio delante del texto (solo HTML generado por Oficina, como los iconos del formato condicional).
        if (it.html) label.insertAdjacentHTML('afterbegin', `<span class="ctx-icons">${it.html}</span>`);
        b.append(label);
        if (it.hint || it.items) {
          const hint = document.createElement('span');
          hint.className = 'ctx-hint';
          hint.textContent = it.items ? '▸' : it.hint;
          b.append(hint);
        }
        const openChild = (focus) => {
          closeChild();
          child = buildPanel(it.items, true);
          document.body.append(child);
          const box = b.getBoundingClientRect();
          const width = child.getBoundingClientRect().width;
          place(child, box.right + width + 4 > window.innerWidth ? box.left - width : box.right, box.top - 5);
          if (focus) buttonsOf(child)[0].focus({ preventScroll: true });
        };
        b.onclick = () => {
          if (it.items) { openChild(true); return; }
          close();
          it.action();
        };
        if (!isChild) b.onmouseenter = () => { if (it.items && !it.disabled) openChild(false); else closeChild(); };
        if (it.items) b.dataset.submenu = '1';
        b.openChild = it.items ? openChild : null;
        panel.append(b);
      }
      panel.addEventListener('contextmenu', (e) => e.preventDefault());
      return panel;
    }

    const menu = buildPanel(items, false);
    const inside = (node) => menu.contains(node) || (child && child.contains(node));
    const onDown = (e) => { if (!inside(e.target)) close(); };
    const onScroll = (e) => { if (!inside(e.target)) close(); };
    const onClose = () => close();
    const onKey = (e) => {
      const active = document.activeElement;
      const panel = child && child.contains(active) ? child : menu;
      if (e.key === 'Escape' || (e.key === 'ArrowLeft' && panel === child)) {
        e.preventDefault();
        e.stopPropagation();
        if (panel === child) {
          const owner = menu.querySelector('.ctx-item[data-open="1"]') || buttonsOf(menu)[0];
          closeChild();
          owner.focus({ preventScroll: true });
        } else close();
        return;
      }
      if (e.key === 'ArrowRight' && active.openChild) {
        e.preventDefault();
        active.dataset.open = '1';
        active.openChild(true);
        return;
      }
      if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
      e.preventDefault();
      e.stopPropagation();
      const list = buttonsOf(panel);
      const i = list.indexOf(active);
      list[(i + (e.key === 'ArrowDown' ? 1 : list.length - 1)) % list.length].focus({ preventScroll: true });
    };
    function close() {
      closeCurrent = null;
      closeChild();
      menu.remove();
      document.removeEventListener('mousedown', onDown, true);
      document.removeEventListener('keydown', onKey, true);
      document.removeEventListener('scroll', onScroll, true);
      window.removeEventListener('resize', onClose);
      window.removeEventListener('blur', onClose);
      if (previous && previous.isConnected) previous.focus({ preventScroll: true });
    }
    closeCurrent = close;
    document.body.append(menu);
    // Se mantiene dentro de la ventana aunque el clic este cerca de un borde.
    place(menu, x, y);
    document.addEventListener('mousedown', onDown, true);
    document.addEventListener('keydown', onKey, true);
    document.addEventListener('scroll', onScroll, true);
    window.addEventListener('resize', onClose);
    window.addEventListener('blur', onClose);
    const first = buttonsOf(menu)[0];
    if (first) first.focus({ preventScroll: true });
  };

  // fields: { name, label, type: text|textarea|number|color|checkbox|select, value, options: [[valor, texto]] }.
  // onOk recibe los valores y puede devolver un texto de error para mantener abierto el cuadro.
  O.formDialog = ({ title, fields, okLabel, onOk }) => {
    const previous = document.activeElement;
    const dlg = document.createElement('dialog');
    dlg.className = 'form-dialog';
    const form = document.createElement('form');
    const h = document.createElement('h2');
    h.textContent = title;
    form.append(h);
    const inputs = {};
    for (const f of fields) {
      const row = document.createElement('label');
      row.className = 'form-row' + (f.type === 'checkbox' ? ' check' : '');
      const text = document.createElement('span');
      text.textContent = f.label;
      let input;
      if (f.type === 'textarea') {
        input = document.createElement('textarea');
        input.rows = 4;
      } else if (f.type === 'select') {
        input = document.createElement('select');
        for (const [value, label] of f.options) {
          const opt = document.createElement('option');
          opt.value = value;
          opt.textContent = label;
          input.append(opt);
        }
      } else {
        input = document.createElement('input');
        input.type = f.type || 'text';
      }
      input.name = f.name;
      if (f.type === 'checkbox') input.checked = !!f.value;
      else if (f.value !== undefined && f.value !== null) input.value = f.value;
      inputs[f.name] = input;
      if (f.type === 'checkbox') row.append(input, text);
      else row.append(text, input);
      form.append(row);
    }
    const error = document.createElement('p');
    error.className = 'form-error';
    error.setAttribute('role', 'alert');
    const actions = document.createElement('div');
    actions.className = 'form-actions';
    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.className = 'btn';
    cancel.textContent = 'Cancelar';
    const ok = document.createElement('button');
    ok.type = 'submit';
    ok.className = 'btn primary';
    ok.textContent = okLabel || 'Aceptar';
    actions.append(cancel, ok);
    form.append(error, actions);
    dlg.append(form);
    const finish = () => {
      dlg.close();
      dlg.remove();
      if (previous && previous.isConnected) previous.focus({ preventScroll: true });
    };
    cancel.onclick = finish;
    dlg.addEventListener('cancel', (e) => { e.preventDefault(); finish(); });
    // Las teclas del cuadro no deben llegar a la hoja que hay debajo.
    dlg.addEventListener('keydown', (e) => e.stopPropagation());
    form.onsubmit = (e) => {
      e.preventDefault();
      const values = {};
      for (const f of fields) values[f.name] = f.type === 'checkbox' ? inputs[f.name].checked : inputs[f.name].value;
      const problem = onOk(values);
      if (problem) { error.textContent = problem; return; }
      finish();
    };
    document.body.append(dlg);
    dlg.showModal();
  };
})();
