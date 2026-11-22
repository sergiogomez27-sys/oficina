// Tacto para movil y tablet: el dedo arrastra formas y graficos, y el doble toque hace lo que el doble clic.
(function () {
  // Lo que se arrastra con el dedo: formas de Presentaciones y objetos de Calculo (graficos, imagenes, segmentaciones).
  const DRAG = '.stage .el, .stage [data-handle], .dr-layer .dr-item';
  const SKIP = '.editing, input, textarea, select, [contenteditable="true"]';
  let active = null;
  let lastTap = null;

  const mouse = (type, t, target) => target.dispatchEvent(new MouseEvent(type, {
    bubbles: true, cancelable: true, view: window, clientX: t.clientX, clientY: t.clientY,
    screenX: t.screenX, screenY: t.screenY, button: 0, buttons: type === 'mouseup' || type === 'click' ? 0 : 1,
  }));
  const elOf = (n) => (n && n.nodeType === 1 ? n : n && n.parentElement);

  document.addEventListener('touchstart', (e) => {
    active = null;
    if (e.touches.length !== 1) return;
    const target = elOf(e.target);
    if (!target || !target.closest(DRAG) || target.closest(SKIP)) return;
    const t = e.touches[0];
    active = { target, x: t.clientX, y: t.clientY, moved: false };
    // Sin esto el navegador desplazaria la pagina y luego simularia un clic.
    e.preventDefault();
    mouse('mousedown', t, target);
  }, { passive: false, capture: true });

  document.addEventListener('touchmove', (e) => {
    if (!active) return;
    e.preventDefault();
    const t = e.touches[0];
    if (Math.hypot(t.clientX - active.x, t.clientY - active.y) > 6) active.moved = true;
    mouse('mousemove', t, active.target);
  }, { passive: false, capture: true });

  const end = (e) => {
    const t = e.changedTouches[0];
    const target = elOf(e.target);
    if (active) {
      const a = active;
      active = null;
      mouse('mouseup', t, a.target);
      if (!a.moved) mouse('click', t, a.target);
    }
    // Doble toque = doble clic (editar una celda, una forma o un grafico).
    if (e.type !== 'touchend' || !target || !target.closest('.grid-wrap, .stage, .dr-layer')) return;
    const now = Date.now();
    if (lastTap && now - lastTap.time < 350 && Math.hypot(t.clientX - lastTap.x, t.clientY - lastTap.y) < 24 && !target.closest(SKIP)) {
      lastTap = null;
      e.preventDefault();
      mouse('dblclick', t, target);
    } else {
      lastTap = { time: now, x: t.clientX, y: t.clientY };
    }
  };
  document.addEventListener('touchend', end, { passive: false, capture: true });
  document.addEventListener('touchcancel', end, { passive: false, capture: true });
})();
