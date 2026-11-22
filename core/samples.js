// Documentos en blanco y de ejemplo. Los ejemplos estan marcados como tales en su titulo.
(function (g) {
  const O = g.Oficina;
  const box = (id, x, y, w, h, size, html, extra = {}) => ({
    id, type: 'shape', geom: 'rect', x, y, w, h, fill: null, line: null, size, color: '1F2A33', align: 'left', anchor: 't', bold: false, html, ...extra,
  });

  O.samples = {
    writer: {
      blank: () => ({ html: '<p><br></p>' }),
      sample: () => ({
        html: '<h1>Acta de la reunión de vecinos</h1><p><span style="color:#5B6B75">Ejemplo. Sustituye este texto por el tuyo.</span></p><h2>Asistentes</h2><ul><li>Portal 1: 6 viviendas</li><li>Portal 2: 4 viviendas</li></ul><h2>Acuerdos</h2><ol><li>Reparar la puerta del garaje antes de fin de mes.</li><li>Pedir dos presupuestos para pintar el portal.</li></ol><table><tbody><tr><td><b>Tarea</b></td><td><b>Responsable</b></td><td><b>Fecha</b></td></tr><tr><td>Puerta del garaje</td><td>Presidencia</td><td>30/11/2026</td></tr><tr><td>Presupuestos</td><td>Administración</td><td>15/12/2026</td></tr></tbody></table><p><br></p>',
      }),
    },
    sheet: {
      blank: () => ({ sheets: [{ name: 'Hoja1', cells: {}, cols: {} }], active: 0 }),
      sample: () => {
        const head = { b: true, fill: 'DCE8E2' };
        const money = { nf: '#,##0.00 "€"' };
        const rows = [['Vivienda', 180, 72.4], ['Suministros', 64.5, 58.1], ['Transporte', 40, 55], ['Ocio', 90, 120.75]];
        const cells = {
          A1: { v: 'Concepto', s: head }, B1: { v: 'Presupuesto', s: head }, C1: { v: 'Gasto', s: head }, D1: { v: 'Diferencia', s: head },
        };
        rows.forEach((r, i) => {
          const n = i + 2;
          cells['A' + n] = { v: r[0] };
          cells['B' + n] = { v: r[1], s: money };
          cells['C' + n] = { v: r[2], s: money };
          cells['D' + n] = { f: `B${n}-C${n}`, s: money };
        });
        cells.A6 = { v: 'Total', s: { b: true } };
        for (const c of 'BCD') cells[c + '6'] = { f: `SUM(${c}2:${c}5)`, s: { b: true, ...money } };
        cells.A8 = { v: 'Partidas excedidas' };
        cells.B8 = { f: 'COUNTIF(D2:D5,"<0")' };
        return { sheets: [{ name: 'Gastos (ejemplo)', cells, cols: { 0: 150, 1: 100, 2: 100, 3: 100 } }], active: 0 };
      },
    },
    slides: {
      // Presentacion nueva de PowerPoint 2024: 16:9 con la diapositiva de titulo en Aptos Display y Aptos.
      blank: () => ({
        w: 960, h: 540, nextId: 3, fonts: { major: 'Aptos Display', minor: 'Aptos' },
        slides: [{ bg: 'FFFFFF', els: [
          box(1, 120, 88.4, 720, 188, 60, '<div><br></div>', { ph: 'ctrTitle', color: '000000', align: 'center', anchor: 'b' }),
          box(2, 120, 283.6, 720, 130.4, 24, '<div><br></div>', { ph: 'subTitle', color: '000000', align: 'center' }),
        ] }],
      }),
      sample: () => ({
        w: 960, h: 540, nextId: 8,
        slides: [
          { bg: '1F3A5F', els: [
            box(1, 70, 170, 820, 110, 48, '<div>Plan del huerto comunitario</div>', { color: 'FFFFFF', bold: true, anchor: 'b' }),
            box(2, 70, 295, 820, 50, 22, '<div>Ejemplo de presentación</div>', { color: 'C9D8E6' }),
          ] },
          { bg: 'FFFFFF', els: [
            box(3, 60, 40, 840, 70, 34, '<div>Primeros pasos</div>', { color: '1F3A5F', bold: true }),
            box(4, 60, 130, 500, 330, 22, '<div class="bu">Preparar 12 bancales</div><div class="bu">Instalar riego por goteo</div><div class="bu">Plantar en marzo</div>'),
            box(5, 620, 150, 270, 270, 26, '<div>12</div><div><span style="font-size:16pt">bancales</span></div>', { type: 'shape', geom: 'ellipse', fill: 'B4532A', color: 'FFFFFF', align: 'center', anchor: 'ctr', bold: true }),
          ] },
        ],
      }),
    },
  };
})(globalThis);
