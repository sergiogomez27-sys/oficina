// Lectura y escritura de .docx. El editor trabaja con HTML; aqui se convierte en ambos sentidos.
(function (g) {
  const O = g.Oficina;
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const EMU_PX = 9525;
  // Altura de una linea sencilla respecto al tamano de letra; el interlineado multiple de Word se multiplica por ella.
  const LINE = 1.2;
  const pt = (twips) => Math.round(+twips / 2) / 10;

  // ---------- Lectura ----------
  async function readDocx(buffer) {
    const zip = await g.JSZip.loadAsync(buffer);
    const main = await O.mainPart(zip);
    const doc = await O.readXml(zip, main);
    const rels = await O.readRels(zip, main);
    const warnings = new Set();
    const relOf = (suffix) => Object.values(rels).find((r) => r.type.endsWith(suffix));

    // Fuentes del tema, para los estilos que dicen "fuente de titulos" o "fuente de cuerpo".
    const themeFonts = {};
    const themeRel = relOf('/theme');
    if (themeRel) {
      const xml = await O.readXml(zip, themeRel.target);
      for (const k of ['major', 'minor']) {
        const f = xml && O.all(xml, k + 'Font')[0];
        const latin = f && O.kid(f, 'latin');
        if (latin && O.attr(latin, 'typeface')) themeFonts[k] = O.attr(latin, 'typeface');
      }
    }
    const fontOf = (rFonts) => {
      if (!rFonts) return null;
      const theme = O.attr(rFonts, 'asciiTheme') || O.attr(rFonts, 'hAnsiTheme');
      if (theme) return themeFonts[theme.startsWith('major') ? 'major' : 'minor'] || null;
      return O.attr(rFonts, 'ascii') || O.attr(rFonts, 'hAnsi') || null;
    };
    // Formato de un estilo o de los valores por defecto: fuente, tamano, color, negrita y espaciado.
    const props = (rPr, pPr) => {
      const p = {};
      const font = fontOf(O.kid(rPr, 'rFonts'));
      if (font) p.font = font;
      const sz = O.attr(O.kid(rPr, 'sz'), 'val');
      if (sz) p.size = +sz / 2;
      const color = O.attr(O.kid(rPr, 'color'), 'val');
      if (color && color !== 'auto') p.color = color.toUpperCase();
      if (O.kid(rPr, 'b')) p.b = O.isOn(O.kid(rPr, 'b'));
      const sp = O.kid(pPr, 'spacing');
      if (O.attr(sp, 'before') !== null) p.before = pt(O.attr(sp, 'before'));
      if (O.attr(sp, 'after') !== null) p.after = pt(O.attr(sp, 'after'));
      const rule = O.attr(sp, 'lineRule');
      if (O.attr(sp, 'line') !== null && (!rule || rule === 'auto')) p.line = Math.round(+O.attr(sp, 'line') / 240 * 100) / 100;
      return p;
    };
    const styles = {};
    const styleEls = {};
    let base = null;
    const heads = {};
    const stylesRel = relOf('/styles');
    if (stylesRel) {
      const xml = await O.readXml(zip, stylesRel.target);
      const root = xml && xml.documentElement;
      for (const s of O.kids(root, 'style')) {
        styles[O.attr(s, 'styleId')] = (O.attr(O.kid(s, 'name'), 'val') || '').toLowerCase();
        styleEls[O.attr(s, 'styleId')] = s;
      }
      const defs = O.kid(root, 'docDefaults');
      base = props(O.path(defs, 'rPrDefault', 'rPr'), O.path(defs, 'pPrDefault', 'pPr'));
      const normal = O.kids(root, 'style').find((s) => O.attr(s, 'type') === 'paragraph' && ['1', 'true', 'on'].includes(O.attr(s, 'default')));
      if (normal) Object.assign(base, props(O.kid(normal, 'rPr'), O.kid(normal, 'pPr')));
      // Lo que nadie define toma el valor de la especificacion: Times New Roman 10, sin espacio, interlineado sencillo.
      base = { font: 'Times New Roman', size: 10, after: 0, line: 1, ...base };
      delete base.color; delete base.b; delete base.before;
      for (const [id, name] of Object.entries(styles)) {
        const m = /^heading ([1-3])$/.exec(name);
        if (!m) continue;
        const st = styleEls[id];
        heads['H' + m[1]] = { font: base.font, size: base.size, color: '000000', b: false, before: 0, after: base.after, ...props(O.kid(st, 'rPr'), O.kid(st, 'pPr')) };
        delete heads['H' + m[1]].line;
      }
    }
    const numbering = {};
    const numRel = relOf('/numbering');
    if (numRel) {
      const xml = await O.readXml(zip, numRel.target);
      const root = xml && xml.documentElement;
      const abstract = {};
      for (const a of O.kids(root, 'abstractNum')) {
        const lvls = {};
        for (const l of O.kids(a, 'lvl')) lvls[O.attr(l, 'ilvl')] = O.attr(O.kid(l, 'numFmt'), 'val');
        abstract[O.attr(a, 'abstractNumId')] = lvls;
      }
      for (const n of O.kids(root, 'num')) numbering[O.attr(n, 'numId')] = abstract[O.attr(O.kid(n, 'abstractNumId'), 'val')] || {};
    }
    if (Object.keys(zip.files).some((f) => /word\/(header|footer)\d*\.xml$/.test(f))) warnings.add('encabezados y pies de página');
    if (zip.file('word/comments.xml')) warnings.add('comentarios');
    if (zip.file('word/footnotes.xml') && O.all(doc, 'footnoteReference').length) warnings.add('notas al pie');

    const imageHtml = async (drawing) => {
      const blip = O.all(drawing, 'blip')[0];
      const rel = blip && rels[O.attr(blip, 'embed')];
      const url = rel && (await O.partToDataUrl(zip, rel.target));
      if (!url) { warnings.add('gráficos o formas dibujadas'); return ''; }
      const ext = O.all(drawing, 'extent')[0];
      const w = ext ? Math.round(+O.attr(ext, 'cx') / EMU_PX) : '';
      const h = ext ? Math.round(+O.attr(ext, 'cy') / EMU_PX) : '';
      return `<img src="${url}" width="${w}" height="${h}" alt="">`;
    };

    const runHtml = async (run) => {
      const pr = O.kid(run, 'rPr');
      const rs = {};
      if (pr) {
        rs.b = O.isOn(O.kid(pr, 'b'));
        rs.i = O.isOn(O.kid(pr, 'i'));
        rs.u = O.isOn(O.kid(pr, 'u'));
        rs.s = O.isOn(O.kid(pr, 'strike'));
        const color = O.attr(O.kid(pr, 'color'), 'val');
        if (color && color !== 'auto') rs.color = color;
        const sz = O.attr(O.kid(pr, 'sz'), 'val');
        if (sz) rs.size = +sz / 2;
        const font = O.attr(O.kid(pr, 'rFonts'), 'ascii');
        if (font) rs.font = font;
        const shd = O.attr(O.kid(pr, 'shd'), 'fill');
        if (shd && shd !== 'auto') rs.bg = shd;
        const hl = O.attr(O.kid(pr, 'highlight'), 'val');
        if (hl && hl !== 'none') rs.bg = O.hex(hl) || rs.bg;
        const va = O.attr(O.kid(pr, 'vertAlign'), 'val');
        if (va === 'subscript' || va === 'superscript') rs.va = va;
      }
      let inner = '';
      for (const c of O.kids(run)) {
        const n = c.localName;
        if (n === 't') inner += O.esc(c.textContent);
        else if (n === 'tab') inner += '&emsp;';
        else if (n === 'noBreakHyphen') inner += '-';
        else if (n === 'br') inner += O.attr(c, 'type') === 'page' ? '<!--salto-->' : '<br>';
        else if (n === 'drawing') inner += await imageHtml(c);
        else if (n === 'pict' || n === 'AlternateContent' || n === 'object') warnings.add('gráficos o formas dibujadas');
      }
      if (!inner) return '';
      const css = O.runCss(rs);
      return css ? `<span style="${css}">${inner}</span>` : inner;
    };

    const inlineHtml = async (parent) => {
      let out = '';
      for (const c of O.kids(parent)) {
        const n = c.localName;
        if (n === 'r') out += await runHtml(c);
        else if (n === 'hyperlink') {
          const rel = rels[O.attr(c, 'id')];
          const inner = await inlineHtml(c);
          out += rel && rel.external ? `<a href="${O.esc(rel.target)}">${inner}</a>` : inner;
        } else if (n === 'ins') {
          warnings.add('control de cambios');
          out += await inlineHtml(c);
        } else if (n === 'smartTag' || n === 'fldSimple') out += await inlineHtml(c);
        else if (n === 'sdt') out += await inlineHtml(O.kid(c, 'sdtContent'));
        else if (n === 'del') warnings.add('control de cambios');
      }
      return out;
    };

    const paragraphTag = (p) => {
      const name = styles[O.attr(O.path(p, 'pPr', 'pStyle'), 'val')] || '';
      if (name === 'title' || name === 'heading 1') return 'h1';
      if (name === 'heading 2') return 'h2';
      if (/^heading [3-9]$/.test(name)) return 'h3';
      return 'p';
    };

    const blocks = async (parent) => {
      let out = '', list = null;
      const closeList = () => {
        if (list) { out += `</${list.tag}>`; list = null; }
      };
      const items = O.kids(parent).flatMap((el) => (el.localName === 'sdt' ? O.kids(O.kid(el, 'sdtContent')) : [el]));
      for (const el of items) {
        if (el.localName === 'p') {
          const pPr = O.kid(el, 'pPr');
          const jc = O.attr(O.kid(pPr, 'jc'), 'val');
          const align = { center: 'center', right: 'right', end: 'right', both: 'justify' }[jc];
          // Interlineado, espacio despues y sangria propios del parrafo.
          let pcss = align ? `text-align:${align};` : '';
          const sp = O.kid(pPr, 'spacing');
          const rule = O.attr(sp, 'lineRule');
          if (O.attr(sp, 'line') && (!rule || rule === 'auto')) pcss += `line-height:${(+O.attr(sp, 'line') / 240 * LINE).toFixed(3)};`;
          if (O.attr(sp, 'after') !== null) pcss += `margin-bottom:${pt(O.attr(sp, 'after'))}pt;`;
          const indLeft = O.attr(O.kid(pPr, 'ind'), 'left') || O.attr(O.kid(pPr, 'ind'), 'start');
          if (indLeft && +indLeft > 0 && !O.kid(pPr, 'numPr')) pcss += `margin-left:${pt(indLeft)}pt;`;
          const style = pcss ? ` style="${pcss}"` : '';
          const numPr = O.kid(pPr, 'numPr');
          const numId = O.attr(O.kid(numPr, 'numId'), 'val');
          let inner = await inlineHtml(el);
          const hasBreak = inner.includes('<!--salto-->');
          inner = inner.replace(/<!--salto-->/g, '');
          if (numId && numId !== '0') {
            const lvl = O.attr(O.kid(numPr, 'ilvl'), 'val') || '0';
            const fmt = (numbering[numId] || {})[lvl] || 'bullet';
            const tag = fmt === 'bullet' ? 'ul' : 'ol';
            if (!list || list.tag !== tag || list.numId !== numId) {
              closeList();
              out += `<${tag}>`;
              list = { tag, numId };
            }
            const lvlAttr = lvl !== '0' ? ` data-level="${lvl}" style="margin-left:${+lvl * 24}px"` : '';
            out += `<li${lvlAttr}>${inner || '<br>'}</li>`;
          } else {
            closeList();
            const tag = paragraphTag(el);
            out += `<${tag}${style}>${inner || '<br>'}</${tag}>`;
          }
          if (hasBreak) { closeList(); out += '<hr class="salto">'; }
        } else if (el.localName === 'tbl') {
          closeList();
          out += '<table><tbody>';
          for (const tr of O.kids(el, 'tr')) {
            out += '<tr>';
            for (const tc of O.kids(tr, 'tc')) {
              const span = O.attr(O.path(tc, 'tcPr', 'gridSpan'), 'val');
              const merge = O.path(tc, 'tcPr', 'vMerge');
              if (merge && O.attr(merge, 'val') !== 'restart') { out += '<td></td>'; warnings.add('celdas combinadas verticalmente'); continue; }
              out += `<td${span ? ` colspan="${O.esc(span)}"` : ''}>${await blocks(tc)}</td>`;
            }
            out += '</tr>';
          }
          out += '</tbody></table>';
        }
      }
      closeList();
      return out;
    };

    const html = await blocks(O.path(doc.documentElement, 'body'));
    const state = { html: html || '<p><br></p>' };
    if (base) { state.base = base; state.heads = heads; }
    // Tamano, orientacion y margenes de la pagina (en twips, como Word).
    const sect = O.path(doc.documentElement, 'body', 'sectPr');
    if (sect) {
      const sz = O.path(sect, 'pgSz'), mar = O.path(sect, 'pgMar');
      const n = (el, k) => { const v = parseInt(O.attr(el, k), 10); return Number.isFinite(v) ? v : undefined; };
      const pg = { w: n(sz, 'w'), h: n(sz, 'h'), top: n(mar, 'top'), right: n(mar, 'right'), bottom: n(mar, 'bottom'), left: n(mar, 'left') };
      for (const k of Object.keys(pg)) if (pg[k] === undefined) delete pg[k];
      if (Object.keys(pg).length) state.page = pg;
    }
    return { state, warnings: [...warnings] };
  }

  // ---------- Escritura ----------
  const BLOCK = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'TABLE', 'HR', 'BLOCKQUOTE', 'SECTION', 'ARTICLE']);
  const HEADING = { H1: 'Heading1', H2: 'Heading2', H3: 'Heading3', H4: 'Heading3', H5: 'Heading3', H6: 'Heading3' };

  function rPr(s) {
    let x = '';
    if (s.font) { const f = O.esc(s.font); x += `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`; }
    if (s.b) x += '<w:b/>';
    if (s.i) x += '<w:i/>';
    if (s.s) x += '<w:strike/>';
    if (s.color) x += `<w:color w:val="${s.color}"/>`;
    if (s.size) { const v = Math.round(s.size * 2); x += `<w:sz w:val="${v}"/><w:szCs w:val="${v}"/>`; }
    if (s.u) x += '<w:u w:val="single"/>';
    if (s.bg) x += `<w:shd w:val="clear" w:color="auto" w:fill="${s.bg}"/>`;
    if (s.va) x += `<w:vertAlign w:val="${s.va}"/>`;
    return x ? `<w:rPr>${x}</w:rPr>` : '';
  }

  function writeBody(html) {
    const root = new DOMParser().parseFromString(`<!doctype html><body>${html}`, 'text/html').body;
    const ctx = { images: [], links: [], lists: [], docPr: 1 };

    const drawing = (img) => {
      const data = O.splitDataUrl(img.getAttribute('src'));
      if (!data) return '';
      const n = ctx.images.length + 1;
      ctx.images.push({ name: `image${n}.${data.ext}`, b64: data.b64, ext: data.ext, mime: data.mime });
      const w = (parseInt(img.getAttribute('width') || img.style.width, 10) || 400) * EMU_PX;
      const h = (parseInt(img.getAttribute('height') || img.style.height, 10) || 300) * EMU_PX;
      const id = ctx.docPr++;
      return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${w}" cy="${h}"/><wp:docPr id="${id}" name="Imagen ${id}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="Imagen ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg${n}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    };

    const runs = (node, style) => {
      let out = '';
      for (const c of node.childNodes) {
        if (c.nodeType === 3) {
          const text = c.nodeValue.replace(/[\r\n\t ]+/g, ' ').replace(/ /g, '\t');
          if (!text) continue;
          for (const [i, piece] of text.split('\t').entries()) {
            if (i) out += `<w:r>${rPr(style)}<w:tab/></w:r>`;
            if (piece) out += `<w:r>${rPr(style)}<w:t xml:space="preserve">${O.esc(piece)}</w:t></w:r>`;
          }
        } else if (c.nodeType === 1) {
          if (c.tagName === 'BR') out += `<w:r>${rPr(style)}<w:br/></w:r>`;
          else if (c.tagName === 'IMG') out += drawing(c);
          else if (c.tagName === 'A' && /^(https?:|mailto:)/i.test(c.getAttribute('href') || '')) {
            ctx.links.push(c.getAttribute('href'));
            const s = { ...O.inlineStyle(c, style), u: true, color: '0563C1' };
            out += `<w:hyperlink r:id="rIdLink${ctx.links.length}">${runs(c, s)}</w:hyperlink>`;
          } else if (!BLOCK.has(c.tagName)) out += runs(c, O.inlineStyle(c, style));
          else out += runs(c, style);
        }
      }
      return out;
    };

    const para = (el, pStyle, extra = '') => {
      const inner = runs(el, O.inlineStyle(el, {}));
      const st = el.style || {};
      const align = { center: 'center', right: 'right', justify: 'both' }[st.textAlign || ''];
      const lh = parseFloat(st.lineHeight);
      const after = O.toPt(st.marginBottom);
      const spacing = lh || after !== null ? `<w:spacing${after !== null ? ` w:after="${Math.round(after * 20)}"` : ''}${lh ? ` w:line="${Math.round(lh / LINE * 240)}" w:lineRule="auto"` : ''}/>` : '';
      const left = extra ? null : O.toPt(st.marginLeft);
      const ind = left ? `<w:ind w:left="${Math.round(left * 20)}"/>` : '';
      const pr = (pStyle ? `<w:pStyle w:val="${pStyle}"/>` : '') + extra + spacing + ind + (align ? `<w:jc w:val="${align}"/>` : '');
      return `<w:p>${pr ? `<w:pPr>${pr}</w:pPr>` : ''}${inner}</w:p>`;
    };

    const list = (el, depth) => {
      const numId = ctx.lists.length + 1;
      ctx.lists.push(el.tagName === 'OL' ? 'decimal' : 'bullet');
      const walk = (ul, d) => {
        let out = '';
        for (const li of ul.children) {
          if (li.tagName !== 'LI') continue;
          const nested = [...li.children].filter((k) => k.tagName === 'UL' || k.tagName === 'OL');
          nested.forEach((k) => k.remove());
          const lvl = Math.min(d + (+li.getAttribute('data-level') || 0), 2);
          out += para(li, 'ListParagraph', `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${numId}"/></w:numPr>`);
          for (const k of nested) out += walk(k, d + 1);
        }
        return out;
      };
      return walk(el, depth);
    };

    const table = (el) => {
      const rows = [...el.querySelectorAll('tr')].filter((tr) => tr.closest('table') === el);
      const cols = Math.max(1, ...rows.map((tr) => [...tr.children].reduce((n, td) => n + (+td.getAttribute('colspan') || 1), 0)));
      const grid = Math.floor(9000 / cols);
      let x = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid>${`<w:gridCol w:w="${grid}"/>`.repeat(cols)}</w:tblGrid>`;
      for (const tr of rows) {
        x += '<w:tr>';
        let used = 0;
        for (const td of tr.children) {
          const span = +td.getAttribute('colspan') || 1;
          used += span;
          const inner = blocks(td);
          x += `<w:tc><w:tcPr><w:tcW w:w="${grid * span}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}</w:tcPr>${inner.endsWith('</w:p>') ? inner : inner + '<w:p/>'}</w:tc>`;
        }
        for (; used < cols; used++) x += `<w:tc><w:tcPr><w:tcW w:w="${grid}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`;
        x += '</w:tr>';
      }
      return x + '</w:tbl>';
    };

    function blocks(parent) {
      let out = '';
      let loose = null;
      const flush = () => {
        if (!loose) return;
        if (loose.textContent.trim() || loose.querySelector('img,br')) out += para(loose, null);
        loose = null;
      };
      for (const c of [...parent.childNodes]) {
        const isBlock = c.nodeType === 1 && BLOCK.has(c.tagName);
        if (!isBlock) {
          if (c.nodeType === 3 && !c.nodeValue.trim() && !loose) continue;
          if (c.nodeType !== 1 && c.nodeType !== 3) continue;
          loose = loose || root.ownerDocument.createElement('p');
          loose.appendChild(c.cloneNode(true));
          continue;
        }
        flush();
        const tag = c.tagName;
        if (HEADING[tag]) out += para(c, HEADING[tag]);
        else if (tag === 'UL' || tag === 'OL') out += list(c, 0);
        else if (tag === 'TABLE') out += table(c);
        else if (tag === 'HR') out += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
        else if ([...c.children].some((k) => BLOCK.has(k.tagName))) out += blocks(c);
        else out += para(c, null);
      }
      flush();
      return out;
    }

    return { body: blocks(root) || '<w:p/>', ...ctx };
  }

  // Sin estilos guardados se usan los de un documento nuevo de Word 2024.
  // Pagina A4 vertical con los margenes de Word 2024 (2,5 cm arriba y abajo, 3 cm a los lados).
  const DEF_PAGE = { w: 11906, h: 16838, top: 1417, right: 1701, bottom: 1417, left: 1701 };
  const sectPr = (page) => {
    const p = { ...DEF_PAGE, ...(page || {}) };
    const orient = p.w > p.h ? ' w:orient="landscape"' : '';
    return `<w:sectPr><w:pgSz w:w="${p.w}" w:h="${p.h}"${orient}/><w:pgMar w:top="${p.top}" w:right="${p.right}" w:bottom="${p.bottom}" w:left="${p.left}" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;
  };
  O.DOCX_PAGE = DEF_PAGE;

  const DEF_BASE = { font: 'Aptos', size: 11, after: 8, line: 1.08 };
  const DEF_HEADS = {
    H1: { font: 'Aptos Display', size: 20, color: '0F4761', b: false, before: 18, after: 4 },
    H2: { font: 'Aptos Display', size: 16, color: '0F4761', b: false, before: 8, after: 4 },
    H3: { font: 'Aptos', size: 14, color: '0F4761', b: false, before: 8, after: 4 },
  };
  const fontsXml = (f) => { const e = O.esc(f); return `<w:rFonts w:ascii="${e}" w:hAnsi="${e}" w:eastAsia="${e}" w:cs="${e}"/>`; };
  const STYLES = (state) => {
    const base = { ...DEF_BASE, ...(state.base || {}) };
    const heads = { ...DEF_HEADS, ...(state.heads || {}) };
    const sz = (v) => `<w:sz w:val="${Math.round(v * 2)}"/><w:szCs w:val="${Math.round(v * 2)}"/>`;
    const heading = (tag, lvl) => {
      const h = heads[tag];
      return `<w:style w:type="paragraph" w:styleId="Heading${lvl + 1}"><w:name w:val="heading ${lvl + 1}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${Math.round((h.before || 0) * 20)}" w:after="${Math.round((h.after || 0) * 20)}"/><w:outlineLvl w:val="${lvl}"/></w:pPr><w:rPr>${fontsXml(h.font || base.font)}${h.b ? '<w:b/>' : ''}<w:color w:val="${h.color || '000000'}"/>${sz(h.size)}</w:rPr></w:style>`;
    };
    return `${O.XML}<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${fontsXml(base.font)}${sz(base.size)}<w:lang w:val="es-ES"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="${Math.round(base.after * 20)}" w:line="${Math.round(base.line * 240)}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>${heading('H1', 0)}${heading('H2', 1)}${heading('H3', 2)}<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:contextualSpacing/></w:pPr></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>`;
  };

  function numberingXml(lists) {
    let abs = '', nums = '';
    lists.forEach((kind, i) => {
      const id = i + 1;
      let lvls = '';
      for (let l = 0; l < 3; l++) {
        const text = kind === 'bullet' ? ['•', '–', '•'][l] : `%${l + 1}.`;
        lvls += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${kind}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 + l * 360}" w:hanging="360"/></w:pPr></w:lvl>`;
      }
      abs += `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${lvls}</w:abstractNum>`;
      nums += `<w:num w:numId="${id}"><w:abstractNumId w:val="${id}"/></w:num>`;
    });
    return `${O.XML}<w:numbering xmlns:w="${W}">${abs}${nums}</w:numbering>`;
  }

  async function writeDocx(state) {
    const b = writeBody(state.html);
    const zip = new g.JSZip();
    const exts = new Map(b.images.map((i) => [i.ext, i.mime]));
    zip.file('[Content_Types].xml', `${O.XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${[...exts].map(([e, m]) => `<Default Extension="${e}" ContentType="${m}"/>`).join('')}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`);
    zip.file('_rels/.rels', `${O.XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file('word/document.xml', `${O.XML}<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${b.body}${sectPr(state.page)}</w:body></w:document>`);
    zip.file('word/styles.xml', STYLES(state));
    zip.file('word/numbering.xml', numberingXml(b.lists));
    const rels = [
      `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>`,
      `<Relationship Id="rIdNum" Type="${R}/numbering" Target="numbering.xml"/>`,
      ...b.images.map((im, i) => `<Relationship Id="rIdImg${i + 1}" Type="${R}/image" Target="media/${im.name}"/>`),
      ...b.links.map((l, i) => `<Relationship Id="rIdLink${i + 1}" Type="${R}/hyperlink" Target="${O.esc(l)}" TargetMode="External"/>`),
    ];
    zip.file('word/_rels/document.xml.rels', `${O.XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
    for (const im of b.images) zip.file('word/media/' + im.name, im.b64, { base64: true });
    return O.zipOut(zip, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  }

  O.docx = { read: readDocx, write: writeDocx };
})(globalThis);
