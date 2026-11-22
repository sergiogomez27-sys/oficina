// Lectura y escritura de .pptx. Las medidas del modelo estan en puntos (1 pt = 12700 EMU).
(function (g) {
  const O = g.Oficina;
  const EMU = 12700;
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';

  // ---------- Lectura ----------
  function colorOf(fill, theme) {
    if (!fill) return null;
    const srgb = O.kid(fill, 'srgbClr');
    if (srgb) return O.attr(srgb, 'val').toUpperCase();
    const scheme = O.kid(fill, 'schemeClr');
    if (scheme) {
      const alias = { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' };
      const name = O.attr(scheme, 'val');
      return theme[alias[name] || name] || null;
    }
    const sys = O.kid(fill, 'sysClr');
    return sys ? O.attr(sys, 'lastClr') : null;
  }

  const xfrmOf = (spPr) => {
    const x = O.kid(spPr, 'xfrm');
    const off = O.kid(x, 'off'), ext = O.kid(x, 'ext');
    if (!off || !ext) return null;
    return {
      x: +O.attr(off, 'x') / EMU, y: +O.attr(off, 'y') / EMU,
      w: +O.attr(ext, 'cx') / EMU, h: +O.attr(ext, 'cy') / EMU,
    };
  };

  async function readTheme(zip, masterPart) {
    const rels = await O.readRels(zip, masterPart);
    const rel = Object.values(rels).find((r) => r.type.endsWith('/theme'));
    const xml = rel && (await O.readXml(zip, rel.target));
    const theme = {};
    const scheme = xml && O.all(xml, 'clrScheme')[0];
    for (const c of O.kids(scheme)) theme[c.localName] = colorOf(c, {});
    // Fuentes del tema: la de titulos (major) y la de cuerpo (minor).
    const fontsOf = (k) => { const f = xml && O.all(xml, k)[0]; return f ? O.attr(O.kid(f, 'latin'), 'typeface') : null; };
    Object.defineProperty(theme, 'fonts', { value: { major: fontsOf('majorFont'), minor: fontsOf('minorFont') }, enumerable: false });
    return theme;
  }

  const placeholders = (doc) => {
    const out = [];
    for (const sp of O.all(doc, 'sp')) {
      const ph = O.kid(O.path(sp, 'nvSpPr', 'nvPr'), 'ph');
      if (ph) out.push({ type: O.attr(ph, 'type') || 'body', idx: O.attr(ph, 'idx'), sp });
    }
    return out;
  };
  const findPh = (list, ph) => {
    const type = O.attr(ph, 'type') || 'body', idx = O.attr(ph, 'idx');
    const sameFamily = (t) => t === type || (type === 'ctrTitle' && t === 'title') || (type === 'subTitle' && t === 'body');
    return (idx !== null && list.find((p) => p.idx === idx)) || list.find((p) => sameFamily(p.type)) || null;
  };

  async function readPptx(buffer) {
    const zip = await g.JSZip.loadAsync(buffer);
    const main = await O.mainPart(zip);
    const pres = await O.readXml(zip, main);
    const rels = await O.readRels(zip, main);
    const warnings = new Set();
    const size = O.kid(pres.documentElement, 'sldSz');
    const w = +O.attr(size, 'cx') / EMU, h = +O.attr(size, 'cy') / EMU;
    const slides = [];
    let nextId = 1;
    let fonts = null;

    for (const id of O.kids(O.kid(pres.documentElement, 'sldIdLst'), 'sldId')) {
      // sldId tiene dos atributos llamados id: el propio y el r:id de la relacion.
      const part = rels[id.getAttributeNS(R, 'id')].target;
      const doc = await O.readXml(zip, part);
      const slideRels = await O.readRels(zip, part);
      const layoutPart = Object.values(slideRels).find((r) => r.type.endsWith('/slideLayout')).target;
      const layout = await O.readXml(zip, layoutPart);
      const layoutRels = await O.readRels(zip, layoutPart);
      const masterPart = Object.values(layoutRels).find((r) => r.type.endsWith('/slideMaster')).target;
      const master = await O.readXml(zip, masterPart);
      const theme = await readTheme(zip, masterPart);
      if (!fonts && theme.fonts.major) fonts = { major: theme.fonts.major, minor: theme.fonts.minor || theme.fonts.major };
      const layoutPh = placeholders(layout), masterPh = placeholders(master);
      const txStyles = O.kid(master.documentElement, 'txStyles');
      const lvl1 = (name) => O.kid(O.kid(O.kid(txStyles, name), 'lvl1pPr'), 'defRPr');
      const baseSize = (name, fallback) => { const v = O.attr(lvl1(name), 'sz'); return v ? +v / 100 : fallback; };

      const textOf = (sp, def) => {
        const tx = O.kid(sp, 'txBody');
        if (!tx) return '';
        const scale = +O.attr(O.kid(O.kid(tx, 'bodyPr'), 'normAutofit'), 'fontScale') || 100000;
        let html = '';
        for (const p of O.kids(tx, 'p')) {
          const pPr = O.kid(p, 'pPr');
          const algn = { ctr: 'center', r: 'right', just: 'justify' }[O.attr(pPr, 'algn')];
          const lvl = +O.attr(pPr, 'lvl') || 0;
          const bullet = !!O.kid(pPr, 'buChar') || (def.bullets && !O.kid(pPr, 'buNone'));
          let css = algn ? `text-align:${algn};` : '';
          if (lvl) css += `margin-left:${lvl * 24}pt;`;
          let inner = '';
          for (const r of O.kids(p)) {
            if (r.localName === 'br') { inner += '<br>'; continue; }
            if (r.localName !== 'r' && r.localName !== 'fld') continue;
            const rp = O.kid(r, 'rPr');
            const rs = {};
            if (O.attr(rp, 'b') === '1') rs.b = true;
            if (O.attr(rp, 'i') === '1') rs.i = true;
            if (O.attr(rp, 'u') && O.attr(rp, 'u') !== 'none') rs.u = true;
            const sz = O.attr(rp, 'sz');
            if (sz) rs.size = Math.round((+sz / 100) * (scale / 100000) * 10) / 10;
            else if (scale !== 100000 && def.size) rs.size = Math.round(def.size * (scale / 100000) * 10) / 10;
            const color = colorOf(O.kid(rp, 'solidFill'), theme);
            if (color) rs.color = color;
            const face = O.attr(O.kid(rp, 'latin'), 'typeface');
            if (face === '+mj-lt' && theme.fonts.major) rs.font = theme.fonts.major;
            else if (face === '+mn-lt' && theme.fonts.minor) rs.font = theme.fonts.minor;
            else if (face && !face.startsWith('+')) rs.font = face;
            const text = O.esc(O.kid(r, 't') ? O.kid(r, 't').textContent : '');
            const st = O.runCss(rs);
            inner += st ? `<span style="${st}">${text}</span>` : text;
          }
          html += `<div${bullet && inner ? ' class="bu"' : ''}${css ? ` style="${css}"` : ''}>${inner || '<br>'}</div>`;
        }
        return html;
      };

      const walk = async (tree, tf) => {
        const out = [];
        for (const n of O.kids(tree)) {
          const place = (b) => ({ x: b.x * tf.sx + tf.dx, y: b.y * tf.sy + tf.dy, w: b.w * tf.sx, h: b.h * tf.sy });
          if (n.localName === 'sp') {
            const spPr = O.kid(n, 'spPr');
            const ph = O.kid(O.path(n, 'nvSpPr', 'nvPr'), 'ph');
            const type = ph ? O.attr(ph, 'type') || 'body' : null;
            if (['dt', 'ftr', 'sldNum'].includes(type)) continue;
            const inhL = ph ? findPh(layoutPh, ph) : null;
            const inhM = ph ? findPh(masterPh, ph) : null;
            const box = xfrmOf(spPr) || (inhL && xfrmOf(O.kid(inhL.sp, 'spPr'))) || (inhM && xfrmOf(O.kid(inhM.sp, 'spPr')));
            if (!box) continue;
            const isTitle = type === 'title' || type === 'ctrTitle';
            const def = {
              size: ph ? (isTitle ? baseSize('titleStyle', 40) : baseSize('bodyStyle', 24)) : 18,
              bullets: ph && !isTitle && type !== 'subTitle',
            };
            const geom = O.attr(O.kid(spPr, 'prstGeom'), 'prst') || 'rect';
            if (O.kid(spPr, 'custGeom')) warnings.add('formas personalizadas (se muestran como rectángulos)');
            const anchor = O.attr(O.kid(O.kid(n, 'txBody'), 'bodyPr'), 'anchor')
              || O.attr(O.kid(O.kid(inhL && inhL.sp, 'txBody'), 'bodyPr'), 'anchor')
              || O.attr(O.kid(O.kid(inhM && inhM.sp, 'txBody'), 'bodyPr'), 'anchor') || 't';
            const line = O.kid(spPr, 'ln');
            if (O.kid(spPr, 'gradFill') || O.kid(spPr, 'blipFill')) warnings.add('rellenos degradados o de imagen');
            // Lo que la forma no define lo hereda de su estilo del tema; indice 0 significa ninguno.
            const styleColor = (name) => {
              const ref = O.kid(O.kid(n, 'style'), name);
              return ref && O.attr(ref, 'idx') !== '0' ? colorOf(ref, theme) : null;
            };
            out.push({
              id: nextId++, type: 'shape', ...place(box),
              geom: { ellipse: 'ellipse', roundRect: 'roundRect' }[geom] || 'rect',
              fill: O.kid(spPr, 'noFill') ? null : colorOf(O.kid(spPr, 'solidFill'), theme) || styleColor('fillRef'),
              line: line && O.kid(line, 'noFill') ? null : (line && colorOf(O.kid(line, 'solidFill'), theme)) || styleColor('lnRef'),
              size: def.size, anchor, align: type === 'ctrTitle' || type === 'subTitle' ? 'center' : 'left',
              color: styleColor('fontRef') || theme.dk1 || '000000', bold: false,
              html: textOf(n, def),
              ...(type ? { ph: type } : {}),
            });
          } else if (n.localName === 'pic') {
            const rel = slideRels[O.attr(O.all(n, 'blip')[0], 'embed')];
            const url = rel && (await O.partToDataUrl(zip, rel.target));
            const box = xfrmOf(O.kid(n, 'spPr'));
            if (url && box) out.push({ id: nextId++, type: 'image', ...place(box), src: url });
            else warnings.add('imágenes en formato no compatible');
          } else if (n.localName === 'grpSp') {
            const x = O.kid(O.kid(n, 'grpSpPr'), 'xfrm');
            const off = O.kid(x, 'off'), ext = O.kid(x, 'ext'), chOff = O.kid(x, 'chOff'), chExt = O.kid(x, 'chExt');
            if (off && ext && chOff && chExt) {
              const sx = +O.attr(ext, 'cx') / (+O.attr(chExt, 'cx') || 1);
              const sy = +O.attr(ext, 'cy') / (+O.attr(chExt, 'cy') || 1);
              out.push(...(await walk(n, {
                sx: tf.sx * sx, sy: tf.sy * sy,
                dx: tf.dx + tf.sx * (+O.attr(off, 'x') / EMU - (+O.attr(chOff, 'x') / EMU) * sx),
                dy: tf.dy + tf.sy * (+O.attr(off, 'y') / EMU - (+O.attr(chOff, 'y') / EMU) * sy),
              })));
            } else out.push(...(await walk(n, tf)));
          } else if (n.localName === 'graphicFrame') warnings.add('tablas y gráficos');
          else if (n.localName === 'cxnSp') warnings.add('líneas conectoras');
        }
        return out;
      };

      const bgFill = (d) => O.kid(O.kid(O.path(d.documentElement, 'cSld', 'bg'), 'bgPr'), 'solidFill');
      const bg = colorOf(bgFill(doc) || bgFill(layout) || bgFill(master), theme) || 'FFFFFF';
      if (O.kids(O.path(layout.documentElement, 'cSld', 'spTree'), 'pic').length || O.kids(O.path(master.documentElement, 'cSld', 'spTree'), 'pic').length) warnings.add('imágenes del patrón de diapositivas');
      // Transicion: la primera p:transition es la de mc:Choice (con la duracion exacta de PowerPoint 2010 y posteriores).
      let transition;
      const tr = doc.documentElement.getElementsByTagNameNS(P, 'transition')[0];
      if (tr) {
        const fx = [...tr.children].find((c) => !['sndAc', 'extLst'].includes(c.localName));
        const known = ['fade', 'push', 'wipe', 'split', 'cover'];
        let type = fx ? fx.localName : 'fade';
        if (!known.includes(type)) { warnings.add(`la transición «${type}» se muestra como Desvanecer`); type = 'fade'; }
        const durAttr = [...tr.attributes].find((a) => a.localName === 'dur');
        const spd = { fast: 0.5, med: 0.75, slow: 1 }[tr.getAttribute('spd') || 'fast'] || 0.5;
        transition = { type, dur: durAttr ? Math.round(parseInt(durAttr.value, 10) / 10) / 100 : spd };
        if (fx && fx.getAttribute('dir')) transition.dir = fx.getAttribute('dir');
        if (fx && fx.getAttribute('orient')) transition.orient = fx.getAttribute('orient');
      }
      if (doc.documentElement.getElementsByTagNameNS(P, 'timing').length) warnings.add('animaciones');
      const els = await walk(O.path(doc.documentElement, 'cSld', 'spTree'), { sx: 1, sy: 1, dx: 0, dy: 0 });
      slides.push(transition ? { bg, els, transition } : { bg, els });
    }
    if (!slides.length) slides.push({ bg: 'FFFFFF', els: [] });
    return { state: { w, h, slides, nextId, fonts: fonts || { major: 'Calibri Light', minor: 'Calibri' } }, warnings: [...warnings] };
  }

  // ---------- Escritura ----------
  const THEME_XML = `${O.XML}<a:theme xmlns:a="${A}" name="Oficina"><a:themeElements><a:clrScheme name="Oficina"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F3A5F"/></a:dk2><a:lt2><a:srgbClr val="EEF2F5"/></a:lt2><a:accent1><a:srgbClr val="B4532A"/></a:accent1><a:accent2><a:srgbClr val="2E7D5B"/></a:accent2><a:accent3><a:srgbClr val="2B5F8A"/></a:accent3><a:accent4><a:srgbClr val="C9A227"/></a:accent4><a:accent5><a:srgbClr val="6B4E8A"/></a:accent5><a:accent6><a:srgbClr val="4A6670"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Oficina"><a:majorFont><a:latin typeface="@MAJOR@"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="@MINOR@"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Oficina"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

  const emptyTree = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
  const NS = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`;
  const rId = (n) => `rId${n}`;
  const emu = (v) => Math.round(v * EMU);
  const fillXml = (c) => (c ? `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>` : '<a:noFill/>');
  // Transicion como la escribe PowerPoint: duracion exacta para 2010 y posteriores, velocidad aproximada para los antiguos.
  const transitionXml = (t) => {
    if (!t || !['fade', 'push', 'wipe', 'split', 'cover'].includes(t.type)) return '';
    const attrs = [];
    if (t.dir && t.type !== 'fade') attrs.push(`dir="${O.esc(t.dir)}"`);
    if (t.orient && t.type === 'split') attrs.push(`orient="${O.esc(t.orient)}"`);
    const fx = `<p:${t.type}${attrs.length ? ' ' + attrs.join(' ') : ''}/>`;
    const dur = Math.max(1, Math.round((t.dur || 1) * 1000));
    const spd = dur <= 500 ? 'fast' : dur <= 750 ? 'med' : 'slow';
    return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition spd="${spd}" p14:dur="${dur}">${fx}</p:transition></mc:Choice><mc:Fallback><p:transition spd="${spd}">${fx}</p:transition></mc:Fallback></mc:AlternateContent>`;
  };

  function parasXml(el, fonts) {
    const root = new DOMParser().parseFromString(`<!doctype html><body>${el.html || ''}`, 'text/html').body;
    const base = { size: el.size || 18, color: el.color || '000000', b: !!el.bold, i: !!el.italic, font: el.font || null };
    if (!base.font && fonts) base.font = ['title', 'ctrTitle'].includes(el.ph) ? fonts.major : fonts.minor;
    const rPr = (tag, s) => {
      const attrs = `lang="es-ES" sz="${Math.round(s.size * 100)}" b="${s.b ? 1 : 0}" i="${s.i ? 1 : 0}"${s.u ? ' u="sng"' : ''}${s.s ? ' strike="sngStrike"' : ''}`;
      return `<a:${tag} ${attrs}>${s.color ? fillXml(s.color) : ''}${s.font ? `<a:latin typeface="${O.esc(s.font)}"/>` : ''}</a:${tag}>`;
    };
    const runs = (node, style) => {
      let out = '';
      for (const c of node.childNodes) {
        if (c.nodeType === 3) {
          const text = c.nodeValue.replace(/[\r\n\t ]+/g, ' ');
          if (text) out += `<a:r>${rPr('rPr', style)}<a:t>${O.esc(text)}</a:t></a:r>`;
        } else if (c.nodeType === 1) {
          if (c.tagName === 'BR') out += `<a:br>${rPr('rPr', style)}</a:br>`;
          else out += runs(c, O.inlineStyle(c, style));
        }
      }
      return out;
    };
    const para = (node) => {
      const st = O.inlineStyle(node, base);
      let inner = runs(node, st);
      // Un <br> solo al final representa una linea vacia en el editor, no un salto real.
      inner = inner.replace(/<a:br>(?:(?!<a:br>).)*<\/a:br>$/, '');
      const align = { center: 'ctr', right: 'r', justify: 'just' }[(node.style && node.style.textAlign) || el.align];
      const margin = node.style ? parseFloat(node.style.marginLeft) || 0 : 0;
      const bu = node.classList && node.classList.contains('bu');
      const ind = bu ? ` marL="${emu(margin + 18)}" indent="${-emu(18)}"` : margin ? ` marL="${emu(margin)}"` : '';
      const attrs = `${ind}${align ? ` algn="${align}"` : ''}`;
      const buXml = bu ? '<a:buChar char="&#8226;"/>' : '<a:buNone/>';
      return `<a:p><a:pPr${attrs}>${buXml}</a:pPr>${inner}${rPr('endParaRPr', st)}</a:p>`;
    };
    let out = '';
    let loose = null;
    const flush = () => {
      if (loose) out += para(loose);
      loose = null;
    };
    const isBlock = (n) => n.nodeType === 1 && ['DIV', 'P', 'LI', 'H1', 'H2', 'H3'].includes(n.tagName);
    for (const c of [...root.childNodes]) {
      if (isBlock(c)) { flush(); out += para(c); continue; }
      if (c.nodeType === 1 && (c.tagName === 'UL' || c.tagName === 'OL')) {
        flush();
        for (const li of c.children) { li.classList.add('bu'); out += para(li); }
        continue;
      }
      if (c.nodeType === 3 && !c.nodeValue.trim() && !loose) continue;
      loose = loose || root.ownerDocument.createElement('div');
      loose.appendChild(c.cloneNode(true));
    }
    flush();
    return out || `<a:p>${rPr('endParaRPr', base)}</a:p>`;
  }

  async function writePptx(state) {
    const zip = new g.JSZip();
    const images = [];
    const count = state.slides.length;

    state.slides.forEach((slide, si) => {
      const rels = [`<Relationship Id="rId1" Type="${R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`];
      let shapes = '';
      slide.els.forEach((el, ei) => {
        const id = ei + 2;
        const xfrm = `<a:xfrm><a:off x="${emu(el.x)}" y="${emu(el.y)}"/><a:ext cx="${emu(el.w)}" cy="${emu(el.h)}"/></a:xfrm>`;
        if (el.type === 'image') {
          const data = O.splitDataUrl(el.src);
          if (!data) return;
          const name = `image${images.length + 1}.${data.ext}`;
          images.push({ name, ...data });
          const rel = rId(rels.length + 1);
          rels.push(`<Relationship Id="${rel}" Type="${R}/image" Target="../media/${name}"/>`);
          shapes += `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Imagen ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rel}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
        } else {
          const geom = { ellipse: 'ellipse', roundRect: 'roundRect' }[el.geom] || 'rect';
          const line = el.line ? `<a:ln w="12700">${fillXml(el.line)}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
          const anchor = ['ctr', 'b'].includes(el.anchor) ? el.anchor : 't';
          shapes += `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Forma ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm}<a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${fillXml(el.fill)}${line}</p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="${anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>${parasXml(el, { major: 'Aptos Display', minor: 'Aptos', ...(state.fonts || {}) })}</p:txBody></p:sp>`;
        }
      });
      zip.file(`ppt/slides/slide${si + 1}.xml`, `${O.XML}<p:sld ${NS}><p:cSld><p:bg><p:bgPr>${fillXml(slide.bg || 'FFFFFF')}<a:effectLst/></p:bgPr></p:bg><p:spTree>${emptyTree}${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transitionXml(slide.transition)}</p:sld>`);
      zip.file(`ppt/slides/_rels/slide${si + 1}.xml.rels`, `${O.XML}<Relationships xmlns="${PKG}">${rels.join('')}</Relationships>`);
    });

    const slideRels = state.slides.map((s, i) => `<Relationship Id="rId${i + 2}" Type="${R}/slide" Target="slides/slide${i + 1}.xml"/>`).join('');
    zip.file('ppt/presentation.xml', `${O.XML}<p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${state.slides.map((s, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${emu(state.w)}" cy="${emu(state.h)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
    zip.file('ppt/_rels/presentation.xml.rels', `${O.XML}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/slideMaster" Target="slideMasters/slideMaster1.xml"/>${slideRels}<Relationship Id="rId${count + 2}" Type="${R}/theme" Target="theme/theme1.xml"/></Relationships>`);
    zip.file('ppt/slideMasters/slideMaster1.xml', `${O.XML}<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${emptyTree}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`);
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `${O.XML}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${R}/theme" Target="../theme/theme1.xml"/></Relationships>`);
    zip.file('ppt/slideLayouts/slideLayout1.xml', `${O.XML}<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="En blanco"><p:spTree>${emptyTree}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`);
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `${O.XML}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
    const fonts = { major: 'Aptos Display', minor: 'Aptos', ...(state.fonts || {}) };
    zip.file('ppt/theme/theme1.xml', THEME_XML.replace('@MAJOR@', O.esc(fonts.major)).replace('@MINOR@', O.esc(fonts.minor)));
    for (const im of images) zip.file('ppt/media/' + im.name, im.b64, { base64: true });
    zip.file('_rels/.rels', `${O.XML}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
    const exts = new Map(images.map((i) => [i.ext, i.mime]));
    zip.file('[Content_Types].xml', `${O.XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${[...exts].map(([e, m]) => `<Default Extension="${e}" ContentType="${m}"/>`).join('')}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${state.slides.map((s, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
    return O.zipOut(zip, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  }

  O.pptx = { read: readPptx, write: writePptx };
})(globalThis);
