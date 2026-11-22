// Nucleo compartido: utilidades de XML, ZIP, colores y estilos de texto usadas por las tres aplicaciones.
(function (g) {
  const O = (g.Oficina = g.Oficina || {});
  O.apps = O.apps || {};
  O.zipType = O.zipType || 'blob';
  O.XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  O.esc = (s) =>
    String(s)
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');

  O.parseXml = (text) => new DOMParser().parseFromString(text, 'application/xml');
  O.kids = (el, name) => (el ? [...el.children].filter((c) => !name || c.localName === name) : []);
  O.kid = (el, name) => O.kids(el, name)[0] || null;
  O.path = (el, ...names) => names.reduce((cur, n) => (cur ? O.kid(cur, n) : null), el);
  O.all = (el, name) => (el ? [...el.getElementsByTagName('*')].filter((c) => c.localName === name) : []);
  O.attr = (el, name) => {
    if (!el) return null;
    for (const a of el.attributes) if (a.localName === name) return a.value;
    return null;
  };
  O.isOn = (el) => {
    if (!el) return false;
    const v = O.attr(el, 'val');
    return v === null || !['0', 'false', 'off', 'none'].includes(v);
  };

  O.resolvePart = (base, target) => {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/').slice(0, -1);
    for (const p of target.split('/')) {
      if (p === '..') parts.pop();
      else if (p !== '.') parts.push(p);
    }
    return parts.join('/');
  };

  O.readXml = async (zip, part) => {
    const f = zip.file(part);
    return f ? O.parseXml(await f.async('string')) : null;
  };

  O.readRels = async (zip, part) => {
    const i = part.lastIndexOf('/');
    const relPath = part.slice(0, i + 1) + '_rels/' + part.slice(i + 1) + '.rels';
    const xml = await O.readXml(zip, relPath);
    const map = {};
    if (xml) {
      for (const r of O.kids(xml.documentElement, 'Relationship')) {
        const external = O.attr(r, 'TargetMode') === 'External';
        map[O.attr(r, 'Id')] = {
          type: O.attr(r, 'Type'),
          external,
          target: external ? O.attr(r, 'Target') : O.resolvePart(part, O.attr(r, 'Target')),
        };
      }
    }
    return map;
  };

  O.mainPart = async (zip) => {
    const rels = await O.readRels(zip, '');
    const main = Object.values(rels).find((r) => r.type.endsWith('/officeDocument'));
    if (!main) throw new Error('El archivo no contiene un documento de Office valido.');
    return main.target;
  };

  const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };
  O.partToDataUrl = async (zip, part) => {
    const f = zip.file(part);
    if (!f) return null;
    const mime = MIME[part.split('.').pop().toLowerCase()];
    if (!mime) return null;
    return `data:${mime};base64,${await f.async('base64')}`;
  };
  // Separa una imagen data: en extension y contenido base64 para incrustarla en el ZIP.
  O.splitDataUrl = (url) => {
    const m = /^data:image\/([a-z+]+);base64,(.*)$/is.exec(url || '');
    if (!m) return null;
    const ext = m[1].toLowerCase() === 'jpeg' ? 'jpg' : m[1].toLowerCase().replace('+xml', '');
    return { ext, b64: m[2], mime: 'image/' + m[1].toLowerCase() };
  };

  // SHA-512 en enteros de 32 bits, para las contraseñas de hoja de Excel (100 000 vueltas en menos de un segundo).
  const K512 = ('428a2f98d728ae22 7137449123ef65cd b5c0fbcfec4d3b2f e9b5dba58189dbbc 3956c25bf348b538 59f111f1b605d019 923f82a4af194f9b ab1c5ed5da6d8118 '
    + 'd807aa98a3030242 12835b0145706fbe 243185be4ee4b28c 550c7dc3d5ffb4e2 72be5d74f27b896f 80deb1fe3b1696b1 9bdc06a725c71235 c19bf174cf692694 '
    + 'e49b69c19ef14ad2 efbe4786384f25e3 0fc19dc68b8cd5b5 240ca1cc77ac9c65 2de92c6f592b0275 4a7484aa6ea6e483 5cb0a9dcbd41fbd4 76f988da831153b5 '
    + '983e5152ee66dfab a831c66d2db43210 b00327c898fb213f bf597fc7beef0ee4 c6e00bf33da88fc2 d5a79147930aa725 06ca6351e003826f 142929670a0e6e70 '
    + '27b70a8546d22ffc 2e1b21385c26c926 4d2c6dfc5ac42aed 53380d139d95b3df 650a73548baf63de 766a0abb3c77b2a8 81c2c92e47edaee6 92722c851482353b '
    + 'a2bfe8a14cf10364 a81a664bbc423001 c24b8b70d0f89791 c76c51a30654be30 d192e819d6ef5218 d69906245565a910 f40e35855771202a 106aa07032bbd1b8 '
    + '19a4c116b8d2d0c8 1e376c085141ab53 2748774cdf8eeb99 34b0bcb5e19b48a8 391c0cb3c5c95a63 4ed8aa4ae3418acb 5b9cca4f7763e373 682e6ff3d6b2b8a3 '
    + '748f82ee5defb2fc 78a5636f43172f60 84c87814a1f0ab72 8cc702081a6439ec 90befffa23631e28 a4506cebde82bde9 bef9a3f7b2c67915 c67178f2e372532b '
    + 'ca273eceea26619c d186b8c721c0c207 eada7dd6cde0eb1e f57d4f7fee6ed178 06f067aa72176fba 0a637dc5a2c898a6 113f9804bef90dae 1b710b35131c471b '
    + '28db77f523047d84 32caab7b40c72493 3c9ebe0a15c9bebc 431d67c49c100d4c 4cc5d4becb3e42b6 597f299cfc657e2a 5fcb6fab3ad6faec 6c44198c4a475817')
    .split(' ').flatMap((x) => [parseInt(x.slice(0, 8), 16) | 0, parseInt(x.slice(8), 16) | 0]);
  const H512 = '6a09e667f3bcc908 bb67ae8584caa73b 3c6ef372fe94f82b a54ff53a5f1d36f1 510e527fade682d1 9b05688c2b3e6c1f 1f83d9abfb41bd6b 5be0cd19137e2179'
    .split(' ').flatMap((x) => [parseInt(x.slice(0, 8), 16) | 0, parseInt(x.slice(8), 16) | 0]);
  O.sha512 = (bytes) => {
    const n = bytes.length, total = ((n + 17 + 127) >> 7) << 7;
    const buf = new Uint8Array(total);
    buf.set(bytes);
    buf[n] = 0x80;
    const dv = new DataView(buf.buffer);
    dv.setUint32(total - 4, (n * 8) >>> 0);
    dv.setUint32(total - 8, Math.floor((n * 8) / 4294967296));
    const h = H512.slice(), w = new Array(160);
    // Rotación a la derecha de un número de 64 bits guardado como (alto, bajo).
    const rh = (x, y, k) => (k < 32 ? (x >>> k) | (y << (32 - k)) : (y >>> (k - 32)) | (x << (64 - k)));
    const rl = (x, y, k) => (k < 32 ? (y >>> k) | (x << (32 - k)) : (x >>> (k - 32)) | (y << (64 - k)));
    for (let off = 0; off < total; off += 128) {
      for (let i = 0; i < 32; i++) w[i] = dv.getInt32(off + i * 4);
      for (let i = 16; i < 80; i++) {
        const xh = w[(i - 15) * 2], xl = w[(i - 15) * 2 + 1], yh = w[(i - 2) * 2], yl = w[(i - 2) * 2 + 1];
        const s0h = rh(xh, xl, 1) ^ rh(xh, xl, 8) ^ (xh >>> 7), s0l = rl(xh, xl, 1) ^ rl(xh, xl, 8) ^ ((xl >>> 7) | (xh << 25));
        const s1h = rh(yh, yl, 19) ^ rh(yh, yl, 61) ^ (yh >>> 6), s1l = rl(yh, yl, 19) ^ rl(yh, yl, 61) ^ ((yl >>> 6) | (yh << 26));
        const lo = (s0l >>> 0) + (s1l >>> 0) + (w[(i - 7) * 2 + 1] >>> 0) + (w[(i - 16) * 2 + 1] >>> 0);
        w[i * 2] = (s0h + s1h + w[(i - 7) * 2] + w[(i - 16) * 2] + Math.floor(lo / 4294967296)) | 0;
        w[i * 2 + 1] = lo | 0;
      }
      let [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl] = h;
      for (let i = 0; i < 80; i++) {
        const S1h = rh(eh, el, 14) ^ rh(eh, el, 18) ^ rh(eh, el, 41), S1l = rl(eh, el, 14) ^ rl(eh, el, 18) ^ rl(eh, el, 41);
        const chh = (eh & fh) ^ (~eh & gh), chl = (el & fl) ^ (~el & gl);
        const t1l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (K512[i * 2 + 1] >>> 0) + (w[i * 2 + 1] >>> 0);
        const t1h = (hh + S1h + chh + K512[i * 2] + w[i * 2] + Math.floor(t1l / 4294967296)) | 0;
        const S0h = rh(ah, al, 28) ^ rh(ah, al, 34) ^ rh(ah, al, 39), S0l = rl(ah, al, 28) ^ rl(ah, al, 34) ^ rl(ah, al, 39);
        const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch), mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
        const t2l = (S0l >>> 0) + (mjl >>> 0), t2h = (S0h + mjh + Math.floor(t2l / 4294967296)) | 0;
        hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
        const el2 = (dl >>> 0) + (t1l >>> 0);
        eh = (dh + t1h + Math.floor(el2 / 4294967296)) | 0; el = el2 | 0;
        dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
        const al2 = (t1l >>> 0) + (t2l >>> 0);
        ah = (t1h + t2h + Math.floor(al2 / 4294967296)) | 0; al = al2 | 0;
      }
      [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl].forEach((v, i) => {
        if (i % 2) {
          const lo = (h[i] >>> 0) + (v >>> 0);
          h[i - 1] = (h[i - 1] + Math.floor(lo / 4294967296)) | 0;
          h[i] = lo | 0;
        } else h[i] = (h[i] + v) | 0;
      });
    }
    const out = new Uint8Array(64), odv = new DataView(out.buffer);
    h.forEach((v, i) => odv.setInt32(i * 4, v));
    return out;
  };
  const b64 = (bytes) => btoa(String.fromCharCode(...bytes));
  const unb64 = (text) => Uint8Array.from(atob(text), (ch) => ch.charCodeAt(0));
  // Hash de contraseña de hoja o libro, como lo guarda Excel: SHA-512 de sal + contraseña en UTF-16 y spinCount vueltas.
  O.passwordHash = (password, salt, spin) => {
    const pw = new Uint8Array(password.length * 2);
    for (let i = 0; i < password.length; i++) { pw[i * 2] = password.charCodeAt(i) & 255; pw[i * 2 + 1] = password.charCodeAt(i) >> 8; }
    const s = unb64(salt);
    let hash = O.sha512(new Uint8Array([...s, ...pw]));
    const buf = new Uint8Array(68);
    for (let i = 0; i < spin; i++) {
      buf.set(hash);
      buf[64] = i & 255; buf[65] = (i >> 8) & 255; buf[66] = (i >> 16) & 255; buf[67] = (i >>> 24) & 255;
      hash = O.sha512(buf);
    }
    return b64(hash);
  };
  // Contraseña antigua de 16 bits (atributo password).
  O.legacyPasswordHash = (password) => {
    let h = 0;
    [...password].forEach((ch, i) => {
      const v = ch.charCodeAt(0) << (i + 1);
      h ^= (v & 0x7fff) | (v >> 15);
    });
    return ((h ^ password.length ^ 0xce4b) >>> 0).toString(16).toUpperCase();
  };
  // Atributos de protección con contraseña nueva; sin contraseña, solo la marca de protegido.
  O.newProtection = (password) => {
    if (!password) return {};
    const salt = new Uint8Array(16);
    if (g.crypto && g.crypto.getRandomValues) g.crypto.getRandomValues(salt);
    else for (let i = 0; i < 16; i++) salt[i] = Math.floor(Math.random() * 256);
    const saltValue = b64(salt);
    return { algorithmName: 'SHA-512', hashValue: O.passwordHash(password, saltValue, 100000), saltValue, spinCount: '100000' };
  };
  // ¿Abre esta contraseña la protección? Sin contraseña guardada, siempre.
  O.checkProtection = (attrs, password) => {
    if (!attrs) return true;
    if (attrs.hashValue && attrs.saltValue) {
      if (!/SHA-?512/i.test(attrs.algorithmName || '')) return null;
      return O.passwordHash(password, attrs.saltValue, +attrs.spinCount || 0) === attrs.hashValue;
    }
    const legacy = attrs.password || attrs.workbookPassword;
    if (legacy) return O.legacyPasswordHash(password) === legacy.toUpperCase();
    return true;
  };
  O.hasPassword = (attrs) => !!(attrs && (attrs.hashValue || attrs.password || attrs.workbookPassword || attrs.workbookHashValue));

  O.zipOut = (zip, mimeType) => zip.generateAsync({ type: O.zipType, compression: 'DEFLATE', mimeType });

  // Convierte cualquier color CSS habitual en RRGGBB; null si es transparente o desconocido.
  const NAMED = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080' };
  O.hex = (c) => {
    if (!c) return null;
    c = String(c).trim().toLowerCase();
    if (NAMED[c]) return NAMED[c];
    let m = /^#([0-9a-f]{3})$/.exec(c);
    if (m) return [...m[1]].map((x) => x + x).join('').toUpperCase();
    m = /^#([0-9a-f]{6})/.exec(c);
    if (m) return m[1].toUpperCase();
    m = /^rgba?\(\s*(\d+)[ ,]+(\d+)[ ,]+(\d+)(?:[ ,/]+([\d.]+))?/.exec(c);
    if (m) {
      if (m[4] !== undefined && parseFloat(m[4]) === 0) return null;
      return [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase();
    }
    return null;
  };

  const KEYWORD_PT = { 'xx-small': 7.5, 'x-small': 8, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24, 'xxx-large': 36 };
  O.toPt = (size) => {
    if (!size) return null;
    size = String(size).trim().toLowerCase();
    if (KEYWORD_PT[size]) return KEYWORD_PT[size];
    const m = /^([\d.]+)(pt|px)?$/.exec(size);
    if (!m) return null;
    return m[2] === 'px' ? parseFloat(m[1]) * 0.75 : parseFloat(m[1]);
  };

  // Acumula el formato en linea de un nodo HTML sobre el heredado de sus antecesores.
  O.inlineStyle = (el, base) => {
    const s = { ...base };
    const tag = el.tagName;
    if (tag === 'B' || tag === 'STRONG') s.b = true;
    if (tag === 'I' || tag === 'EM') s.i = true;
    if (tag === 'U') s.u = true;
    if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') s.s = true;
    if (tag === 'SUB') s.va = 'subscript';
    if (tag === 'SUP') s.va = 'superscript';
    if (tag === 'FONT') {
      if (el.getAttribute('color')) s.color = O.hex(el.getAttribute('color')) || s.color;
      if (el.getAttribute('face')) s.font = el.getAttribute('face');
    }
    const st = el.style;
    if (st) {
      const w = st.fontWeight;
      if (w === 'bold' || w === 'bolder' || parseInt(w, 10) >= 600) s.b = true;
      if (w === 'normal' || w === '400') s.b = false;
      if (st.fontStyle === 'italic') s.i = true;
      if (st.fontStyle === 'normal') s.i = false;
      const deco = st.textDecorationLine || st.textDecoration || '';
      if (deco.includes('underline')) s.u = true;
      if (deco.includes('line-through')) s.s = true;
      if (st.color) s.color = O.hex(st.color) || s.color;
      const pt = O.toPt(st.fontSize);
      if (pt) s.size = pt;
      if (st.fontFamily) s.font = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (st.backgroundColor) s.bg = O.hex(st.backgroundColor) || s.bg;
      if (st.verticalAlign === 'sub' || st.verticalAlign === 'super') s.va = st.verticalAlign === 'sub' ? 'subscript' : 'superscript';
    }
    return s;
  };

  O.runCss = (r) => {
    const css = [];
    if (r.b) css.push('font-weight:700');
    if (r.i) css.push('font-style:italic');
    const deco = [r.u && 'underline', r.s && 'line-through'].filter(Boolean).join(' ');
    if (deco) css.push('text-decoration:' + deco);
    if (r.color) css.push('color:#' + r.color);
    if (r.size) css.push(`font-size:${r.size}pt`);
    if (r.font) css.push(`font-family:'${String(r.font).replace(/['"<>]/g, '')}'`);
    if (r.bg) css.push('background-color:#' + r.bg);
    if (r.va) css.push('vertical-align:' + (r.va === 'subscript' ? 'sub' : 'super'));
    return css.join(';');
  };
})(globalThis);
