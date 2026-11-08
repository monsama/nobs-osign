  // ---------------------------------------------------------------------------------------------
  // A PNG with a palette (at most 256 colours, as pngquant or TinyPNG make them): a logo or an icon
  // is often a fraction of its size so, and keeps its transparency, which a JPEG cannot. Made in the
  // browser: the colours counted (and, lossy, the most used ones kept, each pixel the nearest of
  // them), the PNG written by hand (IHDR, PLTE, tRNS, IDAT), compressed with the browser's own
  // CompressionStream (zlib, as PNG wants it).
  // ---------------------------------------------------------------------------------------------
  const CRC_TABLE = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(bytes) { let c = 0xffffffff; for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
  // One chunk: its length, type, data and CRC (over type and data).
  function pngChunk(type, data) {
    const out = new Uint8Array(12 + data.length), dv = new DataView(out.buffer);
    dv.setUint32(0, data.length);
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
    out.set(data, 8);
    dv.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
    return out;
  }
  async function zlib(bytes) {
    const cs = new CompressionStream('deflate');
    const buf = await new Response(new Blob([bytes]).stream().pipeThrough(cs)).arrayBuffer();
    return new Uint8Array(buf);
  }
  // The palette of an image: exact when it has 256 colours or fewer; else (lossy only) the 256
  // most used, the colours first made coarser until few enough are left to choose from. null: not
  // possible this way (more colours and not lossy). max: fewer colours than 256 (smaller still).
  function paletteOf(px, lossy, max) {
    max = max || 256;
    const count = new Map();
    for (let i = 0; i < px.length; i += 4) {
      // Fully transparent pixels are one colour, whatever their RGB says.
      const k = px[i + 3] === 0 ? 0 : ((px[i] << 24) | (px[i + 1] << 16) | (px[i + 2] << 8) | px[i + 3]) >>> 0;
      count.set(k, (count.get(k) || 0) + 1);
      if (count.size > max && !lossy) return null;
    }
    if (count.size <= max) return [...count.keys()];
    // Coarser buckets (fewer bits a channel) until at most 4096 are left, each the mean of its colours.
    for (let drop = 1; drop <= 5; drop++) {
      const mask = (0xff << drop) & 0xff, buckets = new Map();
      for (const [k, n] of count) {
        const r = k >>> 24, g = (k >>> 16) & 0xff, b = (k >>> 8) & 0xff, a = k & 0xff;
        const bk = a === 0 ? 0 : (((r & mask) << 24) | ((g & mask) << 16) | ((b & mask) << 8) | (a & mask)) >>> 0;
        const e = buckets.get(bk) || { n: 0, r: 0, g: 0, b: 0, a: 0 };
        e.n += n; e.r += r * n; e.g += g * n; e.b += b * n; e.a += a * n;
        buckets.set(bk, e);
      }
      if (buckets.size <= 4096 || drop === 5) {
        return [...buckets.values()].sort((x, y) => y.n - x.n).slice(0, max)
          .map((e) => ((Math.round(e.r / e.n) << 24) | (Math.round(e.g / e.n) << 16) | (Math.round(e.b / e.n) << 8) | Math.round(e.a / e.n)) >>> 0);
      }
    }
    return null;
  }
  // The image (ImageData) as a PNG with a palette, or null when it would not be one.
  async function palettePng(img, lossy, max) {
    const px = img.data, pal = paletteOf(px, lossy, max);
    if (!pal || !pal.length) return null;
    const index = new Map(pal.map((c, i) => [c, i]));
    const pr = pal.map((c) => c >>> 24), pg = pal.map((c) => (c >>> 16) & 0xff), pb = pal.map((c) => (c >>> 8) & 0xff), pa = pal.map((c) => c & 0xff);
    // The nearest palette colour of a pixel (counted once per colour).
    const near = (k) => {
      let j = index.get(k); if (j !== undefined) return j;
      const r = k >>> 24, g = (k >>> 16) & 0xff, b = (k >>> 8) & 0xff, a = k & 0xff;
      let best = 0, bd = Infinity;
      for (let i = 0; i < pal.length; i++) { const d = (pr[i] - r) ** 2 + (pg[i] - g) ** 2 + (pb[i] - b) ** 2 + 2 * (pa[i] - a) ** 2; if (d < bd) { bd = d; best = i; } }
      index.set(k, best); return best;
    };
    const w = img.width, h = img.height, raw = new Uint8Array((w + 1) * h);
    for (let y = 0; y < h; y++) {
      raw[y * (w + 1)] = 0; // no filter: a palette's indices do not gain from one
      for (let x = 0; x < w; x++) {
        const i = (y * w + x) * 4, k = px[i + 3] === 0 ? 0 : ((px[i] << 24) | (px[i + 1] << 16) | (px[i + 2] << 8) | px[i + 3]) >>> 0;
        raw[y * (w + 1) + 1 + x] = near(k);
      }
    }
    const ihdr = new Uint8Array(13), dv = new DataView(ihdr.buffer);
    dv.setUint32(0, w); dv.setUint32(4, h); ihdr[8] = 8; ihdr[9] = 3; // 8 bits, palette
    const plte = new Uint8Array(pal.length * 3);
    pal.forEach((c, i) => { plte[i * 3] = pr[i]; plte[i * 3 + 1] = pg[i]; plte[i * 3 + 2] = pb[i]; });
    // Transparency per palette entry, up to the last one that is not opaque.
    let last = -1; pa.forEach((a, i) => { if (a < 255) last = i; });
    const parts = [new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), pngChunk('IHDR', ihdr), pngChunk('PLTE', plte)];
    if (last >= 0) parts.push(pngChunk('tRNS', Uint8Array.from(pa.slice(0, last + 1))));
    parts.push(pngChunk('IDAT', await zlib(raw)), pngChunk('IEND', new Uint8Array(0)));
    return new Blob(parts, { type: 'image/png' });
  }
