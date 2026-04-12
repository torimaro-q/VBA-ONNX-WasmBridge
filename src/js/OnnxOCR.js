async function initModels() {
  const [session, sessionRec, dictText] = await Promise.all([
    ort.InferenceSession.create('det.onnx', { executionProviders: ['wasm'] }),
    ort.InferenceSession.create('rec.onnx', { executionProviders: ['wasm'] }),
    fetch("ppocrv5_dict.txt").then(r => r.text())
  ]);
  const dict = dictText.split('\n').map(s => s.trim()).filter(Boolean).map(s => [...s][0]);
  return { session, sessionRec, dict };
}
async function loadOpenCV() {
  return new Promise(resolve => {
    const script = document.createElement("script");
    script.src = "opencv.js";
    script.onload = () => (cv.onRuntimeInitialized = resolve);
    document.body.appendChild(script);
  });
}

function to32sCeil(x) {
  return Math.ceil(x / 32) * 32;
}

function toCHW(imageData, w, h) {
  const chw = new Float32Array(3 * w * h);
  for (let i = 0, p = 0; i < w * h; i++) {
    const r = imageData[p] / 255;
    const g = imageData[p + 1] / 255;
    const b = imageData[p + 2] / 255;
    p += 4;
    chw[i] = r;
    chw[i + w * h] = g;
    chw[i + 2 * w * h] = b;
  }
  return chw;
}

function cropCanvas(image, box) {
  const canvas = document.createElement('canvas');
  canvas.width = box.w;
  canvas.height = box.h;
  canvas.getContext('2d').drawImage(
    image,
    box.x, box.y, box.w, box.h,
    0, 0, box.w, box.h
  );
  return canvas;
}

function prepareRecInput(canvas) {
  const targetH = 48;
  const ratio = targetH / canvas.height;
  const targetW = Math.round(canvas.width * ratio);
  const tmp = document.createElement('canvas');
  tmp.width = targetW;
  tmp.height = targetH;
  tmp.getContext('2d').drawImage(canvas, 0, 0, targetW, targetH);
  const { data } = tmp.getContext('2d').getImageData(0, 0, targetW, targetH);
  const chw = toCHW(data, targetW, targetH);
  return new ort.Tensor('float32', chw, [1, 3, targetH, targetW]);
}

function decodeCTC(tensor, dict) {
  const data = tensor.data;
  const [n, seq, classes] = tensor.dims;
  let prev = -1;
  const result = [];
  for (let t = 0; t < seq; t++) {
    let maxIdx = 0;
    let maxVal = -Infinity;
    for (let c = 0; c < classes; c++) {
      const v = data[t * classes + c];
      if (v > maxVal) {maxVal = v;maxIdx = c;}
    }
    if (maxIdx !== 0 && maxIdx !== prev && maxIdx < dict.length) {result.push(dict[maxIdx - 2].codePointAt(0));}
    prev = maxIdx;
  }
  return result;
}

class SimpleDBPostProcess {
  constructor(thresh = 0.3, minSize = 3, margin = 5) {this.thresh = thresh;this.minSize = minSize;this.margin = margin;}
  call(predTensor, [src_h, src_w]) {
    const [n, c, h, w] = predTensor.dims;
    const pred = predTensor.data;
    const bitmap = new Uint8Array(h * w);
    for (let i = 0; i < h * w; i++) bitmap[i] = pred[i] > this.thresh ? 1 : 0;
    const mat = cv.matFromArray(h, w, cv.CV_8UC1, bitmap);
    const contours = new cv.MatVector();
    const hierarchy = new cv.Mat();
    cv.findContours(mat, contours, hierarchy, cv.RETR_LIST, cv.CHAIN_APPROX_SIMPLE);
    const results = [];
    for (let i = 0; i < contours.size(); i++) {
      const cnt = contours.get(i);
      const rect = cv.minAreaRect(cnt);
      const pts = cv.RotatedRect.points(rect);
      if (Math.min(rect.size.width, rect.size.height) < this.minSize) {cnt.delete();continue;}
      const scaled = pts.map(p => {
        let x = Math.round(p.x / w * src_w);
        let y = Math.round(p.y / h * src_h);
        return [Math.max(0, Math.min(x, src_w)),Math.max(0, Math.min(y, src_h))];
      });
      const xs = scaled.map(p => p[0]);
      const ys = scaled.map(p => p[1]);
      let x = Math.min(...xs) - this.margin;
      let y = Math.min(...ys) - this.margin;
      x = Math.max(0, x);
      y = Math.max(0, y);
      let wBox = Math.max(...xs) - x + this.margin * 2;
      let hBox = Math.max(...ys) - y + this.margin * 2;
      hBox = Math.min(src_h - y, hBox);
      results.push({ id: results.length, x, y, w: wBox, h: hBox });
      cnt.delete();
    }
    mat.delete();
    contours.delete();
    hierarchy.delete();
    results.sort((a, b) => {
      const ay = Math.floor(a.y / 10);
      const by = Math.floor(b.y / 10);
      if (ay !== by) return ay - by;
      return a.x - b.x;
    });
    return results;
  }
}

async function runInference(image) {
  await loadOpenCV();

  const { session, sessionRec, dict } = await initModels();
  const ow = image.naturalWidth;
  const oh = image.naturalHeight;
  const inputW = to32sCeil(ow);
  const inputH = to32sCeil(oh);

  const canvas = document.createElement('canvas');
  canvas.width = inputW;
  canvas.height = inputH;

  const ctx = canvas.getContext('2d');
  ctx.drawImage(image, 0, 0, ow, oh);

  const { data } = ctx.getImageData(0, 0, inputW, inputH);
  const chw = toCHW(data, inputW, inputH);

  const input = new ort.Tensor('float32', chw, [1, 3, inputH, inputW]);  
  const out = await session.run({ x: input });

  const predTensor = out[Object.keys(out)[0]];
  const post = new SimpleDBPostProcess(0.3);
  const boxes = post.call(predTensor, [inputH, inputW]);
  const recResults = [];

  for (const box of boxes) {
    const crop = cropCanvas(image, box);
    const recInput = prepareRecInput(crop);
    const recOut = await sessionRec.run({ x: recInput });
    const recTensor = recOut[Object.keys(recOut)[0]];
    recResults.push({
      ...box,
      u_label: decodeCTC(recTensor, dict)
    });
  }
  await fetch('/api/result', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(recResults)
  });
}

const img = document.getElementById('img');
if (img.complete && img.naturalWidth > 0) {
  runInference(img);
} else {
  img.onload = () => runInference(img);
}

