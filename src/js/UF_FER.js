// UltraFace input size
const INPUT_W = 640;
const INPUT_H = 480;
const SIZE = 64;
const emotions = ["neutral", "happiness", "surprise", "sadness", "anger", "disgust", "fear", "contempt"];

async function initModels() {
  const [sessionRFB, sessionFER] = await Promise.all([
    ort.InferenceSession.create('version-RFB-640.onnx', { executionProviders: ['wasm'] }),
    ort.InferenceSession.create('emotion-ferplus-12-int8.onnx', { executionProviders: ['wasm'] })
  ]);
  return { sessionRFB, sessionFER};
}

function calcIoU(a, b) {
  const x1 = Math.max(a.x, b.x);
  const y1 = Math.max(a.y, b.y);
  const x2 = Math.min(a.x + a.w, b.x + b.w);
  const y2 = Math.min(a.y + a.h, b.y + b.h);
  const inter = Math.max(0, x2 - x1) * Math.max(0, y2 - y1);
  const union = a.w * a.h + b.w * b.h - inter;
  return union <= 0 ? 0 : inter / union;
}

function softmax(arr) {
  const max = Math.max(...arr);
  const exps = arr.map(v => Math.exp(v - max));
  const sum = exps.reduce((a, b) => a + b, 0);
  return exps.map(v => v / sum);
}

function nonMaxSuppression(boxes, iouThreshold = 0.45, scoreThreshold = 0.3) {
  boxes = boxes.filter(b => b.score >= scoreThreshold);
  boxes.sort((a, b) => b.score - a.score);
  const selected = [];
  while (boxes.length > 0) {
    const best = boxes.shift();
    selected.push(best);
    boxes = boxes.filter(box => calcIoU(best, box) < iouThreshold);
  }
  return selected;
}

function decode(scores, boxes, origW, origH, scoreThres = 0.3) {
  const results = [];
  const predictionCount = scores.length / 2;
  const resizeRatio = Math.min(INPUT_W / origW, INPUT_H / origH);
  const Tp = origH - INPUT_H / resizeRatio;
  const Lp = origW - INPUT_W / resizeRatio;
  for (let i = 0; i < predictionCount; i++) {
    const scoreFace = scores[i * 2 + 1];
    if (scoreFace < scoreThres) continue;
    const bi = i * 4;
    const x1 = 0.5 * Lp + (origW - Lp) * boxes[bi + 0];
    const y1 = 0.5 * Tp + (origH - Tp) * boxes[bi + 1];
    const x2 = 0.5 * Lp + (origW - Lp) * boxes[bi + 2];
    const y2 = 0.5 * Tp + (origH - Tp) * boxes[bi + 3];
    results.push({
      x: x1,
      y: y1,
      w: x2 - x1,
      h: y2 - y1,
      score: scoreFace,
      class: 1
    });
  }
  return results;
}

function cropFace(image, det) {
  const canvas = document.createElement('canvas');
  canvas.width = det.w;
  canvas.height = det.h;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(
    image,
    det.x, det.y, det.w, det.h,
    0, 0, det.w, det.h
  );
  return canvas;
}

function resizeTo64(src) {
  const dst = document.createElement('canvas');
  dst.width = SIZE;
  dst.height = SIZE;
  const ctx = dst.getContext('2d');
  ctx.drawImage(src, 0, 0, SIZE, SIZE);
  return dst;
}

async function runInference(image) {
  const { sessionRFB, sessionFER } = await initModels();
  const origW = image.naturalWidth;
  const origH = image.naturalHeight;
  const canvas = document.createElement('canvas');
  canvas.width = INPUT_W;
  canvas.height = INPUT_H;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(image, 0, 0, INPUT_W, INPUT_H);
  const { data } = ctx.getImageData(0, 0, INPUT_W, INPUT_H);
  const chw = new Float32Array(3 * INPUT_W * INPUT_H);
  for (let i = 0, p = 0; i < INPUT_W * INPUT_H; i++) {
    const r = data[p];
    const g = data[p + 1];
    const b = data[p + 2];
    p += 4;
    chw[i] = (r - 127) / 128;
    chw[i + INPUT_W * INPUT_H] = (g - 127) / 128;
    chw[i + 2 * INPUT_W * INPUT_H] = (b - 127) / 128;
  }
  const input = new ort.Tensor('float32', chw, [1, 3, INPUT_H, INPUT_W]);
  const results = await sessionRFB.run({ input });
  const detections = decode(results["scores"].data, results["boxes"].data, origW, origH);
  const nmsDetections = nonMaxSuppression(detections);

  const restored = [];
  for (const det of nmsDetections) {
    const faceCanvas = cropFace(image, det);
    const resized = resizeTo64(faceCanvas);
    const ctx2 = resized.getContext('2d');
    const { data } = ctx2.getImageData(0, 0, SIZE, SIZE);
    const input = new Float32Array(SIZE * SIZE);
    for (let i = 0, p = 0; i < SIZE * SIZE; i++) {
      const r = data[p];
      const g = data[p + 1];
      const b = data[p + 2];
      p += 4;
      const gray = 0.299 * r + 0.587 * g + 0.114 * b;
      input[i] = gray;
    }
    const tensor = new ort.Tensor('float32', input, [1, 1, SIZE, SIZE]);
    const result = await sessionFER.run({ Input3: tensor });
    const rawScores = result["Plus692_Output_0"].data;
    const maxIndex = rawScores.indexOf(Math.max(...rawScores));
    const scores = softmax(rawScores);

    restored.push({
      label: "face",
      score: det.score,
      x: det.x,
      y: det.y,
      w: det.w,
      h: det.h,
      emotion: emotions[maxIndex],
      eScore: scores[maxIndex],
      neutral: scores[0],
      happiness: scores[1],
      surprise: scores[2],
      sadness: scores[3],
      anger: scores[4],
      disgust: scores[5],
      fear: scores[6],
      contempt: scores[7]
    });
  }
  await fetch("/api/result", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(restored)
  });
}

const img = document.getElementById('img');
if (img.complete) {
    runInference(img);
} else {
    img.onload = () => runInference(img);
}
