const classTxt = await fetch("coco.names").then(r => r.text());
const classLabels = classTxt.split("\n").map(s => s.trim()).filter(s => s.length > 0);

const session = await ort.InferenceSession.create('yolox_nano.onnx', {
  executionProviders: ['wasm']
});

const img = document.getElementById('img');

async function runInference(image) {
  const inputSize = 416;

  const origW = image.naturalWidth;
  const origH = image.naturalHeight;

  const canvas = document.createElement('canvas');
  canvas.width = inputSize;
  canvas.height = inputSize;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(image, 0, 0, inputSize, inputSize);

  const { data } = ctx.getImageData(0, 0, inputSize, inputSize);
  const chw = new Float32Array(3 * inputSize * inputSize);

  for (let i = 0, p = 0; i < inputSize * inputSize; i++) {
    const r = data[p];
    const g = data[p + 1];
    const b = data[p + 2];
    p += 4;

    chw[i] = r;
    chw[i + inputSize * inputSize] = g;
    chw[i + 2 * inputSize * inputSize] = b;
  }

  const input = new ort.Tensor('float32', chw, [1, 3, inputSize, inputSize]);
  const feeds = { images: input };

  const results = await session.run(feeds);
  const output = results[Object.keys(results)[0]].data;

  const detections = decodeYOLOX(output, inputSize, inputSize);
  const nmsDetections = nonMaxSuppression(detections, 0.45, 0.3);

  const scaleX = origW / inputSize;
  const scaleY = origH / inputSize;

  const restored = nmsDetections.map(det => ({
    label: classLabels[det.class] || "Unknown",
    score: det.score,
    class: det.class,
    x: det.x * scaleX,
    y: det.y * scaleY,
    w: det.w * scaleX,
    h: det.h * scaleY
  }));
  await fetch("/api/result", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(restored)
  });
}
function decodeYOLOX(output, imgW, imgH) {
  const strideList = [8, 16, 32];
  const grids = [];
  const expandedStrides = [];
  for (let s of strideList) {
    const ws = imgW / s;
    const hs = imgH / s;
    for (let y = 0; y < hs; y++) {
      for (let x = 0; x < ws; x++) {
        grids.push([x, y]);
        expandedStrides.push(s);
      }
    }
  }
  const results = [];
  const numClasses = 80;
  const numAnchors = grids.length;
  for (let i = 0; i < numAnchors; i++) {
    const offset = i * (5 + numClasses);
    const gx = grids[i][0];
    const gy = grids[i][1];
    const stride = expandedStrides[i];
    const x = (output[offset] + gx) * stride;
    const y = (output[offset + 1] + gy) * stride;
    const w = Math.exp(output[offset + 2]) * stride;
    const h = Math.exp(output[offset + 3]) * stride;
    const obj = output[offset + 4]; // 0?1
    let maxClass = 0;
    let maxProb = 0;
    for (let c = 0; c < numClasses; c++) {
      const p = output[offset + 5 + c]; // 0?1
      if (p > maxProb) {
        maxProb = p;
        maxClass = c;
      }
    }
    const score = obj * maxProb;
    if (score > 0.3) {
      results.push({
        x: x - w / 2,
        y: y - h / 2,
        w,
        h,
        score,
        class: maxClass
      });
    }
  }
  return results;
}
function nonMaxSuppression(boxes, iouThreshold = 0.45, scoreThreshold = 0.3) {
  boxes = boxes.filter(b => b.score >= scoreThreshold);
  boxes.sort((a, b) => b.score - a.score);
  const selected = [];
  while (boxes.length > 0) {
    const best = boxes.shift();
    selected.push(best);
    boxes = boxes.filter(box => {
      const iou = calcIoU(best, box);
      return iou < iouThreshold;
    });
  }
  return selected;
}
function calcIoU(a, b) {
  const x1 = Math.max(a.x, b.x);
  const y1 = Math.max(a.y, b.y);
  const x2 = Math.min(a.x + a.w, b.x + b.w);
  const y2 = Math.min(a.y + a.h, b.y + b.h);
  const inter = Math.max(0, x2 - x1) * Math.max(0, y2 - y1);
  const union = a.w * a.h + b.w * b.h - inter;
  return inter / union;
}
if (img.complete) {
  runInference(img);
} else {
  img.onload = () => runInference(img);
}

