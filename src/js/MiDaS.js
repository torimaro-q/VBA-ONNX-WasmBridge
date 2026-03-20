const session = await ort.InferenceSession.create('midas_v21_small_256.onnx', {
  executionProviders: ['wasm']
});
const img = document.getElementById('img');
async function runInference(image) {
  const inputSize = 256;
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
    const r = data[p] / 255;
    const g = data[p + 1] / 255;
    const b = data[p + 2] / 255;
    p += 4;
    chw[i] = r;
    chw[i + inputSize * inputSize] = g;
    chw[i + 2 * inputSize * inputSize] = b;
  }
  const input = new ort.Tensor('float32', chw, [1, 3, inputSize, inputSize]);
  const feeds = { input_image: input };
  const results = await session.run(feeds);
  const output = results[Object.keys(results)[0]].data;
  const depthMap = Array.from(output);
  const min = Math.min(...depthMap);
  const max = Math.max(...depthMap);
  const range = max - min;

  const quantized = depthMap.map(v => {
    const norm = (v - min) / range;
    return Math.floor(norm * 255);
  });
  const resized = new Array(origW * origH);
  for (let y = 0; y < origH; y++) {
    for (let x = 0; x < origW; x++) {
      const srcX = Math.floor((x / origW) * inputSize);
      const srcY = Math.floor((y / origH) * inputSize);
      const srcIdx = parseInt(srcY * inputSize + srcX);
      const dstIdx = parseInt(y * origW + x);
      resized[dstIdx] = quantized[srcIdx];
    }
  }
  await fetch("/api/result", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ depth: quantized })
  });
}
if (img.complete) {
  runInference(img);
} else {
  img.onload = () => runInference(img);
}

