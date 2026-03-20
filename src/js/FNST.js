const session = await ort.InferenceSession.create('mosaic-8.onnx', {
  executionProviders: ['wasm']
});
const img = document.getElementById('img');
async function runInference(image) {
  const inputSize = 224;
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
  const feeds = { input1: input };
  const results = await session.run(feeds);
  const output = results[Object.keys(results)[0]].data; 

  const rgb = new Uint8Array(3 * inputSize * inputSize);
  for (let i = 0; i < rgb.length; i++) {
    rgb[i] = Math.min(255, Math.max(0, output[i]));
  }
  let outW, outH;
  if (origW === origH) {
    outW = 224;
    outH = 224;
  } else if (origW > origH) {
    outH = 224;
    outW = Math.round(origW * (224 / origH));
  } else {
    outW = 224;
    outH = Math.round(origH * (224 / origW));
  }
  const resized = new Uint8Array(outW * outH * 3);
  for (let y = 0; y < outH; y++) {
    for (let x = 0; x < outW; x++) {
      const srcX = Math.floor((x / outW) * inputSize);
      const srcY = Math.floor((y / outH) * inputSize);
      const srcIdx = srcY * inputSize + srcX;
      const dstIdx = (y * outW + x) * 3;
      resized[dstIdx]     = rgb[srcIdx];
      resized[dstIdx + 1] = rgb[srcIdx + inputSize * inputSize];
      resized[dstIdx + 2] = rgb[srcIdx + 2 * inputSize * inputSize];
    }
  }
  await fetch("/api/result", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({width: outW,height: outH, array: resized.join()  })
  });
}
if (img.complete) {
  runInference(img);
} else {
  img.onload = () => runInference(img);
}

