const session = await ort.InferenceSession.create('Shinkai_53.onnx', {
  executionProviders: ['wasm']
});
const img = document.getElementById('img');
const MAX_SIZE = 1024;

function to32sCeil(x) {
  return Math.ceil(x / 32) * 32;
}

async function runInference(image) {
  let w = image.naturalWidth;
  let h = image.naturalHeight;

  if (w > MAX_SIZE || h > MAX_SIZE) {
    const scale = MAX_SIZE / Math.max(w, h);
    w = Math.floor(w * scale);
    h = Math.floor(h * scale);
  }

  const inputW = to32sCeil(w);
  const inputH = to32sCeil(h);

  const canvas = document.createElement('canvas');
  canvas.width = inputW;
  canvas.height = inputH;
  const ctx = canvas.getContext('2d');

  ctx.drawImage(image, 0, 0, w, h);

  const { data } = ctx.getImageData(0, 0, inputW, inputH);
  const nhwc = new Float32Array(inputW * inputH * 3);

  for (let i = 0, p = 0; i < inputW * inputH; i++) {
    const r = data[p] / 127.5 - 1;
    const g = data[p + 1] / 127.5 - 1;
    const b = data[p + 2] / 127.5 - 1;
    p += 4;

    const idx = i * 3;
    nhwc[idx]     = r;
    nhwc[idx + 1] = g;
    nhwc[idx + 2] = b;
  }

  const input = new ort.Tensor('float32', nhwc, [1, inputH, inputW, 3]);
  const feeds = { "generator_input:0": input };

  const results = await session.run(feeds);
  const output = results["generator/G_MODEL/out_layer/Tanh:0"].data;

  const rgb = new Uint8Array(inputW * inputH * 3);
  for (let i = 0; i < rgb.length; i++) {
    const v = (output[i] + 1) * 127.5;
    rgb[i] = Math.min(255, Math.max(0, v));
  }
  await fetch("/api/result", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      width: inputW,
      height: inputH,
      array: rgb.join()
    })
  });
}

if (img.complete) {
  runInference(img);
} else {
  img.onload = () => runInference(img);
}
