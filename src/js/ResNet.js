const classTxt = await fetch('imagenet_classes.txt').then(r => r.text());
const classLabels = classTxt.split('\n').map(s => s.trim()).filter(s => s.length > 0);

const session = await ort.InferenceSession.create('resnet18_Opset16.onnx', {executionProviders: ['wasm']});
const img = document.getElementById('img');

function softmax(logits) {
  const maxLogit = Math.max(...logits);
  const exps = logits.map(v => Math.exp(v - maxLogit));
  const sumExps = exps.reduce((a, b) => a + b, 0);
  return exps.map(v => v / sumExps);
}

async function runInference(image) {
  const canvas = document.createElement('canvas');
  canvas.width = 224; canvas.height = 224;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(image, 0, 0, 224, 224);
  const { data } = ctx.getImageData(0, 0, 224, 224);

  const mean = [0.485, 0.456, 0.406];
  const std  = [0.229, 0.224, 0.225];
  const chw = new Float32Array(3 * 224 * 224);
  for (let i = 0, p = 0; i < 224*224; i++) {
    const r = data[p]/255, g = data[p+1]/255, b = data[p+2]/255;
    p += 4;
    chw[i] = (r - mean[0]) / std[0];
    chw[i+224*224] = (g - mean[1]) / std[1];
    chw[i+2*224*224] = (b - mean[2]) / std[2];
  }

  const input = new ort.Tensor('float32', chw, [1, 3, 224, 224]);
  const results = await session.run({ x: input });
  const output = results[Object.keys(results)[0]].data;

  const probs = softmax(output);

  const top3 = [...probs]
    .map((p, i) => ({ index: i, prob: p }))
    .sort((a, b) => b.prob - a.prob)
    .slice(0, 3)
    .map(item => ({
      label: classLabels[item.index] || 'Unknown',
      probability: item.prob,
      index: item.index
    }));

  await fetch('/api/result', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(top3)
  });
}

if (img.complete) {
    runInference(img);
} else {
    img.onload = () => runInference(img);
}
