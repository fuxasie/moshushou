# PP-OCRv6 model sources

This application bundles the official PaddlePaddle PP-OCRv6 medium ONNX
exports for local CPU inference:

- Detection runtime model:
  `PaddlePaddle/PP-OCRv6_medium_det_onnx`
- Recognition runtime model:
  `PaddlePaddle/PP-OCRv6_medium_rec_onnx`
- Detection safetensors counterpart:
  `PaddlePaddle/PP-OCRv6_medium_det_safetensors`
- Recognition safetensors counterpart:
  `PaddlePaddle/PP-OCRv6_medium_rec_safetensors`

The ONNX and safetensors repositories are official exports of the same
PP-OCRv6 medium model family. The ONNX exports are used because this is a
.NET desktop application and already uses ONNX Runtime.

Model license: Apache License 2.0.

Downloaded on 2026-07-29.

SHA-256:

- `PP-OCRv6_medium_det.onnx`:
  `EB13B44B25BB36F89528B68720AF8A61D9CF381176107F465DB1757B65D086E1`
- `PP-OCRv6_medium_rec.onnx`:
  `9C09ABF0957F7968C7586464B7397B84AD2387A0497A351AF40E9ACC71B673BA`
