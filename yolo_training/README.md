# Logo Detector with YOLOv11

Complete pipeline to label slides/documents with the corporate logo and train a YOLOv11 detector (Ultralytics).

## Requirements
- Python 3.10+
- GPU optional (CPU training works but is slower)

## Setup
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate
pip install -r requirements.txt
```

## Main flow (Label Studio)

The pipeline now assumes only manual annotations exported from Label Studio in YOLO format. Example (assuming the export is unzipped in `dataset/label_studio/...`):
```bash
python prepare_and_train.py \
  --screenshots_dir screenshots \
  --out_dir dataset \
  --label_studio_export dataset/label_studio/project-1-at-2025-11-03-11-11-e96cd009 \
  --train_ratio 0.8 --val_ratio 0.1 --test_ratio 0.1 \
  --imgs_sz 960 \
  --model_size n \
  --overwrite
```
The script maps the exported `.txt` files (even with hash prefixes/underscores), assumes 1 class (`logo`), and treats images without a label file as negatives. The `dataset/images` and `dataset/labels` directories are recreated, but other subfolders (like the original export) are preserved.

### Useful parameters
- `--model_size`: choose between `n`, `s`, `m` (try bigger if you have GPU headroom).
- `--imgs_sz`: use `1280` if the logo is small or you have more hardware headroom.
- `--label_studio_export`: path to the Label Studio export (required).
- `--overwrite`: clears `dataset/images` and `dataset/labels` before reloading (keeps other files).

## Visual check
```bash
python visual_check.py --dataset_dir dataset --split val --n 20
```
Generates annotated images in `debug_samples/{split}` for quick inspection.

## Best practices
- Keep a set of unseen slides only for the final test.
- Check that all exported `.txt` files were mapped (logs will warn about missing matches).
- Clean `runs_logo/*` periodically to avoid confusion between old and new results; the script uses the real run name in reports.

## Ready for web app integration
- Model weights: `models/best.pt` (copy of the trained `best.pt`).
- Inference module: `inference.py` with `predict(image_bytes)` returning detections as a list of dicts.
- Requirements: `requirements.txt` includes `ultralytics`, `flask`, `lxml`, `pandas`, `openpyxl` (plus training libs).
- Dockerfile: CPU image based on `python:3.10-slim` with system deps for lxml/pandas and default `flask run` command.

### Suggested contract for `/predict` route (Flask)
- `POST /predict` with `multipart/form-data` field `file`.
- Accepted extensions: `.jpg`, `.jpeg`, `.png`.
- 200 response (JSON): list of detections `[{class_id, class_name, score, box}]`.
- Errors: 400 for missing/invalid file; 500 for inference failure.

### Quick test after integration
```bash
curl -X POST -F "file=@samples/sua_imagem.jpg" http://localhost:5000/predict
```
Run outside or inside the container (`docker build -t minha-app:latest .` and `docker run -p 5000:5000 minha-app:latest`).
