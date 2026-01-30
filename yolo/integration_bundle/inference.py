"""
Small inference helper for Flask:
- Loads the YOLO model once (avoids reloading per request).
- Exposes predict(image_bytes) returning a list of detections ready for JSON.
"""
from typing import List, Dict, Any
from ultralytics import YOLO
from PIL import Image
from pathlib import Path
import io
import os

_BASE_DIR = Path(__file__).resolve().parent
_MODEL_PATH = Path(os.environ.get("YOLO_MODEL_PATH", _BASE_DIR / "models" / "best.pt"))

# Load the model once in memory
_MODEL = YOLO(str(_MODEL_PATH))


def predict(image_bytes: bytes, conf: float = 0.25, iou: float = 0.45) -> List[Dict[str, Any]]:
    """
    Takes image bytes, runs the model, and returns detections in this format:
    [
      {"class_id": 0, "class_name": "objeto", "score": 0.87, "box": [x1, y1, x2, y2]},
      ...
    ]
    """
    # Convert bytes to RGB image
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")

    # Run the model with default thresholds
    results = _MODEL.predict(img, conf=conf, iou=iou, verbose=False)
    dets: List[Dict[str, Any]] = []
    if not results:
        return dets

    r = results[0]
    names = r.names  # maps class_id -> name
    boxes = r.boxes

    for b in boxes:
        cls_id = int(b.cls.item())
        score = float(b.conf.item())
        x1, y1, x2, y2 = [float(v) for v in b.xyxy[0].tolist()]
        dets.append({
            "class_id": cls_id,
            "class_name": names.get(cls_id, str(cls_id)),
            "score": round(score, 4),
            "box": [round(x1, 2), round(y1, 2), round(x2, 2), round(y2, 2)],
        })

    return dets
