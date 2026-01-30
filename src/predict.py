from typing import List, Dict, Any
import io
import os
from pathlib import Path
from PIL import Image

try:
    from ultralytics import YOLO
except ImportError:
    pass

class YOLOInference:
    def __init__(self, model_path: str):
        self.model = YOLO(model_path)

    def predict(self, image_input, conf: float = 0.25, iou: float = 0.45) -> List[Dict[str, Any]]:
        """
        Runs prediction on an image (path or bytes/PIL) and returns list of dicts.
        """
        # If input is bytes, convert to PIL
        if isinstance(image_input, bytes):
            image = Image.open(io.BytesIO(image_input)).convert("RGB")
        else:
            image = image_input

        results = self.model.predict(image, conf=conf, iou=iou, verbose=False)
        dets: List[Dict[str, Any]] = []
        if not results:
            return dets

        r = results[0]
        names = r.names
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

if __name__ == "__main__":
    # Simple CLI test
    import sys
    if len(sys.argv) > 2:
        model_path = sys.argv[1]
        image_path = sys.argv[2]
        predictor = YOLOInference(model_path)
        d = predictor.predict(image_path)
        print(d)
    else:
        print("Usage: python src/predict.py <model_path> <image_path>")
