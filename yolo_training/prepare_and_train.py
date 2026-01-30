import argparse
import csv
import logging
import math
import random
import shutil
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
import re
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import cv2
import numpy as np
from tqdm import tqdm


SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}
DEFAULT_PROJECT = "runs_logo"
DEFAULT_TRAIN_RUN = "y11_logo"


@dataclass
class ImageAnnotation:
    path: Path
    width: int
    height: int
    classes: List[int]
    boxes: List[Tuple[float, float, float, float]]
    scores: List[Optional[float]]
    scales: List[Optional[float]]
    candidates_before_nms: int
    candidates_after_nms: int


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Load Label Studio annotations and train YOLOv11."
    )
    parser.add_argument("--screenshots_dir", type=Path, required=True)
    parser.add_argument("--out_dir", type=Path, required=True)
    parser.add_argument("--train_ratio", type=float, default=0.8)
    parser.add_argument("--val_ratio", type=float, default=0.1)
    parser.add_argument("--test_ratio", type=float, default=0.1)
    parser.add_argument("--imgs_sz", type=int, default=960)
    parser.add_argument("--model_size", choices=["n", "s", "m"], default="n")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing dataset directory.")
    parser.add_argument(
        "--label_studio_export",
        type=Path,
        required=True,
        help="Path to a Label Studio YOLO export directory. Template matching is disabled.",
    )
    parser.add_argument("--debug", action="store_true", help="Enable debug logging.")
    return parser.parse_args()


def setup_logging(debug: bool) -> None:
    level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=level,
        format="[%(levelname)s] %(message)s",
    )


def validate_args(args: argparse.Namespace) -> None:
    for ratio_name in ("train_ratio", "val_ratio", "test_ratio"):
        ratio_value = getattr(args, ratio_name)
        if ratio_value <= 0:
            raise ValueError(f"{ratio_name} must be > 0 (received {ratio_value}).")
    ratio_sum = args.train_ratio + args.val_ratio + args.test_ratio
    if not math.isclose(ratio_sum, 1.0, rel_tol=1e-3, abs_tol=1e-3):
        raise ValueError(
            f"Train/val/test ratios must sum to 1.0 (received {ratio_sum:.3f})."
        )
    if args.imgs_sz <= 0:
        raise ValueError("imgs_sz must be positive.")


def list_images(directory: Path) -> List[Path]:
    if not directory.exists():
        raise FileNotFoundError(f"Screenshots directory not found: {directory}")
    paths = [p for p in directory.rglob("*") if p.suffix.lower() in SUPPORTED_EXTENSIONS]
    return sorted(paths)


def describe_logo_position(x_center: float, y_center: float) -> str:
    vertical = "top" if y_center < 0.5 else "bottom"
    horizontal = "left" if x_center < 0.5 else "right"
    return f"{vertical} {horizontal}"


def normalize_label_name(name: str) -> str:
    normalized = name.replace("_", " ")
    normalized = re.sub(r"\s+", " ", normalized).strip().lower()
    return normalized


def load_label_studio_annotations(
    export_dir: Path,
    images: Sequence[Path],
) -> Dict[Path, Tuple[List[int], List[Tuple[float, float, float, float]]]]:
    label_dir = export_dir / "labels"
    if not label_dir.exists():
        raise FileNotFoundError(f"Label Studio export missing 'labels' folder: {label_dir}")

    image_lookup: Dict[str, Path] = {}
    for image_path in images:
        key = normalize_label_name(image_path.stem)
        image_lookup[key] = image_path

    annotations: Dict[Path, Tuple[List[int], List[Tuple[float, float, float, float]]]] = {}
    unmatched_labels: List[Path] = []

    for label_file in label_dir.glob("*.txt"):
        stem = label_file.stem
        if "-" in stem:
            stem = stem.split("-", 1)[1]
        key = normalize_label_name(stem)
        image_path = image_lookup.get(key)
        if image_path is None:
            unmatched_labels.append(label_file)
            continue

        classes: List[int] = []
        boxes: List[Tuple[float, float, float, float]] = []
        content = label_file.read_text(encoding="utf-8").strip()
        if content:
            for line in content.splitlines():
                parts = line.strip().split()
                if len(parts) != 5:
                    logging.warning("Skipping malformed line in %s: %s", label_file, line)
                    continue
                try:
                    cls = int(parts[0])
                    x_center, y_center, width, height = map(float, parts[1:])
                except ValueError:
                    logging.warning("Skipping non-numeric line in %s: %s", label_file, line)
                    continue
                classes.append(cls)
                boxes.append((x_center, y_center, width, height))
        annotations[image_path] = (classes, boxes)

    if unmatched_labels:
        logging.warning("Could not match %d label file(s) to screenshots.", len(unmatched_labels))
        for label in unmatched_labels[:5]:
            logging.warning("  Unmatched label: %s", label)

    return annotations


def build_manual_annotation(
    image_path: Path,
    manual_map: Dict[Path, Tuple[List[int], List[Tuple[float, float, float, float]]]],
) -> ImageAnnotation:
    image = cv2.imread(str(image_path), cv2.IMREAD_COLOR)
    if image is None:
        raise ValueError(f"Unable to read image: {image_path}")
    img_h, img_w = image.shape[:2]

    classes, boxes = manual_map.get(image_path, ([], []))
    num = len(boxes)
    scores: List[Optional[float]] = [None] * num
    scales: List[Optional[float]] = [None] * num

    return ImageAnnotation(
        path=image_path,
        width=img_w,
        height=img_h,
        classes=classes,
        boxes=boxes,
        scores=scores,
        scales=scales,
        candidates_before_nms=num,
        candidates_after_nms=num,
    )


def split_dataset(
    images: Sequence[Path],
    annotations: Dict[Path, ImageAnnotation],
    train_ratio: float,
    val_ratio: float,
    test_ratio: float,
    seed: int = 42,
) -> Dict[str, List[Path]]:
    rng = random.Random(seed)
    splits: Dict[str, List[Path]] = {"train": [], "val": [], "test": []}

    positives: List[Path] = []
    negatives: List[Path] = []
    for img in images:
        annotation = annotations.get(img)
        if annotation is None:
            continue
        if annotation.boxes:
            positives.append(img)
        else:
            negatives.append(img)

    def allocate(group: List[Path], label: str) -> None:
        if not group:
            return
        shuffled = list(group)
        rng.shuffle(shuffled)
        total = len(shuffled)
        ratios = {"train": train_ratio, "val": val_ratio, "test": test_ratio}
        raw_counts = {split: total * ratios[split] for split in splits}
        counts = {split: int(math.floor(raw_counts[split])) for split in splits}
        remainder = total - sum(counts.values())
        for split in sorted(splits, key=lambda s: raw_counts[s] - counts[s], reverse=True):
            if remainder <= 0:
                break
            counts[split] += 1
            remainder -= 1

        if label == "positive" and total:
            if counts["train"] == 0:
                donor = max(
                    (split for split in ("val", "test") if counts[split] > 0),
                    key=lambda s: counts[s],
                    default=None,
                )
                if donor is not None:
                    counts["train"] += 1
                    counts[donor] -= 1
                else:
                    counts["train"] = 1
            active = [split for split in ("val", "test") if ratios[split] > 0]
            if len(active) == 2 and total >= 2:
                for split in active:
                    if counts[split] == 0:
                        donor = max(
                            (candidate for candidate in ("train", "val", "test") if counts[candidate] > 1),
                            key=lambda s: counts[s],
                            default=None,
                        )
                        if donor is None:
                            break
                        counts[donor] -= 1
                        counts[split] += 1

        index = 0
        for split in ("train", "val", "test"):
            amount = counts[split]
            if amount <= 0:
                continue
            splits[split].extend(shuffled[index : index + amount])
            index += amount

    allocate(positives, "positive")
    allocate(negatives, "negative")
    return splits


def prepare_output_directories(out_dir: Path, overwrite: bool) -> None:
    if out_dir.exists():
        if any(out_dir.iterdir()) and not overwrite:
            raise FileExistsError(
                f"{out_dir} already exists and is not empty. Use --overwrite to replace it."
            )
        if overwrite:
            for subdir in ("images", "labels"):
                target = out_dir / subdir
                if target.exists():
                    shutil.rmtree(target)
            matches_file = out_dir / "matches.csv"
            if matches_file.exists():
                matches_file.unlink()
    else:
        out_dir.mkdir(parents=True, exist_ok=True)
    images_dir = out_dir / "images"
    labels_dir = out_dir / "labels"
    for split in ("train", "val", "test"):
        (images_dir / split).mkdir(parents=True, exist_ok=True)
        (labels_dir / split).mkdir(parents=True, exist_ok=True)


def write_label_file(
    label_path: Path,
    boxes: Iterable[Tuple[float, float, float, float]],
    classes: Iterable[int],
) -> None:
    lines = [
        f"{cls} {x_center:.6f} {y_center:.6f} {width:.6f} {height:.6f}"
        for cls, (x_center, y_center, width, height) in zip(classes, boxes)
    ]
    label_path.write_text("\n".join(lines), encoding="utf-8")


def copy_images_and_labels(
    annotations: Dict[Path, ImageAnnotation],
    splits: Dict[str, List[Path]],
    out_dir: Path,
) -> Dict[Path, str]:
    assignment: Dict[Path, str] = {}
    for split, paths in splits.items():
        image_dir = out_dir / "images" / split
        label_dir = out_dir / "labels" / split
        for src in paths:
            annotation = annotations.get(src)
            if annotation is None:
                continue
            dest_image = image_dir / src.name
            shutil.copy2(src, dest_image)
            label_path = label_dir / (src.stem + ".txt")
            write_label_file(label_path, annotation.boxes, annotation.classes)
            assignment[src] = split
    return assignment


def save_logo_yaml(out_dir: Path) -> Path:
    if out_dir.is_absolute():
        try:
            rel = out_dir.relative_to(Path.cwd())
            dataset_path = f"./{rel.as_posix()}"
        except ValueError:
            dataset_path = str(out_dir.resolve())
    else:
        dataset_path = f"./{out_dir.as_posix()}"
        if not dataset_path.startswith("./"):
            dataset_path = f"./{dataset_path}"

    yaml_lines = [
        f"path: {dataset_path}",
        "train: images/train",
        "val: images/val",
        "test: images/test",
        "names:",
        "  0: logo",
        "",
    ]
    yaml_path = Path("logo.yaml")
    yaml_path.write_text("\n".join(yaml_lines), encoding="utf-8")
    return yaml_path


def save_matches_csv(
    annotations: Dict[Path, ImageAnnotation],
    assignment: Dict[Path, str],
    csv_path: Path,
) -> None:
    header = [
        "image",
        "split",
        "class",
        "score",
        "x_center",
        "y_center",
        "width",
        "height",
        "scale",
    ]
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        for path, annotation in annotations.items():
            split = assignment.get(path, "unassigned")
            if not annotation.boxes:
                writer.writerow([str(path), split, "", "", "", "", "", "", ""])
                continue
            scores = annotation.scores if annotation.scores else [None] * len(annotation.boxes)
            scales = annotation.scales if annotation.scales else [None] * len(annotation.boxes)
            for cls, (x_center, y_center, width, height), score, scale in zip(
                annotation.classes,
                annotation.boxes,
                scores,
                scales,
            ):
                score_str = "" if score is None else f"{score:.4f}"
                scale_str = "" if scale is None else f"{scale:.4f}"
                writer.writerow(
                    [
                        str(path),
                        split,
                        cls,
                        score_str,
                        f"{x_center:.6f}",
                        f"{y_center:.6f}",
                        f"{width:.6f}",
                        f"{height:.6f}",
                        scale_str,
                    ]
                )


def collect_statistics(annotations: Dict[Path, ImageAnnotation]) -> Dict[str, object]:
    total_images = len(annotations)
    logos_per_image = Counter(len(ann.boxes) for ann in annotations.values())
    detected_images = sum(freq for count, freq in logos_per_image.items() if count > 0)
    ratios = {
        "with_logo_pct": (detected_images / total_images * 100) if total_images else 0.0,
        "without_logo_pct": ((total_images - detected_images) / total_images * 100) if total_images else 0.0,
    }
    widths = [box[2] for ann in annotations.values() for box in ann.boxes]
    heights = [box[3] for ann in annotations.values() for box in ann.boxes]
    candidate_counts = sum(ann.candidates_before_nms for ann in annotations.values())
    retained_counts = sum(ann.candidates_after_nms for ann in annotations.values())
    filtered_counts = candidate_counts - retained_counts

    stats = {
        "total_images": total_images,
        "logos_per_image": logos_per_image,
        "ratios": ratios,
        "widths": widths,
        "heights": heights,
        "candidate_counts": candidate_counts,
        "retained_counts": retained_counts,
        "filtered_counts": filtered_counts,
    }
    return stats


def summarize_sizes(values: Sequence[float], label: str) -> None:
    if not values:
        logging.info("%s: no annotations", label)
        return
    arr = np.array(values, dtype=np.float32)
    logging.info(
        "%s (normalized): min=%.4f mean=%.4f median=%.4f max=%.4f",
        label,
        float(np.min(arr)),
        float(np.mean(arr)),
        float(np.median(arr)),
        float(np.max(arr)),
    )


def print_report(stats: Dict[str, object]) -> None:
    logging.info("Total images processed: %d", stats["total_images"])
    logging.info(
        "Images with logo: %.1f%% | without logo: %.1f%%",
        stats["ratios"]["with_logo_pct"],
        stats["ratios"]["without_logo_pct"],
    )
    logging.info("Distribution of logos per image:")
    for count, freq in sorted(stats["logos_per_image"].items()):
        logging.info("  %d logos -> %d image(s)", count, freq)
    summarize_sizes(stats["widths"], "Bounding box width")
    summarize_sizes(stats["heights"], "Bounding box height")
    logging.info(
        "Candidates before NMS: %d | retained after NMS: %d | filtered out: %d",
        stats["candidate_counts"],
        stats["retained_counts"],
        stats["filtered_counts"],
    )


def train_yolo(
    data_yaml: Path,
    imgsz: int,
    model_size: str,
    dataset_dir: Path,
) -> Dict[str, float]:
    try:
        from ultralytics import YOLO  # type: ignore
    except ImportError as exc:
        logging.error("Ultralytics is not installed: %s", exc)
        return {}

    weights = Path(f"yolo11{model_size}.pt")
    if not weights.exists():
        logging.warning(
            "Weights file %s not found locally. Ultralytics will attempt to download it.",
            weights,
        )

    logging.info("Starting YOLOv11 training with %s...", weights)
    model = YOLO(str(weights))
    train_results = model.train(
        data=str(data_yaml),
        imgsz=imgsz,
        epochs=100,
        batch=16,
        project=DEFAULT_PROJECT,
        name=DEFAULT_TRAIN_RUN,
    )
    train_save_dir = Path(getattr(train_results, "save_dir", Path(DEFAULT_PROJECT) / DEFAULT_TRAIN_RUN))
    logging.info("Training outputs saved to %s", train_save_dir)
    val_name = f"{train_save_dir.name}_val"
    pred_name = f"{train_save_dir.name}_pred"

    logging.info("Running validation...")
    model.val(
        data=str(data_yaml),
        imgsz=imgsz,
        project=str(train_save_dir.parent),
        name=val_name,
    )
    logging.info("Running predictions on test split...")
    model.predict(
        source=str(dataset_dir / "images" / "test"),
        imgsz=imgsz,
        conf=0.25,
        project=str(train_save_dir.parent),
        name=pred_name,
        save=True,
    )
    results_csv = train_save_dir / "results.csv"
    metrics = read_metrics_from_results(results_csv)
    if metrics:
        logging.info("Training metrics (final epoch):")
        for key, value in metrics.items():
            logging.info("  %s = %.4f", key, value)
    else:
        logging.info(
            "Training metrics not available yet. Check %s after training finishes.",
            results_csv,
        )
    return metrics


def read_metrics_from_results(results_csv: Path) -> Dict[str, float]:
    if not results_csv.exists():
        return {}
    last_row: Optional[Dict[str, float]] = None
    with results_csv.open("r", newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            last_row = {
                key: float(value)
                for key, value in row.items()
                if value not in ("", None) and is_float(value)
            }
    return last_row or {}


def is_float(value: str) -> bool:
    try:
        float(value)
    except ValueError:
        return False
    return True


def main() -> None:
    args = parse_args()
    setup_logging(args.debug)
    random.seed(42)
    np.random.seed(42)

    validate_args(args)

    images = list_images(args.screenshots_dir)
    logging.info("Found %d image(s) in %s", len(images), args.screenshots_dir)
    if not images:
        raise RuntimeError("No images found for processing.")

    logging.info("Using manual annotations from %s", args.label_studio_export)
    manual_map = load_label_studio_annotations(args.label_studio_export, images)
    positives = sum(1 for classes, _ in manual_map.values() if classes)
    logging.info(
        "Matched %d image(s) with manual labels (positives: %d).",
        len(manual_map),
        positives,
    )
    logging.info(
        "Images without explicit manual labels: %d (treated as background).",
        len(images) - len(manual_map),
    )
    tqdm_desc = "Loading labels"

    annotations: Dict[Path, ImageAnnotation] = {}
    failed_images: List[Tuple[Path, str]] = []

    tqdm_bar = tqdm(images, desc=tqdm_desc, unit="image")
    for image_path in tqdm_bar:
        try:
            annotation = build_manual_annotation(image_path, manual_map)
            tqdm_bar.set_postfix({"labels": len(annotation.boxes)})
            annotations[image_path] = annotation
        except Exception as exc:  # pylint: disable=broad-except
            failed_images.append((image_path, str(exc)))
            logging.warning("Failed to process %s: %s", image_path, exc)

    if failed_images:
        logging.error("Encountered %d problematic image(s). They will be skipped.", len(failed_images))

    successful_images = list(annotations.keys())
    if not successful_images:
        raise RuntimeError("No images were successfully processed. Aborting.")

    prepare_output_directories(args.out_dir, args.overwrite)
    splits = split_dataset(
        successful_images,
        annotations,
        args.train_ratio,
        args.val_ratio,
        args.test_ratio,
    )
    for split_name, paths in splits.items():
        logging.info("%s split: %d image(s)", split_name.capitalize(), len(paths))
    assignment = copy_images_and_labels(annotations, splits, args.out_dir)
    csv_path = args.out_dir / "matches.csv"
    save_matches_csv(annotations, assignment, csv_path)
    logging.info("Saved match details to %s", csv_path)

    stats = collect_statistics(annotations)
    print_report(stats)

    data_yaml = save_logo_yaml(args.out_dir)
    logging.info("Generated %s", data_yaml)

    train_yolo(data_yaml, args.imgs_sz, args.model_size, args.out_dir)


if __name__ == "__main__":
    main()
