 # Lameco Format Correction Application

This Open Source project is a specialized **format correction tool** designed for **Lameco**. It automates the checking and standardization of **PowerPoint (.pptx)** and **Word (.docx)** documents based on brand guidelines.

The core of the system is a rule-based engine that checks for fonts, sizes, and layout compliance, combined with an **AI-powered YOLO model** used specifically for detecting and validating the Lameco logo in documents.

## Project Overview

The main objective is to provide a user-friendly application where checking formatted documents is automated.

### Key Features
1.  **Rule-Based Logic Engine**:
    *   **PowerPoint (.pptx)**: Validates slide size, font families (e.g., Montserrat), font sizes (titles, body), and specific slide layouts (Agenda, Closing).
    *   **Word (.docx)**: Checks page margins, font usage, headings, and line spacing against corporate identity.
    *   **Configurable Rules**: settings are persisted in `ui_settings.json` and can be adjusted via the UI.
2.  **Logo Detection (AI)**:
    *   Uses a **YOLOv11** object detection model to verify the presence and correct placement of the Lameco logo in images/slides.
3.  **MLOps Pipeline**:
    *   An automated CI/CD pipeline ensures the YOLO model is reproducible and valid. It handles training, testing, and artifact generation automatically upon code changes.

## Project Structure

```
lameco-open-source
├── app_lameco_updated.py  # MAIN APPLICATION: The Flask web app running rule-based checks & UI
├── ui_settings.json       # Configuration file for rule-based logic (fonts, sizes, margins)
├── yolo_implementation_final_version.py 
├── src
│   ├── train.py           # MLOps: Script to retrain the YOLO logo model
│   ├── predict.py         # MLOps: Helper for model inference
│   └── data               # helper scripts for data processing
├── yolo_training          # Original dataset and training resources
├── params.yaml            # DVC/MLOps parameters configuration
├── .github/workflows      # CI/CD pipelines
└── requirements.txt       # Dependencies for both App and AI
```

## How to Use

### 1. Running the Format Correction App
The main application is a web-based tool powered by Flask.

**Prerequisites**:
*   Python 3.10+ installed.
*   Dependencies installed: `pip install -r requirements.txt`

**Start the App**:
```bash
python app_lameco_updated.py
```
*   Open your browser at `http://localhost:5000` (or the port shown in terminal).
*   Upload a `.pptx` or `.docx` file to see the compliance report.

### 2. MLOps & Model Training
The MLOps component is "added value" to ensure the AI part (logo detection) remains robust.

*   **Pipeline**: defined in `.github/workflows/mlops.yml`. It runs automatically on GitHub.
*   **Manual Training**: If you need to retrain the logo detector manually:
    ```bash
    python src/train.py \
      --screenshots_dir "yolo_training/dataset_v2/images" \
      --label_studio_export "yolo_training/dataset_v2" \
      --out_dir "data/processed/dataset_prepared" \
      --imgs_sz 640 \
      --model_size n \
      --overwrite
    ```

## Development Logic

*   **Rule-Based Logic**: Located in `app_lameco_updated.py`. Look for `check_pptx_rules` and `check_docx_rules` functions to understand how fonts and margins are validated.
*   **AI Integration**: The app uses `ultralytics` to load the trained model (`models/best.pt`) and inference scripts to check images embedded within the documents.
