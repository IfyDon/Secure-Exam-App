# Secure Computer-Based Exam Application

A full‑screen, proctored examination system built with Python and Tkinter.  
It enforces strict security rules (no tab switching, no copy/paste, camera monitoring) and automatically saves results to an Excel spreadsheet.

## Features

- **Full‑screen mode** for both registration and exam – prevents distraction.
- **Security monitoring**:
  - Detects tab/window switches – terminates the exam.
  - Detects full‑screen exit – terminates the exam.
  - Disables right‑click, copy/paste, and common keyboard shortcuts (Ctrl+C, Ctrl+V, etc.).
- **Camera overlay** – shows a live feed from the webcam (if available).
- **2‑minute timer** – automatically submits the exam when time runs out.
- **5 Python‑related multiple‑choice questions** (easily customizable).
- **Excel result storage** – saves student details, score, percentage, pass/fail status, date/time, and remarks.
- **Responsive layout** – all UI elements scale with screen size; scrollbars appear if content exceeds the screen height.

## Requirements

- **Python 3.7 or higher** (Tkinter is included by default)
- **Required packages** (install with `pip`):
  - `openpyxl` – for Excel result files
  - `opencv-python` – for camera access
  - `Pillow` – for displaying camera feed in the GUI

> **Note:** If you don't need camera monitoring, you can omit `opencv-python` and `Pillow`. The application will still run (camera features will be disabled).

## Installation

### 1. Clone or download the project

```bash
git clone https://github.com/yourusername/EXAMTEST.git
cd examtest

## Create a virtual environment (recommended)
# Windows
python -m venv venv
venv\Scripts\activate

# macOS / Linux
python3 -m venv venv
source venv/bin/activate

## Install Required Packages
pip install openpyxl opencv-python Pillow

## Run the application
python main.py

Usage

    Registration screen – Enter your full name, class, and registration number.
    Press Esc to exit fullscreen (with confirmation) if needed.

    Exam screen – Answer the 5 Python multiple‑choice questions.

        The timer starts immediately (2 minutes).

        You can navigate using Previous / Next buttons.

        All answers are saved automatically when you move to another question.

        Camera monitoring appears in the bottom‑right corner.

        Security rules are active: any attempt to switch tabs, leave fullscreen, or use forbidden shortcuts will instantly terminate the exam.

    Results – After submission (or timeout), you see your score and whether you passed.
    Results are saved to exam_results.xlsx in the same folder. The Excel file is created automatically if it doesn't exist.


##  Troubleshooting

    Camera not working
    Ensure your webcam is connected and not used by another application.
    On Linux, you may need to install additional packages: sudo apt install python3-opencv.

    Tkinter not found
    On Linux, install python3-tk: sudo apt install python3-tk.

    Full‑screen exit not detected
    The detection relies on window size events. If your window manager interferes, the exam might be terminated incorrectly. In that case, run the application in a standard windowed mode by commenting out the full‑screen lines in the code.

    Error: "No module named 'openpyxl'"
    Install the missing package: pip install openpyxl.


Security Notes

    The application is designed for trusted exam environments. It cannot prevent all possible cheating methods (e.g., external devices), but it enforces basic in‑browser restrictions.

    For stronger proctoring, consider combining with screen recording or AI monitoring.

License

This project is provided for educational purposes. You are free to use, modify, and distribute it as needed.