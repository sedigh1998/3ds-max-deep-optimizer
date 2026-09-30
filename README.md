# 3ds Max Deep Optimizer
Automate batch processing, optimization, and export of 3ds Max projects.
[English](README.md) | [Persian](docs/README.fa.md)
## Overview
A lightweight bilingual desktop tool that recursively scans folders for `.max` files, exports them in a chosen format, applies a `ProOptimizer` modifier, and re-exports the optimized version - all in a single automated run.
## Features
- Recursive scanning of all subfolders
- Batch export in `_Normal` and `_Optimized` versions
- Adjustable optimization percentage (1% to 100%)
- 13 output formats: FBX, OBJ, STL, 3DS, Collada, Alembic, glTF, DXF, DWG, VRML97, IGES, ASE, SAT
- Automatic V-Ray dialog dismissal
- Live progress bar with face-count reporting
- Bilingual UI (Persian / English) with RTL support
- Smart timeout system
- Persistent settings
- Cancellable at any moment
## Requirements
- Windows 10 / 11 (64-bit)
- Python 3.9+
- 3ds Max 2020 - 2025
## Installation
    git clone https://github.com/SEDDIGH/3ds-max-deep-optimizer.git
    cd 3ds-max-deep-optimizer
    pip install -r requirements.txt
    python deep_optimizer.py
## Usage
1. Select a root folder
2. Choose an output format
3. Adjust the optimization slider (1% = max reduction, 100% = no change)
4. Click Start
5. Each file produces `filename_Normal.<ext>` and `filename_Optimized.<ext>`
6. Report saved as `deep_process_report.txt`
## License
MIT License - see LICENSE file.
