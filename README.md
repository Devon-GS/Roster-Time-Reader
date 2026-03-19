# Roster-Time-Reader 🕒📊

A Python-based automation tool designed to calculate fortnightly working hours based on custom company roster logic. It streamlines payroll preparation by calculating hours and exporting formatted data directly into Excel workbooks.

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/)
[![Dependency Manager](https://img.shields.io/badge/dependency-poetry-purple)](https://python-poetry.org/)

---

## ✨ Key Features

- **Automated Calculations:** Automatically determines fortnightly hours worked based on specific roster rules.
- **Excel Integration:** Generates and updates Excel workbooks with roster times and calculated totals.
- **Professional Formatting:** Auto-formats Excel sheets for readability, including headers, bold totals, and cell alignment.
- **Error Handling:** Robust validation to handle blank cells and data inconsistencies without crashing.
- **Production Ready:** Optimized paths for databases, icons, and external assets.

---

## 🚀 Getting Started

### Prerequisites

- Python 3.10 or higher
- [Poetry](https://python-poetry.org/docs/#installation) (Dependency Manager)

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/Roster-Time-Reader.git
   cd Roster-Time-Reader

2. **Set up the virtual environment:**
    ```bash
    python -m venv venv
    ```
3. **Install Poetry and dependencies:**
    ```bash
    pip install poetry
    poetry install
    ```
4. **Running the Application:**
    ```bash
    python main.py
    ```
## 📜 Changelog

All notable changes and bug fixes are documented below.

### [Patch 008] — 2026-03-19
- **Cleaned:** Conducted major code refactor and optimization.
- **Fixed:** Resolved issue where blank cells threw errors during calculation.
- **Added:** New functionality to calculate and append **Total Hours**.
- **Changed:** Enhanced Excel workbook formatting for better visual clarity.

### [Patch 005] — 2026-03-09
- **Fixed:** Resolved a bug preventing the Excel workbook from saving while the application was active.

### [Patch 004] — 2025-11-19
- **Changed:** Configured production-ready paths for Excel files, databases, and application icons.

### [Patch 003] — 2025-11-11
- **Added:** Automated Excel sheet formatting.
- **Changed:** Completed full rewrite of the core logic.

### [Patch 002] — 2025-11-05
- **Added:** Implemented global error handling and data validation.

### [Patch 001] — 2024-03-08
- **Fixed:** Corrected logic errors in the hour calculation algorithm.

---

## 🛠️ Technical Details

- **Language:** Python
- **Environment Management:** Poetry
- **Output Format:** `.xlsx` (Microsoft Excel)
- **Data Handling:** Integrated database support for roster configurations.