# README.md (English version)

# 📊 MRP Critical Items Analyzer

A desktop application built with Python and a graphical interface to analyze and visualize critical items from an MRP (Material Requirements Planning) spreadsheet. Offers interactive filtering, charting, exporting, and quick statistics — all in a clean and modern interface.

---

## 🖥️ Features

- Load and analyze Excel `.xlsx` spreadsheets
- Interactive bar chart showing quantities to request
- Table view with:
  - Column filtering
  - Clickable column headers for sorting
  - Pagination (50 items per page)
  - Summary statistics
  - Export to `.csv` or `.xlsx`
  - Double-click to view detailed item information
- Light/dark theme toggle (`Ctrl+T`)

---

## ⚙️ Installation (Developer Mode)

> Requires Python 3.9 or higher.

1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-user/your-repo.git
   cd your-repo
   ```

2. **Create a virtual environment (recommended)**:
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   source venv/bin/activate  # macOS/Linux
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the app**:
   ```bash
   python mrp_gui.py
   ```

---

## 📦 Building the `.exe` Executable (Windows)

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Build the executable:
   ```bash
   pyinstaller --noconfirm --onefile --windowed mrp_gui.py
   ```

3. The executable will be available in:
   ```
   dist/mrp_gui.exe
   ```

---

## 📁 Expected Spreadsheet Format

The worksheet must contain the following columns (names must match exactly):

- `CÓD`
- `DESCRIÇÃOPROMOB`
- `ESTOQ10`
- `ESTOQ20`
- `DEMANDAMRP`
- `ESTOQSEG`
- `STATUS`
- `FORNECEDORPRINCIPAL`
- `PEDIDOS`
- `OBS`

These can be in any order, as long as the column headers match.

---

## 🧪 How to Use

1. Open the application.
2. Select the Excel spreadsheet.
3. Enter the worksheet name (e.g., `Cálculo MRP`).
4. Click **Analyze MRP**.
5. Review the results via the **Chart** and **Table** tabs.

---

## 🧱 Tech Stack

- **Python 3**
- **Tkinter + ttkbootstrap** (GUI)
- **Pandas + XlsxWriter** (data processing)
- **Matplotlib** (chart rendering)
- **PyInstaller** (packaging into executable)

---

## 📄 License

MIT License — see [`LICENSE`](LICENSE) for details.

---

## ✨ Author

Developed by **Humberto Rodrigues** — Fullstack Developer focused on frontend solutions.
