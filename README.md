# README.md

# 📊 MRP Critical Items Analyzer

This program analyzes and identifies critical items in a Material Requirements Planning (MRP) process. It connects to a database or reads from an Excel spreadsheet, processes inventory and demand data, and generates a report highlighting items that require urgent attention or replenishment.

**Main features:**
- Calculates available stock and compares it to demand and safety stock.
- Identifies items with insufficient stock to meet requirements.
- Suggests the quantity to order for each critical item.
- Exports the results to a formatted Excel file for easy review and sharing.
- Maintains a historical log of each analysis for traceability.

The program is designed to help companies quickly visualize and act on their most urgent material needs, improving supply chain efficiency and reducing the risk of stockouts.

---

## ⚙️ Installation Tutorial

> Requires Python 3.9 or higher.

1. **Clone the repository:**
   ```bash
   git clone https://github.com/humberto-rodrigz/MRPCriticalItemsAnalyzer
   cd MRPCriticalItemsAnalyzer
   ```

2. **Create a virtual environment (recommended):**
   ```bash
   python -m venv venv
   venv\Scripts\activate  # On Windows
   source venv/bin/activate  # On macOS/Linux
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application:**
   ```bash
   python mrp_gui.py
   ```

---

## 📄 License

MIT License — see [`LICENSE`](LICENSE) for details.

---

## ✨ Author

Maintained by [@humberto-rodrigz](https://github.com/humberto-rodrigz).
