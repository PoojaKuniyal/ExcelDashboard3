## 📊 Sales Dashboard with Conditional Slicer Connectivity (Excel + VBA)

### Overview
This interactive Excel dashboard showcases sales performance across multiple regions using pivot tables, slicers, and custom VBA logic. Designed to highlight both top-performing and underperforming sales executives, it helps businesses quickly assess target achievement across different cities.

The dashboard dynamically connects slicers to specific pivot tables based on user selection, enabling focused regional analysis while maintaining stability in other views. It’s ideal for organizations tracking KPIs and regional sales effectiveness.

---

### 🔧 Features

- ✅ Dynamic slicer linkage via VBA (based on dashboard selection)
- 📍 Region slicer with values: *Mumbai, Patna, Delhi, Chennai, Nagpur, Pune, Ranchi, Surat*
- 🏆 Top 5 sales executives by total sales
- 📉 Bottom 5 sales executives by total sales
- 🎯 Target achievement visualization
  - % of top 5 executives who hit targets
  - % of bottom 5 executives who missed targets or underperformed
- 📈 Individual pivot charts for each performance group

---

### 🧠 How It Works

1. **Macro Recording:**  
   - Start by recording a macro while manually linking/unlinking a slicer to a pivot table.
   - Ensure no slicer options are selected before recording begins.

2. **VBA Adjustment:**  
   - After recording, inspect the generated VBA module.
   - Customize logic to connect each slicer only to its respective dashboard pivot table using conditional statements.
   - Use cell-based control (e.g., `Sheet1.Range("A1").Value`) to toggle dashboard focus.

3. **Save as Excel Macro-Enabled File (`.xlsm`)**

---

### 🚀 Getting Started

1. Clone the repo or download `Sales_Dashboard.xlsm`
2. Open the file in Excel (enable macros)
3. Use the slicer and control cells to explore different dashboards

---

🎥 Demo Preview
⚠️ Note: Excel macro-enabled files (.xlsm) cannot be executed or previewed directly on GitHub.
To demonstrate the dashboard's interactivity, slicer behavior, and conditional chart logic, here’s a short video walkthrough:

📽️ Watch the demo → https://vimeo.com/1103277057
