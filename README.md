# SAIL-EGOM-IRON ORE-DESPATCH-DASHBOARD

A **Streamlit-powered interactive dashboard** for tracking and analyzing daily iron ore despatch performance data from SAIL (Steel Authority of India Limited) Eastern Ghats Open Mines (EGOM).

## 🎯 Features

- **Upload & Parse**: Accept Word (.docx) or text exports of daily WhatsApp despatch messages
- **Interactive Dashboard**: Real-time visualization of despatch trends vs. plan
- **Multi-Level Reporting**:
  - Daily, Weekly, Monthly, Quarterly, and Financial Year summaries
  - Mine-wise performance tracking
  - Plant distribution analysis
  - Rake type statistics
- **Performance Metrics**:
  - Achievement % tracking
  - Variance analysis (vs. COD plan)
  - Cumulative progress visualization
- **Data Quality Checks**: Automatic validation of message consistency
- **Export Capabilities**: Download reports as CSV or Excel

## 📊 What It Analyzes

- **Total Rakes**: Daily despatch volumes
- **COD Plan**: Daily planned despatch
- **Mine Performance**: Per-mine despatches and COD plans
- **Plant Distribution**: Routing to Bokaro, Durgapur, Rourkela, ISP, Bhilai
- **Rake Types**: BOXN, BOYN, BOBSN, GPWIS classifications
- **Remarks**: Reasons for shortfalls and operational notes

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/ayushmaanganguly2001/SAIL-EGOM-IRON-ORE-DESPATCH-DASHBOARD.git
cd SAIL-EGOM-IRON-ORE-DESPATCH-DASHBOARD

# Install dependencies
pip install -r requirements.txt
```

### Run Locally

```bash
streamlit run despatch_dashboard.py
```

The app will open in your browser at `http://localhost:8501`

### Usage

1. **Load Messages**: Upload Word files or paste text from WhatsApp despatch messages
2. **Filter**: Select financial year, quarter, months, days, or specific mines
3. **Analyze**: View trends, performance metrics, and detailed breakdowns
4. **Export**: Download data as CSV or Excel for further analysis

## 📋 Reporting Conventions

- **Financial Year**: Runs 1 April to 31 March (e.g., F.Y.2025-2026)
- **Quarters**: Q1 (Apr-Jun), Q2 (Jul-Sep), Q3 (Oct-Dec), Q4 (Jan-Mar)
- **Weeks**: Sunday to Saturday, never cross month boundaries
- **Rakes/Day**: Divides by calendar days (unless otherwise specified)

## 🏭 Mines & Plants

### EGOM Mines (in reporting order):
- **Jharkhand Group**: Manoharpur, Gua, Kiriburu, Meghataburu
- **Odisha Group**: Bolani, Barsuan, Taldih, Kalta

### Plants:
- BSL (Bokaro), DSP (Durgapur), RSP (Rourkela), ISP (Burnpur), BSP (Bhilai)

## 📁 File Structure

```
.
├── despatch_dashboard.py   # Streamlit frontend & UI
├── despatch_report.py      # Message parser & data processing
└── requirements.txt        # Python dependencies
```

## 🛠️ Dependencies

- **Core**: streamlit, pandas
- **Optional**: 
  - `openpyxl` - Excel export support
  - `matplotlib` - Enhanced color gradients for achievement %

Install all with:
```bash
pip install -r requirements.txt
```

## 🔍 Message Format

Expected WhatsApp message structure:

```
D-Month-YYYY

Total Rakes: 20
Daily COD Plan // Current Mth Avg : 22.5 // 21.0
Cumulative: 150 Rakes

KBR(5): (Daily COD Plan- 2.84)
  (1 L+2 F)BSL / 1FRSP

Plant Distribution
  BSL : 2.00L+3.00F: 5 Rakes
  DSP : 1.00L+2.00F: 3 Rakes

Rake Types
  BOXN/BOXNHL(10): 4BSL|2DSP|4RSP

Remarks ::
  Power failure at mine for 2 hours
```

## 🐛 Troubleshooting

**Parser not finding lines?**
- Check the format matches the expected patterns
- See **Unparsed Lines** tab for unrecognized content
- Send sample lines for parser extension

**Excel export not working?**
- Install openpyxl: `pip install openpyxl`

**Need help?**
- Open an issue on GitHub
- Include sample despatch messages and error details

## 📝 License

This project is open source and available for use by SAIL and authorized stakeholders.

## 👤 Author

Created for SAIL-EGOM operational intelligence and reporting.

---

**📌 Note**: This dashboard is maintained for daily operational use. The parser adapts to format variations in WhatsApp messages as they occur.
