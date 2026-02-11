# Slot-X Sales & Inventory Reports

Production-ready Streamlit application for generating professional brand-level Excel reports for:

- Alexandria
- Zamalek
- Merged

## 🚀 Features

- Dynamic column detection
- Refund cleaning (per branch before merge)
- Deals logic per mode (Zamalek / Alexandria / Merged tabs)
- File-per-brand Excel generation
- Structured ZIP output
- Empty Brand Guard detection
- Inventory split in merged mode
- Executive KPI summary
- Status per product

## 📁 Output Structure

Single Mode:
```
/Alexandria/
    /Reports/
    /No_Deal/
    /Empty_Brand_Guard/
```

Merged Mode:
```
/Zamalek/
/Alexandria/
/Merged/
```

Each contains:
```
/Reports/
/No_Deal/
/Empty_Brand_Guard/
```

## 📊 Excel Structure (Per Brand)

1. Sales  
2. Inventory  
3. Report  
4. Metadata  

## 🧠 Deal Logic

- Percentage deducted first
- Rent deducted after percentage
- Mode-specific deal tabs required

## 🛠 Installation (Local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

## ☁ Deployment (Streamlit Cloud)

1. Push to GitHub
2. Go to https://share.streamlit.io
3. Connect repository
4. Set main file = `app.py`
5. Deploy

---

Powered by Slot-X Solutions  
Version 1.0
