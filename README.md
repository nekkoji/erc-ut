# ERC PDF Utility Tool

A multifunctional PyQt5-based desktop application to streamline and automate PDF-related document processing tasks for the Energy Regulatory Commission.

---

## 🚀 Features

### 📂 File Management
- **Extract & Rename of OBR, NCA, and SARO PDFs** based on content (e.g., Serial No., NCA No., SARO No.)
- **Split PDFs** into individual pages
- **Manage PDFs**: Open, rename, delete, and move files

### 🧾 OBR Extractor
- OCR-based data extraction from PDF forms
- Intelligent parsing of Payee, Date, Particulars, Amount
- Table editing with:
  - Cell-level editing
  - Undo/Redo
  - Copy/Paste
  - Inline audit logs
- Manual OCR region scan per cell (crop & extract)
- Save to CSV, Excel, and PDF
- Smart summary row calculation

### 💸 Earmark Monitoring
- Track budget allotment, realignment, obligations, and earmarked funds
- Auto-calculated balances and utilization percentages
- Live updates in Excel with formulas

### 🏢 SharePoint Integration
- Authenticate to a SharePoint site
- Extract PDF links from folders
- Export result as a hyperlinked Excel file

### 🎨 Theme Support
- 🌗 Dark/Light mode toggle
- Theme preference is saved between sessions

---

## 📁 Project Structure
```
erc_app/
│
├── main.py
├── theme_config.json (can be deleted since the system can make on its own after closing the app for the first time)
├── budget.xlsx
├── icon.png
├── sun.png
├── moon.png
├── README.md
├── obr_extractor.py
├── 2fa.png
├── njz.png
├── rename.png
├── theme_manager.py
├── users.json (can be deleted since the system can make on its own after signing up the first account)
│
├── core/
│   ├── budget_utils.py
│   ├── email_utils.py
│   ├── excel_utils.py
│   ├── file_utils.py
|   ├── logger.py
|   ├── pdf_tools.py
|   ├── pdf_utils.py
|   ├── sharepoint_tools.py
|   ├── sharepoint_utils.py
|   ├── user_auth.py
│
├── config/
│   ├── _init_.py
│   └── constants.py
|   ├── theme_config.json
|   ├── theme_config.py
│
├── ui_pages/
|   ├── activity_log_page.py
|   ├── earmark_page.py
|   ├── login_page.py
|   ├── main_menu.py
|   ├── main_window.py
|   ├── merge_page.py
|   ├── obr_fallback_dialog.py
|   ├── obr_page.py
|   ├── rename_option_dialog.py
|   ├── rename_page.py
|   ├── saro_fallback_dialog.py
|   ├── sharepoint_page.py
|   ├── signup_dialog.py
|   ├── split_page.py
|   ├── two_factor_dialog.py
|
|
├── utils/
    ├── dialogs.py
    ├── helpers.py
    ├── image_utils.py
```x`

---

## 🤝 Developers:
- Engr. Rolando Celeste - celeste.landon667@gmail.com
- Engr. Cris John Perez - perezcj2003@gmail.com

