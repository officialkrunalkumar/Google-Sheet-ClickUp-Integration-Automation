# 🛠 Krunal's Google Sheets → ClickUp Automation

This Google Apps Script automates task management between a Google Sheet and [ClickUp](https://clickup.com/).  
It provides menu options, automatic row management, and ClickUp task creation directly from your spreadsheet.

---

## 📌 Features

- **Custom Menu in Google Sheets**
  - Move overdue or completed rows to an `Archived` sheet.
  - Create a ClickUp task for the current row.
  - Bulk-create ClickUp tasks for all qualifying rows.

- **Automatic Row Archiving**
  - Detects overdue or completed tasks based on **Closing Date** and **Status**.
  - Moves them from `Active` to `Archived` automatically on edit or manually via menu.

- **ClickUp Integration**
  - Creates ClickUp tasks from rows marked `"Moved to Clickup"`.
  - Stores ClickUp `Task ID` and `Task Link` in new columns.
  - Prevents duplicate task creation.
  - Ensures required columns are filled before creating tasks.

- **Bulk Task Creation**
  - Automatically processes all rows with `"Moved to Clickup"` status in one go.

---

## 📂 Sheet Setup

Your Google Sheet must have:
- **Sheet Names**
  - `Active` → Source sheet for current tenders.
  - `Archived` → Destination for overdue/done tenders.

- **Headers**
  These must be present in the first row:
  ```text
  Portal
  Tender Title
  Buyer/Issuing Authority
  Short Description
  Closing Date
  Estimated Value
  Link to Tender
  Status
