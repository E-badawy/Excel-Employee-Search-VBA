# 🧠 Excel Employee Search & Automation System (VBA Project)

## 📘 Overview
This project demonstrates how to use **Excel VBA** and **ActiveX Controls** to create a functional and automated **Employee Record Search System**.  
It combines clean data management, user-friendly search capability, and dynamic column formatting — all within Microsoft Excel.

---

## ⚙️ Features
- 🔍 **Search Bar Functionality** — Search employee records by ID or Name.
- 📋 **Automated Result Display** — Matching results appear dynamically from Row 7.
- ⚡ **AutoFit Columns** — Automatically resizes columns for better readability.
- 🧩 **Macro-Enabled Workbook** — Saved as `.xlsm` for VBA compatibility.
- 🛡️ **Secure Trust Center Setup** — Ensures smooth ActiveX operation.

---

## 🧭 Setup Guide

### 1️⃣ Enable ActiveX Controls (Trust Center Configuration)
Before running any VBA macros, ensure that Excel can execute ActiveX components safely.

1. Go to **File → Options → Trust Center → Trust Center Settings**  
2. Click **ActiveX Settings**  
   - ✅ Enable all controls without restrictions  
   - ✅ Prompt me before enabling unsafe ActiveX controls  
3. Click **Macro Settings**  
   - ✅ Enable all macros  
   - ✅ Trust access to the VBA project object model  
4. Save and restart Excel.

---

### 2️⃣ Prepare Workbook Sheets
| Sheet Name | Purpose |
|-------------|----------|
| **Search Interface** | User search input and results display |
| **EmployeeData** | Holds full dataset (31 columns, starting from Row 1) |

- Add a **TextBox** (`TextBox1`) and a **CommandButton** (`CommandButton1`) on the *Search Interface* sheet.
- Save the workbook as **Macro-Enabled (.xlsm)**.

---

### 3️⃣ VBA Code Setup

#### 🔸 Search Function (Paste in CommandButton code)
```vba
Private Sub CommandButton1_Click()
    Dim wsData As Worksheet, wsSearch As Worksheet
    Dim searchValue As String
    Dim lastRow As Long, destRow As Long
    Dim cell As Range, rng As Range

    Set wsData = ThisWorkbook.Sheets("EmployeeData")
    Set wsSearch = ThisWorkbook.Sheets("Search Interface")

    searchValue = Trim(wsSearch.TextBox1.Text)
    If searchValue = "" Then
        MsgBox "Please enter a search term.", vbExclamation
        Exit Sub
    End If

    wsSearch.Rows("7:" & wsSearch.Rows.Count).ClearContents

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    Set rng = wsData.Range("A2:A" & lastRow)
    destRow = 7

    For Each cell In rng
        If InStr(1, cell.Value, searchValue, vbTextCompare) > 0 _
            Or InStr(1, cell.Offset(0, 1).Value, searchValue, vbTextCompare) > 0 Then
            cell.EntireRow.Copy wsSearch.Cells(destRow, "A")
            destRow = destRow + 1
        End If
    Next cell

    If destRow = 7 Then
        MsgBox "No matching records found.", vbInformation
    Else
        MsgBox "Search completed.", vbInformation
    End If
End Sub
```

---

#### 🔸 AutoFit Columns on Sheet Activation
*(Paste this into the “Search Interface” sheet module)*

```vba
Private Sub Worksheet_Activate()
    Cells.EntireColumn.AutoFit
End Sub
```

---

## 🧩 Project Logic Summary
1. The user enters a **search keyword** (Employee ID or Name).  
2. VBA loops through `EmployeeData` (Sheet2).  
3. If a match is found, the entire row is copied to `Search Interface` (Sheet1) starting from Row 7.  
4. When the sheet is reactivated, columns automatically resize for clarity.  

This setup is ideal for HR dashboards, employee directories, or data filtering tools.

---

## 🧠 Demo Section

### 🎥 Video Demo
📺 [![Watch the Demo](screenshots/search_interface.png)](excel_search_demo.mp4)  
*(Click image to view recorded demo)*  

> You can record your screen using **Xbox Game Bar (Windows + G)** or **Mac Screen Recorder (Command + Shift + 5)**, then save it as `excel_search_demo.mp4` in your project root.

---

### 📸 Screenshots
| Search Interface | Results Display |
|------------------|-----------------|
| ![Search Interface](screenshots/search_interface.png) | ![Results Display](screenshots/results_display.png) |

---

## 🗂️ Folder Structure
```
Excel-Employee-Search-VBA/
│
├── README.md
├── EmployeeSearch.xlsm
├── excel_search_demo.mp4
└── screenshots/
    ├── search_interface.png
    └── results_display.png
```

---

## 🧑‍💻 Author
**Badawi Aminu Muhammed**  
Data Analyst | Project Manager | Researcher  
📧 cigma.generalsolutions@gmail.com  
🌐 [linkedin.com/in/elameenbadawy](https://linkedin.com/in/elameenbadawy)

> *Cigma General Solutions — "…significant difference"*

---

## 💡 Notes
- To upload your Excel file on GitHub, ensure it’s a **macro-enabled `.xlsm` file**.  
- Videos larger than 25 MB should be uploaded to **Google Drive** or **YouTube**, then link to it here.  
- To keep your repo tidy, push screenshots and small demo files inside dedicated folders.

---

### 🌟 Tags
`Excel VBA` • `ActiveX Controls` • `Automation` • `Data Cleaning` • `Search Function` • `HR Analytics`
