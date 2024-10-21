# Inventory Tracking System

## Purpose of the File

This Excel workbook serves as an **inventory tracking system**, where each sheet represents different parts or product codes. The purpose of this file is to manage and monitor the availability, updates, and organization of specific items, typically used in **technical operations**, **logistics**, or **supply chain management**. It supports various part numbers and item codes, allowing users to easily track, update, and analyze the status of items in real-time.

### Features:
- **Each sheet** corresponds to a unique part or product code, enabling organized tracking of items.
- The first sheet, named `"ADD(ACTUAL)"`, serves as the **primary or summary sheet**, where new data or item adjustments can be inputted.
- Other sheets store detailed information about specific items or part numbers, potentially including **quantities**, **pricing**, or **availability**.
- Some sheets are marked with special labels like `™` or `(FT)` to signify certain classifications, such as **Fluteck** and **Tongmayong**.

## VBA Code for Automated Operations

The Excel workbook is automated using **Visual Basic for Applications (VBA)**, which enables repetitive tasks and dynamic management of sheets and data. Below is the VBA code used to handle various operations, such as creating new sheets, adding hyperlinks, and standardizing formatting.

### 1. Automatically Create Sheets for New Items

This VBA script dynamically creates new sheets when new part numbers are added, with the sheet name set to the part number in a given range.

```vba
Sub CreateNewSheets()
    Dim ws As Worksheet
    Dim partNumber As Range
    Dim newSheet As Worksheet
    
    ' Loop through a defined range in the main sheet ("ADD(ACTUAL)")
    For Each partNumber In Sheets("ADD(ACTUAL)").Range("D3:D56")
        ' Check if sheet already exists, if not, create it
        On Error Resume Next
        Set newSheet = Sheets(partNumber.Value)
        On Error GoTo 0
        
        If newSheet Is Nothing Then
            ' Add new sheet with part number as the name
            Set newSheet = Sheets.Add(After:=Sheets(Sheets.Count))
            newSheet.Name = partNumber.Value
        End If
    Next partNumber
End Sub
```

### 2. Hyperlink to Sheets

This script automatically adds hyperlinks to the corresponding sheet in the part number column.

```vba
Sub AddHyperlinks()
    Dim ws As Worksheet
    Dim partNumber As Range
    Dim mainSheet As Worksheet
    
    Set mainSheet = Sheets("ADD(ACTUAL)")
    
    ' Loop through part numbers in the main sheet and add hyperlinks
    For Each partNumber In mainSheet.Range("D3:D56")
        On Error Resume Next
        ' Create hyperlink to the sheet with the same name as the part number
        mainSheet.Hyperlinks.Add Anchor:=partNumber, Address:="", SubAddress:="'" & partNumber.Value & "'!A1", TextToDisplay:=partNumber.Value
        On Error GoTo 0
    Next partNumber
End Sub
```

### 3. Update Cell with Sheet Name

This script ensures that cell `E1` in each sheet (except "ADD(ACTUAL)") contains the sheet name.

```vba
Sub UpdateSheetNames()
    Dim ws As Worksheet
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "ADD(ACTUAL)" Then
            ws.Range("E1").Value = ws.Name
        End If
    Next ws
End Sub
```

### 4. Standardize Column Width

To ensure consistency in layout, this VBA script sets the column width of all columns to **115 pixels** across all sheets except "ADD(ACTUAL)".

```vba
Sub SetColumnWidth()
    Dim ws As Worksheet
    Dim colWidth As Double
    
    colWidth = 16.43 ' Approx 115 pixels
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "ADD(ACTUAL)" Then
            ws.Columns.ColumnWidth = colWidth
        End If
    Next ws
End Sub
```

### 5. Copy a Formula Across Sheets

If there’s a formula in `B3` of a source sheet (like `"K9000930"`), you can automatically copy it to `B3` of all other sheets.

```vba
Sub CopyFormulaToSheets()
    Dim ws As Worksheet
    Dim sourceSheet As Worksheet
    
    Set sourceSheet = Sheets("K9000930")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "ADD(ACTUAL)" And ws.Name <> "K9000930" Then
            ws.Range("B3").Formula = sourceSheet.Range("B3").Formula
        End If
    Next ws
End Sub
```

## How to Use:

1. **Open the Excel file** and press `Alt + F11` to access the VBA editor.
2. **Insert the VBA code** into a new module (`Insert` > `Module`).
3. **Run the macros** as needed by pressing `Alt + F8` and selecting the macro name.

These VBA scripts help automate the creation of new sheets, management of hyperlinks, and updating of specific cells. This makes the file easier to use for ongoing inventory tracking and management.

---

You can copy and paste this directly into your GitHub README file. Let me know if you need any more modifications!
