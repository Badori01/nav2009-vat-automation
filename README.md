# 📊 DSP - VAT Report Automation Tool (Microsoft Dynamics NAV 2009)

A high-utility **VBA (Excel Macro)** automation tool designed to clean, merge, and reconcile VAT reports extracted from **Microsoft Dynamics NAV 2009 Classic** (backed by **MS SQL Server**) for **DSP**, a subsidiary of **BT Bin Laden**.

---

## 📝 Project Overview

This project automates a tedious, error-prone manual data cleaning process. Previously, financial data had to be exported as a PDF, converted via *I love pdf*, and then manually structured and filtered. 

This macro completely replaces the manual workflow, compressing hours of tedious data scrubbing into a single button click with 100% precision.

---

## 🛠️ The Manual Workflow vs. Automated Solution

| Manual Steps (Before) | Automated Solution (After) |
| :--- | :--- |
| Extract VAT report from NAV 2009 as a PDF. | Raw converted Excel data is processed directly. |
| Use *I love pdf* to convert PDF into multi-sheet Excel files. | Streamlined macro runs instantly on the workbook. |
| Manually find and delete odd-numbered "Table #" sheets. | **Automated:** Deletes all odd sheets using dynamic string analysis. |
| Copy rows 2 to 47 from even sheets into a master sheet. | **Automated:** Consolidates rows 2–47 accurately, ignoring blanks. |
| Scan Column H to filter out zeros, empty spaces, and text. | **Automated:** Safely deletes invalid or zero-value VAT rows (Bottom-up). |
| Manually look up and offset matching positive/negative pairs. | **Automated:** Cross-references `Credit Mem` rows via strict multi-column matching. |

---

## 🚀 Key Technical Features

* **Odd-Sheet Elimination:** Automatically scans sheet names and drops odd-numbered tables (`Table 1`, `Table 3`, etc.) generated during PDF conversion.
* **Smart Data Consolidation:** Dynamically builds a `Cleaned_Data` sheet, preserves the original header from the first data page, and appends records.
* **Advanced Row Cleaning:** Multi-condition filtering on Column H to wipe out empty rows, string errors, and absolute zero `0.00` values.
* **Dictionary-Based Deduplication:** Utilizes `Scripting.Dictionary` to look up unique identifiers in Column C, accumulating totals for columns G, H, and I while wiping duplicates.
* **Strict Reconciliation (Reversal Matching):** Tracks negative values in Column G, verifies if they are flagged as `Credit Mem` in Column B, and runs a strict verification across Columns A, E, and F against their positive counterparts before executing a clean dual-row elimination.

---

## 📂 System Architecture & Technologies

* **Source ERP:** Microsoft Dynamics NAV 2009 Classic
* **Database Backend:** Microsoft SQL Server
* **Automation Language:** VBA (Visual Basic for Applications)
* **Host Environment:** Microsoft Excel

---

## 💻 Source Code

Below is the complete automation script. You can save this inside a standard Excel module (`.bas` file) or paste it directly into your workbook's VBA project.

```vba
Sub CleanMergeAndFilterColumnH()

    Dim wb As Workbook

    Dim wsNew As Worksheet

    Dim ws As Worksheet

    Dim i As Long, r As Long, j As Long

    Dim sheetName As String

    Dim sheetNum As Variant

    Dim lastRow As Long

    Dim destLastRow As Long

    Dim cellValue As Variant

   

    ' Additional variables for merging duplicates

    Dim dict As Object

    Dim key As String

    Dim firstRowIndex As Long

   

    ' Additional variables for positive/negative matching

    Dim valR As Double, valJ As Double

    Dim matchFound As Boolean

   

    Set wb = ThisWorkbook

   

    ' 1. Delete sheets with odd numbers in their names (e.g., Table 1, Table 3...)

    Application.DisplayAlerts = False

    For i = wb.Worksheets.Count To 1 Step -1

        Set ws = wb.Worksheets(i)

        sheetName = ws.Name

       

        If InStr(1, sheetName, "Table", vbTextCompare) > 0 Then

            sheetNum = val(Mid(sheetName, InStr(sheetName, " ") + 1))

            If sheetNum Mod 2 <> 0 Then

                ws.Delete

            End If

        End If

    Next i

    Application.DisplayAlerts = True

   

    ' 2. Create a new fresh worksheet named "Cleaned_Data"

    Application.DisplayAlerts = False

    On Error Resume Next

    wb.Worksheets("Cleaned_Data").Delete ' Delete if it already exists

    On Error GoTo 0

    Application.DisplayAlerts = True

   

    Set wsNew = wb.Worksheets.Add(Before:=wb.Worksheets(1))

    wsNew.Name = "Cleaned_Data"

   

    ' Copy the header (Row 1) from the first remaining data sheet (now sheet number 2)

    wb.Worksheets(2).Rows(1).Copy Destination:=wsNew.Rows(1)

   

    ' 3. Loop through all old data sheets and merge data from row 2 to 47

    i = 2

    Do While wb.Worksheets.Count >= i

        Set ws = wb.Worksheets(i)

       

        ' Find the actual last row using an accurate method for PDF converted sheets

        If ws.UsedRange.Rows.Count <= 1 And ws.Cells(1, "A").Value = "" Then

            lastRow = 0

        Else

            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            If lastRow = 1 And ws.Cells(2, "B").Value <> "" Then

                lastRow = ws.UsedRange.Rows.Count

            End If

        End If

       

        ' Force the upper limit to 47

        If lastRow > 47 Then lastRow = 47

       

        ' STRICT CHECK: Only copy if there is actual data starting from row 2

        If lastRow >= 2 Then

            destLastRow = wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).Row + 1

            ws.Range("A2:Z" & lastRow).Copy Destination:=wsNew.Range("A" & destLastRow)

        End If

       

        ' Delete the old processed sheet

        Application.DisplayAlerts = False

        ws.Delete

        Application.DisplayAlerts = True

    Loop

   

    ' 4. Filter and clean Column H in the "Cleaned_Data" sheet

    lastRow = wsNew.Cells(wsNew.Rows.Count, "H").End(xlUp).Row

   

    For r = lastRow To 2 Step -1

        cellValue = wsNew.Cells(r, "H").Value

       

        ' Check if the cell is empty, contains text (letters), or equals 0

        If IsEmpty(wsNew.Cells(r, "H")) Then

            wsNew.Rows(r).Delete

        ElseIf IsText(cellValue) Then

            wsNew.Rows(r).Delete

        ElseIf IsNumeric(cellValue) And val(cellValue) = 0 Then

            wsNew.Rows(r).Delete

        End If

    Next r

   

    ' 5. Process duplicates in Column C, sum G, H, I and flag rows for deletion

    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsNew.Cells(wsNew.Rows.Count, "C").End(xlUp).Row

   

    For r = 2 To lastRow

        key = Trim(CStr(wsNew.Cells(r, "C").Value))

       

        If key <> "" Then

            If Not dict.Exists(key) Then

                ' If the number appears for the first time, store its row index

                dict.Add key, r

            Else

                ' If the number is a duplicate, get the first row index

                firstRowIndex = dict(key)

               

                ' Sum values for columns G, H, and I into the first row

                wsNew.Cells(firstRowIndex, "G").Value = val(wsNew.Cells(firstRowIndex, "G").Value) + val(wsNew.Cells(r, "G").Value)

                wsNew.Cells(firstRowIndex, "H").Value = val(wsNew.Cells(firstRowIndex, "H").Value) + val(wsNew.Cells(r, "H").Value)

                wsNew.Cells(firstRowIndex, "I").Value = val(wsNew.Cells(firstRowIndex, "I").Value) + val(wsNew.Cells(r, "I").Value)

               

                ' Flag the current row for deletion

                wsNew.Cells(r, "C").Value = "DELETE_ROW"

            End If

        End If

    Next r

   

    ' Delete the rows flagged as duplicates (Bottom to Top)

    For r = lastRow To 2 Step -1

        If wsNew.Cells(r, "C").Value = "DELETE_ROW" Then

            wsNew.Rows(r).Delete

        End If

    Next r

   

    ' 6. Match Positive/Negative values in Column G with Credit Memo and strict column verification

    lastRow = wsNew.Cells(wsNew.Rows.Count, "G").End(xlUp).Row

    

    For r = lastRow To 2 Step -1

        ' Check if row still exists and Column G has a numeric value

        If IsNumeric(wsNew.Cells(r, "G").Value) And wsNew.Cells(r, "G").Value <> "" Then

            valR = CDbl(wsNew.Cells(r, "G").Value)

           

            ' Look for negative value first to check "Credit Mem" condition

            If valR < 0 Then

                ' Verify if Column B contains "Credit Mem" (case-insensitive check)

                If InStr(1, CStr(wsNew.Cells(r, "B").Value), "Credit Mem", vbTextCompare) > 0 Then

                    matchFound = False

                   

                    ' Search for the matching positive value in other rows

                    For j = 2 To lastRow

                        If r <> j And IsNumeric(wsNew.Cells(j, "G").Value) And wsNew.Cells(j, "G").Value <> "" Then

                            valJ = CDbl(wsNew.Cells(j, "G").Value)

                           

                            ' Check if values are equal but opposite in sign

                            If valJ = Abs(valR) Then

                                ' Strict verification: Check if Columns A, E, and F are identical

                                If CStr(wsNew.Cells(r, "A").Value) = CStr(wsNew.Cells(j, "A").Value) And _

                                   CStr(wsNew.Cells(r, "E").Value) = CStr(wsNew.Cells(j, "E").Value) And _

                                   CStr(wsNew.Cells(r, "F").Value) = CStr(wsNew.Cells(j, "F").Value) Then

                                   

                                    ' Flag both rows for combined elimination

                                    wsNew.Cells(r, "G").Value = "MATCHED_DEL"

                                    wsNew.Cells(j, "G").Value = "MATCHED_DEL"

                                    matchFound = True

                                    Exit For

                                End If

                            End If

                        End If

                    Next j

                End If

            End If

        End If

    Next r

   

    ' Delete the matched positive/negative rows (Bottom to Top)

    For r = lastRow To 2 Step -1

        If wsNew.Cells(r, "G").Value = "MATCHED_DEL" Then

            wsNew.Rows(r).Delete

        End If

    Next r

   

    MsgBox "Data cleaning, merging, and Column H filtering completed successfully!", vbInformation, "Success"

End Sub

 

' Helper function to accurately check if a cell contains text/letters

Function IsText(val As Variant) As Boolean

    If VarType(val) = vbString Then

        If IsNumeric(val) = False Then

            IsText = True

            Exit Function

        End If

    End If

    IsText = False

End Function 
```
 ## 🚀 How to Run the Macro

1. Convert your extracted **NAV 2009** PDF report to Excel.
2. Open the converted workbook and press `ALT + F11` to trigger the VBA editor.
3. Click `Insert` -> `Module` and paste the script above.
4. Close the editor, return to your Excel sheet, and press `ALT + F8`.
5. Select `CleanMergeAndFilterColumnH` and click **Run**.

---

### 💡 Project Notes & Optimization
* **Automated Reconciliation:** This script includes an active validation pass that automatically matches positive and negative ledger balances based on `Credit Mem` conditions.
* **Dual-Row Elimination:** Once matching records are detected across identical columns (A, E, and F), both offsetting lines are cleanly purged from the sheet.
* **Process Safeguards:** Includes integrated bottom-up loops to prevent row-skipping errors and fires a final success notification box upon completion.
