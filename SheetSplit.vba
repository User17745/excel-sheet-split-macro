' --------------------------------------------------------------------------------
' Title: VBA Excel Sheet Split
' Filename: SheetSplit.vba
' Description: Simple VBA script to break a large excel sheet into multiple smaller ones.
' Usage: 1. Enable the developer tab from Excel options.
'        2. Go to the required sheet -> ribbon menu -> Developer tab and select the Visual Basic Editor
'        3. Select your workbook -> right click -> Insert -> Module.
'        4. Copy paste this code, save and exit the VBA editor
'        5. Developer tab -> Marcos -> Select "SheetSplit"
' Tested on: - Windows 10 Pro 64 bits + Office 2016 64 bits

' Credit:
'   - User "Fer García" & User "pnuts" from Github:
'       https://stackoverflow.com/questions/17997851/how-to-split-spreadsheet-into-multiple-spreadsheets-with-set-number-of-rows/18001183#18001183 
' --------------------------------------------------------------------------------

Sub SheetSplit()
  Dim wb As Workbook
  Dim ThisSheet As Worksheet
  Dim NumOfColumns As Integer
  Dim RangeToCopy As Range
  Dim RangeOfHeader As Range        'data (range) of header row
  Dim WorkbookCounter As Integer
  Dim RowsInFile                    'how many rows (incl. header) in new files?

  Application.ScreenUpdating = False

  'Initialize data
  Set ThisSheet = ThisWorkbook.ActiveSheet
  NumOfColumns = ThisSheet.UsedRange.Columns.Count
  WorkbookCounter = 1
  RowsInFile = 1001                   'Limiting to 1001 rows per file

  'Copy the data of the first row (header)
  Set RangeOfHeader = ThisSheet.Range(ThisSheet.Cells(1, 1), ThisSheet.Cells(1, NumOfColumns))

  For p = 2 To ThisSheet.UsedRange.Rows.Count Step RowsInFile - 1
    Set wb = Workbooks.Add

    'Paste the header row in new file
    RangeOfHeader.Copy wb.Sheets(1).Range("A1")

    'Paste the chunk of rows for this file
    Set RangeToCopy = ThisSheet.Range(ThisSheet.Cells(p, 1), ThisSheet.Cells(p + RowsInFile - 2, NumOfColumns))
    RangeToCopy.Copy wb.Sheets(1).Range("A2")

    'Save the new workbook, and close it
    wb.SaveAs ThisWorkbook.Path & "\kx_catalog_chunk_" & WorkbookCounter
    wb.Close

    'Increment file counter
    WorkbookCounter = WorkbookCounter + 1
  Next p

  Application.ScreenUpdating = True
  Set wb = Nothing
End Sub