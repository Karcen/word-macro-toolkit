' ============================================================
'  Macro 1: SelectAllTables          — Select every table in the document
'  Macro 2: ConvertToThreeLineTables — Convert all tables to three-line style
' ============================================================

Option Explicit

' ------------------------------------------------------------
'  Macro 1: Select All Tables
' ------------------------------------------------------------
Sub SelectAllTables()

    Dim doc        As Document
    Dim tbl        As Table
    Dim tableCount As Integer

    Set doc = ActiveDocument
    tableCount = doc.Tables.Count

    If tableCount = 0 Then
        MsgBox "No tables found in this document.", vbInformation, "Notice"
        Exit Sub
    End If

    ' Select each table one by one
    ' Word does not support discontinuous selections,
    ' so each call moves the cursor to that table.
    ' All tables are visited; the last one remains selected.
    For Each tbl In doc.Tables
        tbl.Select
    Next tbl

    MsgBox tableCount & " table(s) found and selected.", vbInformation, "Done"
End Sub


' ------------------------------------------------------------
'  Macro 2: Convert All Tables to Three-Line Style
'
'  Three-line rules:
'    Top border    — 1.5 pt thick
'    Bottom border — 1.5 pt thick
'    Row-1 bottom  — 0.75 pt thin  (header rule)
'    All others    — removed
' ------------------------------------------------------------
Sub ConvertToThreeLineTables()

    Dim doc        As Document
    Dim tableCount As Integer
    Dim i          As Integer

    Set doc = ActiveDocument
    tableCount = doc.Tables.Count

    If tableCount = 0 Then
        MsgBox "No tables found in this document.", vbInformation, "Notice"
        Exit Sub
    End If

    For i = 1 To tableCount
        Call ApplyThreeLineStyle(doc.Tables(i))
    Next i

    MsgBox tableCount & " table(s) converted to three-line style.", _
           vbInformation, "Done"
End Sub


' ------------------------------------------------------------
'  Helper: Apply three-line style to one table
' ------------------------------------------------------------
Private Sub ApplyThreeLineStyle(tbl As Table)

    Dim rw As Row
    Dim cl As Cell
    Dim b  As Integer

    ' Table-level borders (no diagonals — they are cell-only)
    Dim tblB As Variant
    tblB = Array(wdBorderTop, wdBorderBottom, wdBorderLeft, wdBorderRight, _
                 wdBorderHorizontal, wdBorderVertical)

    ' Cell-level borders (diagonals included)
    Dim cellB As Variant
    cellB = Array(wdBorderTop, wdBorderBottom, wdBorderLeft, wdBorderRight, _
                  wdBorderDiagonalDown, wdBorderDiagonalUp)

    ' Step 1 — Clear all table-level borders
    For b = 0 To UBound(tblB)
        tbl.Borders(tblB(b)).LineStyle = wdLineStyleNone
    Next b

    ' Step 2 — Clear all cell-level borders
    For Each rw In tbl.Rows
        For Each cl In rw.Cells
            For b = 0 To UBound(cellB)
                cl.Borders(cellB(b)).LineStyle = wdLineStyleNone
            Next b
        Next cl
    Next rw

    ' Step 3 — Top rule: 1.5 pt
    With tbl.Borders(wdBorderTop)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .Color = wdColorBlack
    End With

    ' Step 4 — Bottom rule: 1.5 pt
    With tbl.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .Color = wdColorBlack
    End With

    ' Step 5 — Header rule (row 1 bottom): 0.75 pt
    If tbl.Rows.Count >= 1 Then
        For Each cl In tbl.Rows(1).Cells
            With cl.Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth075pt
                .Color = wdColorBlack
            End With
        Next cl
    End If

End Sub
