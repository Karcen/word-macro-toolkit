' ============================================================
'  Word Macro Toolkit — Table Utilities
'  Macro 1: SelectAllTables          — Select every table in the document
'  Macro 2: ConvertToThreeLineTables — Convert all tables to three-line style
'
'  How to install:
'    Open Word -> Alt+F11 -> Insert Module -> Paste this code -> Ctrl+S
'    Save the document as a macro-enabled file (.docm)
' ============================================================

Option Explicit

' ------------------------------------------------------------
'  Macro 1: Select All Tables
' ------------------------------------------------------------
Sub SelectAllTables()

    Dim doc        As Document
    Dim tableCount As Integer

    Set doc = ActiveDocument
    tableCount = doc.Tables.Count

    If tableCount = 0 Then
        MsgBox "No tables were found in this document.", _
               vbInformation, "Notice"
        Exit Sub
    End If

    If tableCount = 1 Then
        doc.Tables(1).Select
        MsgBox "1 table selected.", vbInformation, "Done"
        Exit Sub
    End If

    Dim fullRange As Range
    Set fullRange = doc.Range( _
        doc.Tables(1).Range.Start, _
        doc.Tables(tableCount).Range.End)
    fullRange.Select

    MsgBox "All " & tableCount & " tables selected " & _
           "(including any content between them).", _
           vbInformation, "Done"
End Sub


' ------------------------------------------------------------
'  Macro 2: Convert All Tables to Three-Line Style
' ------------------------------------------------------------
Sub ConvertToThreeLineTables()

    Dim doc        As Document
    Dim tableCount As Integer
    Dim i          As Integer

    Set doc = ActiveDocument
    tableCount = doc.Tables.Count

    If tableCount = 0 Then
        MsgBox "No tables were found in this document.", _
               vbInformation, "Notice"
        Exit Sub
    End If

    For i = 1 To tableCount
        Call ApplyThreeLineStyle(doc.Tables(i))
    Next i

    MsgBox "All " & tableCount & " table(s) converted to three-line style.", _
           vbInformation, "Done"
End Sub


' ------------------------------------------------------------
'  Helper: Apply three-line formatting to a single table
' ------------------------------------------------------------
Private Sub ApplyThreeLineStyle(tbl As Table)

    Const THICK_PT As Single = 1.5
    Const THIN_PT  As Single = 0.75
    Const LINE_CLR As Long   = 0     ' Black

    Dim b  As Integer
    Dim rw As Row
    Dim cl As Cell

    ' --- Border lists ---
    ' Table-level: diagonal borders are NOT valid here (causes error 5941)
    Dim tblBorders As Variant
    tblBorders = Array( _
        wdBorderTop, wdBorderBottom, wdBorderLeft, wdBorderRight, _
        wdBorderHorizontal, wdBorderVertical)

    ' Cell-level: diagonals are valid on individual cells
    Dim cellBorders As Variant
    cellBorders = Array( _
        wdBorderTop, wdBorderBottom, wdBorderLeft, wdBorderRight, _
        wdBorderDiagonalDown, wdBorderDiagonalUp)

    ' Step 1 — Remove all borders from the table object
    For b = 0 To UBound(tblBorders)
        tbl.Borders(tblBorders(b)).LineStyle = wdLineStyleNone
    Next b

    ' Step 2 — Remove all borders from every individual cell
    For Each rw In tbl.Rows
        For Each cl In rw.Cells
            For b = 0 To UBound(cellBorders)
                cl.Borders(cellBorders(b)).LineStyle = wdLineStyleNone
            Next b
        Next cl
    Next rw

    ' Step 3 — Top rule (thick)
    With tbl.Borders(wdBorderTop)
        .LineStyle = wdLineStyleSingle
        .LineWidth = PointsToLineWidth(THICK_PT)
        .Color     = LINE_CLR
    End With

    ' Step 4 — Bottom rule (thick)
    With tbl.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = PointsToLineWidth(THICK_PT)
        .Color     = LINE_CLR
    End With

    ' Step 5 — Header rule: bottom border of row 1 only (thin)
    If tbl.Rows.Count >= 1 Then
        For Each cl In tbl.Rows(1).Cells
            With cl.Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = PointsToLineWidth(THIN_PT)
                .Color     = LINE_CLR
            End With
        Next cl
    End If

End Sub


' ------------------------------------------------------------
'  Helper: Map a point value to the nearest WdLineWidth constant
' ------------------------------------------------------------
Private Function PointsToLineWidth(pt As Single) As WdLineWidth
    Select Case True
        Case pt <= 0.375:  PointsToLineWidth = wdLineWidth025pt
        Case pt <= 0.625:  PointsToLineWidth = wdLineWidth050pt
        Case pt <= 0.875:  PointsToLineWidth = wdLineWidth075pt
        Case pt <= 1.25:   PointsToLineWidth = wdLineWidth100pt
        Case pt <= 1.875:  PointsToLineWidth = wdLineWidth150pt
        Case pt <= 2.625:  PointsToLineWidth = wdLineWidth225pt
        Case pt <= 3.75:   PointsToLineWidth = wdLineWidth300pt
        Case pt <= 5.25:   PointsToLineWidth = wdLineWidth450pt
        Case Else:         PointsToLineWidth = wdLineWidth600pt
    End Select
End Function
