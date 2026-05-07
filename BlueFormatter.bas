Attribute VB_Name = "BlueFormatter"
'=====================================================================
' BlueFormatter
' ---------------------------------------------------------------------
' One-click recolor: cross-references, headings, captions, abstract,
' keywords -> blue. Supports both English and Chinese style names.
'
' Usage:
'   Open VBA editor (Alt+F11) -> File -> Import File -> select this .bas
'   Run macro: ChangeToBlue
'
' Author: generated for personal use. Public domain.
'=====================================================================

Option Explicit

' ---- Configurable color (Office Blue by default) --------------------
Private Const TARGET_R As Integer = 0
Private Const TARGET_G As Integer = 112
Private Const TARGET_B As Integer = 192


'======================================================================
' MAIN ENTRY POINT
'======================================================================
Public Sub ChangeToBlue()
    If Documents.Count = 0 Then
        MsgBox "No active document. Please open a Word file first." & vbCrLf & _
               "没有打开的文档，请先打开一个 Word 文件。", _
               vbExclamation, "BlueFormatter"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = ActiveDocument

    Dim blueColor As Long
    blueColor = RGB(TARGET_R, TARGET_G, TARGET_B)

    ' Save and disable state that could interfere
    Dim screenWasOn As Boolean
    Dim trackWasOn As Boolean
    Dim showRevWasOn As Boolean
    screenWasOn = Application.ScreenUpdating
    trackWasOn = doc.TrackRevisions
    showRevWasOn = doc.ShowRevisions

    Application.ScreenUpdating = False
    doc.TrackRevisions = False

    ' Counters
    Dim styleCount As Long, paraCount As Long, fldCount As Long
    styleCount = 0: paraCount = 0: fldCount = 0

    '------------------------------------------------------------------
    ' Step 1: Update style definitions (built-in IDs + named styles)
    '------------------------------------------------------------------
    Dim builtInIds As Variant
    builtInIds = Array(wdStyleHeading1, wdStyleHeading2, wdStyleHeading3, _
                       wdStyleHeading4, wdStyleHeading5, wdStyleHeading6, _
                       wdStyleHeading7, wdStyleHeading8, wdStyleHeading9, _
                       wdStyleCaption, wdStyleHyperlink, wdStyleTOC1, _
                       wdStyleTOC2, wdStyleTOC3, wdStyleTOC4, wdStyleTOC5, _
                       wdStyleTOC6, wdStyleTOC7, wdStyleTOC8, wdStyleTOC9)

    Dim bId As Variant
    For Each bId In builtInIds
        If SetStyleColorById(doc, CLng(bId), blueColor) Then
            styleCount = styleCount + 1
        End If
    Next bId

    ' Custom / non-built-in style names (English + Chinese variants)
    Dim customNames As Variant
    customNames = Array( _
        "Abstract", "Abstract Title", "Abstract Heading", "Abstract Body", _
        "Keywords", "Keyword", "Key Words", "Keyword Text", "Keywords Heading", _
        "摘要", "摘要标题", "摘要正文", "摘要内容", "中文摘要", "英文摘要", _
        "关键词", "关键字", "关键词标题", "关键词正文", "中文关键词", "英文关键词", _
        "题注", "超链接", _
        "Heading 1 Char", "Heading 2 Char", "Heading 3 Char", _
        "Heading 4 Char", "Heading 5 Char", "Heading 6 Char", _
        "Caption Char")

    Dim cn As Variant
    For Each cn In customNames
        If SetStyleColorByName(doc, CStr(cn), blueColor) Then
            styleCount = styleCount + 1
        End If
    Next cn

    '------------------------------------------------------------------
    ' Step 2: Recolor cross-reference fields in every story range
    '------------------------------------------------------------------
    Dim story As Range
    For Each story In doc.StoryRanges
        RecolorFieldsInRange story, blueColor, fldCount
        ' Walk linked stories (multi-section headers/footers)
        Dim nextStory As Range
        Set nextStory = story.NextStoryRange
        Do Until nextStory Is Nothing
            RecolorFieldsInRange nextStory, blueColor, fldCount
            Set nextStory = nextStory.NextStoryRange
        Loop
    Next story

    '------------------------------------------------------------------
    ' Step 3: Force-apply blue to paragraphs that use target styles.
    '         Catches direct formatting that overrides style definition.
    '------------------------------------------------------------------
    For Each story In doc.StoryRanges
        RecolorParagraphsInRange story, blueColor, paraCount
        Dim nxt As Range
        Set nxt = story.NextStoryRange
        Do Until nxt Is Nothing
            RecolorParagraphsInRange nxt, blueColor, paraCount
            Set nxt = nxt.NextStoryRange
        Loop
    Next story

    '------------------------------------------------------------------
    ' Step 4: Handle text inside shapes / text boxes (best effort)
    '------------------------------------------------------------------
    RecolorShapes doc, blueColor, fldCount, paraCount

    '------------------------------------------------------------------
    ' Restore state
    '------------------------------------------------------------------
    doc.TrackRevisions = trackWasOn
    doc.ShowRevisions = showRevWasOn
    Application.ScreenUpdating = screenWasOn
    Application.ScreenRefresh

    MsgBox "Done! / 完成！" & vbCrLf & vbCrLf & _
           "Styles updated / 样式更新:     " & styleCount & vbCrLf & _
           "Paragraphs updated / 段落更新: " & paraCount & vbCrLf & _
           "Fields updated / 字段更新:     " & fldCount, _
           vbInformation, "BlueFormatter"
End Sub


'======================================================================
' DIAGNOSTIC: list every style name actually used in the document.
' Run this if the main macro misses something - it tells you the
' exact style name to add to customNames.
'======================================================================
Public Sub ListUsedStyles()
    If Documents.Count = 0 Then Exit Sub
    Dim doc As Document
    Set doc = ActiveDocument

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim para As Paragraph
    Dim sname As String
    For Each para In doc.Paragraphs
        sname = SafeStyleName(para)
        If Len(sname) > 0 Then
            If Not dict.Exists(sname) Then dict.Add sname, 1
        End If
    Next para

    Dim out As String
    out = "Styles in use (" & dict.Count & "):" & vbCrLf & String(40, "-") & vbCrLf
    Dim k As Variant
    For Each k In dict.Keys
        out = out & k & vbCrLf
    Next k
    MsgBox out, vbInformation, "BlueFormatter - ListUsedStyles"
End Sub


'======================================================================
' HELPERS
'======================================================================

' Set color of a built-in style by its constant ID.
' Returns True on success, False if style is unavailable.
Private Function SetStyleColorById(doc As Document, builtInId As Long, _
                                    clr As Long) As Boolean
    On Error Resume Next
    Dim s As Style
    Set s = doc.Styles(builtInId)
    If Err.Number <> 0 Or s Is Nothing Then
        SetStyleColorById = False
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    s.Font.ColorIndex = wdAuto
    Err.Clear
    s.Font.Color = clr
    SetStyleColorById = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


' Set color of a style looked up by name. Tries the name as-is.
Private Function SetStyleColorByName(doc As Document, styleName As String, _
                                      clr As Long) As Boolean
    If Len(styleName) = 0 Then
        SetStyleColorByName = False
        Exit Function
    End If

    On Error Resume Next
    Dim s As Style
    Set s = doc.Styles(styleName)
    If Err.Number <> 0 Or s Is Nothing Then
        SetStyleColorByName = False
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    s.Font.ColorIndex = wdAuto
    Err.Clear
    s.Font.Color = clr
    SetStyleColorByName = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


' Recolor cross-reference fields inside a Range.
Private Sub RecolorFieldsInRange(rng As Range, clr As Long, ByRef fldCount As Long)
    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    Dim fld As Field
    For Each fld In rng.Fields
        Select Case fld.Type
            Case wdFieldRef, wdFieldPageRef, wdFieldNoteRef, _
                 wdFieldFootnoteRef, wdFieldEndnoteRef, _
                 wdFieldHyperlink, wdFieldStyleRef, _
                 wdFieldSequence, wdFieldTOC, wdFieldTOA
                Err.Clear
                fld.Result.Font.Color = clr
                If Err.Number = 0 Then fldCount = fldCount + 1
                Err.Clear
        End Select
    Next fld
    On Error GoTo 0
End Sub


' Recolor every paragraph in Range that uses a target style.
Private Sub RecolorParagraphsInRange(rng As Range, clr As Long, _
                                      ByRef paraCount As Long)
    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    Dim para As Paragraph
    Dim sname As String
    For Each para In rng.Paragraphs
        sname = SafeStyleName(para)
        If IsTargetStyle(sname) Then
            Err.Clear
            para.Range.Font.Color = clr
            If Err.Number = 0 Then paraCount = paraCount + 1
            Err.Clear
        End If
    Next para
    On Error GoTo 0
End Sub


' Recolor inside floating shapes / text boxes.
Private Sub RecolorShapes(doc As Document, clr As Long, _
                           ByRef fldCount As Long, ByRef paraCount As Long)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In doc.Shapes
        If shp.TextFrame.HasText Then
            Dim r As Range
            Set r = shp.TextFrame.TextRange
            RecolorFieldsInRange r, clr, fldCount
            RecolorParagraphsInRange r, clr, paraCount
        End If
    Next shp

    Dim ishp As InlineShape
    For Each ishp In doc.InlineShapes
        ' Inline shapes rarely host text bodies but try anyway
        If ishp.OLEFormat Is Nothing Then
            ' nothing to do
        End If
    Next ishp
    On Error GoTo 0
End Sub


' Get a paragraph's style name without throwing.
Private Function SafeStyleName(para As Paragraph) As String
    On Error Resume Next
    Dim n As String
    n = ""
    n = CStr(para.Style)
    SafeStyleName = n
    Err.Clear
    On Error GoTo 0
End Function


' Decide whether a style name is one we want to recolor.
' Matches English and Chinese variants, with tolerance for
' spacing/case and "Char" companion styles.
Private Function IsTargetStyle(styleName As String) As Boolean
    IsTargetStyle = False
    If Len(styleName) = 0 Then Exit Function

    Dim lower As String
    lower = LCase$(Trim$(styleName))

    '--- Heading 1..9 (English) ---
    Dim i As Integer
    For i = 1 To 9
        If lower = "heading " & i Then IsTargetStyle = True: Exit Function
        If lower = "heading" & i Then IsTargetStyle = True: Exit Function
        If lower = "heading " & i & " char" Then IsTargetStyle = True: Exit Function
    Next i

    '--- 标题 1..9 (Chinese) ---
    For i = 1 To 9
        If styleName = "标题 " & i Then IsTargetStyle = True: Exit Function
        If styleName = "标题" & i Then IsTargetStyle = True: Exit Function
    Next i

    '--- Caption / 题注 ---
    If lower = "caption" Then IsTargetStyle = True: Exit Function
    If lower = "caption char" Then IsTargetStyle = True: Exit Function
    If styleName = "题注" Then IsTargetStyle = True: Exit Function

    '--- Abstract / 摘要 ---
    If InStr(lower, "abstract") > 0 Then IsTargetStyle = True: Exit Function
    If InStr(styleName, "摘要") > 0 Then IsTargetStyle = True: Exit Function

    '--- Keywords / 关键词 ---
    If InStr(lower, "keyword") > 0 Then IsTargetStyle = True: Exit Function
    If InStr(lower, "key words") > 0 Then IsTargetStyle = True: Exit Function
    If InStr(styleName, "关键词") > 0 Then IsTargetStyle = True: Exit Function
    If InStr(styleName, "关键字") > 0 Then IsTargetStyle = True: Exit Function

    '--- TOC entries (often colored differently) ---
    If lower Like "toc#" Or lower Like "toc #" Then IsTargetStyle = True: Exit Function
    If styleName Like "目录 #" Then IsTargetStyle = True: Exit Function
End Function
