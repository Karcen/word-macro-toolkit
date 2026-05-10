
'=====================================================================
' BlueFormatter
' ---------------------------------------------------------------------
' One-click recolor:
'   - Cross-references
'   - Headings
'   - Captions
'   - Abstract
'   - Keywords
'
' Supports English style names.
'
' Usage:
'   1. Open VBA Editor (Alt + F11)
'   2. File -> Import File
'   3. Select this .bas file
'   4. Run macro: ChangeToBlue
'
' Author: Karcen Zheng
'=====================================================================

Option Explicit

'---------------------------------------------------------------------
' Configurable target color (Office Blue)
'---------------------------------------------------------------------
Private Const TARGET_R As Integer = 0
Private Const TARGET_G As Integer = 112
Private Const TARGET_B As Integer = 192


'=====================================================================
' MAIN ENTRY
'=====================================================================
Public Sub ChangeToBlue()

    If Documents.count = 0 Then
        MsgBox "No active document. Please open a Word document first.", _
               vbExclamation, "BlueFormatter"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = ActiveDocument

    Dim blueColor As Long
    blueColor = RGB(TARGET_R, TARGET_G, TARGET_B)

    ' Save current application state
    Dim screenUpdatingState As Boolean
    Dim trackRevisionState As Boolean
    Dim showRevisionState As Boolean

    screenUpdatingState = Application.ScreenUpdating
    trackRevisionState = doc.TrackRevisions
    showRevisionState = doc.ShowRevisions

    Application.ScreenUpdating = False
    doc.TrackRevisions = False

    ' Counters
    Dim styleCount As Long
    Dim paragraphCount As Long
    Dim fieldCount As Long

    styleCount = 0
    paragraphCount = 0
    fieldCount = 0

    '-----------------------------------------------------------------
    ' Step 1: Update built-in styles
    '-----------------------------------------------------------------
    Dim builtInStyles As Variant

    builtInStyles = Array( _
        wdStyleHeading1, wdStyleHeading2, wdStyleHeading3, _
        wdStyleHeading4, wdStyleHeading5, wdStyleHeading6, _
        wdStyleHeading7, wdStyleHeading8, wdStyleHeading9, _
        wdStyleCaption, _
        wdStyleHyperlink, _
        wdStyleTOC1, wdStyleTOC2, wdStyleTOC3, _
        wdStyleTOC4, wdStyleTOC5, wdStyleTOC6, _
        wdStyleTOC7, wdStyleTOC8, wdStyleTOC9)

    Dim styleId As Variant

    For Each styleId In builtInStyles
        If SetStyleColorById(doc, CLng(styleId), blueColor) Then
            styleCount = styleCount + 1
        End If
    Next styleId

    '-----------------------------------------------------------------
    ' Step 2: Update custom style names
    '-----------------------------------------------------------------
    Dim customStyles As Variant

    customStyles = Array( _
        "Abstract", _
        "Abstract Title", _
        "Abstract Heading", _
        "Abstract Body", _
        "Keywords", _
        "Keyword", _
        "Key Words", _
        "Keyword Text", _
        "Keywords Heading", _
        "Heading 1 Char", _
        "Heading 2 Char", _
        "Heading 3 Char", _
        "Heading 4 Char", _
        "Heading 5 Char", _
        "Heading 6 Char", _
        "Caption Char")

    Dim styleName As Variant

    For Each styleName In customStyles
        If SetStyleColorByName(doc, CStr(styleName), blueColor) Then
            styleCount = styleCount + 1
        End If
    Next styleName

    '-----------------------------------------------------------------
    ' Step 3: Recolor fields in all story ranges
    '-----------------------------------------------------------------
    Dim story As Range

    For Each story In doc.StoryRanges

        RecolorFieldsInRange story, blueColor, fieldCount

        Dim nextStory As Range
        Set nextStory = story.NextStoryRange

        Do Until nextStory Is Nothing
            RecolorFieldsInRange nextStory, blueColor, fieldCount
            Set nextStory = nextStory.NextStoryRange
        Loop

    Next story

    '-----------------------------------------------------------------
    ' Step 4: Recolor paragraphs using target styles
    '-----------------------------------------------------------------
    For Each story In doc.StoryRanges

        RecolorParagraphsInRange story, blueColor, paragraphCount

        Dim linkedStory As Range
        Set linkedStory = story.NextStoryRange

        Do Until linkedStory Is Nothing
            RecolorParagraphsInRange linkedStory, blueColor, paragraphCount
            Set linkedStory = linkedStory.NextStoryRange
        Loop

    Next story

    '-----------------------------------------------------------------
    ' Step 5: Recolor shapes and text boxes
    '-----------------------------------------------------------------
    RecolorShapes doc, blueColor, fieldCount, paragraphCount

    '-----------------------------------------------------------------
    ' Restore original application state
    '-----------------------------------------------------------------
    doc.TrackRevisions = trackRevisionState
    doc.ShowRevisions = showRevisionState
    Application.ScreenUpdating = screenUpdatingState

    Application.ScreenRefresh

    MsgBox _
        "Completed successfully." & vbCrLf & vbCrLf & _
        "Styles updated: " & styleCount & vbCrLf & _
        "Paragraphs updated: " & paragraphCount & vbCrLf & _
        "Fields updated: " & fieldCount, _
        vbInformation, "BlueFormatter"

End Sub


'=====================================================================
' DIAGNOSTIC TOOL
' Lists all styles currently used in the document
'=====================================================================
Public Sub ListUsedStyles()

    If Documents.count = 0 Then Exit Sub

    Dim doc As Document
    Set doc = ActiveDocument

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim para As Paragraph
    Dim currentStyle As String

    For Each para In doc.Paragraphs

        currentStyle = SafeStyleName(para)

        If Len(currentStyle) > 0 Then
            If Not dict.Exists(currentStyle) Then
                dict.Add currentStyle, 1
            End If
        End If

    Next para

    Dim outputText As String
    outputText = "Styles currently in use (" & dict.count & "):" & _
                 vbCrLf & String(40, "-") & vbCrLf

    Dim key As Variant

    For Each key In dict.Keys
        outputText = outputText & key & vbCrLf
    Next key

    MsgBox outputText, vbInformation, "BlueFormatter - Style Report"

End Sub


'=====================================================================
' HELPER FUNCTIONS
'=====================================================================

'---------------------------------------------------------------------
' Set font color for a built-in style
'---------------------------------------------------------------------
Private Function SetStyleColorById( _
    doc As Document, _
    builtInId As Long, _
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
    s.Font.Color = clr

    SetStyleColorById = (Err.Number = 0)

    Err.Clear
    On Error GoTo 0

End Function


'---------------------------------------------------------------------
' Set font color for a style by name
'---------------------------------------------------------------------
Private Function SetStyleColorByName( _
    doc As Document, _
    styleName As String, _
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
    s.Font.Color = clr

    SetStyleColorByName = (Err.Number = 0)

    Err.Clear
    On Error GoTo 0

End Function


'---------------------------------------------------------------------
' Recolor cross-reference related fields
'---------------------------------------------------------------------
Private Sub RecolorFieldsInRange( _
    rng As Range, _
    clr As Long, _
    ByRef fieldCount As Long)

    If rng Is Nothing Then Exit Sub

    On Error Resume Next

    Dim fld As Field

    For Each fld In rng.Fields

        Select Case fld.Type

            Case wdFieldRef, _
                 wdFieldPageRef, _
                 wdFieldNoteRef, _
                 wdFieldFootnoteRef, _
                 wdFieldHyperlink, _
                 wdFieldStyleRef, _
                 wdFieldSequence, _
                 wdFieldTOC, _
                 wdFieldTOA

                fld.Result.Font.Color = clr

                If Err.Number = 0 Then
                    fieldCount = fieldCount + 1
                End If

                Err.Clear

        End Select

    Next fld

    On Error GoTo 0

End Sub


'---------------------------------------------------------------------
' Recolor paragraphs using matching styles
'---------------------------------------------------------------------
Private Sub RecolorParagraphsInRange( _
    rng As Range, _
    clr As Long, _
    ByRef paragraphCount As Long)

    If rng Is Nothing Then Exit Sub

    On Error Resume Next

    Dim para As Paragraph
    Dim styleName As String

    For Each para In rng.Paragraphs

        styleName = SafeStyleName(para)

        If IsTargetStyle(styleName) Then

            para.Range.Font.Color = clr

            If Err.Number = 0 Then
                paragraphCount = paragraphCount + 1
            End If

            Err.Clear

        End If

    Next para

    On Error GoTo 0

End Sub


'---------------------------------------------------------------------
' Recolor text inside shapes and text boxes
'---------------------------------------------------------------------
Private Sub RecolorShapes( _
    doc As Document, _
    clr As Long, _
    ByRef fieldCount As Long, _
    ByRef paragraphCount As Long)

    On Error Resume Next

    Dim shp As Shape

    For Each shp In doc.Shapes

        If shp.TextFrame.HasText Then

            Dim r As Range
            Set r = shp.TextFrame.TextRange

            RecolorFieldsInRange r, clr, fieldCount
            RecolorParagraphsInRange r, clr, paragraphCount

        End If

    Next shp

    On Error GoTo 0

End Sub


'---------------------------------------------------------------------
' Safely retrieve paragraph style name
'---------------------------------------------------------------------
Private Function SafeStyleName(para As Paragraph) As String

    On Error Resume Next

    Dim styleName As String
    styleName = CStr(para.Style)

    SafeStyleName = styleName

    Err.Clear
    On Error GoTo 0

End Function


'---------------------------------------------------------------------
' Determine whether a style should be recolored
'---------------------------------------------------------------------
Private Function IsTargetStyle(styleName As String) As Boolean

    IsTargetStyle = False

    If Len(styleName) = 0 Then Exit Function

    Dim lowerName As String
    lowerName = LCase$(Trim$(styleName))

    Dim i As Integer

    ' Heading styles
    For i = 1 To 9

        If lowerName = "heading " & i Then
            IsTargetStyle = True
            Exit Function
        End If

        If lowerName = "heading" & i Then
            IsTargetStyle = True
            Exit Function
        End If

        If lowerName = "heading " & i & " char" Then
            IsTargetStyle = True
            Exit Function
        End If

    Next i

    ' Caption styles
    If lowerName = "caption" Then
        IsTargetStyle = True
        Exit Function
    End If

    If lowerName = "caption char" Then
        IsTargetStyle = True
        Exit Function
    End If

    ' Abstract styles
    If InStr(lowerName, "abstract") > 0 Then
        IsTargetStyle = True
        Exit Function
    End If

    ' Keyword styles
    If InStr(lowerName, "keyword") > 0 Then
        IsTargetStyle = True
        Exit Function
    End If

    If InStr(lowerName, "key words") > 0 Then
        IsTargetStyle = True
        Exit Function
    End If

    ' Table of contents styles
    If lowerName Like "toc#" Or lowerName Like "toc #" Then
        IsTargetStyle = True
        Exit Function
    End If

End Function
