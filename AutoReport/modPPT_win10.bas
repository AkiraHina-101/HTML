'Attribute VB_Name = "modPPT_win10"
Option Explicit

' =============================================================================
' PPT export for Windows 10 / older PowerPoint builds.
' Charts are copied via CopyPicture(xlScreen, xlPicture) + PasteSpecial(EMF),
' matching the result of manual copy-paste from Excel into PowerPoint.
' This avoids SVG rendering issues on Windows 10.
'
' All settings read from ExportConfig sheet via modConfig.
' Entry point: ExportToPPT
' =============================================================================

Private Const ppPasteHTML             As Long = 8
Private Const ppPasteEnhancedMetafile As Long = 2
Private Const msoFalse                As Long = 0
Private Const msoTrue                 As Long = -1
Private Const xlScreen                As Long = 1
Private Const xlPicture               As Long = -4147

' =============================================================================
Public Sub ExportToPPT()
' =============================================================================
    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanFail

    modConfig.InvalidateCache

    Dim boundsName  As String: boundsName = modConfig.CfgStr("SlideBoundsName", "PPT_SlideBounds")
    Dim tableName   As String: tableName = modConfig.CfgStr("DataTableName", "PPT_XL_DataTable")
    Dim chartPfx    As String: chartPfx = modConfig.CfgStr("ChartShapePrefix", "Chart_")
    Dim tblShpName  As String: tblShpName = modConfig.CfgStr("DataTableShapeName", "XL_DataTable")
    Dim tblFont     As String: tblFont = modConfig.CfgStr("DataTableFontName", "")
    Dim tblFontSz   As Double: tblFontSz = modConfig.CfgDbl("DataTableFontSize", 0)
    Dim lblFontSz   As Double: lblFontSz = modConfig.CfgDbl("LabelFontSize", 0)
    Dim lineWtScale As Double: lineWtScale = modConfig.CfgDbl("LineWeightScale", 1)
    If lineWtScale <= 0 Then lineWtScale = 1

    Dim pres As Object: Set pres = OpenPres()
    If pres Is Nothing Then
        Debug.Print "ExportToPPT: cannot open presentation": GoTo CleanExit
    End If
    Debug.Print "=== ExportToPPT_win10 start: " & pres.Name & " (" & pres.Slides.Count & " slides) ==="

    Dim slideW As Double: slideW = pres.PageSetup.SlideWidth
    Dim slideH As Double: slideH = pres.PageSetup.SlideHeight
    Dim cfgWs As Worksheet
    Set cfgWs = modConfig.GetConfigSheet()
    If Not cfgWs Is Nothing Then
        cfgWs.Range("M1").Value2 = "PPT ratio"
        cfgWs.Range("N1").Value2 = slideW / slideH
    End If

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim slideIdx As Long
        slideIdx = SlideIdxFromConfig(ws.Name)
        If slideIdx < 1 Or slideIdx > pres.Slides.Count Then GoTo NextSheet

        Dim bounds As Range
        Set bounds = modLayout.FindNamedRange(ws, boundsName)
        If bounds Is Nothing Then
            Debug.Print "  [SKIP] " & ws.Name & ": " & boundsName & " missing"
            GoTo NextSheet
        End If
        Debug.Print "--- " & ws.Name & " -> slide " & slideIdx & " ---"

        Dim sld As Object: Set sld = pres.Slides(slideIdx)

        On Error Resume Next
        pres.Application.ActiveWindow.View.GotoSlide slideIdx
        On Error GoTo CleanFail

        ' Export all DataTable ranges
        Dim nm      As Name
        Dim nmLocal As String
        Dim dtRng   As Range
        For Each nm In ThisWorkbook.Names
            nmLocal = nm.Name
            If InStr(nmLocal, "!") > 0 Then nmLocal = Mid$(nmLocal, InStr(nmLocal, "!") + 1)
            If Left$(nmLocal, Len(tableName)) = tableName Then
                Set dtRng = Nothing
                On Error Resume Next
                Set dtRng = nm.RefersToRange
                On Error GoTo CleanFail
                If Not dtRng Is Nothing Then
                    If StrComp(dtRng.Parent.Name, ws.Name, vbTextCompare) = 0 Then
                        ExportTable dtRng, sld, bounds, nmLocal, slideW, slideH
                    End If
                End If
            End If
        Next nm

        ' Sync all tables to match _1: total W, column widths, row heights
        SyncTableColumns sld, tableName

        Dim co As ChartObject
        For Each co In ws.ChartObjects
            ExportChart co, sld, bounds, chartPfx, slideW, slideH
        Next co

        Dim lineShp As Shape
        For Each lineShp In ws.Shapes
            If Left$(lineShp.Name, 5) = "Line_" Then
                ExportLineShape lineShp, sld, bounds, slideW, slideH, lineWtScale
            End If
        Next lineShp

        Dim labelShp As Shape
        For Each labelShp In ws.Shapes
            If Left$(labelShp.Name, 9) = "LabelOut_" Or Left$(labelShp.Name, 10) = "PPT_Label_" Then
                ExportLabelShape labelShp, sld, bounds, slideW, slideH, tblFont, lblFontSz
            End If
        Next labelShp

NextSheet:
    Next ws
    Debug.Print "=== ExportToPPT_win10 done ==="

CleanExit:
    Application.EnableEvents = prevEvents
    Exit Sub
CleanFail:
    Debug.Print "[ERROR] " & ws.Name & ": " & Err.Number & " - " & Err.Description
    Resume NextSheet
End Sub

' --- Sync all DataTable column widths to match _1 ----------------------------
Private Sub SyncTableColumns(ByVal sld As Object, ByVal tablePrefix As String)
    ' Find reference shape (tableName & "_1")
    Dim refShp As Object
    On Error Resume Next
    Set refShp = sld.Shapes(tablePrefix & "_1")
    On Error GoTo 0
    If refShp Is Nothing Then Exit Sub

    Dim refTbl As Object
    On Error Resume Next
    Set refTbl = refShp.Table
    On Error GoTo 0
    If refTbl Is Nothing Then Exit Sub

    ' Read reference column widths and row heights
    Dim nCols As Long: nCols = refTbl.Columns.Count
    Dim nRows As Long: nRows = refTbl.Rows.Count
    Dim refColW() As Double: ReDim refColW(1 To nCols)
    Dim refRowH() As Double: ReDim refRowH(1 To nRows)
    Dim refTotalW As Double: refTotalW = 0
    Dim c As Long, r As Long
    For c = 1 To nCols
        refColW(c) = refTbl.Columns(c).Width
        refTotalW = refTotalW + refColW(c)
    Next c
    For r = 1 To nRows
        refRowH(r) = refTbl.Rows(r).Height
    Next r

    ' Apply to all other matching shapes: set W, col widths, row heights
    Dim i As Long
    For i = 1 To sld.Shapes.Count
        Dim shp As Object: Set shp = sld.Shapes(i)
        If Left$(shp.Name, Len(tablePrefix)) = tablePrefix And _
           shp.Name <> tablePrefix & "_1" Then
            On Error Resume Next
            Dim tbl As Object: Set tbl = shp.Table
            If Not tbl Is Nothing Then
                If tbl.Columns.Count = nCols Then
                    ' Set total width first
                    shp.Width = refTotalW
                    ' Sync column widths (absolute)
                    For c = 1 To nCols
                        tbl.Columns(c).Width = refColW(c)
                    Next c
                End If
                ' Sync row heights (absolute) if row count matches
                If tbl.Rows.Count = nRows Then
                    For r = 1 To nRows
                        tbl.Rows(r).Height = refRowH(r)
                    Next r
                End If
                Debug.Print "  [SYNC] " & shp.Name
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

' --- DataTable ----------------------------------------------------------------
Private Sub ExportTable(ByVal rng As Range, ByVal sld As Object, _
                         ByVal bounds As Range, ByVal shapeName As String, _
                         ByVal slideW As Double, ByVal slideH As Double, _
                         Optional ByVal sharedW As Double = 0)
    DeleteByName sld, shapeName
    rng.Copy

    Dim shp As Object
    On Error Resume Next
    Err.Clear
    Set shp = sld.Shapes.PasteSpecial(DataType:=ppPasteHTML)
    On Error GoTo 0
    Application.CutCopyMode = False
    If shp Is Nothing Then Debug.Print "  [ERR] DataTable paste failed": Exit Sub

    Dim scaleX As Double: scaleX = slideW / bounds.Width
    Dim scaleY As Double: scaleY = slideH / bounds.Height
    Dim pptL   As Double: pptL = (rng.Left - bounds.Left) * scaleX
    Dim pptT   As Double: pptT = (rng.Top  - bounds.Top)  * scaleY
    Dim pptW   As Double: pptW = rng.Width  * scaleX
    Dim pptH   As Double: pptH = rng.Height * scaleY

    shp.Name            = shapeName
    shp.LockAspectRatio = msoFalse
    shp.Left            = pptL
    shp.Top             = pptT
    shp.Width = pptW:   shp.Height = pptH
    shp.Width = pptW:   shp.Height = pptH

    Debug.Print "  [OK] " & shapeName & " L=" & Pt(pptL) & " T=" & Pt(pptT) & _
                " W=" & Pt(pptW) & " H=" & Pt(pptH)
End Sub

' --- Charts -------------------------------------------------------------------
Private Sub ExportChart(ByVal co As ChartObject, ByVal sld As Object, _
                         ByVal bounds As Range, ByVal prefix As String, _
                         ByVal slideW As Double, ByVal slideH As Double)
    Dim sName As String
    sName = prefix & co.Name
    DeleteByName sld, sName

    Dim scaleX As Double: scaleX = slideW / bounds.Width
    Dim scaleY As Double: scaleY = slideH / bounds.Height
    Dim pptL As Double: pptL = (co.Left - bounds.Left) * scaleX
    Dim pptT As Double: pptT = (co.Top  - bounds.Top)  * scaleY
    Dim pptW As Double: pptW = co.Width  * scaleX
    Dim pptH As Double: pptH = co.Height * scaleY

    Dim pptShp As Object
    Dim pasteErr As Long
    Dim attempt  As Long
    On Error Resume Next
    For attempt = 1 To 3
        co.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        DoEvents
        Err.Clear
        Set pptShp = sld.Shapes.PasteSpecial(DataType:=ppPasteEnhancedMetafile)
        pasteErr = Err.Number
        Application.CutCopyMode = False
        If pasteErr = 0 And Not pptShp Is Nothing Then Exit For
        Debug.Print "  [RETRY " & attempt & "] " & co.Name & " err=" & pasteErr
        DoEvents
    Next attempt
    On Error GoTo 0

    If pasteErr <> 0 Or pptShp Is Nothing Then
        Debug.Print "  [ERR] ExportChart failed after 3 attempts: " & co.Name
        Exit Sub
    End If

    On Error Resume Next
    Set pptShp = FirstShapeFromPaste(pptShp)
    pptShp.LockAspectRatio = msoFalse
    pptShp.Left   = pptL
    pptShp.Top    = pptT
    pptShp.Width  = pptW
    pptShp.Height = pptH
    pptShp.Name   = sName
    On Error GoTo 0

    Debug.Print "  [OK] " & sName & " L=" & Pt(pptL) & " T=" & Pt(pptT) & _
                " W=" & Pt(pptW) & " H=" & Pt(pptH)
End Sub

Private Function FirstShapeFromPaste(ByVal pasted As Object) As Object
    On Error Resume Next
    Set FirstShapeFromPaste = pasted.Item(1)
    If FirstShapeFromPaste Is Nothing Then Set FirstShapeFromPaste = pasted
    On Error GoTo 0
End Function

' --- Line shapes (Line_ prefix) -----------------------------------------------
' Type=9 (msoLine): AddConnector in PPT — preserves color/dash/weight + lineWtScale.
' Oval markers    : AddShape (exact fill color, no outline).
Private Sub ExportLineShape(ByVal xlShp As Shape, ByVal sld As Object, _
                             ByVal bounds As Range, ByVal slideW As Double, _
                             ByVal slideH As Double, ByVal lineWtScale As Double)
    DeleteByName sld, xlShp.Name

    Dim scaleX As Double: scaleX = slideW / bounds.Width
    Dim scaleY As Double: scaleY = slideH / bounds.Height
    Dim pptL   As Double: pptL = (xlShp.Left - bounds.Left) * scaleX
    Dim pptT   As Double: pptT = (xlShp.Top  - bounds.Top)  * scaleY
    Dim pptW   As Double: pptW = xlShp.Width  * scaleX
    Dim pptH   As Double: pptH = xlShp.Height * scaleY

    Dim pptShp As Object
    On Error Resume Next
    Err.Clear

    If xlShp.Type = 9 Then
        ' msoConnectorStraight = 1 — vertical line from (pptL, pptT) to (pptL, pptT+pptH)
        Set pptShp = sld.Shapes.AddConnector(1, pptL, pptT, pptL, pptT + pptH)
        If Err.Number = 0 And Not pptShp Is Nothing Then
            With pptShp.Line
                .ForeColor.RGB = xlShp.Line.ForeColor.RGB
                .Weight        = xlShp.Line.Weight * lineWtScale
                .DashStyle     = XlDashToMso(xlShp.Line.DashStyle)
            End With
            pptShp.Name = xlShp.Name
        End If
        On Error GoTo 0
    Else
        ' Oval marker — AddShape (exact fill, no outline)
        Set pptShp = sld.Shapes.AddShape(9, pptL, pptT, pptW, pptH)
        If Err.Number = 0 And Not pptShp Is Nothing Then
            pptShp.Line.Visible       = msoFalse
            pptShp.Fill.ForeColor.RGB = xlShp.Fill.ForeColor.RGB
            pptShp.Name               = xlShp.Name
        End If
        On Error GoTo 0
    End If

    Debug.Print "  [OK] " & xlShp.Name & " L=" & Pt(pptL) & " T=" & Pt(pptT) & _
                " W=" & Pt(pptW) & " H=" & Pt(pptH)
End Sub

' Map Excel XlLineStyle dash constants → PPT MsoDashStyle constants
Private Function XlDashToMso(ByVal xlDash As Long) As Long
    Select Case xlDash
        Case 1       ' xlSolid
            XlDashToMso = 1   ' msoLineSolid
        Case -4142   ' xlDot
            XlDashToMso = 2   ' msoLineDot
        Case -4115   ' xlDash
            XlDashToMso = 4   ' msoLineDash
        Case 4       ' xlDashDot
            XlDashToMso = 5   ' msoLineDashDot
        Case 5       ' xlDashDotDot
            XlDashToMso = 6   ' msoLineDashDotDot
        Case Else
            XlDashToMso = 4   ' fallback: dash
    End Select
End Function

Private Function OpenPres() As Object
    Dim cfgPath As String: cfgPath = modConfig.CfgStr("PptxPath", "")
    If Len(cfgPath) = 0 Then
        Debug.Print "OpenPres: PptxPath not configured": Exit Function
    End If

    Dim pptxPath As String
    If Mid$(cfgPath, 2, 1) = ":" Or Left$(cfgPath, 2) = "\\" Then
        pptxPath = cfgPath
    Else
        pptxPath = ThisWorkbook.Path & "\" & cfgPath
    End If

    Dim pptApp As Object
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application")
        pptApp.Visible = True
    End If

    Dim p As Object
    For Each p In pptApp.Presentations
        If StrComp(p.FullName, pptxPath, vbTextCompare) = 0 Then
            Set OpenPres = p: Exit Function
        End If
    Next p
    Set OpenPres = pptApp.Presentations.Open(pptxPath)
End Function

Private Sub DeleteByName(ByVal sld As Object, ByVal sName As String)
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        If StrComp(sld.Shapes(i).Name, sName, vbTextCompare) = 0 Then
            sld.Shapes(i).Delete
        End If
    Next i
End Sub

' --- Label shapes (LabelOut_ prefix) ------------------------------------------
Private Sub ExportLabelShape(ByVal xlShp As Shape, ByVal sld As Object, _
                              ByVal bounds As Range, ByVal slideW As Double, _
                              ByVal slideH As Double, ByVal fontName As String, _
                              ByVal fontSize As Double)
    DeleteByName sld, xlShp.Name

    Dim scaleX As Double: scaleX = slideW / bounds.Width
    Dim scaleY As Double: scaleY = slideH / bounds.Height
    Dim pptL   As Double: pptL = (xlShp.Left - bounds.Left) * scaleX
    Dim pptT   As Double: pptT = (xlShp.Top  - bounds.Top)  * scaleY
    Dim pptW   As Double: pptW = xlShp.Width  * scaleX
    Dim pptH   As Double: pptH = xlShp.Height * scaleY

    ' ── Tạo native PPT TextBox (editable) ────────────────────
    Dim pptShp As Object
    On Error Resume Next
    Set pptShp = sld.Shapes.AddTextBox(1, pptL, pptT, pptW, pptH)
    On Error GoTo 0

    If pptShp Is Nothing Then
        Debug.Print "  [ERR] ExportLabelShape AddTextBox '" & xlShp.Name & "' failed"
        Exit Sub
    End If

    On Error Resume Next

    Dim xlTf  As TextFrame: Set xlTf = xlShp.TextFrame
    Dim pptTf As Object:    Set pptTf = pptShp.TextFrame

    ' ── Text + font (1 lần cho toàn bộ TextRange) ───────────
    pptTf.TextRange.Text = xlTf.Characters.Text

    Dim xlFont As Font: Set xlFont = xlTf.Characters.Font
    With pptTf.TextRange.Font
        If Len(fontName) > 0 Then .Name = fontName Else .Name = xlFont.Name
        .Size      = xlFont.Size * scaleY
        .Color.RGB = xlFont.Color
        .Bold      = (xlFont.Bold = True)
        .Italic    = (xlFont.Italic = True)
        .Underline = (xlFont.Underline <> xlUnderlineStyleNone)
    End With

    ' ── Paragraph alignment (H) ──────────────────────────────
    ' Excel: xlLeft=-4131, xlCenter=-4108, xlRight=-4152, xlJustify=-4130
    ' PPT  : ppAlignLeft=1, ppAlignCenter=2, ppAlignRight=3, ppAlignJustify=4
    Dim hAlign As Long
    Select Case xlTf.HorizontalAlignment
        Case -4108: hAlign = 2  ' xlCenter → ppAlignCenter
        Case -4152: hAlign = 3  ' xlRight  → ppAlignRight
        Case -4130: hAlign = 4  ' xlJustify → ppAlignJustify
        Case Else:  hAlign = 1  ' xlLeft   → ppAlignLeft
    End Select
    pptTf.TextRange.ParagraphFormat.Alignment = hAlign

    ' ── Vertical anchor ──────────────────────────────────────
    ' Excel: xlTop=-4160, xlCenter=-4108, xlBottom=-4107
    ' PPT  : msoAnchorTop=1, msoAnchorMiddle=3, msoAnchorBottom=4
    Dim vAlign As Long
    Select Case xlTf.VerticalAlignment
        Case -4108: vAlign = 3  ' xlCenter → msoAnchorMiddle
        Case -4107: vAlign = 4  ' xlBottom → msoAnchorBottom
        Case Else:  vAlign = 1  ' xlTop    → msoAnchorTop
    End Select
    pptTf.VerticalAnchor = vAlign

    ' ── Word wrap ────────────────────────────────────────────
    pptTf.WordWrap = xlTf.WordWrap

    ' ── Margins ──────────────────────────────────────────────
    pptTf.MarginLeft   = xlTf.MarginLeft
    pptTf.MarginRight  = xlTf.MarginRight
    pptTf.MarginTop    = xlTf.MarginTop
    pptTf.MarginBottom = xlTf.MarginBottom

    ' ── Fill ─────────────────────────────────────────────────
    If xlShp.Fill.Visible = msoTrue Then
        pptShp.Fill.Solid
        pptShp.Fill.ForeColor.RGB = xlShp.Fill.ForeColor.RGB
        pptShp.Fill.Transparency  = xlShp.Fill.Transparency
    Else
        pptShp.Fill.Visible = msoFalse
    End If

    ' ── Border (Line) ─────────────────────────────────────────
    If xlShp.Line.Visible = msoTrue Then
        pptShp.Line.Visible        = msoTrue
        pptShp.Line.ForeColor.RGB  = xlShp.Line.ForeColor.RGB
        pptShp.Line.Weight         = xlShp.Line.Weight
        pptShp.Line.DashStyle      = XlDashToMso(xlShp.Line.DashStyle)
    Else
        pptShp.Line.Visible = msoFalse
    End If

    pptShp.Name = xlShp.Name
    On Error GoTo 0

    Debug.Print "  [OK] " & xlShp.Name & " L=" & Pt(pptL) & " T=" & Pt(pptT) & _
                " W=" & Pt(pptW) & " H=" & Pt(pptH)
End Sub

Private Function Pt(ByVal v As Double) As String
    Pt = Format$(v, "0.0")
End Function

' --- Create PPT_XL_DataTable_N named ranges via repeated InputBox ---
Public Sub CreateDataTableRanges()
    Dim prefix As String: prefix = "PPT_XL_DataTable_"
    Dim idx    As Long:   idx = 1

    Do
        Dim rng As Range
        Set rng = Nothing
        On Error Resume Next
        Set rng = Application.InputBox( _
            Prompt:="Select range for " & prefix & idx & vbCrLf & "(Cancel to stop)", _
            Title:="Create DataTable " & idx, _
            Type:=8)
        Dim cancelErr As Long: cancelErr = Err.Number
        On Error GoTo 0

        If cancelErr <> 0 Or rng Is Nothing Then Exit Do

        Dim nmName As String: nmName = prefix & idx
        ThisWorkbook.Names.Add Name:=nmName, RefersTo:=rng
        Debug.Print "[OK] " & nmName & " = " & rng.Address(External:=True)

        idx = idx + 1
    Loop

    MsgBox "Created " & (idx - 1) & " named range(s)." & vbCrLf & _
           prefix & "1 to " & prefix & (idx - 1), vbInformation, "Done"
End Sub

Private Function SlideIdxFromConfig(ByVal sheetName As String) As Long
    Dim cfgWs As Worksheet
    Set cfgWs = modConfig.GetConfigSheet()
    If cfgWs Is Nothing Then Exit Function

    Dim colSlide As Long: colSlide = modConfig.FindHeaderCol(cfgWs, modConfig.HDR_SLIDE)
    Dim colSheet As Long: colSheet = modConfig.FindHeaderCol(cfgWs, modConfig.HDR_SHEET)
    Dim colGroup As Long: colGroup = modConfig.FindHeaderCol(cfgWs, modConfig.HDR_GROUP)
    If colSlide = 0 Or colSheet = 0 Then
        Debug.Print "SlideIdxFromConfig: missing header '" & modConfig.HDR_SLIDE & _
                    "' or '" & modConfig.HDR_SHEET & "'"
        Exit Function
    End If

    Dim r As Long
    For r = modConfig.HDR_ROW + 1 To modConfig.CFG_MAX_ROW
        Dim aVal As Variant: aVal = cfgWs.Cells(r, colSlide).Value2
        Dim bVal As Variant: bVal = cfgWs.Cells(r, colSheet).Value2
        Dim cVal As Variant
        If colGroup > 0 Then cVal = cfgWs.Cells(r, colGroup).Value2 Else cVal = ""
        If (IsEmpty(aVal) Or Len(Trim$(CStr(aVal))) = 0) And _
           (IsEmpty(bVal) Or Len(Trim$(CStr(bVal))) = 0) And _
           (IsEmpty(cVal) Or Len(Trim$(CStr(cVal))) = 0) Then Exit For
        If Not IsEmpty(bVal) Then
            If StrComp(Trim$(CStr(bVal)), sheetName, vbTextCompare) = 0 Then
                If IsNumeric(aVal) Then SlideIdxFromConfig = CLng(aVal)
                Exit Function
            End If
        End If
    Next r
End Function
