Option Explicit
' Self-test for the VDOT LRS tool. Not needed to use the tool.
'
'   VBA RUN [VDOT_LRS_Linear_Milepost]modLrsSelfTest.LrsSelfTest
'
' Drives the real dialog and writes what happened to %TEMP%\vdot_lrs_vba_selftest.txt.
' It only creates events when the active file is the scratch copy lrs_vba_scratch.dgn and
' %TEMP%\lrs_selftest_create.flag exists.

Private mHandle As Integer

Private Sub Out(ByVal key As String, ByVal value As Variant)
    Print #mHandle, key & "=" & Replace(Replace(CStr(value), vbCrLf, " | "), vbLf, " | ")
End Sub

Public Sub LrsSelfTest()
    Dim path As String, i As Long, n As Long, names As String, x As Double, y As Double
    Dim gx As Double, gy As Double, ex As Double, ey As Double, before As Long, doCreate As Boolean
    path = Environ$("TEMP") & "\vdot_lrs_vba_selftest.txt"
    mHandle = FreeFile
    Open path For Output As #mHandle
    On Error GoTo Failed
    Out "started", Now
    ' Events are only created in the scratch copy, and only when the flag file exists.
    doCreate = (Dir$(Environ$("TEMP") & "\lrs_selftest_create.flag") <> "") And _
               (LCase$(ActiveDesignFile.Name) = "lrs_vba_scratch.dgn")
    Out "file", ActiveDesignFile.Name
    Out "model3d", Model3D()
    Out "modelGcs", ModelGcsName()

    ' Pure helpers
    Out "fmt", FmtMeasure(-0.0004) & " " & FmtMeasure(15.8096) & " " & FmtThousands(24753.66, 1)
    Dim digits As String, decimalText As String
    MilepostText 11.5, "0.5", digits, decimalText: Out "milepost11.5", digits & decimalText
    MilepostText 11#, "1", digits, decimalText: Out "milepost11", digits & "[" & decimalText & "]"
    MilepostText 15.81, MILEPOST_ENDS, digits, decimalText: Out "milepost15.81", digits & decimalText

    ' Box against segment, and a label beside a road that turns sharply (east, then north)
    Out "segmentBox", SegmentHitsBox(0, 0, 10, 10, 4, 4, 6, 6) & " " & SegmentHitsBox(0, 0, 10, 0, 4, 4, 6, 6) & " " & _
        SegmentHitsBox(0, 7, 7, 0, 4, 4, 6, 6) & " " & SegmentHitsBox(0, 9, 9, 0, 4, 4, 6, 6)
    Dim bend As clsLrsGeometry, bx(0 To 2) As Double, by(0 To 2) As Double, bm(0 To 2) As Double
    bx(0) = 0: by(0) = 0: bm(0) = 0: bx(1) = 100: by(1) = 0: bm(1) = 1: bx(2) = 100: by(2) = 100: bm(2) = 2
    Set bend = New clsLrsGeometry
    bend.AddPath bx, by, bm, 3
    bend.Finish
    Out "bendTouches", bend.TouchesBox(-1, 3, 50, 2, 110, 6) & " " & bend.TouchesBox(-1, 3, 50, -6, 110, -2) & _
        " rate " & FmtNum(bend.MeasurePerUnit(), 3)

    ' Coordinate conversion: the macro's own EPSG:2284 against the model GCS (when it has one)
    If ModelGcsName() <> "" Then
        InitProjector MODE_MODEL_GCS, 0, 0, 0
        ProjectPoint -82.2636, 36.6359, gx, gy
        InitProjector MODE_EPSG, 2284, 0, 0
        ProjectPoint -82.2636, 36.6359, ex, ey
        Out "gcsPoint", FmtNum(gx, 3) & ", " & FmtNum(gy, 3)
        Out "epsg2284Point", FmtNum(ex, 3) & ", " & FmtNum(ey, 3)
        Out "difference", FmtNum(Hypot(gx - ex, gy - ey), 4)
    End If
    InitProjector MODE_EPSG, 2283, 0, 0
    ProjectPoint -77.3, 38.85, ex, ey: Out "epsg2283Point", FmtNum(ex, 3) & ", " & FmtNum(ey, 3)
    InitProjector MODE_EPSG, 3968, 0, 0
    ProjectPoint -77.3, 38.85, ex, ey: Out "epsg3968Point", FmtNum(ex, 3) & ", " & FmtNum(ey, 3)

    ' The dialog
    frmLrsEvent.NoZoom = True
    frmLrsEvent.Show vbModeless
    DoEvents
    Out "formSize", frmLrsEvent.TestValue("width") & " x " & frmLrsEvent.TestValue("height")
    Out "logoShown", frmLrsEvent.TestValue("logoShown")
    Out "status", frmLrsEvent.TestValue("status")
    Out "jurisdictions", frmLrsEvent.TestValue("jurisdictions")
    Out "firstJurisdictions", frmLrsEvent.TestValue("jurisdiction", 0) & "; " & frmLrsEvent.TestValue("jurisdiction", 1)
    Out "modelHint", frmLrsEvent.TestValue("modelHint")
    Out "coords", frmLrsEvent.TestValue("coords") & " (" & frmLrsEvent.TestValue("coordsCount") & " choices)"
    Out "routeHintOnOpen", frmLrsEvent.TestValue("routeHint")

    Out "choseJurisdiction", frmLrsEvent.TestChoose("jurisdiction", "Washington County")
    n = frmLrsEvent.TestValue("routes")
    Out "routes", n
    For i = 0 To 19
        If i < n Then names = names & IIf(i = 0, "", "; ") & frmLrsEvent.TestValue("route", i)
    Next i
    Out "first20", names
    Out "routeBlank", (frmLrsEvent.TestValue("routeText") = "")

    Out "choseI81", frmLrsEvent.TestChoose("route", "I-81N")
    Out "i81Hint", frmLrsEvent.TestValue("routeHint")
    Out "i81Measures", frmLrsEvent.TestValue("from") & " to " & frmLrsEvent.TestValue("to")
    Out "i81Preview", frmLrsEvent.TestValue("routePreview") & " route, " & frmLrsEvent.TestValue("eventPreview") & " event"
    Out "toolActiveAfterRoute", gLrsToolActive

    Out "choseSC700", frmLrsEvent.TestChoose("route", "SC-700E (Washington County)")
    Out "sc700Hint", frmLrsEvent.TestValue("routeHint")
    Out "sc700Measures", frmLrsEvent.TestValue("from") & " to " & frmLrsEvent.TestValue("to")
    Out "sc700From", frmLrsEvent.TestValue("fromLocation")
    Out "sc700To", frmLrsEvent.TestValue("toLocation")
    Out "sc700Event", frmLrsEvent.TestValue("eventHint")
    Out "converted", frmLrsEvent.TestValue("converted")

    ' Picking: From at MP 12.6, then To behind it at MP 5.14 (must swap)
    frmLrsEvent.TestLocate 12.6, x, y: frmLrsEvent.TestPick x, y
    Out "pick1", frmLrsEvent.TestValue("from") & " to " & frmLrsEvent.TestValue("to") & " | " & frmLrsEvent.PickPrompt()
    frmLrsEvent.TestLocate 5.14, x, y: frmLrsEvent.TestPick x, y
    Out "pick2", frmLrsEvent.TestValue("from") & " to " & frmLrsEvent.TestValue("to") & " | " & frmLrsEvent.TestValue("eventHint")
    Out "offer", frmLrsEvent.TestLastOffer() & " | height " & frmLrsEvent.TestValue("height")
    frmLrsEvent.TestOffer False
    Out "offerNo", LastLogLine() & " | bar hidden " & (frmLrsEvent.TestLastOffer() = "") & _
        " | height " & frmLrsEvent.TestValue("height")

    ' Validation
    n = frmLrsEvent.TestValue("logLines")
    frmLrsEvent.TestSet "to", "abc"
    Out "badMeasure", frmLrsEvent.TestValue("eventHint") & " | error " & frmLrsEvent.TestValue("eventHintError")
    frmLrsEvent.TestSet "to", "50"
    Out "outsideRoute", frmLrsEvent.TestValue("eventHint") & " | logged " & (frmLrsEvent.TestValue("logLines") - n)
    ' Swap, Full route, Linear / Point, pick indicator, opposite direction, log
    frmLrsEvent.TestSet "from", "9": frmLrsEvent.TestSet "to", "4": frmLrsEvent.TestClick "swap"
    Out "swap", frmLrsEvent.TestValue("from") & " to " & frmLrsEvent.TestValue("to") & " | error " & frmLrsEvent.TestValue("eventHintError")
    frmLrsEvent.TestClick "full"
    Out "full", frmLrsEvent.TestValue("from") & " to " & frmLrsEvent.TestValue("to") & " | " & frmLrsEvent.TestValue("eventHint")
    frmLrsEvent.TestSet "type", EVENT_POINT: Out "pointHidesTo", Not frmLrsEvent.TestValue("toVisible")
    frmLrsEvent.TestSet "type", EVENT_LINEAR: Out "linearShowsTo", frmLrsEvent.TestValue("toVisible")
    Out "pickActive", gLrsToolActive & " | " & frmLrsEvent.TestValue("pickCaption") & " | from " & _
        frmLrsEvent.TestValue("fromLabel") & " | to " & frmLrsEvent.TestValue("toLabel")
    frmLrsEvent.TestClick "pick"
    Out "pickStopped", gLrsToolActive & " | " & frmLrsEvent.TestValue("pickCaption") & " | from " & frmLrsEvent.TestValue("fromLabel")
    frmLrsEvent.TestClick "pick"
    Out "pickRestarted", gLrsToolActive
    Out "sc700Opposite", frmLrsEvent.TestValue("oppositeEnabled")
    Out "choseI81Again", frmLrsEvent.TestChoose("route", "I-81N") & " | opposite " & frmLrsEvent.TestValue("oppositeEnabled")
    frmLrsEvent.TestClick "opposite"
    Out "oppositeLoaded", frmLrsEvent.TestValue("routeId") & " | " & frmLrsEvent.TestValue("routeText")
    x = Timer: frmLrsEvent.TestClick "opposite"
    Out "backCached", frmLrsEvent.TestValue("routeId") & " in " & FmtNum(Timer - x, 2) & " s"
    x = Timer: frmLrsEvent.TestChoose "jurisdiction", "Washington County"
    Out "routesCached", frmLrsEvent.TestValue("routes") & " in " & FmtNum(Timer - x, 2) & " s"
    frmLrsEvent.TestChoose "route", "SC-700E (Washington County)"
    frmLrsEvent.TestLog "WARNING: sample warning": frmLrsEvent.TestLog "ERROR: sample error"
    frmLrsEvent.TestClick "log": Out "logClosed", (Not CBool(frmLrsEvent.TestValue("logOpen"))) & " height " & frmLrsEvent.TestValue("height")
    frmLrsEvent.TestClick "log": Out "logOpen", frmLrsEvent.TestValue("logOpen") & " height " & frmLrsEvent.TestValue("height")
    frmLrsEvent.TestSet "from", "10": frmLrsEvent.TestSet "to", "20"
    Out "gapEvent", frmLrsEvent.TestValue("eventHint") & " | preview " & frmLrsEvent.TestValue("eventPreview")
    frmLrsEvent.TestSet "type", EVENT_POINT: frmLrsEvent.TestSet "from", "3.25"
    Out "pointEvent", frmLrsEvent.TestValue("eventHint") & " | " & frmLrsEvent.TestValue("fromLocation") & " | preview " & frmLrsEvent.TestValue("eventPreview")
    frmLrsEvent.TestSet "type", EVENT_LINEAR

    ' Other coordinate choices
    For i = 1 To CLng(frmLrsEvent.TestValue("coordsCount")) - 1
        frmLrsEvent.TestSet "coords", CStr(i)
        frmLrsEvent.TestLocate 0#, x, y
        Out "coords" & i, frmLrsEvent.TestValue("coords") & " -> " & FmtNum(x, 2) & ", " & FmtNum(y, 2)
    Next i
    frmLrsEvent.TestSet "coords", "0"

    If doCreate Then
        before = ElementCount()
        frmLrsEvent.TestSet "from", "2": frmLrsEvent.TestSet "to", "5"
        frmLrsEvent.TestSet "mileposts", "1"            ' At From and To
        frmLrsEvent.TestCreate
        Out "createShort", LastLogLine() & " | elements added " & (ElementCount() - before)
        before = ElementCount()
        frmLrsEvent.TestSet "from", "14": frmLrsEvent.TestSet "to", "17"      ' crosses the gap at 15.81
        frmLrsEvent.TestSet "mileposts", "3"            ' Every 0.5 mile
        frmLrsEvent.TestCreate
        Out "createGap", LastLogLine() & " | elements added " & (ElementCount() - before)
        before = ElementCount()
        frmLrsEvent.TestSet "type", EVENT_POINT: frmLrsEvent.TestSet "from", "3.25"
        frmLrsEvent.TestSet "mileposts", "0"
        frmLrsEvent.TestCreate
        Out "createPoint", LastLogLine() & " | elements added " & (ElementCount() - before)
        Out "itemsOnLast", LastItemSummary()
        before = ElementCount()
        frmLrsEvent.TestLocate 7, x, y: frmLrsEvent.TestPick x, y        ' Point mode: one pick
        Out "offerPoint", frmLrsEvent.TestLastOffer()
        frmLrsEvent.TestOffer True
        Out "offerYes", LastLogLine() & " | elements added " & (ElementCount() - before) & _
            " | bar hidden " & (frmLrsEvent.TestLastOffer() = "")
        Dim summary As String
        Out "foundCreated", FindCreatedElements(summary).count & " | " & summary
        before = ElementCount()
        frmLrsEvent.TestDelete
        Out "deleteOffer", frmLrsEvent.TestLastOffer() & " | elements removed " & (before - ElementCount())
        frmLrsEvent.TestOffer False
        Out "deleteNo", LastLogLine() & " | elements removed " & (before - ElementCount())
        frmLrsEvent.TestDelete
        frmLrsEvent.TestOffer True
        Out "deleteCreated", LastLogLine() & " | elements removed " & (before - ElementCount())
        summary = ""
        Out "foundAfterDelete", FindCreatedElements(summary).count
    End If

    Out "log", frmLrsEvent.TestValue("log")
    Unload frmLrsEvent
    DoEvents
    Out "toolActiveAfterClose", gLrsToolActive
    Out "ok", True
    GoTo Finish
Failed:
    Out "error", Err.Number & " " & Err.Description
    Out "ok", False
Finish:
    Out "finished", Now
    Close #mHandle
End Sub

Public Sub LrsDemoDark()
    LrsDemo "0"
End Sub

Public Sub LrsDemoLight()
    LrsDemo "1"
End Sub

Private Sub LrsDemo(ByVal themeIndex As String)
    ' Opens the dialog with SC-700E loaded, for a screenshot.
    frmLrsEvent.NoZoom = False
    frmLrsEvent.Show vbModeless
    DoEvents
    frmLrsEvent.TestSet "theme", themeIndex
    frmLrsEvent.TestChoose "jurisdiction", "Washington County"
    frmLrsEvent.TestChoose "route", "SC-700E (Washington County)"
End Sub

Public Sub LrsClose()
    Unload frmLrsEvent
End Sub

Public Sub LrsOpenBlank()
    frmLrsEvent.NoZoom = False
    frmLrsEvent.Show vbModeless
    frmLrsEvent.TestSet "theme", "0"
End Sub

Public Sub LrsFocusJurisdiction()
    frmLrsEvent.TestFocus "jurisdiction"
End Sub

Public Sub LrsFocusRoute()
    frmLrsEvent.TestFocus "route"
End Sub

Public Sub LrsProbe()
    ' Writes the state of the dialog to %TEMP%\vdot_lrs_vba_probe.txt.
    mHandle = FreeFile
    Open Environ$("TEMP") & "\vdot_lrs_vba_probe.txt" For Output As #mHandle
    On Error Resume Next
    Out "jurisdictionText", frmLrsEvent.TestValue("jurisdictionText")
    Out "jurisdictionList", frmLrsEvent.TestValue("jurisdictionList")
    Out "routeText", frmLrsEvent.TestValue("routeText")
    Out "routeList", frmLrsEvent.TestValue("routeList")
    Out "routes", frmLrsEvent.TestValue("routes")
    Out "loaded", frmLrsEvent.TestValue("loaded")
    Out "routeHint", frmLrsEvent.TestValue("routeHint")
    Out "measures", frmLrsEvent.TestValue("from") & " to " & frmLrsEvent.TestValue("to")
    Out "eventHint", frmLrsEvent.TestValue("eventHint")
    Out "toolActive", gLrsToolActive
    Out "logGeometry", frmLrsEvent.TestValue("logGeometry")
    Out "log", frmLrsEvent.TestValue("log")
    Close #mHandle
End Sub

Public Sub LrsInspect()
    ' Lists the text elements and cells of the active model in %TEMP%\vdot_lrs_vba_probe.txt.
    Dim scan As ElementEnumerator, item As Element, parts As ElementEnumerator, heights As String, n As Long
    mHandle = FreeFile
    Open Environ$("TEMP") & "\vdot_lrs_vba_probe.txt" For Output As #mHandle
    Out "activeTextHeight", ActiveSettings.TextStyle.height
    Set scan = ActiveModelReference.GraphicalElementCache.Scan()
    Do While scan.MoveNext
        Set item = scan.Current
        If item.IsTextElement Then
            Out DLongToString(item.ID), "text '" & item.AsTextElement.text & "' height " & _
                FmtNum(item.AsTextElement.TextStyle.height, 1) & " at " & FmtNum(item.AsTextElement.origin.x, 0) & _
                ", " & FmtNum(item.AsTextElement.origin.y, 0)
        ElseIf item.IsCellElement Then
            heights = "": n = 0
            Set parts = item.AsCellElement.GetSubElements
            Do While parts.MoveNext
                n = n + 1
                If parts.Current.IsTextElement Then
                    heights = heights & " " & parts.Current.AsTextElement.text & ":" & _
                              FmtNum(parts.Current.AsTextElement.TextStyle.height, 1)
                End If
            Loop
            Out DLongToString(item.ID), "cell " & item.AsCellElement.Name & " (" & n & " parts)" & heights
        End If
    Loop
    Close #mHandle
End Sub

Public Sub LrsZoomSign()
    ' Zooms the view to the last milepost sign that shows a decimal, for a screenshot.
    Dim scan As ElementEnumerator, item As Element, box As Range3d, found As Boolean
    Set scan = ActiveModelReference.GraphicalElementCache.Scan()
    Do While scan.MoveNext
        Set item = scan.Current
        If item.IsCellElement Then
            If item.AsCellElement.Name = "MILEPOST" Then
                If Hypot(item.Range.High.x - item.Range.Low.x, 0) > 0 Then box = item.Range: found = True
            End If
        End If
    Loop
    If found Then ZoomToExtent box.Low.x - 600, box.Low.y - 100, box.High.x + 600, box.High.y + 100
    mHandle = FreeFile
    Open Environ$("TEMP") & "\vdot_lrs_vba_probe.txt" For Output As #mHandle
    Out "found", found
    Out "box", FmtNum(box.Low.x, 0) & ", " & FmtNum(box.Low.y, 0) & " to " & FmtNum(box.High.x, 0) & ", " & FmtNum(box.High.y, 0)
    Out "viewOrigin", FmtNum(ActiveDesignFile.Views(1).origin.x, 0) & ", " & FmtNum(ActiveDesignFile.Views(1).origin.y, 0)
    Out "viewExtents", FmtNum(ActiveDesignFile.Views(1).Extents.x, 0) & ", " & FmtNum(ActiveDesignFile.Views(1).Extents.y, 0)
    Close #mHandle
End Sub

Public Sub LrsZoomMeasures()
    ' Zooms to the stretch between two measures of the loaded route ("from to" in
    ' %TEMP%\lrs_zoom.txt), for a screenshot.
    Dim handle As Integer, m1 As Double, m2 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim pad As Double
    handle = FreeFile
    Open Environ$("TEMP") & "\lrs_zoom.txt" For Input As #handle
    Input #handle, m1, m2
    Close #handle
    frmLrsEvent.TestLocate m1, x1, y1
    frmLrsEvent.TestLocate m2, x2, y2
    pad = 0.2 * Hypot(x2 - x1, y2 - y1) + 100
    ZoomToExtent MinD(x1, x2) - pad, MinD(y1, y2) - pad, MaxD(x1, x2) + pad, MaxD(y1, y2) + pad
End Sub

Public Sub LrsInspectCell()
    ' Details of the last MILEPOST cell in %TEMP%\vdot_lrs_vba_probe.txt.
    Dim scan As ElementEnumerator, item As Element, last As CellElement, parts As ElementEnumerator, part As Element
    Dim line As String
    Set scan = ActiveModelReference.GraphicalElementCache.Scan()
    Do While scan.MoveNext
        Set item = scan.Current
        If item.IsCellElement Then
            If item.AsCellElement.Name = "MILEPOST" Then Set last = item.AsCellElement
        End If
    Loop
    mHandle = FreeFile
    Open Environ$("TEMP") & "\vdot_lrs_vba_probe.txt" For Output As #mHandle
    Out "cell", DLongToString(last.ID) & " origin " & FmtNum(last.origin.x, 0) & ", " & FmtNum(last.origin.y, 0)
    Set parts = last.GetSubElements
    Do While parts.MoveNext
        Set part = parts.Current
        line = "type " & part.Type & " range " & FmtNum(part.Range.Low.x, 0) & ", " & FmtNum(part.Range.Low.y, 0) & _
               " to " & FmtNum(part.Range.High.x, 0) & ", " & FmtNum(part.Range.High.y, 0)
        If part.IsTextElement Then
            line = line & " text '" & part.AsTextElement.text & "' h " & FmtNum(part.AsTextElement.TextStyle.height, 1) & _
                   " w " & FmtNum(part.AsTextElement.TextStyle.width, 1) & " origin " & _
                   FmtNum(part.AsTextElement.origin.x, 0) & ", " & FmtNum(part.AsTextElement.origin.y, 0)
        End If
        Out "part", line
    Loop
    Close #mHandle
End Sub

Public Sub LrsRemoveRange()
    ' Removes scratch elements whose ID is within lrs_remove_range.txt ("low high") in %TEMP%.
    Dim scan As ElementEnumerator, item As Element, low As Long, high As Long, n As Long, handle As Integer
    Dim doomed() As Element, i As Long
    If LCase$(ActiveDesignFile.Name) <> "lrs_vba_scratch.dgn" Then Exit Sub
    handle = FreeFile
    Open Environ$("TEMP") & "\lrs_remove_range.txt" For Input As #handle
    Input #handle, low, high
    Close #handle
    Set scan = ActiveModelReference.GraphicalElementCache.Scan()
    Do While scan.MoveNext
        Set item = scan.Current
        If item.ID.High = 0 And item.ID.low >= low And item.ID.low <= high Then
            ReDim Preserve doomed(0 To n)
            Set doomed(n) = item: n = n + 1
        End If
    Loop
    For i = 0 To n - 1
        ActiveModelReference.RemoveElement doomed(i)
    Next i
    mHandle = FreeFile
    Open Environ$("TEMP") & "\vdot_lrs_vba_probe.txt" For Output As #mHandle
    Out "removed", n
    Close #mHandle
End Sub

Private Function LastLogLine() As String
    Dim lines() As String
    lines = Split(CStr(frmLrsEvent.TestValue("log")), vbCrLf)
    LastLogLine = lines(UBound(lines))
End Function

Private Function ElementCount() As Long
    Dim scan As ElementEnumerator
    Set scan = ActiveModelReference.GraphicalElementCache.Scan()
    Do While scan.MoveNext
        ElementCount = ElementCount + 1
    Loop
End Function

Private Function LastItemSummary() As String
    ' Reads the LRS item back from the most recently added element that has one.
    Dim scan As ElementEnumerator, item As Element, libraries As ItemTypeLibraries, library As ItemTypeLibrary
    Dim kind As ItemType, handler As ItemTypePropertyHandler, found As String
    Set libraries = New ItemTypeLibraries
    Set library = libraries.FindByName(ITEM_LIBRARY)
    Set kind = library.GetItemTypeByName(ITEM_TYPE)
    Set scan = ActiveModelReference.GraphicalElementCache.Scan()
    On Error Resume Next
    Do While scan.MoveNext
        Set item = scan.Current
        Set handler = Nothing
        Set handler = kind.FindItem(item)
        If Not handler Is Nothing Then
            found = "Route=" & handler.GetPropertyValue("Route") & ", Name=" & handler.GetPropertyValue("Name") & _
                    ", EventType=" & handler.GetPropertyValue("EventType") & ", StartMeasure=" & _
                    handler.GetPropertyValue("StartMeasure") & ", EndMeasure=" & handler.GetPropertyValue("EndMeasure") & _
                    ", FromLatitude=" & handler.GetPropertyValue("FromLatitude")
        End If
    Loop
    LastItemSummary = found
End Function
