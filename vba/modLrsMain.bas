Option Explicit
' VDOT LRS Linear Milepost - MicroStation VBA port of VDOT_LRS_Linear_Milepost.py
'
' Run:  Utilities > Macros > VBA, project VDOT_LRS_Linear_Milepost, macro VdotLrsEvent
'       or key in:  VBA RUN VdotLrsEvent
'
' Modules
'   modLrsMain     constants, entry point, shared helpers, embedded logo
'   modLrsJson     JSON reader for the REST replies
'   modLrsDraw     everything that touches the DGN (projection, elements, item type)
'   clsLrsService  VDOT LRS Route Master REST client
'   clsLrsGeometry route polyline with measures
'   clsLrsRoute    one route (attributes + geometry)
'   clsLrsPickTool pick tool: previews the route and picks From / To
'   frmLrsEvent    the dialog

Public Const APP_TITLE As String = "VDOT LRS Linear Milepost"
Public Const APP_AUTHOR As String = "Robert Ratliff"
Public Const BUILD_DATE As String = "2026-11-25"

Public Const SERVICE_URL As String = "https://services.arcgis.com/p5v98VHDX9Atv3l7/arcgis/rest/services/LRS_Route_Master/FeatureServer/0"
Public Const STATEWIDE_LABEL As String = "Statewide routes (Interstate, US, VA)"

Public Const ITEM_LIBRARY As String = "VDOT LRS"
Public Const ITEM_TYPE As String = "Linear Event"
' Labels, start/stop markers and milepost signs carry this item type, so "Delete created"
' can find everything the tool placed, in this session or a later one.
Public Const GRAPHIC_TYPE As String = "LRS Graphic"
Public Const GRAPHIC_PART As String = "Part"
Public Const PART_LABEL As String = "Label"
Public Const PART_FROM As String = "Start marker"
Public Const PART_TO As String = "Stop marker"
Public Const PART_MILEPOST As String = "Milepost"

Public Const MODE_MODEL_GCS As String = "model"
Public Const MODE_EPSG As String = "epsg"
Public Const MODE_LOCAL As String = "local"

Public Const EVENT_LINEAR As String = "Linear"
Public Const EVENT_POINT As String = "Point"

Public Const LABELS_ALL As String = "all"
Public Const LABELS_MEASURES As String = "measures"
Public Const LABELS_NONE As String = "none"

' Milepost sign modes: "none", "ends" or a spacing in miles ("1", "0.5", "0.1").
Public Const MILEPOST_NONE As String = "none"
Public Const MILEPOST_ENDS As String = "ends"
Public Const MAX_MILEPOSTS As Long = 400
Public Const MILEPOST_SIZE As Double = 1#          ' multiplier on the automatic sign size
Public Const LABEL_HEIGHT_MIN_M As Double = 5#     ' smallest and largest label text height, metres
Public Const LABEL_HEIGHT_MAX_M As Double = 50#
Public Const SIGN_HEIGHT_MIN_M As Double = 20#     ' smallest and largest milepost sign height, metres
Public Const SIGN_HEIGHT_MAX_M As Double = 250#
Public Const LABEL_CHAR_WIDTH As Double = 0.7      ' estimated character width as a share of text height

' Preview symbology: colour-table index and weight.
Public Const PREVIEW_ROUTE_COLOR As Long = 1
Public Const PREVIEW_ROUTE_WEIGHT As Long = 1
Public Const PREVIEW_EVENT_COLOR As Long = 7
Public Const PREVIEW_EVENT_WEIGHT As Long = 5
Public Const PREVIEW_FROM_COLOR As Long = 2
Public Const PREVIEW_TO_COLOR As Long = 3
Public Const PREVIEW_PIN_WEIGHT As Long = 30
Public Const PREVIEW_CURSOR_COLOR As Long = 4
Public Const PREVIEW_CURSOR_WEIGHT As Long = 16

' Symbology of what "Create alignment" adds besides the line itself.
Public Const CREATED_FROM_COLOR As Long = 2        ' start marker - green
Public Const CREATED_TO_COLOR As Long = 3          ' stop marker - red
Public Const POINT_EVENT_COLOR As Long = 4         ' Point event - yellow
Public Const CREATED_PIN_WEIGHT As Long = 15

Public Const MAX_LINESTRING_POINTS As Long = 4500
Public Const WGS84_A As Double = 6378137#
Public Const WGS84_E2 As Double = 0.00669437999014

Public Const ERR_SERVICE As Long = vbObjectError + 513     ' REST service problem
Public Const ERR_INPUT As Long = vbObjectError + 514       ' invalid user input
Public Const ERR_CANCELLED As Long = vbObjectError + 515   ' a request was abandoned (dialog closing)
Public Const ERR_DATA As Long = vbObjectError + 516        ' the service answered, but not usefully

Public gLrsFormOpen As Boolean
Public gLrsToolActive As Boolean

' ---------------------------------------------------------------------------
' Entry point
' ---------------------------------------------------------------------------
Public Sub VdotLrsEvent()
    ' Showing the form again while it is open just brings it forward.
    frmLrsEvent.Show vbModeless
End Sub

' ---------------------------------------------------------------------------
' Number and text helpers
' ---------------------------------------------------------------------------
Public Function FmtNum(ByVal value As Double, ByVal decimals As Long) As String
    ' Fixed decimals with "." whatever the Windows regional settings are.
    Dim text As String
    If Abs(value) < 0.5 * 10 ^ (-decimals) Then value = 0#
    If decimals > 0 Then
        text = Format$(value, "0." & String$(decimals, "0"))
    Else
        text = Format$(value, "0")
    End If
    FmtNum = Replace(text, ",", ".")
End Function

Public Function FmtMeasure(ByVal measure As Double) As String
    FmtMeasure = FmtNum(Round(measure, 3), 3)
End Function

Public Function FmtThousands(ByVal value As Double, ByVal decimals As Long) As String
    Dim text As String, whole As String, frac As String, result As String, dot As Long
    text = FmtNum(Abs(value), decimals)
    dot = InStr(text, ".")
    If dot > 0 Then
        whole = Left$(text, dot - 1): frac = Mid$(text, dot)
    Else
        whole = text
    End If
    Do While Len(whole) > 3
        result = "," & Right$(whole, 3) & result
        whole = Left$(whole, Len(whole) - 3)
    Loop
    FmtThousands = IIf(value < 0, "-", "") & whole & result & frac
End Function

Public Function NumStr(ByVal value As Double) As String
    ' Full-precision number for a REST request ("." decimal, no thousands separator).
    NumStr = Trim$(Str$(value))
End Function

Public Function Ceil(ByVal value As Double) As Long
    Ceil = -Int(-value)
End Function

Public Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Public Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Public Function ClampD(ByVal value As Double, ByVal low As Double, ByVal high As Double) As Double
    ClampD = MinD(MaxD(value, low), high)
End Function

Public Sub SideCentre(ByVal x As Double, ByVal y As Double, ByVal dirX As Double, ByVal dirY As Double, _
                      ByVal side As Double, ByVal width As Double, ByVal height As Double, ByVal gap As Double, _
                      ByRef centreX As Double, ByRef centreY As Double)
    ' Centre of an upright width x height box placed beside the route at (x, y). side is +1 for
    ' the right of the direction of travel, -1 for the left; gap is the clear distance between
    ' the route and the nearest part of the box.
    Dim nx As Double, ny As Double, reach As Double
    nx = dirY * side: ny = -dirX * side
    reach = gap + Abs(nx) * width / 2# + Abs(ny) * height / 2#
    centreX = x + nx * reach: centreY = y + ny * reach
End Sub

Public Function SegmentHitsBox(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double, _
                               ByVal xMin As Double, ByVal yMin As Double, ByVal xMax As Double, _
                               ByVal yMax As Double) As Boolean
    ' True when the segment a-b touches the axis-aligned box (Liang-Barsky clipping).
    Dim p(0 To 3) As Double, q(0 To 3) As Double, i As Long, enter As Double, leave As Double, t As Double
    If MaxD(ax, bx) < xMin Or MinD(ax, bx) > xMax Or MaxD(ay, by) < yMin Or MinD(ay, by) > yMax Then Exit Function
    p(0) = ax - bx: q(0) = ax - xMin
    p(1) = bx - ax: q(1) = xMax - ax
    p(2) = ay - by: q(2) = ay - yMin
    p(3) = by - ay: q(3) = yMax - ay
    enter = 0#: leave = 1#
    For i = 0 To 3
        If p(i) = 0# Then
            If q(i) < 0# Then Exit Function
        Else
            t = q(i) / p(i)
            If p(i) < 0# Then
                If t > leave Then Exit Function
                If t > enter Then enter = t
            Else
                If t < enter Then Exit Function
                If t < leave Then leave = t
            End If
        End If
    Next i
    SegmentHitsBox = True
End Function

Public Function Hypot(ByVal dx As Double, ByVal dy As Double) As Double
    Hypot = Sqr(dx * dx + dy * dy)
End Function

Public Function HexColor(ByVal hexText As String) As Long
    ' "#rrggbb" -> VBA colour
    HexColor = RGB(CLng("&H" & Mid$(hexText, 2, 2)), CLng("&H" & Mid$(hexText, 4, 2)), CLng("&H" & Mid$(hexText, 6, 2)))
End Function

Public Sub SortDoubles(values() As Double, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As Double, temp As Double
    If low >= high Then Exit Sub
    i = low: j = high: pivot = values((low + high) \ 2)
    Do While i <= j
        Do While values(i) < pivot: i = i + 1: Loop
        Do While values(j) > pivot: j = j - 1: Loop
        If i <= j Then
            temp = values(i): values(i) = values(j): values(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then SortDoubles values, low, j
    If i < high Then SortDoubles values, i, high
End Sub

Public Sub SortStrings(values() As String, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As String, temp As String
    If low >= high Then Exit Sub
    i = low: j = high: pivot = values((low + high) \ 2)
    Do While i <= j
        Do While StrComp(values(i), pivot, vbBinaryCompare) < 0: i = i + 1: Loop
        Do While StrComp(values(j), pivot, vbBinaryCompare) > 0: j = j - 1: Loop
        If i <= j Then
            temp = values(i): values(i) = values(j): values(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then SortStrings values, low, j
    If i < high Then SortStrings values, i, high
End Sub

Public Sub SortIndexByKey(keys() As String, index() As Long, ByVal low As Long, ByVal high As Long)
    ' Sorts index() so that keys(index(i)) ascends.
    Dim i As Long, j As Long, pivot As String, temp As Long
    If low >= high Then Exit Sub
    i = low: j = high: pivot = keys(index((low + high) \ 2))
    Do While i <= j
        Do While StrComp(keys(index(i)), pivot, vbBinaryCompare) < 0: i = i + 1: Loop
        Do While StrComp(keys(index(j)), pivot, vbBinaryCompare) > 0: j = j - 1: Loop
        If i <= j Then
            temp = index(i): index(i) = index(j): index(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then SortIndexByKey keys, index, low, j
    If i < high Then SortIndexByKey keys, index, i, high
End Sub

Public Sub PauseSeconds(ByVal seconds As Double)
    Dim finish As Double
    finish = Timer + seconds
    Do While Timer < finish And Timer > finish - seconds - 1
        DoEvents
    Loop
End Sub

' ---------------------------------------------------------------------------
' Pure geometry helpers shared by the dialog and the drawing code
' ---------------------------------------------------------------------------
Public Sub LocalXYMetres(ByVal lon As Double, ByVal lat As Double, ByVal lon0 As Double, ByVal lat0 As Double, _
                         ByRef east As Double, ByRef north As Double)
    ' Metres east/north of (lon0, lat0) on a tangent plane at that point.
    Const PI As Double = 3.14159265358979
    Dim phi As Double, w As Double
    phi = lat0 * PI / 180#
    w = Sqr(1# - WGS84_E2 * Sin(phi) ^ 2)
    east = (lon - lon0) * PI / 180# * (WGS84_A / w) * Cos(phi)
    north = (lat - lat0) * PI / 180# * (WGS84_A * (1# - WGS84_E2) / w ^ 3)
End Sub

Public Sub MilepostText(ByVal measure As Double, ByVal mode As String, ByRef digits As String, ByRef decimalText As String)
    ' Sign legend for a measure: whole-mile digits and a decimal such as ".5" (or "").
    Dim value As Double, whole As Long
    value = Round(measure, 3)
    If value < 0.0005 Then value = 0#
    whole = Int(value + 0.0000001)
    digits = CStr(whole)
    decimalText = Mid$(FmtNum(value - whole, 3), 2)             ' ".500"
    Do While Right$(decimalText, 1) = "0"
        decimalText = Left$(decimalText, Len(decimalText) - 1)
    Loop
    If decimalText = "." Then
        If mode = "1" Or mode = MILEPOST_ENDS Then decimalText = "" Else decimalText = ".0"
    End If
End Sub

Public Sub MilepostLayout(ByVal digits As String, ByVal decimalText As String, ByVal unit As Double, _
                          ByRef width As Double, ByRef height As Double, ByRef texts() As String, _
                          ByRef centreY() As Double, ByRef textHeight() As Double, ByRef textCount As Long, _
                          ByRef hasRule As Boolean, ByRef ruleY As Double)
    ' Sign face for a legend, y measured up from the bottom edge. unit = height of one digit.
    Dim margin As Double, gap As Double, mileHeight As Double, y As Double
    Dim rows As Long, i As Long, rowHeight As Double
    margin = 0.45 * unit: gap = 0.35 * unit: mileHeight = 0.5 * unit
    rows = 1 + Len(digits)
    ReDim texts(0 To rows), centreY(0 To rows), textHeight(0 To rows)
    height = 2 * margin + mileHeight + Len(digits) * unit + gap * (rows - 1)
    If decimalText <> "" Then height = height + 2 * gap + unit
    y = height - margin
    For i = 0 To rows - 1
        If i = 0 Then
            texts(i) = "MILE": rowHeight = mileHeight
        Else
            texts(i) = Mid$(digits, i, 1): rowHeight = unit
        End If
        centreY(i) = y - rowHeight / 2#
        textHeight(i) = rowHeight
        y = y - rowHeight - gap
    Next i
    textCount = rows
    hasRule = (decimalText <> "")
    If hasRule Then
        ruleY = y
        texts(rows) = decimalText
        centreY(rows) = y - gap - unit / 2#
        textHeight(rows) = unit
        textCount = rows + 1
    End If
    width = 3# * unit
End Sub

Public Sub RoundedRectangle(ByVal xCentre As Double, ByVal bottom As Double, ByVal width As Double, _
                            ByVal height As Double, ByVal radius As Double, ByRef px() As Double, _
                            ByRef py() As Double, ByRef count As Long)
    ' Closed outline of a rounded rectangle.
    Const PI As Double = 3.14159265358979
    Const STEPS As Long = 6
    Dim leftX As Double, rightX As Double, topY As Double
    Dim cx(0 To 3) As Double, cy(0 To 3) As Double, startAngle(0 To 3) As Double
    Dim corner As Long, stepIndex As Long, angle As Double
    leftX = xCentre - width / 2#: rightX = xCentre + width / 2#: topY = bottom + height
    cx(0) = rightX - radius: cy(0) = bottom + radius: startAngle(0) = -90#
    cx(1) = rightX - radius: cy(1) = topY - radius: startAngle(1) = 0#
    cx(2) = leftX + radius: cy(2) = topY - radius: startAngle(2) = 90#
    cx(3) = leftX + radius: cy(3) = bottom + radius: startAngle(3) = 180#
    ReDim px(0 To 4 * (STEPS + 1)), py(0 To 4 * (STEPS + 1))
    count = 0
    For corner = 0 To 3
        For stepIndex = 0 To STEPS
            angle = (startAngle(corner) + 90# * stepIndex / STEPS) * PI / 180#
            px(count) = cx(corner) + radius * Cos(angle)
            py(count) = cy(corner) + radius * Sin(angle)
            count = count + 1
        Next stepIndex
    Next corner
    px(count) = px(0): py(count) = py(0)
    count = count + 1
End Sub

' ---------------------------------------------------------------------------
' Embedded VDOT logo
' ---------------------------------------------------------------------------
Public Function LogoFile() As String
    ' Writes the embedded logo (JPEG) to the temp folder and returns its path, "" on failure.
    Dim xml As Object, node As Object, bytes() As Byte, path As String, handle As Integer
    On Error GoTo Failed
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.text = LogoBase64()
    bytes = node.nodeTypedValue
    path = Environ$("TEMP") & "\vdot_lrs_logo.jpg"
    If Dir$(path) <> "" Then Kill path
    handle = FreeFile
    Open path For Binary Access Write As #handle
    Put #handle, , bytes
    Close #handle
    LogoFile = path
    Exit Function
Failed:
    On Error Resume Next
    Close #handle
    LogoFile = ""
End Function
