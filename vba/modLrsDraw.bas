Option Explicit
' Everything that touches the DGN: coordinate conversion, preview and created elements,
' the "VDOT LRS / Linear Event" item type, labels and milepost signs.
' Coordinates here are MicroStation master units.

Private mMode As String
Private mGcs As GeographicCoordinateSystem
Private mMasterPerMetre As Double
Private mLon0 As Double, mLat0 As Double
' Lambert Conformal Conic constants of the chosen EPSG system
Private mLccN As Double, mLccF As Double, mLccRho0 As Double, mLccLon0 As Double
Private mLccFE As Double, mLccFN As Double

' Kept referenced while item types from them are in use.
Private mItemLibraries As ItemTypeLibraries
Private mItemLibrary As ItemTypeLibrary
Private mGraphicType As ItemType

Private Const PI As Double = 3.14159265358979
Private Const GRS80_A As Double = 6378137#
Private Const GRS80_E2 As Double = 0.0066943800229

' ---------------------------------------------------------------------------
' Active model
' ---------------------------------------------------------------------------
Public Function Model3D() As Boolean
    Model3D = ActiveModelReference.Is3D
End Function

Public Function ModelGcsName() As String
    ' Name of the active model's geographic coordinate system, "" when it has none.
    Dim gcs As GeographicCoordinateSystem
    On Error Resume Next
    Set gcs = ActiveModelReference.GetGCS(True)
    If Err.Number <> 0 Or gcs Is Nothing Then Exit Function
    ModelGcsName = gcs.Name
    If ModelGcsName = "" Then ModelGcsName = "unnamed GCS"
End Function

Private Function MasterPerMetre() As Double
    Dim unit As MeasurementUnit
    unit = ActiveModelReference.MasterUnit
    If unit.UnitsPerBaseDenominator = 0# Or unit.UnitsPerBaseNumerator = 0# Then
        MasterPerMetre = 1#
    Else
        MasterPerMetre = unit.UnitsPerBaseNumerator / unit.UnitsPerBaseDenominator
    End If
End Function

' ---------------------------------------------------------------------------
' Coordinate systems offered when the model GCS is not used
' ---------------------------------------------------------------------------
Public Function EpsgCount() As Long
    EpsgCount = 5
End Function

Public Function EpsgCode(ByVal index As Long) As Long
    EpsgCode = Choose(index + 1, 2284, 2283, 6595, 6593, 3968)
End Function

Public Function EpsgName(ByVal index As Long) As String
    EpsgName = Choose(index + 1, "NAD83 Virginia South (ftUS)", "NAD83 Virginia North (ftUS)", _
                      "NAD83(2011) Virginia South (ftUS)", "NAD83(2011) Virginia North (ftUS)", _
                      "NAD83 Virginia Lambert (m)")
End Function

Private Function LccM(ByVal phi As Double) As Double
    LccM = Cos(phi) / Sqr(1# - GRS80_E2 * Sin(phi) ^ 2)
End Function

Private Function LccT(ByVal phi As Double) As Double
    Dim e As Double
    e = Sqr(GRS80_E2)
    LccT = Tan(PI / 4# - phi / 2#) / ((1# - e * Sin(phi)) / (1# + e * Sin(phi))) ^ (e / 2#)
End Function

Private Sub SetupLcc(ByVal code As Long)
    ' Lambert Conformal Conic (2 standard parallels) on GRS80; false origin in metres.
    Dim lat0 As Double, sp1 As Double, sp2 As Double
    Select Case code
        Case 2284, 6595             ' Virginia South
            lat0 = 36# + 20# / 60#: mLccLon0 = -78.5
            sp1 = 36# + 46# / 60#: sp2 = 37# + 58# / 60#
            mLccFE = 3500000#: mLccFN = 1000000#
        Case 2283, 6593             ' Virginia North
            lat0 = 37# + 40# / 60#: mLccLon0 = -78.5
            sp1 = 38# + 2# / 60#: sp2 = 39# + 12# / 60#
            mLccFE = 3500000#: mLccFN = 2000000#
        Case 3968                   ' Virginia Lambert
            lat0 = 36#: mLccLon0 = -79.5
            sp1 = 37#: sp2 = 39.5
            mLccFE = 0#: mLccFN = 0#
        Case Else
            Err.Raise ERR_INPUT, "modLrsDraw", "EPSG:" & code & " is not supported."
    End Select
    lat0 = lat0 * PI / 180#: sp1 = sp1 * PI / 180#: sp2 = sp2 * PI / 180#
    mLccN = (Log(LccM(sp1)) - Log(LccM(sp2))) / (Log(LccT(sp1)) - Log(LccT(sp2)))
    mLccF = LccM(sp1) / (mLccN * LccT(sp1) ^ mLccN)
    mLccRho0 = GRS80_A * mLccF * LccT(lat0) ^ mLccN
End Sub

Public Function InitProjector(ByVal mode As String, ByVal epsg As Long, ByVal originLon As Double, _
                              ByVal originLat As Double) As String
    ' Prepares ProjectPoint. Returns the line to write to the log.
    mMode = mode
    Set mGcs = Nothing
    mMasterPerMetre = MasterPerMetre()
    Select Case mode
        Case MODE_LOCAL
            mLon0 = originLon: mLat0 = originLat
            InitProjector = "Local tangent plane: the route start is placed at the model origin."
        Case MODE_MODEL_GCS
            On Error Resume Next
            Set mGcs = ActiveModelReference.GetGCS(True)
            On Error GoTo 0
            If mGcs Is Nothing Then
                Err.Raise ERR_INPUT, "modLrsDraw", "The active model has no Geographic Coordinate System."
            End If
            InitProjector = "Converting with the model GCS '" & mGcs.Name & "'."
        Case Else
            SetupLcc epsg
            InitProjector = "Converting with EPSG:" & epsg & " (computed in the macro)."
    End Select
End Function

Public Sub ProjectPoint(ByVal lon As Double, ByVal lat As Double, ByRef x As Double, ByRef y As Double)
    ' Longitude/latitude -> master units of the active model.
    Dim east As Double, north As Double, geo As GeoPoint3D, point As Point3d
    Dim rho As Double, theta As Double
    Select Case mMode
        Case MODE_LOCAL
            LocalXYMetres lon, lat, mLon0, mLat0, east, north
            x = east * mMasterPerMetre: y = north * mMasterPerMetre
        Case MODE_MODEL_GCS
            geo.Longitude = lon: geo.Latitude = lat: geo.Elevation = 0#
            On Error Resume Next
            point = mGcs.MasterUnitsFromLatLong(geo)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Err.Raise ERR_INPUT, "modLrsDraw", "(" & FmtNum(lat, 6) & ", " & FmtNum(lon, 6) & _
                          ") could not be converted with the model GCS."
            End If
            On Error GoTo 0
            x = point.x: y = point.y
        Case Else
            rho = GRS80_A * mLccF * LccT(lat * PI / 180#) ^ mLccN
            theta = mLccN * (lon - mLccLon0) * PI / 180#
            x = (mLccFE + rho * Sin(theta)) * mMasterPerMetre
            y = (mLccFN + mLccRho0 - rho * Cos(theta)) * mMasterPerMetre
    End Select
End Sub

Public Function ProjectGeometry(ByVal source As clsLrsGeometry) As clsLrsGeometry
    Dim xs() As Double, ys() As Double, i As Long, n As Long
    n = source.VertexCount
    ReDim xs(0 To n - 1), ys(0 To n - 1)
    For i = 0 To n - 1
        ProjectPoint source.VX(i), source.VY(i), xs(i), ys(i)
    Next i
    Set ProjectGeometry = source.Mapped(xs, ys)
End Function

Public Function MetresFromMaster(ByVal length As Double) As Double
    MetresFromMaster = length / MasterPerMetre()
End Function

' ---------------------------------------------------------------------------
' Elements
' ---------------------------------------------------------------------------
Public Sub AppendElement(ByRef elements() As Element, ByRef count As Long, ByVal item As Element)
    ReDim Preserve elements(0 To count)
    Set elements(count) = item
    count = count + 1
End Sub

Public Function PartPoints(ByVal shape As clsLrsGeometry, ByVal part As Long) As Variant
    ' Vertices of one part as a 2-D array (0..n-1, 0..1).
    Dim points() As Double, k As Long, n As Long
    ReDim points(0 To shape.PartEnd(part) - shape.PartStart(part), 0 To 1)
    For k = shape.PartStart(part) To shape.PartEnd(part)
        points(n, 0) = shape.VX(k): points(n, 1) = shape.VY(k): n = n + 1
    Next k
    PartPoints = points
End Function

Public Function PolylineLength(points As Variant) As Double
    Dim i As Long
    For i = LBound(points, 1) + 1 To UBound(points, 1)
        PolylineLength = PolylineLength + Hypot(points(i, 0) - points(i - 1, 0), points(i, 1) - points(i - 1, 1))
    Next i
End Function

Public Sub AppendLineStrings(points As Variant, ByVal useSymbology As Boolean, ByVal colorIndex As Long, _
                             ByVal weight As Long, ByRef elements() As Element, ByRef count As Long)
    ' Line string element(s) through the points, split at the vertex limit.
    Dim total As Long, start As Long, n As Long, i As Long, vertices() As Point3d, item As LineElement
    total = UBound(points, 1) + 1
    Do While start < total - 1
        n = total - start
        If n > MAX_LINESTRING_POINTS Then n = MAX_LINESTRING_POINTS
        ReDim vertices(0 To n - 1)
        For i = 0 To n - 1
            vertices(i).x = points(start + i, 0): vertices(i).y = points(start + i, 1): vertices(i).Z = 0#
        Next i
        Set item = CreateLineElement1(Nothing, vertices)
        If useSymbology Then
            item.Color = colorIndex
            item.LineWeight = weight
        End If
        AppendElement elements, count, item
        start = start + MAX_LINESTRING_POINTS - 1
    Loop
End Sub

Public Function PointElement(ByVal x As Double, ByVal y As Double, ByVal useSymbology As Boolean, _
                             ByVal colorIndex As Long, ByVal weight As Long) As Element
    ' Zero-length line: displays as a dot whose size follows its weight.
    Dim point As Point3d, item As LineElement
    point = Point3dFromXY(x, y)
    Set item = CreateLineElement2(Nothing, point, point)
    If useSymbology Then
        item.Color = colorIndex
        item.LineWeight = weight
    End If
    Set PointElement = item
End Function

Private Function TargetView() As View
    Dim candidate As View
    On Error Resume Next
    For Each candidate In ActiveDesignFile.Views
        If candidate.IsOpen And candidate.IsSelected Then
            Set TargetView = candidate
            Exit Function
        End If
    Next candidate
    For Each candidate In ActiveDesignFile.Views
        If candidate.IsOpen Then
            Set TargetView = candidate
            Exit Function
        End If
    Next candidate
End Function

Public Sub ZoomToExtent(ByVal xMin As Double, ByVal yMin As Double, ByVal xMax As Double, ByVal yMax As Double)
    ' Centres the selected view on the box with a 20% margin.
    Dim target As View, size As Point3d, origin As Point3d, rotation As Matrix3d, factor As Double
    Dim cx As Double, cy As Double, halfX As Double, halfY As Double
    Set target = TargetView()
    If target Is Nothing Then Exit Sub
    size = target.Extents
    If size.x <= 0# Or size.y <= 0# Then Exit Sub
    factor = 1.2 * MaxD((xMax - xMin) / size.x, (yMax - yMin) / size.y)
    If factor <= 0# Then Exit Sub
    origin = target.origin
    rotation = target.rotation
    cx = (xMin + xMax) / 2#: cy = (yMin + yMax) / 2#
    size.x = size.x * factor: size.y = size.y * factor
    halfX = size.x / 2#: halfY = size.y / 2#
    ' The view origin is its lower-left corner; rows of the rotation are the view axes.
    origin.x = cx - rotation.RowX.x * halfX - rotation.RowY.x * halfY
    origin.y = cy - rotation.RowX.y * halfX - rotation.RowY.y * halfY
    target.origin = origin
    target.Extents = size
    target.Redraw
End Sub

' ---------------------------------------------------------------------------
' Event properties and the item type
' ---------------------------------------------------------------------------
Public Function EventLength(ByVal eventType As String, ByVal fromMeasure As Double, ByVal toMeasure As Double) As Double
    If eventType = EVENT_LINEAR Then EventLength = toMeasure - fromMeasure
End Function

Private Sub EventProperties(ByVal route As clsLrsRoute, ByVal eventType As String, ByVal fromMeasure As Double, _
                            ByVal toMeasure As Double, ByRef names() As String, ByRef values() As Variant, _
                            ByRef count As Long)
    Dim endMeasure As Double, fromLon As Double, fromLat As Double, toLon As Double, toLat As Double
    If eventType = EVENT_LINEAR Then endMeasure = toMeasure Else endMeasure = fromMeasure
    route.Geometry.PointAt fromMeasure, fromLon, fromLat
    route.Geometry.PointAt endMeasure, toLon, toLat
    count = 13
    ReDim names(0 To count - 1), values(0 To count - 1)
    names(0) = "Route": values(0) = route.RouteName
    names(1) = "Name": values(1) = route.NameId
    names(2) = "Jurisdiction": values(2) = route.Jurisdiction
    names(3) = "Category": values(3) = route.Category
    names(4) = "EventType": values(4) = eventType
    names(5) = "StartMeasure": values(5) = Round(fromMeasure, 3) + 0#
    names(6) = "EndMeasure": values(6) = Round(endMeasure, 3) + 0#
    names(7) = "EventLength": values(7) = Round(EventLength(eventType, fromMeasure, toMeasure), 3)
    names(8) = "RouteLength": values(8) = Round(route.Geometry.MMax, 3)
    names(9) = "FromLatitude": values(9) = Round(fromLat, 6)
    names(10) = "FromLongitude": values(10) = Round(fromLon, 6)
    names(11) = "ToLatitude": values(11) = Round(toLat, 6)
    names(12) = "ToLongitude": values(12) = Round(toLon, 6)
End Sub

Private Function EventItemType(names() As String, values() As Variant, ByVal count As Long) As ItemType
    ' Returns the item type, creating or extending the library as needed. The library objects
    ' stay referenced in this module for as long as the item type is used.
    Dim target As ItemType, changed As Boolean, i As Long, kind As ItemTypePropertyType, graphic As ItemType
    Set mItemLibraries = New ItemTypeLibraries
    Set mItemLibrary = Nothing
    On Error Resume Next
    Set mItemLibrary = mItemLibraries.FindByName(ITEM_LIBRARY)
    On Error GoTo 0
    If mItemLibrary Is Nothing Then
        Set mItemLibrary = mItemLibraries.CreateLib(ITEM_LIBRARY, False)
        changed = True
    End If
    On Error Resume Next
    Set target = mItemLibrary.GetItemTypeByName(ITEM_TYPE)
    On Error GoTo 0
    If target Is Nothing Then
        Set target = mItemLibrary.AddItemType(ITEM_TYPE, False)
        changed = True
    End If
    For i = 0 To count - 1
        If Not HasItemProperty(target, names(i)) Then
            If VarType(values(i)) = vbString Then kind = ItemPropertyTypeString Else kind = ItemPropertyTypeDouble
            target.AddProperty names(i), kind, False
            changed = True
        End If
    Next i
    Set graphic = Nothing
    On Error Resume Next
    Set graphic = mItemLibrary.GetItemTypeByName(GRAPHIC_TYPE)
    On Error GoTo 0
    If graphic Is Nothing Then
        Set graphic = mItemLibrary.AddItemType(GRAPHIC_TYPE, False)
        changed = True
    End If
    If Not HasItemProperty(graphic, GRAPHIC_PART) Then
        graphic.AddProperty GRAPHIC_PART, ItemPropertyTypeString, False
        changed = True
    End If
    If changed Then
        If Not mItemLibrary.Write() Then
            Err.Raise ERR_INPUT, "modLrsDraw", "Could not write item type library '" & ITEM_LIBRARY & "'."
        End If
    End If
    If target Is Nothing Or graphic Is Nothing Then
        Err.Raise ERR_INPUT, "modLrsDraw", "Item type '" & ITEM_LIBRARY & " / " & ITEM_TYPE & "' is not available."
    End If
    Set mGraphicType = graphic
    Set EventItemType = target
End Function

Private Sub TagGraphic(ByVal item As Element, ByVal part As String)
    ' Marks an element the tool added (label, marker or sign) so Delete can find it.
    Dim handler As ItemTypePropertyHandler
    If mGraphicType Is Nothing Then Exit Sub
    Set handler = mGraphicType.AttachItem(item)
    If handler Is Nothing Then
        Err.Raise ERR_INPUT, "modLrsDraw", "The LRS Graphic item type could not be applied to the element."
    End If
    handler.SetPropertyValue GRAPHIC_PART, part
End Sub

' ---------------------------------------------------------------------------
' Delete what the tool created
' ---------------------------------------------------------------------------
Private Function CreatedKind(ByVal item As Element, ByVal eventType As ItemType, ByVal graphic As ItemType) As String
    ' ITEM_TYPE or the Part of an element the tool created, else "".
    Dim handler As ItemTypePropertyHandler
    On Error Resume Next
    If Not eventType Is Nothing Then
        Set handler = eventType.FindItem(item)
        If Not handler Is Nothing Then CreatedKind = ITEM_TYPE: Exit Function
    End If
    If Not graphic Is Nothing Then
        Set handler = graphic.FindItem(item)
        If Not handler Is Nothing Then
            CreatedKind = CStr(handler.GetPropertyValue(GRAPHIC_PART))
            If CreatedKind = "" Then CreatedKind = GRAPHIC_TYPE
        End If
    End If
End Function

Public Function FindCreatedElements(ByRef summary As String) As Collection
    ' Elements in the active model that the tool created; summary counts them by kind.
    Dim libraries As ItemTypeLibraries, library As ItemTypeLibrary, eventType As ItemType, graphic As ItemType
    Dim criteria As ElementScanCriteria, found As ElementEnumerator, kind As String, i As Long
    Dim kinds As Variant, names As Variant, counts(0 To 5) As Long
    Set FindCreatedElements = New Collection
    kinds = Array(ITEM_TYPE, PART_FROM, PART_TO, PART_LABEL, PART_MILEPOST, GRAPHIC_TYPE)
    names = Array("event", "start marker", "stop marker", "label", "milepost sign", "other")
    Set libraries = New ItemTypeLibraries
    On Error Resume Next
    Set library = libraries.FindByName(ITEM_LIBRARY)
    If Not library Is Nothing Then
        Set eventType = library.GetItemTypeByName(ITEM_TYPE)
        Set graphic = library.GetItemTypeByName(GRAPHIC_TYPE)
    End If
    On Error GoTo 0
    If eventType Is Nothing And graphic Is Nothing Then Exit Function
    Set criteria = New ElementScanCriteria
    criteria.ExcludeNonGraphical
    Set found = ActiveModelReference.Scan(criteria)
    Do While found.MoveNext
        kind = CreatedKind(found.Current, eventType, graphic)
        If kind <> "" Then
            FindCreatedElements.Add found.Current
            For i = 0 To 5
                If kind = kinds(i) Or i = 5 Then counts(i) = counts(i) + 1: Exit For
            Next i
        End If
    Loop
    For i = 0 To 5
        If counts(i) > 0 Then
            summary = summary & IIf(summary = "", "", ", ") & counts(i) & " " & names(i) & IIf(counts(i) = 1, "", "s")
        End If
    Next i
End Function

Public Function DeleteCreatedElements(ByVal items As Collection) As Long
    ' Deletes the given elements; returns the number deleted.
    Dim item As Element
    For Each item In items
        On Error Resume Next
        ActiveModelReference.RemoveElement item
        If Err.Number = 0 Then DeleteCreatedElements = DeleteCreatedElements + 1
        Err.Clear
        On Error GoTo 0
    Next item
End Function

Private Function HasItemProperty(ByVal target As ItemType, ByVal propertyName As String) As Boolean
    Dim found As ItemTypeProperty
    On Error Resume Next
    Set found = target.GetPropertyByName(propertyName)
    HasItemProperty = Not (found Is Nothing)
End Function

Private Sub AttachMetadata(ByVal item As Element, ByVal target As ItemType, names() As String, _
                           values() As Variant, ByVal count As Long)
    Dim handler As ItemTypePropertyHandler, i As Long
    Set handler = target.AttachItem(item)
    If handler Is Nothing Then
        Err.Raise ERR_INPUT, "modLrsDraw", "The item type could not be applied to the element."
    End If
    For i = 0 To count - 1
        handler.SetPropertyValue names(i), values(i)
    Next i
End Sub

' ---------------------------------------------------------------------------
' Labels
' ---------------------------------------------------------------------------
Private Sub UseTrueTypeFont(ByVal item As TextElement)
    Dim face As MicroStationDGN.Font
    On Error Resume Next
    Set face = ActiveDesignFile.Fonts.Find(msdFontTypeWindowsTrueType, "Arial")
    If Not face Is Nothing Then Set item.TextStyle.Font = face
End Sub

Private Sub SizeText(ByVal item As TextElement, ByVal height As Double, ByVal width As Double)
    ' Sets the size the text is drawn at. When the model's annotation scale applies to new
    ' text, MicroStation multiplies the stored size by that scale, so it is divided out here.
    Dim factor As Double
    On Error Resume Next
    factor = item.AnnotationScaleFactor
    On Error GoTo 0
    If factor <= 0# Then factor = 1#
    item.TextStyle.height = height / factor
    item.TextStyle.width = width / factor
End Sub

Private Sub PlaceLabel(ByVal text As String, ByVal x As Double, ByVal y As Double, ByVal height As Double)
    ' Places text centred on (x, y).
    Dim item As TextElement
    Set item = CreateTextElement1(Nothing, text, Point3dFromXY(x, y), Matrix3dIdentity)
    UseTrueTypeFont item
    SizeText item, height, height
    item.TextStyle.Justification = msdTextJustificationCenterCenter
    ActiveModelReference.AddElement item
    TagGraphic item, PART_LABEL
End Sub

Private Sub BesideRoute(ByVal shape As clsLrsGeometry, ByVal measure As Double, ByVal side As Double, _
                        ByVal width As Double, ByVal height As Double, ByVal gap As Double, ByVal stack As Double, _
                        ByRef centreX As Double, ByRef centreY As Double)
    ' Centre of an upright width x height box beside the route at a measure. The box starts
    ' gap clear of the route (see SideCentre) and, because a road that curves can swing back
    ' into it, is moved further out in steps until it no longer touches the route nearby.
    ' If no position on that side is clear (a sharp bend), the other side is tried; failing
    ' that the box stays at its first position. stack shifts it vertically (a second line
    ' of text).
    Dim x As Double, y As Double, dx As Double, dy As Double, stepOut As Double, span As Double
    Dim extra As Long, pass As Long, trySide As Double, firstX As Double, firstY As Double, margin As Double
    margin = gap / 2#                            ' the box must clear the route by this much
    shape.PointAt measure, x, y
    shape.DirectionAt measure, dx, dy
    stepOut = 0.75 * MinD(width, height)
    span = 1.5 * (width + height + gap + 12 * stepOut + Abs(stack)) * shape.MeasurePerUnit()
    For pass = 0 To 1
        If pass = 0 Then trySide = side Else trySide = -side
        For extra = 0 To 12
            SideCentre x, y, dx, dy, trySide, width, height, gap + extra * stepOut, centreX, centreY
            centreY = centreY + stack
            If pass = 0 And extra = 0 Then firstX = centreX: firstY = centreY
            If Not shape.TouchesBox(measure - span, measure + span, centreX - width / 2# - margin, _
                                    centreY - height / 2# - margin, centreX + width / 2# + margin, _
                                    centreY + height / 2# + margin) Then Exit Sub
        Next extra
    Next pass
    centreX = firstX: centreY = firstY
End Sub

Private Sub PlaceSideLabel(ByVal shape As clsLrsGeometry, ByVal text As String, ByVal measure As Double, _
                           ByVal height As Double, ByVal line As Long)
    ' Places a label to the left of the direction of travel, clear of the route. line 1 is
    ' stacked on a first label at the same measure, on the side away from the route.
    Dim dx As Double, dy As Double, cx As Double, cy As Double, stack As Double
    shape.DirectionAt measure, dx, dy
    If dx >= 0# Then stack = line * 1.6 * height Else stack = -line * 1.6 * height
    BesideRoute shape, measure, -1#, LABEL_CHAR_WIDTH * height * Len(text), height, 0.8 * height, stack, cx, cy
    PlaceLabel text, cx, cy, height
End Sub

' ---------------------------------------------------------------------------
' Milepost signs
' ---------------------------------------------------------------------------
Public Function MilepostMeasures(ByVal shape As clsLrsGeometry, ByVal eventType As String, ByVal fromMeasure As Double, _
                                 ByVal toMeasure As Double, ByVal mode As String, ByRef measures() As Double) As Long
    ' Measures that get a milepost sign. Returns the count.
    Dim stepTenths As Long, firstTenth As Long, lastTenth As Long, tenth As Long, n As Long
    If mode = MILEPOST_NONE Then Exit Function
    If eventType = EVENT_POINT Then
        ReDim measures(0 To 0): measures(0) = fromMeasure
        MilepostMeasures = 1
        Exit Function
    End If
    If mode = MILEPOST_ENDS Then
        ReDim measures(0 To 1): measures(0) = fromMeasure: measures(1) = toMeasure
        MilepostMeasures = 2
        Exit Function
    End If
    stepTenths = CLng(Val(mode) * 10)
    firstTenth = Ceil(fromMeasure * 10 - 0.000001)
    lastTenth = Int(toMeasure * 10 + 0.000001)
    If lastTenth < firstTenth Then Exit Function
    ReDim measures(0 To lastTenth - firstTenth)
    For tenth = firstTenth To lastTenth
        If tenth Mod stepTenths = 0 Then
            If shape.HasMeasure(tenth / 10#) Then
                measures(n) = tenth / 10#: n = n + 1
            End If
        End If
    Next tenth
    MilepostMeasures = n
End Function

Private Function MilepostCell(ByVal x As Double, ByVal y As Double, ByVal digits As String, ByVal decimalText As String, _
                              ByVal unit As Double, ByVal green As Long, ByVal white As Long, _
                              ByVal signX As Double, ByVal signY As Double) As CellElement
    ' A milepost sign as one cell: green panel, white border and legend, on a post. The sign
    ' stands upright with its middle at (signX, signY); the post runs from the route point
    ' (routeX, routeY) to the sign.
    Dim routeX As Double, routeY As Double, centreY2 As Double
    Dim width As Double, height As Double, texts() As String, centreY() As Double, textHeight() As Double
    Dim textCount As Long, hasRule As Boolean, ruleY As Double, bottom As Double, border As Double
    Dim children() As Element, childCount As Long, pass As Long, inset As Double, colour As Long
    Dim px() As Double, py() As Double, n As Long, i As Long, vertices() As Point3d
    Dim post As LineElement, panel As ShapeElement, rule As LineElement, legend As TextElement

    MilepostLayout digits, decimalText, unit, width, height, texts, centreY, textHeight, textCount, hasRule, ruleY
    routeX = x: routeY = y
    x = signX: centreY2 = signY
    bottom = centreY2 - height / 2#

    Set post = CreateLineElement2(Nothing, Point3dFromXY(routeX, routeY), Point3dFromXY(x, centreY2))
    post.Color = green: post.LineWeight = 2
    AppendElement children, childCount, post

    border = 0.1 * unit
    For pass = 0 To 1
        If pass = 0 Then
            inset = 0#: colour = white
        Else
            inset = border: colour = green
        End If
        RoundedRectangle x, bottom + inset, width - 2 * inset, height - 2 * inset, 0.4 * unit - inset, px, py, n
        ReDim vertices(0 To n - 1)
        For i = 0 To n - 1
            vertices(i).x = px(i): vertices(i).y = py(i): vertices(i).Z = 0#
        Next i
        Set panel = CreateShapeElement1(Nothing, vertices, msdFillModeFilled)
        panel.Color = colour: panel.LineWeight = 0
        panel.FillColor = colour
        AppendElement children, childCount, panel
    Next pass

    If hasRule Then
        Set rule = CreateLineElement2(Nothing, Point3dFromXY(x - 0.3 * width, bottom + ruleY), _
                                      Point3dFromXY(x + 0.3 * width, bottom + ruleY))
        rule.Color = white: rule.LineWeight = 2
        AppendElement children, childCount, rule
    End If

    For i = 0 To textCount - 1
        Set legend = CreateTextElement1(Nothing, texts(i), Point3dFromXY(x, bottom + centreY(i)), Matrix3dIdentity)
        UseTrueTypeFont legend
        SizeText legend, textHeight(i), 0.85 * textHeight(i)
        legend.TextStyle.Justification = msdTextJustificationCenterCenter
        legend.TextStyle.IsBold = True
        legend.Color = white: legend.LineWeight = 0
        AppendElement children, childCount, legend
    Next i

    Set MilepostCell = CreateCellElement1("MILEPOST", children, Point3dFromXY(routeX, routeY))
End Function

Private Function PlaceMileposts(ByVal route As clsLrsRoute, ByVal shape As clsLrsGeometry, ByVal eventType As String, _
                                ByVal fromMeasure As Double, ByVal toMeasure As Double, ByVal mode As String, _
                                ByVal lengthMaster As Double) As Long
    Dim measures() As Double, n As Long, i As Long, signHeight As Double, unit As Double
    Dim xMin As Double, yMin As Double, xMax As Double, yMax As Double, x As Double, y As Double
    Dim green As Long, white As Long, digits As String, decimalText As String, signX As Double, signY As Double
    Dim width As Double, height As Double, texts() As String, centreY() As Double, textHeight() As Double
    Dim textCount As Long, hasRule As Boolean, ruleY As Double, sign As CellElement
    n = MilepostMeasures(route.Geometry, eventType, fromMeasure, toMeasure, mode, measures)
    If n = 0 Then Exit Function
    If eventType = EVENT_LINEAR And mode <> MILEPOST_ENDS Then
        signHeight = 0.35 * Val(mode) * lengthMaster / (toMeasure - fromMeasure)
    ElseIf lengthMaster > 0# Then
        signHeight = lengthMaster / 15#
    Else
        shape.Extent xMin, yMin, xMax, yMax
        signHeight = Hypot(xMax - xMin, yMax - yMin) / 30#
    End If
    signHeight = ClampD(signHeight, SIGN_HEIGHT_MIN_M * MasterPerMetre(), SIGN_HEIGHT_MAX_M * MasterPerMetre())
    unit = MILEPOST_SIZE * signHeight / 5.8          ' 5.8 units = two digits plus a decimal
    green = ActiveModelReference.InternalColorFromRGBColor(RGB(0, 110, 81))
    white = ActiveModelReference.InternalColorFromRGBColor(RGB(255, 255, 255))
    For i = 0 To n - 1
        shape.PointAt measures(i), x, y
        MilepostText measures(i), mode, digits, decimalText
        MilepostLayout digits, decimalText, unit, width, height, texts, centreY, textHeight, textCount, hasRule, ruleY
        BesideRoute shape, measures(i), 1#, width, height, 0.8 * unit, 0#, signX, signY
        Set sign = MilepostCell(x, y, digits, decimalText, unit, green, white, signX, signY)
        ActiveModelReference.AddElement sign
        TagGraphic sign, PART_MILEPOST
    Next i
    PlaceMileposts = n
End Function

' ---------------------------------------------------------------------------
' Create the event
' ---------------------------------------------------------------------------
Public Function CreateEventElements(ByVal route As clsLrsRoute, ByVal shape As clsLrsGeometry, ByVal eventType As String, _
                                    ByVal fromMeasure As Double, ByVal toMeasure As Double, ByVal labels As String, _
                                    ByVal mileposts As String, ByVal logger As Object) As Long
    ' Adds the event to the active model. shape is the route in master units.
    ' Returns the number of event elements (those that carry the LRS metadata).
    Dim measures() As Double, signCount As Long
    Dim elements() As Element, count As Long, overrideColor() As Long, hasMetadata() As Boolean
    Dim fromX As Double, fromY As Double, toX As Double, toY As Double, lengthMaster As Double
    Dim parts As Collection, part As Variant, pieces() As Element, pieceCount As Long, chain() As ChainableElement
    Dim i As Long, names() As String, values() As Variant, propertyCount As Long, target As ItemType
    Dim idText As String, eventCount As Long, metres As Double, endMeasure As Double
    Dim xMin As Double, yMin As Double, xMax As Double, yMax As Double, height As Double
    Dim placed As Long

    signCount = MilepostMeasures(route.Geometry, eventType, fromMeasure, toMeasure, mileposts, measures)
    If signCount > MAX_MILEPOSTS Then
        Err.Raise ERR_INPUT, "modLrsDraw", signCount & " milepost signs is more than the limit of " & _
                  MAX_MILEPOSTS & ". Choose a wider spacing or a shorter event."
    End If
    If Not Model3D() Then logger.Log "Active model is 2D - elements are created 2D."

    shape.PointAt fromMeasure, fromX, fromY
    toX = fromX: toY = fromY
    If eventType = EVENT_POINT Then
        AppendElement elements, count, PointElement(fromX, fromY, False, 0, 0)
        ReDim overrideColor(0 To 0), hasMetadata(0 To 0)
        overrideColor(0) = POINT_EVENT_COLOR: hasMetadata(0) = True
    Else
        Set parts = shape.Slice(fromMeasure, toMeasure)
        If parts.count = 0 Then
            Err.Raise ERR_INPUT, "modLrsDraw", "The route has no geometry between these measures."
        End If
        shape.PointAt toMeasure, toX, toY
        For Each part In parts
            lengthMaster = lengthMaster + PolylineLength(part)
            pieceCount = 0
            Erase pieces
            AppendLineStrings part, False, 0, 0, pieces, pieceCount
            If pieceCount = 1 Then
                AppendElement elements, count, pieces(0)
            Else
                ReDim chain(0 To pieceCount - 1)
                For i = 0 To pieceCount - 1
                    Set chain(i) = pieces(i)
                Next i
                AppendElement elements, count, CreateComplexStringElement1(chain)
            End If
        Next part
        ReDim overrideColor(0 To count + 1), hasMetadata(0 To count + 1)
        For i = 0 To count - 1
            overrideColor(i) = -1: hasMetadata(i) = True
        Next i
        overrideColor(count) = CREATED_FROM_COLOR: overrideColor(count + 1) = CREATED_TO_COLOR
        AppendElement elements, count, PointElement(fromX, fromY, False, 0, 0)
        AppendElement elements, count, PointElement(toX, toY, False, 0, 0)
    End If

    EventProperties route, eventType, fromMeasure, toMeasure, names, values, propertyCount
    Set target = EventItemType(names, values, propertyCount)
    For i = 0 To count - 1
        If overrideColor(i) >= 0 Then
            elements(i).Color = overrideColor(i)
            elements(i).LineWeight = CREATED_PIN_WEIGHT
        End If
        ActiveModelReference.AddElement elements(i)
        If hasMetadata(i) Then
            AttachMetadata elements(i), target, names, values, propertyCount
            idText = idText & IIf(idText = "", "", ", ") & DLongToString(elements(i).ID)
            eventCount = eventCount + 1
        ElseIf overrideColor(i) = CREATED_FROM_COLOR Then
            TagGraphic elements(i), PART_FROM
        Else
            TagGraphic elements(i), PART_TO
        End If
    Next i

    endMeasure = values(6)
    logger.Log "Added " & eventCount & " event element(s) on the active level: " & idText & "."
    If eventType = EVENT_LINEAR Then
        logger.Log "Added the start marker (green) and stop marker (red)."
        metres = MetresFromMaster(lengthMaster)
        logger.Log "Geometry length " & FmtThousands(metres, 1) & " m (" & FmtNum(metres / 1609.344, 3) & _
                   " mi) for a " & FmtNum(toMeasure - fromMeasure, 3) & " mi event."
    Else
        logger.Log "The point event is drawn yellow."
    End If
    logger.Log "Attached '" & ITEM_LIBRARY & " / " & ITEM_TYPE & "': Route=" & route.RouteName & ", Name=" & _
               route.NameId & ", StartMeasure=" & FmtMeasure(fromMeasure) & ", EndMeasure=" & FmtMeasure(endMeasure) & "."

    If labels <> LABELS_NONE Then
        shape.Extent xMin, yMin, xMax, yMax
        height = lengthMaster / 75#
        If height = 0# Then height = Hypot(xMax - xMin, yMax - yMin) / 150#
        height = ClampD(height, LABEL_HEIGHT_MIN_M * MasterPerMetre(), LABEL_HEIGHT_MAX_M * MasterPerMetre())
        On Error GoTo LabelFailed
        PlaceSideLabel shape, "MP " & FmtMeasure(fromMeasure), fromMeasure, height, 0
        If eventType = EVENT_LINEAR Then PlaceSideLabel shape, "MP " & FmtMeasure(toMeasure), toMeasure, height, 0
        If labels = LABELS_ALL Then
            If eventType = EVENT_LINEAR Then
                PlaceSideLabel shape, route.RouteName & "  " & route.NameId & "  (" & _
                               FmtNum(toMeasure - fromMeasure, 3) & " mi)", (fromMeasure + toMeasure) / 2#, height, 0
            Else
                PlaceSideLabel shape, route.RouteName & "  " & route.NameId, fromMeasure, height, 1
            End If
        End If
        logger.Log "Placed labels to the left of the direction of travel."
        GoTo LabelsDone
LabelFailed:
        logger.Log "WARNING: labels were not placed (" & Err.Description & ")."
        Resume LabelsDone
LabelsDone:
        On Error GoTo 0
    End If

    If mileposts <> MILEPOST_NONE Then
        On Error GoTo MilepostFailed
        placed = PlaceMileposts(route, shape, eventType, fromMeasure, toMeasure, mileposts, lengthMaster)
        If placed > 0 Then logger.Log "Placed " & placed & " milepost sign(s) to the right of the direction of travel."
        GoTo MilepostsDone
MilepostFailed:
        logger.Log "WARNING: milepost signs were not all placed (" & Err.Description & ")."
        Resume MilepostsDone
MilepostsDone:
        On Error GoTo 0
    End If

    ShowMessage "LRS event " & route.RouteName & " created (" & eventCount & " element(s)).", _
                route.NameId & "  measure " & FmtMeasure(fromMeasure) & " to " & FmtMeasure(endMeasure) & " mi", _
                msdMessageCenterPriorityInfo, False
    CreateEventElements = eventCount
End Function
