Option Explicit
' VDOT LRS Route Master REST client (ArcGIS FeatureServer layer, polylines with M = miles).

' Interstates, US routes and most VA primaries are single statewide records with no
' jurisdiction, so they are added to a jurisdiction's list by location instead.
Private Const STATEWIDE As String = "(RTE_JURIS_PROPER_NM IS NULL OR RTE_JURIS_PROPER_NM = '')"
Private Const NOT_ROUTES As String = "RTE_CATEGORY_NM NOT LIKE 'RNS Datum%'"       ' internal LRS layers
Private Const ROUTE_FIELDS As String = "RTE_NM,RTE_COMMON_NM,RTE_NBR,RTE_CATEGORY_NM"

Public Url As String
' Optional. Its OnServiceRetry(attempt) is called before a retry and its OnServiceWaiting
' about 20 times a second while a request is in progress.
Public Listener As Object
Public Cancelled As Boolean      ' set to abandon the request in progress

' Replies already downloaded while this object lives; a failed request is never stored.
Private mJurisdictions As Variant
Private mRoutesCache As Object, mRouteCache As Object

Private Sub Class_Initialize()
    Url = SERVICE_URL
    Set mRoutesCache = CreateObject("Scripting.Dictionary")
    Set mRouteCache = CreateObject("Scripting.Dictionary")
End Sub

' ---------------------------------------------------------------------------
' Requests
' ---------------------------------------------------------------------------
Private Function UrlEncode(ByVal text As String) As String
    Dim i As Long, code As Long, result As String, c As String
    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        code = AscW(c)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) _
           Or c = "-" Or c = "_" Or c = "." Or c = "~" Then
            result = result & c
        ElseIf code = 32 Then
            result = result & "+"
        ElseIf code >= 0 And code < 128 Then
            result = result & "%" & Right$("0" & Hex$(code), 2)
        Else
            result = result & "%3F"          ' the service data is ASCII
        End If
    Next i
    UrlEncode = result
End Function

Private Function Body(ParamArray pairs() As Variant) As String
    ' Body("where", "1=1", "outFields", "*") -> "f=json&where=1%3D1&outFields=%2A"
    Dim i As Long, result As String
    result = "f=json"
    For i = LBound(pairs) To UBound(pairs) - 1 Step 2
        result = result & "&" & CStr(pairs(i)) & "=" & UrlEncode(CStr(pairs(i + 1)))
    Next i
    Body = result
End Function

Private Function Query(ByVal requestBody As String) As Object
    ' Posts one query and returns the parsed reply. A dropped connection is retried twice.
    ' The request is asynchronous: while it runs, MicroStation and the dialog keep handling
    ' their own events, so neither freezes.
    Dim attempt As Long, http As Object, failure As String, replyText As String, payload As Object
    Dim status As Long, finished As Boolean, started As Double, elapsed As Double
    For attempt = 0 To 2
        failure = ""
        On Error Resume Next
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.setTimeouts 15000, 15000, 45000, 45000
        http.Open "POST", Url & "/query", True
        http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.setRequestHeader "User-Agent", "MicroStation-VDOT-LRS-VBA/1.0"
        http.send requestBody
        If Err.Number <> 0 Then failure = Err.Description
        started = Timer
        Do While failure = ""
            finished = http.waitForResponse(0.05)
            If Err.Number <> 0 Then failure = Err.Description: Exit Do
            If finished Then Exit Do
            DoEvents
            If Not Listener Is Nothing Then Listener.OnServiceWaiting
            elapsed = Timer - started
            If elapsed < 0# Then elapsed = elapsed + 86400#         ' midnight
            If Cancelled Then http.abort: Exit Do
            If elapsed > 75# Then http.abort: failure = "the request timed out"
        Loop
        If failure = "" And Not Cancelled Then
            status = http.status
            replyText = http.responseText
            If Err.Number <> 0 Then failure = Err.Description
        End If
        On Error GoTo 0
        If Cancelled Then Err.Raise ERR_CANCELLED, "clsLrsService", "The request was cancelled."
        If failure = "" Then Exit For
        If attempt = 2 Then
            Err.Raise ERR_SERVICE, "clsLrsService", "Could not reach the VDOT LRS service: " & Trim$(failure)
        End If
        If Not Listener Is Nothing Then Listener.OnServiceRetry attempt + 1
        PauseSeconds 1# + attempt
        If Cancelled Then Err.Raise ERR_CANCELLED, "clsLrsService", "The request was cancelled."
    Next attempt
    If status <> 200 Then
        Err.Raise ERR_SERVICE, "clsLrsService", "Could not reach the VDOT LRS service: HTTP " & status
    End If

    On Error Resume Next
    Set payload = JsonParse(replyText)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise ERR_DATA, "clsLrsService", "The VDOT LRS service returned an unreadable reply."
    End If
    On Error GoTo 0
    If payload.Exists("error") Then
        Err.Raise ERR_DATA, "clsLrsService", "VDOT LRS service error " & JText(payload("error"), "code") & _
                  ": " & JText(payload("error"), "message")
    End If
    Set Query = payload
End Function

Private Function Features(ByVal requestBody As String) As Collection
    ' All features of a query, following the service's paging.
    Dim result As Collection, payload As Object, page As Collection, item As Variant, offset As Long
    Dim more As Boolean
    Set result = New Collection
    Do
        Set payload = Query(requestBody & "&resultOffset=" & offset)
        If Not payload.Exists("features") Then Exit Do
        Set page = payload("features")
        For Each item In page
            result.Add item
        Next item
        more = False
        If payload.Exists("exceededTransferLimit") Then
            If Not IsNull(payload("exceededTransferLimit")) Then more = CBool(payload("exceededTransferLimit"))
        End If
        If (Not more) Or page.count = 0 Then Exit Do
        offset = offset + page.count
    Loop
    Set Features = result
End Function

Private Function Quote(ByVal text As String) As String
    Quote = "'" & Replace(text, "'", "''") & "'"
End Function

Private Function JurisdictionWhere(ByVal jurisdictionName As String) As String
    If jurisdictionName <> "" Then
        JurisdictionWhere = "RTE_JURIS_PROPER_NM = " & Quote(jurisdictionName)
    Else
        JurisdictionWhere = STATEWIDE
    End If
End Function

' ---------------------------------------------------------------------------
' Jurisdictions
' ---------------------------------------------------------------------------
Public Function Jurisdictions(ByRef names() As String) As Long
    ' Distinct jurisdiction names, sorted; "" (no jurisdiction) comes first. Returns the count.
    Dim rows As Collection, row As Variant, seen As Object, key As Variant, n As Long, text As String
    If Not IsEmpty(mJurisdictions) Then
        names = mJurisdictions
        Jurisdictions = UBound(names) + 1
        Exit Function
    End If
    Set rows = Features(Body("where", "1=1", "outFields", "RTE_JURIS_PROPER_NM", _
                             "returnDistinctValues", "true", "orderByFields", "RTE_JURIS_PROPER_NM", _
                             "returnGeometry", "false"))
    Set seen = CreateObject("Scripting.Dictionary")
    For Each row In rows
        text = Trim$(JText(row("attributes"), "RTE_JURIS_PROPER_NM"))
        seen(text) = True
    Next row
    If seen.count = 0 Then Exit Function
    ReDim names(0 To seen.count - 1)
    For Each key In seen.Keys
        names(n) = CStr(key): n = n + 1
    Next key
    SortStrings names, 0, n - 1
    mJurisdictions = names
    Jurisdictions = n
End Function

' ---------------------------------------------------------------------------
' Route list
' ---------------------------------------------------------------------------
Private Function HighwayOrder(ByVal categoryName As String) As Long
    Select Case categoryName
        Case "Interstate": HighwayOrder = 0
        Case "US Highway Primary": HighwayOrder = 1
        Case "State Highway Primary": HighwayOrder = 2
        Case Else: HighwayOrder = 3
    End Select
End Function

Private Sub AppendRows(ByVal rows As Collection, ByVal highwaysFirst As Boolean, ByRef names() As String, _
                       ByRef ids() As String, ByRef total As Long)
    ' Appends (common name, route name ID) of each feature; optionally Interstate, US, VA first.
    Dim n As Long, i As Long, row As Variant, attributes As Object
    Dim rowName() As String, rowId() As String, keys() As String, order() As Long, number As String
    n = rows.count
    If n = 0 Then Exit Sub
    ReDim rowName(0 To n - 1), rowId(0 To n - 1), keys(0 To n - 1), order(0 To n - 1)
    For Each row In rows
        Set attributes = row("attributes")
        rowId(i) = JText(attributes, "RTE_NM")
        rowName(i) = JText(attributes, "RTE_COMMON_NM")
        If rowName(i) = "" Then rowName(i) = rowId(i)
        If highwaysFirst Then
            If IsNull(attributes("RTE_NBR")) Then
                number = "1000000000"
            Else
                number = Right$("0000000000" & CStr(CLng(attributes("RTE_NBR"))), 10)
            End If
            ' Chr(1) sorts below every character of a name, so "US-1N" stays ahead of
            ' "US-1N Ramp"; the sequence keeps rows with equal keys in service order.
            keys(i) = HighwayOrder(JText(attributes, "RTE_CATEGORY_NM")) & number & rowName(i) & _
                      Chr$(1) & Right$("000000" & i, 7)
        End If
        order(i) = i
        i = i + 1
    Next row
    If highwaysFirst Then SortIndexByKey keys, order, 0, n - 1
    ReDim Preserve names(0 To total + n - 1), ids(0 To total + n - 1)
    For i = 0 To n - 1
        names(total + i) = rowName(order(i))
        ids(total + i) = rowId(order(i))
    Next i
    total = total + n
End Sub

Private Function CoreExtent(values() As Double, ByVal count As Long, ByRef low As Double, ByRef high As Double) As Boolean
    ' Range of the values ignoring far-off strays. A few records carry the wrong jurisdiction
    ' (one "City of Richmond" street lies near Roanoke); a plain range would span half the state.
    Dim p2 As Double, p98 As Double, pad As Double, i As Long
    If count = 0 Then Exit Function
    SortDoubles values, 0, count - 1
    p2 = values(Int(0.02 * (count - 1)))
    p98 = values(CLng(Round(0.98 * (count - 1), 0)))
    pad = MaxD(0.5 * (p98 - p2), 0.02)            ' 0.02 degrees is about a mile
    For i = 0 To count - 1
        If values(i) >= p2 - pad Then low = values(i): Exit For
    Next i
    For i = count - 1 To 0 Step -1
        If values(i) <= p98 + pad Then high = values(i): Exit For
    Next i
    CoreExtent = True
End Function

Public Sub Routes(ByVal jurisdictionName As String, ByRef names() As String, ByRef ids() As String, _
                  ByRef total As Long, ByRef statewideCount As Long)
    ' Routes offered for a jurisdiction. The statewide routes (Interstate, US, VA, then ramps
    ' and frontage roads) that cross the jurisdiction's bounding box come first, then the
    ' jurisdiction's own routes in route-number order. "" lists every statewide route.
    Dim localRows As Collection, statewideRows As Collection, row As Variant, path As Variant
    Dim xs() As Double, ys() As Double, n As Long, capacity As Long, k As Long, vertex As Variant
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, cached As Variant
    total = 0: statewideCount = 0
    Erase names: Erase ids
    If mRoutesCache.Exists(jurisdictionName) Then
        cached = mRoutesCache(jurisdictionName)
        names = cached(0): ids = cached(1): total = cached(2): statewideCount = cached(3)
        Exit Sub
    End If

    If jurisdictionName = "" Then
        Set statewideRows = Features(Body("where", STATEWIDE & " AND " & NOT_ROUTES, "outFields", ROUTE_FIELDS, _
                                          "orderByFields", "RTE_NBR,RTE_COMMON_NM", "returnGeometry", "false"))
        AppendRows statewideRows, True, names, ids, total
        statewideCount = total
        If total > 0 Then mRoutesCache(jurisdictionName) = Array(names, ids, total, statewideCount)
        Exit Sub
    End If

    ' The jurisdiction's routes come back with heavily simplified geometry, used only to
    ' work out the area the jurisdiction covers.
    Set localRows = Features(Body("where", JurisdictionWhere(jurisdictionName), "outFields", ROUTE_FIELDS, _
                                  "orderByFields", "RTE_NBR,RTE_COMMON_NM", "returnGeometry", "true", _
                                  "outSR", "4326", "maxAllowableOffset", "0.05", "geometryPrecision", "4"))
    capacity = 1024
    ReDim xs(0 To capacity - 1), ys(0 To capacity - 1)
    For Each row In localRows
        If row.Exists("geometry") Then
            If row("geometry").Exists("paths") Then
                For Each path In row("geometry")("paths")
                    If path.count > 0 Then
                        For k = 0 To 1
                            If k = 0 Then Set vertex = path(1) Else Set vertex = path(path.count)
                            If n >= capacity Then
                                capacity = capacity * 2
                                ReDim Preserve xs(0 To capacity - 1), ys(0 To capacity - 1)
                            End If
                            xs(n) = vertex(1): ys(n) = vertex(2): n = n + 1
                        Next k
                    End If
                Next path
            End If
        End If
    Next row

    If CoreExtent(xs, n, xMin, xMax) Then
        CoreExtent ys, n, yMin, yMax
        Set statewideRows = Features(Body("where", STATEWIDE & " AND " & NOT_ROUTES, "outFields", ROUTE_FIELDS, _
                                          "geometry", NumStr(xMin) & "," & NumStr(yMin) & "," & NumStr(xMax) & "," & NumStr(yMax), _
                                          "geometryType", "esriGeometryEnvelope", "inSR", "4326", _
                                          "spatialRel", "esriSpatialRelIntersects", _
                                          "orderByFields", "RTE_NBR,RTE_COMMON_NM", "returnGeometry", "false"))
        AppendRows statewideRows, True, names, ids, total
        statewideCount = total
    End If
    AppendRows localRows, False, names, ids, total
    If total > 0 Then mRoutesCache(jurisdictionName) = Array(names, ids, total, statewideCount)
End Sub

' ---------------------------------------------------------------------------
' One route with its geometry
' ---------------------------------------------------------------------------
Public Function Route(ByVal nameId As String) As clsLrsRoute
    Dim payload As Object, rows As Collection, row As Variant, path As Variant, vertex As Variant
    Dim attributes As Object, result As clsLrsRoute, shape As clsLrsGeometry
    Dim x() As Double, y() As Double, m() As Double, n As Long, countReply As Object
    If mRouteCache.Exists(nameId) Then
        Set Route = mRouteCache(nameId)
        Exit Function
    End If
    Set payload = Query(Body("where", "RTE_NM = " & Quote(nameId), "returnGeometry", "true", "returnM", "true", _
                             "outSR", "4326", "outFields", _
                             "RTE_NM,RTE_COMMON_NM,RTE_JURIS_PROPER_NM,RTE_CATEGORY_NM,RTE_DIRECTION_CD,RTE_OPPOSITE_DIRECTION_RTE_NM"))
    If payload.Exists("features") Then Set rows = payload("features") Else Set rows = New Collection
    If rows.count = 0 Then
        Err.Raise ERR_DATA, "clsLrsService", "Route " & nameId & " was not found in the LRS service."
    End If

    Set shape = New clsLrsGeometry
    For Each row In rows
        If row.Exists("geometry") Then
            If row("geometry").Exists("paths") Then
                For Each path In row("geometry")("paths")
                    If path.count > 0 Then
                        ReDim x(0 To path.count - 1), y(0 To path.count - 1), m(0 To path.count - 1)
                        n = 0
                        For Each vertex In path
                            If vertex.count > 2 Then
                                If Not IsNull(vertex(3)) Then
                                    x(n) = vertex(1): y(n) = vertex(2): m(n) = vertex(3): n = n + 1
                                End If
                            End If
                        Next vertex
                        shape.AddPath x, y, m, n
                    End If
                Next path
            End If
        End If
    Next row
    On Error Resume Next
    shape.Finish
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise ERR_DATA, "clsLrsService", "Route " & nameId & ": The route has no geometry with measures."
    End If
    On Error GoTo 0

    Set attributes = rows(1)("attributes")
    Set result = New clsLrsRoute
    result.NameId = JText(attributes, "RTE_NM")
    result.CommonName = JText(attributes, "RTE_COMMON_NM")
    If result.CommonName = "" Then result.CommonName = result.NameId
    result.Jurisdiction = JText(attributes, "RTE_JURIS_PROPER_NM")
    result.Category = JText(attributes, "RTE_CATEGORY_NM")
    result.Direction = JText(attributes, "RTE_DIRECTION_CD")
    result.OppositeName = JText(attributes, "RTE_OPPOSITE_DIRECTION_RTE_NM")
    Set result.Geometry = shape
    If result.OppositeName <> "" Then
        Set countReply = Query(Body("where", "RTE_NM = " & Quote(result.OppositeName), _
                                    "returnCountOnly", "true", "returnGeometry", "false"))
        If countReply.Exists("count") Then result.OppositeExists = (CLng(countReply("count")) > 0)
    End If
    Set mRouteCache(nameId) = result
    Set Route = result
End Function
