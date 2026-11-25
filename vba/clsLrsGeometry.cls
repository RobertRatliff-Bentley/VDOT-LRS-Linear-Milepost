Option Explicit
' Route polyline: parts of (x, y, m) vertices with m increasing along each part.
' Vertices of all parts are stored end to end; mStart/mEnd give each part's range.

Private mX() As Double, mY() As Double, mM() As Double
Private mStart() As Long, mEnd() As Long
Private mParts As Long, mCount As Long, mCapacity As Long
Public MMin As Double
Public MMax As Double

Public Sub AddPath(x() As Double, y() As Double, m() As Double, ByVal count As Long)
    ' Adds one part (arrays 0..count-1). Parts with fewer than 2 vertices are ignored;
    ' a part whose measures decrease is reversed.
    Dim i As Long, source As Long
    If count < 2 Then Exit Sub
    If mCount + count > mCapacity Then
        mCapacity = (mCount + count) * 2
        ReDim Preserve mX(0 To mCapacity - 1), mY(0 To mCapacity - 1), mM(0 To mCapacity - 1)
    End If
    ReDim Preserve mStart(0 To mParts), mEnd(0 To mParts)
    mStart(mParts) = mCount
    For i = 0 To count - 1
        If m(0) > m(count - 1) Then source = count - 1 - i Else source = i
        mX(mCount) = x(source): mY(mCount) = y(source): mM(mCount) = m(source)
        mCount = mCount + 1
    Next i
    mEnd(mParts) = mCount - 1
    mParts = mParts + 1
End Sub

Public Sub Finish()
    ' Call after the last AddPath: orders the parts by their first measure.
    Dim order() As Long, i As Long, j As Long, temp As Long, p As Long, k As Long, n As Long
    Dim nx() As Double, ny() As Double, nm() As Double, ns() As Long, ne() As Long
    If mParts = 0 Then Err.Raise ERR_INPUT, "clsLrsGeometry", "The route has no geometry with measures."
    ReDim order(0 To mParts - 1)
    For i = 0 To mParts - 1: order(i) = i: Next i
    For i = 1 To mParts - 1                     ' insertion sort, stable
        temp = order(i): j = i - 1
        Do While j >= 0
            If mM(mStart(order(j))) <= mM(mStart(temp)) Then Exit Do
            order(j + 1) = order(j): j = j - 1
        Loop
        order(j + 1) = temp
    Next i
    ReDim nx(0 To mCount - 1), ny(0 To mCount - 1), nm(0 To mCount - 1)
    ReDim ns(0 To mParts - 1), ne(0 To mParts - 1)
    For i = 0 To mParts - 1
        p = order(i)
        ns(i) = n
        For k = mStart(p) To mEnd(p)
            nx(n) = mX(k): ny(n) = mY(k): nm(n) = mM(k): n = n + 1
        Next k
        ne(i) = n - 1
    Next i
    mX = nx: mY = ny: mM = nm: mStart = ns: mEnd = ne
    mCapacity = mCount
    MMin = mM(mStart(0))
    MMax = mM(mEnd(0))
    For i = 1 To mParts - 1
        If mM(mEnd(i)) > MMax Then MMax = mM(mEnd(i))
    Next i
End Sub

Friend Sub InitRaw(x() As Double, y() As Double, m() As Double, starts() As Long, ends() As Long, _
                   ByVal parts As Long, ByVal count As Long, ByVal lowM As Double, ByVal highM As Double)
    mX = x: mY = y: mM = m: mStart = starts: mEnd = ends
    mParts = parts: mCount = count: mCapacity = count
    MMin = lowM: MMax = highM
End Sub

Public Function Mapped(xs() As Double, ys() As Double) As clsLrsGeometry
    ' The same route with new coordinates (xs, ys: 0..VertexCount-1).
    Dim result As clsLrsGeometry
    Set result = New clsLrsGeometry
    result.InitRaw xs, ys, mM, mStart, mEnd, mParts, mCount, MMin, MMax
    Set Mapped = result
End Function

Public Property Get PartCount() As Long
    PartCount = mParts
End Property

Public Property Get VertexCount() As Long
    VertexCount = mCount
End Property

Public Property Get PartStart(ByVal part As Long) As Long
    PartStart = mStart(part)
End Property

Public Property Get PartEnd(ByVal part As Long) As Long
    PartEnd = mEnd(part)
End Property

Public Property Get VX(ByVal index As Long) As Double
    VX = mX(index)
End Property

Public Property Get VY(ByVal index As Long) As Double
    VY = mY(index)
End Property

Public Property Get VM(ByVal index As Long) As Double
    VM = mM(index)
End Property

Private Sub Interpolate(ByVal part As Long, ByVal measure As Double, ByRef x As Double, ByRef y As Double)
    Dim low As Long, high As Long, middle As Long, index As Long, t As Double
    low = mStart(part): high = mEnd(part) + 1
    Do While low < high                          ' first vertex with m > measure
        middle = (low + high) \ 2
        If mM(middle) <= measure Then low = middle + 1 Else high = middle
    Loop
    index = low
    If index < mStart(part) + 1 Then index = mStart(part) + 1
    If index > mEnd(part) Then index = mEnd(part)
    If mM(index) = mM(index - 1) Then
        t = 0#
    Else
        t = (measure - mM(index - 1)) / (mM(index) - mM(index - 1))
        If t < 0# Then t = 0#
        If t > 1# Then t = 1#
    End If
    x = mX(index - 1) + (mX(index) - mX(index - 1)) * t
    y = mY(index - 1) + (mY(index) - mY(index - 1)) * t
End Sub

Public Sub PointAt(ByVal measure As Double, ByRef x As Double, ByRef y As Double)
    ' Location of a measure; a measure in a gap snaps to the nearest part end.
    Dim part As Long, best As Long, bestGap As Double, gap As Double, k As Long, index As Long
    For part = 0 To mParts - 1
        If mM(mStart(part)) <= measure And measure <= mM(mEnd(part)) Then
            Interpolate part, measure, x, y
            Exit Sub
        End If
    Next part
    best = -1
    For part = 0 To mParts - 1
        For k = 0 To 1
            If k = 0 Then index = mStart(part) Else index = mEnd(part)
            gap = Abs(mM(index) - measure)
            If best < 0 Or gap < bestGap Then best = index: bestGap = gap
        Next k
    Next part
    x = mX(best): y = mY(best)
End Sub

Public Function MeasurePerUnit() As Double
    ' Average measure covered by one unit of length along the route.
    Dim part As Long, i As Long, length As Double, covered As Double
    For part = 0 To mParts - 1
        covered = covered + mM(mEnd(part)) - mM(mStart(part))
        For i = mStart(part) + 1 To mEnd(part)
            length = length + Hypot(mX(i) - mX(i - 1), mY(i) - mY(i - 1))
        Next i
    Next part
    If length > 0# Then MeasurePerUnit = covered / length
End Function

Public Function TouchesBox(ByVal fromMeasure As Double, ByVal toMeasure As Double, ByVal xMin As Double, _
                           ByVal yMin As Double, ByVal xMax As Double, ByVal yMax As Double) As Boolean
    ' True when the route between (about) two measures touches the box.
    Dim part As Long, low As Long, high As Long, middle As Long, i As Long
    For part = 0 To mParts - 1
        If mM(mEnd(part)) >= fromMeasure And mM(mStart(part)) <= toMeasure Then
            low = mStart(part): high = mEnd(part)
            Do While low < high                  ' first vertex with m >= fromMeasure
                middle = (low + high) \ 2
                If mM(middle) < fromMeasure Then low = middle + 1 Else high = middle
            Loop
            If low > mStart(part) Then low = low - 1
            For i = low + 1 To mEnd(part)
                If SegmentHitsBox(mX(i - 1), mY(i - 1), mX(i), mY(i), xMin, yMin, xMax, yMax) Then
                    TouchesBox = True
                    Exit Function
                End If
                If mM(i) > toMeasure Then Exit For
            Next i
        End If
    Next part
End Function

Public Sub DirectionAt(ByVal measure As Double, ByRef dx As Double, ByRef dy As Double)
    ' Unit vector along the route (towards higher measures) at a measure.
    Dim part As Long, chosen As Long, bestGap As Double, gap As Double
    Dim low As Long, high As Long, middle As Long, index As Long, offset As Long, k As Long, i As Long
    Dim length As Double
    chosen = -1
    For part = 0 To mParts - 1
        If mM(mStart(part)) <= measure And measure <= mM(mEnd(part)) Then chosen = part: Exit For
    Next part
    If chosen < 0 Then
        For part = 0 To mParts - 1
            gap = MinD(Abs(mM(mStart(part)) - measure), Abs(mM(mEnd(part)) - measure))
            If chosen < 0 Or gap < bestGap Then chosen = part: bestGap = gap
        Next part
    End If
    low = mStart(chosen): high = mEnd(chosen) + 1
    Do While low < high                          ' first vertex with m > measure
        middle = (low + high) \ 2
        If mM(middle) <= measure Then low = middle + 1 Else high = middle
    Loop
    index = low
    If index < mStart(chosen) + 1 Then index = mStart(chosen) + 1
    If index > mEnd(chosen) Then index = mEnd(chosen)
    For offset = 0 To mEnd(chosen) - mStart(chosen)        ' skips zero-length segments
        For k = 0 To 1
            If k = 0 Then i = index + offset Else i = index - offset
            If i >= mStart(chosen) + 1 And i <= mEnd(chosen) Then
                length = Hypot(mX(i) - mX(i - 1), mY(i) - mY(i - 1))
                If length > 0# Then
                    dx = (mX(i) - mX(i - 1)) / length: dy = (mY(i) - mY(i - 1)) / length
                    Exit Sub
                End If
            End If
        Next k
    Next offset
    dx = 1#: dy = 0#
End Sub

Public Function Slice(ByVal fromMeasure As Double, ByVal toMeasure As Double) As Collection
    ' Parts of the route between two measures. Each item is a 2-D array (0..n-1, 0..1) of x, y.
    Dim result As Collection, part As Long, low As Double, high As Double
    Dim k As Long, n As Long, points() As Double, x As Double, y As Double
    Set result = New Collection
    For part = 0 To mParts - 1
        low = MaxD(fromMeasure, mM(mStart(part)))
        high = MinD(toMeasure, mM(mEnd(part)))
        If high > low Then
            ReDim points(0 To mEnd(part) - mStart(part) + 2, 0 To 1)
            Interpolate part, low, x, y
            points(0, 0) = x: points(0, 1) = y
            n = 1
            For k = mStart(part) To mEnd(part)
                If mM(k) > low And mM(k) < high Then
                    points(n, 0) = mX(k): points(n, 1) = mY(k): n = n + 1
                End If
            Next k
            Interpolate part, high, x, y
            points(n, 0) = x: points(n, 1) = y
            n = n + 1
            result.Add TrimPoints(points, n)
        End If
    Next part
    Set Slice = result
End Function

Private Function TrimPoints(points() As Double, ByVal count As Long) As Variant
    Dim result() As Double, i As Long
    ReDim result(0 To count - 1, 0 To 1)
    For i = 0 To count - 1
        result(i, 0) = points(i, 0): result(i, 1) = points(i, 1)
    Next i
    TrimPoints = result
End Function

Public Function SliceCount(ByVal fromMeasure As Double, ByVal toMeasure As Double) As Long
    Dim part As Long
    For part = 0 To mParts - 1
        If MinD(toMeasure, mM(mEnd(part))) > MaxD(fromMeasure, mM(mStart(part))) Then SliceCount = SliceCount + 1
    Next part
End Function

Public Sub Locate(ByVal x As Double, ByVal y As Double, ByRef measure As Double, ByRef px As Double, ByRef py As Double)
    ' Nearest location on the route.
    Dim part As Long, k As Long, dx As Double, dy As Double, length2 As Double, t As Double
    Dim qx As Double, qy As Double, distance2 As Double, best As Double, found As Boolean
    For part = 0 To mParts - 1
        For k = mStart(part) + 1 To mEnd(part)
            dx = mX(k) - mX(k - 1): dy = mY(k) - mY(k - 1)
            length2 = dx * dx + dy * dy
            If length2 = 0# Then
                t = 0#
            Else
                t = ((x - mX(k - 1)) * dx + (y - mY(k - 1)) * dy) / length2
                If t < 0# Then t = 0#
                If t > 1# Then t = 1#
            End If
            qx = mX(k - 1) + dx * t: qy = mY(k - 1) + dy * t
            distance2 = (x - qx) * (x - qx) + (y - qy) * (y - qy)
            If (Not found) Or distance2 < best Then
                found = True: best = distance2
                measure = mM(k - 1) + (mM(k) - mM(k - 1)) * t
                px = qx: py = qy
            End If
        Next k
    Next part
End Sub

Public Sub Extent(ByRef xMin As Double, ByRef yMin As Double, ByRef xMax As Double, ByRef yMax As Double)
    Dim i As Long
    xMin = mX(0): xMax = mX(0): yMin = mY(0): yMax = mY(0)
    For i = 1 To mCount - 1
        If mX(i) < xMin Then xMin = mX(i)
        If mX(i) > xMax Then xMax = mX(i)
        If mY(i) < yMin Then yMin = mY(i)
        If mY(i) > yMax Then yMax = mY(i)
    Next i
End Sub

Public Function HasMeasure(ByVal measure As Double) As Boolean
    Dim part As Long
    For part = 0 To mParts - 1
        If mM(mStart(part)) - 0.000001 <= measure And measure <= mM(mEnd(part)) + 0.000001 Then
            HasMeasure = True
            Exit Function
        End If
    Next part
End Function

Public Function Stops(ByRef values() As Double) As Long
    ' Measures offered in the drop-downs: part ends and every whole mile. Returns the count.
    Dim seen As Object, part As Long, mile As Long, key As Variant, n As Long
    Set seen = CreateObject("Scripting.Dictionary")
    For part = 0 To mParts - 1
        seen(FmtMeasure(mM(mStart(part)))) = Round(mM(mStart(part)), 3)
        seen(FmtMeasure(mM(mEnd(part)))) = Round(mM(mEnd(part)), 3)
        For mile = Ceil(mM(mStart(part))) To Int(mM(mEnd(part)))
            seen(FmtMeasure(mile)) = CDbl(mile)
        Next mile
    Next part
    ReDim values(0 To seen.count - 1)
    For Each key In seen.Keys
        values(n) = seen(key)
        If Abs(values(n)) < 0.0005 Then values(n) = 0#
        n = n + 1
    Next key
    SortDoubles values, 0, n - 1
    Stops = n
End Function
