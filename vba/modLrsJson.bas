Option Explicit
' Small JSON reader. Objects become Scripting.Dictionary, arrays become Collection
' (1-based), null becomes Null. Raises an error on anything that is not JSON.

Private mText As String
Private mB() As Byte          ' the text as UTF-16 bytes
Private mN As Long            ' length in characters
Private mP As Long            ' position in characters, 0-based

Public Function JsonParse(ByRef text As String) As Object
    mText = text
    mN = Len(mText)
    If mN = 0 Then Err.Raise ERR_SERVICE, "JsonParse", "Empty reply."
    mB = mText
    mP = 0
    SkipWs
    If Ch() <> 123 Then Err.Raise ERR_SERVICE, "JsonParse", "The reply is not a JSON object."
    Set JsonParse = ParseObject()
    mText = vbNullString
    Erase mB
End Function

Public Function JText(ByVal dict As Object, ByVal key As String) As String
    ' String value of a key; "" when the key is missing or null.
    If dict.Exists(key) Then
        If Not IsNull(dict(key)) Then JText = CStr(dict(key))
    End If
End Function

Private Function Ch() As Long
    If mP >= mN Then
        Ch = -1
    Else
        Ch = mB(mP + mP) Or (CLng(mB(mP + mP + 1)) * 256&)
    End If
End Function

Private Sub SkipWs()
    Dim c As Long
    Do While mP < mN
        If mB(mP + mP + 1) <> 0 Then Exit Do
        c = mB(mP + mP)
        If c = 32 Or c = 9 Or c = 10 Or c = 13 Then mP = mP + 1 Else Exit Do
    Loop
End Sub

Private Sub Fail(ByVal what As String)
    Err.Raise ERR_SERVICE, "JsonParse", "Unexpected " & what & " at character " & mP & "."
End Sub

Private Function ParseScalar() As Variant
    ' Anything that is not an object or an array.
    Select Case Ch()
        Case 34: ParseScalar = ParseString()
        Case 116: mP = mP + 4: ParseScalar = True          ' true
        Case 102: mP = mP + 5: ParseScalar = False         ' false
        Case 110: mP = mP + 4: ParseScalar = Null          ' null
        Case 78: mP = mP + 3: ParseScalar = Null           ' NaN
        Case -1: Fail "end of text"
        Case Else: ParseScalar = ParseNumber()
    End Select
End Function

Private Function ParseObject() As Object
    Dim dict As Object, key As String, c As Long
    Set dict = CreateObject("Scripting.Dictionary")
    mP = mP + 1
    SkipWs
    If Ch() = 125 Then
        mP = mP + 1
    Else
        Do
            SkipWs
            If Ch() <> 34 Then Fail "character (a name was expected)"
            key = ParseString()
            SkipWs
            If Ch() <> 58 Then Fail "character (':' was expected)"
            mP = mP + 1
            SkipWs
            c = Ch()
            If c = 123 Then
                Set dict(key) = ParseObject()
            ElseIf c = 91 Then
                Set dict(key) = ParseArray()
            Else
                dict(key) = ParseScalar()
            End If
            SkipWs
            c = Ch()
            mP = mP + 1
            If c = 125 Then Exit Do
            If c <> 44 Then Fail "character in an object"
        Loop
    End If
    Set ParseObject = dict
End Function

Private Function ParseArray() As Collection
    Dim items As Collection, c As Long
    Set items = New Collection
    mP = mP + 1
    SkipWs
    If Ch() = 93 Then
        mP = mP + 1
    Else
        Do
            SkipWs
            c = Ch()
            If c = 123 Then
                items.Add ParseObject()
            ElseIf c = 91 Then
                items.Add ParseArray()
            Else
                items.Add ParseScalar()
            End If
            SkipWs
            c = Ch()
            mP = mP + 1
            If c = 93 Then Exit Do
            If c <> 44 Then Fail "character in an array"
        Loop
    End If
    Set ParseArray = items
End Function

Private Function ParseNumber() As Double
    Dim start As Long, c As Long
    start = mP
    Do While mP < mN
        If mB(mP + mP + 1) <> 0 Then Exit Do
        c = mB(mP + mP)
        If (c >= 48 And c <= 57) Or c = 45 Or c = 43 Or c = 46 Or c = 101 Or c = 69 Then
            mP = mP + 1
        Else
            Exit Do
        End If
    Loop
    If mP = start Then Fail "character"
    ParseNumber = Val(Mid$(mText, start + 1, mP - start))
End Function

Private Function ParseString() As String
    Dim start As Long, c As Long, hasEscape As Boolean, text As String
    mP = mP + 1
    start = mP
    Do
        If mP >= mN Then Fail "end of text in a string"
        c = Ch()
        If c = 34 Then Exit Do
        If c = 92 Then hasEscape = True: mP = mP + 1
        mP = mP + 1
    Loop
    text = Mid$(mText, start + 1, mP - start)
    mP = mP + 1
    If hasEscape Then text = Unescape(text)
    ParseString = text
End Function

Private Function Unescape(ByVal text As String) As String
    Dim i As Long, result As String, c As String
    i = 1
    Do While i <= Len(text)
        c = Mid$(text, i, 1)
        If c = "\" And i < Len(text) Then
            i = i + 1
            Select Case Mid$(text, i, 1)
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "b": result = result & Chr$(8)
                Case "f": result = result & Chr$(12)
                Case "u"
                    result = result & ChrW$(CLng("&H" & Mid$(text, i + 1, 4)))
                    i = i + 4
                Case Else: result = result & Mid$(text, i, 1)      ' \" \\ \/
            End Select
        Else
            result = result & c
        End If
        i = i + 1
    Loop
    Unescape = result
End Function
