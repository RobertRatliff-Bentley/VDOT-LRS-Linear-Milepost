Option Explicit
' One route from the LRS Route Master service.

Public NameId As String
Public CommonName As String
Public Jurisdiction As String
Public Category As String
Public Direction As String
Public OppositeName As String
Public OppositeExists As Boolean          ' most undivided roads name an opposite that is not a route
Public Geometry As clsLrsGeometry        ' x = longitude, y = latitude (WGS84)

Public Property Get RouteName() As String
    ' "SC-700E (Washington County)" -> "SC-700E"
    Dim cut As Long
    cut = InStr(CommonName, " (")
    If cut > 0 Then RouteName = Left$(CommonName, cut - 1) Else RouteName = CommonName
End Property
