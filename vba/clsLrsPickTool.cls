Option Explicit
' Pick tool: while it is active the route and event are drawn in the view as dynamics.
' A data point sets the measure nearest the cursor (From first, then To); reset ends it.

Implements IPrimitiveCommandEvents

Private mReportedError As Boolean

Private Sub IPrimitiveCommandEvents_Start()
    gLrsToolActive = True
    ShowCommand "VDOT LRS Event"
    If gLrsFormOpen Then
        ShowPrompt frmLrsEvent.PickPrompt()
        frmLrsEvent.OnToolStarted
    End If
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    On Error GoTo Failed
    If Not gLrsFormOpen Then Exit Sub
    frmLrsEvent.OnPick Point.X, Point.Y
    ShowPrompt frmLrsEvent.PickPrompt()
    Exit Sub
Failed:
    If gLrsFormOpen Then frmLrsEvent.Log "ERROR while picking: " & Err.Description
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
    On Error GoTo Failed
    If Not gLrsFormOpen Then Exit Sub
    frmLrsEvent.DrawDynamics Point.X, Point.Y, DrawMode
    Exit Sub
Failed:
    If Not mReportedError Then
        mReportedError = True
        If gLrsFormOpen Then frmLrsEvent.Log "ERROR in the drawing preview: " & Err.Description
    End If
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    CommandState.StartDefaultCommand
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
    gLrsToolActive = False
    If gLrsFormOpen Then frmLrsEvent.OnToolEnded
End Sub
