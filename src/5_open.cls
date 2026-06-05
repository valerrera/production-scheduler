Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    Application.EnableEvents = False
    If Not Intersect(Target, Me.Range("C12,C13,C14")) Is Nothing Then
        RedibujarInputs_OS Me
    End If
SafeExit:
    Application.EnableEvents = True
End Sub
