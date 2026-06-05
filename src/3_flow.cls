Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit

    If Intersect(Target, Me.Range("C12,C13")) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    RedibujarInputs_FS Me

SafeExit:
    Application.EnableEvents = True
End Sub
