Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    Application.EnableEvents = False
    If Not Intersect(Target, Me.Range("C12,C13,C14")) Is Nothing Then
        RedibujarInputs_FJS Me
        GoTo SafeExit
    End If
    Dim ckRng As Range
    Set ckRng = FJS_CkRange(Me)
    If Not ckRng Is Nothing Then
        If Not Intersect(Target, ckRng) Is Nothing Then
            FJS_RebuildDecisionTable Me
        End If
    End If
SafeExit:
    Application.EnableEvents = True
End Sub
