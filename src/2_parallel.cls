Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit

    Application.EnableEvents = False

    If Not Intersect(Target, Me.Range("C12,C13")) Is Nothing Then
        RedibujarDistribucion_PM Me
        GoTo SafeExit
    End If

    Dim n As Long, m As Long
    If PM_ReadNM(Me, n, m) Then
        Dim rngDistribucion As Range
        Set rngDistribucion = Me.Range(Me.Cells(18, 3), Me.Cells(18 + m - 1, 3))

        If Not Intersect(Target, rngDistribucion) Is Nothing Then
            PM_IntentarMostrarResto
            GoTo SafeExit
        End If
    End If

SafeExit:
    Application.EnableEvents = True
End Sub
