Private Sub Worksheet_Change(ByVal Target As Range)
     If Intersect(Target, Me.Range("C12")) Is Nothing Then Exit Sub
     Application.EnableEvents = False
     RedibujarInputs_SM Me
     Application.EnableEvents = True
 End Sub
