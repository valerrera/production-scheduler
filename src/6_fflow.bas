Option Explicit

Private Const SHEET_NAME As String = "FlexibleFlowShop"

Private Const CELL_N As String = "C12"
Private Const CELL_M As String = "C13"
Private Const CELL_FFS_READY As String = "Z2"

Private Const DYNAMIC_TOP_ROW As Long = 15
Private Const INPUT_CLEAR_ROWS As Long = 1500

Private Const PAR_C0 As Long = 2
Private Const DEC_C0 As Long = 2

Private Const WS_COL_LABEL As Long = 2
Private Const WS_COL_C As Long = 3
Private Const WS_COL_RMAQ As Long = 4

Private Const MAX_JOBS As Long = 400
Private Const MAX_WS As Long = 50
Private Const MAX_MACHS_PER_WS As Long = 50

Private Const CHART_NAME_TIMELINE As String = "chTimeline_FFS"

Private Type FFSLayout
    parTitle As Long: parHeader As Long: parFirst As Long: parLast As Long
    wsTitle As Long: wsHeader As Long: wsFirst As Long: wsLast As Long
    decTitle As Long: decHeader As Long: decFirst As Long: decLast As Long
    instrLoad As Long: instrGen As Long
End Type

' CACHE
Private ffsLoaded As Boolean
Private ffsN As Long, ffsM As Long, ffsTotalMach As Long
Private ffsSeq() As Long
Private ffsJob() As String
Private ffsR() As Double, ffsD() As Double, ffsW() As Double
Private ffsP() As Double, ffsS() As Double
Private ffsCk() As Long, ffsRmaq() As Double

' =========================================================
Public Sub RedibujarInputs_FFS(ByVal ws As Worksheet)
    On Error GoTo ErrH
    Dim n As Long, m As Long
    If Not FFS_ReadNM(ws, n, m) Then Exit Sub
    ws.Unprotect
    ffsLoaded = False: ffsN = 0: ffsM = 0: ffsTotalMach = 0
    ws.Range(CELL_FFS_READY).Value = ""
    FFS_ClearDynamicZone ws: FFS_DeleteChartIfExists ws

    Dim L As FFSLayout: FFS_GetLayout n, m, L

    ' Tabla 1: parámetros por job
    FFS_SectionTitle ws, L.parTitle, PAR_C0, "2. Parámetros por job [PARÁMETRO] (r, p1..pm, s1..sm, d, w)."
    ws.Cells(L.parHeader, PAR_C0).Value = "Job": ws.Cells(L.parHeader, PAR_C0 + 1).Value = "r"
    Dim k As Long
    For k = 1 To m: ws.Cells(L.parHeader, PAR_C0 + 1 + k).Value = "p" & k: Next k
    For k = 1 To m: ws.Cells(L.parHeader, PAR_C0 + 1 + m + k).Value = "s" & k: Next k
    ws.Cells(L.parHeader, PAR_C0 + 2 + 2 * m).Value = "d"
    ws.Cells(L.parHeader, PAR_C0 + 3 + 2 * m).Value = "w"
    Dim lastColPar As Long: lastColPar = PAR_C0 + 3 + 2 * m
    FFS_FormatHeaderRow ws, L.parHeader, PAR_C0, lastColPar
    FFS_FormatEditableBlock ws, L.parFirst, L.parLast, PAR_C0, lastColPar, RGB(248, 248, 248)
    Dim i As Long
    For i = 1 To n
        ws.Cells(L.parFirst + i - 1, PAR_C0).Value = "J" & i
        ws.Cells(L.parFirst + i - 1, PAR_C0 + 3 + 2 * m).Value = 1
    Next i
    ws.Range(ws.Cells(L.parFirst, PAR_C0), ws.Cells(L.parLast, PAR_C0)).Locked = True
    FFS_SetJobListNamedRange ws, L.parFirst, L.parLast

    ' Tabla 2: workstations (c, rmaq)
    FFS_SectionTitle ws, L.wsTitle, WS_COL_LABEL, "3. Workstations [PARÁMETRO] — # máquinas idénticas y rmaq por workstation."
    ws.Cells(L.wsHeader, WS_COL_LABEL).Value = "Workstation"
    ws.Cells(L.wsHeader, WS_COL_C).Value = "c (# máquinas)"
    ws.Cells(L.wsHeader, WS_COL_RMAQ).Value = "rmaq"
    FFS_FormatHeaderRow ws, L.wsHeader, WS_COL_LABEL, WS_COL_RMAQ
    FFS_FormatEditableBlock ws, L.wsFirst, L.wsLast, WS_COL_LABEL, WS_COL_RMAQ, RGB(248, 248, 248)
    For k = 1 To m
        ws.Cells(L.wsFirst + k - 1, WS_COL_LABEL).Value = "M" & k
        ws.Cells(L.wsFirst + k - 1, WS_COL_C).Value = 1
    Next k
    ws.Range(ws.Cells(L.wsFirst, WS_COL_LABEL), ws.Cells(L.wsLast, WS_COL_LABEL)).Locked = True

    ' Tabla 3: decisión (permutación; FIFO determina máquina dentro del workstation)
    FFS_SectionTitle ws, L.decTitle, DEC_C0, "4. Decisión [DECISIÓN] — escribe la permutación de jobs (mismo orden en todos los workstations; la asignación a máquina específica es FIFO)."
    ws.Cells(L.decHeader, DEC_C0).Value = "Secuencia"
    ws.Cells(L.decHeader, DEC_C0 + 1).Value = "Job"
    FFS_FormatHeaderRow ws, L.decHeader, DEC_C0, DEC_C0 + 1
    FFS_FormatEditableBlock ws, L.decFirst, L.decLast, DEC_C0, DEC_C0 + 1, 0
    Dim rngDV As Range
    Set rngDV = ws.Range(ws.Cells(L.decFirst, DEC_C0 + 1), ws.Cells(L.decLast, DEC_C0 + 1))
    On Error Resume Next: rngDV.Validation.Delete: On Error GoTo 0
    rngDV.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=FFS_JobList"
    rngDV.Validation.IgnoreBlank = True: rngDV.Validation.InCellDropdown = True
    For i = 1 To n: ws.Cells(L.decFirst + i - 1, DEC_C0).Value = i: Next i

    With ws.Cells(L.instrLoad, DEC_C0)
        .Value = "5. Presione 'Cargar datos' para validar la información."
        .Font.Bold = True: .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en RedibujarInputs_FFS: " & Err.Description, vbExclamation: Resume Salir
End Sub

Public Sub FFS_CargarDatos()
    On Error GoTo ErrH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Dim n As Long, m As Long
    If Not FFS_ReadNM(ws, n, m) Then
        MsgBox "Verifica que # jobs y # workstations sean enteros positivos.", vbExclamation: Exit Sub
    End If
    ws.Unprotect
    Dim L As FFSLayout: FFS_GetLayout n, m, L
    FFS_ClearOutputArea ws, L: FFS_DeleteChartIfExists ws
    ws.Range(CELL_FFS_READY).Value = ""

    Dim warn As String
    If Not FFS_ValidateInputs(ws, n, m, L, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos - Flexible Flow Shop": GoTo Salir
    End If
    FFS_LoadInputsToCache ws, n, m, L
    ws.Range(CELL_FFS_READY).Value = "OK"

    With ws.Cells(L.instrGen, DEC_C0)
        .Value = "6. Datos válidos. Presione 'Generar outputs'."
        .Font.Bold = True: .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en FFS_CargarDatos: " & Err.Description, vbExclamation: Resume Salir
End Sub

Public Sub FFS_GenerarOutputs()
    On Error GoTo ErrH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If UCase$(Trim$(CStr(ws.Range(CELL_FFS_READY).Value))) <> "OK" Then
        MsgBox "Primero presione 'Cargar datos'.", vbExclamation: Exit Sub
    End If
    Dim n As Long, m As Long
    If Not FFS_ReadNM(ws, n, m) Then Exit Sub
    If Not ffsLoaded Or ffsN <> n Or ffsM <> m Then
        MsgBox "Presione 'Cargar datos' nuevamente.", vbExclamation: Exit Sub
    End If
    ws.Unprotect
    Dim L As FFSLayout: FFS_GetLayout n, m, L
    FFS_DeleteChartIfExists ws: FFS_ClearOutputArea ws, L

    Dim outLineRow As Long: outLineRow = L.instrGen + 2
    Dim outTitleRow As Long: outTitleRow = outLineRow + 2
    Dim ganttTopRow As Long: ganttTopRow = outTitleRow + 2
    Dim indTopRow As Long: indTopRow = ganttTopRow + 16 + 6

    With ws.Range(ws.Cells(outLineRow, 1), ws.Cells(outLineRow, 40)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(0, 0, 0): .Weight = xlThin
    End With
    With ws.Cells(outTitleRow, DEC_C0)
        .Value = "ZONA DE OUTPUTS": .Font.Bold = True: .HorizontalAlignment = xlLeft: .Locked = True
    End With

    Dim c() As Double, startJob() As Double, assign() As Long, Cmax As Double
    FFS_Schedule c, startJob, assign, Cmax

    Dim Cj() As Double, Flow() As Double, Lj() As Double, Tard() As Double, wT() As Double
    Dim avgFlow As Double, Lmax As Double, avgT As Double, sumWT As Double
    Dim lateCount As Long, pctLate As Double, pctOnTime As Double
    FFS_ComputeJobMetrics c, Cj, Flow, Lj, Tard, wT, avgFlow, Lmax, avgT, sumWT, lateCount, pctLate, pctOnTime

    Dim chObj As ChartObject
    Set chObj = ws.ChartObjects.Add(Left:=ws.Cells(ganttTopRow, 2).Left, top:=ws.Cells(ganttTopRow, 2).top, _
        Width:=1100, Height:=240 + 22 * ffsTotalMach)
    chObj.name = CHART_NAME_TIMELINE
    With chObj.Chart
        .ChartType = xlBarStacked: .HasTitle = True: .ChartTitle.text = "Gantt Flexible Flow Shop": .HasLegend = False
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
    End With

    FFS_DrawIndicatorsStructure ws, indTopRow
    Dim jmFirstCol As Long: jmFirstCol = 8
    Dim jmFirstDataRow As Long
    FFS_DrawJobMetricsTable ws, n, indTopRow, jmFirstCol, jmFirstDataRow

    ws.Cells(indTopRow + 1, 5).Value = Cmax
    ws.Cells(indTopRow + 2, 5).Value = avgFlow
    ws.Cells(indTopRow + 3, 5).Value = Lmax
    ws.Cells(indTopRow + 4, 5).Value = avgT
    ws.Cells(indTopRow + 5, 5).Value = sumWT
    ws.Cells(indTopRow + 6, 5).Value = lateCount
    ws.Cells(indTopRow + 7, 5).Value = pctLate: ws.Cells(indTopRow + 7, 5).NumberFormat = "0%"
    ws.Cells(indTopRow + 8, 5).Value = pctOnTime: ws.Cells(indTopRow + 8, 5).NumberFormat = "0%"

    Dim j As Long, k As Long, machStr As String
    For j = 1 To n
        machStr = ""
        For k = 1 To m
            machStr = machStr & "M" & k & "." & assign(j, k)
            If k < m Then machStr = machStr & ", "
        Next k
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol).Value = ffsJob(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 1).Value = startJob(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 2).Value = Cj(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 3).Value = Flow(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 4).Value = Lj(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 5).Value = Tard(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 6).Value = wT(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 7).Value = machStr
    Next j

    FFS_BuildGantt chObj.Chart, c, assign, Cmax

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en FFS_GenerarOutputs: " & Err.Description, vbExclamation: Resume Salir
End Sub

' =========================================================
' LAYOUT
' =========================================================
Private Sub FFS_GetLayout(ByVal n As Long, ByVal m As Long, ByRef L As FFSLayout)
    L.parTitle = 16: L.parHeader = 17: L.parFirst = 18: L.parLast = L.parFirst + n - 1
    L.wsTitle = L.parLast + 3: L.wsHeader = L.wsTitle + 1: L.wsFirst = L.wsHeader + 1: L.wsLast = L.wsFirst + m - 1
    L.decTitle = L.wsLast + 3: L.decHeader = L.decTitle + 1: L.decFirst = L.decHeader + 1: L.decLast = L.decFirst + n - 1
    L.instrLoad = L.decLast + 3: L.instrGen = L.instrLoad + 2
End Sub

Public Function FFS_ReadNM(ByVal ws As Worksheet, ByRef n As Long, ByRef m As Long) As Boolean
    FFS_ReadNM = False
    If Not IsNumeric(ws.Range(CELL_N).Value) Then Exit Function
    If Not IsNumeric(ws.Range(CELL_M).Value) Then Exit Function
    n = CLng(ws.Range(CELL_N).Value): m = CLng(ws.Range(CELL_M).Value)
    If n <= 0 Or m <= 0 Then Exit Function
    If n > MAX_JOBS Or m > MAX_WS Then Exit Function
    FFS_ReadNM = True
End Function

' =========================================================
' VALIDACIÓN
' =========================================================
Private Function FFS_ValidateInputs(ByVal ws As Worksheet, ByVal n As Long, ByVal m As Long, _
        ByRef L As FFSLayout, ByRef warn As String) As Boolean
    FFS_ValidateInputs = False: warn = ""

    Dim dictJobs As Object: Set dictJobs = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As Long, rr As Long, jb As String, v As Variant
    For i = 1 To n
        rr = L.parFirst + i - 1
        jb = UCase$(Trim$(CStr(ws.Cells(rr, PAR_C0).Value)))
        If Len(jb) = 0 Then warn = "Falta Job en parámetros (fila " & rr & ").": Exit Function
        If dictJobs.Exists(jb) Then warn = "Job repetido en parámetros: '" & jb & "'.": Exit Function
        dictJobs.Add jb, True
        v = ws.Cells(rr, PAR_C0 + 1).Value
        If Not IsNumeric(v) Or CDbl(v) < 0 Then warn = "r inválido para " & jb & ".": Exit Function
        For k = 1 To m
            v = ws.Cells(rr, PAR_C0 + 1 + k).Value
            If Not IsNumeric(v) Or CDbl(v) <= 0 Then warn = "p" & k & " debe ser > 0 para " & jb & ".": Exit Function
            v = ws.Cells(rr, PAR_C0 + 1 + m + k).Value
            If Not IsNumeric(v) Or CDbl(v) < 0 Then warn = "s" & k & " inválido para " & jb & ".": Exit Function
        Next k
        v = ws.Cells(rr, PAR_C0 + 2 + 2 * m).Value
        If Not IsNumeric(v) Or CDbl(v) <= 0 Then warn = "d debe ser > 0 para " & jb & ".": Exit Function
        v = ws.Cells(rr, PAR_C0 + 3 + 2 * m).Value
        If Not IsNumeric(v) Or CDbl(v) <= 0 Then warn = "w debe ser > 0 para " & jb & ".": Exit Function
    Next i

    For k = 1 To m
        rr = L.wsFirst + k - 1
        v = ws.Cells(rr, WS_COL_C).Value
        If Not IsNumeric(v) Then warn = "c debe ser numérico para M" & k & ".": Exit Function
        If CLng(v) <> CDbl(v) Or CLng(v) <= 0 Then warn = "c debe ser entero ≥ 1 para M" & k & ".": Exit Function
        If CLng(v) > MAX_MACHS_PER_WS Then warn = "c excede el máximo (" & MAX_MACHS_PER_WS & ") en M" & k & ".": Exit Function
        v = ws.Cells(rr, WS_COL_RMAQ).Value
        If Not IsNumeric(v) Or CDbl(v) < 0 Then warn = "rmaq inválido para M" & k & ".": Exit Function
    Next k

    Dim dictSeq As Object: Set dictSeq = CreateObject("Scripting.Dictionary")
    Dim dictDecJobs As Object: Set dictDecJobs = CreateObject("Scripting.Dictionary")
    Dim vSeq As Variant, vJob As String
    For i = 1 To n
        rr = L.decFirst + i - 1
        vSeq = ws.Cells(rr, DEC_C0).Value
        vJob = UCase$(Trim$(CStr(ws.Cells(rr, DEC_C0 + 1).Value)))
        If Not IsNumeric(vSeq) Then warn = "Secuencia debe ser numérica (fila " & rr & ").": Exit Function
        If CLng(vSeq) <> CDbl(vSeq) Or CLng(vSeq) <= 0 Then warn = "Secuencia debe ser entero positivo (fila " & rr & ").": Exit Function
        If dictSeq.Exists(CStr(CLng(vSeq))) Then warn = "Secuencia repetida: " & CLng(vSeq) & ".": Exit Function
        dictSeq.Add CStr(CLng(vSeq)), True
        If Len(vJob) = 0 Then warn = "Falta Job en decisión (fila " & rr & ").": Exit Function
        If Not dictJobs.Exists(vJob) Then warn = "Job '" & vJob & "' no existe en parámetros.": Exit Function
        If dictDecJobs.Exists(vJob) Then warn = "Job repetido en decisión: '" & vJob & "'.": Exit Function
        dictDecJobs.Add vJob, True
    Next i
    Dim need As Long
    For need = 1 To n
        If Not dictSeq.Exists(CStr(need)) Then warn = "Secuencia debe ser 1.." & n & " sin saltos. Falta: " & need & ".": Exit Function
    Next need
    FFS_ValidateInputs = True
End Function

' =========================================================
' CACHE
' =========================================================
Private Sub FFS_LoadInputsToCache(ByVal ws As Worksheet, ByVal n As Long, ByVal m As Long, ByRef L As FFSLayout)
    ffsLoaded = False: ffsN = n: ffsM = m
    ReDim ffsSeq(1 To n): ReDim ffsJob(1 To n)
    ReDim ffsR(1 To n): ReDim ffsD(1 To n): ReDim ffsW(1 To n)
    ReDim ffsP(1 To n, 1 To m): ReDim ffsS(1 To n, 1 To m)
    ReDim ffsCk(1 To m): ReDim ffsRmaq(1 To m)

    Dim k As Long, total As Long: total = 0
    For k = 1 To m
        ffsCk(k) = CLng(ws.Cells(L.wsFirst + k - 1, WS_COL_C).Value)
        ffsRmaq(k) = CDbl(ws.Cells(L.wsFirst + k - 1, WS_COL_RMAQ).Value)
        total = total + ffsCk(k)
    Next k
    ffsTotalMach = total

    Dim dictR As Object, dictP As Object, dictS As Object, dictD As Object, dictW As Object
    Set dictR = CreateObject("Scripting.Dictionary"): Set dictP = CreateObject("Scripting.Dictionary")
    Set dictS = CreateObject("Scripting.Dictionary"): Set dictD = CreateObject("Scripting.Dictionary")
    Set dictW = CreateObject("Scripting.Dictionary")

    Dim i As Long, rr As Long, jb As String, arrP() As Double, arrS() As Double
    For i = 1 To n
        rr = L.parFirst + i - 1
        jb = UCase$(Trim$(CStr(ws.Cells(rr, PAR_C0).Value)))
        dictR(jb) = CDbl(ws.Cells(rr, PAR_C0 + 1).Value)
        dictD(jb) = CDbl(ws.Cells(rr, PAR_C0 + 2 + 2 * m).Value)
        dictW(jb) = CDbl(ws.Cells(rr, PAR_C0 + 3 + 2 * m).Value)
        ReDim arrP(1 To m): ReDim arrS(1 To m)
        For k = 1 To m
            arrP(k) = CDbl(ws.Cells(rr, PAR_C0 + 1 + k).Value)
            arrS(k) = CDbl(ws.Cells(rr, PAR_C0 + 1 + m + k).Value)
        Next k
        dictP(jb) = arrP: dictS(jb) = arrS
    Next i

    For i = 1 To n
        rr = L.decFirst + i - 1
        jb = UCase$(Trim$(CStr(ws.Cells(rr, DEC_C0 + 1).Value)))
        ffsSeq(i) = CLng(ws.Cells(rr, DEC_C0).Value)
        ffsJob(i) = jb
        ffsR(i) = dictR(jb): ffsD(i) = dictD(jb): ffsW(i) = dictW(jb)
        arrP = dictP(jb): arrS = dictS(jb)
        For k = 1 To m
            ffsP(i, k) = arrP(k): ffsS(i, k) = arrS(k)
        Next k
    Next i
    FFS_SortBySequence n, m
    ffsLoaded = True
End Sub

Private Sub FFS_SortBySequence(ByVal n As Long, ByVal m As Long)
    Dim i As Long, j As Long, t As Long, ts As String, td As Double, k As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If ffsSeq(j) < ffsSeq(i) Then
                t = ffsSeq(i): ffsSeq(i) = ffsSeq(j): ffsSeq(j) = t
                ts = ffsJob(i): ffsJob(i) = ffsJob(j): ffsJob(j) = ts
                td = ffsR(i): ffsR(i) = ffsR(j): ffsR(j) = td
                td = ffsD(i): ffsD(i) = ffsD(j): ffsD(j) = td
                td = ffsW(i): ffsW(i) = ffsW(j): ffsW(j) = td
                For k = 1 To m
                    td = ffsP(i, k): ffsP(i, k) = ffsP(j, k): ffsP(j, k) = td
                    td = ffsS(i, k): ffsS(i, k) = ffsS(j, k): ffsS(j, k) = td
                Next k
            End If
        Next j
    Next i
End Sub

' =========================================================
' SCHEDULING (permutation FFS + FIFO assignment within workstation)
' Para cada workstation k = 1..m, en orden de la permutación:
'   asigna el job a la máquina de k que se libera primero (argmin, desempate por índice)
' =========================================================
Private Sub FFS_Schedule(ByRef c() As Double, ByRef startJob() As Double, ByRef assign() As Long, ByRef Cmax As Double)
    Dim n As Long: n = ffsN: Dim m As Long: m = ffsM
    ReDim c(1 To n, 1 To m): ReDim startJob(1 To n): ReDim assign(1 To n, 1 To m)
    Cmax = 0#

    Dim machFree() As Double: ReDim machFree(1 To m, 1 To MAX_MACHS_PER_WS)
    Dim k As Long, i As Long
    For k = 1 To m
        For i = 1 To ffsCk(k): machFree(k, i) = ffsRmaq(k): Next i
    Next k

    Dim j As Long, arr As Double, iStar As Long, freeMin As Double, startBeforeSetup As Double
    For k = 1 To m
        For j = 1 To n
            If k = 1 Then arr = ffsR(j) Else arr = c(j, k - 1)
            iStar = 1: freeMin = machFree(k, 1)
            For i = 2 To ffsCk(k)
                If machFree(k, i) < freeMin Then iStar = i: freeMin = machFree(k, i)
            Next i
            startBeforeSetup = freeMin
            If arr > startBeforeSetup Then startBeforeSetup = arr
            c(j, k) = startBeforeSetup + ffsS(j, k) + ffsP(j, k)
            machFree(k, iStar) = c(j, k)
            assign(j, k) = iStar
            If k = 1 Then startJob(j) = startBeforeSetup
            If c(j, k) > Cmax Then Cmax = c(j, k)
        Next j
    Next k
End Sub

Private Sub FFS_ComputeJobMetrics(ByRef c() As Double, ByRef Cj() As Double, ByRef Flow() As Double, _
        ByRef Lj() As Double, ByRef Tard() As Double, ByRef wT() As Double, _
        ByRef avgFlow As Double, ByRef Lmax As Double, ByRef avgT As Double, ByRef sumWT As Double, _
        ByRef lateCount As Long, ByRef pctLate As Double, ByRef pctOnTime As Double)
    Dim n As Long: n = ffsN: Dim m As Long: m = ffsM
    ReDim Cj(1 To n): ReDim Flow(1 To n): ReDim Lj(1 To n): ReDim Tard(1 To n): ReDim wT(1 To n)
    avgFlow = 0#: avgT = 0#: sumWT = 0#: lateCount = 0: Lmax = -1E+30
    Dim j As Long
    For j = 1 To n
        Cj(j) = c(j, m)
        Flow(j) = Cj(j) - ffsR(j)
        Lj(j) = Cj(j) - ffsD(j)
        If Lj(j) > 0# Then Tard(j) = Lj(j) Else Tard(j) = 0#
        wT(j) = ffsW(j) * Tard(j)
        avgFlow = avgFlow + Flow(j): avgT = avgT + Tard(j): sumWT = sumWT + wT(j)
        If Lj(j) > Lmax Then Lmax = Lj(j)
        If Tard(j) > 0.000001 Then lateCount = lateCount + 1
    Next j
    avgFlow = avgFlow / CDbl(n): avgT = avgT / CDbl(n)
    pctLate = lateCount / CDbl(n): pctOnTime = 1# - pctLate
End Sub

' =========================================================
' GANTT (categorías = M1.1, M1.2, M2.1, ...)
' =========================================================
Private Sub FFS_BuildGantt(ByVal ch As Chart, ByRef c() As Double, ByRef assign() As Long, ByVal Cmax As Double)
    Dim n As Long: n = ffsN: Dim m As Long: m = ffsM
    Dim total As Long: total = ffsTotalMach
    Dim cats() As Variant: ReDim cats(1 To total)
    Dim catIdx() As Long: ReDim catIdx(1 To m, 1 To MAX_MACHS_PER_WS)
    Dim k As Long, i As Long, p As Long: p = 0
    For k = 1 To m
        For i = 1 To ffsCk(k)
            p = p + 1: cats(p) = "M" & k & "." & i: catIdx(k, i) = p
        Next i
    Next k

    With ch
        .ChartType = xlBarStacked: .HasLegend = False
        On Error Resume Next: Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop: On Error GoTo 0
        Dim base() As Double: ReDim base(1 To total)
        Dim srs As Series
        Set srs = .SeriesCollection.NewSeries: srs.Values = base: srs.XValues = cats
        srs.Format.Fill.Visible = msoFalse: srs.Format.Line.Visible = msoFalse

        Dim machOff() As Double: ReDim machOff(1 To total)
        Dim jobColor As Object: Set jobColor = CreateObject("Scripting.Dictionary")
        Dim baseCol As Long, gap As Double, setupDur As Double, procDur As Double, startBeforeSetup As Double
        Dim mIdx As Long, j As Long
        For k = 1 To m
            For j = 1 To n
                mIdx = catIdx(k, assign(j, k))
                setupDur = ffsS(j, k): procDur = ffsP(j, k)
                startBeforeSetup = c(j, k) - setupDur - procDur
                gap = startBeforeSetup - machOff(mIdx)
                If gap > 0.000001 Then
                    Set srs = .SeriesCollection.NewSeries
                    srs.Values = FFS_OneHot(total, mIdx, gap): srs.XValues = cats
                    srs.Format.Fill.Visible = msoFalse: srs.Format.Line.Visible = msoFalse
                    machOff(mIdx) = machOff(mIdx) + gap
                End If
                baseCol = FFS_ColorForJobName(ffsJob(j), jobColor)
                If setupDur > 0.000001 Then
                    Set srs = .SeriesCollection.NewSeries
                    srs.Values = FFS_OneHot(total, mIdx, setupDur): srs.XValues = cats
                    srs.Format.Fill.ForeColor.RGB = FFS_LightTone(baseCol): srs.Format.Line.Visible = msoFalse
                    machOff(mIdx) = machOff(mIdx) + setupDur
                End If
                If procDur > 0.000001 Then
                    Set srs = .SeriesCollection.NewSeries
                    srs.Values = FFS_OneHot(total, mIdx, procDur): srs.XValues = cats
                    srs.Format.Fill.ForeColor.RGB = FFS_DarkTone(baseCol): srs.Format.Line.Visible = msoFalse
                    With srs.Points(mIdx)
                        .HasDataLabel = True: .DataLabel.text = ffsJob(j): .DataLabel.Font.Size = 9
                    End With
                    machOff(mIdx) = machOff(mIdx) + procDur
                End If
            Next j
        Next k

        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.text = "Tiempo"
        FFS_ConfigurarEjeTiempo ch, Cmax
    End With
End Sub

' =========================================================
' OUTPUTS (estructura)
' =========================================================
Private Sub FFS_DrawIndicatorsStructure(ByVal ws As Worksheet, ByVal indTopRow As Long)
    With ws.Cells(indTopRow, 2)
        .Value = "Indicadores": .Font.Bold = True: .HorizontalAlignment = xlLeft: .Locked = True
    End With
    Dim labels As Variant: labels = Array("Makespan (Cmax)", "Tiempo de flujo promedio", "Máximo lateness (Lmax)", "Tardanza media", "Tardanza ponderada total", "# Jobs tarde", "% Jobs tarde", "% Jobs a tiempo")
    Dim i As Long
    For i = 0 To UBound(labels)
        With ws.Range(ws.Cells(indTopRow + 1 + i, 2), ws.Cells(indTopRow + 1 + i, 4))
            On Error Resume Next: .UnMerge: On Error GoTo 0
            .Merge: .Value = labels(i): .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
        End With
        With ws.Cells(indTopRow + 1 + i, 5)
            .ClearContents: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Locked = True
        End With
    Next i
    With ws.Range(ws.Cells(indTopRow + 1, 2), ws.Cells(indTopRow + 8, 5))
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin
    End With
End Sub

Private Sub FFS_DrawJobMetricsTable(ByVal ws As Worksheet, ByVal n As Long, _
        ByVal indTopRow As Long, ByVal firstCol As Long, ByRef firstDataRow As Long)
    Dim headerRow As Long: headerRow = indTopRow + 1: firstDataRow = headerRow + 1
    Dim lastRow As Long: lastRow = firstDataRow + n - 1
    With ws.Cells(indTopRow, firstCol)
        .Value = "Indicadores por job": .Font.Bold = True: .HorizontalAlignment = xlLeft: .Locked = True
    End With
    Dim headers As Variant: headers = Array("Job", "Inicio", "Cj", "Flow (Cj-rj)", "L (Cj-dj)", "T=max(L,0)", "w*T", "Máquinas")
    Dim h As Long
    For h = 0 To UBound(headers)
        With ws.Cells(headerRow, firstCol + h)
            .Value = headers(h): .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Interior.Color = RGB(230, 230, 230): .Locked = True
        End With
    Next h
    With ws.Range(ws.Cells(firstDataRow, firstCol), ws.Cells(lastRow, firstCol + UBound(headers)))
        .ClearContents: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Locked = True
    End With
    With ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(lastRow, firstCol + UBound(headers)))
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin
    End With
End Sub

' =========================================================
' HELPERS DE DIBUJO / LIMPIEZA
' =========================================================
Private Sub FFS_SectionTitle(ByVal ws As Worksheet, ByVal rowN As Long, ByVal col As Long, ByVal txt As String)
    With ws.Cells(rowN, col)
        .Value = txt: .Font.Bold = True: .Font.Italic = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With
End Sub

Private Sub FFS_FormatHeaderRow(ByVal ws As Worksheet, ByVal rowN As Long, ByVal col1 As Long, ByVal col2 As Long)
    With ws.Range(ws.Cells(rowN, col1), ws.Cells(rowN, col2))
        .Font.Bold = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin
        .Locked = True
    End With
End Sub

Private Sub FFS_FormatEditableBlock(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal col1 As Long, ByVal col2 As Long, ByVal bgColor As Long)
    With ws.Range(ws.Cells(firstRow, col1), ws.Cells(lastRow, col2))
        .ClearContents
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        If bgColor = 0 Then .Interior.Pattern = xlNone Else .Interior.Color = bgColor
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(160, 160, 160): .Borders.Weight = xlThin
        .Locked = False
    End With
End Sub

Private Sub FFS_ClearDynamicZone(ByVal ws As Worksheet)
    With ws.Range(ws.Cells(DYNAMIC_TOP_ROW, 1), ws.Cells(DYNAMIC_TOP_ROW + INPUT_CLEAR_ROWS, 50))
        On Error Resume Next: .UnMerge: On Error GoTo 0
        .ClearContents: .Borders.LineStyle = xlNone: .Interior.Pattern = xlNone
        .Font.Bold = False: .Font.Italic = False: .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral: .VerticalAlignment = xlCenter: .Validation.Delete
    End With
End Sub

Private Sub FFS_ClearOutputArea(ByVal ws As Worksheet, ByRef L As FFSLayout)
    Dim outStart As Long: outStart = L.instrGen + 2
    With ws.Range(ws.Cells(outStart, 1), ws.Cells(outStart + 900, 40))
        .ClearContents: .Borders.LineStyle = xlNone: .Interior.Pattern = xlNone
        .Font.Bold = False: .Font.Italic = False: .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral: .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub FFS_DeleteChartIfExists(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.ChartObjects.Count To 1 Step -1: ws.ChartObjects(i).Delete: Next i
    On Error GoTo 0
End Sub

Private Sub FFS_SetJobListNamedRange(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error Resume Next: ws.Parent.Names("FFS_JobList").Delete: On Error GoTo 0
    ws.Parent.Names.Add name:="FFS_JobList", _
        RefersTo:=ws.Range(ws.Cells(firstRow, PAR_C0), ws.Cells(lastRow, PAR_C0))
End Sub

' =========================================================
' COLORES / EJES
' =========================================================
Private Function FFS_OneHot(ByVal total As Long, ByVal idx As Long, ByVal v As Double) As Variant
    Dim a() As Double, t As Long: ReDim a(1 To total)
    For t = 1 To total: a(t) = 0#: Next t
    a(idx) = v: FFS_OneHot = a
End Function

Private Function FFS_ColorForJobName(ByVal jn As String, ByVal jc As Object) As Long
    Dim k As String: k = UCase$(Trim$(jn))
    If Not jc.Exists(k) Then jc.Add k, jc.Count + 1
    FFS_ColorForJobName = FFS_BasePalette(jc(k))
End Function

Private Function FFS_BasePalette(ByVal i As Long) As Long
    Dim p As Variant
    p = Array(RGB(52, 96, 174), RGB(46, 204, 113), RGB(155, 89, 182), RGB(241, 196, 15), _
              RGB(231, 76, 60), RGB(26, 188, 156), RGB(127, 140, 141), RGB(52, 152, 219), _
              RGB(39, 174, 96), RGB(243, 156, 18))
    FFS_BasePalette = p((i - 1) Mod (UBound(p) + 1))
End Function

Private Function FFS_LightTone(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = (c And &HFF): g = (c \ &H100) And &HFF: b = (c \ &H10000) And &HFF
    FFS_LightTone = RGB(IIf(r + 80 < 255, r + 80, 255), IIf(g + 80 < 255, g + 80, 255), IIf(b + 80 < 255, b + 80, 255))
End Function

Private Function FFS_DarkTone(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = (c And &HFF): g = (c \ &H100) And &HFF: b = (c \ &H10000) And &HFF
    FFS_DarkTone = RGB(IIf(r > 40, r - 40, 0), IIf(g > 40, g - 40, 0), IIf(b > 40, b - 40, 0))
End Function

Private Sub FFS_ConfigurarEjeTiempo(ByVal ch As Chart, ByVal maxTime As Double)
    Dim eje As Axis: Set eje = ch.Axes(xlValue)
    If maxTime < 0.000001 Then maxTime = 1
    Dim majorU As Double: majorU = FFS_NiceMajorUnit(maxTime, 30)
    Dim minorU As Double: minorU = majorU / 5#
    If minorU < 1# And majorU >= 5# Then minorU = 1#
    eje.MinimumScale = 0
    eje.MaximumScale = Application.WorksheetFunction.Ceiling(maxTime, majorU)
    eje.MajorUnit = majorU: eje.MinorUnit = minorU
    eje.TickLabelPosition = xlTickLabelPositionNextToAxis
    On Error Resume Next: eje.TickLabels.Orientation = xlHorizontal: On Error GoTo 0
    eje.HasMajorGridlines = True: eje.HasMinorGridlines = True
    On Error Resume Next
    With eje.MajorGridlines.Format.Line
        .ForeColor.RGB = RGB(180, 180, 180): .Weight = 0.5
    End With
    With eje.MinorGridlines.Format.Line
        .ForeColor.RGB = RGB(225, 225, 225): .DashStyle = msoLineDash: .Weight = 0.25
    End With
    On Error GoTo 0
End Sub

Private Function FFS_NiceMajorUnit(ByVal mx As Double, ByVal tt As Long) As Double
    If mx <= 0 Then FFS_NiceMajorUnit = 1: Exit Function
    Dim rs As Double: rs = mx / tt: If rs < 1 Then rs = 1
    Dim p10 As Double: p10 = 10 ^ Int(Log(rs) / Log(10))
    Dim fr As Double: fr = rs / p10
    Select Case fr
        Case Is <= 1: FFS_NiceMajorUnit = 1 * p10
        Case Is <= 2: FFS_NiceMajorUnit = 2 * p10
        Case Is <= 5: FFS_NiceMajorUnit = 5 * p10
        Case Else: FFS_NiceMajorUnit = 10 * p10
    End Select
End Function
