Option Explicit

' =========================================================
' FLOW SHOP CON HORARIO (BETA)
' Permite definir turnos de trabajo (ej. 08:00-12:00, 14:00-16:00).
' Tiempo t = 0 corresponde a la hora de inicio del primer turno del día 1.
' Cada operación debe caber completa dentro de un único turno; si no, se
' programa en el siguiente turno disponible.
' =========================================================

Private Const SHEET_NAME As String = "FlowShopHorario"

Private Const CELL_N As String = "C12"          ' # jobs
Private Const CELL_M As String = "C13"          ' # máquinas
Private Const CELL_S As String = "C14"          ' # turnos por día
Private Const CELL_FSH_READY As String = "Z2"

Private Const DYNAMIC_TOP_ROW As Long = 15
Private Const INPUT_CLEAR_ROWS As Long = 1500

Private Const PAR_C0 As Long = 2
Private Const RM_COL_MACH As Long = 2
Private Const RM_COL_RMAQ As Long = 3
Private Const SH_COL_TURNO As Long = 2
Private Const SH_COL_INI As Long = 3
Private Const SH_COL_FIN As Long = 4
Private Const DEC_C0 As Long = 2

Private Const MAX_SHIFTS As Long = 6
Private Const MAX_HORIZON_DAYS As Long = 60
Private Const CHART_NAME_TIMELINE As String = "chTimeline_FSH"

Private Type FSHLayout
    parTitle As Long: parHeader As Long: parFirst As Long: parLast As Long
    rmaqTitle As Long: rmaqHeader As Long: rmaqFirst As Long: rmaqLast As Long
    shTitle As Long: shHeader As Long: shFirst As Long: shLast As Long
    decTitle As Long: decHeader As Long: decFirst As Long: decLast As Long
    instr4Row As Long: instr5Row As Long
End Type

' CACHE
Private fshLoaded As Boolean
Private fshN As Long, fshM As Long, fshNumShifts As Long
Private fshSeq() As Long
Private fshJob() As String
Private fshR() As Double, fshP() As Double, fshS() As Double, fshD() As Double, fshW() As Double
Private fshRmaq() As Double
Private fshShiftIni() As Double, fshShiftFin() As Double  ' en horas decimales 24h del día (ej 8.0)
Private fshStartOfDay As Double                          ' = primer ShiftIni; t=0 corresponde a esta hora del día 1

' =========================================================
Public Sub RedibujarInputs_FSH(ByVal ws As Worksheet)
    On Error GoTo ErrH
    Dim n As Long, m As Long, ns As Long
    If Not FSH_ReadNMS(ws, n, m, ns) Then Exit Sub

    ws.Unprotect
    fshLoaded = False: fshN = 0: fshM = 0: fshNumShifts = 0
    ws.Range(CELL_FSH_READY).Value = ""
    FSH_ClearDynamicZone ws: FSH_DeleteChartIfExists ws

    Dim L As FSHLayout: FSH_GetLayout n, m, ns, L

    ' Tabla 1: Parámetros
    With ws.Cells(L.parTitle, PAR_C0)
        .Value = "2. Parámetros por job [PARÁMETRO] — r, p1..pm, s1..sm, d, w (todo en HORAS)."
        .Font.Bold = True: .Font.Italic = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With
    ws.Cells(L.parHeader, PAR_C0).Value = "Job"
    ws.Cells(L.parHeader, PAR_C0 + 1).Value = "r"
    Dim k As Long
    For k = 1 To m: ws.Cells(L.parHeader, PAR_C0 + 1 + k).Value = "p" & k: Next k
    For k = 1 To m: ws.Cells(L.parHeader, PAR_C0 + 1 + m + k).Value = "s" & k: Next k
    ws.Cells(L.parHeader, PAR_C0 + 2 + 2 * m).Value = "d"
    ws.Cells(L.parHeader, PAR_C0 + 3 + 2 * m).Value = "w"
    Dim lastColPar As Long: lastColPar = PAR_C0 + 3 + 2 * m
    FSH_FormatHeaderRow ws, L.parHeader, PAR_C0, lastColPar
    FSH_FormatEditableBlock ws, L.parFirst, L.parLast, PAR_C0, lastColPar, RGB(248, 248, 248)
    Dim i As Long
    For i = 1 To n
        ws.Cells(L.parFirst + i - 1, PAR_C0).Value = "J" & i
        ws.Cells(L.parFirst + i - 1, PAR_C0 + 3 + 2 * m).Value = 1
    Next i
    ws.Range(ws.Cells(L.parFirst, PAR_C0), ws.Cells(L.parLast, PAR_C0)).Locked = True
    FSH_SetJobListNamedRange ws, L.parFirst, L.parLast

    ' Tabla 2: rmaq
    With ws.Cells(L.rmaqTitle, RM_COL_MACH)
        .Value = "3. Disponibilidad de cada máquina rmaq [PARÁMETRO] (en horas desde t=0)."
        .Font.Bold = True: .Font.Italic = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With
    ws.Cells(L.rmaqHeader, RM_COL_MACH).Value = "Máquina"
    ws.Cells(L.rmaqHeader, RM_COL_RMAQ).Value = "rmaq"
    FSH_FormatHeaderRow ws, L.rmaqHeader, RM_COL_MACH, RM_COL_RMAQ
    FSH_FormatEditableBlock ws, L.rmaqFirst, L.rmaqLast, RM_COL_MACH, RM_COL_RMAQ, RGB(248, 248, 248)
    For k = 1 To m: ws.Cells(L.rmaqFirst + k - 1, RM_COL_MACH).Value = "M" & k: Next k
    ws.Range(ws.Cells(L.rmaqFirst, RM_COL_MACH), ws.Cells(L.rmaqLast, RM_COL_MACH)).Locked = True

    ' Tabla 3: Turnos
    With ws.Cells(L.shTitle, SH_COL_TURNO)
        .Value = "4. Turnos [PARÁMETRO] — hora de inicio y fin en formato 24h (ej. 8 y 12). t=0 = inicio del primer turno del día 1."
        .Font.Bold = True: .Font.Italic = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With
    ws.Cells(L.shHeader, SH_COL_TURNO).Value = "Turno"
    ws.Cells(L.shHeader, SH_COL_INI).Value = "Hora inicio"
    ws.Cells(L.shHeader, SH_COL_FIN).Value = "Hora fin"
    FSH_FormatHeaderRow ws, L.shHeader, SH_COL_TURNO, SH_COL_FIN
    FSH_FormatEditableBlock ws, L.shFirst, L.shLast, SH_COL_TURNO, SH_COL_FIN, RGB(248, 248, 248)
    Dim sDefault As Variant: sDefault = Array(Array(8, 12), Array(14, 16), Array(0, 0), Array(0, 0), Array(0, 0), Array(0, 0))
    For k = 1 To ns
        ws.Cells(L.shFirst + k - 1, SH_COL_TURNO).Value = "T" & k
        If k <= 2 Then
            ws.Cells(L.shFirst + k - 1, SH_COL_INI).Value = sDefault(k - 1)(0)
            ws.Cells(L.shFirst + k - 1, SH_COL_FIN).Value = sDefault(k - 1)(1)
        End If
    Next k
    ws.Range(ws.Cells(L.shFirst, SH_COL_TURNO), ws.Cells(L.shLast, SH_COL_TURNO)).Locked = True

    ' Tabla 4: Decisión (permutación)
    With ws.Cells(L.decTitle, DEC_C0)
        .Value = "5. Decisión [DECISIÓN] — escribe la secuencia y selecciona el job."
        .Font.Bold = True: .Font.Italic = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With
    ws.Cells(L.decHeader, DEC_C0).Value = "Secuencia"
    ws.Cells(L.decHeader, DEC_C0 + 1).Value = "Job"
    FSH_FormatHeaderRow ws, L.decHeader, DEC_C0, DEC_C0 + 1
    FSH_FormatEditableBlock ws, L.decFirst, L.decLast, DEC_C0, DEC_C0 + 1, 0
    Dim rngDV As Range
    Set rngDV = ws.Range(ws.Cells(L.decFirst, DEC_C0 + 1), ws.Cells(L.decLast, DEC_C0 + 1))
    On Error Resume Next: rngDV.Validation.Delete: On Error GoTo 0
    rngDV.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=FSH_JobList"
    rngDV.Validation.IgnoreBlank = True: rngDV.Validation.InCellDropdown = True
    For i = 1 To n: ws.Cells(L.decFirst + i - 1, DEC_C0).Value = i: Next i

    With ws.Cells(L.instr4Row, DEC_C0)
        .Value = "6. Presione 'Cargar datos' para validar la información."
        .Font.Bold = True: .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False: ws.Range(CELL_S).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en RedibujarInputs_FSH: " & Err.Description, vbExclamation: Resume Salir
End Sub

Public Sub FSH_CargarDatos()
    On Error GoTo ErrH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Dim n As Long, m As Long, ns As Long
    If Not FSH_ReadNMS(ws, n, m, ns) Then
        MsgBox "Verifica que # jobs, # máquinas y # turnos sean enteros positivos.", vbExclamation: Exit Sub
    End If
    ws.Unprotect
    Dim L As FSHLayout: FSH_GetLayout n, m, ns, L
    FSH_ClearOutputArea ws, L: FSH_DeleteChartIfExists ws
    ws.Range(CELL_FSH_READY).Value = ""

    Dim warn As String
    If Not FSH_ValidateInputs(ws, n, m, ns, L, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos - Flow Shop Horario": GoTo Salir
    End If
    FSH_LoadInputsToCache ws, n, m, ns, L

    ws.Range(CELL_FSH_READY).Value = "OK"
    With ws.Cells(L.instr5Row, DEC_C0)
        .Value = "7. Datos válidos. Presione 'Generar outputs'."
        .Font.Bold = True: .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False: ws.Range(CELL_S).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en FSH_CargarDatos: " & Err.Description, vbExclamation: Resume Salir
End Sub

Public Sub FSH_GenerarOutputs()
    On Error GoTo ErrH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If UCase$(Trim$(CStr(ws.Range(CELL_FSH_READY).Value))) <> "OK" Then
        MsgBox "Primero presione 'Cargar datos'.", vbExclamation: Exit Sub
    End If
    Dim n As Long, m As Long, ns As Long
    If Not FSH_ReadNMS(ws, n, m, ns) Then Exit Sub
    If Not fshLoaded Or fshN <> n Or fshM <> m Or fshNumShifts <> ns Then
        MsgBox "Presione 'Cargar datos' nuevamente.", vbExclamation: Exit Sub
    End If

    ws.Unprotect
    Dim L As FSHLayout: FSH_GetLayout n, m, ns, L
    FSH_DeleteChartIfExists ws: FSH_ClearOutputArea ws, L

    Dim outLineRow As Long: outLineRow = L.instr5Row + 1
    Dim outTitleRow As Long: outTitleRow = outLineRow + 2
    Dim ganttTopRow As Long: ganttTopRow = outTitleRow + 2
    Dim indTopRow As Long: indTopRow = ganttTopRow + 16 + 6

    With ws.Range(ws.Cells(outLineRow, 1), ws.Cells(outLineRow, 40)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(0, 0, 0): .Weight = xlThin
    End With
    With ws.Cells(outTitleRow, DEC_C0)
        .Value = "ZONA DE OUTPUTS  (t=0 = " & FSH_HoraFmt(0) & ")"
        .Font.Bold = True: .HorizontalAlignment = xlLeft: .Locked = True
    End With

    Dim c() As Double, st() As Double, Cmax As Double, warn As String
    If Not FSH_Schedule(c, st, Cmax, warn) Then
        MsgBox warn, vbExclamation, "Revisar programación": GoTo Salir
    End If

    Dim Cj() As Double, Flow() As Double, Lj() As Double, Tard() As Double, wT() As Double
    Dim avgFlow As Double, Lmax As Double, avgT As Double, sumWT As Double
    Dim lateCount As Long, pctLate As Double, pctOnTime As Double
    FSH_ComputeJobMetrics c, Cj, Flow, Lj, Tard, wT, avgFlow, Lmax, avgT, sumWT, lateCount, pctLate, pctOnTime

    Dim chObj As ChartObject
    Set chObj = ws.ChartObjects.Add(Left:=ws.Cells(ganttTopRow, 2).Left, top:=ws.Cells(ganttTopRow, 2).top, _
        Width:=1100, Height:=220 + 24 * m)
    chObj.name = CHART_NAME_TIMELINE
    With chObj.Chart
        .ChartType = xlBarStacked: .HasTitle = True: .ChartTitle.text = "Gantt Flow Shop con horario": .HasLegend = False
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
    End With

    FSH_DrawIndicatorsStructure ws, indTopRow
    Dim jmFirstCol As Long: jmFirstCol = 8
    Dim jmFirstDataRow As Long
    FSH_DrawJobMetricsTable ws, n, m, indTopRow, jmFirstCol, jmFirstDataRow

    ws.Cells(indTopRow + 1, 5).Value = Cmax
    ws.Cells(indTopRow + 2, 5).Value = avgFlow
    ws.Cells(indTopRow + 3, 5).Value = Lmax
    ws.Cells(indTopRow + 4, 5).Value = avgT
    ws.Cells(indTopRow + 5, 5).Value = sumWT
    ws.Cells(indTopRow + 6, 5).Value = lateCount
    ws.Cells(indTopRow + 7, 5).Value = pctLate: ws.Cells(indTopRow + 7, 5).NumberFormat = "0%"
    ws.Cells(indTopRow + 8, 5).Value = pctOnTime: ws.Cells(indTopRow + 8, 5).NumberFormat = "0%"

    Dim j As Long
    For j = 1 To n
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol).Value = fshJob(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 1).Value = st(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 2).Value = FSH_HoraFmt(st(j))
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 3).Value = Cj(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 4).Value = FSH_HoraFmt(Cj(j))
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 5).Value = Flow(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 6).Value = Lj(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 7).Value = Tard(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 8).Value = wT(j)
    Next j

    FSH_BuildGantt chObj.Chart, n, m, c, Cmax

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False: ws.Range(CELL_S).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en FSH_GenerarOutputs: " & Err.Description, vbExclamation: Resume Salir
End Sub

' =========================================================
' LAYOUT
' =========================================================
Private Sub FSH_GetLayout(ByVal n As Long, ByVal m As Long, ByVal ns As Long, ByRef L As FSHLayout)
    L.parTitle = 16: L.parHeader = 17: L.parFirst = 18: L.parLast = L.parFirst + n - 1
    L.rmaqTitle = L.parLast + 3: L.rmaqHeader = L.rmaqTitle + 1: L.rmaqFirst = L.rmaqHeader + 1: L.rmaqLast = L.rmaqFirst + m - 1
    L.shTitle = L.rmaqLast + 3: L.shHeader = L.shTitle + 1: L.shFirst = L.shHeader + 1: L.shLast = L.shFirst + ns - 1
    L.decTitle = L.shLast + 3: L.decHeader = L.decTitle + 1: L.decFirst = L.decHeader + 1: L.decLast = L.decFirst + n - 1
    L.instr4Row = L.decLast + 3: L.instr5Row = L.instr4Row + 2
End Sub

Public Function FSH_ReadNMS(ByVal ws As Worksheet, ByRef n As Long, ByRef m As Long, ByRef ns As Long) As Boolean
    FSH_ReadNMS = False
    If Not IsNumeric(ws.Range(CELL_N).Value) Then Exit Function
    If Not IsNumeric(ws.Range(CELL_M).Value) Then Exit Function
    If Not IsNumeric(ws.Range(CELL_S).Value) Then Exit Function
    n = CLng(ws.Range(CELL_N).Value): m = CLng(ws.Range(CELL_M).Value): ns = CLng(ws.Range(CELL_S).Value)
    If n <= 0 Or m <= 0 Or ns <= 0 Then Exit Function
    If ns > MAX_SHIFTS Then Exit Function
    FSH_ReadNMS = True
End Function

' =========================================================
' VALIDACIÓN
' =========================================================
Private Function FSH_ValidateInputs(ByVal ws As Worksheet, ByVal n As Long, ByVal m As Long, ByVal ns As Long, _
        ByRef L As FSHLayout, ByRef warn As String) As Boolean
    FSH_ValidateInputs = False: warn = ""

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
        v = ws.Cells(L.rmaqFirst + k - 1, RM_COL_RMAQ).Value
        If Not IsNumeric(v) Or CDbl(v) < 0 Then warn = "rmaq inválido para M" & k & ".": Exit Function
    Next k

    Dim prevFin As Double: prevFin = -1
    For k = 1 To ns
        Dim viIni As Variant, viFin As Variant
        viIni = ws.Cells(L.shFirst + k - 1, SH_COL_INI).Value
        viFin = ws.Cells(L.shFirst + k - 1, SH_COL_FIN).Value
        If Not IsNumeric(viIni) Or CDbl(viIni) < 0 Or CDbl(viIni) >= 24 Then warn = "Turno " & k & ": hora inicio debe estar en [0, 24).": Exit Function
        If Not IsNumeric(viFin) Or CDbl(viFin) <= 0 Or CDbl(viFin) > 24 Then warn = "Turno " & k & ": hora fin debe estar en (0, 24].": Exit Function
        If CDbl(viFin) <= CDbl(viIni) Then warn = "Turno " & k & ": hora fin debe ser > hora inicio.": Exit Function
        If CDbl(viIni) < prevFin Then warn = "Los turnos deben ir en orden creciente y sin superponerse.": Exit Function
        prevFin = CDbl(viFin)
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
    FSH_ValidateInputs = True
End Function

' =========================================================
' CACHE
' =========================================================
Private Sub FSH_LoadInputsToCache(ByVal ws As Worksheet, ByVal n As Long, ByVal m As Long, ByVal ns As Long, ByRef L As FSHLayout)
    fshLoaded = False: fshN = n: fshM = m: fshNumShifts = ns
    ReDim fshSeq(1 To n): ReDim fshJob(1 To n)
    ReDim fshR(1 To n): ReDim fshD(1 To n): ReDim fshW(1 To n)
    ReDim fshP(1 To n, 1 To m): ReDim fshS(1 To n, 1 To m)
    ReDim fshRmaq(1 To m)
    ReDim fshShiftIni(1 To ns): ReDim fshShiftFin(1 To ns)

    Dim k As Long
    For k = 1 To m: fshRmaq(k) = CDbl(ws.Cells(L.rmaqFirst + k - 1, RM_COL_RMAQ).Value): Next k
    For k = 1 To ns
        fshShiftIni(k) = CDbl(ws.Cells(L.shFirst + k - 1, SH_COL_INI).Value)
        fshShiftFin(k) = CDbl(ws.Cells(L.shFirst + k - 1, SH_COL_FIN).Value)
    Next k
    fshStartOfDay = fshShiftIni(1)

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
        fshSeq(i) = CLng(ws.Cells(rr, DEC_C0).Value)
        fshJob(i) = jb
        fshR(i) = dictR(jb): fshD(i) = dictD(jb): fshW(i) = dictW(jb)
        arrP = dictP(jb): arrS = dictS(jb)
        For k = 1 To m
            fshP(i, k) = arrP(k): fshS(i, k) = arrS(k)
        Next k
    Next i
    FSH_SortBySequence n, m
    fshLoaded = True
End Sub

Private Sub FSH_SortBySequence(ByVal n As Long, ByVal m As Long)
    Dim i As Long, j As Long, t As Long, ts As String, td As Double, k As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If fshSeq(j) < fshSeq(i) Then
                t = fshSeq(i): fshSeq(i) = fshSeq(j): fshSeq(j) = t
                ts = fshJob(i): fshJob(i) = fshJob(j): fshJob(j) = ts
                td = fshR(i): fshR(i) = fshR(j): fshR(j) = td
                td = fshD(i): fshD(i) = fshD(j): fshD(j) = td
                td = fshW(i): fshW(i) = fshW(j): fshW(j) = td
                For k = 1 To m
                    td = fshP(i, k): fshP(i, k) = fshP(j, k): fshP(j, k) = td
                    td = fshS(i, k): fshS(i, k) = fshS(j, k): fshS(j, k) = td
                Next k
            End If
        Next j
    Next i
End Sub

' =========================================================
' SCHEDULING con turnos
' c(i, k) = tiempo de finalización del job i en la máquina k.
' st(i) = tiempo de inicio del job i en M1.
' La operación de duración (s+p) debe caber completa dentro de un único turno.
' =========================================================
Private Function FSH_Schedule(ByRef c() As Double, ByRef st() As Double, ByRef Cmax As Double, ByRef warn As String) As Boolean
    Dim n As Long: n = fshN: Dim m As Long: m = fshM
    ReDim c(1 To n, 1 To m): ReDim st(1 To n)
    Cmax = 0#: warn = ""

    Dim lastOnMac() As Double: ReDim lastOnMac(1 To m)
    Dim k As Long: For k = 1 To m: lastOnMac(k) = fshRmaq(k): Next k

    Dim i As Long, earliestStart As Double, duration As Double, finStart As Double
    For i = 1 To n
        For k = 1 To m
            If k = 1 Then
                earliestStart = lastOnMac(k)
                If fshR(i) > earliestStart Then earliestStart = fshR(i)
            Else
                earliestStart = lastOnMac(k)
                If c(i, k - 1) > earliestStart Then earliestStart = c(i, k - 1)
            End If
            duration = fshS(i, k) + fshP(i, k)
            finStart = FSH_NextFeasibleStart(earliestStart, duration)
            If finStart < 0 Then
                warn = "El job '" & fshJob(i) & "' op M" & k & " (duración " & duration & " h) no cabe en ningún turno dentro de " & MAX_HORIZON_DAYS & " días. Aumenta la longitud de algún turno."
                FSH_Schedule = False: Exit Function
            End If
            If k = 1 Then st(i) = finStart
            c(i, k) = finStart + duration
            lastOnMac(k) = c(i, k)
            If c(i, k) > Cmax Then Cmax = c(i, k)
        Next k
    Next i
    FSH_Schedule = True
End Function

' Devuelve el primer instante >= earliestStart tal que [t, t+duration] cabe completo en un turno.
' Tiempo en horas desde t=0 (= primer ShiftIni del día 1).
Private Function FSH_NextFeasibleStart(ByVal earliestStart As Double, ByVal duration As Double) As Double
    Dim d As Long, i As Long
    Dim shiftStart As Double, shiftEnd As Double, effStart As Double
    For d = 0 To MAX_HORIZON_DAYS
        For i = 1 To fshNumShifts
            shiftStart = d * 24# + (fshShiftIni(i) - fshStartOfDay)
            shiftEnd = d * 24# + (fshShiftFin(i) - fshStartOfDay)
            If shiftEnd < earliestStart Then GoTo NextShift
            If shiftEnd - shiftStart + 0.000001 < duration Then GoTo NextShift
            effStart = earliestStart
            If shiftStart > effStart Then effStart = shiftStart
            If effStart + duration <= shiftEnd + 0.000001 Then
                FSH_NextFeasibleStart = effStart
                Exit Function
            End If
NextShift:
        Next i
    Next d
    FSH_NextFeasibleStart = -1
End Function

' =========================================================
' MÉTRICAS
' =========================================================
Private Sub FSH_ComputeJobMetrics(ByRef c() As Double, _
        ByRef Cj() As Double, ByRef Flow() As Double, ByRef Lj() As Double, ByRef Tard() As Double, ByRef wT() As Double, _
        ByRef avgFlow As Double, ByRef Lmax As Double, ByRef avgT As Double, ByRef sumWT As Double, _
        ByRef lateCount As Long, ByRef pctLate As Double, ByRef pctOnTime As Double)
    Dim n As Long: n = fshN: Dim m As Long: m = fshM
    ReDim Cj(1 To n): ReDim Flow(1 To n): ReDim Lj(1 To n): ReDim Tard(1 To n): ReDim wT(1 To n)
    avgFlow = 0#: avgT = 0#: sumWT = 0#: lateCount = 0: Lmax = -1E+30
    Dim i As Long
    For i = 1 To n
        Cj(i) = c(i, m)
        Flow(i) = Cj(i) - fshR(i)
        Lj(i) = Cj(i) - fshD(i)
        If Lj(i) > 0# Then Tard(i) = Lj(i) Else Tard(i) = 0#
        wT(i) = fshW(i) * Tard(i)
        avgFlow = avgFlow + Flow(i): avgT = avgT + Tard(i): sumWT = sumWT + wT(i)
        If Lj(i) > Lmax Then Lmax = Lj(i)
        If Tard(i) > 0.000001 Then lateCount = lateCount + 1
    Next i
    avgFlow = avgFlow / CDbl(n): avgT = avgT / CDbl(n)
    pctLate = lateCount / CDbl(n): pctOnTime = 1# - pctLate
End Sub

' =========================================================
' GANTT
' =========================================================
Private Sub FSH_BuildGantt(ByVal ch As Chart, ByVal n As Long, ByVal m As Long, ByRef c() As Double, ByVal Cmax As Double)
    Dim cats() As Variant: ReDim cats(1 To m)
    Dim k As Long: For k = 1 To m: cats(k) = "M" & k: Next k
    With ch
        .ChartType = xlBarStacked: .HasLegend = False
        On Error Resume Next: Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop: On Error GoTo 0
        Dim base() As Double: ReDim base(1 To m)
        Dim srs As Series
        Set srs = .SeriesCollection.NewSeries: srs.Values = base: srs.XValues = cats
        srs.Format.Fill.Visible = msoFalse: srs.Format.Line.Visible = msoFalse

        Dim offByMac() As Double: ReDim offByMac(1 To m)
        Dim jobColor As Object: Set jobColor = CreateObject("Scripting.Dictionary")
        Dim i As Long, baseCol As Long, gap As Double, setupDur As Double, procDur As Double, startBeforeSetup As Double
        For i = 1 To n
            For k = 1 To m
                setupDur = fshS(i, k): procDur = fshP(i, k)
                startBeforeSetup = c(i, k) - setupDur - procDur
                gap = startBeforeSetup - offByMac(k)
                If gap > 0.000001 Then
                    Set srs = .SeriesCollection.NewSeries
                    srs.Values = FSH_OneHot(m, k, gap): srs.XValues = cats
                    srs.Format.Fill.Visible = msoFalse: srs.Format.Line.Visible = msoFalse
                    offByMac(k) = offByMac(k) + gap
                End If
                baseCol = FSH_ColorForJobName(fshJob(i), jobColor)
                If setupDur > 0.000001 Then
                    Set srs = .SeriesCollection.NewSeries
                    srs.Values = FSH_OneHot(m, k, setupDur): srs.XValues = cats
                    srs.Format.Fill.ForeColor.RGB = FSH_LightTone(baseCol): srs.Format.Line.Visible = msoFalse
                    offByMac(k) = offByMac(k) + setupDur
                End If
                If procDur > 0.000001 Then
                    Set srs = .SeriesCollection.NewSeries
                    srs.Values = FSH_OneHot(m, k, procDur): srs.XValues = cats
                    srs.Format.Fill.ForeColor.RGB = FSH_DarkTone(baseCol): srs.Format.Line.Visible = msoFalse
                    With srs.Points(k)
                        .HasDataLabel = True: .DataLabel.text = fshJob(i): .DataLabel.Font.Size = 9
                    End With
                    offByMac(k) = offByMac(k) + procDur
                End If
            Next k
        Next i

        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.text = "Tiempo (horas desde t=0 = " & FSH_HoraFmt(0) & ")"
        FSH_ConfigurarEjeTiempo ch, Cmax
    End With
End Sub

' =========================================================
' OUTPUT: estructura de indicadores
' =========================================================
Private Sub FSH_DrawIndicatorsStructure(ByVal ws As Worksheet, ByVal indTopRow As Long)
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

Private Sub FSH_DrawJobMetricsTable(ByVal ws As Worksheet, ByVal n As Long, ByVal m As Long, _
        ByVal indTopRow As Long, ByVal firstCol As Long, ByRef firstDataRow As Long)
    Dim headerRow As Long: headerRow = indTopRow + 1: firstDataRow = headerRow + 1
    Dim lastRow As Long: lastRow = firstDataRow + n - 1
    With ws.Cells(indTopRow, firstCol)
        .Value = "Indicadores por job": .Font.Bold = True: .HorizontalAlignment = xlLeft: .Locked = True
    End With
    Dim headers As Variant: headers = Array("Job", "Inicio (t)", "Inicio (hora)", "Cj (t)", "Cj (hora)", "Flow (Cj-rj)", "L (Cj-dj)", "T=max(L,0)", "w*T")
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
' HELPERS DE DIBUJO
' =========================================================
Private Sub FSH_FormatHeaderRow(ByVal ws As Worksheet, ByVal rowN As Long, ByVal col1 As Long, ByVal col2 As Long)
    With ws.Range(ws.Cells(rowN, col1), ws.Cells(rowN, col2))
        .Font.Bold = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin
        .Locked = True
    End With
End Sub

Private Sub FSH_FormatEditableBlock(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal col1 As Long, ByVal col2 As Long, ByVal bgColor As Long)
    With ws.Range(ws.Cells(firstRow, col1), ws.Cells(lastRow, col2))
        .ClearContents
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        If bgColor = 0 Then .Interior.Pattern = xlNone Else .Interior.Color = bgColor
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(160, 160, 160): .Borders.Weight = xlThin
        .Locked = False
    End With
End Sub

Private Sub FSH_ClearDynamicZone(ByVal ws As Worksheet)
    With ws.Range(ws.Cells(DYNAMIC_TOP_ROW, 1), ws.Cells(DYNAMIC_TOP_ROW + INPUT_CLEAR_ROWS, 50))
        On Error Resume Next: .UnMerge: On Error GoTo 0
        .ClearContents: .Borders.LineStyle = xlNone: .Interior.Pattern = xlNone
        .Font.Bold = False: .Font.Italic = False: .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral: .VerticalAlignment = xlCenter: .Validation.Delete
    End With
End Sub

Private Sub FSH_ClearOutputArea(ByVal ws As Worksheet, ByRef L As FSHLayout)
    Dim outStart As Long: outStart = L.instr5Row + 2
    With ws.Range(ws.Cells(outStart, 1), ws.Cells(outStart + 900, 40))
        .ClearContents: .Borders.LineStyle = xlNone: .Interior.Pattern = xlNone
        .Font.Bold = False: .Font.Italic = False: .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral: .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub FSH_DeleteChartIfExists(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.ChartObjects.Count To 1 Step -1: ws.ChartObjects(i).Delete: Next i
    On Error GoTo 0
End Sub

Private Sub FSH_SetJobListNamedRange(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error Resume Next: ws.Parent.Names("FSH_JobList").Delete: On Error GoTo 0
    ws.Parent.Names.Add name:="FSH_JobList", _
        RefersTo:=ws.Range(ws.Cells(firstRow, PAR_C0), ws.Cells(lastRow, PAR_C0))
End Sub

' =========================================================
' CONVERSIÓN HORA <-> TIEMPO
' =========================================================
Private Function FSH_HoraFmt(ByVal t As Double) As String
    ' t en horas desde t=0 (que corresponde a fshStartOfDay del día 1)
    Dim d As Long, horaDec As Double
    d = Int(t / 24#)
    horaDec = fshStartOfDay + (t - d * 24#)
    If horaDec >= 24# Then horaDec = horaDec - 24#: d = d + 1
    Dim H As Long, M As Long
    H = Int(horaDec)
    M = CLng((horaDec - H) * 60#)
    If M = 60 Then M = 0: H = H + 1
    FSH_HoraFmt = Format(H, "00") & ":" & Format(M, "00") & " D" & (d + 1)
End Function

' =========================================================
' COLORES / EJES
' =========================================================
Private Function FSH_OneHot(ByVal m As Long, ByVal idx As Long, ByVal v As Double) As Variant
    Dim a() As Double, t As Long: ReDim a(1 To m)
    For t = 1 To m: a(t) = 0#: Next t
    a(idx) = v: FSH_OneHot = a
End Function
Private Function FSH_ColorForJobName(ByVal jn As String, ByVal jc As Object) As Long
    Dim k As String: k = UCase$(Trim$(jn))
    If Not jc.Exists(k) Then jc.Add k, jc.Count + 1
    FSH_ColorForJobName = FSH_BasePalette(jc(k))
End Function
Private Function FSH_BasePalette(ByVal i As Long) As Long
    Dim p As Variant
    p = Array(RGB(52, 96, 174), RGB(46, 204, 113), RGB(155, 89, 182), RGB(241, 196, 15), _
              RGB(231, 76, 60), RGB(26, 188, 156), RGB(127, 140, 141), RGB(52, 152, 219), _
              RGB(39, 174, 96), RGB(243, 156, 18))
    FSH_BasePalette = p((i - 1) Mod (UBound(p) + 1))
End Function
Private Function FSH_LightTone(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = (c And &HFF): g = (c \ &H100) And &HFF: b = (c \ &H10000) And &HFF
    FSH_LightTone = RGB(IIf(r + 80 < 255, r + 80, 255), IIf(g + 80 < 255, g + 80, 255), IIf(b + 80 < 255, b + 80, 255))
End Function
Private Function FSH_DarkTone(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = (c And &HFF): g = (c \ &H100) And &HFF: b = (c \ &H10000) And &HFF
    FSH_DarkTone = RGB(IIf(r > 40, r - 40, 0), IIf(g > 40, g - 40, 0), IIf(b > 40, b - 40, 0))
End Function
Private Sub FSH_ConfigurarEjeTiempo(ByVal ch As Chart, ByVal maxTime As Double)
    Dim eje As Axis: Set eje = ch.Axes(xlValue)
    If maxTime < 0.000001 Then maxTime = 1
    Dim majorU As Double: majorU = FSH_NiceMajorUnit(maxTime, 30)
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
Private Function FSH_NiceMajorUnit(ByVal mx As Double, ByVal tt As Long) As Double
    If mx <= 0 Then FSH_NiceMajorUnit = 1: Exit Function
    Dim rs As Double: rs = mx / tt: If rs < 1 Then rs = 1
    Dim p10 As Double: p10 = 10 ^ Int(Log(rs) / Log(10))
    Dim fr As Double: fr = rs / p10
    Select Case fr
        Case Is <= 1: FSH_NiceMajorUnit = 1 * p10
        Case Is <= 2: FSH_NiceMajorUnit = 2 * p10
        Case Is <= 5: FSH_NiceMajorUnit = 5 * p10
        Case Else: FSH_NiceMajorUnit = 10 * p10
    End Select
End Function
