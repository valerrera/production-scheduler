
Option Explicit

Private Const SHEET_NAME As String = "JobShop"

Private Const CELL_N As String = "C12"
Private Const CELL_M As String = "C13"
Private Const CELL_MAXOPS As String = "C14"
Private Const CELL_JS_READY As String = "Z2"

Private Const DYNAMIC_TOP_ROW As Long = 15
Private Const INPUT_CLEAR_ROWS As Long = 2500

Private Const P_COL_JOB As Long = 2
Private Const P_COL_R As Long = 3
Private Const P_COL_D As Long = 4
Private Const P_COL_W As Long = 5

Private Const RM_COL_MACH As Long = 2
Private Const RM_COL_RMAQ As Long = 3

Private Const MAT_COL_LABEL As Long = 2
Private Const MAT_COL_FIRST As Long = 3

Private Const MAX_JOBS As Long = 400
Private Const MAX_MACHS As Long = 100

Private Const CHART_NAME_TIMELINE As String = "chTimeline_JS"

Private Type JSLayout
    pTitleRow As Long
    pHeaderRow As Long
    pFirstRow As Long
    pLastRow As Long

    rmaqTitleRow As Long
    rmaqHeaderRow As Long
    rmaqFirstRow As Long
    rmaqLastRow As Long

    routeTitleRow As Long
    routeHelpRow As Long
    routeHeaderRow As Long
    routeFirstRow As Long
    routeLastRow As Long

    pmatTitleRow As Long
    pmatHeaderRow As Long
    pmatFirstRow As Long
    pmatLastRow As Long

    smatTitleRow As Long
    smatHeaderRow As Long
    smatFirstRow As Long
    smatLastRow As Long

    decTitleRow As Long
    decHeaderRow As Long
    decFirstRow As Long
    decLastRow As Long

    instrCargaRow As Long
    instrGenRow As Long
End Type

' CACHE
Private jsLoaded As Boolean
Private jsN As Long
Private jsM As Long
Private jsMaxOps As Long
Private jsTotalOps As Long

Private jsJobs() As String
Private jsR() As Double
Private jsD() As Double
Private jsW() As Double
Private jsRmaq() As Double

Private jsOpJob() As Long
Private jsOpNum() As Long
Private jsOpMach() As Long
Private jsOpP() As Double
Private jsOpS() As Double

Private jsOpsPerJob() As Long
Private jsOpIdx() As Long
Private jsOpsOnMach() As Long

Private jsMachSeq() As Long
Private jsMaxDecCols As Long

' =========================================================
Public Sub RedibujarInputs_JS(ByVal ws As Worksheet)
    On Error GoTo ErrH

    Dim n As Long, m As Long, maxOps As Long
    If Not JS_ReadNMO(ws, n, m, maxOps) Then Exit Sub

    ws.Unprotect
    jsLoaded = False: jsN = 0: jsM = 0: jsMaxOps = 0: jsTotalOps = 0
    ws.Range(CELL_JS_READY).Value = ""

    JS_ClearDynamicZone ws
    JS_DeleteChartIfExists ws

    Dim L As JSLayout
    Dim maxDecCols As Long: maxDecCols = n * maxOps
    JS_GetLayout n, m, maxOps, maxDecCols, L

    ' Tabla 1
    JS_DrawSectionTitle ws, L.pTitleRow, P_COL_JOB, "2. Parámetros por job (escribe r, d y w)."
    JS_DrawHeaderRow ws, L.pHeaderRow, P_COL_JOB, P_COL_W, Array("Job", "r", "d", "w")
    JS_DrawEditableBlock ws, L.pFirstRow, L.pLastRow, P_COL_JOB, P_COL_W

    Dim i As Long
    For i = 1 To n
        ws.Cells(L.pFirstRow + i - 1, P_COL_JOB).Value = "J" & i
        ws.Cells(L.pFirstRow + i - 1, P_COL_W).Value = 1
    Next i
    ws.Range(ws.Cells(L.pFirstRow, P_COL_JOB), ws.Cells(L.pLastRow, P_COL_JOB)).Locked = True
    JS_SetJobListNamedRange ws, L.pFirstRow, L.pLastRow

    ' Tabla 2
    JS_DrawSectionTitle ws, L.rmaqTitleRow, RM_COL_MACH, "3. Escribe la fecha de disponibilidad de la(s) máquina(s)."
    JS_DrawHeaderRow ws, L.rmaqHeaderRow, RM_COL_MACH, RM_COL_RMAQ, Array("Máquina", "rmaq")
    JS_DrawEditableBlock ws, L.rmaqFirstRow, L.rmaqLastRow, RM_COL_MACH, RM_COL_RMAQ

    Dim k As Long
    For k = 1 To m
        ws.Cells(L.rmaqFirstRow + k - 1, RM_COL_MACH).Value = "M" & k
    Next k
    ws.Range(ws.Cells(L.rmaqFirstRow, RM_COL_MACH), ws.Cells(L.rmaqLastRow, RM_COL_MACH)).Locked = True

    ' Tabla 3: Ruta
    JS_DrawSectionTitle ws, L.routeTitleRow, MAT_COL_LABEL, "4. Ruta [PARÁMETRO] — escribe la máquina que usa cada job en cada operación."
    With ws.Cells(L.routeHelpRow, MAT_COL_LABEL)
        .Value = "Op 1, Op 2, ... son las operaciones en orden tecnológico. Si el job tiene menos operaciones, deja vacías las de la derecha (sin huecos). Se permite repetir máquina."
        .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With
    JS_DrawMatrix ws, L.routeHeaderRow, L.routeFirstRow, L.routeLastRow, n, maxOps, "Job", "Op", True, RGB(248, 248, 248)
    JS_PrellenarLabelsJobs ws, L.routeFirstRow, n
    JS_ApplyMachineDropdownMatrix ws, m, L.routeFirstRow, L.routeLastRow, MAT_COL_FIRST, MAT_COL_FIRST + maxOps - 1

    ' Tabla 3': p
    JS_DrawSectionTitle ws, L.pmatTitleRow, MAT_COL_LABEL, "5. Tiempo de procesamiento p [PARÁMETRO]."
    JS_DrawMatrix ws, L.pmatHeaderRow, L.pmatFirstRow, L.pmatLastRow, n, maxOps, "Job", "Op", True, RGB(248, 248, 248)
    JS_PrellenarLabelsJobs ws, L.pmatFirstRow, n

    ' Tabla 3'': s
    JS_DrawSectionTitle ws, L.smatTitleRow, MAT_COL_LABEL, "6. Tiempo de setup s [PARÁMETRO]."
    JS_DrawMatrix ws, L.smatHeaderRow, L.smatFirstRow, L.smatLastRow, n, maxOps, "Job", "Op", True, RGB(248, 248, 248)
    JS_PrellenarLabelsJobs ws, L.smatFirstRow, n

    ' Tabla 4: Decisión
    JS_DrawSectionTitle ws, L.decTitleRow, MAT_COL_LABEL, "7. Secuencia por máquina [DECISIÓN] — elige el orden en que cada máquina procesa los jobs."
    JS_DrawMatrix ws, L.decHeaderRow, L.decFirstRow, L.decLastRow, m, maxDecCols, "Máquina", "Pos", False, 0
    JS_PrellenarLabelsMachs ws, L.decFirstRow, m
    JS_ApplyJobDropdownMatrix ws, L.decFirstRow, L.decLastRow, MAT_COL_FIRST, MAT_COL_FIRST + maxDecCols - 1

    With ws.Cells(L.instrCargaRow, P_COL_JOB)
        .Value = "8. Presione el botón Cargar datos para validar la información ingresada."
        .Font.Bold = True: .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False
    ws.Range(CELL_M).Locked = False
    ws.Range(CELL_MAXOPS).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en RedibujarInputs_JS: " & Err.Description, vbExclamation
    Resume Salir
End Sub

Public Sub JS_CargarDatos()
    On Error GoTo ErrH

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Dim n As Long, m As Long, maxOps As Long
    If Not JS_ReadNMO(ws, n, m, maxOps) Then
        MsgBox "Verifica que # jobs, # máquinas y # máx ops sean enteros positivos.", vbExclamation: Exit Sub
    End If

    ws.Unprotect
    Dim maxDecCols As Long: maxDecCols = n * maxOps
    Dim L As JSLayout: JS_GetLayout n, m, maxOps, maxDecCols, L
    JS_ClearOutputArea ws, L: JS_DeleteChartIfExists ws
    ws.Range(CELL_JS_READY).Value = ""

    Dim warn As String
    Dim jobs() As String, rJobArr() As Double, d() As Double, w() As Double
    If Not JS_ReadAndValidateJobTable(ws, n, L, jobs, rJobArr, d, w, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos": GoTo Salir
    End If

    Dim rmaq() As Double
    If Not JS_ReadAndValidateRmaq(ws, m, L, rmaq, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos": GoTo Salir
    End If

    Dim opJob() As Long, opNum() As Long, opMach() As Long, opP() As Double, opS() As Double
    Dim opsPerJob() As Long, opsOnMach() As Long, totalOps As Long, opIdxMat() As Long
    If Not JS_ReadAndValidateRouteMatrices(ws, n, m, maxOps, L, opJob, opNum, opMach, opP, opS, _
                                           opsPerJob, opsOnMach, opIdxMat, totalOps, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos": GoTo Salir
    End If

    Dim machSeq() As Long
    If Not JS_ReadAndValidateDecisionMatrix(ws, n, m, maxDecCols, L, opsOnMach, opMach, opJob, opNum, jobs, opsPerJob, machSeq, warn) Then
        MsgBox warn, vbExclamation, "Revisar decisión": GoTo Salir
    End If

    JS_StoreInCache n, m, maxOps, totalOps, maxDecCols, jobs, rJobArr, d, w, rmaq, _
                    opJob, opNum, opMach, opP, opS, opsPerJob, opsOnMach, opIdxMat, machSeq

    Dim st() As Double, ct() As Double, Cmax As Double
    If Not JS_ScheduleFromCache(st, ct, Cmax, warn) Then
        MsgBox warn, vbExclamation, "Revisar decisión": jsLoaded = False: GoTo Salir
    End If

    jsLoaded = True
    ws.Range(CELL_JS_READY).Value = "OK"
    With ws.Cells(L.instrGenRow, P_COL_JOB)
        .Value = "9. Datos válidos. Presione el botón Generar outputs."
        .Font.Bold = True: .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False: ws.Range(CELL_MAXOPS).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en JS_CargarDatos: " & Err.Description, vbExclamation: Resume Salir
End Sub

Public Sub JS_GenerarOutputs()
    On Error GoTo ErrH

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If UCase$(Trim$(CStr(ws.Range(CELL_JS_READY).Value))) <> "OK" Then
        MsgBox "Primero presione 'Cargar datos'.", vbExclamation: Exit Sub
    End If

    Dim n As Long, m As Long, maxOps As Long
    If Not JS_ReadNMO(ws, n, m, maxOps) Then Exit Sub
    If Not jsLoaded Or jsN <> n Or jsM <> m Or jsMaxOps <> maxOps Then
        MsgBox "Presione 'Cargar datos' nuevamente.", vbExclamation: Exit Sub
    End If

    ws.Unprotect
    Dim L As JSLayout: JS_GetLayout n, m, maxOps, jsMaxDecCols, L
    JS_DeleteChartIfExists ws: JS_ClearOutputArea ws, L

    Dim outLineRow As Long: outLineRow = L.instrGenRow + 2
    Dim outTitleRow As Long: outTitleRow = outLineRow + 2
    Dim ganttTopRow As Long: ganttTopRow = outTitleRow + 2
    Dim indTopRow As Long: indTopRow = ganttTopRow + 16 + 6

    With ws.Range(ws.Cells(outLineRow, 1), ws.Cells(outLineRow, 40)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(0, 0, 0): .Weight = xlThin
    End With
    With ws.Cells(outTitleRow, P_COL_JOB)
        .Value = "ZONA DE OUTPUTS": .Font.Bold = True: .HorizontalAlignment = xlLeft: .Locked = True
    End With

    Dim st() As Double, ct() As Double, Cmax As Double, warn As String
    If Not JS_ScheduleFromCache(st, ct, Cmax, warn) Then
        MsgBox warn, vbExclamation: GoTo Salir
    End If

    Dim startJob() As Double, Cj() As Double, Flow() As Double, L2() As Double, Tard() As Double, wT() As Double
    Dim avgFlow As Double, Lmax As Double, avgT As Double, sumWT As Double
    Dim lateCount As Long, pctLate As Double, pctOnTime As Double
    JS_ComputeJobMetrics st, ct, startJob, Cj, Flow, L2, Tard, wT, avgFlow, Lmax, avgT, sumWT, lateCount, pctLate, pctOnTime

    Dim chObj As ChartObject
    Set chObj = ws.ChartObjects.Add(Left:=ws.Cells(ganttTopRow, 2).Left, top:=ws.Cells(ganttTopRow, 2).top, Width:=1100, Height:=240 + 24 * m)
    chObj.name = CHART_NAME_TIMELINE
    With chObj.Chart
        .ChartType = xlBarStacked: .HasTitle = True: .ChartTitle.text = "Gantt Job Shop": .HasLegend = False
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
    End With

    JS_DrawIndicatorsStructure ws, indTopRow
    Dim jmFirstCol As Long: jmFirstCol = 8
    Dim jmFirstDataRow As Long
    JS_DrawJobMetricsTable ws, n, indTopRow, jmFirstCol, jmFirstDataRow

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
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol).Value = jsJobs(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 1).Value = startJob(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 2).Value = Cj(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 3).Value = Flow(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 4).Value = L2(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 5).Value = Tard(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 6).Value = wT(j)
    Next j

    JS_BuildGantt chObj.Chart, st, ct, Cmax

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False: ws.Range(CELL_MAXOPS).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en JS_GenerarOutputs: " & Err.Description, vbExclamation: Resume Salir
End Sub

' =========================================================
' LAYOUT
' =========================================================
Private Sub JS_GetLayout(ByVal n As Long, ByVal m As Long, ByVal maxOps As Long, ByVal maxDecCols As Long, ByRef L As JSLayout)
    L.pTitleRow = 16: L.pHeaderRow = 17: L.pFirstRow = 18: L.pLastRow = L.pFirstRow + n - 1
    L.rmaqTitleRow = L.pLastRow + 3: L.rmaqHeaderRow = L.rmaqTitleRow + 1: L.rmaqFirstRow = L.rmaqHeaderRow + 1: L.rmaqLastRow = L.rmaqFirstRow + m - 1
    L.routeTitleRow = L.rmaqLastRow + 3: L.routeHelpRow = L.routeTitleRow + 1: L.routeHeaderRow = L.routeHelpRow + 1: L.routeFirstRow = L.routeHeaderRow + 1: L.routeLastRow = L.routeFirstRow + n - 1
    L.pmatTitleRow = L.routeLastRow + 3: L.pmatHeaderRow = L.pmatTitleRow + 1: L.pmatFirstRow = L.pmatHeaderRow + 1: L.pmatLastRow = L.pmatFirstRow + n - 1
    L.smatTitleRow = L.pmatLastRow + 3: L.smatHeaderRow = L.smatTitleRow + 1: L.smatFirstRow = L.smatHeaderRow + 1: L.smatLastRow = L.smatFirstRow + n - 1
    L.decTitleRow = L.smatLastRow + 3: L.decHeaderRow = L.decTitleRow + 1: L.decFirstRow = L.decHeaderRow + 1: L.decLastRow = L.decFirstRow + m - 1
    L.instrCargaRow = L.decLastRow + 3: L.instrGenRow = L.instrCargaRow + 2
End Sub

' =========================================================
' LECTURA Y VALIDACIÓN
' =========================================================
Public Function JS_ReadNMO(ByVal ws As Worksheet, ByRef n As Long, ByRef m As Long, ByRef maxOps As Long) As Boolean
    JS_ReadNMO = False
    If Not IsNumeric(ws.Range(CELL_N).Value) Then Exit Function
    If Not IsNumeric(ws.Range(CELL_M).Value) Then Exit Function
    If Not IsNumeric(ws.Range(CELL_MAXOPS).Value) Then Exit Function
    n = CLng(ws.Range(CELL_N).Value): m = CLng(ws.Range(CELL_M).Value): maxOps = CLng(ws.Range(CELL_MAXOPS).Value)
    If n <= 0 Or m <= 0 Or maxOps <= 0 Then Exit Function
    If n > MAX_JOBS Or m > MAX_MACHS Then Exit Function
    JS_ReadNMO = True
End Function

Private Function JS_ReadAndValidateJobTable(ByVal ws As Worksheet, ByVal n As Long, ByRef L As JSLayout, _
        ByRef jobs() As String, ByRef rJobArr() As Double, ByRef d() As Double, ByRef w() As Double, ByRef warn As String) As Boolean
    JS_ReadAndValidateJobTable = False: warn = ""
    ReDim jobs(1 To n): ReDim rJobArr(1 To n): ReDim d(1 To n): ReDim w(1 To n)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, rr As Long, jb As String
    For i = 1 To n
        rr = L.pFirstRow + i - 1: jb = UCase$(Trim$(CStr(ws.Cells(rr, P_COL_JOB).Value)))
        If Len(jb) = 0 Then warn = "Falta Job (fila " & rr & ").": Exit Function
        If dict.Exists(jb) Then warn = "Job repetido: '" & jb & "'.": Exit Function
        dict.Add jb, True: jobs(i) = jb
        If Not IsNumeric(ws.Cells(rr, P_COL_R).Value) Or CDbl(ws.Cells(rr, P_COL_R).Value) < 0 Then warn = "r inválido para " & jb & ".": Exit Function
        rJobArr(i) = CDbl(ws.Cells(rr, P_COL_R).Value)
        If Not IsNumeric(ws.Cells(rr, P_COL_D).Value) Or CDbl(ws.Cells(rr, P_COL_D).Value) <= 0 Then warn = "d debe ser > 0 para " & jb & ".": Exit Function
        d(i) = CDbl(ws.Cells(rr, P_COL_D).Value)
        If Not IsNumeric(ws.Cells(rr, P_COL_W).Value) Or CDbl(ws.Cells(rr, P_COL_W).Value) <= 0 Then warn = "w debe ser > 0 para " & jb & ".": Exit Function
        w(i) = CDbl(ws.Cells(rr, P_COL_W).Value)
    Next i
    JS_ReadAndValidateJobTable = True
End Function

Private Function JS_ReadAndValidateRmaq(ByVal ws As Worksheet, ByVal m As Long, ByRef L As JSLayout, _
        ByRef rmaq() As Double, ByRef warn As String) As Boolean
    JS_ReadAndValidateRmaq = False: warn = "": ReDim rmaq(1 To m)
    Dim k As Long, rv As Variant
    For k = 1 To m
        rv = ws.Cells(L.rmaqFirstRow + k - 1, RM_COL_RMAQ).Value
        If Len(Trim$(CStr(rv))) = 0 Then warn = "Falta rmaq para M" & k & ".": Exit Function
        If Not IsNumeric(rv) Or CDbl(rv) < 0 Then warn = "rmaq inválido para M" & k & ".": Exit Function
        rmaq(k) = CDbl(rv)
    Next k
    JS_ReadAndValidateRmaq = True
End Function

Private Function JS_ReadAndValidateRouteMatrices(ByVal ws As Worksheet, ByVal n As Long, ByVal m As Long, ByVal maxOps As Long, _
        ByRef L As JSLayout, ByRef opJob() As Long, ByRef opNum() As Long, ByRef opMach() As Long, _
        ByRef opP() As Double, ByRef opS() As Double, ByRef opsPerJob() As Long, ByRef opsOnMach() As Long, _
        ByRef opIdxMat() As Long, ByRef totalOps As Long, ByRef warn As String) As Boolean
    JS_ReadAndValidateRouteMatrices = False: warn = ""
    ReDim opsPerJob(1 To n): ReDim opsOnMach(1 To m): ReDim opIdxMat(1 To n, 1 To maxOps)
    totalOps = 0
    Dim j As Long, k As Long, cellVal As Variant, foundEmpty As Boolean, macStr As String, macIdx As Long
    For j = 1 To n
        foundEmpty = False
        For k = 1 To maxOps
            cellVal = ws.Cells(L.routeFirstRow + j - 1, MAT_COL_FIRST + k - 1).Value
            If Len(Trim$(CStr(cellVal))) = 0 Then
                foundEmpty = True
            Else
                If foundEmpty Then warn = "J" & j & " tiene hueco antes de Op " & k & ".": Exit Function
                macStr = UCase$(Trim$(CStr(cellVal)))
                If Not JS_IsValidMachine(macStr, m) Then warn = "Máquina inválida '" & macStr & "' en J" & j & " Op " & k & ".": Exit Function
                macIdx = JS_MachineIndex(macStr)
                opsPerJob(j) = opsPerJob(j) + 1
                opsOnMach(macIdx) = opsOnMach(macIdx) + 1
                totalOps = totalOps + 1
            End If
        Next k
        If opsPerJob(j) = 0 Then warn = "J" & j & " no tiene operaciones.": Exit Function
    Next j

    ReDim opJob(1 To totalOps): ReDim opNum(1 To totalOps): ReDim opMach(1 To totalOps)
    ReDim opP(1 To totalOps): ReDim opS(1 To totalOps)
    Dim cnt As Long: cnt = 0
    Dim pVal As Variant, sVal As Variant
    For j = 1 To n
        For k = 1 To opsPerJob(j)
            cnt = cnt + 1
            macStr = UCase$(Trim$(CStr(ws.Cells(L.routeFirstRow + j - 1, MAT_COL_FIRST + k - 1).Value)))
            macIdx = JS_MachineIndex(macStr)
            opJob(cnt) = j: opNum(cnt) = k: opMach(cnt) = macIdx: opIdxMat(j, k) = cnt
            pVal = ws.Cells(L.pmatFirstRow + j - 1, MAT_COL_FIRST + k - 1).Value
            If Not IsNumeric(pVal) Or CDbl(pVal) <= 0 Then warn = "p debe ser > 0 en J" & j & " Op " & k & ".": Exit Function
            opP(cnt) = CDbl(pVal)
            sVal = ws.Cells(L.smatFirstRow + j - 1, MAT_COL_FIRST + k - 1).Value
            If Not IsNumeric(sVal) Or CDbl(sVal) < 0 Then warn = "s inválido en J" & j & " Op " & k & ".": Exit Function
            opS(cnt) = CDbl(sVal)
        Next k
        For k = opsPerJob(j) + 1 To maxOps
            pVal = ws.Cells(L.pmatFirstRow + j - 1, MAT_COL_FIRST + k - 1).Value
            sVal = ws.Cells(L.smatFirstRow + j - 1, MAT_COL_FIRST + k - 1).Value
            If Len(Trim$(CStr(pVal))) > 0 Then warn = "J" & j & " no tiene Op " & k & " pero tiene p. Borra esa celda.": Exit Function
            If Len(Trim$(CStr(sVal))) > 0 Then warn = "J" & j & " no tiene Op " & k & " pero tiene s. Borra esa celda.": Exit Function
        Next k
    Next j
    JS_ReadAndValidateRouteMatrices = True
End Function

Private Function JS_ReadAndValidateDecisionMatrix(ByVal ws As Worksheet, ByVal n As Long, ByVal m As Long, ByVal maxDecCols As Long, _
        ByRef L As JSLayout, ByRef opsOnMach() As Long, ByRef opMach() As Long, ByRef opJob() As Long, ByRef opNum() As Long, _
        ByRef jobs() As String, ByRef opsPerJob() As Long, ByRef machSeq() As Long, ByRef warn As String) As Boolean
    JS_ReadAndValidateDecisionMatrix = False: warn = ""
    ReDim machSeq(1 To m, 1 To maxDecCols)

    Dim dictJob As Object: Set dictJob = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To n: dictJob(UCase$(jobs(i))) = i: Next i

    Dim totalOps As Long: totalOps = UBound(opMach)
    Dim k As Long, pos As Long, jb As String, cellVal As Variant, foundEmpty As Boolean

    For k = 1 To m
        ' Construir lista de ops disponibles en esta máquina como (jobName|opNum) -> opIndex
        Dim dictAvail As Object: Set dictAvail = CreateObject("Scripting.Dictionary")
        For i = 1 To totalOps
            If opMach(i) = k Then
                dictAvail.Add UCase$(jobs(opJob(i))) & "|" & CStr(opNum(i)), i
            End If
        Next i

        ' Construir conteo de cuántas veces cada job pasa por esta máquina
        Dim dictJobCount As Object: Set dictJobCount = CreateObject("Scripting.Dictionary")
        For i = 1 To totalOps
            If opMach(i) = k Then
                Dim jn As String: jn = UCase$(jobs(opJob(i)))
                If Not dictJobCount.Exists(jn) Then dictJobCount.Add jn, 0
                dictJobCount(jn) = CLng(dictJobCount(jn)) + 1
            End If
        Next i

        ' Conteo de apariciones de cada job en la decisión
        Dim dictUsed As Object: Set dictUsed = CreateObject("Scripting.Dictionary")

        foundEmpty = False
        Dim filled As Long: filled = 0
        For pos = 1 To maxDecCols
            cellVal = ws.Cells(L.decFirstRow + k - 1, MAT_COL_FIRST + pos - 1).Value
            If Len(Trim$(CStr(cellVal))) = 0 Then
                foundEmpty = True
            Else
                If foundEmpty Then warn = "M" & k & " tiene hueco antes de Pos " & pos & ".": Exit Function
                jb = UCase$(Trim$(CStr(cellVal)))
                If Not dictJob.Exists(jb) Then warn = "Job inválido '" & jb & "' en M" & k & " Pos " & pos & ".": Exit Function
                If Not dictJobCount.Exists(jb) Then warn = "'" & jb & "' no pasa por M" & k & ".": Exit Function

                If Not dictUsed.Exists(jb) Then dictUsed.Add jb, 0
                dictUsed(jb) = CLng(dictUsed(jb)) + 1

                If CLng(dictUsed(jb)) > CLng(dictJobCount(jb)) Then
                    warn = "'" & jb & "' aparece más veces de las que pasa por M" & k & "."
                    Exit Function
                End If

                ' Encontrar la op correspondiente (la N-ésima vez que ese job pasa por esta máquina)
                Dim occurrence As Long: occurrence = CLng(dictUsed(jb))
                Dim cnt2 As Long: cnt2 = 0
                Dim found As Long: found = 0
                For i = 1 To totalOps
                    If opMach(i) = k And UCase$(jobs(opJob(i))) = jb Then
                        cnt2 = cnt2 + 1
                        If cnt2 = occurrence Then found = i: Exit For
                    End If
                Next i

                machSeq(k, pos) = found
                filled = filled + 1
            End If
        Next pos

        If filled <> opsOnMach(k) Then warn = "M" & k & " tiene " & filled & " ops pero deben ser " & opsOnMach(k) & ".": Exit Function
    Next k
    JS_ReadAndValidateDecisionMatrix = True
End Function

' =========================================================
' CACHE
' =========================================================
Private Sub JS_StoreInCache(ByVal n As Long, ByVal m As Long, ByVal maxOps As Long, ByVal totalOps As Long, ByVal maxDecCols As Long, _
        ByRef jobs() As String, ByRef rJobArr() As Double, ByRef d() As Double, ByRef w() As Double, ByRef rmaq() As Double, _
        ByRef opJob() As Long, ByRef opNum() As Long, ByRef opMach() As Long, ByRef opP() As Double, ByRef opS() As Double, _
        ByRef opsPerJob() As Long, ByRef opsOnMach() As Long, ByRef opIdxMat() As Long, ByRef machSeq() As Long)
    jsLoaded = False: jsN = n: jsM = m: jsMaxOps = maxOps: jsTotalOps = totalOps: jsMaxDecCols = maxDecCols
    ReDim jsJobs(1 To n): ReDim jsR(1 To n): ReDim jsD(1 To n): ReDim jsW(1 To n): ReDim jsRmaq(1 To m)
    ReDim jsOpJob(1 To totalOps): ReDim jsOpNum(1 To totalOps): ReDim jsOpMach(1 To totalOps)
    ReDim jsOpP(1 To totalOps): ReDim jsOpS(1 To totalOps)
    ReDim jsOpsPerJob(1 To n): ReDim jsOpsOnMach(1 To m)
    ReDim jsOpIdx(1 To n, 1 To maxOps): ReDim jsMachSeq(1 To m, 1 To maxDecCols)
    Dim i As Long, k As Long
    For i = 1 To n: jsJobs(i) = jobs(i): jsR(i) = rJobArr(i): jsD(i) = d(i): jsW(i) = w(i): jsOpsPerJob(i) = opsPerJob(i): Next i
    For k = 1 To m: jsRmaq(k) = rmaq(k): jsOpsOnMach(k) = opsOnMach(k): Next k
    For i = 1 To totalOps: jsOpJob(i) = opJob(i): jsOpNum(i) = opNum(i): jsOpMach(i) = opMach(i): jsOpP(i) = opP(i): jsOpS(i) = opS(i): Next i
    For i = 1 To n: For k = 1 To maxOps: jsOpIdx(i, k) = opIdxMat(i, k): Next k: Next i
    For k = 1 To m: For i = 1 To maxDecCols: jsMachSeq(k, i) = machSeq(k, i): Next i: Next k
End Sub

' =========================================================
' SCHEDULING
' =========================================================
Private Function JS_ScheduleFromCache(ByRef st() As Double, ByRef ct() As Double, ByRef Cmax As Double, ByRef warn As String) As Boolean
    JS_ScheduleFromCache = False: warn = "": Cmax = 0#
    Dim totalOps As Long: totalOps = jsTotalOps
    Dim n As Long: n = jsN: Dim m As Long: m = jsM
    ReDim st(1 To totalOps): ReDim ct(1 To totalOps)

    Dim maxEdges As Long: maxEdges = 2 * totalOps + 20
    Dim head() As Long, toN() As Long, nxt() As Long, indeg() As Long
    ReDim head(1 To totalOps): ReDim toN(1 To maxEdges): ReDim nxt(1 To maxEdges): ReDim indeg(1 To totalOps)
    Dim eCount As Long: eCount = 0
    Dim u As Long, v As Long, j As Long, op As Long, k As Long, pos As Long

    For j = 1 To n
        For op = 2 To jsOpsPerJob(j)
            u = jsOpIdx(j, op - 1): v = jsOpIdx(j, op)
            JS_AddEdge u, v, head, toN, nxt, indeg, eCount
        Next op
    Next j

    For k = 1 To m
        For pos = 2 To jsOpsOnMach(k)
            u = jsMachSeq(k, pos - 1): v = jsMachSeq(k, pos)
            JS_AddEdge u, v, head, toN, nxt, indeg, eCount
        Next pos
    Next k

    Dim est() As Double: ReDim est(1 To totalOps)
    Dim i As Long
    For i = 1 To totalOps
        est(i) = jsR(jsOpJob(i))
        If jsRmaq(jsOpMach(i)) > est(i) Then est(i) = jsRmaq(jsOpMach(i))
    Next i

    Dim q() As Long, qh As Long, qt As Long: ReDim q(1 To totalOps): qh = 1: qt = 0
    For i = 1 To totalOps: If indeg(i) = 0 Then qt = qt + 1: q(qt) = i
    Next i

    Dim processed As Long: processed = 0
    Dim e As Long
    Do While qh <= qt
        u = q(qh): qh = qh + 1: processed = processed + 1
        st(u) = est(u): ct(u) = st(u) + jsOpS(u) + jsOpP(u)
        If ct(u) > Cmax Then Cmax = ct(u)
        e = head(u)
        Do While e <> 0
            v = toN(e)
            If ct(u) > est(v) Then est(v) = ct(u)
            indeg(v) = indeg(v) - 1
            If indeg(v) = 0 Then qt = qt + 1: q(qt) = v
            e = nxt(e)
        Loop
    Loop
    If processed <> totalOps Then warn = "Ciclo en las precedencias.": Exit Function
    JS_ScheduleFromCache = True
End Function

Private Sub JS_AddEdge(ByVal u As Long, ByVal v As Long, ByRef head() As Long, ByRef toN() As Long, ByRef nxt() As Long, ByRef indeg() As Long, ByRef eCount As Long)
    eCount = eCount + 1: toN(eCount) = v: nxt(eCount) = head(u): head(u) = eCount: indeg(v) = indeg(v) + 1
End Sub

' =========================================================
' MÉTRICAS
' =========================================================
Private Sub JS_ComputeJobMetrics(ByRef st() As Double, ByRef ct() As Double, _
        ByRef startJob() As Double, ByRef Cj() As Double, ByRef Flow() As Double, ByRef L() As Double, _
        ByRef Tard() As Double, ByRef wT() As Double, ByRef avgFlow As Double, ByRef Lmax As Double, _
        ByRef avgT As Double, ByRef sumWT As Double, ByRef lateCount As Long, ByRef pctLate As Double, ByRef pctOnTime As Double)
    Dim n As Long: n = jsN: Dim totalOps As Long: totalOps = jsTotalOps
    ReDim startJob(1 To n): ReDim Cj(1 To n): ReDim Flow(1 To n): ReDim L(1 To n): ReDim Tard(1 To n): ReDim wT(1 To n)
    Dim j As Long
    For j = 1 To n: startJob(j) = 1E+30: Cj(j) = 0#: Next j
    Dim i As Long, jj As Long
    For i = 1 To totalOps
        jj = jsOpJob(i)
        If st(i) < startJob(jj) Then startJob(jj) = st(i)
        If ct(i) > Cj(jj) Then Cj(jj) = ct(i)
    Next i
    Lmax = -1E+30: avgFlow = 0#: avgT = 0#: sumWT = 0#: lateCount = 0
    For j = 1 To n
        If startJob(j) > 1E+20 Then startJob(j) = 0#
        Flow(j) = Cj(j) - jsR(j): L(j) = Cj(j) - jsD(j)
        If L(j) > 0# Then Tard(j) = L(j) Else Tard(j) = 0#
        wT(j) = jsW(j) * Tard(j)
        avgFlow = avgFlow + Flow(j): avgT = avgT + Tard(j): sumWT = sumWT + wT(j)
        If L(j) > Lmax Then Lmax = L(j)
        If Tard(j) > 0.000001 Then lateCount = lateCount + 1
    Next j
    avgFlow = avgFlow / CDbl(n): avgT = avgT / CDbl(n)
    pctLate = lateCount / CDbl(n): pctOnTime = 1# - pctLate
End Sub

' =========================================================
' GANTT
' =========================================================
Private Sub JS_BuildGantt(ByVal ch As Chart, ByRef st() As Double, ByRef ct() As Double, ByVal Cmax As Double)
    Dim m As Long: m = jsM: Dim totalOps As Long: totalOps = jsTotalOps
    Dim cats() As Variant: ReDim cats(1 To m)
    Dim k As Long
    For k = 1 To m: cats(k) = "M" & k: Next k
    With ch
        .ChartType = xlBarStacked: .HasLegend = False
        On Error Resume Next: Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop: On Error GoTo 0
        Dim srs As Series, base() As Double: ReDim base(1 To m)
        Set srs = .SeriesCollection.NewSeries: srs.Values = base: srs.XValues = cats: srs.Format.Fill.Visible = msoFalse: srs.Format.Line.Visible = msoFalse

        Dim idx() As Long: ReDim idx(1 To totalOps)
        Dim i As Long: For i = 1 To totalOps: idx(i) = i: Next i
        JS_SortIdxByMachStart totalOps, idx, st

        Dim offByMac() As Double: ReDim offByMac(1 To m)
        Dim jobColor As Object: Set jobColor = CreateObject("Scripting.Dictionary")
        Dim ii As Long, macIdx As Long, gap As Double, baseCol As Long, setupDur As Double, procDur As Double

        For ii = 1 To totalOps
            i = idx(ii): macIdx = jsOpMach(i): setupDur = jsOpS(i): procDur = jsOpP(i)
            gap = st(i) - offByMac(macIdx)
            If gap > 0.000001 Then
                Set srs = .SeriesCollection.NewSeries: srs.Values = JS_OneHot(m, macIdx, gap): srs.XValues = cats
                srs.Format.Fill.Visible = msoFalse: srs.Format.Line.Visible = msoFalse: offByMac(macIdx) = offByMac(macIdx) + gap
            End If
            baseCol = JS_ColorForJobName(jsJobs(jsOpJob(i)), jobColor)
            If setupDur > 0.000001 Then
                Set srs = .SeriesCollection.NewSeries: srs.Values = JS_OneHot(m, macIdx, setupDur): srs.XValues = cats
                srs.Format.Fill.ForeColor.RGB = JS_LightTone(baseCol): srs.Format.Line.Visible = msoFalse: offByMac(macIdx) = offByMac(macIdx) + setupDur
            End If
            If procDur > 0.000001 Then
                Set srs = .SeriesCollection.NewSeries: srs.Values = JS_OneHot(m, macIdx, procDur): srs.XValues = cats
                srs.Format.Fill.ForeColor.RGB = JS_DarkTone(baseCol): srs.Format.Line.Visible = msoFalse
                With srs.Points(macIdx): .HasDataLabel = True: .DataLabel.text = jsJobs(jsOpJob(i)) & " (Op " & jsOpNum(i) & ")": .DataLabel.Font.Size = 9: End With
                offByMac(macIdx) = offByMac(macIdx) + procDur
            End If
        Next ii
        .Axes(xlCategory).ReversePlotOrder = True: .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.text = "Tiempo"
        JS_ConfigurarEjeTiempo ch, Cmax
    End With
End Sub

Private Sub JS_SortIdxByMachStart(ByVal n As Long, ByRef idx() As Long, ByRef st() As Double)
    Dim i As Long, j As Long, t As Long
    For i = 1 To n - 1: For j = i + 1 To n
        If (jsOpMach(idx(j)) < jsOpMach(idx(i))) Or ((jsOpMach(idx(j)) = jsOpMach(idx(i))) And (st(idx(j)) < st(idx(i)))) Then
            t = idx(i): idx(i) = idx(j): idx(j) = t
        End If
    Next j: Next i
End Sub

' =========================================================
' HELPERS DE DIBUJO
' =========================================================
Private Sub JS_DrawSectionTitle(ByVal ws As Worksheet, ByVal rowN As Long, ByVal col As Long, ByVal text As String)
    With ws.Cells(rowN, col): .Value = text: .Font.Bold = True: .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True: End With
End Sub
Private Sub JS_DrawHeaderRow(ByVal ws As Worksheet, ByVal rowN As Long, ByVal col1 As Long, ByVal col2 As Long, ByVal labels As Variant)
    Dim i As Long: For i = 0 To UBound(labels): ws.Cells(rowN, col1 + i).Value = labels(i): Next i
    With ws.Range(ws.Cells(rowN, col1), ws.Cells(rowN, col2)): .Font.Bold = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Interior.Color = RGB(230, 230, 230): .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin: .Locked = True: End With
End Sub
Private Sub JS_DrawEditableBlock(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal col1 As Long, ByVal col2 As Long)
    With ws.Range(ws.Cells(firstRow, col1), ws.Cells(lastRow, col2)): .ClearContents: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Interior.Pattern = xlNone: .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(160, 160, 160): .Borders.Weight = xlThin: .Locked = False: End With
End Sub
Private Sub JS_DrawMatrix(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal firstRow As Long, ByVal lastRow As Long, ByVal numRows As Long, ByVal numCols As Long, ByVal rowLabelHeader As String, ByVal colLabelPrefix As String, ByVal grayBg As Boolean, ByVal bgColor As Long)
    With ws.Cells(headerRow, MAT_COL_LABEL): .Value = rowLabelHeader: .Font.Bold = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Interior.Color = RGB(230, 230, 230): .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin: .Locked = True: End With
    Dim k As Long
    For k = 1 To numCols
        With ws.Cells(headerRow, MAT_COL_FIRST + k - 1): .Value = colLabelPrefix & " " & k: .Font.Bold = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Interior.Color = RGB(230, 230, 230): .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin: .Locked = True: End With
    Next k
    With ws.Range(ws.Cells(firstRow, MAT_COL_LABEL), ws.Cells(lastRow, MAT_COL_FIRST + numCols - 1)): .ClearContents: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(160, 160, 160): .Borders.Weight = xlThin: End With
    With ws.Range(ws.Cells(firstRow, MAT_COL_FIRST), ws.Cells(lastRow, MAT_COL_FIRST + numCols - 1))
        .Locked = False
        If grayBg Then .Interior.Color = bgColor Else .Interior.Pattern = xlNone
    End With
    ws.Range(ws.Cells(firstRow, MAT_COL_LABEL), ws.Cells(lastRow, MAT_COL_LABEL)).Locked = True
End Sub
Private Sub JS_PrellenarLabelsJobs(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal n As Long)
    Dim i As Long: For i = 1 To n: With ws.Cells(firstRow + i - 1, MAT_COL_LABEL): .Value = "J" & i: .Font.Bold = True: .Interior.Color = RGB(230, 230, 230): End With: Next i
End Sub
Private Sub JS_PrellenarLabelsMachs(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal m As Long)
    Dim k As Long: For k = 1 To m: With ws.Cells(firstRow + k - 1, MAT_COL_LABEL): .Value = "M" & k: .Font.Bold = True: .Interior.Color = RGB(230, 230, 230): End With: Next k
End Sub
Private Sub JS_DrawIndicatorsStructure(ByVal ws As Worksheet, ByVal indTopRow As Long)
    With ws.Cells(indTopRow, 2): .Value = "Indicadores": .Font.Bold = True: .HorizontalAlignment = xlLeft: .Locked = True: End With
    Dim labels As Variant: labels = Array("Makespan (Cmax)", "Tiempo de flujo promedio", "Máximo lateness (Lmax)", "Tardanza media", "Tardanza ponderada total", "# Jobs tarde", "% Jobs tarde", "% Jobs a tiempo")
    Dim i As Long
    For i = 0 To UBound(labels)
        With ws.Range(ws.Cells(indTopRow + 1 + i, 2), ws.Cells(indTopRow + 1 + i, 4)): On Error Resume Next: .UnMerge: On Error GoTo 0: .Merge: .Value = labels(i): .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True: End With
        With ws.Cells(indTopRow + 1 + i, 5): .ClearContents: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Locked = True: End With
    Next i
    With ws.Range(ws.Cells(indTopRow + 1, 2), ws.Cells(indTopRow + 8, 5)): .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin: End With
End Sub
Private Sub JS_DrawJobMetricsTable(ByVal ws As Worksheet, ByVal n As Long, ByVal indTopRow As Long, ByVal firstCol As Long, ByRef firstDataRow As Long)
    Dim headerRow As Long: headerRow = indTopRow + 1: firstDataRow = headerRow + 1
    Dim lastRow As Long: lastRow = firstDataRow + n - 1
    With ws.Cells(indTopRow, firstCol): .Value = "Indicadores por job": .Font.Bold = True: .HorizontalAlignment = xlLeft: .Locked = True: End With
    Dim headers As Variant: headers = Array("Job", "Inicio", "Cj", "Flow (Cj-rj)", "L (Cj-dj)", "T=max(L,0)", "w*T")
    Dim h As Long
    For h = 0 To UBound(headers): With ws.Cells(headerRow, firstCol + h): .Value = headers(h): .Font.Bold = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Interior.Color = RGB(230, 230, 230): .Locked = True: End With: Next h
    With ws.Range(ws.Cells(firstDataRow, firstCol), ws.Cells(lastRow, firstCol + 6)): .ClearContents: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Locked = True: End With
    With ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(lastRow, firstCol + 6)): .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin: End With
End Sub

' =========================================================
' UTILIDADES
' =========================================================
Private Sub JS_ClearDynamicZone(ByVal ws As Worksheet)
    With ws.Range(ws.Cells(DYNAMIC_TOP_ROW, 1), ws.Cells(DYNAMIC_TOP_ROW + INPUT_CLEAR_ROWS, 50))
        On Error Resume Next: .UnMerge: On Error GoTo 0
        .ClearContents: .Borders.LineStyle = xlNone: .Interior.Pattern = xlNone: .Font.Bold = False: .Font.Italic = False: .Font.Underline = xlUnderlineStyleNone: .NumberFormat = "General": .HorizontalAlignment = xlGeneral: .VerticalAlignment = xlCenter: .Validation.Delete
    End With
End Sub
Private Sub JS_ClearOutputArea(ByVal ws As Worksheet, ByRef L As JSLayout)
    Dim outStart As Long: outStart = L.instrGenRow + 2
    With ws.Range(ws.Cells(outStart, 1), ws.Cells(outStart + 900, 40)): .ClearContents: .Borders.LineStyle = xlNone: .Interior.Pattern = xlNone: .Font.Bold = False: .Font.Italic = False: .NumberFormat = "General": .HorizontalAlignment = xlGeneral: .VerticalAlignment = xlCenter: End With
End Sub
Private Sub JS_DeleteChartIfExists(ByVal ws As Worksheet)
    Dim i As Long: On Error Resume Next: For i = ws.ChartObjects.Count To 1 Step -1: ws.ChartObjects(i).Delete: Next i: On Error GoTo 0
End Sub
Private Sub JS_ApplyMachineDropdownMatrix(ByVal ws As Worksheet, ByVal m As Long, ByVal fr As Long, ByVal lR As Long, ByVal fC As Long, ByVal lC As Long)
    Dim s As String, k As Long: s = ""
    For k = 1 To m: s = s & "M" & k & IIf(k < m, ",", ""): Next k
    With ws.Range(ws.Cells(fr, fC), ws.Cells(lR, lC)).Validation: .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=s: .IgnoreBlank = True: .InCellDropdown = True: End With
End Sub
Private Sub JS_ApplyJobDropdownMatrix(ByVal ws As Worksheet, ByVal fr As Long, ByVal lR As Long, ByVal fC As Long, ByVal lC As Long)
    With ws.Range(ws.Cells(fr, fC), ws.Cells(lR, lC)).Validation: .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=JS_JobList": .IgnoreBlank = True: .InCellDropdown = True: End With
End Sub
Private Sub JS_SetJobListNamedRange(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error Resume Next: ws.Parent.Names("JS_JobList").Delete: On Error GoTo 0
    ws.Parent.Names.Add name:="JS_JobList", RefersTo:=ws.Range(ws.Cells(firstRow, P_COL_JOB), ws.Cells(lastRow, P_COL_JOB))
End Sub
Private Function JS_IsValidMachine(ByVal s As String, ByVal m As Long) As Boolean
    s = UCase$(Trim$(s)): If Len(s) < 2 Then Exit Function: If Left$(s, 1) <> "M" Then Exit Function
    JS_IsValidMachine = (val(Mid$(s, 2)) >= 1 And val(Mid$(s, 2)) <= m)
End Function
Private Function JS_MachineIndex(ByVal s As String) As Long
    s = UCase$(Trim$(s)): JS_MachineIndex = val(Mid$(s, 2)): If JS_MachineIndex < 1 Then JS_MachineIndex = 1
End Function
Private Function JS_OneHot(ByVal m As Long, ByVal idx As Long, ByVal val As Double) As Variant
    Dim a() As Double, t As Long: ReDim a(1 To m): For t = 1 To m: a(t) = 0#: Next t: a(idx) = val: JS_OneHot = a
End Function
Private Function JS_ColorForJobName(ByVal jn As String, ByVal jc As Object) As Long
    Dim k As String: k = UCase$(Trim$(jn)): If Not jc.Exists(k) Then jc.Add k, jc.Count + 1
    JS_ColorForJobName = JS_BasePalette(jc(k))
End Function
Private Function JS_BasePalette(ByVal i As Long) As Long
    Dim p As Variant: p = Array(RGB(52, 96, 174), RGB(46, 204, 113), RGB(155, 89, 182), RGB(241, 196, 15), RGB(231, 76, 60), RGB(26, 188, 156), RGB(127, 140, 141), RGB(52, 152, 219), RGB(39, 174, 96), RGB(243, 156, 18))
    JS_BasePalette = p((i - 1) Mod (UBound(p) + 1))
End Function
Private Function JS_LightTone(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long: r = (c And &HFF): g = (c \ &H100) And &HFF: b = (c \ &H10000) And &HFF
    JS_LightTone = RGB(IIf(r + 80 < 255, r + 80, 255), IIf(g + 80 < 255, g + 80, 255), IIf(b + 80 < 255, b + 80, 255))
End Function
Private Function JS_DarkTone(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long: r = (c And &HFF): g = (c \ &H100) And &HFF: b = (c \ &H10000) And &HFF
    JS_DarkTone = RGB(IIf(r > 40, r - 40, 0), IIf(g > 40, g - 40, 0), IIf(b > 40, b - 40, 0))
End Function
Private Sub JS_ConfigurarEjeTiempo(ByVal ch As Chart, ByVal maxTime As Double)
    Dim eje As Axis: Set eje = ch.Axes(xlValue)
    If maxTime < 0.000001 Then maxTime = 1
    Dim majorU As Double: majorU = JS_NiceMajorUnit(maxTime, 30)
    Dim minorU As Double: minorU = majorU / 5#
    If minorU < 1# And majorU >= 5# Then minorU = 1#
    eje.MinimumScale = 0
    eje.MaximumScale = Application.WorksheetFunction.Ceiling(maxTime, majorU)
    eje.MajorUnit = majorU
    eje.MinorUnit = minorU
    eje.TickLabelPosition = xlTickLabelPositionNextToAxis
    On Error Resume Next: eje.TickLabels.Orientation = xlHorizontal: On Error GoTo 0
    eje.HasMajorGridlines = True
    eje.HasMinorGridlines = True
    On Error Resume Next
    With eje.MajorGridlines.Format.Line
        .ForeColor.RGB = RGB(180, 180, 180): .Weight = 0.5
    End With
    With eje.MinorGridlines.Format.Line
        .ForeColor.RGB = RGB(225, 225, 225): .DashStyle = msoLineDash: .Weight = 0.25
    End With
    On Error GoTo 0
End Sub
Private Function JS_NiceMajorUnit(ByVal mx As Double, ByVal tt As Long) As Double
    If mx <= 0 Then JS_NiceMajorUnit = 1: Exit Function
    Dim rs As Double: rs = mx / tt: If rs < 1 Then rs = 1
    Dim p10 As Double: p10 = 10 ^ Int(Log(rs) / Log(10))
    Dim fr As Double: fr = rs / p10
    Select Case fr
        Case Is <= 1: JS_NiceMajorUnit = 1 * p10
        Case Is <= 2: JS_NiceMajorUnit = 2 * p10
        Case Is <= 5: JS_NiceMajorUnit = 5 * p10
        Case Else: JS_NiceMajorUnit = 10 * p10
    End Select
End Function


