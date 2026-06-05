Option Explicit

Private Const SHEET_NAME As String = "FlexibleJobShop"

Private Const CELL_N As String = "C12"
Private Const CELL_M As String = "C13"
Private Const CELL_MAXOPS As String = "C14"
Private Const CELL_FJS_READY As String = "Z2"

Private Const DYNAMIC_TOP_ROW As Long = 15
Private Const INPUT_CLEAR_ROWS As Long = 3000

Private Const P_COL_JOB As Long = 2
Private Const P_COL_R As Long = 3
Private Const P_COL_D As Long = 4
Private Const P_COL_W As Long = 5

Private Const WS_COL_LABEL As Long = 2
Private Const WS_COL_C As Long = 3
Private Const WS_COL_RMAQ As Long = 4

Private Const MAT_COL_LABEL As Long = 2
Private Const MAT_COL_FIRST As Long = 3

Private Const MAX_JOBS As Long = 400
Private Const MAX_WS As Long = 50
Private Const MAX_MACHS_PER_WS As Long = 50

Private Const CHART_NAME_TIMELINE As String = "chTimeline_FJS"

Private Type FJSLayout
    pTitleRow As Long: pHeaderRow As Long: pFirstRow As Long: pLastRow As Long
    wsTitleRow As Long: wsHeaderRow As Long: wsFirstRow As Long: wsLastRow As Long
    routeTitleRow As Long: routeHelpRow As Long: routeHeaderRow As Long: routeFirstRow As Long: routeLastRow As Long
    pmatTitleRow As Long: pmatHeaderRow As Long: pmatFirstRow As Long: pmatLastRow As Long
    smatTitleRow As Long: smatHeaderRow As Long: smatFirstRow As Long: smatLastRow As Long
    decTitleRow As Long: decHeaderRow As Long: decFirstRow As Long: decLastRow As Long
    instrCargaRow As Long: instrGenRow As Long
End Type

' CACHE
Private fjsLoaded As Boolean
Private fjsN As Long, fjsM As Long, fjsMaxOps As Long, fjsTotalOps As Long, fjsMaxDecCols As Long, fjsTotalMach As Long
Private fjsJobs() As String, fjsR() As Double, fjsD() As Double, fjsW() As Double
Private fjsCk() As Long, fjsRmaq() As Double
Private fjsOpJob() As Long, fjsOpNum() As Long, fjsOpWS() As Long, fjsOpP() As Double, fjsOpS() As Double
Private fjsOpsPerJob() As Long, fjsOpIdx() As Long, fjsOpsOnWS() As Long
Private fjsMachSeq() As Long      ' (globalMachIdx, pos) -> opIndex
Private fjsOpsOnMach() As Long    ' (globalMachIdx) -> # ops asignados
Private fjsOpAssign() As Long     ' (opIdx) -> machine 1..c_k dentro del workstation (set en validación)
Private fjsGlobalStart() As Long  ' globalStart(k) = globalIdx de la máquina 1 de ws k - 1

' =========================================================
' REDIBUJO PRINCIPAL (al cambiar n, m o maxOps)
' =========================================================
Public Sub RedibujarInputs_FJS(ByVal ws As Worksheet)
    On Error GoTo ErrH
    Dim n As Long, m As Long, maxOps As Long
    If Not FJS_ReadNMO(ws, n, m, maxOps) Then Exit Sub

    ws.Unprotect
    fjsLoaded = False: fjsN = 0: fjsM = 0: fjsMaxOps = 0: fjsTotalOps = 0
    ws.Range(CELL_FJS_READY).Value = ""
    FJS_ClearDynamicZone ws: FJS_DeleteChartIfExists ws

    ' Dibuja inicialmente con c_k = 1 (totalMach = m). Cuando el alumno cambie c_k,
    ' Worksheet_Change llama a FJS_RebuildDecisionTable para reajustar la decisión.
    Dim totalMach As Long: totalMach = m
    Dim maxDecCols As Long: maxDecCols = n * maxOps
    Dim L As FJSLayout: FJS_GetLayout n, m, maxOps, totalMach, maxDecCols, L

    ' Tabla 1
    FJS_DrawSectionTitle ws, L.pTitleRow, P_COL_JOB, "2. Parámetros por job [PARÁMETRO] (escribe r, d y w)."
    FJS_DrawHeaderRow ws, L.pHeaderRow, P_COL_JOB, P_COL_W, Array("Job", "r", "d", "w")
    FJS_DrawEditableBlock ws, L.pFirstRow, L.pLastRow, P_COL_JOB, P_COL_W
    Dim i As Long
    For i = 1 To n
        ws.Cells(L.pFirstRow + i - 1, P_COL_JOB).Value = "J" & i
        ws.Cells(L.pFirstRow + i - 1, P_COL_W).Value = 1
    Next i
    ws.Range(ws.Cells(L.pFirstRow, P_COL_JOB), ws.Cells(L.pLastRow, P_COL_JOB)).Locked = True
    FJS_SetJobListNamedRange ws, L.pFirstRow, L.pLastRow

    ' Tabla 2
    FJS_DrawSectionTitle ws, L.wsTitleRow, WS_COL_LABEL, "3. Workstations [PARÁMETRO] — # máquinas idénticas y rmaq por workstation. Al cambiar c, la Tabla 7 se reajusta."
    FJS_DrawHeaderRow ws, L.wsHeaderRow, WS_COL_LABEL, WS_COL_RMAQ, Array("Workstation", "c (# máquinas)", "rmaq")
    FJS_DrawEditableBlock ws, L.wsFirstRow, L.wsLastRow, WS_COL_LABEL, WS_COL_RMAQ
    Dim k As Long
    For k = 1 To m
        ws.Cells(L.wsFirstRow + k - 1, WS_COL_LABEL).Value = "M" & k
        ws.Cells(L.wsFirstRow + k - 1, WS_COL_C).Value = 1
    Next k
    ws.Range(ws.Cells(L.wsFirstRow, WS_COL_LABEL), ws.Cells(L.wsLastRow, WS_COL_LABEL)).Locked = True

    ' Tablas de ruta, p, s
    FJS_DrawSectionTitle ws, L.routeTitleRow, MAT_COL_LABEL, "4. Ruta [PARÁMETRO] — workstation que usa cada job en cada operación."
    With ws.Cells(L.routeHelpRow, MAT_COL_LABEL)
        .Value = "Op 1, Op 2, ... son las operaciones en orden tecnológico. Si el job tiene menos operaciones, deja vacías las de la derecha (sin huecos). Se permite repetir workstation."
        .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With
    FJS_DrawMatrix ws, L.routeHeaderRow, L.routeFirstRow, L.routeLastRow, n, maxOps, "Job", "Op", True, RGB(248, 248, 248)
    FJS_PrellenarLabelsJobs ws, L.routeFirstRow, n
    FJS_ApplyWorkstationDropdownMatrix ws, m, L.routeFirstRow, L.routeLastRow, MAT_COL_FIRST, MAT_COL_FIRST + maxOps - 1

    FJS_DrawSectionTitle ws, L.pmatTitleRow, MAT_COL_LABEL, "5. Tiempo de procesamiento p [PARÁMETRO]."
    FJS_DrawMatrix ws, L.pmatHeaderRow, L.pmatFirstRow, L.pmatLastRow, n, maxOps, "Job", "Op", True, RGB(248, 248, 248)
    FJS_PrellenarLabelsJobs ws, L.pmatFirstRow, n

    FJS_DrawSectionTitle ws, L.smatTitleRow, MAT_COL_LABEL, "6. Tiempo de setup s [PARÁMETRO]."
    FJS_DrawMatrix ws, L.smatHeaderRow, L.smatFirstRow, L.smatLastRow, n, maxOps, "Job", "Op", True, RGB(248, 248, 248)
    FJS_PrellenarLabelsJobs ws, L.smatFirstRow, n

    ' Tabla 7: decisión por máquina individual (filas Mk.i)
    Dim ck() As Long: ReDim ck(1 To m)
    For k = 1 To m: ck(k) = 1: Next k
    FJS_DrawDecisionTable ws, L, n, m, maxOps, ck

    With ws.Cells(L.instrCargaRow, P_COL_JOB)
        .Value = "8. Presione 'Cargar datos' para validar la información ingresada."
        .Font.Bold = True: .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False: ws.Range(CELL_MAXOPS).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en RedibujarInputs_FJS: " & Err.Description, vbExclamation: Resume Salir
End Sub

' =========================================================
' RECONSTRUYE LA TABLA DE DECISIÓN cuando cambia c_k
' (llamado desde Worksheet_Change en el .cls)
' =========================================================
Public Sub FJS_RebuildDecisionTable(ByVal ws As Worksheet)
    On Error GoTo ErrH
    Dim n As Long, m As Long, maxOps As Long
    If Not FJS_ReadNMO(ws, n, m, maxOps) Then Exit Sub
    ws.Unprotect
    ws.Range(CELL_FJS_READY).Value = ""
    fjsLoaded = False

    Dim L0 As FJSLayout: FJS_GetLayout n, m, maxOps, m, n * maxOps, L0
    ' Lee c_k tolerante (si una celda está vacía o inválida, usa 1)
    Dim ck() As Long: ReDim ck(1 To m)
    Dim k As Long, v As Variant, total As Long: total = 0
    For k = 1 To m
        v = ws.Cells(L0.wsFirstRow + k - 1, WS_COL_C).Value
        If IsNumeric(v) Then
            If CLng(v) >= 1 And CLng(v) <= MAX_MACHS_PER_WS Then ck(k) = CLng(v) Else ck(k) = 1
        Else
            ck(k) = 1
        End If
        total = total + ck(k)
    Next k

    Dim L As FJSLayout: FJS_GetLayout n, m, maxOps, total, n * maxOps, L
    ' Limpia desde decTitle hacia abajo
    With ws.Range(ws.Cells(L.decTitleRow, 1), ws.Cells(L.decTitleRow + 1200, 50))
        On Error Resume Next: .UnMerge: On Error GoTo 0
        .ClearContents: .Borders.LineStyle = xlNone: .Interior.Pattern = xlNone
        .Font.Bold = False: .Font.Italic = False: .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral: .VerticalAlignment = xlCenter: .Validation.Delete
    End With

    FJS_DrawDecisionTable ws, L, n, m, maxOps, ck
    With ws.Cells(L.instrCargaRow, P_COL_JOB)
        .Value = "8. Presione 'Cargar datos' para validar la información ingresada."
        .Font.Bold = True: .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False: ws.Range(CELL_MAXOPS).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en FJS_RebuildDecisionTable: " & Err.Description, vbExclamation: Resume Salir
End Sub

' =========================================================
' Range del c_k para Worksheet_Change
' =========================================================
Public Function FJS_CkRange(ByVal ws As Worksheet) As Range
    Dim n As Long, m As Long, maxOps As Long
    If Not FJS_ReadNMO(ws, n, m, maxOps) Then Exit Function
    Dim L As FJSLayout: FJS_GetLayout n, m, maxOps, m, n * maxOps, L
    Set FJS_CkRange = ws.Range(ws.Cells(L.wsFirstRow, WS_COL_C), ws.Cells(L.wsLastRow, WS_COL_C))
End Function

' =========================================================
Public Sub FJS_CargarDatos()
    On Error GoTo ErrH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Dim n As Long, m As Long, maxOps As Long
    If Not FJS_ReadNMO(ws, n, m, maxOps) Then
        MsgBox "Verifica que # jobs, # workstations y # máx ops sean enteros positivos.", vbExclamation: Exit Sub
    End If

    ws.Unprotect
    Dim warn As String
    Dim jobs() As String, rJobArr() As Double, d() As Double, w() As Double
    Dim L0 As FJSLayout: FJS_GetLayout n, m, maxOps, m, n * maxOps, L0
    If Not FJS_ReadAndValidateJobTable(ws, n, L0, jobs, rJobArr, d, w, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos": GoTo SalirFail
    End If
    Dim ck() As Long, rmaq() As Double
    If Not FJS_ReadAndValidateWorkstations(ws, m, L0, ck, rmaq, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos": GoTo SalirFail
    End If

    ' Layout definitivo con totalMach correcto
    Dim totalMach As Long: totalMach = 0
    Dim k As Long
    For k = 1 To m: totalMach = totalMach + ck(k): Next k
    Dim maxDecCols As Long: maxDecCols = n * maxOps
    Dim L As FJSLayout: FJS_GetLayout n, m, maxOps, totalMach, maxDecCols, L
    FJS_ClearOutputArea ws, L: FJS_DeleteChartIfExists ws
    ws.Range(CELL_FJS_READY).Value = ""

    Dim opJob() As Long, opNum() As Long, opWS() As Long, opP() As Double, opS() As Double
    Dim opsPerJob() As Long, opsOnWS() As Long, totalOps As Long, opIdxMat() As Long
    If Not FJS_ReadAndValidateRouteMatrices(ws, n, m, maxOps, L, opJob, opNum, opWS, opP, opS, _
                                            opsPerJob, opsOnWS, opIdxMat, totalOps, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos": GoTo SalirFail
    End If

    Dim machSeq() As Long, opsOnMach() As Long, opAssign() As Long, globalStart() As Long
    If Not FJS_ReadAndValidateDecisionMatrix(ws, n, m, maxOps, maxDecCols, totalMach, ck, L, _
            opsOnWS, opWS, opJob, opNum, jobs, opsPerJob, _
            machSeq, opsOnMach, opAssign, globalStart, warn) Then
        MsgBox warn, vbExclamation, "Revisar decisión": GoTo SalirFail
    End If

    FJS_StoreInCache n, m, maxOps, totalOps, maxDecCols, totalMach, jobs, rJobArr, d, w, ck, rmaq, _
                     opJob, opNum, opWS, opP, opS, opsPerJob, opsOnWS, opIdxMat, _
                     machSeq, opsOnMach, opAssign, globalStart

    Dim st() As Double, ct() As Double, Cmax As Double
    If Not FJS_ScheduleFromCache(st, ct, Cmax, warn) Then
        MsgBox warn, vbExclamation, "Revisar decisión": fjsLoaded = False: GoTo SalirFail
    End If

    fjsLoaded = True
    ws.Range(CELL_FJS_READY).Value = "OK"
    With ws.Cells(L.instrGenRow, P_COL_JOB)
        .Value = "9. Datos válidos. Presione 'Generar outputs'."
        .Font.Bold = True: .Font.Italic = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False: ws.Range(CELL_MAXOPS).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
SalirFail:
    fjsLoaded = False: GoTo Salir
ErrH:
    MsgBox "Error en FJS_CargarDatos: " & Err.Description, vbExclamation: Resume Salir
End Sub

Public Sub FJS_GenerarOutputs()
    On Error GoTo ErrH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If UCase$(Trim$(CStr(ws.Range(CELL_FJS_READY).Value))) <> "OK" Then
        MsgBox "Primero presione 'Cargar datos'.", vbExclamation: Exit Sub
    End If
    Dim n As Long, m As Long, maxOps As Long
    If Not FJS_ReadNMO(ws, n, m, maxOps) Then Exit Sub
    If Not fjsLoaded Or fjsN <> n Or fjsM <> m Or fjsMaxOps <> maxOps Then
        MsgBox "Presione 'Cargar datos' nuevamente.", vbExclamation: Exit Sub
    End If

    ws.Unprotect
    Dim L As FJSLayout: FJS_GetLayout n, m, maxOps, fjsTotalMach, fjsMaxDecCols, L
    FJS_DeleteChartIfExists ws: FJS_ClearOutputArea ws, L

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
    If Not FJS_ScheduleFromCache(st, ct, Cmax, warn) Then MsgBox warn, vbExclamation: GoTo Salir

    Dim startJob() As Double, Cj() As Double, Flow() As Double, Lj() As Double, Tard() As Double, wT() As Double
    Dim avgFlow As Double, Lmax As Double, avgT As Double, sumWT As Double
    Dim lateCount As Long, pctLate As Double, pctOnTime As Double
    FJS_ComputeJobMetrics st, ct, startJob, Cj, Flow, Lj, Tard, wT, avgFlow, Lmax, avgT, sumWT, lateCount, pctLate, pctOnTime

    Dim chObj As ChartObject
    Set chObj = ws.ChartObjects.Add(Left:=ws.Cells(ganttTopRow, 2).Left, top:=ws.Cells(ganttTopRow, 2).top, _
        Width:=1100, Height:=240 + 22 * fjsTotalMach)
    chObj.name = CHART_NAME_TIMELINE
    With chObj.Chart
        .ChartType = xlBarStacked: .HasTitle = True: .ChartTitle.text = "Gantt Flexible Job Shop": .HasLegend = False
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
    End With

    FJS_DrawIndicatorsStructure ws, indTopRow
    Dim jmFirstCol As Long: jmFirstCol = 8
    Dim jmFirstDataRow As Long
    FJS_DrawJobMetricsTable ws, n, indTopRow, jmFirstCol, jmFirstDataRow

    ws.Cells(indTopRow + 1, 5).Value = Cmax
    ws.Cells(indTopRow + 2, 5).Value = avgFlow
    ws.Cells(indTopRow + 3, 5).Value = Lmax
    ws.Cells(indTopRow + 4, 5).Value = avgT
    ws.Cells(indTopRow + 5, 5).Value = sumWT
    ws.Cells(indTopRow + 6, 5).Value = lateCount
    ws.Cells(indTopRow + 7, 5).Value = pctLate: ws.Cells(indTopRow + 7, 5).NumberFormat = "0%"
    ws.Cells(indTopRow + 8, 5).Value = pctOnTime: ws.Cells(indTopRow + 8, 5).NumberFormat = "0%"

    Dim j As Long, machStr As String, op As Long, opIdx As Long
    For j = 1 To n
        machStr = ""
        For op = 1 To fjsOpsPerJob(j)
            opIdx = fjsOpIdx(j, op)
            machStr = machStr & "M" & fjsOpWS(opIdx) & "." & fjsOpAssign(opIdx)
            If op < fjsOpsPerJob(j) Then machStr = machStr & ", "
        Next op
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol).Value = fjsJobs(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 1).Value = startJob(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 2).Value = Cj(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 3).Value = Flow(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 4).Value = Lj(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 5).Value = Tard(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 6).Value = wT(j)
        ws.Cells(jmFirstDataRow + j - 1, jmFirstCol + 7).Value = machStr
    Next j

    FJS_BuildGantt chObj.Chart, st, ct, Cmax

Salir:
    ws.Range(CELL_N).Locked = False: ws.Range(CELL_M).Locked = False: ws.Range(CELL_MAXOPS).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub
ErrH:
    MsgBox "Error en FJS_GenerarOutputs: " & Err.Description, vbExclamation: Resume Salir
End Sub

' =========================================================
' LAYOUT (decFirst/decLast usan totalMach)
' =========================================================
Private Sub FJS_GetLayout(ByVal n As Long, ByVal m As Long, ByVal maxOps As Long, _
        ByVal totalMach As Long, ByVal maxDecCols As Long, ByRef L As FJSLayout)
    L.pTitleRow = 16: L.pHeaderRow = 17: L.pFirstRow = 18: L.pLastRow = L.pFirstRow + n - 1
    L.wsTitleRow = L.pLastRow + 3: L.wsHeaderRow = L.wsTitleRow + 1: L.wsFirstRow = L.wsHeaderRow + 1: L.wsLastRow = L.wsFirstRow + m - 1
    L.routeTitleRow = L.wsLastRow + 3: L.routeHelpRow = L.routeTitleRow + 1: L.routeHeaderRow = L.routeHelpRow + 1
    L.routeFirstRow = L.routeHeaderRow + 1: L.routeLastRow = L.routeFirstRow + n - 1
    L.pmatTitleRow = L.routeLastRow + 3: L.pmatHeaderRow = L.pmatTitleRow + 1
    L.pmatFirstRow = L.pmatHeaderRow + 1: L.pmatLastRow = L.pmatFirstRow + n - 1
    L.smatTitleRow = L.pmatLastRow + 3: L.smatHeaderRow = L.smatTitleRow + 1
    L.smatFirstRow = L.smatHeaderRow + 1: L.smatLastRow = L.smatFirstRow + n - 1
    L.decTitleRow = L.smatLastRow + 3: L.decHeaderRow = L.decTitleRow + 1
    L.decFirstRow = L.decHeaderRow + 1: L.decLastRow = L.decFirstRow + totalMach - 1
    L.instrCargaRow = L.decLastRow + 3: L.instrGenRow = L.instrCargaRow + 2
End Sub

Public Function FJS_ReadNMO(ByVal ws As Worksheet, ByRef n As Long, ByRef m As Long, ByRef maxOps As Long) As Boolean
    FJS_ReadNMO = False
    If Not IsNumeric(ws.Range(CELL_N).Value) Then Exit Function
    If Not IsNumeric(ws.Range(CELL_M).Value) Then Exit Function
    If Not IsNumeric(ws.Range(CELL_MAXOPS).Value) Then Exit Function
    n = CLng(ws.Range(CELL_N).Value): m = CLng(ws.Range(CELL_M).Value): maxOps = CLng(ws.Range(CELL_MAXOPS).Value)
    If n <= 0 Or m <= 0 Or maxOps <= 0 Then Exit Function
    If n > MAX_JOBS Or m > MAX_WS Then Exit Function
    FJS_ReadNMO = True
End Function

' =========================================================
' DIBUJO DE LA TABLA DE DECISIÓN (totalMach filas Mk.i)
' =========================================================
Private Sub FJS_DrawDecisionTable(ByVal ws As Worksheet, ByRef L As FJSLayout, _
        ByVal n As Long, ByVal m As Long, ByVal maxOps As Long, ByRef ck() As Long)
    Dim totalMach As Long, k As Long, i As Long
    totalMach = 0: For k = 1 To m: totalMach = totalMach + ck(k): Next k
    Dim maxDecCols As Long: maxDecCols = n * maxOps

    FJS_DrawSectionTitle ws, L.decTitleRow, MAT_COL_LABEL, "7. Secuencia por máquina individual [DECISIÓN] — orden de jobs en cada máquina específica."
    FJS_DrawMatrix ws, L.decHeaderRow, L.decFirstRow, L.decLastRow, totalMach, maxDecCols, "Máquina", "Pos", False, 0

    Dim row As Long: row = L.decFirstRow
    For k = 1 To m
        For i = 1 To ck(k)
            With ws.Cells(row, MAT_COL_LABEL)
                .Value = "M" & k & "." & i: .Font.Bold = True: .Interior.Color = RGB(230, 230, 230)
            End With
            row = row + 1
        Next i
    Next k

    FJS_ApplyJobDropdownMatrix ws, L.decFirstRow, L.decLastRow, MAT_COL_FIRST, MAT_COL_FIRST + maxDecCols - 1
End Sub

' =========================================================
' LECTURA Y VALIDACIÓN
' =========================================================
Private Function FJS_ReadAndValidateJobTable(ByVal ws As Worksheet, ByVal n As Long, ByRef L As FJSLayout, _
        ByRef jobs() As String, ByRef rJobArr() As Double, ByRef d() As Double, ByRef w() As Double, ByRef warn As String) As Boolean
    FJS_ReadAndValidateJobTable = False: warn = ""
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
    FJS_ReadAndValidateJobTable = True
End Function

Private Function FJS_ReadAndValidateWorkstations(ByVal ws As Worksheet, ByVal m As Long, ByRef L As FJSLayout, _
        ByRef ck() As Long, ByRef rmaq() As Double, ByRef warn As String) As Boolean
    FJS_ReadAndValidateWorkstations = False: warn = "": ReDim ck(1 To m): ReDim rmaq(1 To m)
    Dim k As Long, cv As Variant, rv As Variant
    For k = 1 To m
        cv = ws.Cells(L.wsFirstRow + k - 1, WS_COL_C).Value
        If Not IsNumeric(cv) Then warn = "c debe ser numérico para M" & k & ".": Exit Function
        If CLng(cv) <> CDbl(cv) Or CLng(cv) <= 0 Then warn = "c debe ser entero ≥ 1 para M" & k & ".": Exit Function
        If CLng(cv) > MAX_MACHS_PER_WS Then warn = "c excede el máximo (" & MAX_MACHS_PER_WS & ") en M" & k & ".": Exit Function
        ck(k) = CLng(cv)
        rv = ws.Cells(L.wsFirstRow + k - 1, WS_COL_RMAQ).Value
        If Not IsNumeric(rv) Or CDbl(rv) < 0 Then warn = "rmaq inválido para M" & k & ".": Exit Function
        rmaq(k) = CDbl(rv)
    Next k
    FJS_ReadAndValidateWorkstations = True
End Function

Private Function FJS_ReadAndValidateRouteMatrices(ByVal ws As Worksheet, ByVal n As Long, ByVal m As Long, ByVal maxOps As Long, _
        ByRef L As FJSLayout, ByRef opJob() As Long, ByRef opNum() As Long, ByRef opWS() As Long, _
        ByRef opP() As Double, ByRef opS() As Double, ByRef opsPerJob() As Long, ByRef opsOnWS() As Long, _
        ByRef opIdxMat() As Long, ByRef totalOps As Long, ByRef warn As String) As Boolean
    FJS_ReadAndValidateRouteMatrices = False: warn = ""
    ReDim opsPerJob(1 To n): ReDim opsOnWS(1 To m): ReDim opIdxMat(1 To n, 1 To maxOps)
    totalOps = 0
    Dim j As Long, k As Long, cellVal As Variant, foundEmpty As Boolean, wsStr As String, wsIdx As Long
    For j = 1 To n
        foundEmpty = False
        For k = 1 To maxOps
            cellVal = ws.Cells(L.routeFirstRow + j - 1, MAT_COL_FIRST + k - 1).Value
            If Len(Trim$(CStr(cellVal))) = 0 Then
                foundEmpty = True
            Else
                If foundEmpty Then warn = "J" & j & " tiene hueco antes de Op " & k & ".": Exit Function
                wsStr = UCase$(Trim$(CStr(cellVal)))
                If Not FJS_IsValidWS(wsStr, m) Then warn = "Workstation inválido '" & wsStr & "' en J" & j & " Op " & k & ".": Exit Function
                wsIdx = FJS_WSIndex(wsStr)
                opsPerJob(j) = opsPerJob(j) + 1
                opsOnWS(wsIdx) = opsOnWS(wsIdx) + 1
                totalOps = totalOps + 1
            End If
        Next k
        If opsPerJob(j) = 0 Then warn = "J" & j & " no tiene operaciones.": Exit Function
    Next j
    ReDim opJob(1 To totalOps): ReDim opNum(1 To totalOps): ReDim opWS(1 To totalOps)
    ReDim opP(1 To totalOps): ReDim opS(1 To totalOps)
    Dim cnt As Long: cnt = 0
    Dim pVal As Variant, sVal As Variant
    For j = 1 To n
        For k = 1 To opsPerJob(j)
            cnt = cnt + 1
            wsStr = UCase$(Trim$(CStr(ws.Cells(L.routeFirstRow + j - 1, MAT_COL_FIRST + k - 1).Value)))
            wsIdx = FJS_WSIndex(wsStr)
            opJob(cnt) = j: opNum(cnt) = k: opWS(cnt) = wsIdx: opIdxMat(j, k) = cnt
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
            If Len(Trim$(CStr(pVal))) > 0 Then warn = "J" & j & " no tiene Op " & k & " pero tiene p.": Exit Function
            If Len(Trim$(CStr(sVal))) > 0 Then warn = "J" & j & " no tiene Op " & k & " pero tiene s.": Exit Function
        Next k
    Next j
    FJS_ReadAndValidateRouteMatrices = True
End Function

' Cada workstation k tiene c_k máquinas (filas Mk.1..Mk.ck).
' Por workstation, la N-ésima aparición de un job J (en orden Mk.1 pos1..maxDecCols, luego Mk.2, ...)
' corresponde a la N-ésima operación de J que pasa por k (orden de la ruta).
' Asigna a cada operación la máquina específica (1..c_k) que le tocó.
Private Function FJS_ReadAndValidateDecisionMatrix(ByVal ws As Worksheet, _
        ByVal n As Long, ByVal m As Long, ByVal maxOps As Long, ByVal maxDecCols As Long, _
        ByVal totalMach As Long, ByRef ck() As Long, ByRef L As FJSLayout, _
        ByRef opsOnWS() As Long, ByRef opWS() As Long, ByRef opJob() As Long, ByRef opNum() As Long, _
        ByRef jobs() As String, ByRef opsPerJob() As Long, _
        ByRef machSeq() As Long, ByRef opsOnMach() As Long, ByRef opAssign() As Long, _
        ByRef globalStart() As Long, ByRef warn As String) As Boolean
    FJS_ReadAndValidateDecisionMatrix = False: warn = ""

    ReDim machSeq(1 To totalMach, 1 To maxDecCols)
    ReDim opsOnMach(1 To totalMach)
    Dim totalOps As Long: totalOps = UBound(opWS)
    ReDim opAssign(1 To totalOps)
    ReDim globalStart(1 To m)
    Dim k As Long, sum As Long: sum = 0
    For k = 1 To m: globalStart(k) = sum: sum = sum + ck(k): Next k

    Dim dictJobIdx As Object: Set dictJobIdx = CreateObject("Scripting.Dictionary")
    Dim i As Long: For i = 1 To n: dictJobIdx(UCase$(jobs(i))) = i: Next i

    ' Para cada workstation, lista ordenada por ruta de las ops de cada job que pasan por k
    For k = 1 To m
        ' Construye dictJobOpsList: jobName -> array de opIdx (en orden de ruta)
        Dim dictJobOpsList As Object: Set dictJobOpsList = CreateObject("Scripting.Dictionary")
        For i = 1 To totalOps
            If opWS(i) = k Then
                Dim jn As String: jn = UCase$(jobs(opJob(i)))
                If Not dictJobOpsList.Exists(jn) Then dictJobOpsList(jn) = ""
                dictJobOpsList(jn) = dictJobOpsList(jn) & i & ","
            End If
        Next i

        ' Contador de cuántas veces hemos visto cada job en este ws (en la decisión)
        Dim dictJobSeen As Object: Set dictJobSeen = CreateObject("Scripting.Dictionary")

        Dim mi As Long, pos As Long, cellVal As Variant, jb As String
        Dim filledInWS As Long: filledInWS = 0
        For mi = 1 To ck(k)
            Dim foundEmpty As Boolean: foundEmpty = False
            Dim filledInMach As Long: filledInMach = 0
            Dim gIdx As Long: gIdx = globalStart(k) + mi
            For pos = 1 To maxDecCols
                cellVal = ws.Cells(L.decFirstRow + gIdx - 1, MAT_COL_FIRST + pos - 1).Value
                If Len(Trim$(CStr(cellVal))) = 0 Then
                    foundEmpty = True
                Else
                    If foundEmpty Then warn = "M" & k & "." & mi & " tiene hueco antes de Pos " & pos & ".": Exit Function
                    jb = UCase$(Trim$(CStr(cellVal)))
                    If Not dictJobIdx.Exists(jb) Then warn = "Job inválido '" & jb & "' en M" & k & "." & mi & " Pos " & pos & ".": Exit Function
                    If Not dictJobOpsList.Exists(jb) Then warn = "'" & jb & "' no pasa por workstation M" & k & ".": Exit Function
                    If Not dictJobSeen.Exists(jb) Then dictJobSeen(jb) = 0
                    dictJobSeen(jb) = CLng(dictJobSeen(jb)) + 1
                    ' Obtener N-ésima op de J en ws k según orden de ruta
                    Dim list As String: list = CStr(dictJobOpsList(jb))
                    Dim parts() As String: parts = Split(list, ",")
                    Dim occurrence As Long: occurrence = CLng(dictJobSeen(jb))
                    Dim opsInWS As Long: opsInWS = UBound(parts)  ' parts(UBound) es "" porque list termina en coma
                    If occurrence > opsInWS Then warn = "'" & jb & "' aparece más veces de las que pasa por M" & k & ".": Exit Function
                    Dim opIdx As Long: opIdx = CLng(parts(occurrence - 1))
                    machSeq(gIdx, pos) = opIdx
                    opAssign(opIdx) = mi
                    filledInMach = filledInMach + 1: filledInWS = filledInWS + 1
                End If
            Next pos
            opsOnMach(gIdx) = filledInMach
        Next mi
        If filledInWS <> opsOnWS(k) Then warn = "Workstation M" & k & " tiene " & filledInWS & " ops asignadas pero deben ser " & opsOnWS(k) & ".": Exit Function
    Next k
    FJS_ReadAndValidateDecisionMatrix = True
End Function

' =========================================================
' CACHE
' =========================================================
Private Sub FJS_StoreInCache(ByVal n As Long, ByVal m As Long, ByVal maxOps As Long, ByVal totalOps As Long, _
        ByVal maxDecCols As Long, ByVal totalMach As Long, _
        ByRef jobs() As String, ByRef rJobArr() As Double, ByRef d() As Double, ByRef w() As Double, _
        ByRef ck() As Long, ByRef rmaq() As Double, _
        ByRef opJob() As Long, ByRef opNum() As Long, ByRef opWS() As Long, ByRef opP() As Double, ByRef opS() As Double, _
        ByRef opsPerJob() As Long, ByRef opsOnWS() As Long, ByRef opIdxMat() As Long, _
        ByRef machSeq() As Long, ByRef opsOnMach() As Long, ByRef opAssign() As Long, ByRef globalStart() As Long)
    fjsLoaded = False
    fjsN = n: fjsM = m: fjsMaxOps = maxOps: fjsTotalOps = totalOps
    fjsMaxDecCols = maxDecCols: fjsTotalMach = totalMach
    ReDim fjsJobs(1 To n): ReDim fjsR(1 To n): ReDim fjsD(1 To n): ReDim fjsW(1 To n)
    ReDim fjsCk(1 To m): ReDim fjsRmaq(1 To m): ReDim fjsGlobalStart(1 To m)
    ReDim fjsOpJob(1 To totalOps): ReDim fjsOpNum(1 To totalOps): ReDim fjsOpWS(1 To totalOps)
    ReDim fjsOpP(1 To totalOps): ReDim fjsOpS(1 To totalOps): ReDim fjsOpAssign(1 To totalOps)
    ReDim fjsOpsPerJob(1 To n): ReDim fjsOpsOnWS(1 To m)
    ReDim fjsOpIdx(1 To n, 1 To maxOps)
    ReDim fjsMachSeq(1 To totalMach, 1 To maxDecCols): ReDim fjsOpsOnMach(1 To totalMach)
    Dim i As Long, k As Long
    For i = 1 To n: fjsJobs(i) = jobs(i): fjsR(i) = rJobArr(i): fjsD(i) = d(i): fjsW(i) = w(i): fjsOpsPerJob(i) = opsPerJob(i): Next i
    For k = 1 To m: fjsCk(k) = ck(k): fjsRmaq(k) = rmaq(k): fjsOpsOnWS(k) = opsOnWS(k): fjsGlobalStart(k) = globalStart(k): Next k
    For i = 1 To totalOps
        fjsOpJob(i) = opJob(i): fjsOpNum(i) = opNum(i): fjsOpWS(i) = opWS(i)
        fjsOpP(i) = opP(i): fjsOpS(i) = opS(i): fjsOpAssign(i) = opAssign(i)
    Next i
    For i = 1 To n: For k = 1 To maxOps: fjsOpIdx(i, k) = opIdxMat(i, k): Next k: Next i
    For i = 1 To totalMach
        fjsOpsOnMach(i) = opsOnMach(i)
        Dim pos As Long: For pos = 1 To maxDecCols: fjsMachSeq(i, pos) = machSeq(i, pos): Next pos
    Next i
End Sub

' =========================================================
' SCHEDULING — Kahn topológico con precedencias por ruta + por máquina (estricta)
' Cada máquina procesa una op a la vez, así que la cola por máquina = precedencia "termina antes que empiece".
' =========================================================
Private Function FJS_ScheduleFromCache(ByRef st() As Double, ByRef ct() As Double, ByRef Cmax As Double, ByRef warn As String) As Boolean
    FJS_ScheduleFromCache = False: warn = "": Cmax = 0#
    Dim totalOps As Long: totalOps = fjsTotalOps
    Dim n As Long: n = fjsN: Dim m As Long: m = fjsM: Dim totalMach As Long: totalMach = fjsTotalMach
    ReDim st(1 To totalOps): ReDim ct(1 To totalOps)

    Dim maxEdges As Long: maxEdges = 2 * totalOps + 20
    Dim head() As Long, toN() As Long, nxt() As Long, indeg() As Long
    ReDim head(1 To totalOps): ReDim toN(1 To maxEdges): ReDim nxt(1 To maxEdges): ReDim indeg(1 To totalOps)
    Dim eCount As Long: eCount = 0
    Dim u As Long, v As Long, j As Long, op As Long, g As Long, pos As Long

    ' Aristas por ruta
    For j = 1 To n
        For op = 2 To fjsOpsPerJob(j)
            u = fjsOpIdx(j, op - 1): v = fjsOpIdx(j, op)
            FJS_AddEdge u, v, head, toN, nxt, indeg, eCount
        Next op
    Next j
    ' Aristas por máquina individual (cola estricta)
    For g = 1 To totalMach
        For pos = 2 To fjsOpsOnMach(g)
            u = fjsMachSeq(g, pos - 1): v = fjsMachSeq(g, pos)
            FJS_AddEdge u, v, head, toN, nxt, indeg, eCount
        Next pos
    Next g

    ' machFree por máquina individual (inicia en rmaq del workstation)
    Dim machFree() As Double: ReDim machFree(1 To totalMach)
    Dim k As Long, i As Long
    For k = 1 To m
        For i = 1 To fjsCk(k)
            machFree(fjsGlobalStart(k) + i) = fjsRmaq(k)
        Next i
    Next k

    Dim est() As Double: ReDim est(1 To totalOps)
    For i = 1 To totalOps: est(i) = fjsR(fjsOpJob(i)): Next i

    Dim q() As Long, qh As Long, qt As Long: ReDim q(1 To totalOps): qh = 1: qt = 0
    For i = 1 To totalOps: If indeg(i) = 0 Then qt = qt + 1: q(qt) = i
    Next i

    Dim processed As Long: processed = 0
    Dim e As Long, gIdx As Long, startU As Double
    Do While qh <= qt
        u = q(qh): qh = qh + 1: processed = processed + 1
        gIdx = fjsGlobalStart(fjsOpWS(u)) + fjsOpAssign(u)
        startU = est(u)
        If machFree(gIdx) > startU Then startU = machFree(gIdx)
        st(u) = startU
        ct(u) = startU + fjsOpS(u) + fjsOpP(u)
        machFree(gIdx) = ct(u)
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
    FJS_ScheduleFromCache = True
End Function

Private Sub FJS_AddEdge(ByVal u As Long, ByVal v As Long, ByRef head() As Long, ByRef toN() As Long, _
        ByRef nxt() As Long, ByRef indeg() As Long, ByRef eCount As Long)
    eCount = eCount + 1: toN(eCount) = v: nxt(eCount) = head(u): head(u) = eCount: indeg(v) = indeg(v) + 1
End Sub

Private Sub FJS_ComputeJobMetrics(ByRef st() As Double, ByRef ct() As Double, _
        ByRef startJob() As Double, ByRef Cj() As Double, ByRef Flow() As Double, ByRef Lj() As Double, _
        ByRef Tard() As Double, ByRef wT() As Double, ByRef avgFlow As Double, ByRef Lmax As Double, _
        ByRef avgT As Double, ByRef sumWT As Double, ByRef lateCount As Long, ByRef pctLate As Double, ByRef pctOnTime As Double)
    Dim n As Long: n = fjsN: Dim totalOps As Long: totalOps = fjsTotalOps
    ReDim startJob(1 To n): ReDim Cj(1 To n): ReDim Flow(1 To n): ReDim Lj(1 To n): ReDim Tard(1 To n): ReDim wT(1 To n)
    Dim j As Long: For j = 1 To n: startJob(j) = 1E+30: Cj(j) = 0#: Next j
    Dim i As Long, jj As Long
    For i = 1 To totalOps
        jj = fjsOpJob(i)
        If st(i) < startJob(jj) Then startJob(jj) = st(i)
        If ct(i) > Cj(jj) Then Cj(jj) = ct(i)
    Next i
    Lmax = -1E+30: avgFlow = 0#: avgT = 0#: sumWT = 0#: lateCount = 0
    For j = 1 To n
        If startJob(j) > 1E+20 Then startJob(j) = 0#
        Flow(j) = Cj(j) - fjsR(j): Lj(j) = Cj(j) - fjsD(j)
        If Lj(j) > 0# Then Tard(j) = Lj(j) Else Tard(j) = 0#
        wT(j) = fjsW(j) * Tard(j)
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
Private Sub FJS_BuildGantt(ByVal ch As Chart, ByRef st() As Double, ByRef ct() As Double, ByVal Cmax As Double)
    Dim m As Long: m = fjsM: Dim totalOps As Long: totalOps = fjsTotalOps
    Dim total As Long: total = fjsTotalMach
    Dim cats() As Variant: ReDim cats(1 To total)
    Dim k As Long, i As Long, p As Long: p = 0
    For k = 1 To m
        For i = 1 To fjsCk(k)
            p = p + 1: cats(p) = "M" & k & "." & i
        Next i
    Next k

    With ch
        .ChartType = xlBarStacked: .HasLegend = False
        On Error Resume Next: Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop: On Error GoTo 0
        Dim base() As Double: ReDim base(1 To total)
        Dim srs As Series
        Set srs = .SeriesCollection.NewSeries: srs.Values = base: srs.XValues = cats
        srs.Format.Fill.Visible = msoFalse: srs.Format.Line.Visible = msoFalse

        Dim idx() As Long: ReDim idx(1 To totalOps)
        For i = 1 To totalOps: idx(i) = i: Next i
        FJS_SortIdxByMachStart totalOps, idx, st

        Dim machOff() As Double: ReDim machOff(1 To total)
        Dim jobColor As Object: Set jobColor = CreateObject("Scripting.Dictionary")
        Dim ii As Long, mIdx As Long, gap As Double, baseCol As Long, setupDur As Double, procDur As Double
        For ii = 1 To totalOps
            i = idx(ii)
            mIdx = fjsGlobalStart(fjsOpWS(i)) + fjsOpAssign(i)
            setupDur = fjsOpS(i): procDur = fjsOpP(i)
            gap = st(i) - machOff(mIdx)
            If gap > 0.000001 Then
                Set srs = .SeriesCollection.NewSeries
                srs.Values = FJS_OneHot(total, mIdx, gap): srs.XValues = cats
                srs.Format.Fill.Visible = msoFalse: srs.Format.Line.Visible = msoFalse
                machOff(mIdx) = machOff(mIdx) + gap
            End If
            baseCol = FJS_ColorForJobName(fjsJobs(fjsOpJob(i)), jobColor)
            If setupDur > 0.000001 Then
                Set srs = .SeriesCollection.NewSeries
                srs.Values = FJS_OneHot(total, mIdx, setupDur): srs.XValues = cats
                srs.Format.Fill.ForeColor.RGB = FJS_LightTone(baseCol): srs.Format.Line.Visible = msoFalse
                machOff(mIdx) = machOff(mIdx) + setupDur
            End If
            If procDur > 0.000001 Then
                Set srs = .SeriesCollection.NewSeries
                srs.Values = FJS_OneHot(total, mIdx, procDur): srs.XValues = cats
                srs.Format.Fill.ForeColor.RGB = FJS_DarkTone(baseCol): srs.Format.Line.Visible = msoFalse
                With srs.Points(mIdx)
                    .HasDataLabel = True: .DataLabel.text = fjsJobs(fjsOpJob(i)) & " (Op " & fjsOpNum(i) & ")": .DataLabel.Font.Size = 9
                End With
                machOff(mIdx) = machOff(mIdx) + procDur
            End If
        Next ii

        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.text = "Tiempo"
        FJS_ConfigurarEjeTiempo ch, Cmax
    End With
End Sub

Private Sub FJS_SortIdxByMachStart(ByVal n As Long, ByRef idx() As Long, ByRef st() As Double)
    Dim i As Long, j As Long, t As Long, ai As Long, aj As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            ai = fjsGlobalStart(fjsOpWS(idx(i))) + fjsOpAssign(idx(i))
            aj = fjsGlobalStart(fjsOpWS(idx(j))) + fjsOpAssign(idx(j))
            If (aj < ai) Or ((aj = ai) And (st(idx(j)) < st(idx(i)))) Then
                t = idx(i): idx(i) = idx(j): idx(j) = t
            End If
        Next j
    Next i
End Sub

' =========================================================
' HELPERS DE DIBUJO
' =========================================================
Private Sub FJS_DrawSectionTitle(ByVal ws As Worksheet, ByVal rowN As Long, ByVal col As Long, ByVal txt As String)
    With ws.Cells(rowN, col)
        .Value = txt: .Font.Bold = True: .Font.Italic = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Locked = True
    End With
End Sub
Private Sub FJS_DrawHeaderRow(ByVal ws As Worksheet, ByVal rowN As Long, ByVal col1 As Long, ByVal col2 As Long, ByVal labels As Variant)
    Dim i As Long: For i = 0 To UBound(labels): ws.Cells(rowN, col1 + i).Value = labels(i): Next i
    With ws.Range(ws.Cells(rowN, col1), ws.Cells(rowN, col2))
        .Font.Bold = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin
        .Locked = True
    End With
End Sub
Private Sub FJS_DrawEditableBlock(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal col1 As Long, ByVal col2 As Long)
    With ws.Range(ws.Cells(firstRow, col1), ws.Cells(lastRow, col2))
        .ClearContents
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Interior.Pattern = xlNone
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(160, 160, 160): .Borders.Weight = xlThin
        .Locked = False
    End With
End Sub
Private Sub FJS_DrawMatrix(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal numRows As Long, ByVal numCols As Long, ByVal rowLabelHeader As String, ByVal colLabelPrefix As String, _
        ByVal grayBg As Boolean, ByVal bgColor As Long)
    With ws.Cells(headerRow, MAT_COL_LABEL)
        .Value = rowLabelHeader: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin: .Locked = True
    End With
    Dim k As Long
    For k = 1 To numCols
        With ws.Cells(headerRow, MAT_COL_FIRST + k - 1)
            .Value = colLabelPrefix & " " & k: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Interior.Color = RGB(230, 230, 230)
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 0, 0): .Borders.Weight = xlThin: .Locked = True
        End With
    Next k
    With ws.Range(ws.Cells(firstRow, MAT_COL_LABEL), ws.Cells(lastRow, MAT_COL_FIRST + numCols - 1))
        .ClearContents
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(160, 160, 160): .Borders.Weight = xlThin
    End With
    With ws.Range(ws.Cells(firstRow, MAT_COL_FIRST), ws.Cells(lastRow, MAT_COL_FIRST + numCols - 1))
        .Locked = False
        If grayBg Then .Interior.Color = bgColor Else .Interior.Pattern = xlNone
    End With
    ws.Range(ws.Cells(firstRow, MAT_COL_LABEL), ws.Cells(lastRow, MAT_COL_LABEL)).Locked = True
End Sub
Private Sub FJS_PrellenarLabelsJobs(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal n As Long)
    Dim i As Long
    For i = 1 To n
        With ws.Cells(firstRow + i - 1, MAT_COL_LABEL)
            .Value = "J" & i: .Font.Bold = True: .Interior.Color = RGB(230, 230, 230)
        End With
    Next i
End Sub
Private Sub FJS_DrawIndicatorsStructure(ByVal ws As Worksheet, ByVal indTopRow As Long)
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
Private Sub FJS_DrawJobMetricsTable(ByVal ws As Worksheet, ByVal n As Long, ByVal indTopRow As Long, ByVal firstCol As Long, ByRef firstDataRow As Long)
    Dim headerRow As Long: headerRow = indTopRow + 1: firstDataRow = headerRow + 1
    Dim lastRow As Long: lastRow = firstDataRow + n - 1
    With ws.Cells(indTopRow, firstCol)
        .Value = "Indicadores por job": .Font.Bold = True: .HorizontalAlignment = xlLeft: .Locked = True
    End With
    Dim headers As Variant: headers = Array("Job", "Inicio", "Cj", "Flow (Cj-rj)", "L (Cj-dj)", "T=max(L,0)", "w*T", "Máquinas (por op)")
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

Private Sub FJS_ClearDynamicZone(ByVal ws As Worksheet)
    With ws.Range(ws.Cells(DYNAMIC_TOP_ROW, 1), ws.Cells(DYNAMIC_TOP_ROW + INPUT_CLEAR_ROWS, 50))
        On Error Resume Next: .UnMerge: On Error GoTo 0
        .ClearContents: .Borders.LineStyle = xlNone: .Interior.Pattern = xlNone
        .Font.Bold = False: .Font.Italic = False: .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral: .VerticalAlignment = xlCenter: .Validation.Delete
    End With
End Sub
Private Sub FJS_ClearOutputArea(ByVal ws As Worksheet, ByRef L As FJSLayout)
    Dim outStart As Long: outStart = L.instrGenRow + 2
    With ws.Range(ws.Cells(outStart, 1), ws.Cells(outStart + 900, 40))
        .ClearContents: .Borders.LineStyle = xlNone: .Interior.Pattern = xlNone
        .Font.Bold = False: .Font.Italic = False: .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral: .VerticalAlignment = xlCenter
    End With
End Sub
Private Sub FJS_DeleteChartIfExists(ByVal ws As Worksheet)
    Dim i As Long: On Error Resume Next
    For i = ws.ChartObjects.Count To 1 Step -1: ws.ChartObjects(i).Delete: Next i
    On Error GoTo 0
End Sub
Private Sub FJS_ApplyWorkstationDropdownMatrix(ByVal ws As Worksheet, ByVal m As Long, ByVal fr As Long, ByVal lR As Long, ByVal fC As Long, ByVal lC As Long)
    Dim s As String, k As Long: s = ""
    For k = 1 To m: s = s & "M" & k & IIf(k < m, ",", ""): Next k
    With ws.Range(ws.Cells(fr, fC), ws.Cells(lR, lC)).Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=s
        .IgnoreBlank = True: .InCellDropdown = True
    End With
End Sub
Private Sub FJS_ApplyJobDropdownMatrix(ByVal ws As Worksheet, ByVal fr As Long, ByVal lR As Long, ByVal fC As Long, ByVal lC As Long)
    With ws.Range(ws.Cells(fr, fC), ws.Cells(lR, lC)).Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=FJS_JobList"
        .IgnoreBlank = True: .InCellDropdown = True
    End With
End Sub
Private Sub FJS_SetJobListNamedRange(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error Resume Next: ws.Parent.Names("FJS_JobList").Delete: On Error GoTo 0
    ws.Parent.Names.Add name:="FJS_JobList", _
        RefersTo:=ws.Range(ws.Cells(firstRow, P_COL_JOB), ws.Cells(lastRow, P_COL_JOB))
End Sub
Private Function FJS_IsValidWS(ByVal s As String, ByVal m As Long) As Boolean
    s = UCase$(Trim$(s)): If Len(s) < 2 Then Exit Function: If Left$(s, 1) <> "M" Then Exit Function
    FJS_IsValidWS = (val(Mid$(s, 2)) >= 1 And val(Mid$(s, 2)) <= m)
End Function
Private Function FJS_WSIndex(ByVal s As String) As Long
    s = UCase$(Trim$(s)): FJS_WSIndex = val(Mid$(s, 2)): If FJS_WSIndex < 1 Then FJS_WSIndex = 1
End Function
Private Function FJS_OneHot(ByVal total As Long, ByVal idx As Long, ByVal v As Double) As Variant
    Dim a() As Double, t As Long: ReDim a(1 To total)
    For t = 1 To total: a(t) = 0#: Next t
    a(idx) = v: FJS_OneHot = a
End Function
Private Function FJS_ColorForJobName(ByVal jn As String, ByVal jc As Object) As Long
    Dim k As String: k = UCase$(Trim$(jn))
    If Not jc.Exists(k) Then jc.Add k, jc.Count + 1
    FJS_ColorForJobName = FJS_BasePalette(jc(k))
End Function
Private Function FJS_BasePalette(ByVal i As Long) As Long
    Dim p As Variant
    p = Array(RGB(52, 96, 174), RGB(46, 204, 113), RGB(155, 89, 182), RGB(241, 196, 15), _
              RGB(231, 76, 60), RGB(26, 188, 156), RGB(127, 140, 141), RGB(52, 152, 219), _
              RGB(39, 174, 96), RGB(243, 156, 18))
    FJS_BasePalette = p((i - 1) Mod (UBound(p) + 1))
End Function
Private Function FJS_LightTone(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = (c And &HFF): g = (c \ &H100) And &HFF: b = (c \ &H10000) And &HFF
    FJS_LightTone = RGB(IIf(r + 80 < 255, r + 80, 255), IIf(g + 80 < 255, g + 80, 255), IIf(b + 80 < 255, b + 80, 255))
End Function
Private Function FJS_DarkTone(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = (c And &HFF): g = (c \ &H100) And &HFF: b = (c \ &H10000) And &HFF
    FJS_DarkTone = RGB(IIf(r > 40, r - 40, 0), IIf(g > 40, g - 40, 0), IIf(b > 40, b - 40, 0))
End Function
Private Sub FJS_ConfigurarEjeTiempo(ByVal ch As Chart, ByVal maxTime As Double)
    Dim eje As Axis: Set eje = ch.Axes(xlValue)
    If maxTime < 0.000001 Then maxTime = 1
    Dim majorU As Double: majorU = FJS_NiceMajorUnit(maxTime, 30)
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
Private Function FJS_NiceMajorUnit(ByVal mx As Double, ByVal tt As Long) As Double
    If mx <= 0 Then FJS_NiceMajorUnit = 1: Exit Function
    Dim rs As Double: rs = mx / tt: If rs < 1 Then rs = 1
    Dim p10 As Double: p10 = 10 ^ Int(Log(rs) / Log(10))
    Dim fr As Double: fr = rs / p10
    Select Case fr
        Case Is <= 1: FJS_NiceMajorUnit = 1 * p10
        Case Is <= 2: FJS_NiceMajorUnit = 2 * p10
        Case Is <= 5: FJS_NiceMajorUnit = 5 * p10
        Case Else: FJS_NiceMajorUnit = 10 * p10
    End Select
End Function
