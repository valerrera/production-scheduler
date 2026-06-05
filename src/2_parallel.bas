Option Explicit

Private Const SHEET_NAME As String = "ParallelMachine"

' =========================
' CELDAS FIJAS
' =========================
Private Const CELL_N As String = "C12"          ' # jobs
Private Const CELL_M As String = "C13"          ' # máquinas
Private Const CELL_PM_READY As String = "Z2"    ' flag interno

' =========================
' ZONA DINÁMICA
' =========================
Private Const DYNAMIC_TOP_ROW As Long = 15
Private Const INPUT_CLEAR_ROWS As Long = 900

' =========================
' TABLA DISTRIBUCIÓN
' =========================
Private Const MACH_HEADER_ROW As Long = 17
Private Const MACH_FIRST_ROW As Long = 18
Private Const COL_MACH As Long = 2              ' B
Private Const COL_JOBS As Long = 3              ' C
Private Const MACH_MAX As Long = 50

' =========================
' PARÁMETROS POR JOB
' =========================
Private Const GAP_AFTER_DIST_TABLE As Long = 3
Private Const GAP_AFTER_PARAM_TABLE As Long = 3
Private Const BLOCK_GAP_ROWS As Long = 1

Private Const P_COL_JOB As Long = 2             ' B
Private Const P_COL_R As Long = 3               ' C
Private Const P_COL_P As Long = 4               ' D
Private Const P_COL_S As Long = 5               ' E
Private Const P_COL_D As Long = 6               ' F
Private Const P_COL_W As Long = 7               ' G
Private Const P_NUM_COLS As Long = 6

' =========================
' TABLA 3.1 rmaq
' =========================
Private Const RM_COL_MACH As Long = 2           ' B
Private Const RM_COL_RMAQ As Long = 3           ' C

' =========================
' DECISIÓN
' =========================
Private Const D_COL_SEQ As Long = 2             ' B
Private Const D_COL_JOB As Long = 3             ' C

' =========================
' CHART
' =========================
Private Const CHART_NAME_TIMELINE As String = "chLineaTiempo_PM"

' =========================
' CACHE EN MEMORIA
' =========================
Private pmLoaded As Boolean
Private pmN As Long
Private pmM As Long

Private pmJobName() As String
Private pmMachName() As String
Private pmMachIdx() As Long
Private pmSeq() As Long

Private pmR() As Double
Private pmP() As Double
Private pmS() As Double
Private pmD() As Double
Private pmW() As Double
Private pmRmaq() As Double

' =========================================================
' ETAPA 1: SOLO DISTRIBUCIÓN
' =========================================================
Public Sub RedibujarDistribucion_PM(ByVal ws As Worksheet)
    On Error GoTo ErrH

    Dim n As Long, m As Long
    If Not PM_ReadNM(ws, n, m) Then Exit Sub

    ws.Unprotect

    pmLoaded = False
    pmN = 0
    pmM = 0
    ws.Range(CELL_PM_READY).Value = ""

    PM_ClearDynamicZone ws
    DeleteChartIfExists_PM ws, CHART_NAME_TIMELINE

    With ws.Cells(MACH_HEADER_ROW - 1, COL_MACH)
        .Value = "2. Distribución de jobs por máquina (escribe cuántos jobs se asignarán a cada máquina)."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    PM_DrawDistributionTable ws, m

Salir:
    ws.Range(CELL_N).Locked = False
    ws.Range(CELL_M).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub

ErrH:
    MsgBox "Error en RedibujarDistribucion_PM: " & Err.Description, vbExclamation
    Resume Salir
End Sub
Public Sub PM_IntentarMostrarResto()
    On Error GoTo ErrH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim n As Long, m As Long
    If Not PM_ReadNM(ws, n, m) Then Exit Sub

    ' Siempre resetear lo de abajo cuando cambie la distribución
    PM_OcultarRestoInputs ws

    Dim jobsPerMac() As Long
    Dim warn As String

    If Not PM_ReadAndValidateJobsPerMachine(ws, n, m, jobsPerMac, warn) Then
        Exit Sub
    End If

    PM_MostrarRestoInputs ws, n, m, jobsPerMac
    Exit Sub

ErrH:
    MsgBox "Error en PM_IntentarMostrarResto: " & Err.Description, vbExclamation
End Sub

Private Sub PM_OcultarRestoInputs(ByVal ws As Worksheet)
    On Error GoTo ErrH

    Dim n As Long, m As Long
    If Not PM_ReadNM(ws, n, m) Then Exit Sub

    ws.Unprotect

    Dim pTitleRow As Long
    pTitleRow = MACH_FIRST_ROW + m - 1 + GAP_AFTER_DIST_TABLE

    With ws.Range(ws.Cells(pTitleRow, 1), ws.Cells(pTitleRow + 700, 35))
        On Error Resume Next
        .UnMerge
        On Error GoTo ErrH

        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Pattern = xlNone
        .Font.Bold = False
        .Font.Italic = False
        .Font.Underline = xlUnderlineStyleNone
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Validation.Delete
    End With

    DeleteChartIfExists_PM ws, CHART_NAME_TIMELINE

    ws.Range(CELL_PM_READY).Value = ""
    pmLoaded = False
    pmN = 0
    pmM = 0

Salir:
    ws.Range(CELL_N).Locked = False
    ws.Range(CELL_M).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub

ErrH:
    MsgBox "Error en PM_OcultarRestoInputs: " & Err.Description, vbExclamation
    Resume Salir
End Sub

Private Sub PM_MostrarRestoInputs( _
    ByVal ws As Worksheet, _
    ByVal n As Long, _
    ByVal m As Long, _
    ByRef jobsPerMac() As Long)

    On Error GoTo ErrH

    ws.Unprotect

    Dim pTitleRow As Long, pHeaderRow As Long, pFirstDataRow As Long, pLastDataRow As Long
    Dim rmaqTitleRow As Long, rmaqHeaderRow As Long, rmaqFirstRow As Long, rmaqLastRow As Long
    Dim helpRow As Long, firstBlockHeaderRow As Long, instr4Row As Long, instr5Row As Long

    PM_GetLayoutRows_FromDistribution_WithRmaq n, m, jobsPerMac, _
        pTitleRow, pHeaderRow, pFirstDataRow, pLastDataRow, _
        rmaqTitleRow, rmaqHeaderRow, rmaqFirstRow, rmaqLastRow, _
        helpRow, firstBlockHeaderRow, instr4Row, instr5Row

    With ws.Range(ws.Cells(pTitleRow, 1), ws.Cells(pTitleRow + 700, 35))
        On Error Resume Next
        .UnMerge
        On Error GoTo ErrH

        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Pattern = xlNone
        .Font.Bold = False
        .Font.Italic = False
        .Font.Underline = xlUnderlineStyleNone
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Validation.Delete
    End With

    With ws.Cells(pTitleRow, P_COL_JOB)
        .Value = "3. Parámetros por job (llena r (disponibilidad), p (tiempo de procesamiento), s (tiempo de alistamiento), d (fecha de entrega), w (peso))."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    PM_DrawParamTable ws, n, pHeaderRow, pFirstDataRow, pLastDataRow
    PM_SetJobListNamedRange ws, pFirstDataRow, pLastDataRow

    With ws.Cells(rmaqTitleRow, RM_COL_MACH)
        .Value = "3.1. Escribe la fecha de disponibilidad de la(s) máquina(s)."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    PM_DrawRmaqTable ws, m, rmaqHeaderRow, rmaqFirstRow, rmaqLastRow

    With ws.Cells(helpRow, D_COL_SEQ)
        .Value = "4. Decisión (en cada máquina escribe la secuencia y selecciona el job en el dropdown)."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    PM_DrawDecisionBlocks ws, m, jobsPerMac, firstBlockHeaderRow

    With ws.Cells(instr4Row, D_COL_SEQ)
        .Value = "5. Presione el botón Cargar datos para validar la información ingresada."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False
    ws.Range(CELL_M).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub

ErrH:
    MsgBox "Error en PM_MostrarRestoInputs: " & Err.Description, vbExclamation
    Resume Salir
End Sub

' =========================================================
' BOTÓN 1: CARGAR DATOS
' =========================================================
Public Sub PM_CargarDatos()
    On Error GoTo ErrH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim n As Long, m As Long
    If Not PM_ReadNM(ws, n, m) Then
        MsgBox "Verifica que # jobs y # máquinas sean enteros positivos.", vbExclamation
        Exit Sub
    End If

    ws.Unprotect
    ws.Range(CELL_PM_READY).Value = ""

    Dim jobsPerMac() As Long
    Dim warn As String

    If Not PM_ReadAndValidateJobsPerMachine(ws, n, m, jobsPerMac, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos - Parallel Machine"
        GoTo Salir
    End If

    Dim pTitleRow As Long, pHeaderRow As Long, pFirstDataRow As Long, pLastDataRow As Long
    Dim rmaqTitleRow As Long, rmaqHeaderRow As Long, rmaqFirstRow As Long, rmaqLastRow As Long
    Dim helpRow As Long, firstBlockHeaderRow As Long, instr4Row As Long, instr5Row As Long

    PM_GetLayoutRows_FromDistribution_WithRmaq n, m, jobsPerMac, _
        pTitleRow, pHeaderRow, pFirstDataRow, pLastDataRow, _
        rmaqTitleRow, rmaqHeaderRow, rmaqFirstRow, rmaqLastRow, _
        helpRow, firstBlockHeaderRow, instr4Row, instr5Row

    With ws.Cells(instr5Row, D_COL_SEQ)
        .ClearContents
        .Font.Bold = False
        .Font.Italic = False
    End With

    If Not ValidateInputs_PM(ws, n, m, pFirstDataRow, firstBlockHeaderRow, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos - Parallel Machine"
        GoTo Salir
    End If

    PM_LoadInputsToCache ws, n, m, pFirstDataRow, firstBlockHeaderRow

    ws.Range(CELL_PM_READY).Value = "OK"
    ws.Range(CELL_PM_READY).Locked = True

    With ws.Cells(instr5Row, D_COL_SEQ)
        .Value = "6. Datos válidos. Presione el botón Generar outputs."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False
    ws.Range(CELL_M).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub

ErrH:
    MsgBox "Error en PM_CargarDatos: " & Err.Description, vbExclamation
    Resume Salir
End Sub

Private Function ValidateInputs_PM( _
    ByVal ws As Worksheet, _
    ByVal n As Long, _
    ByVal m As Long, _
    ByVal pFirstDataRow As Long, _
    ByVal firstBlockHeaderRow As Long, _
    ByRef warn As String) As Boolean

    ValidateInputs_PM = False
    warn = ""

    Dim jobsPerMac() As Long
    If Not PM_ReadAndValidateJobsPerMachine(ws, n, m, jobsPerMac, warn) Then Exit Function

    Dim dictJobs As Object
    Set dictJobs = CreateObject("Scripting.Dictionary")

    Dim i As Long, rr As Long, jb As String
    For i = 1 To n
        rr = pFirstDataRow + i - 1

        jb = UCase$(Trim$(CStr(ws.Cells(rr, P_COL_JOB).Value)))
        If Len(jb) = 0 Then
            warn = "Tabla 3: falta Job en fila " & rr & "."
            Exit Function
        End If
        If dictJobs.Exists(jb) Then
            warn = "Tabla 3: Job repetido (" & jb & ")."
            Exit Function
        End If
        dictJobs.Add jb, True

        If Not IsNumeric(ws.Cells(rr, P_COL_R).Value) Then
            warn = "Tabla 3: r debe ser numérico para " & jb & "."
            Exit Function
        End If
        If Not IsNumeric(ws.Cells(rr, P_COL_P).Value) Then
            warn = "Tabla 3: p debe ser numérico para " & jb & "."
            Exit Function
        End If
        If Not IsNumeric(ws.Cells(rr, P_COL_S).Value) Then
            warn = "Tabla 3: s debe ser numérico para " & jb & "."
            Exit Function
        End If
        If Not IsNumeric(ws.Cells(rr, P_COL_D).Value) Then
            warn = "Tabla 3: d debe ser numérico para " & jb & "."
            Exit Function
        End If
        If Not IsNumeric(ws.Cells(rr, P_COL_W).Value) Then
            warn = "Tabla 3: w debe ser numérico para " & jb & "."
            Exit Function
        End If

        If NzD_PM(ws.Cells(rr, P_COL_R).Value) < 0 Then
            warn = "Tabla 3: r no puede ser negativo para " & jb & "."
            Exit Function
        End If
        If NzD_PM(ws.Cells(rr, P_COL_S).Value) < 0 Then
            warn = "Tabla 3: s no puede ser negativo para " & jb & "."
            Exit Function
        End If
        If NzD_PM(ws.Cells(rr, P_COL_P).Value) <= 0 Then
            warn = "Tabla 3: p debe ser > 0 para " & jb & "."
            Exit Function
        End If
        If NzD_PM(ws.Cells(rr, P_COL_D).Value) <= 0 Then
            warn = "Tabla 3: d debe ser > 0 para " & jb & "."
            Exit Function
        End If
        If NzD_PM(ws.Cells(rr, P_COL_W).Value) <= 0 Then
            warn = "Tabla 3: w debe ser > 0 para " & jb & "."
            Exit Function
        End If
    Next i

    Dim rmaqTitleRow As Long, rmaqHeaderRow As Long, rmaqFirstRow As Long, rmaqLastRow As Long
    PM_GetRmaqRows_FromDistribution n, m, rmaqTitleRow, rmaqHeaderRow, rmaqFirstRow, rmaqLastRow

    Dim rk As Long, rv As Variant
    For rk = rmaqFirstRow To rmaqLastRow
        rv = ws.Cells(rk, RM_COL_RMAQ).Value

        If Len(Trim$(CStr(rv))) = 0 Then
            warn = "Tabla 3.1: falta rmaq para la máquina en fila " & rk & "."
            Exit Function
        End If
        If Not IsNumeric(rv) Then
            warn = "Tabla 3.1: rmaq debe ser numérico en fila " & rk & "."
            Exit Function
        End If
        If CDbl(rv) < 0 Then
            warn = "Tabla 3.1: rmaq no puede ser negativo en fila " & rk & "."
            Exit Function
        End If
    Next rk

    Dim seenJobs As Object
    Set seenJobs = CreateObject("Scripting.Dictionary")

    Dim k As Long, t As Long
    Dim headerRow As Long, firstDataRow As Long, rowsOnMac As Long
    Dim vSeq As Variant, vJob As String
    Dim dictSeq As Object

    headerRow = firstBlockHeaderRow

    For k = 1 To m
        rowsOnMac = jobsPerMac(k)
        firstDataRow = headerRow + 1

        Set dictSeq = CreateObject("Scripting.Dictionary")

        For t = 0 To rowsOnMac - 1
            rr = firstDataRow + t

            vSeq = ws.Cells(rr, D_COL_SEQ).Value
            vJob = UCase$(Trim$(CStr(ws.Cells(rr, D_COL_JOB).Value)))

            If Len(Trim$(CStr(vSeq))) = 0 Then
                warn = "Bloque M" & k & ": falta Secuencia en fila " & rr & "."
                Exit Function
            End If
            If Not IsNumeric(vSeq) Then
                warn = "Bloque M" & k & ": Secuencia debe ser numérica en fila " & rr & "."
                Exit Function
            End If
            If CLng(vSeq) <> CDbl(vSeq) Or CLng(vSeq) <= 0 Then
                warn = "Bloque M" & k & ": Secuencia debe ser entero positivo en fila " & rr & "."
                Exit Function
            End If
            If dictSeq.Exists(CStr(CLng(vSeq))) Then
                warn = "Bloque M" & k & ": Secuencia repetida (" & CLng(vSeq) & ")."
                Exit Function
            End If
            dictSeq.Add CStr(CLng(vSeq)), True

            If Len(vJob) = 0 Then
                warn = "Bloque M" & k & ": falta Job en fila " & rr & "."
                Exit Function
            End If
            If Not dictJobs.Exists(vJob) Then
                warn = "Bloque M" & k & ": el Job '" & vJob & "' no existe en la tabla 3."
                Exit Function
            End If
            If seenJobs.Exists(vJob) Then
                warn = "Bloque 4: el Job '" & vJob & "' está repetido."
                Exit Function
            End If
            seenJobs.Add vJob, True
        Next t

        Dim need As Long
        For need = 1 To rowsOnMac
            If Not dictSeq.Exists(CStr(need)) Then
                warn = "Bloque M" & k & ": la secuencia debe ser 1.." & rowsOnMac & " sin saltos. Falta: " & need & "."
                Exit Function
            End If
        Next need

        headerRow = PM_NextBlockHeaderRow(headerRow, rowsOnMac)
    Next k

    If seenJobs.Count <> n Then
        warn = "Cada Job debe aparecer exactamente una vez en la decisión. Faltan jobs por asignar."
        Exit Function
    End If

    ValidateInputs_PM = True
End Function

' =========================================================
' BOTÓN 2: GENERAR OUTPUTS
' =========================================================
Public Sub PM_GenerarOutputs()
    On Error GoTo ErrH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    If UCase$(Trim$(CStr(ws.Range(CELL_PM_READY).Value))) <> "OK" Then
        MsgBox "Primero presione 'Cargar datos' y asegúrese de que los datos sean válidos.", vbExclamation
        Exit Sub
    End If

    Dim n As Long, m As Long
    If Not PM_ReadNM(ws, n, m) Then Exit Sub

    If Not pmLoaded Or pmN <> n Or pmM <> m Then
        MsgBox "Los datos no están cargados en memoria. Presione 'Cargar datos' nuevamente.", vbExclamation
        Exit Sub
    End If

    ws.Unprotect

    Dim jobsPerMac() As Long
    Dim warn As String
    If Not PM_ReadAndValidateJobsPerMachine(ws, n, m, jobsPerMac, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos - Parallel Machine"
        GoTo Salir
    End If

    Dim pTitleRow As Long, pHeaderRow As Long, pFirstDataRow As Long, pLastDataRow As Long
    Dim rmaqTitleRow As Long, rmaqHeaderRow As Long, rmaqFirstRow As Long, rmaqLastRow As Long
    Dim helpRow As Long, firstBlockHeaderRow As Long, instr4Row As Long, instr5Row As Long

    PM_GetLayoutRows_FromDistribution_WithRmaq n, m, jobsPerMac, _
        pTitleRow, pHeaderRow, pFirstDataRow, pLastDataRow, _
        rmaqTitleRow, rmaqHeaderRow, rmaqFirstRow, rmaqLastRow, _
        helpRow, firstBlockHeaderRow, instr4Row, instr5Row

    Dim outLineRow As Long: outLineRow = instr5Row + 1
    Dim outTitleRow As Long: outTitleRow = outLineRow + 2
    Dim ganttTopRow As Long: ganttTopRow = outTitleRow + 2
    Dim indTopRow As Long: indTopRow = ganttTopRow + 14 + 5

    DeleteChartIfExists_PM ws, CHART_NAME_TIMELINE

    Dim lastClearRow As Long: lastClearRow = outLineRow + 500
    With ws.Range(ws.Cells(outLineRow, 1), ws.Cells(lastClearRow, 35))
        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Pattern = xlNone
        .Font.Bold = False
        .Font.Italic = False
        .Font.Underline = xlUnderlineStyleNone
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(outLineRow, 1), ws.Cells(outLineRow, 35))
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .Weight = xlThin
        End With
    End With

    With ws.Cells(outTitleRow, D_COL_SEQ)
        .Value = "ZONA DE OUTPUTS"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    Dim chObj As ChartObject
    Set chObj = ws.ChartObjects.Add( _
        Left:=ws.Cells(ganttTopRow, 2).Left, _
        top:=ws.Cells(ganttTopRow, 2).top, _
        Width:=1100, Height:=220 + 24 * m)
    chObj.name = CHART_NAME_TIMELINE

    With chObj.Chart
        .ChartType = xlBarStacked
        .HasTitle = True
        .ChartTitle.text = "Gantt Parallel Machine"
        .HasLegend = False
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
    End With

    PM_DrawIndicatorsStructure ws, indTopRow

    Dim jmFirstCol As Long: jmFirstCol = 8   ' H
    Dim jmNumCols As Long: jmNumCols = 9
    Dim jmTitleRow As Long: jmTitleRow = indTopRow
    Dim jmHeaderRow As Long: jmHeaderRow = jmTitleRow + 1
    Dim jmFirstDataRow As Long: jmFirstDataRow = jmHeaderRow + 1
    Dim jmLastRow As Long: jmLastRow = jmFirstDataRow + n - 1

    With ws.Cells(jmTitleRow, jmFirstCol)
        .Value = "Indicadores por job"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Locked = True
    End With

    Dim jmHeaders As Variant
    jmHeaders = Array("Job", "Máquina", "rmaq", "Inicio", "Cj", "Flow (Cj-rj)", "L (Cj-dj)", "T=max(L,0)", "w*T")

    Dim j As Long
    For j = 0 To UBound(jmHeaders)
        With ws.Cells(jmHeaderRow, jmFirstCol + j)
            .Value = jmHeaders(j)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(230, 230, 230)
            .Locked = True
        End With
    Next j

    With ws.Range(ws.Cells(jmFirstDataRow, jmFirstCol), ws.Cells(jmLastRow, jmFirstCol + jmNumCols - 1))
        .ClearContents
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
        .Interior.Pattern = xlNone
        .Locked = True
    End With

    With ws.Range(ws.Cells(jmHeaderRow, jmFirstCol), ws.Cells(jmLastRow, jmFirstCol + jmNumCols - 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
    End With

    Dim makespan As Double
    PM_CalcAndWriteOutputs ws, n, m, indTopRow, jmFirstCol, jmFirstDataRow, makespan
    PM_BuildGanttFromCache chObj.Chart, n, m, makespan

Salir:
    ws.Range(CELL_N).Locked = False
    ws.Range(CELL_M).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub

ErrH:
    MsgBox "Error en PM_GenerarOutputs: " & Err.Description, vbExclamation
    Resume Salir
End Sub

' =========================================================
' CARGAR INPUTS A CACHE
' =========================================================
Private Sub PM_LoadInputsToCache( _
    ByVal ws As Worksheet, _
    ByVal n As Long, _
    ByVal m As Long, _
    ByVal pFirstDataRow As Long, _
    ByVal firstBlockHeaderRow As Long)

    pmLoaded = False
    pmN = n
    pmM = m

    Dim jobsPerMac() As Long
    Dim warn As String
    If Not PM_ReadAndValidateJobsPerMachine(ws, n, m, jobsPerMac, warn) Then Exit Sub

    Dim dictR As Object, dictP As Object, dictS As Object, dictD As Object, dictW As Object
    Set dictR = CreateObject("Scripting.Dictionary")
    Set dictP = CreateObject("Scripting.Dictionary")
    Set dictS = CreateObject("Scripting.Dictionary")
    Set dictD = CreateObject("Scripting.Dictionary")
    Set dictW = CreateObject("Scripting.Dictionary")

    Dim i As Long, rr As Long, jb As String
    For i = 1 To n
        rr = pFirstDataRow + i - 1
        jb = UCase$(Trim$(CStr(ws.Cells(rr, P_COL_JOB).Value)))

        dictR(jb) = NzD_PM(ws.Cells(rr, P_COL_R).Value)
        dictP(jb) = NzD_PM(ws.Cells(rr, P_COL_P).Value)
        dictS(jb) = NzD_PM(ws.Cells(rr, P_COL_S).Value)
        dictD(jb) = NzD_PM(ws.Cells(rr, P_COL_D).Value)
        dictW(jb) = NzD_PM(ws.Cells(rr, P_COL_W).Value)
    Next i

    ReDim pmJobName(1 To n)
    ReDim pmMachName(1 To n)
    ReDim pmMachIdx(1 To n)
    ReDim pmSeq(1 To n)
    ReDim pmR(1 To n)
    ReDim pmP(1 To n)
    ReDim pmS(1 To n)
    ReDim pmD(1 To n)
    ReDim pmW(1 To n)
    ReDim pmRmaq(1 To m)

    Dim rmaqTitleRow As Long, rmaqHeaderRow As Long, rmaqFirstRow As Long, rmaqLastRow As Long
    PM_GetRmaqRows_FromDistribution n, m, rmaqTitleRow, rmaqHeaderRow, rmaqFirstRow, rmaqLastRow

    Dim mk As Long
    For mk = 1 To m
        pmRmaq(mk) = NzD_PM(ws.Cells(rmaqFirstRow + mk - 1, RM_COL_RMAQ).Value)
    Next mk

    Dim k As Long, t As Long, pos As Long
    Dim headerRow As Long, firstDataRow As Long, rowsOnMac As Long

    pos = 0
    headerRow = firstBlockHeaderRow

    For k = 1 To m
        rowsOnMac = jobsPerMac(k)
        firstDataRow = headerRow + 1

        For t = 0 To rowsOnMac - 1
            rr = firstDataRow + t
            pos = pos + 1

            jb = UCase$(Trim$(CStr(ws.Cells(rr, D_COL_JOB).Value)))

            pmJobName(pos) = jb
            pmMachName(pos) = "M" & k
            pmMachIdx(pos) = k
            pmSeq(pos) = CLng(ws.Cells(rr, D_COL_SEQ).Value)

            pmR(pos) = dictR(jb)
            pmP(pos) = dictP(jb)
            pmS(pos) = dictS(jb)
            pmD(pos) = dictD(jb)
            pmW(pos) = dictW(jb)
        Next t

        headerRow = PM_NextBlockHeaderRow(headerRow, rowsOnMac)
    Next k

    PM_SortByMachineAndSeq pmMachIdx, pmSeq, pmJobName, pmMachName, pmR, pmP, pmS, pmD, pmW, n
    pmLoaded = True
End Sub

' =========================================================
' CÁLCULO + ESCRITURA DE OUTPUTS
' =========================================================
Private Sub PM_CalcAndWriteOutputs( _
    ByVal ws As Worksheet, _
    ByVal n As Long, _
    ByVal m As Long, _
    ByVal indTopRow As Long, _
    ByVal jmFirstCol As Long, _
    ByVal jmFirstDataRow As Long, _
    ByRef makespan As Double)

    Dim startT() As Double, Cj() As Double, flowT() As Double, Lj() As Double, tJ() As Double, wTj() As Double
    ReDim startT(1 To n)
    ReDim Cj(1 To n)
    ReDim flowT(1 To n)
    ReDim Lj(1 To n)
    ReDim tJ(1 To n)
    ReDim wTj(1 To n)

    Dim macClock() As Double
    ReDim macClock(1 To m)

    Dim k As Long
    For k = 1 To m
        macClock(k) = pmRmaq(k)
    Next k

    Dim i As Long, idx As Long
    For i = 1 To n
        idx = pmMachIdx(i)

        startT(i) = WorksheetFunction.Max(macClock(idx), pmR(i))
        macClock(idx) = startT(i) + pmS(i) + pmP(i)
        Cj(i) = macClock(idx)

        flowT(i) = Cj(i) - pmR(i)
        Lj(i) = Cj(i) - pmD(i)
        If Lj(i) > 0# Then
            tJ(i) = Lj(i)
        Else
            tJ(i) = 0#
        End If
        wTj(i) = pmW(i) * tJ(i)
    Next i

    Dim Cmax As Double: Cmax = 0#
    Dim sumFlow As Double, avgFlow As Double
    Dim Lmax As Double: Lmax = -1E+30
    Dim sumT As Double, avgT As Double
    Dim sumWT As Double
    Dim lateCount As Long

    For i = 1 To n
        If Cj(i) > Cmax Then Cmax = Cj(i)
        sumFlow = sumFlow + flowT(i)
        If Lj(i) > Lmax Then Lmax = Lj(i)
        sumT = sumT + tJ(i)
        sumWT = sumWT + wTj(i)
        If tJ(i) > 0.000001 Then lateCount = lateCount + 1
    Next i

    avgFlow = sumFlow / CDbl(n)
    avgT = sumT / CDbl(n)
    makespan = Cmax

    ws.Cells(indTopRow + 1, 5).Value = Cmax
    ws.Cells(indTopRow + 2, 5).Value = avgFlow
    ws.Cells(indTopRow + 3, 5).Value = Lmax
    ws.Cells(indTopRow + 4, 5).Value = avgT
    ws.Cells(indTopRow + 5, 5).Value = sumWT
    ws.Cells(indTopRow + 6, 5).Value = lateCount
    ws.Cells(indTopRow + 7, 5).Value = lateCount / CDbl(n)
    ws.Cells(indTopRow + 8, 5).Value = 1# - lateCount / CDbl(n)

    ws.Cells(indTopRow + 7, 5).NumberFormat = "0%"
    ws.Cells(indTopRow + 8, 5).NumberFormat = "0%"

    For i = 1 To n
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 0).Value = pmJobName(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 1).Value = pmMachName(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 2).Value = pmRmaq(pmMachIdx(i))
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 3).Value = startT(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 4).Value = Cj(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 5).Value = flowT(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 6).Value = Lj(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 7).Value = tJ(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 8).Value = wTj(i)
    Next i
End Sub

' =========================================================
' GANTT
' =========================================================
Private Sub PM_BuildGanttFromCache( _
    ByVal ch As Chart, _
    ByVal n As Long, _
    ByVal m As Long, _
    ByVal makespan As Double)

    Dim cats() As String
    ReDim cats(1 To m)

    Dim i As Long
    For i = 1 To m
        cats(i) = "M" & i
    Next i

    With ch
        .ChartType = xlBarStacked
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.text = "Gantt Parallel Machine"

        On Error Resume Next
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        On Error GoTo 0

        Dim srs As Series
        Dim base() As Double
        ReDim base(1 To m)

        Set srs = .SeriesCollection.NewSeries
        srs.Values = base
        srs.XValues = cats
        srs.Format.Fill.Visible = msoFalse
        srs.Format.Line.Visible = msoFalse

        Dim offByMac() As Double
        ReDim offByMac(1 To m)

        For i = 1 To m
            offByMac(i) = 0
        Next i
        Dim gap As Double
        Dim baseColor As Long, cLight As Long, cDark As Long
        Dim idx As Long, startVal As Double

        For i = 1 To n
            idx = pmMachIdx(i)
            startVal = PM_StartTime_FromCache(i)

            gap = startVal - offByMac(idx)
            If gap > 0.000001 Then
                Set srs = .SeriesCollection.NewSeries
                srs.Values = PM_OneHot(m, idx, gap)
                srs.XValues = cats
                srs.Format.Fill.Visible = msoFalse
                srs.Format.Line.Visible = msoFalse
                offByMac(idx) = offByMac(idx) + gap
            End If

            baseColor = ColorForJob_PM(i)
            cLight = LightTone_PM(baseColor)
            cDark = DarkTone_PM(baseColor)

            If pmS(i) > 0# Then
                Set srs = .SeriesCollection.NewSeries
                srs.Values = PM_OneHot(m, idx, pmS(i))
                srs.XValues = cats
                srs.Format.Fill.ForeColor.RGB = cLight
                srs.Format.Line.Visible = msoFalse
                offByMac(idx) = offByMac(idx) + pmS(i)
            End If

            If pmP(i) > 0# Then
                Set srs = .SeriesCollection.NewSeries
                srs.Values = PM_OneHot(m, idx, pmP(i))
                srs.XValues = cats
                srs.Format.Fill.ForeColor.RGB = cDark
                srs.Format.Line.Visible = msoFalse

                If srs.Points.Count >= idx Then
                    With srs.Points(idx)
                        .HasDataLabel = True
                        .DataLabel.text = pmJobName(i)
                        .DataLabel.Font.Size = 9
                    End With
                End If

                offByMac(idx) = offByMac(idx) + pmP(i)
            End If
        Next i

        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlCategory).Crosses = xlMinimum

        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.text = "Tiempo"
        ConfigurarEjeTiempo_Inteligente_PM ch, makespan
    End With
End Sub

Private Function PM_StartTime_FromCache(ByVal pos As Long) As Double
    Dim idx As Long, i As Long
    Dim macClock() As Double
    ReDim macClock(1 To pmM)

    Dim k As Long
    For k = 1 To pmM
        macClock(k) = pmRmaq(k)
    Next k

    For i = 1 To pos
        idx = pmMachIdx(i)

        If i = pos Then
            PM_StartTime_FromCache = WorksheetFunction.Max(macClock(idx), pmR(i))
            Exit Function
        End If

        macClock(idx) = WorksheetFunction.Max(macClock(idx), pmR(i)) + pmS(i) + pmP(i)
    Next i
End Function

' =========================================================
' LAYOUT
' =========================================================
Private Sub PM_GetLayoutRows_FromDistribution_WithRmaq( _
    ByVal n As Long, _
    ByVal m As Long, _
    ByRef jobsPerMac() As Long, _
    ByRef pTitleRow As Long, _
    ByRef pHeaderRow As Long, _
    ByRef pFirstDataRow As Long, _
    ByRef pLastDataRow As Long, _
    ByRef rmaqTitleRow As Long, _
    ByRef rmaqHeaderRow As Long, _
    ByRef rmaqFirstRow As Long, _
    ByRef rmaqLastRow As Long, _
    ByRef helpRow As Long, _
    ByRef firstBlockHeaderRow As Long, _
    ByRef instr4Row As Long, _
    ByRef instr5Row As Long)

    Dim lastDistRow As Long
    lastDistRow = MACH_FIRST_ROW + m - 1

    pTitleRow = lastDistRow + GAP_AFTER_DIST_TABLE
    pHeaderRow = pTitleRow + 1
    pFirstDataRow = pHeaderRow + 1
    pLastDataRow = pFirstDataRow + n - 1

    rmaqTitleRow = pLastDataRow + 3
    rmaqHeaderRow = rmaqTitleRow + 1
    rmaqFirstRow = rmaqHeaderRow + 1
    rmaqLastRow = rmaqFirstRow + m - 1

    helpRow = rmaqLastRow + GAP_AFTER_PARAM_TABLE
    firstBlockHeaderRow = helpRow + 3

    Dim k As Long, headerRow As Long
    headerRow = firstBlockHeaderRow

    For k = 1 To m
        headerRow = PM_NextBlockHeaderRow(headerRow, jobsPerMac(k))
    Next k

    instr4Row = headerRow
    instr5Row = instr4Row + 2
End Sub

Private Sub PM_GetRmaqRows_FromDistribution( _
    ByVal n As Long, _
    ByVal m As Long, _
    ByRef rmaqTitleRow As Long, _
    ByRef rmaqHeaderRow As Long, _
    ByRef rmaqFirstRow As Long, _
    ByRef rmaqLastRow As Long)

    Dim lastDistRow As Long
    lastDistRow = MACH_FIRST_ROW + m - 1

    Dim pTitleRow As Long, pHeaderRow As Long, pFirstDataRow As Long, pLastDataRow As Long
    pTitleRow = lastDistRow + GAP_AFTER_DIST_TABLE
    pHeaderRow = pTitleRow + 1
    pFirstDataRow = pHeaderRow + 1
    pLastDataRow = pFirstDataRow + n - 1

    rmaqTitleRow = pLastDataRow + 3
    rmaqHeaderRow = rmaqTitleRow + 1
    rmaqFirstRow = rmaqHeaderRow + 1
    rmaqLastRow = rmaqFirstRow + m - 1
End Sub

' =========================================================
' DIBUJO
' =========================================================
Private Sub PM_ClearDynamicZone(ByVal ws As Worksheet)
    With ws.Range(ws.Cells(DYNAMIC_TOP_ROW, 1), ws.Cells(DYNAMIC_TOP_ROW + INPUT_CLEAR_ROWS, 35))
        On Error Resume Next
        .UnMerge
        On Error GoTo 0

        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Pattern = xlNone
        .Font.Bold = False
        .Font.Italic = False
        .Font.Underline = xlUnderlineStyleNone
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Validation.Delete
    End With
End Sub

Private Sub PM_DrawDistributionTable(ByVal ws As Worksheet, ByVal m As Long)
    Dim maxLast As Long
    maxLast = MACH_FIRST_ROW + MACH_MAX - 1

    With ws.Range(ws.Cells(MACH_HEADER_ROW, COL_MACH), ws.Cells(maxLast, COL_JOBS))
        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Pattern = xlNone
        On Error Resume Next
        .UnMerge
        On Error GoTo 0
    End With

    ws.Cells(MACH_HEADER_ROW, COL_MACH).Value = "Máquina"
    ws.Cells(MACH_HEADER_ROW, COL_JOBS).Value = "#Jobs"

    With ws.Range(ws.Cells(MACH_HEADER_ROW, COL_MACH), ws.Cells(MACH_HEADER_ROW, COL_JOBS))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = True
    End With

    Dim i As Long
    For i = 1 To m
        ws.Cells(MACH_FIRST_ROW + i - 1, COL_MACH).Value = "M" & i
    Next i

    With ws.Range(ws.Cells(MACH_FIRST_ROW, COL_MACH), ws.Cells(MACH_FIRST_ROW + m - 1, COL_JOBS))
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(160, 160, 160)
        .Borders.Weight = xlThin
    End With

    ws.Range(ws.Cells(MACH_FIRST_ROW, COL_MACH), ws.Cells(MACH_FIRST_ROW + m - 1, COL_MACH)).Locked = True
    ws.Range(ws.Cells(MACH_FIRST_ROW, COL_JOBS), ws.Cells(MACH_FIRST_ROW + m - 1, COL_JOBS)).Locked = False
End Sub

Private Sub PM_DrawParamTable( _
    ByVal ws As Worksheet, _
    ByVal n As Long, _
    ByVal pHeaderRow As Long, _
    ByVal pFirstDataRow As Long, _
    ByVal pLastDataRow As Long)

    Dim lastCol As Long
    lastCol = P_COL_JOB + P_NUM_COLS - 1

    ws.Cells(pHeaderRow, P_COL_JOB + 0).Value = "Job"
    ws.Cells(pHeaderRow, P_COL_R).Value = "r"
    ws.Cells(pHeaderRow, P_COL_P).Value = "p"
    ws.Cells(pHeaderRow, P_COL_S).Value = "s"
    ws.Cells(pHeaderRow, P_COL_D).Value = "d"
    ws.Cells(pHeaderRow, P_COL_W).Value = "w"

    With ws.Range(ws.Cells(pHeaderRow, P_COL_JOB), ws.Cells(pHeaderRow, lastCol))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = True
    End With

    With ws.Range(ws.Cells(pFirstDataRow, P_COL_JOB), ws.Cells(pLastDataRow, lastCol))
        .ClearContents
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .Interior.Pattern = xlNone
        .NumberFormat = "General"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(160, 160, 160)
        .Borders.Weight = xlThin
    End With

    Dim i As Long
    For i = 1 To n
        ws.Cells(pFirstDataRow + i - 1, P_COL_JOB).Value = "J" & i
        ws.Cells(pFirstDataRow + i - 1, P_COL_W).Value = 1
    Next i

    ws.Range(ws.Cells(pFirstDataRow, P_COL_JOB), ws.Cells(pLastDataRow, P_COL_JOB)).Locked = True
    ws.Range(ws.Cells(pFirstDataRow, P_COL_R), ws.Cells(pLastDataRow, P_COL_W)).Locked = False
End Sub

Private Sub PM_DrawRmaqTable( _
    ByVal ws As Worksheet, _
    ByVal m As Long, _
    ByVal headerRow As Long, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long)

    ws.Cells(headerRow, RM_COL_MACH).Value = "Máquina"
    ws.Cells(headerRow, RM_COL_RMAQ).Value = "rmaq"

    With ws.Range(ws.Cells(headerRow, RM_COL_MACH), ws.Cells(headerRow, RM_COL_RMAQ))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = True
    End With

    Dim i As Long
    For i = 1 To m
        ws.Cells(firstRow + i - 1, RM_COL_MACH).Value = "M" & i
    Next i

    With ws.Range(ws.Cells(firstRow, RM_COL_MACH), ws.Cells(lastRow, RM_COL_RMAQ))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(160, 160, 160)
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With

    ws.Range(ws.Cells(firstRow, RM_COL_RMAQ), ws.Cells(lastRow, RM_COL_RMAQ)).ClearContents
    ws.Range(ws.Cells(firstRow, RM_COL_MACH), ws.Cells(lastRow, RM_COL_MACH)).Locked = True
    ws.Range(ws.Cells(firstRow, RM_COL_RMAQ), ws.Cells(lastRow, RM_COL_RMAQ)).Locked = False
End Sub

Private Sub PM_SetJobListNamedRange(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error Resume Next
    ws.Parent.Names("PM_JobList").Delete
    On Error GoTo 0

    Dim rngJobs As Range
    Set rngJobs = ws.Range(ws.Cells(firstRow, P_COL_JOB), ws.Cells(lastRow, P_COL_JOB))

    ws.Parent.Names.Add name:="PM_JobList", RefersTo:=rngJobs
End Sub

Private Sub PM_DrawDecisionBlocks( _
    ByVal ws As Worksheet, _
    ByVal m As Long, _
    ByRef jobsPerMac() As Long, _
    ByVal firstBlockHeaderRow As Long)

    Dim k As Long, headerRow As Long
    headerRow = firstBlockHeaderRow

    For k = 1 To m
        PM_DrawDecisionBlockOneMachine ws, k, headerRow, jobsPerMac(k)
        headerRow = PM_NextBlockHeaderRow(headerRow, jobsPerMac(k))
    Next k
End Sub

Private Sub PM_DrawDecisionBlockOneMachine( _
    ByVal ws As Worksheet, _
    ByVal machineIdx As Long, _
    ByVal headerRow As Long, _
    ByVal nRows As Long)

    Dim titleRow As Long
    titleRow = headerRow - 1

    With ws.Range(ws.Cells(titleRow, D_COL_SEQ), ws.Cells(titleRow, D_COL_JOB))
        On Error Resume Next
        .UnMerge
        On Error GoTo 0
        .Merge
        .Value = "M" & machineIdx
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    ws.Cells(headerRow, D_COL_SEQ).Value = "Secuencia"
    ws.Cells(headerRow, D_COL_JOB).Value = "Job"

    With ws.Range(ws.Cells(headerRow, D_COL_SEQ), ws.Cells(headerRow, D_COL_JOB))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = True
    End With

    If nRows > 0 Then
        Dim firstDataRow As Long, lastDataRow As Long
        firstDataRow = headerRow + 1
        lastDataRow = firstDataRow + nRows - 1

        With ws.Range(ws.Cells(firstDataRow, D_COL_SEQ), ws.Cells(lastDataRow, D_COL_JOB))
            .ClearContents
            .Font.Bold = False
            .HorizontalAlignment = xlCenter
            .Interior.Pattern = xlNone
            .NumberFormat = "General"
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(160, 160, 160)
            .Borders.Weight = xlThin
            .Locked = False
        End With

        Dim i As Long
        For i = 1 To nRows
            ws.Cells(firstDataRow + i - 1, D_COL_SEQ).Value = i
        Next i

        ws.Range(ws.Cells(firstDataRow, D_COL_SEQ), ws.Cells(lastDataRow, D_COL_SEQ)).Locked = False
        ws.Range(ws.Cells(firstDataRow, D_COL_JOB), ws.Cells(lastDataRow, D_COL_JOB)).Locked = False

        Dim rngDV As Range
        Set rngDV = ws.Range(ws.Cells(firstDataRow, D_COL_JOB), ws.Cells(lastDataRow, D_COL_JOB))

        On Error Resume Next
        rngDV.Validation.Delete
        On Error GoTo 0

        rngDV.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=PM_JobList"
        rngDV.Validation.IgnoreBlank = True
        rngDV.Validation.InCellDropdown = True
        rngDV.Validation.ShowError = True
    End If
End Sub

Private Function PM_NextBlockHeaderRow(ByVal currentHeaderRow As Long, ByVal nRows As Long) As Long
    PM_NextBlockHeaderRow = (currentHeaderRow + 1 + nRows) + BLOCK_GAP_ROWS + 1
End Function

Private Sub PM_DrawIndicatorsStructure(ByVal ws As Worksheet, ByVal indTopRow As Long)
    With ws.Cells(indTopRow, 2)
        .Value = "Indicadores"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Locked = True
    End With

    Dim labels As Variant
    labels = Array( _
        "Makespan (Cmax)", _
        "Tiempo de flujo promedio", _
        "Máximo lateness (Lmax)", _
        "Tardanza media", _
        "Tardanza ponderada total", _
        "# Jobs tarde", _
        "% Jobs tarde", _
        "% Jobs a tiempo")

    Dim i As Long
    For i = 0 To UBound(labels)
        With ws.Range(ws.Cells(indTopRow + 1 + i, 2), ws.Cells(indTopRow + 1 + i, 4))
            On Error Resume Next
            .UnMerge
            On Error GoTo 0
            .Merge
            .Value = labels(i)
            .Font.Bold = False
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Locked = True
        End With

        With ws.Cells(indTopRow + 1 + i, 5)
            .ClearContents
            .NumberFormat = "General"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Locked = True
        End With
    Next i

    With ws.Range(ws.Cells(indTopRow + 1, 2), ws.Cells(indTopRow + 8, 5))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
    End With
End Sub

' =========================================================
' LECTURA / VALIDACIÓN
' =========================================================
Public Function PM_ReadNM(ByVal ws As Worksheet, ByRef n As Long, ByRef m As Long) As Boolean
    PM_ReadNM = False

    Dim vN As Variant, vM As Variant
    vN = ws.Range(CELL_N).Value
    vM = ws.Range(CELL_M).Value

    If Not IsNumeric(vN) Or CLng(vN) <= 0 Then Exit Function
    If Not IsNumeric(vM) Or CLng(vM) <= 0 Then Exit Function

    n = CLng(vN)
    m = CLng(vM)

    If m > MACH_MAX Then Exit Function
    If n < m Then Exit Function

    PM_ReadNM = True
End Function

Private Function PM_ReadAndValidateJobsPerMachine( _
    ByVal ws As Worksheet, _
    ByVal n As Long, _
    ByVal m As Long, _
    ByRef jobsPerMac() As Long, _
    ByRef warn As String) As Boolean

    PM_ReadAndValidateJobsPerMachine = False
    warn = ""

    ReDim jobsPerMac(1 To m)

    Dim sumJobs As Long: sumJobs = 0
    Dim i As Long, rr As Long, v As Variant

    For i = 1 To m
        rr = MACH_FIRST_ROW + i - 1
        v = ws.Cells(rr, COL_JOBS).Value

        If IsEmpty(v) Or Trim$(CStr(v)) = "" Then
            warn = "Falta #Jobs para M" & i & " (celda " & ws.Cells(rr, COL_JOBS).Address(False, False) & ")."
            Exit Function
        End If
        If Not IsNumeric(v) Then
            warn = "#Jobs debe ser numérico para M" & i & "."
            Exit Function
        End If
        If CLng(v) <> CDbl(v) Then
            warn = "#Jobs debe ser entero para M" & i & "."
            Exit Function
        End If
        If CLng(v) < 0 Then
            warn = "#Jobs no puede ser negativo para M" & i & "."
            Exit Function
        End If

        jobsPerMac(i) = CLng(v)
        sumJobs = sumJobs + jobsPerMac(i)
    Next i

    If sumJobs <> n Then
        warn = "La suma de #Jobs por máquina debe ser n = " & n & ". Actualmente suma " & sumJobs & "."
        Exit Function
    End If

    PM_ReadAndValidateJobsPerMachine = True
End Function

' =========================================================
' UTILIDADES
' =========================================================
Private Function NzD_PM(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        NzD_PM = CDbl(v)
    Else
        NzD_PM = 0#
    End If
End Function

Private Sub DeleteChartIfExists_PM(ByVal ws As Worksheet, ByVal nm As String)
    Dim i As Long
    On Error Resume Next
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
    On Error GoTo 0
End Sub
Private Sub PM_SortByMachineAndSeq( _
    ByRef machIdx() As Long, ByRef seq() As Long, _
    ByRef jobName() As String, ByRef machName() As String, _
    ByRef r() As Double, ByRef p() As Double, ByRef s() As Double, ByRef d() As Double, ByRef w() As Double, _
    ByVal n As Long)

    Dim i As Long, j As Long

    For i = 1 To n - 1
        For j = i + 1 To n
            If (machIdx(j) < machIdx(i)) Or ((machIdx(j) = machIdx(i)) And (seq(j) < seq(i))) Then
                SwapL_PM machIdx(i), machIdx(j)
                SwapL_PM seq(i), seq(j)
                SwapS_PM jobName(i), jobName(j)
                SwapS_PM machName(i), machName(j)
                SwapD_PM r(i), r(j)
                SwapD_PM p(i), p(j)
                SwapD_PM s(i), s(j)
                SwapD_PM d(i), d(j)
                SwapD_PM w(i), w(j)
            End If
        Next j
    Next i
End Sub

Private Sub SwapL_PM(ByRef a As Long, ByRef b As Long)
    Dim t As Long
    t = a: a = b: b = t
End Sub

Private Sub SwapD_PM(ByRef a As Double, ByRef b As Double)
    Dim t As Double
    t = a: a = b: b = t
End Sub

Private Sub SwapS_PM(ByRef a As String, ByRef b As String)
    Dim t As String
    t = a: a = b: b = t
End Sub

Private Function PM_OneHot(ByVal m As Long, ByVal idx As Long, ByVal v As Double) As Variant
    Dim arr() As Double
    Dim i As Long
    ReDim arr(1 To m)
    For i = 1 To m
        arr(i) = 0#
    Next i
    arr(idx) = v
    PM_OneHot = arr
End Function

Private Function ColorForJob_PM(ByVal idx As Long) As Long
    Dim palette As Variant
    palette = Array( _
        RGB(91, 155, 213), RGB(237, 125, 49), RGB(165, 165, 165), RGB(255, 192, 0), _
        RGB(68, 114, 196), RGB(112, 173, 71), RGB(37, 94, 145), RGB(158, 72, 14), _
        RGB(99, 99, 99), RGB(153, 115, 0), RGB(38, 68, 120), RGB(67, 104, 43))
    ColorForJob_PM = palette((idx - 1) Mod (UBound(palette) + 1))
End Function

Private Function LightTone_PM(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = c Mod 256
    g = (c \ 256) Mod 256
    b = (c \ 65536) Mod 256

    r = WorksheetFunction.Min(255, r + 55)
    g = WorksheetFunction.Min(255, g + 55)
    b = WorksheetFunction.Min(255, b + 55)

    LightTone_PM = RGB(r, g, b)
End Function

Private Function DarkTone_PM(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = c Mod 256
    g = (c \ 256) Mod 256
    b = (c \ 65536) Mod 256

    r = WorksheetFunction.Max(0, r - 25)
    g = WorksheetFunction.Max(0, g - 25)
    b = WorksheetFunction.Max(0, b - 25)

    DarkTone_PM = RGB(r, g, b)
End Function

Private Sub ConfigurarEjeTiempo_Inteligente_PM(ByVal ch As Chart, ByVal maxTime As Double)
    If maxTime <= 0# Then maxTime = 1#
    Dim majorU As Double: majorU = NiceStep_12525_PM(maxTime / 30#)
    If majorU <= 0# Then majorU = 1#
    Dim minorU As Double: minorU = majorU / 5#
    If minorU < 1# And majorU >= 5# Then minorU = 1#

    Dim eje As Axis: Set eje = ch.Axes(xlValue)
    eje.MinimumScale = 0#
    eje.MaximumScale = WorksheetFunction.Ceiling(maxTime, majorU)
    eje.MajorUnit = majorU
    eje.MinorUnit = minorU
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

Private Function NiceStep_12525_PM(ByVal x As Double) As Double
    If x <= 0# Then NiceStep_12525_PM = 1#: Exit Function
    Dim p As Double: p = 10# ^ Int(Log(x) / Log(10#))
    Dim m As Double: m = x / p
    Select Case m
        Case Is <= 1#: NiceStep_12525_PM = 1# * p
        Case Is <= 2#: NiceStep_12525_PM = 2# * p
        Case Is <= 5#: NiceStep_12525_PM = 5# * p
        Case Else: NiceStep_12525_PM = 10# * p
    End Select
End Function

