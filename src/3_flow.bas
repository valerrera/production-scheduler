Option Explicit

Private Const SHEET_NAME As String = "FlowShop"

' =========================
' CELDAS FIJAS
' =========================
Private Const CELL_N As String = "C12"          ' # jobs
Private Const CELL_M As String = "C13"          ' # máquinas
Private Const CELL_FS_READY As String = "Z2"    ' flag interno

' =========================
' ZONA DINÁMICA
' =========================
Private Const DYNAMIC_TOP_ROW As Long = 15
Private Const INPUT_CLEAR_ROWS As Long = 1200

' =========================
' COLUMNAS BASE
' =========================
Private Const DEC_C0 As Long = 2                ' B
Private Const PAR_C0 As Long = 2                ' B
Private Const DEC_NUM_COLS As Long = 2          ' Secuencia, Job

Private Const RM_COL_MACH As Long = 2           ' B
Private Const RM_COL_RMAQ As Long = 3           ' C

' =========================
' CHART
' =========================
Private Const CHART_NAME_TIMELINE As String = "chTimeline_FS"

' =========================
' CACHE
' =========================
Private fsLoaded As Boolean
Private fsN As Long
Private fsM As Long

Private fsSeq() As Long
Private fsJob() As String
Private fsR() As Double
Private fsP() As Double
Private fsS() As Double
Private fsD() As Double
Private fsW() As Double
Private fsRmaq() As Double

' =========================================================
' MACROS PRINCIPALES
' =========================================================
Public Sub RedibujarInputs_FS(ByVal ws As Worksheet)
    On Error GoTo ErrH

    Dim n As Long, m As Long
    If Not FS_ReadNM(ws, n, m) Then Exit Sub

    ws.Unprotect

    fsLoaded = False
    fsN = 0
    fsM = 0
    ws.Range(CELL_FS_READY).Value = ""

    FS_ClearDynamicZone ws
    FS_DeleteChartIfExists ws

    Dim parTitleRow As Long, parHeaderRow As Long, parFirstRow As Long, parLastRow As Long
    Dim rmaqTitleRow As Long, rmaqHeaderRow As Long, rmaqFirstRow As Long, rmaqLastRow As Long
    Dim decTitleRow As Long, decHeaderRow As Long, decFirstRow As Long, decLastRow As Long
    Dim instr4Row As Long, instr5Row As Long

    FS_GetLayoutRows n, m, _
        parTitleRow, parHeaderRow, parFirstRow, parLastRow, _
        rmaqTitleRow, rmaqHeaderRow, rmaqFirstRow, rmaqLastRow, _
        decTitleRow, decHeaderRow, decFirstRow, decLastRow, _
        instr4Row, instr5Row

    ' ===== 2. PARÁMETROS =====
    With ws.Cells(parTitleRow, PAR_C0)
        .Value = "2. Parámetros por job (r, p1..pm, s1..sm, d, w)."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    ws.Cells(parHeaderRow, PAR_C0 + 0).Value = "Job"
    ws.Cells(parHeaderRow, PAR_C0 + 1).Value = "r"

    Dim k As Long
    For k = 1 To m
        ws.Cells(parHeaderRow, PAR_C0 + 1 + k).Value = "p" & k
    Next k
    For k = 1 To m
        ws.Cells(parHeaderRow, PAR_C0 + 1 + m + k).Value = "s" & k
    Next k

    ws.Cells(parHeaderRow, PAR_C0 + 2 + 2 * m).Value = "d"
    ws.Cells(parHeaderRow, PAR_C0 + 3 + 2 * m).Value = "w"

    Dim lastColPar As Long
    lastColPar = PAR_C0 + 3 + 2 * m

    With ws.Range(ws.Cells(parHeaderRow, PAR_C0), ws.Cells(parHeaderRow, lastColPar))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = True
    End With

    With ws.Range(ws.Cells(parFirstRow, PAR_C0), ws.Cells(parLastRow, lastColPar))
        .ClearContents
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(160, 160, 160)
        .Borders.Weight = xlThin
        .Locked = False
    End With

    Dim i As Long
    For i = 1 To n
        ws.Cells(parFirstRow + i - 1, PAR_C0).Value = "J" & i
        ws.Cells(parFirstRow + i - 1, PAR_C0 + 3 + 2 * m).Value = 1
    Next i

    ws.Range(ws.Cells(parFirstRow, PAR_C0), ws.Cells(parLastRow, PAR_C0)).Locked = True
    FS_SetJobListNamedRange ws, parFirstRow, parLastRow

    ' ===== 3.1 rmaq =====
    With ws.Cells(rmaqTitleRow, RM_COL_MACH)
        .Value = "3.1. Escribe la fecha de disponibilidad de la(s) máquina(s)."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    ws.Cells(rmaqHeaderRow, RM_COL_MACH).Value = "Máquina"
    ws.Cells(rmaqHeaderRow, RM_COL_RMAQ).Value = "rmaq"

    With ws.Range(ws.Cells(rmaqHeaderRow, RM_COL_MACH), ws.Cells(rmaqHeaderRow, RM_COL_RMAQ))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = True
    End With

    With ws.Range(ws.Cells(rmaqFirstRow, RM_COL_MACH), ws.Cells(rmaqLastRow, RM_COL_RMAQ))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(160, 160, 160)
        .Borders.Weight = xlThin
    End With

    ws.Range(ws.Cells(rmaqFirstRow, RM_COL_RMAQ), ws.Cells(rmaqLastRow, RM_COL_RMAQ)).ClearContents
    For k = 1 To m
        ws.Cells(rmaqFirstRow + k - 1, RM_COL_MACH).Value = "M" & k
    Next k
    ws.Range(ws.Cells(rmaqFirstRow, RM_COL_MACH), ws.Cells(rmaqLastRow, RM_COL_MACH)).Locked = True
    ws.Range(ws.Cells(rmaqFirstRow, RM_COL_RMAQ), ws.Cells(rmaqLastRow, RM_COL_RMAQ)).Locked = False

    ' ===== 4. DECISIÓN =====
    With ws.Cells(decTitleRow, DEC_C0)
        .Value = "4. Decisión (escribe la secuencia y el job)."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    ws.Cells(decHeaderRow, DEC_C0 + 0).Value = "Secuencia"
    ws.Cells(decHeaderRow, DEC_C0 + 1).Value = "Job"

    With ws.Range(ws.Cells(decHeaderRow, DEC_C0), ws.Cells(decHeaderRow, DEC_C0 + DEC_NUM_COLS - 1))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = True
    End With

    With ws.Range(ws.Cells(decFirstRow, DEC_C0), ws.Cells(decLastRow, DEC_C0 + DEC_NUM_COLS - 1))
        .ClearContents
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(160, 160, 160)
        .Borders.Weight = xlThin
        .Locked = False
    End With
    
    Dim rngDV As Range
    Set rngDV = ws.Range(ws.Cells(decFirstRow, DEC_C0 + 1), ws.Cells(decLastRow, DEC_C0 + 1))
    
    On Error Resume Next
    rngDV.Validation.Delete
    On Error GoTo 0
    
    rngDV.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=FS_JobList"
    rngDV.Validation.IgnoreBlank = True
    rngDV.Validation.InCellDropdown = True
    rngDV.Validation.ShowError = True

    For i = 1 To n
        ws.Cells(decFirstRow + i - 1, DEC_C0).Value = i
    Next i

    ' ===== 5. CARGAR DATOS =====
    With ws.Cells(instr4Row, DEC_C0)
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
    MsgBox "Error en RedibujarInputs_FS: " & Err.Description, vbExclamation
    Resume Salir
End Sub

Public Sub FS_CargarDatos()
    On Error GoTo ErrH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim n As Long, m As Long
    If Not FS_ReadNM(ws, n, m) Then
        MsgBox "Verifica que # jobs y # máquinas sean enteros positivos.", vbExclamation
        Exit Sub
    End If

    ws.Unprotect

    Dim parTitleRow As Long, parHeaderRow As Long, parFirstRow As Long, parLastRow As Long
    Dim rmaqTitleRow As Long, rmaqHeaderRow As Long, rmaqFirstRow As Long, rmaqLastRow As Long
    Dim decTitleRow As Long, decHeaderRow As Long, decFirstRow As Long, decLastRow As Long
    Dim instr4Row As Long, instr5Row As Long

    FS_GetLayoutRows n, m, _
        parTitleRow, parHeaderRow, parFirstRow, parLastRow, _
        rmaqTitleRow, rmaqHeaderRow, rmaqFirstRow, rmaqLastRow, _
        decTitleRow, decHeaderRow, decFirstRow, decLastRow, _
        instr4Row, instr5Row
        
    Dim outLineRow As Long
    outLineRow = instr5Row + 1
    
    FS_DeleteChartIfExists ws
    
    With ws.Range(ws.Cells(outLineRow, 1), ws.Cells(outLineRow + 700, 35))
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

    ws.Range(CELL_FS_READY).Value = ""
    With ws.Cells(instr5Row, DEC_C0)
        .ClearContents
        .Font.Bold = False
        .Font.Italic = False
    End With

    Dim warn As String
    If Not FS_ValidateInputs(ws, n, m, parFirstRow, rmaqFirstRow, decFirstRow, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos - Flow Shop"
        GoTo Salir
    End If

    FS_LoadInputsToCache ws, n, m, parFirstRow, rmaqFirstRow, decFirstRow

    ws.Range(CELL_FS_READY).Value = "OK"
    ws.Range(CELL_FS_READY).Locked = True

    With ws.Cells(instr5Row, DEC_C0)
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
    MsgBox "Error en FS_CargarDatos: " & Err.Description, vbExclamation
    Resume Salir
End Sub

Public Sub FS_GenerarOutputs()
    On Error GoTo ErrH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    If UCase$(Trim$(CStr(ws.Range(CELL_FS_READY).Value))) <> "OK" Then
        MsgBox "Primero presione 'Cargar datos' y asegúrese de que los datos sean válidos.", vbExclamation
        Exit Sub
    End If

    Dim n As Long, m As Long
    If Not FS_ReadNM(ws, n, m) Then Exit Sub

    If Not fsLoaded Or fsN <> n Or fsM <> m Then
        MsgBox "Los datos no están cargados en memoria. Presione 'Cargar datos' nuevamente.", vbExclamation
        Exit Sub
    End If

    ws.Unprotect

    Dim parTitleRow As Long, parHeaderRow As Long, parFirstRow As Long, parLastRow As Long
    Dim rmaqTitleRow As Long, rmaqHeaderRow As Long, rmaqFirstRow As Long, rmaqLastRow As Long
    Dim decTitleRow As Long, decHeaderRow As Long, decFirstRow As Long, decLastRow As Long
    Dim instr4Row As Long, instr5Row As Long

    FS_GetLayoutRows n, m, _
        parTitleRow, parHeaderRow, parFirstRow, parLastRow, _
        rmaqTitleRow, rmaqHeaderRow, rmaqFirstRow, rmaqLastRow, _
        decTitleRow, decHeaderRow, decFirstRow, decLastRow, _
        instr4Row, instr5Row

    Dim outLineRow As Long: outLineRow = instr5Row + 1
    Dim outTitleRow As Long: outTitleRow = outLineRow + 2
    Dim ganttTopRow As Long: ganttTopRow = outTitleRow + 2
    Dim indTopRow As Long: indTopRow = ganttTopRow + 14 + 5

    FS_DeleteChartIfExists ws

    With ws.Range(ws.Cells(outLineRow, 1), ws.Cells(outLineRow + 700, 35))
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

    With ws.Cells(outTitleRow, DEC_C0)
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
        .ChartTitle.text = "Gantt Flow Shop"
        .HasLegend = False
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
    End With

    FS_DrawIndicatorsStructure ws, indTopRow

    Dim jmFirstCol As Long: jmFirstCol = 8 ' H
    Dim jmNumCols As Long: jmNumCols = 7
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

    Dim headers As Variant
    headers = Array("Job", "Inicio", "Cj", "Flow (Cj-rj)", "L (Cj-dj)", "T=max(L,0)", "w*T")

    Dim j As Long
    For j = 0 To UBound(headers)
        With ws.Cells(jmHeaderRow, jmFirstCol + j)
            .Value = headers(j)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(230, 230, 230)
            .Locked = True
        End With
    Next j

    With ws.Range(ws.Cells(jmFirstDataRow, jmFirstCol), ws.Cells(jmLastRow, jmFirstCol + jmNumCols - 1))
        .ClearContents
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Pattern = xlNone
        .Locked = True
    End With

    With ws.Range(ws.Cells(jmHeaderRow, jmFirstCol), ws.Cells(jmLastRow, jmFirstCol + jmNumCols - 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
    End With

    Dim c() As Double
    Dim startJob() As Double
    Dim Cj() As Double, Flow() As Double, L() As Double, Tard() As Double, wT() As Double
    Dim Cmax As Double, avgFlow As Double, Lmax As Double, avgT As Double, sumWT As Double
    Dim lateCount As Long, pctLate As Double, pctOnTime As Double

    FS_CalcFromCache n, m, c, startJob, Cj, Flow, L, Tard, wT, Cmax, avgFlow, Lmax, avgT, sumWT, lateCount, pctLate, pctOnTime
    FS_WriteOutputs ws, n, indTopRow, jmFirstCol, jmFirstDataRow, startJob, Cj, Flow, L, Tard, wT, Cmax, avgFlow, Lmax, avgT, sumWT, lateCount, pctLate, pctOnTime
    FS_BuildGanttFromCache chObj.Chart, n, m, c, Cmax

Salir:
    ws.Range(CELL_N).Locked = False
    ws.Range(CELL_M).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub

ErrH:
    MsgBox "Error en FS_GenerarOutputs: " & Err.Description, vbExclamation
    Resume Salir
End Sub

' =========================================================
' VALIDACIÓN
' =========================================================
Private Function FS_ValidateInputs( _
    ByVal ws As Worksheet, _
    ByVal n As Long, _
    ByVal m As Long, _
    ByVal parFirstRow As Long, _
    ByVal rmaqFirstRow As Long, _
    ByVal decFirstRow As Long, _
    ByRef warn As String) As Boolean

    FS_ValidateInputs = False
    warn = ""

    Dim dictParJobs As Object
    Set dictParJobs = CreateObject("Scripting.Dictionary")

    Dim i As Long, k As Long, rowN As Long
    Dim jb As String, vR As Variant, vD As Variant, vW As Variant

    ' =========================
    ' VALIDAR PARÁMETROS
    ' =========================
    For i = 1 To n
        rowN = parFirstRow + i - 1
        jb = UCase$(Trim$(CStr(ws.Cells(rowN, PAR_C0).Value)))

        If Len(jb) = 0 Then
            warn = "Falta Job en PARÁMETROS (fila " & rowN & ")."
            Exit Function
        End If
        If dictParJobs.Exists(jb) Then
            warn = "Job repetido en PARÁMETROS: '" & jb & "'."
            Exit Function
        End If
        dictParJobs.Add jb, True

        vR = ws.Cells(rowN, PAR_C0 + 1).Value
        vD = ws.Cells(rowN, PAR_C0 + 2 + 2 * m).Value
        vW = ws.Cells(rowN, PAR_C0 + 3 + 2 * m).Value

        If Not IsNumeric(vR) Then
            warn = "r debe ser numérico para el job '" & jb & "'."
            Exit Function
        End If
        If CDbl(vR) < 0 Then
            warn = "r no puede ser negativo para el job '" & jb & "'."
            Exit Function
        End If

        For k = 1 To m
            If Not IsNumeric(ws.Cells(rowN, PAR_C0 + 1 + k).Value) Then
                warn = "p" & k & " debe ser numérico para el job '" & jb & "'."
                Exit Function
            End If
            If Not IsNumeric(ws.Cells(rowN, PAR_C0 + 1 + m + k).Value) Then
                warn = "s" & k & " debe ser numérico para el job '" & jb & "'."
                Exit Function
            End If
            If CDbl(ws.Cells(rowN, PAR_C0 + 1 + k).Value) <= 0 Then
                warn = "p" & k & " debe ser > 0 para el job '" & jb & "'."
                Exit Function
            End If
            If CDbl(ws.Cells(rowN, PAR_C0 + 1 + m + k).Value) < 0 Then
                warn = "s" & k & " no puede ser negativo para el job '" & jb & "'."
                Exit Function
            End If
        Next k

        If Not IsNumeric(vD) Or CDbl(vD) <= 0 Then
            warn = "d debe ser > 0 para el job '" & jb & "'."
            Exit Function
        End If

        If Len(Trim$(CStr(vW))) = 0 Then
            warn = "w es obligatorio para el job '" & jb & "'."
            Exit Function
        End If
        If Not IsNumeric(vW) Or CDbl(vW) <= 0 Then
            warn = "w debe ser > 0 para el job '" & jb & "'."
            Exit Function
        End If
    Next i

    ' =========================
    ' VALIDAR DECISIÓN
    ' =========================
    Dim dictSeq As Object, dictDecJobs As Object
    Set dictSeq = CreateObject("Scripting.Dictionary")
    Set dictDecJobs = CreateObject("Scripting.Dictionary")

    Dim vSeq As Variant, vJobDec As String

    For i = 1 To n
        rowN = decFirstRow + i - 1

        vSeq = ws.Cells(rowN, DEC_C0).Value
        vJobDec = UCase$(Trim$(CStr(ws.Cells(rowN, DEC_C0 + 1).Value)))

        If Len(Trim$(CStr(vSeq))) = 0 Then
            warn = "Falta Secuencia en DECISIÓN (fila " & rowN & ")."
            Exit Function
        End If
        If Not IsNumeric(vSeq) Then
            warn = "Secuencia debe ser numérica (fila " & rowN & ")."
            Exit Function
        End If
        If CLng(vSeq) <> CDbl(vSeq) Or CLng(vSeq) <= 0 Then
            warn = "Secuencia debe ser entero positivo (fila " & rowN & ")."
            Exit Function
        End If
        If dictSeq.Exists(CStr(CLng(vSeq))) Then
            warn = "Secuencia repetida: " & CLng(vSeq) & "."
            Exit Function
        End If
        dictSeq.Add CStr(CLng(vSeq)), True

        If Len(vJobDec) = 0 Then
            warn = "Falta Job en DECISIÓN (fila " & rowN & ")."
            Exit Function
        End If
        If Not dictParJobs.Exists(vJobDec) Then
            warn = "El Job '" & vJobDec & "' no existe en la tabla de PARÁMETROS."
            Exit Function
        End If
        If dictDecJobs.Exists(vJobDec) Then
            warn = "Job repetido en DECISIÓN: '" & vJobDec & "'."
            Exit Function
        End If
        dictDecJobs.Add vJobDec, True
    Next i

    Dim need As Long
    For need = 1 To n
        If Not dictSeq.Exists(CStr(need)) Then
            warn = "Secuencia debe ser 1.." & n & " sin saltos. Falta: " & need & "."
            Exit Function
        End If
    Next need

    ' =========================
    ' VALIDAR RMAQ
    ' =========================
    Dim mk As Long, rv As Variant
    For mk = 1 To m
        rv = ws.Cells(rmaqFirstRow + mk - 1, RM_COL_RMAQ).Value

        If Len(Trim$(CStr(rv))) = 0 Then
            warn = "Falta rmaq para M" & mk & "."
            Exit Function
        End If
        If Not IsNumeric(rv) Then
            warn = "rmaq debe ser numérico para M" & mk & "."
            Exit Function
        End If
        If CDbl(rv) < 0 Then
            warn = "rmaq no puede ser negativo para M" & mk & "."
            Exit Function
        End If
    Next mk

    FS_ValidateInputs = True
End Function

' =========================================================
' CACHE
' =========================================================
Private Sub FS_LoadInputsToCache(ByVal ws As Worksheet, ByVal n As Long, ByVal m As Long, ByVal parFirstRow As Long, ByVal rmaqFirstRow As Long, ByVal decFirstRow As Long)
    Dim i As Long, k As Long, rowN As Long, jb As String

    fsLoaded = False
    fsN = n
    fsM = m

    ReDim fsSeq(1 To n)
    ReDim fsJob(1 To n)
    ReDim fsR(1 To n)
    ReDim fsP(1 To n, 1 To m)
    ReDim fsS(1 To n, 1 To m)
    ReDim fsD(1 To n)
    ReDim fsW(1 To n)
    ReDim fsRmaq(1 To m)

    For k = 1 To m
        fsRmaq(k) = CDbl(ws.Cells(rmaqFirstRow + k - 1, RM_COL_RMAQ).Value)
    Next k

    ' diccionarios para buscar parámetros por job
    Dim dictR As Object, dictD As Object, dictW As Object
    Dim dictP As Object, dictS As Object
    Set dictR = CreateObject("Scripting.Dictionary")
    Set dictD = CreateObject("Scripting.Dictionary")
    Set dictW = CreateObject("Scripting.Dictionary")
    Set dictP = CreateObject("Scripting.Dictionary")
    Set dictS = CreateObject("Scripting.Dictionary")

    Dim arrP() As Double, arrS() As Double

    For i = 1 To n
        rowN = parFirstRow + i - 1
        jb = UCase$(Trim$(CStr(ws.Cells(rowN, PAR_C0).Value)))

        dictR(jb) = CDbl(ws.Cells(rowN, PAR_C0 + 1).Value)
        dictD(jb) = CDbl(ws.Cells(rowN, PAR_C0 + 2 + 2 * m).Value)
        dictW(jb) = CDbl(ws.Cells(rowN, PAR_C0 + 3 + 2 * m).Value)

        ReDim arrP(1 To m)
        ReDim arrS(1 To m)

        For k = 1 To m
            arrP(k) = CDbl(ws.Cells(rowN, PAR_C0 + 1 + k).Value)
            arrS(k) = CDbl(ws.Cells(rowN, PAR_C0 + 1 + m + k).Value)
        Next k

        dictP(jb) = arrP
        dictS(jb) = arrS
    Next i

    ' cargar según decisión
    For i = 1 To n
        rowN = decFirstRow + i - 1
        jb = UCase$(Trim$(CStr(ws.Cells(rowN, DEC_C0 + 1).Value)))

        fsSeq(i) = CLng(ws.Cells(rowN, DEC_C0).Value)
        fsJob(i) = jb
        fsR(i) = dictR(jb)
        fsD(i) = dictD(jb)
        fsW(i) = dictW(jb)

        arrP = dictP(jb)
        arrS = dictS(jb)

        For k = 1 To m
            fsP(i, k) = arrP(k)
            fsS(i, k) = arrS(k)
        Next k
    Next i

    FS_SortBySequence fsSeq, fsJob, fsR, fsP, fsS, fsD, fsW, n, m
    fsLoaded = True
End Sub

' =========================================================
' CÁLCULO
' =========================================================
Private Sub FS_CalcFromCache( _
    ByVal n As Long, ByVal m As Long, _
    ByRef c() As Double, ByRef startJob() As Double, _
    ByRef Cj() As Double, ByRef Flow() As Double, ByRef L() As Double, ByRef Tard() As Double, ByRef wT() As Double, _
    ByRef Cmax As Double, ByRef avgFlow As Double, ByRef Lmax As Double, ByRef avgT As Double, _
    ByRef sumWT As Double, ByRef lateCount As Long, ByRef pctLate As Double, ByRef pctOnTime As Double)

    ReDim c(1 To n, 1 To m)
    ReDim startJob(1 To n)
    ReDim Cj(1 To n)
    ReDim Flow(1 To n)
    ReDim L(1 To n)
    ReDim Tard(1 To n)
    ReDim wT(1 To n)

    Dim lastOnMac() As Double
    ReDim lastOnMac(1 To m)

    Dim kk As Long
    For kk = 1 To m
        lastOnMac(kk) = fsRmaq(kk)
    Next kk

    Dim i As Long, k As Long, startBeforeSetup As Double
    For i = 1 To n
        For k = 1 To m
            If k = 1 Then
                startBeforeSetup = WorksheetFunction.Max(fsR(i), lastOnMac(k))
            Else
                startBeforeSetup = WorksheetFunction.Max(c(i, k - 1), lastOnMac(k))
            End If

            If k = 1 Then startJob(i) = startBeforeSetup
            c(i, k) = startBeforeSetup + fsS(i, k) + fsP(i, k)
            lastOnMac(k) = c(i, k)
        Next k
    Next i

    Lmax = -1E+30

    For i = 1 To n
        Cj(i) = c(i, m)
        Flow(i) = Cj(i) - fsR(i)
        L(i) = Cj(i) - fsD(i)
        If L(i) > 0# Then
            Tard(i) = L(i)
        Else
            Tard(i) = 0#
        End If
        wT(i) = fsW(i) * Tard(i)

        If Cj(i) > Cmax Then Cmax = Cj(i)
        avgFlow = avgFlow + Flow(i)
        If L(i) > Lmax Then Lmax = L(i)
        avgT = avgT + Tard(i)
        sumWT = sumWT + wT(i)
        If Tard(i) > 0.000001 Then lateCount = lateCount + 1
    Next i

    avgFlow = avgFlow / CDbl(n)
    avgT = avgT / CDbl(n)
    pctLate = lateCount / CDbl(n)
    pctOnTime = 1# - pctLate
End Sub

' =========================================================
' OUTPUTS
' =========================================================
Private Sub FS_WriteOutputs( _
    ByVal ws As Worksheet, _
    ByVal n As Long, _
    ByVal indTopRow As Long, _
    ByVal jmFirstCol As Long, _
    ByVal jmFirstDataRow As Long, _
    ByRef startJob() As Double, ByRef Cj() As Double, ByRef Flow() As Double, ByRef L() As Double, ByRef Tard() As Double, ByRef wT() As Double, _
    ByVal Cmax As Double, ByVal avgFlow As Double, ByVal Lmax As Double, ByVal avgT As Double, _
    ByVal sumWT As Double, ByVal lateCount As Long, ByVal pctLate As Double, ByVal pctOnTime As Double)

    ws.Cells(indTopRow + 1, 5).Value = Cmax
    ws.Cells(indTopRow + 2, 5).Value = avgFlow
    ws.Cells(indTopRow + 3, 5).Value = Lmax
    ws.Cells(indTopRow + 4, 5).Value = avgT
    ws.Cells(indTopRow + 5, 5).Value = sumWT
    ws.Cells(indTopRow + 6, 5).Value = lateCount
    ws.Cells(indTopRow + 7, 5).Value = pctLate
    ws.Cells(indTopRow + 8, 5).Value = pctOnTime

    ws.Cells(indTopRow + 7, 5).NumberFormat = "0%"
    ws.Cells(indTopRow + 8, 5).NumberFormat = "0%"

    Dim i As Long
    For i = 1 To n
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 0).Value = fsJob(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 1).Value = startJob(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 2).Value = Cj(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 3).Value = Flow(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 4).Value = L(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 5).Value = Tard(i)
        ws.Cells(jmFirstDataRow + i - 1, jmFirstCol + 6).Value = wT(i)
    Next i
End Sub

Private Sub FS_BuildGanttFromCache(ByVal ch As Chart, ByVal n As Long, ByVal m As Long, ByRef c() As Double, ByVal Cmax As Double)
    Dim cats() As Variant
    ReDim cats(1 To m)

    Dim k As Long
    For k = 1 To m
        cats(k) = "M" & k
    Next k

    With ch
        .ChartType = xlBarStacked
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.text = "Gantt Flow Shop"

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
        For k = 1 To m
            offByMac(k) = 0#
        Next k

        Dim i As Long, gap As Double, startBeforeSetup As Double
        Dim setupDur As Double, procDur As Double

        For i = 1 To n
            For k = 1 To m
                startBeforeSetup = c(i, k) - fsS(i, k) - fsP(i, k)

                gap = startBeforeSetup - offByMac(k)
                If gap > 0.000001 Then
                    Set srs = .SeriesCollection.NewSeries
                    srs.Values = FS_OneHot(m, k, gap)
                    srs.XValues = cats
                    srs.Format.Fill.Visible = msoFalse
                    srs.Format.Line.Visible = msoFalse
                    offByMac(k) = offByMac(k) + gap
                End If

                setupDur = fsS(i, k)
                If setupDur > 0.000001 Then
                    Set srs = .SeriesCollection.NewSeries
                    srs.Values = FS_OneHot(m, k, setupDur)
                    srs.XValues = cats
                    srs.Format.Fill.ForeColor.RGB = FS_LightTone(FS_ColorForJob(i))
                    srs.Format.Line.Visible = msoFalse
                    offByMac(k) = offByMac(k) + setupDur
                End If

                procDur = fsP(i, k)
                If procDur > 0.000001 Then
                    Set srs = .SeriesCollection.NewSeries
                    srs.Values = FS_OneHot(m, k, procDur)
                    srs.XValues = cats
                    srs.Format.Fill.ForeColor.RGB = FS_DarkTone(FS_ColorForJob(i))
                    srs.Format.Line.Visible = msoFalse

                    With srs.Points(k)
                        .HasDataLabel = True
                        .DataLabel.text = fsJob(i)
                        .DataLabel.Font.Size = 9
                    End With

                    offByMac(k) = offByMac(k) + procDur
                End If
            Next k
        Next i

        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.text = "Tiempo"
        FS_ConfigurarEjeTiempo ch, Cmax
    End With
End Sub

' =========================================================
' LAYOUT
' =========================================================
Private Sub FS_GetLayoutRows( _
    ByVal n As Long, ByVal m As Long, _
    ByRef parTitleRow As Long, ByRef parHeaderRow As Long, ByRef parFirstRow As Long, ByRef parLastRow As Long, _
    ByRef rmaqTitleRow As Long, ByRef rmaqHeaderRow As Long, ByRef rmaqFirstRow As Long, ByRef rmaqLastRow As Long, _
    ByRef decTitleRow As Long, ByRef decHeaderRow As Long, ByRef decFirstRow As Long, ByRef decLastRow As Long, _
    ByRef instr4Row As Long, ByRef instr5Row As Long)

    ' después de 1. Jobs y máquinas
    parTitleRow = 16
    parHeaderRow = parTitleRow + 1
    parFirstRow = parHeaderRow + 1
    parLastRow = parFirstRow + n - 1

    ' después de parámetros
    rmaqTitleRow = parLastRow + 3
    rmaqHeaderRow = rmaqTitleRow + 1
    rmaqFirstRow = rmaqHeaderRow + 1
    rmaqLastRow = rmaqFirstRow + m - 1

    ' después de rmaq
    decTitleRow = rmaqLastRow + 3
    decHeaderRow = decTitleRow + 1
    decFirstRow = decHeaderRow + 1
    decLastRow = decFirstRow + n - 1

    ' después de decisión
    instr4Row = decLastRow + 3
    instr5Row = instr4Row + 2
End Sub

Private Sub FS_DrawIndicatorsStructure(ByVal ws As Worksheet, ByVal indTopRow As Long)
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
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Locked = True
        End With

        With ws.Cells(indTopRow + 1 + i, 5)
            .ClearContents
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
' LECTURA n,m
' =========================================================
Public Function FS_ReadNM(ByVal ws As Worksheet, ByRef n As Long, ByRef m As Long) As Boolean
    FS_ReadNM = False

    If Not IsNumeric(ws.Range(CELL_N).Value) Then Exit Function
    If Not IsNumeric(ws.Range(CELL_M).Value) Then Exit Function

    n = CLng(ws.Range(CELL_N).Value)
    m = CLng(ws.Range(CELL_M).Value)

    If n <= 0 Or m <= 0 Then Exit Function
    FS_ReadNM = True
End Function

' =========================================================
' UTILIDADES
' =========================================================
Private Sub FS_ClearDynamicZone(ByVal ws As Worksheet)
    With ws.Range(ws.Cells(DYNAMIC_TOP_ROW, 1), ws.Cells(DYNAMIC_TOP_ROW + INPUT_CLEAR_ROWS, 60))
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

Private Sub FS_DeleteChartIfExists(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Sub FS_SortBySequence( _
    ByRef seq() As Long, ByRef job() As String, ByRef rr() As Double, _
    ByRef p() As Double, ByRef s() As Double, ByRef dd() As Double, ByRef ww() As Double, _
    ByVal n As Long, ByVal m As Long)

    Dim i As Long, j As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If seq(j) < seq(i) Then
                FS_SwapRow i, j, seq, job, rr, p, s, dd, ww, m
            End If
        Next j
    Next i
End Sub

Private Sub FS_SwapRow( _
    ByVal i As Long, ByVal j As Long, _
    ByRef seq() As Long, ByRef job() As String, ByRef rr() As Double, _
    ByRef p() As Double, ByRef s() As Double, ByRef dd() As Double, ByRef ww() As Double, _
    ByVal m As Long)

    Dim tL As Long, tS As String, tD As Double
    Dim k As Long

    tL = seq(i): seq(i) = seq(j): seq(j) = tL
    tS = job(i): job(i) = job(j): job(j) = tS
    tD = rr(i): rr(i) = rr(j): rr(j) = tD
    tD = dd(i): dd(i) = dd(j): dd(j) = tD
    tD = ww(i): ww(i) = ww(j): ww(j) = tD

    For k = 1 To m
        tD = p(i, k): p(i, k) = p(j, k): p(j, k) = tD
        tD = s(i, k): s(i, k) = s(j, k): s(j, k) = tD
    Next k
End Sub

Private Function FS_OneHot(ByVal m As Long, ByVal idx As Long, ByVal val As Double) As Variant
    Dim a() As Double, t As Long
    ReDim a(1 To m)
    For t = 1 To m
        a(t) = 0#
    Next t
    a(idx) = val
    FS_OneHot = a
End Function

Private Function FS_ColorForJob(ByVal i As Long) As Long
    Dim pal As Variant
    pal = Array( _
        RGB(52, 96, 174), RGB(46, 204, 113), RGB(155, 89, 182), _
        RGB(241, 196, 15), RGB(231, 76, 60), RGB(26, 188, 156), RGB(127, 140, 141))
    FS_ColorForJob = pal((i - 1) Mod (UBound(pal) + 1))
End Function

Private Function FS_LightTone(ByVal col As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = (col And &HFF): g = (col \ &H100) And &HFF: b = (col \ &H10000) And &HFF
    FS_LightTone = RGB(IIf(r + 80 < 255, r + 80, 255), IIf(g + 80 < 255, g + 80, 255), IIf(b + 80 < 255, b + 80, 255))
End Function

Private Function FS_DarkTone(ByVal col As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = (col And &HFF): g = (col \ &H100) And &HFF: b = (col \ &H10000) And &HFF
    FS_DarkTone = RGB(IIf(r > 40, r - 40, 0), IIf(g > 40, g - 40, 0), IIf(b > 40, b - 40, 0))
End Function

Private Sub FS_ConfigurarEjeTiempo(ByVal ch As Chart, ByVal maxTime As Double)
    Dim eje As Axis: Set eje = ch.Axes(xlValue)
    If maxTime < 0.000001 Then maxTime = 1
    Dim majorU As Double: majorU = FS_NiceMajorUnit(maxTime, 30)
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

Private Function FS_NiceMajorUnit(ByVal maxScale As Double, ByVal targetTicks As Long) As Double
    If maxScale <= 0 Then FS_NiceMajorUnit = 1: Exit Function
    Dim rawStep As Double: rawStep = maxScale / targetTicks
    If rawStep < 1 Then rawStep = 1
    Dim pow10 As Double: pow10 = 10 ^ Int(Log(rawStep) / Log(10))
    Dim frac As Double: frac = rawStep / pow10
    Select Case frac
        Case Is <= 1: FS_NiceMajorUnit = 1 * pow10
        Case Is <= 2: FS_NiceMajorUnit = 2 * pow10
        Case Is <= 5: FS_NiceMajorUnit = 5 * pow10
        Case Else: FS_NiceMajorUnit = 10 * pow10
    End Select
End Function
Private Sub FS_SetJobListNamedRange(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error Resume Next
    ws.Parent.Names("FS_JobList").Delete
    On Error GoTo 0

    Dim rngJobs As Range
    Set rngJobs = ws.Range(ws.Cells(firstRow, PAR_C0), ws.Cells(lastRow, PAR_C0))

    ws.Parent.Names.Add name:="FS_JobList", RefersTo:=rngJobs
End Sub

