Option Explicit

Private Const SHEET_NAME As String = "SingleMachine"
Private Const CELL_N As String = "C12"

' Flag interno para saber si ya pasó validación
Private Const CELL_SM_READY As String = "Z2"

' --- Tabla B (Parámetros): Job r p s d w ---
Private Const B_HEADER_ROW As Long = 16
Private Const B_FIRST_DATA_ROW As Long = 17
Private Const B_FIRST_COL As Long = 2      ' B
Private Const B_NUM_COLS As Long = 6       ' Job + r p s d w
Private Const B_CLEAR_ROWS As Long = 600   ' cuánto limpiamos hacia abajo en inputs

' --- Chart ---
Private Const CHART_NAME_TIMELINE As String = "chLineaTiempo_SM"

' =========================
' CACHE EN MEMORIA (SingleMachine)
' =========================
Private smLoaded As Boolean
Private smN As Long
Private smRmaq As Double

Private smSeq() As Long
Private smJob() As String

Private smParR As Object
Private smParP As Object
Private smParS As Object
Private smParD As Object
Private smParW As Object

' =========================================================
' 1) ZONA DE INPUTS (dinámica con n)
'   - Tabla B: parámetros
'   - 2.1 rmaq: disponibilidad máquina
'   - Tabla A: secuencia
'   - Instrucción 4 siempre; instrucción 5 solo cuando valida
' =========================================================
Public Sub RedibujarInputs_SM(ByVal ws As Worksheet)
On Error GoTo ErrH

    Dim n As Long
    If Not IsNumeric(ws.Range(CELL_N).Value) Then Exit Sub
    n = CLng(ws.Range(CELL_N).Value)
    If n <= 0 Then Exit Sub

    ws.Unprotect

    ' Reset: al cambiar n, obligamos a recargar/validar
    ws.Range(CELL_SM_READY).Value = ""
    smLoaded = False
    smN = 0

    Dim bLastCol As Long: bLastCol = B_FIRST_COL + B_NUM_COLS - 1
    Dim bLastDataRow As Long: bLastDataRow = B_FIRST_DATA_ROW + n - 1

    ' ===== 2.1) rmaq debajo de tabla B =====
    Dim RMAQ_TITLE_ROW As Long: RMAQ_TITLE_ROW = bLastDataRow + 2
    Dim RMAQ_TABLE_ROW As Long: RMAQ_TABLE_ROW = RMAQ_TITLE_ROW + 2
    Dim RMAQ_LABEL_COL As Long: RMAQ_LABEL_COL = B_FIRST_COL          ' B
    Dim RMAQ_INPUT_COL As Long: RMAQ_INPUT_COL = B_FIRST_COL + 1      ' C

    ' ===== Tabla A (secuencia) debajo de rmaq =====
    Dim seqTitleRow As Long: seqTitleRow = RMAQ_TABLE_ROW + 3
    Dim A_HEADER_ROW As Long: A_HEADER_ROW = seqTitleRow + 1
    Dim A_FIRST_DATA_ROW As Long: A_FIRST_DATA_ROW = A_HEADER_ROW + 1
    Dim A_FIRST_COL As Long: A_FIRST_COL = 2
    Dim A_NUM_COLS As Long: A_NUM_COLS = 2
    Dim aLastCol As Long: aLastCol = A_FIRST_COL + A_NUM_COLS - 1     ' B..C
    Dim aLastDataRow As Long: aLastDataRow = A_FIRST_DATA_ROW + n - 1

    ' Instrucciones 4/5 debajo de A
    Dim instr4Row As Long: instr4Row = aLastDataRow + 3
    Dim instr5Row As Long: instr5Row = instr4Row + 2

    ' --- LIMPIEZA FUERTE (inputs + posibles outputs viejos en la zona) ---
    Dim topClear As Long: topClear = B_HEADER_ROW - 1
    Dim botClear As Long: botClear = topClear + B_CLEAR_ROWS

    With ws.Range(ws.Cells(topClear, 1), ws.Cells(botClear, 35)) ' A..AI
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
    End With

    DeleteChartIfExists ws, CHART_NAME_TIMELINE

    ' Limpia instrucción 5 (si existía de un n anterior)
    With ws.Cells(instr5Row, B_FIRST_COL)
        .ClearContents
        .Font.Bold = False
        .Font.Italic = False
    End With

    ' ===== 2) PARÁMETROS (título) =====
    With ws.Cells(B_HEADER_ROW - 1, B_FIRST_COL)
        .Value = "2. Parámetros (llena r (disponibilidad), p (tiempo de procesamiento), s (tiempo de alistamiento), d (fecha de entega), w (peso))."
        .Font.Bold = True
        .Font.Italic = True
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    ' ===== TABLA B =====
    ws.Cells(B_HEADER_ROW, B_FIRST_COL + 0).Value = "Job"
    ws.Cells(B_HEADER_ROW, B_FIRST_COL + 1).Value = "r"
    ws.Cells(B_HEADER_ROW, B_FIRST_COL + 2).Value = "p"
    ws.Cells(B_HEADER_ROW, B_FIRST_COL + 3).Value = "s"
    ws.Cells(B_HEADER_ROW, B_FIRST_COL + 4).Value = "d"
    ws.Cells(B_HEADER_ROW, B_FIRST_COL + 5).Value = "w"

    With ws.Range(ws.Cells(B_HEADER_ROW, B_FIRST_COL), ws.Cells(B_HEADER_ROW, bLastCol))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = True
    End With

    With ws.Range(ws.Cells(B_FIRST_DATA_ROW, B_FIRST_COL), ws.Cells(bLastDataRow, bLastCol))
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
        ws.Cells(B_FIRST_DATA_ROW + i - 1, B_FIRST_COL + 0).Value = "J" & i
        ws.Cells(B_FIRST_DATA_ROW + i - 1, B_FIRST_COL + 5).Value = 1 ' default, pero obligatorio
    Next i

    ws.Range(ws.Cells(B_FIRST_DATA_ROW, B_FIRST_COL), ws.Cells(bLastDataRow, B_FIRST_COL)).Locked = True
    ws.Range(ws.Cells(B_FIRST_DATA_ROW, B_FIRST_COL + 1), ws.Cells(bLastDataRow, bLastCol)).Locked = False

    ' ===== 2.1) rmaq =====
    With ws.Cells(RMAQ_TITLE_ROW, B_FIRST_COL)
        .Value = "2.1. Escribe la fecha de disponibilidad de la máquina"
        .Font.Bold = True
        .Font.Italic = True
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    With ws.Cells(RMAQ_TABLE_ROW, RMAQ_LABEL_COL)
        .Value = "rmaq"
        .Font.Bold = False
        .Font.Italic = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = True
    End With

    With ws.Cells(RMAQ_TABLE_ROW, RMAQ_INPUT_COL)
        .ClearContents
        .Font.Bold = False
        .Font.Italic = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
        .Interior.Color = RGB(255, 255, 0)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = False
    End With

    ' ===== 3) SECUENCIA (título) =====
    With ws.Cells(seqTitleRow, B_FIRST_COL)
        .Value = "3. Secuencia (escribe el orden en que se realizarán los jobs)."
        .Font.Bold = True
        .Font.Italic = True
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = True
    End With

    ' ===== TABLA A =====
    ws.Cells(A_HEADER_ROW, A_FIRST_COL + 0).Value = "Secuencia"
    ws.Cells(A_HEADER_ROW, A_FIRST_COL + 1).Value = "Job"

    With ws.Range(ws.Cells(A_HEADER_ROW, A_FIRST_COL), ws.Cells(A_HEADER_ROW, aLastCol))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
        .Locked = True
    End With

    With ws.Range(ws.Cells(A_FIRST_DATA_ROW, A_FIRST_COL), ws.Cells(aLastDataRow, aLastCol))
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

    For i = 1 To n
        ws.Cells(A_FIRST_DATA_ROW + i - 1, A_FIRST_COL + 0).Value = i
    Next i

    ' ===== Instrucción 4 =====
    With ws.Cells(instr4Row, B_FIRST_COL)
        .Value = "4. Presione el botón Cargar datos para validar la información ingresada."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Locked = True
    End With

    ' --- Dropdown (Job) en Tabla A ---
    Dim colJobA As Long: colJobA = A_FIRST_COL + 1
    On Error Resume Next
    ws.Range(ws.Cells(1, colJobA), ws.Cells(2000, colJobA)).Validation.Delete
    On Error GoTo ErrH

    Dim jobListAddr As String
    jobListAddr = "=" & ws.Range(ws.Cells(B_FIRST_DATA_ROW, B_FIRST_COL), ws.Cells(bLastDataRow, B_FIRST_COL)).Address(True, True)

    Dim rngDV As Range
    Set rngDV = ws.Range(ws.Cells(A_FIRST_DATA_ROW, colJobA), ws.Cells(aLastDataRow, colJobA))

    rngDV.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=jobListAddr
    rngDV.Validation.IgnoreBlank = True
    rngDV.Validation.InCellDropdown = True

Salir:
    ws.Range(CELL_N).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub

ErrH:
    MsgBox "Error en RedibujarInputs_SM: " & Err.Description, vbExclamation
    Resume Salir
End Sub

' =========================================================
' BOTÓN 1: CARGAR DATOS (VALIDAR) + CARGAR A CACHE + INSTRUCCIÓN 5
' =========================================================
Public Sub SM_CargarDatos()
On Error GoTo ErrH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Dim n As Long
    If Not IsNumeric(ws.Range(CELL_N).Value) Then Exit Sub
    n = CLng(ws.Range(CELL_N).Value)
    If n <= 0 Then Exit Sub

    ws.Unprotect

    ' Recalcular layout igual que RedibujarInputs_SM
    Dim bLastDataRow As Long: bLastDataRow = B_FIRST_DATA_ROW + n - 1
    Dim RMAQ_TITLE_ROW As Long: RMAQ_TITLE_ROW = bLastDataRow + 2
    Dim RMAQ_TABLE_ROW As Long: RMAQ_TABLE_ROW = RMAQ_TITLE_ROW + 2
    Dim seqTitleRow As Long: seqTitleRow = RMAQ_TABLE_ROW + 3
    Dim A_HEADER_ROW As Long: A_HEADER_ROW = seqTitleRow + 1
    Dim A_FIRST_DATA_ROW As Long: A_FIRST_DATA_ROW = A_HEADER_ROW + 1
    Dim aLastDataRow As Long: aLastDataRow = A_FIRST_DATA_ROW + n - 1

    Dim instr4Row As Long: instr4Row = aLastDataRow + 3
    Dim instr5Row As Long: instr5Row = instr4Row + 2

    ' limpiar instrucción 5 + flag
    ws.Range(CELL_SM_READY).Value = ""
    With ws.Cells(instr5Row, B_FIRST_COL)
        .ClearContents
        .Font.Bold = False
        .Font.Italic = False
    End With

    Dim warn As String
    If Not ValidateInputs_Separated(ws, n, warn) Then
        MsgBox warn, vbExclamation, "Revisar datos - Single Machine"
        GoTo Salir
    End If

    ' Cargar datos a variables del módulo
    SM_LoadInputsToCache ws, n

    ws.Range(CELL_SM_READY).Value = "OK"
    ws.Range(CELL_SM_READY).Locked = True

    With ws.Cells(instr5Row, B_FIRST_COL)
        .Value = "5. Datos válidos. Presione el botón Generar outputs."
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Locked = True
    End With

Salir:
    ws.Range(CELL_N).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub

ErrH:
    MsgBox "Error en SM_CargarDatos: " & Err.Description, vbExclamation
    Resume Salir
End Sub

' =========================================================
' VALIDACIÓN (B + rmaq + A)
'   - w ES OBLIGATORIO (default 1, pero si el usuario lo borra/0 => error)
' =========================================================
Private Function ValidateInputs_Separated(ws As Worksheet, ByVal n As Long, ByRef warn As String) As Boolean
    ValidateInputs_Separated = False
    warn = ""

    Dim dictB As Object: Set dictB = CreateObject("Scripting.Dictionary")
    Dim i As Long, shRow As Long, jb As String

    ' ---- Tabla B ----
    For i = 1 To n
        shRow = B_FIRST_DATA_ROW + i - 1

        jb = UCase$(Trim$(CStr(ws.Cells(shRow, B_FIRST_COL + 0).Value)))
        If Len(jb) = 0 Then
            warn = "Tabla B: falta Job en fila " & shRow & "."
            Exit Function
        End If
        If dictB.Exists(jb) Then
            warn = "Tabla B: Job repetido '" & jb & "'."
            Exit Function
        End If
        dictB.Add jb, True

        Dim rv As Double, pV As Double, sV As Double, dV As Double, wv As Double
        rv = NzD(ws.Cells(shRow, B_FIRST_COL + 1).Value)
        pV = NzD(ws.Cells(shRow, B_FIRST_COL + 2).Value)
        sV = NzD(ws.Cells(shRow, B_FIRST_COL + 3).Value)
        dV = NzD(ws.Cells(shRow, B_FIRST_COL + 4).Value)
        wv = NzD(ws.Cells(shRow, B_FIRST_COL + 5).Value)

        If rv < 0 Or sV < 0 Then
            warn = "Tabla B: r y s no pueden ser negativos (Job " & jb & ")."
            Exit Function
        End If
        If pV <= 0 Then
            warn = "Tabla B: p debe ser > 0 (Job " & jb & ")."
            Exit Function
        End If
        If dV <= 0 Then
            warn = "Tabla B: d debe ser > 0 (Job " & jb & ")."
            Exit Function
        End If
        If wv <= 0 Then
            warn = "Tabla B: w debe ser > 0 (Job " & jb & ")."
            Exit Function
        End If
    Next i

    ' ---- rmaq ----
    Dim bLastDataRow As Long: bLastDataRow = B_FIRST_DATA_ROW + n - 1
    Dim RMAQ_TITLE_ROW As Long: RMAQ_TITLE_ROW = bLastDataRow + 2
    Dim RMAQ_TABLE_ROW As Long: RMAQ_TABLE_ROW = RMAQ_TITLE_ROW + 2

    Dim vRmaq As Variant
    vRmaq = ws.Cells(RMAQ_TABLE_ROW, B_FIRST_COL + 1).Value  ' C

    If IsEmpty(vRmaq) Or Trim$(CStr(vRmaq)) = "" Then
        warn = "Falta rmaq (disponibilidad de la máquina)."
        Exit Function
    End If
    If Not IsNumeric(vRmaq) Then
        warn = "rmaq debe ser numérico."
        Exit Function
    End If
    If CDbl(vRmaq) < 0 Then
        warn = "rmaq no puede ser negativo."
        Exit Function
    End If

    ' ---- Tabla A ----
    Dim seqTitleRow As Long: seqTitleRow = RMAQ_TABLE_ROW + 3
    Dim A_HEADER_ROW As Long: A_HEADER_ROW = seqTitleRow + 1
    Dim A_FIRST_DATA_ROW As Long: A_FIRST_DATA_ROW = A_HEADER_ROW + 1
    Dim A_FIRST_COL As Long: A_FIRST_COL = 2

    Dim dictSeq As Object: Set dictSeq = CreateObject("Scripting.Dictionary")
    Dim dictAJobs As Object: Set dictAJobs = CreateObject("Scripting.Dictionary")

    For i = 1 To n
        shRow = A_FIRST_DATA_ROW + i - 1

        Dim vSeq As Variant: vSeq = ws.Cells(shRow, A_FIRST_COL + 0).Value
        Dim vJob As String: vJob = UCase$(Trim$(CStr(ws.Cells(shRow, A_FIRST_COL + 1).Value)))

        If IsEmpty(vSeq) Or Trim$(CStr(vSeq)) = "" Then
            warn = "Tabla A: falta Secuencia en fila " & shRow & "."
            Exit Function
        End If
        If Not IsNumeric(vSeq) Then
            warn = "Tabla A: Secuencia debe ser numérica (fila " & shRow & ")."
            Exit Function
        End If
        If CLng(vSeq) <> CDbl(vSeq) Or CLng(vSeq) <= 0 Then
            warn = "Tabla A: Secuencia debe ser entero positivo (fila " & shRow & ")."
            Exit Function
        End If

        If dictSeq.Exists(CStr(CLng(vSeq))) Then
            warn = "Tabla A: Secuencia repetida (" & CLng(vSeq) & ")."
            Exit Function
        End If
        dictSeq.Add CStr(CLng(vSeq)), True

        If Len(vJob) = 0 Then
            warn = "Tabla A: falta Job (elige en el dropdown) en fila " & shRow & "."
            Exit Function
        End If
        If Not dictB.Exists(vJob) Then
            warn = "Tabla A: el Job '" & vJob & "' no existe en Tabla B."
            Exit Function
        End If
        If dictAJobs.Exists(vJob) Then
            warn = "Tabla A: Job repetido en el orden ('" & vJob & "')."
            Exit Function
        End If
        dictAJobs.Add vJob, True
    Next i

    Dim need As Long
    For need = 1 To n
        If Not dictSeq.Exists(CStr(need)) Then
            warn = "Tabla A: la secuencia debe ser 1.." & n & " sin saltos. Falta: " & need & "."
            Exit Function
        End If
    Next need

    ValidateInputs_Separated = True
End Function

' =========================================================
' BOTÓN 2: GENERAR OUTPUTS (Gantt + Indicadores)
'   - Requiere que SM_CargarDatos haya dejado CELL_SM_READY="OK"
' =========================================================
Public Sub SM_GenerarOutputs()
On Error GoTo ErrH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    If UCase$(Trim$(CStr(ws.Range(CELL_SM_READY).Value))) <> "OK" Then
        MsgBox "Primero presione 'Cargar datos' y asegúrese de que los datos sean válidos.", vbExclamation
        Exit Sub
    End If

    Dim n As Long
    If Not IsNumeric(ws.Range(CELL_N).Value) Then Exit Sub
    n = CLng(ws.Range(CELL_N).Value)
    If n <= 0 Then Exit Sub

    ws.Unprotect

    ' Layout (misma lógica que inputs)
    Dim bLastDataRow As Long: bLastDataRow = B_FIRST_DATA_ROW + n - 1
    Dim RMAQ_TITLE_ROW As Long: RMAQ_TITLE_ROW = bLastDataRow + 2
    Dim RMAQ_TABLE_ROW As Long: RMAQ_TABLE_ROW = RMAQ_TITLE_ROW + 2

    Dim seqTitleRow As Long: seqTitleRow = RMAQ_TABLE_ROW + 3
    Dim A_HEADER_ROW As Long: A_HEADER_ROW = seqTitleRow + 1
    Dim A_FIRST_DATA_ROW As Long: A_FIRST_DATA_ROW = A_HEADER_ROW + 1
    Dim aLastDataRow As Long: aLastDataRow = A_FIRST_DATA_ROW + n - 1

    Dim instr4Row As Long: instr4Row = aLastDataRow + 3
    Dim instr5Row As Long: instr5Row = instr4Row + 2

    ' Coordenadas outputs
    Dim outLineRow As Long: outLineRow = instr5Row + 1
    Dim outTitleRow As Long: outTitleRow = outLineRow + 2
    Dim ganttTopRow As Long: ganttTopRow = outTitleRow + 2
    Dim indTopRow As Long: indTopRow = ganttTopRow + 14 + 3

    ' Limpiar outputs viejos + chart viejo
    DeleteChartIfExists ws, CHART_NAME_TIMELINE

    Dim lastClearRow As Long: lastClearRow = outLineRow + 400
    With ws.Range(ws.Cells(outLineRow, 1), ws.Cells(lastClearRow, 35)) ' A..AI
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

    ' Raya horizontal A..AI
    With ws.Range(ws.Cells(outLineRow, 1), ws.Cells(outLineRow, 35)) ' A..AI
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .Weight = xlThin
        End With
    End With

    ' Título outputs
    With ws.Cells(outTitleRow, B_FIRST_COL)
        .Value = "ZONA DE OUTPUTS"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Locked = True
    End With

    ' Crear ChartObject (vacío) y luego lo llenamos desde cache
    Dim chObj As ChartObject
    Set chObj = ws.ChartObjects.Add( _
        Left:=ws.Cells(ganttTopRow, 2).Left, _
        top:=ws.Cells(ganttTopRow, 2).top, _
        Width:=1100, Height:=240)
    chObj.name = CHART_NAME_TIMELINE

    With chObj.Chart
        .ChartType = xlBarStacked
        .HasTitle = True
        .ChartTitle.text = "Gantt Single Machine"
        .HasLegend = False
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
    End With

    ' Indicadores (estructura) en B..E
    With ws.Cells(indTopRow, B_FIRST_COL)
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
        With ws.Range(ws.Cells(indTopRow + 1 + i, B_FIRST_COL), ws.Cells(indTopRow + 1 + i, B_FIRST_COL + 2))
            On Error Resume Next
            .UnMerge
            On Error GoTo ErrH
            .Merge
            .Value = labels(i)
            .Font.Bold = False
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Locked = True
        End With

        With ws.Cells(indTopRow + 1 + i, B_FIRST_COL + 3) ' col E
            .ClearContents
            .NumberFormat = "General"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Locked = True
        End With
    Next i

    With ws.Range(ws.Cells(indTopRow + 1, B_FIRST_COL), ws.Cells(indTopRow + 1 + UBound(labels), B_FIRST_COL + 3))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
    End With

    ws.Cells(indTopRow + 7, B_FIRST_COL + 3).NumberFormat = "0%"
    ws.Cells(indTopRow + 8, B_FIRST_COL + 3).NumberFormat = "0%"

    ' Indicadores por job (H..N)
    Dim jmFirstCol As Long: jmFirstCol = 8  ' H
    Dim jmNumCols As Long: jmNumCols = 7    ' H..N
    Dim jmHeaderRow As Long: jmHeaderRow = indTopRow
    Dim jmFirstDataRow As Long: jmFirstDataRow = jmHeaderRow + 1
    Dim jmLastRow As Long: jmLastRow = jmFirstDataRow + n - 1

    With ws.Cells(jmHeaderRow, jmFirstCol)
        .Value = "Indicadores por job"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Locked = True
    End With

    Dim jmHeaders As Variant
    jmHeaders = Array("Job", "Inicio", "Cj", "Flow (Cj-rj)", "L (Cj-dj)", "T=max(L,0)", "w*T")

    Dim j As Long
    For j = 0 To UBound(jmHeaders)
        With ws.Cells(jmFirstDataRow, jmFirstCol + j)
            .Value = jmHeaders(j)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(230, 230, 230)
            .Locked = True
        End With
    Next j

    With ws.Range(ws.Cells(jmFirstDataRow + 1, jmFirstCol), ws.Cells(jmLastRow + 1, jmFirstCol + jmNumCols - 1))
        .ClearContents
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
        .Interior.Pattern = xlNone
        .Locked = True
    End With

    With ws.Range(ws.Cells(jmFirstDataRow, jmFirstCol), ws.Cells(jmLastRow + 1, jmFirstCol + jmNumCols - 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
    End With

    If Not smLoaded Or smN <> n Then
        MsgBox "Los datos no están cargados en memoria. Presione 'Cargar datos' nuevamente.", vbExclamation
        GoTo Salir
    End If

    Dim makespan As Double
    SM_CalcAndWriteOutputs ws, n, indTopRow, jmFirstCol, jmFirstDataRow, makespan
    SM_BuildGanttFromCache ws, chObj.Chart, n, makespan

Salir:
    ws.Range(CELL_N).Locked = False
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    Exit Sub

ErrH:
    MsgBox "Error en SM_GenerarOutputs: " & Err.Description, vbExclamation
    Resume Salir
End Sub

' =========================================================
' Cargar inputs a cache (después de validar)
' =========================================================
Private Sub SM_LoadInputsToCache(ByVal ws As Worksheet, ByVal n As Long)

    smLoaded = False
    smN = n

    Dim bLastDataRow As Long: bLastDataRow = B_FIRST_DATA_ROW + n - 1
    Dim RMAQ_TITLE_ROW As Long: RMAQ_TITLE_ROW = bLastDataRow + 2
    Dim RMAQ_TABLE_ROW As Long: RMAQ_TABLE_ROW = RMAQ_TITLE_ROW + 2

    Dim seqTitleRow As Long: seqTitleRow = RMAQ_TABLE_ROW + 3
    Dim A_HEADER_ROW As Long: A_HEADER_ROW = seqTitleRow + 1
    Dim A_FIRST_DATA_ROW As Long: A_FIRST_DATA_ROW = A_HEADER_ROW + 1

    ' rmaq
    smRmaq = NzD(ws.Cells(RMAQ_TABLE_ROW, B_FIRST_COL + 1).Value) ' C

    ' arrays orden (tabla A)
    ReDim smSeq(1 To n)
    ReDim smJob(1 To n)

    Dim i As Long
    For i = 1 To n
        smSeq(i) = CLng(ws.Cells(A_FIRST_DATA_ROW + i - 1, 2).Value) ' Secuencia en B
        smJob(i) = UCase$(Trim$(CStr(ws.Cells(A_FIRST_DATA_ROW + i - 1, 3).Value))) ' Job en C
    Next i

    ' diccionarios params (tabla B)
    Set smParR = CreateObject("Scripting.Dictionary")
    Set smParP = CreateObject("Scripting.Dictionary")
    Set smParS = CreateObject("Scripting.Dictionary")
    Set smParD = CreateObject("Scripting.Dictionary")
    Set smParW = CreateObject("Scripting.Dictionary")

    Dim jb As String
    For i = 1 To n
        jb = UCase$(Trim$(CStr(ws.Cells(B_FIRST_DATA_ROW + i - 1, B_FIRST_COL + 0).Value)))

        smParR(jb) = NzD(ws.Cells(B_FIRST_DATA_ROW + i - 1, B_FIRST_COL + 1).Value)
        smParP(jb) = NzD(ws.Cells(B_FIRST_DATA_ROW + i - 1, B_FIRST_COL + 2).Value)
        smParS(jb) = NzD(ws.Cells(B_FIRST_DATA_ROW + i - 1, B_FIRST_COL + 3).Value)
        smParD(jb) = NzD(ws.Cells(B_FIRST_DATA_ROW + i - 1, B_FIRST_COL + 4).Value)

        smParW(jb) = NzD(ws.Cells(B_FIRST_DATA_ROW + i - 1, B_FIRST_COL + 5).Value)
    Next i

    smLoaded = True
End Sub

' =========================================================
' Cálculo + escritura de outputs
' =========================================================
Private Sub SM_CalcAndWriteOutputs( _
    ByVal ws As Worksheet, ByVal n As Long, _
    ByVal indTopRow As Long, _
    ByVal jmFirstCol As Long, ByVal jmFirstDataRow As Long, _
    ByRef makespan As Double)

    SM_SortBySequence smSeq, smJob, n

    Dim startT() As Double, Cj() As Double, flowT() As Double, Lj() As Double, tJ() As Double, wTj() As Double
    ReDim startT(1 To n)
    ReDim Cj(1 To n)
    ReDim flowT(1 To n)
    ReDim Lj(1 To n)
    ReDim tJ(1 To n)
    ReDim wTj(1 To n)

    Dim clock As Double
    clock = smRmaq

    Dim i As Long, jb As String
    Dim rJ As Double, pJ As Double, sJ As Double, dJ As Double, wJ As Double

    For i = 1 To n
        jb = smJob(i)

        rJ = CDbl(smParR(jb))
        pJ = CDbl(smParP(jb))
        sJ = CDbl(smParS(jb))
        dJ = CDbl(smParD(jb))
        wJ = CDbl(smParW(jb))

        If rJ > clock Then clock = rJ
        startT(i) = clock

        clock = clock + sJ + pJ
        Cj(i) = clock

        flowT(i) = Cj(i) - rJ
        Lj(i) = Cj(i) - dJ
        If Lj(i) > 0# Then tJ(i) = Lj(i) Else tJ(i) = 0#
        wTj(i) = wJ * tJ(i)
    Next i

    makespan = clock

    Dim sumFlow As Double, sumT As Double, sumWT As Double
    Dim Lmax As Double: Lmax = -1E+30
    Dim lateCount As Long

    For i = 1 To n
        sumFlow = sumFlow + flowT(i)
        sumT = sumT + tJ(i)
        sumWT = sumWT + wTj(i)
        If Lj(i) > Lmax Then Lmax = Lj(i)
        If tJ(i) > 0.000001 Then lateCount = lateCount + 1
    Next i

    Dim avgFlow As Double: avgFlow = sumFlow / CDbl(n)
    Dim avgT As Double: avgT = sumT / CDbl(n)
    Dim pctLate As Double: pctLate = lateCount / CDbl(n)
    Dim pctOnTime As Double: pctOnTime = 1# - pctLate

    ' Indicadores (valores en E)
    ws.Cells(indTopRow + 1, B_FIRST_COL + 3).Value = makespan
    ws.Cells(indTopRow + 2, B_FIRST_COL + 3).Value = avgFlow
    ws.Cells(indTopRow + 3, B_FIRST_COL + 3).Value = Lmax
    ws.Cells(indTopRow + 4, B_FIRST_COL + 3).Value = avgT
    ws.Cells(indTopRow + 5, B_FIRST_COL + 3).Value = sumWT
    ws.Cells(indTopRow + 6, B_FIRST_COL + 3).Value = lateCount
    ws.Cells(indTopRow + 7, B_FIRST_COL + 3).Value = pctLate
    ws.Cells(indTopRow + 8, B_FIRST_COL + 3).Value = pctOnTime

    ws.Cells(indTopRow + 7, B_FIRST_COL + 3).NumberFormat = "0%"
    ws.Cells(indTopRow + 8, B_FIRST_COL + 3).NumberFormat = "0%"

    ' Indicadores por job: datos arrancan una fila debajo del header de columnas
    Dim row0 As Long: row0 = jmFirstDataRow + 1
    For i = 1 To n
        ws.Cells(row0 + i - 1, jmFirstCol + 0).Value = smJob(i)
        ws.Cells(row0 + i - 1, jmFirstCol + 1).Value = startT(i)
        ws.Cells(row0 + i - 1, jmFirstCol + 2).Value = Cj(i)
        ws.Cells(row0 + i - 1, jmFirstCol + 3).Value = flowT(i)
        ws.Cells(row0 + i - 1, jmFirstCol + 4).Value = Lj(i)
        ws.Cells(row0 + i - 1, jmFirstCol + 5).Value = tJ(i)
        ws.Cells(row0 + i - 1, jmFirstCol + 6).Value = wTj(i)
    Next i
End Sub

' =========================================================
' Gantt desde cache (una sola línea)
' =========================================================
Private Sub SM_BuildGanttFromCache(ByVal ws As Worksheet, ByVal ch As Chart, ByVal n As Long, ByVal makespan As Double)

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ch.ChartType = xlBarStacked
    ch.HasLegend = False
    ch.HasTitle = True
    ch.ChartTitle.text = "Gantt Single Machine"

    SM_SortBySequence smSeq, smJob, n

    Dim clock As Double
    clock = smRmaq

    Dim i As Long, jb As String
    Dim gap As Double, rJ As Double, pJ As Double, sJ As Double
    Dim srs As Series

    ' Gap inicial por rmaq (invisible)
    If smRmaq > 0# Then
        Set srs = ch.SeriesCollection.NewSeries
        srs.Values = Array(smRmaq)
        srs.XValues = Array("Línea")
        srs.Format.Fill.Visible = msoFalse
        srs.Format.Line.Visible = msoFalse
    End If

    For i = 1 To n
        jb = smJob(i)

        rJ = CDbl(smParR(jb))
        pJ = CDbl(smParP(jb))
        sJ = CDbl(smParS(jb))

        If rJ > clock Then gap = rJ - clock Else gap = 0#

        If gap > 0# Then
            Set srs = ch.SeriesCollection.NewSeries
            srs.Values = Array(gap)
            srs.XValues = Array("Línea")
            srs.Format.Fill.Visible = msoFalse
            srs.Format.Line.Visible = msoFalse
            clock = clock + gap
        End If

        Dim baseColor As Long: baseColor = ColorForJob(i)

        Dim rC As Long, gC As Long, bC As Long
        rC = baseColor And &HFF&
        gC = (baseColor \ &H100&) And &HFF&
        bC = (baseColor \ &H10000) And &HFF&

        Dim lightColor As Long, darkColor As Long
        lightColor = RGB(MinL(rC + 80, 255), MinL(gC + 80, 255), MinL(bC + 80, 255))
        darkColor = RGB(MaxL(rC - 40, 0), MaxL(gC - 40, 0), MaxL(bC - 40, 0))

        ' Setup (claro)
        If sJ > 0# Then
            Set srs = ch.SeriesCollection.NewSeries
            srs.Values = Array(sJ)
            srs.XValues = Array("Línea")
            srs.Format.Fill.ForeColor.RGB = lightColor
            srs.Format.Line.Visible = msoFalse
            clock = clock + sJ
        End If

        ' Proceso (oscuro) + label
        If pJ > 0# Then
            Set srs = ch.SeriesCollection.NewSeries
            srs.name = jb
            srs.Values = Array(pJ)
            srs.XValues = Array("Línea")
            srs.Format.Fill.ForeColor.RGB = darkColor
            srs.Format.Line.Visible = msoFalse

            srs.HasDataLabels = True
            srs.DataLabels.ShowSeriesName = True
            srs.DataLabels.ShowValue = False

            clock = clock + pJ
        End If
    Next i

    ch.Axes(xlCategory).ReversePlotOrder = False
    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).AxisTitle.text = "Tiempo"

    ConfigurarEjeTiempo_Inteligente ch, makespan
End Sub

' =========================================================
' Helpers
' =========================================================
Private Function NzD(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Then NzD = 0# Else NzD = CDbl(v)
End Function

Private Sub DeleteChartIfExists(ws As Worksheet, nm As String)
    On Error Resume Next
    ws.ChartObjects(nm).Delete
    On Error GoTo 0
End Sub

Private Sub SM_SortBySequence(ByRef seq() As Long, ByRef job() As String, ByVal n As Long)
    Dim i As Long, j As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If seq(j) < seq(i) Then
                Dim tL As Long: tL = seq(i): seq(i) = seq(j): seq(j) = tL
                Dim tS As String: tS = job(i): job(i) = job(j): job(j) = tS
            End If
        Next j
    Next i
End Sub

Private Function ColorForJob(idx As Long) As Long
    Dim colors As Variant
    colors = Array( _
        RGB(52, 152, 219), RGB(46, 204, 113), RGB(155, 89, 182), _
        RGB(241, 196, 15), RGB(231, 76, 60), RGB(127, 140, 141), _
        RGB(26, 188, 156))
    ColorForJob = colors((idx - 1) Mod (UBound(colors) + 1))
End Function

Private Function MinL(a As Long, b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function

Private Function MaxL(a As Long, b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

' EJE TIEMPO con cuadrícula mayor + menor (subdivisiones finas)
Private Sub ConfigurarEjeTiempo_Inteligente(ByVal ch As Chart, ByVal maxTime As Double)
    Dim eje As Axis
    Set eje = ch.Axes(xlValue)

    If maxTime < 0.000001 Then maxTime = 1

    Dim stepNice As Double: stepNice = NiceStep_12525(maxTime / 30#)
    If stepNice <= 0# Then stepNice = 1#

    Dim minorStep As Double: minorStep = stepNice / 5#
    If minorStep < 1# And stepNice >= 5# Then minorStep = 1#

    eje.MinimumScale = 0
    eje.MaximumScale = Application.WorksheetFunction.Ceiling(maxTime, stepNice)
    eje.MajorUnit = stepNice
    eje.MinorUnit = minorStep
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

Private Function NiceStep_12525(ByVal x As Double) As Double
    If x <= 0# Then NiceStep_12525 = 1#: Exit Function
    Dim p As Double: p = 10 ^ Int(Log(x) / Log(10))
    Dim m As Double: m = x / p
    Select Case m
        Case Is <= 1#: NiceStep_12525 = 1# * p
        Case Is <= 2#: NiceStep_12525 = 2# * p
        Case Is <= 5#: NiceStep_12525 = 5# * p
        Case Else: NiceStep_12525 = 10# * p
    End Select
End Function

