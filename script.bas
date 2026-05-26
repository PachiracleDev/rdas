Attribute VB_Name = "Module_GradoSalarial"
Option Explicit

'===============================================================================
'  ANÁLISIS DE MOVIMIENTOS DE GRADO SALARIAL 2026
'  Autor: Plantilla generada automáticamente
'  Descripción: Calcula transiciones B3→B2 y B2→B1 por mes y segmento
'===============================================================================

' ── Constantes de columnas (basadas en la hoja DATA) ─────────────────────────
Private Const COL_ANO       As Long = 1   ' Columna A: Año
Private Const COL_MES       As Long = 2   ' Columna B: Mes
Private Const COL_MATRICULA As Long = 5   ' Columna E: Matricula
Private Const COL_GS        As Long = 13  ' Columna M: GS (Grado Salarial)
Private Const COL_GCIA      As Long = 23  ' Columna W: Gcia

' ── Nombres de hojas ──────────────────────────────────────────────────────────
Private Const SH_DATA    As String = "DATA"
Private Const SH_RESUMEN As String = "RESUMEN"
Private Const SH_CONFIG  As String = "CONFIG"

' ── Colores ───────────────────────────────────────────────────────────────────
Private Const CLR_DARK_BLUE    As Long = 6591744    ' RGB(31,56,100)
Private Const CLR_MID_BLUE     As Long = 11953198   ' RGB(46,117,182)
Private Const CLR_LIGHT_ORANGE As Long = 13694460   ' RGB(252,228,214)
Private Const CLR_LIGHT_GREEN  As Long = 9823970    ' RGB(226,239,218)
Private Const CLR_LIGHT_BLUE   As Long = 12506606   ' RGB(189,215,238)
Private Const CLR_GOLD         As Long = 39423      ' RGB(255,215,0)
Private Const CLR_WHITE        As Long = 16777215   ' RGB(255,255,255)
Private Const CLR_YELLOW_NOTE  As Long = 16775372   ' RGB(255,242,204)

'===============================================================================
'  MACRO PRINCIPAL: ACTUALIZAR RESUMEN
'===============================================================================
Public Sub ActualizarResumen()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Procesando datos..."

    Dim wsData    As Worksheet
    Dim wsResumen As Worksheet
    Dim wsCfg     As Worksheet

    On Error GoTo ErrHandler

    Set wsData    = ThisWorkbook.Sheets(SH_DATA)
    Set wsResumen = ThisWorkbook.Sheets(SH_RESUMEN)
    Set wsCfg     = ThisWorkbook.Sheets(SH_CONFIG)

    '── 1. Leer gerencias desde CONFIG ───────────────────────────────────────
    ' Usamos arrays dinámicos para evitar desbordamiento por cantidad de gerencias
    Dim chatsG()    As String
    Dim llamadasG() As String
    Dim nC As Long, nL As Long
    nC = 0: nL = 0
    ReDim chatsG(0)
    ReDim llamadasG(0)

    Dim r As Long
    For r = 4 To 100   ' ampliado por si hay muchas gerencias
        Dim tipo As String
        Dim gcia As String
        tipo = UCase(Trim(wsCfg.Cells(r, 1).Value))
        gcia = UCase(Trim(wsCfg.Cells(r, 2).Value))
        If tipo = "" And gcia = "" Then Exit For
        If gcia = "" Then GoTo NextCfgRow
        If tipo = "CHATS" Then
            If nC = 0 Then
                ReDim chatsG(0)
            Else
                ReDim Preserve chatsG(nC)
            End If
            chatsG(nC) = gcia
            nC = nC + 1
        ElseIf tipo = "LLAMADAS" Then
            If nL = 0 Then
                ReDim llamadasG(0)
            Else
                ReDim Preserve llamadasG(nL)
            End If
            llamadasG(nL) = gcia
            nL = nL + 1
        End If
NextCfgRow:
    Next r

    '── 2. Leer orden de meses desde CONFIG ──────────────────────────────────
    Dim mesOrden(1 To 12) As String
    Dim i As Long
    For i = 1 To 12
        Dim rawMes As String
        rawMes = Trim(CStr(wsCfg.Cells(8 + i, 2).Value))
        mesOrden(i) = rawMes
    Next i

    '── 3. Leer DATA → diccionario (mes|mat) → GS|Gcia ──────────────────────
    Dim dictData     As Object
    Dim mesesPresent As Object
    Set dictData     = CreateObject("Scripting.Dictionary")
    Set mesesPresent = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, COL_MATRICULA).End(xlUp).Row

    Application.StatusBar = "Leyendo " & lastRow & " registros..."

    For r = 2 To lastRow
        Dim ano As String
        ano = Trim(CStr(wsData.Cells(r, COL_ANO).Value))
        If ano <> "2026" Then GoTo NextRow

        Dim mes  As String, mat As String
        Dim gs   As String, gciaCel As String
        mes     = Trim(CStr(wsData.Cells(r, COL_MES).Value))
        mat     = Trim(CStr(wsData.Cells(r, COL_MATRICULA).Value))
        gs      = Trim(CStr(wsData.Cells(r, COL_GS).Value))
        gciaCel = Trim(CStr(wsData.Cells(r, COL_GCIA).Value))

        If mes = "" Or mat = "" Then GoTo NextRow

        Dim key As String
        key = mes & "|" & mat

        If Not dictData.Exists(key) Then
            dictData.Add key, gs & "|" & gciaCel
        End If
        If Not mesesPresent.Exists(mes) Then
            mesesPresent.Add mes, True
        End If
NextRow:
    Next r

    '── 4. Construir lista ordenada de meses con datos ───────────────────────
    Dim mesesOrd()  As String
    Dim nMeses      As Long
    nMeses = 0

    For i = 1 To 12
        If mesOrden(i) <> "" Then
            If mesesPresent.Exists(mesOrden(i)) Then
                ReDim Preserve mesesOrd(nMeses)
                mesesOrd(nMeses) = mesOrden(i)
                nMeses = nMeses + 1
            End If
        End If
    Next i

    If nMeses < 2 Then
        MsgBox "Se necesitan al menos 2 meses de datos." & Chr(10) & _
               "Meses encontrados: " & nMeses, vbInformation, "Sin datos suficientes"
        GoTo Cleanup
    End If

    '── 5. Limpiar RESUMEN (filas desde la 5 hacia abajo) ───────────────────
    Dim lastResRow As Long
    lastResRow = wsResumen.Cells(wsResumen.Rows.Count, 1).End(xlUp).Row
    If lastResRow >= 5 Then
        wsResumen.Rows("5:" & lastResRow + 5).Delete
    End If

    '── 6. Calcular y escribir transiciones ──────────────────────────────────
    Dim writeRow As Long
    writeRow = 5

    Dim totB3B2C As Long, totB3B2L As Long
    Dim totB2B1C As Long, totB2B1L As Long
    totB3B2C = 0: totB3B2L = 0: totB2B1C = 0: totB2B1L = 0

    Dim t As Long
    For t = 0 To nMeses - 2

        Application.StatusBar = "Calculando: " & mesesOrd(t) & " → " & mesesOrd(t + 1)

        Dim mesA As String, mesB As String
        mesA = mesesOrd(t): mesB = mesesOrd(t + 1)

        Dim b3b2C As Long, b3b2L As Long
        Dim b2b1C As Long, b2b1L As Long
        b3b2C = 0: b3b2L = 0: b2b1C = 0: b2b1L = 0

        '-- Iterar colaboradores del mes A --
        Dim kv As Variant
        For Each kv In dictData.Keys
            Dim kParts() As String
            kParts = Split(CStr(kv), "|")
            If UBound(kParts) < 1 Then GoTo NextKey
            If kParts(0) <> mesA Then GoTo NextKey

            Dim matV As String
            matV = kParts(1)

            ' ¿Sigue activo en mes B?
            Dim keyB As String
            keyB = mesB & "|" & matV
            If Not dictData.Exists(keyB) Then GoTo NextKey

            ' GS en A y B
            Dim vA() As String, vB() As String
            Dim valA As String, valB As String
            valA = CStr(dictData(kv))
            valB = CStr(dictData(keyB))
            vA = Split(valA, "|")
            vB = Split(valB, "|")

            If UBound(vA) < 1 Or UBound(vB) < 0 Then GoTo NextKey

            Dim gsA As String, gsB As String, gciaV As String
            gsA   = vA(0)
            gsB   = vB(0)
            gciaV = UCase(vA(1))

            ' Determinar servicio
            Dim servicio As String
            servicio = ""
            Dim c As Long
            If nC > 0 Then
                For c = 0 To nC - 1
                    If InStr(gciaV, chatsG(c)) > 0 Or InStr(chatsG(c), gciaV) > 0 Then
                        servicio = "Chats": Exit For
                    End If
                Next c
            End If
            If servicio = "" And nL > 0 Then
                For c = 0 To nL - 1
                    If InStr(gciaV, llamadasG(c)) > 0 Or InStr(llamadasG(c), gciaV) > 0 Then
                        servicio = "Llamadas": Exit For
                    End If
                Next c
            End If
            If servicio = "" Then GoTo NextKey

            ' Contar transición
            If gsA = "B3" And gsB = "B2" Then
                If servicio = "Chats" Then b3b2C = b3b2C + 1 Else b3b2L = b3b2L + 1
            ElseIf gsA = "B2" And gsB = "B1" Then
                If servicio = "Chats" Then b2b1C = b2b1C + 1 Else b2b1L = b2b1L + 1
            End If
NextKey:
        Next kv

        ' Acumulados
        totB3B2C = totB3B2C + b3b2C: totB3B2L = totB3B2L + b3b2L
        totB2B1C = totB2B1C + b2b1C: totB2B1L = totB2B1L + b2b1L

        ' Escribir bloque
        Dim periodoLbl As String
        periodoLbl = mesA & " -> " & mesB

        WriteHeaderRow wsResumen, writeRow, "  >  " & periodoLbl
        writeRow = writeRow + 1

        WriteDataRow wsResumen, writeRow, periodoLbl, "B3 -> B2", "Chats",    b3b2C, b3b2C + b3b2L, True
        writeRow = writeRow + 1
        WriteDataRow wsResumen, writeRow, periodoLbl, "B3 -> B2", "Llamadas", b3b2L, b3b2C + b3b2L, False
        writeRow = writeRow + 1
        WriteDataRow wsResumen, writeRow, periodoLbl, "B2 -> B1", "Chats",    b2b1C, b2b1C + b2b1L, True
        writeRow = writeRow + 1
        WriteDataRow wsResumen, writeRow, periodoLbl, "B2 -> B1", "Llamadas", b2b1L, b2b1C + b2b1L, False
        writeRow = writeRow + 1

    Next t

    '── 7. Totales acumulados ─────────────────────────────────────────────────
    WriteHeaderRow wsResumen, writeRow, "  >  TOTALES ACUMULADOS 2026"
    writeRow = writeRow + 1
    WriteDataRow wsResumen, writeRow, "TOTAL 2026", "B3 -> B2", "Chats",    totB3B2C, totB3B2C + totB3B2L, True
    writeRow = writeRow + 1
    WriteDataRow wsResumen, writeRow, "TOTAL 2026", "B3 -> B2", "Llamadas", totB3B2L, totB3B2C + totB3B2L, False
    writeRow = writeRow + 1
    WriteDataRow wsResumen, writeRow, "TOTAL 2026", "B2 -> B1", "Chats",    totB2B1C, totB2B1C + totB2B1L, True
    writeRow = writeRow + 1
    WriteDataRow wsResumen, writeRow, "TOTAL 2026", "B2 -> B1", "Llamadas", totB2B1L, totB2B1C + totB2B1L, False
    writeRow = writeRow + 1

    '── 8. Timestamp ─────────────────────────────────────────────────────────
    Dim tsRow As Long
    tsRow = writeRow + 1
    Merge_Cells_Safe wsResumen, tsRow, 1, tsRow, 8
    Dim tsCell As Range
    Set tsCell = wsResumen.Cells(tsRow, 1)
    tsCell.Value = "Ultima actualizacion: " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & _
                   "  |  Meses procesados: " & nMeses & _
                   "  |  Transiciones calculadas: " & (nMeses - 1)
    tsCell.Font.Italic = True
    tsCell.Font.Size = 8
    tsCell.Font.Color = RGB(150, 150, 150)
    tsCell.Font.Name = "Arial"

    wsResumen.Activate
    wsResumen.Cells(1, 1).Select

    MsgBox "Resumen actualizado correctamente." & Chr(10) & Chr(10) & _
           "  Meses procesados: " & nMeses & Chr(10) & _
           "  Transiciones calculadas: " & (nMeses - 1) & Chr(10) & Chr(10) & _
           "  B3->B2 Chats: " & totB3B2C & "   Llamadas: " & totB3B2L & Chr(10) & _
           "  B2->B1 Chats: " & totB2B1C & "   Llamadas: " & totB2B1L, _
           vbInformation, "Analisis Grado Salarial 2026"

Cleanup:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description & Chr(10) & _
           "En linea: " & Erl, vbCritical, "Error en macro"
End Sub

'===============================================================================
'  HELPER: ESCRIBIR FILA ENCABEZADO DE PERIODO
'===============================================================================
Private Sub WriteHeaderRow(ws As Worksheet, rowNum As Long, label As String)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8))
    rng.Merge
    rng.Cells(1, 1).Value = label
    rng.Interior.Color = CLR_DARK_BLUE
    rng.Font.Bold = True
    rng.Font.Color = CLR_GOLD
    rng.Font.Size = 11
    rng.Font.Name = "Arial"
    rng.HorizontalAlignment = xlLeft
    rng.VerticalAlignment = xlCenter
    rng.IndentLevel = 1
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = CLR_MID_BLUE
    rng.Borders.Weight = xlMedium
    ws.Rows(rowNum).RowHeight = 26
End Sub

'===============================================================================
'  HELPER: ESCRIBIR FILA DE DATO
'===============================================================================
Private Sub WriteDataRow(ws As Worksheet, rowNum As Long, _
                          periodo As String, transicion As String, _
                          segmento As String, valor As Long, _
                          total As Long, isChats As Boolean)
    Dim bg As Long
    bg = IIf(isChats, CLR_LIGHT_ORANGE, CLR_LIGHT_GREEN)

    Dim pct As Double
    If total > 0 Then
        pct = CDbl(valor) / CDbl(total)
    Else
        pct = 0
    End If
    ' Asegurar que pct esté entre 0 y 1
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1

    ' Barra visual proporcional (caracteres ASCII para evitar Error 5 con Unicode)
    Dim barLen As Long
    barLen = CLng(Int(pct * 15))
    If barLen < 0 Then barLen = 0
    If barLen > 15 Then barLen = 15

    Dim bar As String
    Dim bi As Long
    bar = ""
    For bi = 1 To barLen
        bar = bar & "|"
    Next bi
    For bi = barLen + 1 To 15
        bar = bar & "-"
    Next bi

    Dim values(1 To 7) As Variant
    values(1) = periodo
    values(2) = transicion
    values(3) = segmento
    values(4) = valor
    values(5) = total
    values(6) = pct
    values(7) = bar

    Dim col As Long
    For col = 1 To 7
        Dim cell As Range
        Set cell = ws.Cells(rowNum, col)
        cell.Value = values(col)
        cell.Interior.Color = bg
        cell.Font.Name = "Arial"
        cell.Font.Size = 10
        cell.Font.Bold = False
        cell.VerticalAlignment = xlCenter
        cell.HorizontalAlignment = IIf(col <= 3, xlLeft, xlCenter)

        Dim brd As Border
        For Each brd In cell.Borders
            brd.LineStyle = xlContinuous
            brd.Color = RGB(170, 170, 170)
            brd.Weight = xlThin
        Next brd

        If col = 4 Or col = 5 Then cell.NumberFormat = "#,##0"
        If col = 6 Then cell.NumberFormat = "0.0%"
        If col = 7 Then
            cell.Font.Size = 8
            cell.Font.Color = RGB(46, 117, 182)
            cell.HorizontalAlignment = xlLeft
        End If
    Next col

    ws.Rows(rowNum).RowHeight = 20
End Sub

'===============================================================================
'  HELPER: Merge seguro (evita error si ya está mergeado)
'===============================================================================
Private Sub Merge_Cells_Safe(ws As Worksheet, r1 As Long, c1 As Long, _
                               r2 As Long, c2 As Long)
    On Error Resume Next
    ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Merge
    On Error GoTo 0
End Sub

'===============================================================================
'  MACRO: ACTUALIZAR DESVINCULACIONES POR MES
'  Lógica: colaborador presente en mes N pero ausente en mes N+1 = desvinculado
'  Segmenta por Chats y Llamadas usando CONFIG (igual que ActualizarResumen)
'===============================================================================
Public Sub ActualizarDesvinculaciones()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Procesando desvinculaciones..."

    Dim wsData  As Worksheet
    Dim wsDesv  As Worksheet
    Dim wsCfg   As Worksheet

    On Error GoTo ErrHandlerDesv

    Set wsData = ThisWorkbook.Sheets(SH_DATA)
    Set wsCfg  = ThisWorkbook.Sheets(SH_CONFIG)

    ' Crear hoja si no existe
    Dim shName As String
    shName = "Desvinculaciones por mes"
    On Error Resume Next
    Set wsDesv = ThisWorkbook.Sheets(shName)
    On Error GoTo ErrHandlerDesv
    If wsDesv Is Nothing Then
        Set wsDesv = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("RESUMEN"))
        wsDesv.Name = shName
    End If

    '── 1. Leer gerencias desde CONFIG (misma lógica que ActualizarResumen) ──
    Dim chatsGD()    As String
    Dim llamadasGD() As String
    Dim nCD As Long, nLD As Long
    nCD = 0: nLD = 0
    ReDim chatsGD(0)
    ReDim llamadasGD(0)

    Dim r As Long
    For r = 4 To 100
        Dim tipoD As String, gciaD As String
        tipoD = UCase(Trim(wsCfg.Cells(r, 1).Value))
        gciaD = UCase(Trim(wsCfg.Cells(r, 2).Value))
        If tipoD = "" And gciaD = "" Then Exit For
        If gciaD = "" Then GoTo NextCfgD
        If tipoD = "CHATS" Then
            If nCD > 0 Then ReDim Preserve chatsGD(nCD)
            chatsGD(nCD) = gciaD
            nCD = nCD + 1
        ElseIf tipoD = "LLAMADAS" Then
            If nLD > 0 Then ReDim Preserve llamadasGD(nLD)
            llamadasGD(nLD) = gciaD
            nLD = nLD + 1
        End If
NextCfgD:
    Next r

    '── 2. Leer orden de meses desde CONFIG ──────────────────────────────────
    Dim mesOrdenD(1 To 12) As String
    Dim i As Long
    For i = 1 To 12
        mesOrdenD(i) = Trim(CStr(wsCfg.Cells(8 + i, 2).Value))
    Next i

    '── 3. Leer DATA: (mes|mat) → gcia ──────────────────────────────────────
    '    Un colaborador se considera de un segmento según su Gcia
    Dim dictMat      As Object   ' key="mes|mat"  value="gcia"
    Dim mesesPresD   As Object   ' meses con datos
    Set dictMat    = CreateObject("Scripting.Dictionary")
    Set mesesPresD = CreateObject("Scripting.Dictionary")

    Dim lastRowD As Long
    lastRowD = wsData.Cells(wsData.Rows.Count, COL_MATRICULA).End(xlUp).Row

    Application.StatusBar = "Leyendo " & lastRowD & " registros para desvinculaciones..."

    For r = 2 To lastRowD
        Dim anoD As String
        anoD = Trim(CStr(wsData.Cells(r, COL_ANO).Value))
        If anoD <> "2026" Then GoTo NextRowD

        Dim mesD As String, matD As String, gciaCell As String
        mesD     = Trim(CStr(wsData.Cells(r, COL_MES).Value))
        matD     = Trim(CStr(wsData.Cells(r, COL_MATRICULA).Value))
        gciaCell = Trim(CStr(wsData.Cells(r, COL_GCIA).Value))

        If mesD = "" Or matD = "" Then GoTo NextRowD

        Dim keyD As String
        keyD = mesD & "|" & matD
        If Not dictMat.Exists(keyD) Then
            dictMat.Add keyD, UCase(gciaCell)
        End If
        If Not mesesPresD.Exists(mesD) Then
            mesesPresD.Add mesD, True
        End If
NextRowD:
    Next r

    '── 4. Lista ordenada de meses con datos ─────────────────────────────────
    Dim mesesOrdD() As String
    Dim nMesesD     As Long
    nMesesD = 0
    For i = 1 To 12
        If mesOrdenD(i) <> "" Then
            If mesesPresD.Exists(mesOrdenD(i)) Then
                ReDim Preserve mesesOrdD(nMesesD)
                mesesOrdD(nMesesD) = mesOrdenD(i)
                nMesesD = nMesesD + 1
            End If
        End If
    Next i

    If nMesesD < 2 Then
        MsgBox "Se necesitan al menos 2 meses de datos." & Chr(10) & _
               "Meses encontrados: " & nMesesD, vbInformation, "Sin datos suficientes"
        GoTo CleanupDesv
    End If

    '── 5. Limpiar hoja desde fila 5 hacia abajo ─────────────────────────────
    Dim lastDR As Long
    lastDR = wsDesv.Cells(wsDesv.Rows.Count, 1).End(xlUp).Row
    If lastDR >= 5 Then
        wsDesv.Rows("5:" & lastDR + 2).Delete
    End If

    '── 6. Calcular desvinculaciones por período ──────────────────────────────
    '    Desvinculado = presente en mes A, ausente en mes A+1, del segmento X
    Dim writeRowD As Long
    writeRowD = 5

    Dim totDesvC As Long, totDesvL As Long
    totDesvC = 0: totDesvL = 0

    Dim tD As Long
    For tD = 0 To nMesesD - 2

        Dim mesAD As String, mesBD As String
        mesAD = mesesOrdD(tD)
        mesBD = mesesOrdD(tD + 1)

        Application.StatusBar = "Calculando desvinculaciones: " & mesAD & " → " & mesBD

        Dim desvC As Long, desvL As Long
        desvC = 0: desvL = 0

        ' Recorrer todos los colaboradores del mes A
        Dim kvD As Variant
        For Each kvD In dictMat.Keys
            Dim kpD() As String
            kpD = Split(CStr(kvD), "|")
            If UBound(kpD) < 1 Then GoTo NextKvD
            If kpD(0) <> mesAD Then GoTo NextKvD

            Dim matVD As String
            matVD = kpD(1)

            ' ¿Sigue activo en mes B?
            Dim keyBD As String
            keyBD = mesBD & "|" & matVD
            If dictMat.Exists(keyBD) Then GoTo NextKvD   ' sigue activo → NO es desvinculado

            ' Determinar segmento por su Gcia en mes A
            Dim gciaVD As String
            gciaVD = CStr(dictMat(kvD))

            Dim servicioD As String
            servicioD = ""
            Dim cD As Long
            If nCD > 0 Then
                For cD = 0 To nCD - 1
                    If InStr(gciaVD, chatsGD(cD)) > 0 Or InStr(chatsGD(cD), gciaVD) > 0 Then
                        servicioD = "Chats": Exit For
                    End If
                Next cD
            End If
            If servicioD = "" And nLD > 0 Then
                For cD = 0 To nLD - 1
                    If InStr(gciaVD, llamadasGD(cD)) > 0 Or InStr(llamadasGD(cD), gciaVD) > 0 Then
                        servicioD = "Llamadas": Exit For
                    End If
                Next cD
            End If
            If servicioD = "" Then GoTo NextKvD

            If servicioD = "Chats" Then
                desvC = desvC + 1
            Else
                desvL = desvL + 1
            End If
NextKvD:
        Next kvD

        totDesvC = totDesvC + desvC
        totDesvL = totDesvL + desvL

        ' Escribir fila de datos
        WriteDesvRow wsDesv, writeRowD, mesAD & " → " & mesBD, mesAD, desvC, desvL
        writeRowD = writeRowD + 1

    Next tD

    '── 7. Fila de TOTAL ─────────────────────────────────────────────────────
    WriteDesvRowTotal wsDesv, writeRowD, totDesvC, totDesvL
    writeRowD = writeRowD + 1

    '── 8. Timestamp ─────────────────────────────────────────────────────────
    Dim tsDR As Long
    tsDR = writeRowD + 1
    Merge_Cells_Safe wsDesv, tsDR, 1, tsDR, 7
    With wsDesv.Cells(tsDR, 1)
        .Value = "Ultima actualizacion: " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & _
                 "  |  Meses procesados: " & nMesesD
        .Font.Italic = True
        .Font.Size = 8
        .Font.Color = RGB(150, 150, 150)
        .Font.Name = "Arial"
    End With

    '── 9. Ajustar columnas ───────────────────────────────────────────────────
    wsDesv.Columns("A:G").AutoFit
    If wsDesv.Columns("A").ColumnWidth < 20 Then wsDesv.Columns("A").ColumnWidth = 20
    If wsDesv.Columns("B").ColumnWidth < 14 Then wsDesv.Columns("B").ColumnWidth = 14
    If wsDesv.Columns("C").ColumnWidth < 22 Then wsDesv.Columns("C").ColumnWidth = 22
    If wsDesv.Columns("D").ColumnWidth < 24 Then wsDesv.Columns("D").ColumnWidth = 24
    If wsDesv.Columns("E").ColumnWidth < 12 Then wsDesv.Columns("E").ColumnWidth = 12
    If wsDesv.Columns("F").ColumnWidth < 12 Then wsDesv.Columns("F").ColumnWidth = 12
    If wsDesv.Columns("G").ColumnWidth < 14 Then wsDesv.Columns("G").ColumnWidth = 14

    wsDesv.Activate
    wsDesv.Cells(1, 1).Select

    MsgBox "Desvinculaciones actualizadas correctamente." & Chr(10) & Chr(10) & _
           "  Meses procesados: " & nMesesD & Chr(10) & Chr(10) & _
           "  Total desvinculados Chats:    " & totDesvC & Chr(10) & _
           "  Total desvinculados Llamadas: " & totDesvL & Chr(10) & _
           "  TOTAL GENERAL:               " & (totDesvC + totDesvL), _
           vbInformation, "Desvinculaciones por Mes 2026"

CleanupDesv:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrHandlerDesv:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error en macro desvinculaciones"
End Sub

'===============================================================================
'  HELPER: escribir fila de datos de desvinculacion
'===============================================================================
Private Sub WriteDesvRow(ws As Worksheet, rowNum As Long, _
                          periodo As String, mes As String, _
                          desvChats As Long, desvLlamadas As Long)
    Dim total As Long
    total = desvChats + desvLlamadas

    Dim pctC As Double, pctL As Double
    If total > 0 Then
        pctC = CDbl(desvChats)    / CDbl(total)
        pctL = CDbl(desvLlamadas) / CDbl(total)
    Else
        pctC = 0: pctL = 0
    End If

    Dim vals(1 To 7) As Variant
    vals(1) = periodo
    vals(2) = mes
    vals(3) = desvChats
    vals(4) = desvLlamadas
    vals(5) = total
    vals(6) = pctC
    vals(7) = pctL

    Dim col As Long
    For col = 1 To 7
        Dim cell As Range
        Set cell = ws.Cells(rowNum, col)
        cell.Value = vals(col)
        cell.Font.Name = "Arial"
        cell.Font.Size = 10
        cell.Font.Bold = False
        cell.VerticalAlignment = xlCenter

        ' Fondo alternado: naranja para Chats, verde para Llamadas, azul claro para totales
        If col = 3 Then
            cell.Interior.Color = CLR_LIGHT_ORANGE
        ElseIf col = 4 Then
            cell.Interior.Color = CLR_LIGHT_GREEN
        Else
            cell.Interior.Color = CLR_LIGHT_BLUE
        End If

        cell.HorizontalAlignment = IIf(col <= 2, xlLeft, xlCenter)

        ' Bordes solo perimetrales (sin diagonales)
        Dim s As Long
        Dim sides(3) As Long
        sides(0) = xlEdgeLeft: sides(1) = xlEdgeRight
        sides(2) = xlEdgeTop:  sides(3) = xlEdgeBottom
        For s = 0 To 3
            With cell.Borders(sides(s))
                .LineStyle = xlContinuous
                .Color = RGB(170, 170, 170)
                .Weight = xlThin
            End With
        Next s
        cell.Borders(xlDiagonalDown).LineStyle = xlNone
        cell.Borders(xlDiagonalUp).LineStyle = xlNone

        If col = 3 Or col = 4 Or col = 5 Then cell.NumberFormat = "#,##0"
        If col = 6 Or col = 7 Then cell.NumberFormat = "0.0%"
    Next col

    ws.Rows(rowNum).RowHeight = 20
End Sub

'===============================================================================
'  HELPER: fila TOTAL de desvinculaciones
'===============================================================================
Private Sub WriteDesvRowTotal(ws As Worksheet, rowNum As Long, _
                               totC As Long, totL As Long)
    Dim total As Long
    total = totC + totL

    Dim pctC As Double, pctL As Double
    If total > 0 Then
        pctC = CDbl(totC) / CDbl(total)
        pctL = CDbl(totL) / CDbl(total)
    Else
        pctC = 0: pctL = 0
    End If

    Dim vals(1 To 7) As Variant
    vals(1) = "TOTAL 2026"
    vals(2) = "Acumulado"
    vals(3) = totC
    vals(4) = totL
    vals(5) = total
    vals(6) = pctC
    vals(7) = pctL

    Dim col As Long
    For col = 1 To 7
        Dim cell As Range
        Set cell = ws.Cells(rowNum, col)
        cell.Value = vals(col)
        cell.Font.Name = "Arial"
        cell.Font.Size = 10
        cell.Font.Bold = True
        cell.VerticalAlignment = xlCenter
        cell.Interior.Color = CLR_DARK_BLUE
        cell.Font.Color = CLR_GOLD
        cell.HorizontalAlignment = IIf(col <= 2, xlLeft, xlCenter)

        Dim s As Long
        Dim sides(3) As Long
        sides(0) = xlEdgeLeft: sides(1) = xlEdgeRight
        sides(2) = xlEdgeTop:  sides(3) = xlEdgeBottom
        For s = 0 To 3
            With cell.Borders(sides(s))
                .LineStyle = xlContinuous
                .Color = RGB(46, 117, 182)
                .Weight = xlMedium
            End With
        Next s
        cell.Borders(xlDiagonalDown).LineStyle = xlNone
        cell.Borders(xlDiagonalUp).LineStyle = xlNone

        If col = 3 Or col = 4 Or col = 5 Then cell.NumberFormat = "#,##0"
        If col = 6 Or col = 7 Then cell.NumberFormat = "0.0%"
    Next col

    ws.Rows(rowNum).RowHeight = 24
End Sub

'===============================================================================
'  MACRO: ACTUALIZAR CONTABILIZADO AP Y SUPERVISORES
'  Filtro previo: excluye Canal 2 que contenga "Yape" o "IO" (no case sensitive)
'  Lógica de liderazgo: un AP/Supervisor lidera si su matrícula aparece en la
'  columna "Mat. Jefe actual" de al menos un registro donde ese subordinado
'  tiene "Asesor" en la columna "Tipo de Orgánico"
'===============================================================================
Public Sub ActualizarContabilizado()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Procesando contabilizado AP/Supervisores..."

    Dim wsData As Worksheet
    Dim wsCont As Worksheet
    Dim wsCfg  As Worksheet

    On Error GoTo ErrHandlerCont

    Set wsData = ThisWorkbook.Sheets(SH_DATA)
    Set wsCfg  = ThisWorkbook.Sheets(SH_CONFIG)

    '── Crear / referenciar hoja ─────────────────────────────────────────────
    Dim shCont As String
    shCont = "Contabilizado AP y Supervisores"
    On Error Resume Next
    Set wsCont = ThisWorkbook.Sheets(shCont)
    On Error GoTo ErrHandlerCont
    If wsCont Is Nothing Then
        Set wsCont = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("RESUMEN"))
        wsCont.Name = shCont
    End If
    wsCont.Cells.Clear

    '── Leer orden de meses desde CONFIG (para agrupar por mes) ──────────────
    Dim mesOrdenC(1 To 12) As String
    Dim i As Long
    For i = 1 To 12
        mesOrdenC(i) = Trim(CStr(wsCfg.Cells(8 + i, 2).Value))
    Next i

    '── Leer DATA con filtro Canal 2 ─────────────────────────────────────────
    '
    '  Columnas relevantes:
    '    Col  5  = Matricula
    '    Col  7  = Colaborador (nombre)
    '    Col 16  = Mat. Jefe actual  (a quién reporta)
    '    Col 19  = Canal 2           (filtro exclusión)
    '    Col 20  = Tipo              (AP / SUPERVISOR / etc.)
    '    Col 21  = Tipo de Orgánico  (Asesor = subordinado válido)
    '
    '  Estructuras que construimos:
    '    dictPersonas(mes|mat)  = "TIPO||NOMBRE"
    '       → todos los AP y Supervisores válidos por mes
    '    dictAsesoresJefe(mes|matJefe) = True
    '       → jefes que tienen al menos un Asesor a cargo ese mes
    '

    Dim dictPersonas    As Object   ' key="mes|mat"        val="TIPO||Nombre"
    Dim dictAsesoresJefe As Object  ' key="mes|matJefe"    val=True
    Dim mesesCont       As Object   ' meses con datos
    Set dictPersonas     = CreateObject("Scripting.Dictionary")
    Set dictAsesoresJefe = CreateObject("Scripting.Dictionary")
    Set mesesCont        = CreateObject("Scripting.Dictionary")

    Dim lastRowC As Long
    lastRowC = wsData.Cells(wsData.Rows.Count, COL_MATRICULA).End(xlUp).Row
    Application.StatusBar = "Leyendo " & lastRowC & " registros..."

    Dim excC As Long
    excC = 0

    Dim r As Long
    For r = 2 To lastRowC
        Dim anoC As String
        anoC = Trim(CStr(wsData.Cells(r, COL_ANO).Value))
        If anoC <> "2026" Then GoTo NextRowC

        '── Filtro Canal 2: excluir Yape e IO ────────────────────────────────
        Dim canal2C As String
        canal2C = UCase(Trim(CStr(wsData.Cells(r, 19).Value)))
        If InStr(canal2C, "YAPE") > 0 Then excC = excC + 1: GoTo NextRowC
        If InStr(canal2C, " IO") > 0  Then excC = excC + 1: GoTo NextRowC
        If Left(canal2C, 2) = "IO"    Then excC = excC + 1: GoTo NextRowC

        '── Leer campos clave ────────────────────────────────────────────────
        Dim mesC   As String, matC  As String
        Dim tipoC  As String, nombreC As String
        Dim tipoOrgC As String, matJefeC As String
        mesC      = Trim(CStr(wsData.Cells(r, COL_MES).Value))
        matC      = Trim(CStr(wsData.Cells(r, COL_MATRICULA).Value))
        tipoC     = UCase(Trim(CStr(wsData.Cells(r, 20).Value)))
        nombreC   = Trim(CStr(wsData.Cells(r, 7).Value))
        tipoOrgC  = UCase(Trim(CStr(wsData.Cells(r, 21).Value)))
        matJefeC  = Trim(CStr(wsData.Cells(r, 16).Value))

        If mesC = "" Or matC = "" Then GoTo NextRowC

        '── Registrar AP y Supervisores ──────────────────────────────────────
        If InStr(tipoC, "AP") > 0 Or InStr(tipoC, "SUPERVISOR") > 0 Then
            Dim keyPC As String
            keyPC = mesC & "|" & matC
            If Not dictPersonas.Exists(keyPC) Then
                dictPersonas.Add keyPC, tipoC & "||" & nombreC
            End If
        End If

        '── Registrar jefes con Asesores a cargo ─────────────────────────────
        '   Si este registro tiene "Asesor" en Tipo de Orgánico
        '   → su jefe (matJefeC) lidera al menos un asesor
        If InStr(tipoOrgC, "ASESOR") > 0 Then
            If matJefeC <> "" Then
                Dim keyJC As String
                keyJC = mesC & "|" & matJefeC
                If Not dictAsesoresJefe.Exists(keyJC) Then
                    dictAsesoresJefe.Add keyJC, True
                End If
            End If
        End If

        If Not mesesCont.Exists(mesC) Then mesesCont.Add mesC, True
NextRowC:
    Next r

    '── Construir lista ordenada de meses con datos ───────────────────────────
    Dim mesesOrdC() As String
    Dim nMesesC As Long
    nMesesC = 0
    For i = 1 To 12
        If mesOrdenC(i) <> "" Then
            If mesesCont.Exists(mesOrdenC(i)) Then
                ReDim Preserve mesesOrdC(nMesesC)
                mesesOrdC(nMesesC) = mesOrdenC(i)
                nMesesC = nMesesC + 1
            End If
        End If
    Next i

    '── Formatear encabezado de la hoja ──────────────────────────────────────
    FormatContHeader wsCont, excC

    '── Calcular y escribir por mes ──────────────────────────────────────────
    Dim writeRowC As Long
    writeRowC = 6   ' filas 1-5 = encabezado

    ' Acumulados globales
    Dim totAP As Long, totSup As Long
    Dim totAPLider As Long, totSupLider As Long
    totAP = 0: totSup = 0: totAPLider = 0: totSupLider = 0

    Dim m As Long
    For m = 0 To nMesesC - 1
        Dim mesM As String
        mesM = mesesOrdC(m)
        Application.StatusBar = "Contabilizando: " & mesM

        Dim cntAP As Long, cntSup As Long
        Dim cntAPLid As Long, cntSupLid As Long
        cntAP = 0: cntSup = 0: cntAPLid = 0: cntSupLid = 0

        Dim kv As Variant
        For Each kv In dictPersonas.Keys
            Dim kpC() As String
            kpC = Split(CStr(kv), "|")
            If UBound(kpC) < 1 Then GoTo NextKvC
            If kpC(0) <> mesM Then GoTo NextKvC

            Dim matPersona As String
            matPersona = kpC(1)

            Dim valC() As String
            valC = Split(CStr(dictPersonas(kv)), "||")
            If UBound(valC) < 0 Then GoTo NextKvC

            Dim tipoPersona As String
            tipoPersona = valC(0)

            ' ¿Lidera? → su mes|mat existe en dictAsesoresJefe
            Dim esLider As Boolean
            esLider = dictAsesoresJefe.Exists(mesM & "|" & matPersona)

            If InStr(tipoPersona, "SUPERVISOR") > 0 Then
                cntSup = cntSup + 1
                If esLider Then cntSupLid = cntSupLid + 1
            ElseIf InStr(tipoPersona, "AP") > 0 Then
                cntAP = cntAP + 1
                If esLider Then cntAPLid = cntAPLid + 1
            End If
NextKvC:
        Next kv

        ' Escribir bloque del mes
        WriteContMesHeader wsCont, writeRowC, mesM
        writeRowC = writeRowC + 1
        WriteContDataRow wsCont, writeRowC, "AP", cntAP, False, False
        writeRowC = writeRowC + 1
        WriteContDataRow wsCont, writeRowC, "SUPERVISOR", cntSup, False, True
        writeRowC = writeRowC + 1
        WriteContDataRow wsCont, writeRowC, "AP que lideran un equipo", cntAPLid, True, False
        writeRowC = writeRowC + 1
        WriteContDataRow wsCont, writeRowC, "Supervisores que lideran un equipo", cntSupLid, True, True
        writeRowC = writeRowC + 1
        writeRowC = writeRowC + 1  ' espacio entre meses

        ' Acumular
        totAP = totAP + cntAP: totSup = totSup + cntSup
        totAPLider = totAPLider + cntAPLid: totSupLider = totSupLider + cntSupLid
    Next m

    '── Bloque TOTAL ACUMULADO ────────────────────────────────────────────────
    WriteContMesHeader wsCont, writeRowC, "TOTAL ACUMULADO 2026"
    writeRowC = writeRowC + 1
    WriteContDataRow wsCont, writeRowC, "AP", totAP, False, False
    writeRowC = writeRowC + 1
    WriteContDataRow wsCont, writeRowC, "SUPERVISOR", totSup, False, True
    writeRowC = writeRowC + 1
    WriteContDataRow wsCont, writeRowC, "AP que lideran un equipo", totAPLider, True, False
    writeRowC = writeRowC + 1
    WriteContDataRow wsCont, writeRowC, "Supervisores que lideran un equipo", totSupLider, True, True
    writeRowC = writeRowC + 1

    '── Timestamp ────────────────────────────────────────────────────────────
    Dim tsRC As Long
    tsRC = writeRowC + 1
    Merge_Cells_Safe wsCont, tsRC, 1, tsRC, 4
    With wsCont.Cells(tsRC, 1)
        .Value = "Ultima actualizacion: " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & _
                 "  |  Meses procesados: " & nMesesC & _
                 "  |  Registros excluidos (Yape/IO): " & excC
        .Font.Italic = True: .Font.Size = 8
        .Font.Color = RGB(150, 150, 150): .Font.Name = "Arial"
    End With

    '── Ajustar columnas ─────────────────────────────────────────────────────
    wsCont.Columns("A:D").AutoFit
    If wsCont.Columns("A").ColumnWidth < 38 Then wsCont.Columns("A").ColumnWidth = 38
    If wsCont.Columns("B").ColumnWidth < 14 Then wsCont.Columns("B").ColumnWidth = 14
    If wsCont.Columns("C").ColumnWidth < 14 Then wsCont.Columns("C").ColumnWidth = 14
    If wsCont.Columns("D").ColumnWidth < 18 Then wsCont.Columns("D").ColumnWidth = 18

    wsCont.Activate
    wsCont.Cells(1, 1).Select

    MsgBox "Contabilizado actualizado." & Chr(10) & Chr(10) & _
           "  Meses procesados: " & nMesesC & Chr(10) & Chr(10) & _
           "  AP total:                        " & totAP & Chr(10) & _
           "  Supervisores total:               " & totSup & Chr(10) & _
           "  AP que lideran equipo:            " & totAPLider & Chr(10) & _
           "  Supervisores que lideran equipo:  " & totSupLider & Chr(10) & Chr(10) & _
           "  Excluidos por Yape/IO:            " & excC, _
           vbInformation, "Contabilizado AP y Supervisores"

CleanupCont:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrHandlerCont:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error Contabilizado"
End Sub

'===============================================================================
'  HELPER: Encabezado hoja Contabilizado
'===============================================================================
Private Sub FormatContHeader(ws As Worksheet, excluidos As Long)
    ' Fila 1: Título
    Merge_Cells_Safe ws, 1, 1, 1, 4
    With ws.Cells(1, 1)
        .Value = "CONTABILIZADO AP Y SUPERVISORES 2026"
        .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 13
        .Font.Color = CLR_GOLD
        .Interior.Color = CLR_DARK_BLUE
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(1).RowHeight = 32

    ' Fila 2: Subtítulo
    Merge_Cells_Safe ws, 2, 1, 2, 4
    With ws.Cells(2, 1)
        .Value = "Excluye Canal 2 que contenga 'Yape' o 'IO'  ·  Solo año 2026"
        .Font.Name = "Arial": .Font.Italic = True: .Font.Size = 9
        .Font.Color = RGB(200, 200, 200)
        .Interior.Color = CLR_DARK_BLUE
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(2).RowHeight = 18

    ' Fila 3: Nota de liderazgo
    Merge_Cells_Safe ws, 3, 1, 3, 4
    With ws.Cells(3, 1)
        .Value = "Lidera equipo = su matricula aparece como Jefe de al menos un colaborador con 'Asesor' en Tipo de Organico"
        .Font.Name = "Arial": .Font.Italic = True: .Font.Size = 8
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = CLR_YELLOW_NOTE
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(3).RowHeight = 16

    ' Fila 4: espacio
    ws.Rows(4).RowHeight = 6

    ' Fila 5: encabezados de columna
    Dim headers(1 To 4) As String
    headers(1) = "Indicador"
    headers(2) = "Cantidad"
    headers(3) = "Mes"
    headers(4) = "Detalle"

    Dim col As Long
    For col = 1 To 4
        With ws.Cells(5, col)
            .Value = headers(col)
            .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 10
            .Font.Color = CLR_GOLD
            .Interior.Color = CLR_MID_BLUE
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
        End With
    Next col
    ws.Rows(5).RowHeight = 22
End Sub

'===============================================================================
'  HELPER: Subencabezado de mes
'===============================================================================
Private Sub WriteContMesHeader(ws As Worksheet, rowNum As Long, label As String)
    Merge_Cells_Safe ws, rowNum, 1, rowNum, 4
    With ws.Cells(rowNum, 1)
        .Value = "  >  " & label
        .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 11
        .Font.Color = CLR_GOLD
        .Interior.Color = CLR_DARK_BLUE
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(rowNum).RowHeight = 24
End Sub

'===============================================================================
'  HELPER: Fila de dato en contabilizado
'  isLider:  True = fila de "lidera equipo" (color más intenso)
'  isSup:    True = es Supervisor (verde), False = AP (naranja)
'===============================================================================
Private Sub WriteContDataRow(ws As Worksheet, rowNum As Long, _
                              indicador As String, cantidad As Long, _
                              isLider As Boolean, isSup As Boolean)
    Dim bg As Long
    If isLider Then
        bg = IIf(isSup, CLR_MID_BLUE, CLR_DARK_BLUE)
    Else
        bg = IIf(isSup, CLR_LIGHT_GREEN, CLR_LIGHT_ORANGE)
    End If

    Dim fgColor As Long
    fgColor = IIf(isLider, CLR_GOLD, RGB(50, 50, 50))

    Dim detalle As String
    If isLider Then
        detalle = "Tiene al menos 1 Asesor a cargo"
    Else
        detalle = IIf(isSup, "Tipo contiene SUPERVISOR", "Tipo contiene AP")
    End If

    Dim vals(1 To 4) As Variant
    vals(1) = indicador
    vals(2) = cantidad
    vals(3) = ""          ' mes se rellena desde el bloque de mes
    vals(4) = detalle

    Dim col As Long
    For col = 1 To 4
        With ws.Cells(rowNum, col)
            .Value = vals(col)
            .Interior.Color = bg
            .Font.Name = "Arial": .Font.Size = 10: .Font.Bold = isLider
            .Font.Color = fgColor
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = IIf(col = 1 Or col = 4, xlLeft, xlCenter)

            Dim s As Long, sides(3) As Long
            sides(0) = xlEdgeLeft: sides(1) = xlEdgeRight
            sides(2) = xlEdgeTop:  sides(3) = xlEdgeBottom
            For s = 0 To 3
                With .Borders(sides(s))
                    .LineStyle = xlContinuous
                    .Color = RGB(170, 170, 170): .Weight = xlThin
                End With
            Next s
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone

            If col = 2 Then .NumberFormat = "#,##0"
        End With
    Next col
    ws.Rows(rowNum).RowHeight = 20
End Sub
