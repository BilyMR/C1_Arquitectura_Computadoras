Option Explicit

' ==============================================================================
' SIMULADOR CPU VON NEUMANN - MÓDULO ÚNICO COMPLETO
' ==============================================================================

Public Sub Pausa(Segundos As Double)
    Dim TiempoFinal As Double
    TiempoFinal = Timer + Segundos
    Do While Timer < TiempoFinal
        DoEvents
    Loop
End Sub

Public Function LimpiarHex(ByVal Valor As String) As String
    Valor = Trim(Valor)
    Valor = Replace(Valor, "h", "", , , vbTextCompare)
    Valor = Replace(Valor, "H", "", , , vbTextCompare)
    If Valor = "" Then Valor = "00"
    If Len(Valor) = 1 Then Valor = "0" & Valor
    LimpiarHex = UCase(Valor)
End Function

Public Function HexToLong(ByVal ValorHex As String) As Long
    On Error Resume Next
    HexToLong = CLng("&H" & LimpiarHex(ValorHex))
    If Err.Number <> 0 Then
        HexToLong = 0
        Err.Clear
    End If
End Function

Public Function GetRAMCell(ws As Worksheet, Address As Long) As Range
    Dim Fila As Long, Columna As Long
    Fila = 21 + (Address \ 16)
    Columna = 3 + (Address Mod 16)
    Set GetRAMCell = ws.Cells(Fila, Columna)
End Function

Public Function GetRAMOriginalColor(Address As Long) As Long
    If Address < 16 Then
        GetRAMOriginalColor = RGB(224, 242, 254) ' Código (Azul Claro)
    ElseIf Address >= 240 Then
        GetRAMOriginalColor = RGB(252, 231, 243) ' Pila (Rosa)
    Else
        GetRAMOriginalColor = RGB(255, 255, 255) ' Datos (Blanco)
    End If
End Function

Public Function ReadRAM(ws As Worksheet, Address As Long) As String
    ReadRAM = LimpiarHex(CStr(GetRAMCell(ws, Address).Value))
End Function

Public Sub WriteRAM(ws As Worksheet, Address As Long, Value As String)
    Dim CellRAM As Range
    Set CellRAM = GetRAMCell(ws, Address)
    CellRAM.NumberFormat = "@"
    CellRAM.Value = LimpiarHex(Value)
End Sub

Public Sub AgregarLog(ws As Worksheet, Mensaje As String)
    Dim i As Long
    For i = 34 To 13 Step -1
        ws.Cells(i, 34).Value = ws.Cells(i - 1, 34).Value
    Next i
    ws.Cells(12, 34).Value = "[" & Format(Now, "hh:mm:ss") & "] " & Mensaje
End Sub

' ------------------------------------------------------------------------------
' RESALTAR INSTRUCCIÓN ACTIVA EN LA TABLA "SEGMENTO DE CÓDIGO"
' ------------------------------------------------------------------------------
Public Sub ResaltarInstruccionCodigo(ws As Worksheet, Address As Long)
    Dim r As Long
    For r = 21 To 28
        ws.Cells(r, 20).Interior.Color = RGB(241, 245, 249)
        ws.Range(ws.Cells(r, 21), ws.Cells(r, 23)).Interior.Color = RGB(255, 255, 255)
    Next r
    
    Dim FilaObjetivo As Long
    Select Case Address
        Case 0: FilaObjetivo = 21   ' 00h: MOV AX, 00h
        Case 2: FilaObjetivo = 22   ' 02h: MOV BX, 03h
        Case 4: FilaObjetivo = 23   ' 04h: ADD AX, 04h
        Case 6: FilaObjetivo = 24   ' 06h: DEC BX
        Case 7: FilaObjetivo = 25   ' 07h: JNZ 04h
        Case 9: FilaObjetivo = 26   ' 09h: STORE [10h], AX
        Case 11: FilaObjetivo = 27  ' 0Bh: HLT
    End Select
    
    If FilaObjetivo >= 21 And FilaObjetivo <= 27 Then
        ws.Range(ws.Cells(FilaObjetivo, 20), ws.Cells(FilaObjetivo, 23)).Interior.Color = RGB(254, 240, 138)
    End If
End Sub

' ------------------------------------------------------------------------------
' BOTÓN: REINICIAR CPU (RESET REGISTROS Y MARCADOR VISUAL - MANTECHO RAM)
' ------------------------------------------------------------------------------
Public Sub ResetCPU()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Simulador Von Neumann")
    Application.ScreenUpdating = True
    
    ' Restablecer registros de la CPU a valores iniciales
    ws.Range("C7").Value = "00h"         ' PC (Program Counter)
    ws.Range("H7").Value = "00h 00h"    ' IR (Instruction Register)
    ws.Range("M7").Value = "FFh"         ' SP (Stack Pointer al tope)
    ws.Range("C12").Value = "00h"        ' MAR
    ws.Range("H12").Value = "00h"        ' MDR
    ws.Range("S7").Value = "00h"         ' AX
    ws.Range("W7").Value = "00h"         ' BX
    ws.Range("AA7").Value = "Z:0 C:0 S:0" ' Banderas (Flags)
    
    ' Volver el marcador visual de la tabla de código a la primera línea
    ResaltarInstruccionCodigo ws, 0
    
    ws.Range("AH7").Value = "CPU REINICIADO: Registros en 00h (RAM intacta)"
    AgregarLog ws, "=== CPU Reiniciado. PC restablecido a 00h ==="
End Sub

' ------------------------------------------------------------------------------
' BOTÓN: CARGAR PROGRAMA (RESET COMPLETO + CARGA DE BYTES EN RAM)
' ------------------------------------------------------------------------------
Public Sub CargarProgramaPrueba()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Simulador Von Neumann")
    Application.ScreenUpdating = True
    
    Dim addr As Long
    For addr = 0 To 255
        WriteRAM ws, addr, "00"
        GetRAMCell(ws, addr).Interior.Color = GetRAMOriginalColor(addr)
    Next addr
    
    ' Programa: Multiplicación 3 * 4 = 12 / 0Ch
    WriteRAM ws, 0, "10"  ' 00h: MOV AX, 00h
    WriteRAM ws, 1, "00"
    WriteRAM ws, 2, "11"  ' 02h: MOV BX, 03h
    WriteRAM ws, 3, "03"
    WriteRAM ws, 4, "20"  ' 04h: ADD AX, 04h
    WriteRAM ws, 5, "04"
    WriteRAM ws, 6, "23"  ' 06h: DEC BX
    WriteRAM ws, 7, "32"  ' 07h: JNZ 04h
    WriteRAM ws, 8, "04"
    WriteRAM ws, 9, "13"  ' 09h: STORE [10h], AX
    WriteRAM ws, 10, "10"
    WriteRAM ws, 11, "FF" ' 0Bh: HLT
    
    ResetCPU
    ws.Range("AH7").Value = "PROGRAMA CARGADO: Listo para ejecutar"
    AgregarLog ws, "=== Programa demostrativo (3x4 por sumas) cargado en RAM ==="
End Sub

' ------------------------------------------------------------------------------
' SUBRUTINA INTERNA DE UN CICLO
' ------------------------------------------------------------------------------
Public Function EjecutarUnCiclo(ws As Worksheet) As Boolean
    Dim ColorBusAddrNormal As Long, ColorBusAddrActivo As Long
    Dim ColorBusDataNormal As Long, ColorBusDataActivo As Long
    Dim ColorHighlightRAM As Long, ColorHighlightWrite As Long
    Dim ColorLEDNormal As Long, ColorLEDActivo As Long
    
    ColorBusAddrNormal = RGB(59, 130, 246)
    ColorBusAddrActivo = RGB(0, 255, 255)
    ColorBusDataNormal = RGB(234, 179, 8)
    ColorBusDataActivo = RGB(255, 255, 0)
    
    ColorHighlightRAM = RGB(254, 240, 138)
    ColorHighlightWrite = RGB(134, 239, 172)
    ColorLEDNormal = RGB(15, 23, 42)
    ColorLEDActivo = RGB(30, 58, 138)
    
    Dim PC As Long, MAR As Long
    Dim Opcode As String, Operando As String
    Dim CellRAM As Range
    
    ' --- FETCH ---
    PC = HexToLong(CStr(ws.Range("C7").Value))
    MAR = PC
    
    ResaltarInstruccionCodigo ws, MAR
    
    ws.Range("C7:F8").Interior.Color = ColorLEDActivo
    Pausa 0.2
    ws.Range("C12").Value = Format(Hex(MAR), "00") & "h"
    ws.Range("C12:F13").Interior.Color = ColorLEDActivo
    ws.Range("C7:F8").Interior.Color = ColorLEDNormal
    
    ws.Range("B16:AD16").Interior.Color = ColorBusAddrActivo
    Set CellRAM = GetRAMCell(ws, MAR)
    CellRAM.Interior.Color = ColorHighlightRAM
    Pausa 0.3
    
    Opcode = ReadRAM(ws, MAR)
    ws.Range("H12").Value = Opcode & "h"
    ws.Range("B16:AD16").Interior.Color = ColorBusAddrNormal
    ws.Range("B17:AD17").Interior.Color = ColorBusDataActivo
    ws.Range("H12:K13").Interior.Color = ColorLEDActivo
    Pausa 0.3
    
    If Opcode = "23" Or Opcode = "FF" Then
        Operando = "00"
        ws.Range("H7").Value = Opcode & "h"
    Else
        Operando = ReadRAM(ws, MAR + 1)
        ws.Range("H7").Value = Opcode & "h " & Operando & "h"
    End If
    
    CellRAM.Interior.Color = GetRAMOriginalColor(MAR)
    ws.Range("B17:AD17").Interior.Color = ColorBusDataNormal
    ws.Range("C12:F13").Interior.Color = ColorLEDNormal
    ws.Range("H12:K13").Interior.Color = ColorLEDNormal
    
    If Opcode = "23" Or Opcode = "FF" Then
        PC = PC + 1
    Else
        PC = PC + 2
    End If
    ws.Range("C7").Value = Format(Hex(PC), "00") & "h"
    
    ' --- DECODE ---
    AgregarLog ws, "DECODE: Opcode=" & Opcode & "h | Operando=" & Operando & "h"
    Pausa 0.2
    
    ' --- EXECUTE ---
    Select Case UCase(Opcode)
        Case "10" ' MOV AX, imm
            ws.Range("S7").Value = Operando & "h"
            AgregarLog ws, "EXECUTE [MOV AX]: Cargar " & Operando & "h en AX"
            Pausa 0.3
            
        Case "11" ' MOV BX, imm
            ws.Range("W7").Value = Operando & "h"
            AgregarLog ws, "EXECUTE [MOV BX]: Cargar " & Operando & "h en BX"
            Pausa 0.3
            
        Case "20" ' ADD AX, imm
            Dim ValAX As Long, ValImm As Long, ResADD As Long
            ValAX = HexToLong(CStr(ws.Range("S7").Value))
            ValImm = HexToLong(Operando)
            ResADD = ValAX + ValImm
            
            ws.Range("S7").Value = Format(Hex(ResADD Mod 256), "00") & "h"
            AgregarLog ws, "EXECUTE [ADD AX, " & Operando & "h]: AX = " & Format(Hex(ResADD Mod 256), "00") & "h"
            Pausa 0.3
            
        Case "23" ' DEC BX
            Dim ValBX As Long, ZF As Integer
            ValBX = HexToLong(CStr(ws.Range("W7").Value)) - 1
            If ValBX < 0 Then ValBX = 255
            
            ws.Range("W7").Value = Format(Hex(ValBX), "00") & "h"
            ZF = IIf(ValBX = 0, 1, 0)
            ws.Range("AA7").Value = "Z:" & ZF & " C:0 S:0"
            AgregarLog ws, "EXECUTE [DEC BX]: BX = " & Format(Hex(ValBX), "00") & "h (ZF=" & ZF & ")"
            Pausa 0.3
            
        Case "32" ' JNZ dir
            Dim StrFlags As String, ZeroFlagVal As Integer
            StrFlags = CStr(ws.Range("AA7").Value)
            ZeroFlagVal = CInt(Mid(StrFlags, 3, 1))
            
            If ZeroFlagVal = 0 Then
                Dim DirSalto As Long
                DirSalto = HexToLong(Operando)
                ws.Range("C7").Value = Format(Hex(DirSalto), "00") & "h"
                AgregarLog ws, "EXECUTE [JNZ " & Operando & "h]: BUCLE ACTIVO (ZF=0) -> Salto a " & Operando & "h"
            Else
                AgregarLog ws, "EXECUTE [JNZ " & Operando & "h]: FIN DE BUCLE (ZF=1) -> Continúa secuencialmente"
            End If
            Pausa 0.3
            
        Case "13" ' STORE [dir], AX
            Dim DirDestinoAX As Long
            DirDestinoAX = HexToLong(Operando)
            Set CellRAM = GetRAMCell(ws, DirDestinoAX)
            
            ws.Range("B17:AD17").Interior.Color = ColorBusDataActivo
            WriteRAM ws, DirDestinoAX, CStr(ws.Range("S7").Value)
            CellRAM.Interior.Color = ColorHighlightWrite
            AgregarLog ws, "STORE: AX (" & ws.Range("S7").Value & ") guardado en RAM[" & Operando & "h]"
            Pausa 0.6
            
            ws.Range("B17:AD17").Interior.Color = ColorBusDataNormal
            CellRAM.Interior.Color = GetRAMOriginalColor(DirDestinoAX)
            
        Case "FF" ' HLT
            EjecutarUnCiclo = False
            Exit Function
            
        Case Else
            AgregarLog ws, "ERROR: Opcode " & Opcode & "h no válido."
            EjecutarUnCiclo = False
            Exit Function
    End Select
    
    EjecutarUnCiclo = True
End Function

' ------------------------------------------------------------------------------
' BOTÓN: EJECUTAR PROGRAMA COMPLETO
' ------------------------------------------------------------------------------
Public Sub EjecutarProgramaCompleto()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Simulador Von Neumann")
    Application.ScreenUpdating = True
    
    ws.Range("AH7").Value = "EJECUTANDO PROGRAMA..."
    AgregarLog ws, "=== INICIANDO EJECUCIÓN ==="
    
    Dim Continuar As Boolean
    Continuar = True
    
    Do While Continuar
        Continuar = EjecutarUnCiclo(ws)
    Loop
    
    ws.Range("AH7").Value = "COMPLETADO"
    AgregarLog ws, "=== Ejecución de programa COMPLETADA ==="
End Sub
