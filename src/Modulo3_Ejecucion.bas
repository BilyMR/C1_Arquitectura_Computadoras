Option Explicit

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
