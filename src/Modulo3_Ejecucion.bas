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
    
    Dim valAX As Long
    Dim valBX As Long
    Dim ValOp As Long
    Dim Resultado As Long
    Dim StrFlags As String
    
    ' Variable para controlar si el procesador debe continuar o detenerse
    Dim Exito As Boolean
    Exito = True
    
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
    
    ' Determinar si la instrucción ocupa 1 byte (sin operando inmediato en RAM) o 2 bytes
    Dim EsInstruccion1Byte As Boolean
    Select Case UCase(Opcode)
        Case "12", "21", "22", "23", "24", "26", "28", "FF"
            EsInstruccion1Byte = True
        Case Else
            EsInstruccion1Byte = False
    End Select

    If EsInstruccion1Byte Then
        Operando = "00"
        ws.Range("H7").Value = Opcode & "h"
        PC = PC + 1
    Else
        Operando = ReadRAM(ws, MAR + 1)
        ws.Range("H7").Value = Opcode & "h " & Operando & "h"
        PC = PC + 2
    End If
    ws.Range("C7").Value = Format(Hex(PC), "00") & "h"
    
    ' --- DECODE ---
    AgregarLog ws, "DECODE: Opcode=" & Opcode & "h | Operando=" & Operando & "h"
    Pausa 0.2
    
    ' --- EXECUTE ---
    Select Case UCase(Opcode)

    ' --------------------------------------------------------------------------
    ' FAMILIA 1x: TRANSFERENCIA DE DATOS
    ' --------------------------------------------------------------------------
    Case "10" ' MOV AX, imm
        ws.Range("S7").Value = Format(Operando, "00") & "h"
        AgregarLog ws, "EXECUTE [MOV AX, " & Operando & "h]: AX = " & ws.Range("S7").Value
        
    Case "11" ' MOV BX, imm
        ws.Range("W7").Value = Format(Operando, "00") & "h"
        AgregarLog ws, "EXECUTE [MOV BX, " & Operando & "h]: BX = " & ws.Range("W7").Value
        
    Case "12" ' MOV AX, BX
        ws.Range("S7").Value = ws.Range("W7").Value
        AgregarLog ws, "EXECUTE [MOV AX, BX]: AX = " & ws.Range("S7").Value

    Case "13" ' STORE [dir], AX
        Dim DirDestinoAX As Long
        DirDestinoAX = HexToLong(Operando)
        WriteRAM ws, DirDestinoAX, Replace(CStr(ws.Range("S7").Value), "h", "")
        AgregarLog ws, "EXECUTE [STORE [" & Operando & "h], AX]: RAM[" & Operando & "h] = " & ws.Range("S7").Value

    Case "14" ' STORE [dir], BX
        Dim DirDestinoBX As Long
        DirDestinoBX = HexToLong(Operando)
        WriteRAM ws, DirDestinoBX, Replace(CStr(ws.Range("W7").Value), "h", "")
        AgregarLog ws, "EXECUTE [STORE [" & Operando & "h], BX]: RAM[" & Operando & "h] = " & ws.Range("W7").Value

    Case "15" ' LOAD AX, [dir]
        Dim DirOrigenAX As Long
        DirOrigenAX = HexToLong(Operando)
        ws.Range("S7").Value = ReadRAM(ws, DirOrigenAX) & "h"
        AgregarLog ws, "EXECUTE [LOAD AX, [" & Operando & "h]]: AX = " & ws.Range("S7").Value

    Case "16" ' LOAD BX, [dir]
        Dim DirOrigenBX As Long
        DirOrigenBX = HexToLong(Operando)
        ws.Range("W7").Value = ReadRAM(ws, DirOrigenBX) & "h"
        AgregarLog ws, "EXECUTE [LOAD BX, [" & Operando & "h]]: BX = " & ws.Range("W7").Value

    ' --------------------------------------------------------------------------
    ' FAMILIA 2x: ARITMÉTICA Y LÓGICA (ALU)
    ' --------------------------------------------------------------------------
    Case "20" ' ADD AX, imm
        valAX = HexToLong(CStr(ws.Range("S7").Value))
        ValOp = HexToLong(Operando)
        Resultado = (valAX + ValOp) Mod 256
        ws.Range("S7").Value = Format(Hex(Resultado), "00") & "h"
        ActualizarFlags ws, Resultado
        AgregarLog ws, "EXECUTE [ADD AX, " & Operando & "h]: AX = " & ws.Range("S7").Value

    Case "21" ' ADD AX, BX
        valAX = HexToLong(CStr(ws.Range("S7").Value))
        valBX = HexToLong(CStr(ws.Range("W7").Value))
        Resultado = (valAX + valBX) Mod 256
        ws.Range("S7").Value = Format(Hex(Resultado), "00") & "h"
        ActualizarFlags ws, Resultado
        AgregarLog ws, "EXECUTE [ADD AX, BX]: AX = " & ws.Range("S7").Value

    Case "22" ' INC AX
        valAX = HexToLong(CStr(ws.Range("S7").Value))
        Resultado = (valAX + 1) Mod 256
        ws.Range("S7").Value = Format(Hex(Resultado), "00") & "h"
        ActualizarFlags ws, Resultado
        AgregarLog ws, "EXECUTE [INC AX]: AX = " & ws.Range("S7").Value
        
    Case "29" ' INC BX
        valBX = Val("&H" & Replace(ws.Range("W7").Value, "h", ""))
        
        Dim cfINC_BX As Integer
        cfINC_BX = 0 ' En x86, INC no afecta el Carry (C)
        
        valBX = (valBX + 1) And &HFF
        
        Dim zfINC_BX As Integer, sfINC_BX As Integer
        If valBX = 0 Then zfINC_BX = 1 Else zfINC_BX = 0
        If (valBX And &H80) <> 0 Then sfINC_BX = 1 Else sfINC_BX = 0
        
        ws.Range("W7").Value = Right("0" & Hex(valBX), 2) & "h"
        ws.Range("AA7").Value = "Z:" & zfINC_BX & " C:" & cfINC_BX & " S:" & sfINC_BX
        AgregarLog ws, "EXECUTE [INC BX]: BX = " & ws.Range("W7").Value

    Case "23" ' DEC BX
        valBX = Val("&H" & Replace(ws.Range("W7").Value, "h", ""))
        
        Dim cfDEC_BX As Integer
        If valBX = 0 Then cfDEC_BX = 1 Else cfDEC_BX = 0
        
        valBX = (valBX - 1) And &HFF
        
        Dim zfDEC_BX As Integer, sfDEC_BX As Integer
        If valBX = 0 Then zfDEC_BX = 1 Else zfDEC_BX = 0
        If (valBX & &H80) <> 0 Then sfDEC_BX = 1 Else sfDEC_BX = 0 ' Corrección menor del operador And
        If (valBX And &H80) <> 0 Then sfDEC_BX = 1 Else sfDEC_BX = 0
        
        ws.Range("W7").Value = Right("0" & Hex(valBX), 2) & "h"
        ws.Range("AA7").Value = "Z:" & zfDEC_BX & " C:" & cfDEC_BX & " S:" & sfDEC_BX
        AgregarLog ws, "EXECUTE [DEC BX]: BX = " & ws.Range("W7").Value

    Case "24" ' DEC AX
        valAX = Val("&H" & Replace(ws.Range("S7").Value, "h", ""))
        
        Dim cfDEC_AX As Integer
        If valAX = 0 Then cfDEC_AX = 1 Else cfDEC_AX = 0
        
        valAX = (valAX - 1) And &HFF
        
        Dim zfDEC_AX As Integer, sfDEC_AX As Integer
        If valAX = 0 Then zfDEC_AX = 1 Else zfDEC_AX = 0
        If (valAX And &H80) <> 0 Then sfDEC_AX = 1 Else sfDEC_AX = 0
        
        ws.Range("S7").Value = Right("0" & Hex(valAX), 2) & "h"
        ws.Range("AA7").Value = "Z:" & zfDEC_AX & " C:" & cfDEC_AX & " S:" & sfDEC_AX
        AgregarLog ws, "EXECUTE [DEC AX]: AX = " & ws.Range("S7").Value

    Case "25" ' SUB AX, imm
        Dim valResto25 As Long
        valAX = Val("&H" & Replace(ws.Range("S7").Value, "h", ""))
        valResto25 = Val("&H" & Operando)
        
        Dim cfSUB_Imm As Integer
        If valAX < valResto25 Then cfSUB_Imm = 1 Else cfSUB_Imm = 0
        
        valAX = (valAX - valResto25) And &HFF
        
        Dim zfSUB_Imm As Integer, sfSUB_Imm As Integer
        If valAX = 0 Then zfSUB_Imm = 1 Else zfSUB_Imm = 0
        If (valAX And &H80) <> 0 Then sfSUB_Imm = 1 Else sfSUB_Imm = 0
        
        ws.Range("S7").Value = Right("0" & Hex(valAX), 2) & "h"
        ws.Range("AA7").Value = "Z:" & zfSUB_Imm & " C:" & cfSUB_Imm & " S:" & sfSUB_Imm
        AgregarLog ws, "EXECUTE [SUB AX, " & Operando & "h]: AX = " & ws.Range("S7").Value

    Case "26" ' SUB AX, BX
        valAX = Val("&H" & Replace(ws.Range("S7").Value, "h", ""))
        valBX = Val("&H" & Replace(ws.Range("W7").Value, "h", ""))
        
        Dim cfSUB_BX As Integer
        If valAX < valBX Then cfSUB_BX = 1 Else cfSUB_BX = 0
        
        valAX = (valAX - valBX) And &HFF
        
        Dim zfSUB_BX As Integer, sfSUB_BX As Integer
        If valAX = 0 Then zfSUB_BX = 1 Else zfSUB_BX = 0
        If (valAX And &H80) <> 0 Then sfSUB_BX = 1 Else sfSUB_BX = 0
        
        ws.Range("S7").Value = Right("0" & Hex(valAX), 2) & "h"
        ws.Range("AA7").Value = "Z:" & zfSUB_BX & " C:" & cfSUB_BX & " S:" & sfSUB_BX
        AgregarLog ws, "EXECUTE [SUB AX, BX]: AX = " & ws.Range("S7").Value

    Case "27" ' CMP AX, imm
        valAX = Val("&H" & Replace(ws.Range("S7").Value, "h", ""))
        ValOp = Val("&H" & Operando)
        
        Dim cfCMP_Imm As Integer
        If valAX < ValOp Then cfCMP_Imm = 1 Else cfCMP_Imm = 0
        
        Dim wrappedCMP As Long
        wrappedCMP = (valAX - ValOp) And &HFF
        
        Dim zfCMP_Imm As Integer, sfCMP_Imm As Integer
        If wrappedCMP = 0 Then zfCMP_Imm = 1 Else zfCMP_Imm = 0
        If (wrappedCMP And &H80) <> 0 Then sfCMP_Imm = 1 Else sfCMP_Imm = 0
        
        ws.Range("AA7").Value = "Z:" & zfCMP_Imm & " C:" & cfCMP_Imm & " S:" & sfCMP_Imm
        AgregarLog ws, "EXECUTE [CMP AX, " & Operando & "h]: Banderas actualizadas"

    Case "28" ' CMP AX, BX
        valAX = Val("&H" & Replace(ws.Range("S7").Value, "h", ""))
        valBX = Val("&H" & Replace(ws.Range("W7").Value, "h", ""))
        
        Dim cfCMP_BX As Integer
        If valAX < valBX Then cfCMP_BX = 1 Else cfCMP_BX = 0
        
        Dim wrappedCMPBX As Long
        wrappedCMPBX = (valAX - valBX) And &HFF
        
        Dim zfCMP_BX As Integer, sfCMP_BX As Integer
        If wrappedCMPBX = 0 Then zfCMP_BX = 1 Else zfCMP_BX = 0
        If (wrappedCMPBX And &H80) <> 0 Then sfCMP_BX = 1 Else sfCMP_BX = 0
        
        ws.Range("AA7").Value = "Z:" & zfCMP_BX & " C:" & cfCMP_BX & " S:" & sfCMP_BX
        AgregarLog ws, "EXECUTE [CMP AX, BX]: Banderas actualizadas"


    Case "2A" ' SUB BX, imm (Resta un valor inmediato a BX)
        Dim valResto2A As Long
        valBX = Val("&H" & Replace(ws.Range("W7").Value, "h", ""))
        valResto2A = Val("&H" & Operando)
        
        Dim cfSUB_BX_imm As Integer
        If valBX < valResto2A Then cfSUB_BX_imm = 1 Else cfSUB_BX_imm = 0
        
        valBX = (valBX - valResto2A) And &HFF
        
        Dim zfSUB_BX_imm As Integer, sfSUB_BX_imm As Integer
        If valBX = 0 Then zfSUB_BX_imm = 1 Else zfSUB_BX_imm = 0
        If (valBX And &H80) <> 0 Then sfSUB_BX_imm = 1 Else sfSUB_BX_imm = 0
        
        ws.Range("W7").Value = Right("0" & Hex(valBX), 2) & "h"
        ws.Range("AA7").Value = "Z:" & zfSUB_BX_imm & " C:" & cfSUB_BX_imm & " S:" & sfSUB_BX_imm
        AgregarLog ws, "EXECUTE [SUB BX, " & Operando & "h]: BX = " & ws.Range("W7").Value

    ' --------------------------------------------------------------------------
    ' FAMILIA 3x y FF: CONTROL DE FLUJO Y SISTEMA
    ' --------------------------------------------------------------------------
    Case "30" ' JMP dir (Salto incondicional)
        ws.Range("C7").Value = Format(Hex(HexToLong(Operando)), "00") & "h"
        AgregarLog ws, "EXECUTE [JMP " & Operando & "h]: SALTO REALIZADO"

    Case "31" ' JZ dir (Salto si Zero = 1)
        StrFlags = CStr(ws.Range("AA7").Value)
        If InStr(StrFlags, "Z:1") > 0 Then
            ws.Range("C7").Value = Format(Hex(HexToLong(Operando)), "00") & "h"
            AgregarLog ws, "EXECUTE [JZ " & Operando & "h]: SALTO REALIZADO (ZF=1)"
        Else
            AgregarLog ws, "EXECUTE [JZ " & Operando & "h]: NO SALTA (ZF=0)"
        End If

    Case "32" ' JNZ dir (Salto si Zero = 0)
        StrFlags = CStr(ws.Range("AA7").Value)
        If InStr(StrFlags, "Z:0") > 0 Then
            ws.Range("C7").Value = Format(Hex(HexToLong(Operando)), "00") & "h"
            AgregarLog ws, "EXECUTE [JNZ " & Operando & "h]: SALTO REALIZADO (ZF=0)"
        Else
            AgregarLog ws, "EXECUTE [JNZ " & Operando & "h]: NO SALTA (ZF=1)"
        End If

    Case "FF" ' HLT (Halt)
        AgregarLog ws, "EXECUTE [HLT]: Programa detenido."
        ws.Range("AH7").Value = "EJECUCIÓN FINALIZADA: HLT detectado"
        Exito = False ' Detiene el bucle principal

    Case Else
        AgregarLog ws, "ERROR: Opcode " & Opcode & "h no válido."
        Exito = False ' Detiene el bucle si encuentra basura en la RAM
End Select
    
    EjecutarUnCiclo = Exito
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


' ------------------------------------------------------------------------------
' ACTUALIZAR BANDERAS DE LA ALU (ZF: Zero, CF: Carry, SF: Sign)
' ------------------------------------------------------------------------------
Public Sub ActualizarFlags(ws As Worksheet, Resultado As Long)
    Dim zf As Integer
    Dim cf As Integer
    Dim sf As Integer

    ' 1. Bandera de Cero (ZF): Vale 1 si los 8 bits resultantes son 0
    If (Resultado Mod 256) = 0 Then
        zf = 1
    Else
        zf = 0
    End If

    ' 2. Bandera de Acarreo (CF): Vale 1 si hubo desbordamiento (> 255 o < 0)
    If Resultado > 255 Or Resultado < 0 Then
        cf = 1
    Else
        cf = 0
    End If

    ' 3. Bandera de Signo (SF): Vale 1 si el bit más significativo (Bit 7) es 1
    If ((Resultado Mod 256) And &H80) <> 0 Then
        sf = 1
    Else
        sf = 0
    End If

    ' Escribir el estado de las banderas en la celda AA7
    ws.Range("AA7").Value = "Z:" & zf & " C:" & cf & " S:" & sf
End Sub
