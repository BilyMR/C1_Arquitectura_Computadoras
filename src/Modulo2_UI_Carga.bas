Option Explicit

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
' BOTÓN: REINICIAR CPU (RESET REGISTROS Y MARCADOR VISUAL - MANTIENE RAM)
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
