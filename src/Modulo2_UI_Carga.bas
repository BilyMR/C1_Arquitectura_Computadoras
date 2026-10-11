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
    WriteRAM ws, 4, "20"  ' 04h: ADD AX, 04h / SUB
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
Public Sub CargarPruebaUnderflow()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Simulador Von Neumann")
    
    ' Limpiar RAM
    Dim addr As Long
    For addr = 0 To 255
        WriteRAM ws, addr, "00"
    Next addr
    
    ' Programa de prueba: Carga BX en 00h y le resta 1
    WriteRAM ws, 0, "11"  ' 00h: MOV BX, 00h  (BX = 0)
    WriteRAM ws, 1, "00"
    WriteRAM ws, 2, "23"  ' 02h: DEC BX       (BX pasa a FFh / -1)
    WriteRAM ws, 3, "32"  ' 03h: JNZ 02h      (Como FFh != 0, ZF=0 y salta de nuevo)
    WriteRAM ws, 4, "02"
    WriteRAM ws, 5, "FF"  ' 05h: HLT
    
    ResetCPU
    ws.Range("AH7").Value = "PRUEBA UNDERFLOW CARGADA"
    AgregarLog ws, "=== Prueba de entero negativo (Underflow) cargada ==="
End Sub

' ------------------------------------------------------------------------------
' PROGRAMA DEMOSTRATIVO ADICIONAL: CÁLCULO DE FACTORIAL DE 4 (4! = 24 = 18h)
' ------------------------------------------------------------------------------
Public Sub CargarPruebaFactorial()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Simulador Von Neumann")
    Application.ScreenUpdating = True
    
    ' 1. Limpiar completamente la RAM
    Dim addr As Long
    For addr = 0 To 255
        WriteRAM ws, addr, "00"
        GetRAMCell(ws, addr).Interior.Color = GetRAMOriginalColor(addr)
    Next addr
    
    ' 2. Carga del programa de Factorial para 4! = 24 (18h)
    
    ' --- FASE 1: Calcular 4 * 3 = 12 (0Ch) ---
    WriteRAM ws, 0, "10"   ' 00h: MOV AX, 00h (Inicializa acumulador AX en 0)
    WriteRAM ws, 1, "00"
    WriteRAM ws, 2, "11"   ' 02h: MOV BX, 03h (BX = 3, contador para sumar 4 tres veces)
    WriteRAM ws, 3, "03"
    
    ' Bucle 1 (Dirección 04h): Suma 4 tres veces (0 + 4 + 4 + 4 = 12 / 0Ch)
    WriteRAM ws, 4, "20"   ' 04h: ADD AX, 04h (Suma 4 al acumulador)
    WriteRAM ws, 5, "04"
    WriteRAM ws, 6, "23"   ' 06h: DEC BX      (Decrementa contador BX)
    WriteRAM ws, 7, "32"   ' 07h: JNZ 04h     (Salta a 04h si BX != 0)
    WriteRAM ws, 8, "04"
    
    ' En este punto AX = 0Ch (12). Lo guardamos temporalmente en RAM[30h]
    WriteRAM ws, 9, "13"   ' 09h: STORE [30h], AX (Guarda 12 en RAM[30h])
    WriteRAM ws, 10, "30"
    
    ' --- FASE 2: Calcular 12 * 2 = 24 (18h) ---
    WriteRAM ws, 11, "10"  ' 0Bh: MOV AX, 00h (Reinicia AX a 0)
    WriteRAM ws, 12, "00"
    WriteRAM ws, 13, "11"  ' 0Dh: MOV BX, 02h (BX = 2, contador para sumar 12 dos veces)
    WriteRAM ws, 14, "02"
    
    ' Bucle 2 (Dirección 0Fh): Suma 12 (0Ch) dos veces (0 + 12 + 12 = 24 / 18h)
    WriteRAM ws, 15, "20"  ' 0Fh: ADD AX, 0Ch (Suma 12 a AX)
    WriteRAM ws, 16, "0C"
    WriteRAM ws, 17, "23"  ' 11h: DEC BX      (Decrementa contador BX)
    WriteRAM ws, 18, "32"  ' 12h: JNZ 0Fh     (Salta a 0Fh si BX != 0)
    WriteRAM ws, 19, "0F"
    
    ' --- FASE 3: Guardar el resultado final (24 = 18h) en RAM ---
    WriteRAM ws, 20, "13"  ' 14h: STORE [20h], AX (Guarda 18h en RAM[20h])
    WriteRAM ws, 21, "20"
    
    WriteRAM ws, 22, "FF"  ' 16h: HLT (Fin de ejecución)
    
    ' 3. Reiniciar el CPU
    ResetCPU
    ws.Range("AH7").Value = "PROGRAMA FACTORIAL CARGADO: 4! = 24 (18h)"
    AgregarLog ws, "=== Programa demostrativo Factorial 4! cargado en RAM ==="
End Sub
