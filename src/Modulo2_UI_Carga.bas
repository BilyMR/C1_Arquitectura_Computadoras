Option Explicit
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
Public Sub ResetCPU()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' --- TUS LÍNEAS HABITUALES DE RESET (ej. limpiar registros, banderas, RAM, etc.) ---
    ws.Range("S7").Value = "00h" ' Limpiar AX
    ws.Range("W7").Value = "00h" ' Limpiar BX
    ws.Range("AA7").Value = "Z:0 C:0 S:0" ' Limpiar Banderas
    ws.Range("C7").Value = "00h" ' Limpiar PC
    
    ' --- NUEVA LÍNEA: LIMPIAR EL HISTORIAL DEL LOG (TextBox) ---
    On Error Resume Next ' Evita errores si la hoja activa no tiene el cuadro de texto
    ws.OLEObjects("txtLog").Object.Text = ""
    On Error GoTo 0
    
    ' Mensaje de estado inicial
    ws.Range("AH7").Value = "SISTEMA REINICIADO (READY)"
    
    ' Opcional: Si quieres registrar el reset en el log ya limpio
    ' AgregarLog ws, "CPU Reiniciada a estado inicial."
End Sub
Public Sub CargarProgramaSuma()
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
Private Sub EscribirRAM_Matriz(ws As Worksheet, direccion As Long, valor As String)
    ' Fila donde empieza la dirección 00h (Fila 21)
    Dim filaBase As Long
    filaBase = 21
    
    ' Columna donde empieza el offset +0 (Columna C es la número 3)
    Dim colBase As Long
    colBase = 3
    
    ' Cálculo exacto para matriz de 16 bytes por fila:
    ' - Fila: Se desplaza hacia abajo cada 16 posiciones (Int(direccion / 16))
    ' - Columna: El residuo indica el desplazamiento de +0 a +F (direccion Mod 16)
    Dim filaDestino As Long
    Dim colDestino As Long
    
    filaDestino = filaBase + Int(direccion / 16)
    colDestino = colBase + (direccion Mod 16)
    
    ' Escribe el valor en la celda exacta de la matriz visual
    ws.Cells(filaDestino, colDestino).Value = valor
End Sub

Public Sub CargarPrograma_Bifurcacion()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    On Error GoTo ManejadorError
    
    ' 1. Limpiar registros y estado inicial de la CPU
    Call ResetCPU
    
    ' 2. Secuencia de bytes del programa de Bifurcación
    Dim bytesProg As Variant
    bytesProg = Array("10", "05", "13", "15", "15", "15", "27", "05", "31", "0C", "10", "00", "10", "FF", "13", "20", "FF")
    
    ' 3. Volcar los bytes distribuyéndolos ordenadamente en la matriz de la RAM
    Dim i As Long
    For i = LBound(bytesProg) To UBound(bytesProg)
        EscribirRAM_Matriz ws, CLng(i), bytesProg(i) & "h"
    Next i
    
    ' 4. AGREGAR ESTA LÍNEA: Enviar el mensaje al cuadro de texto del Log
    AgregarLog ws, "=== Prueba de bifurcación cargada ==="
    Exit Sub

ManejadorError:
    MsgBox "Ocurrió un error al cargar: " & Err.Description, vbCritical, "Error"
End Sub
