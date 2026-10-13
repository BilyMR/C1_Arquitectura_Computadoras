Option Explicit

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
    On Error Resume Next ' Evita errores si la hoja activa no tiene el cuadro de texto
    
    Dim nuevoMensaje As String
    nuevoMensaje = "[" & Format(Now, "hh:mm:ss") & "] " & Mensaje
    
    ' Obtener el objeto TextBox llamado "txtLog" de la hoja
    Dim txtBox As Object
    Set txtBox = ws.OLEObjects("txtLog").Object
    
    ' Agregar el mensaje nuevo arriba y conservar el historial abajo
    If txtBox.Text = "" Then
        txtBox.Text = nuevoMensaje
    Else
        txtBox.Text = nuevoMensaje & vbCrLf & txtBox.Text
    End If
    
    On Error GoTo 0
End Sub
