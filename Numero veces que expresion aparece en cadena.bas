Attribute VB_Name = "Módulo1"
Function veces(cadena, expresion As String) As Integer

'Esta funcion calcula el número de veces que <expresion> aparece en <cadena>

Dim i, cuenta As Integer

For i = 1 To Len(cadena)

  If Mid(cadena, i, Len(expresion)) = expresion Then
    cuenta = cuenta + 1
  End If
  
Next

veces = cuenta

End Function
