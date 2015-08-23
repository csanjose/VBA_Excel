Attribute VB_Name = "Módulo1"
Function numeros(valor As String) As String

'Esta funcion extrae todos los números de una cadena

Dim i As Integer

For i = 1 To Len(valor)

  If IsNumeric(Mid(valor, i, 1)) Then
    cadena = cadena + Mid(valor, i, 1)
  End If

Next

numeros = cadena

End Function
