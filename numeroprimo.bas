Attribute VB_Name = "Módulo1"
Sub primo()

numero = Cells(1, 1)
esprimo = 1

For i = 2 To numero - 1
  resto = numero Mod i
  If resto = 0 Then
    esprimo = 0
    Exit For
  End If
Next i

If esprimo = 0 Then
  Cells(2, 1) = "No es primo"
Else
  Cells(2, 1) = "Es primo"
End If

End Sub
