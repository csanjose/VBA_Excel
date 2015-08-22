Attribute VB_Name = "Módulo1"
Function email(cadena As String)

'changelog
'28.02.11; Corregido problemas con correos.e que no tengan espacios antes y/o despues
'28.02.11; Corregido problemas con cadenas que no contengan correos.e o sean nulas

Do
t = t + 1
If Mid$(cadena, t, 1) = "@" Then
    flag = 1
    Exit Do
End If
    
Loop Until t = Len(cadena) Or Len(cadena) = 0

    Do While t - x - 1 > 0 And flag = 1
        If Mid$(cadena, t - x - 1, 1) = " " Then Exit Do
        x = x + 1
    Loop
    
    Do While x + y + 1 < Len(cadena) And flag = 1
        If Mid$(cadena, t + y, 1) = " " Then Exit Do
        y = y + 1
    Loop

'cogemos x caracteres a la izq de la @ e y caracteres a la derecha
'la longitud total es x+1(@)+y

If flag = 1 Then email = Mid$(cadena, t - x, x + y + 1)
If flag = 0 Then email = ""

End Function
