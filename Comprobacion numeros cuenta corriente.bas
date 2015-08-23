Attribute VB_Name = "Módulo4"

Function dc(numerocuenta As String) As String

'*** Digito de control

numerocuenta = Replace(numerocuenta, " ", "")

If Len(numerocuenta) <> 20 And Len(numerocuenta) <> 24 Then
    dc = "El número de cuenta debe tener 20 o 24 dígitos"
    Exit Function
End If

If Len(numerocuenta) = 24 Then
    numerocuenta = UCase$(numerocuenta)
    entidad = Mid(numerocuenta, 5, 4)
    sucursal = Mid(numerocuenta, 9, 4)
    dcnumcuenta1 = Mid(numerocuenta, 13, 4)
    dcnumcuenta2 = Mid(numerocuenta, 17, 4)
    dcnumcuenta3 = Mid(numerocuenta, 21, 4)
Else
    entidad = Mid(numerocuenta, 1, 4)
    sucursal = Mid(numerocuenta, 5, 4)
    dcnumcuenta1 = Mid(numerocuenta, 9, 4)
    dcnumcuenta2 = Mid(numerocuenta, 13, 4)
    dcnumcuenta3 = Mid(numerocuenta, 17, 4)
End If

suma = suma + Val(Mid(entidad, 1, 1) * 4)
suma = suma + Val(Mid(entidad, 2, 1) * 8)
suma = suma + Val(Mid(entidad, 3, 1) * 5)
suma = suma + Val(Mid(entidad, 4, 1) * 10)

suma = suma + Val(Mid(sucursal, 1, 1) * 9)
suma = suma + Val(Mid(sucursal, 2, 1) * 7)
suma = suma + Val(Mid(sucursal, 3, 1) * 3)
suma = suma + Val(Mid(sucursal, 4, 1) * 6)

resto = suma Mod 11
digito1 = 11 - resto
If digito1 = 10 Then digito1 = 1
If digito1 = 11 Then digito1 = 0

suma = 0
suma = suma + Val(Mid(dcnumcuenta1, 3, 1) * 1)
suma = suma + Val(Mid(dcnumcuenta1, 4, 1) * 2)

suma = suma + Val(Mid(dcnumcuenta2, 1, 1) * 4)
suma = suma + Val(Mid(dcnumcuenta2, 2, 1) * 8)
suma = suma + Val(Mid(dcnumcuenta2, 3, 1) * 5)
suma = suma + Val(Mid(dcnumcuenta2, 4, 1) * 10)

suma = suma + Val(Mid(dcnumcuenta3, 1, 1) * 9)
suma = suma + Val(Mid(dcnumcuenta3, 2, 1) * 7)
suma = suma + Val(Mid(dcnumcuenta3, 3, 1) * 3)
suma = suma + Val(Mid(dcnumcuenta3, 4, 1) * 6)

resto = suma Mod 11
digito2 = 11 - resto
If digito2 = 10 Then digito2 = 1
If digito2 = 11 Then digito2 = 0

DigitoControl = Val(Trim(Str(digito1)) + Trim(Str(digito2)))

digitocontrol2 = Val(Left(dcnumcuenta1, 2))

If DigitoControl = digitocontrol2 Then
    dc = "Correcto"
Else
    dc = "Dígito de control erróneo"
    Exit Function
End If

'*** IBAN

If Len(numerocuenta) = 20 Then
    Exit Function
Else
    numeroiban$ = Right(numerocuenta, 20) + "142800"
End If

parte1$ = Mid$(numeroiban$, 1, 9)
parte2$ = Mid$(numeroiban$, 10, 7)
parte3$ = Mid$(numeroiban$, 17, 7)
parte4$ = Mid$(numeroiban$, 24, 3)

a = Val(parte1$) Mod 97
B = Val(Format(a) + parte2$) Mod 97
C = Val(Format(B) + parte3$) Mod 97
D = Val(Format(C) + parte4$) Mod 97
   
digcontrol = Format(98 - D)
   
If Len(Trim(digcontrol)) = 1 Then digcontrol = "0" & digcontrol
     
DigitoControl = "ES" & digcontrol & Right(numerocuenta, 20)

If DigitoControl = numerocuenta Then
    dc = "Correcto"
Else
    dc = "Error en código IBAN"
End If

End Function