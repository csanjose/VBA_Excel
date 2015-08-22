Attribute VB_Name = "Módulo1"
Public Function DigitoControl(ByVal NroCta As String) As String
   
NroCta = Replace(NroCta, " ", "")
numeroiban$ = NroCta + "142800"

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
     
DigitoControl = "ES" & digcontrol

For x = 0 To 4

DigitoControl = DigitoControl & " " & Mid(NroCta, 4 * x + 1, 4)

Next x




End Function
