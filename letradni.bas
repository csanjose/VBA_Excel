Attribute VB_Name = "M�dulo1"
Function letradni(dni As Long) As String

valor = dni - Int(dni / 23) * 23 + 1

letradni = Mid("TRWAGMYFPDXBNJZSQVHLCKEO", valor, 1)


End Function
