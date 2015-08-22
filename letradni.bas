Attribute VB_Name = "Módulo1"
Function letradni(dni As Long) As String

valor = dni - Int(dni / 23) * 23 + 1

letradni = Mid("TRWAGMYFPDXBNJZSQVHLCKEO", valor, 1)


End Function
