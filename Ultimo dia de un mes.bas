Attribute VB_Name = "Módulo1"
Function ultimodia(fecha As Date) As Date

fecha = DateAdd("m", 1, fecha)

ultimodia = DateSerial(Year(fecha), Month(fecha), 1) - 1

End Function
