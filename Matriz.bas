Attribute VB_Name = "Módulo1"
Function matriz(inferior As Integer, superior As Integer, tamaño As Integer)

ReDim m(tamaño, tamaño)

For i = 1 To tamaño
  For j = 1 To tamaño
    m(i, j) = Int((superior - inferior + 1) * Rnd() + inferior)
    mat = mat & " [" & i & ", " & j & "] " & m(i, j)
  Next j
Next i

matriz = mat


End Function
