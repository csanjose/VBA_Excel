Attribute VB_Name = "M�dulo1"
Function matriz(inferior As Integer, superior As Integer, tama�o As Integer)

ReDim m(tama�o, tama�o)

For i = 1 To tama�o
  For j = 1 To tama�o
    m(i, j) = Int((superior - inferior + 1) * Rnd() + inferior)
    mat = mat & " [" & i & ", " & j & "] " & m(i, j)
  Next j
Next i

matriz = mat


End Function
