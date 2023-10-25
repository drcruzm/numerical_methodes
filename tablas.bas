Attribute VB_Name = "tablas"
Function fm(m)
fm = 2 * m ^ 3 + Log(m) - Cos(m) / Exp(m) + Sin(m)
End Function

Sub tabla()


paso = Cells(2, 3)
Final = Cells(3, 3)

x = Cells(12, 2)

    For i = 1 To Final
    
    Cells(11 + i, 2) = x
    Cells(11 + i, 3) = fm(x)
    x = x + paso
    
    Next i

End Sub
