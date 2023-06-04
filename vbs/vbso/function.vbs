Function Sum( a, b, c )
    Sum = a + b + c
End Function

For i = 1 To 10
    WScript.Echo Sum( Rnd * 10, Rnd * 10, Rnd * 10 )
Next
