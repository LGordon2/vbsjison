'Option Explicit

'test...
REM another test...

Dim test
test = 3
test = 4
For i=0 To 5
test = 1
test = 3
Next

Do While test < 5
test = 6
Exit Do
Loop


If test = 7 Then
test = 8
ElseIf test = 6 Then
test = 0
End If 
