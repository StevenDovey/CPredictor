





Function checkinput_htage() As Boolean
If IsEmpty(Cells(14, 3)) Or IsEmpty(Cells(15, 3)) Or _
Cells(14, 3) < 0.1 Or Cells(14, 3) > 100 Or _
Cells(15, 3) < 0.1 Or Cells(15, 3) > 100 Then
If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: height/age measurement")
checkinput_htage = False
Else
checkinput_htage = True
End If
End Function



Function checkinput_htfn() As Boolean
Dim mods As Long, i As Long
mods = 0
For i = 64 To 65
If LCase(Cells(i, 4)) = "x" Then mods = mods + 1
Next i
If mods <> 1 Then
If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: Height function")
checkinput_htfn = False
Else
checkinput_htfn = True
End If
End Function