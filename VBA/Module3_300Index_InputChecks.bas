Attribute VB_Name = "Module3"
'**************************************************************************
'
'       This module contains Input Checking Routines for the 300 Index
'
'**************************************************************************

Option Explicit

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

Function checkinput_I300() As Boolean
    If IsEmpty(Cells(7, 3)) Or IsEmpty(Cells(8, 3)) Or _
        Cells(7, 3) < 0.1 Or Cells(7, 3) > 100 Or _
        Cells(8, 3) < 10 Or Cells(8, 3) > 15000 Or _
        (IsEmpty(Cells(9, 3)) And IsEmpty(Cells(10, 3)) And IsEmpty(Cells(11, 3))) Then
        If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: 300 Index measurement")
        checkinput_I300 = False
    Else
        checkinput_I300 = True
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

Function checkinput_site() As Boolean
    If IsEmpty(Cells(3, 3)) Or Cells(3, 3) < 1 Or Cells(3, 3) > 60 Then
        If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: No 300 Index")
        checkinput_site = False
    Else
        checkinput_site = True
    End If
End Function

Function checkinput_initialstock() As Boolean
    If IsEmpty(Cells(19, 3)) Or Cells(19, 3) < 1 Or Cells(19, 3) > 80000 Then
        If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: initial stocking")
        checkinput_initialstock = False
    Else
        checkinput_initialstock = True
    End If
End Function

Function checkinput_SI() As Boolean
    If IsEmpty(Cells(4, 3)) Or Cells(4, 3) < 5 Or Cells(4, 3) > 60 Then
        If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: no Site Index")
        checkinput_SI = False
    Else
        checkinput_SI = True
    End If
End Function

Function checkinput_prune() As Boolean
    Dim i As Long
    For i = 40 To 44
        If Not IsEmpty(Cells(i, 2)) Then
            If i > 30 And IsEmpty(Cells(i - 1, 2)) Then GoTo Error2
            If IsEmpty(Cells(i, 3)) Then GoTo Error2
            If Cells(i, 3) < 0 Or Cells(i, 3) > 20 Then GoTo Error2
        End If
    Next i
    checkinput_prune = True
    Exit Function
Error2: If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: pruning")
    checkinput_prune = False
End Function

Function checkinput_fellage() As Boolean
    Dim fellrowno As Long
    fellrowno = 47
    If Cells(fellrowno, 3) < 1 Or Cells(fellrowno, 3) > 100 Then
        If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: fell age")
        checkinput_fellage = False
    Else
        checkinput_fellage = True
    End If
End Function

Function checkinput_steplth() As Boolean
        If Cells(48, 3) < 0.01 Or Cells(48, 3) > 2 Then
            If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: step length")
            checkinput_steplth = False
        Else
            checkinput_steplth = True
        End If
End Function

Function checkinput_volfn() As Boolean
    Dim mods As Long, i As Long
        mods = 0
        For i = 51 To 61
            If LCase(Cells(i, 4)) = "x" Then mods = mods + 1
        Next i
        If mods <> 1 Then
            If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: Volume table")
            checkinput_volfn = False
        Else
            checkinput_volfn = True
        End If
End Function

Function checkinput_mortfn() As Boolean
    Dim mods As Long, i As Long
        mods = 0
        For i = 68 To 72
            If LCase(Cells(i, 4)) = "x" Then mods = mods + 1
        Next i
        If mods <> 1 Then
            If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: Mortality function")
            checkinput_mortfn = False
        Else
            checkinput_mortfn = True
        End If
End Function

Function checkinput_stocking() As Boolean
   Dim i As Long, Age As Double, sph1 As Double, sph2 As Double, lastage As Boolean, lastsph As Boolean, _
        oldage As Double, oldsph As Double

    oldage = 0
    oldsph = Cells(19, 3)
    lastage = False
    lastsph = False
    checkinput_stocking = True
    For i = 20 To 36
        Age = Cells(i, 2)
        sph1 = Cells(i, 3)
        sph2 = Cells(i, 4)
        If Age = 0 Then lastage = False
        If sph1 = 0 And sph2 = 0 Then lastsph = False
        If (lastage And Not lastsph) Or (lastsph And Not lastage) Then checkinput_stocking = False
        If lastage And Age <> 0 Then checkinput_stocking = False
        If lastsph And (sph1 <> 0 Or sph2 <> 0) Then checkinput_stocking = False
        If Age <> 0 And Age <= oldage Then checkinput_stocking = False
        If sph1 <> 0 And sph1 > oldsph Then checkinput_stocking = False
        If sph2 <> 0 And sph2 > oldsph Then checkinput_stocking = False
        If sph2 <> 0 And sph1 <> 0 And sph2 > sph1 Then checkinput_stocking = False
        oldage = Age
        oldsph = sph1
        If sph2 > 0 Then oldsph = sph2
    Next i
Error1: If Not checkinput_stocking And (implementation = 1 Or implementation = 5) Then _
        MsgBox ("Input Error: stocking")
End Function


'Earlier stocking input checking procedure
Function checkinput_stock() As Boolean
    Dim i As Long, ci2 As Double, ci3 As Double, ci4 As Double, cim2 As Double, _
        cim3 As Double, cim4 As Double
        
    For i = 20 To 36
        ci2 = Cells(i, 2)
        ci3 = Cells(i, 3)
        ci4 = Cells(i, 4)
        cim2 = Cells(i - 1, 2)
        cim3 = Cells(i - 1, 3)
        cim4 = Cells(i - 1, 4)
        If Not IsEmpty(Cells(i, 2)) Then
            If IsEmpty(Cells(i - 1, 2)) Then GoTo Error1
            If ci2 <> 0 And ci2 > 100 Then GoTo Error1
            If ci2 <> 0 And cim2 <> 0 And ci2 < cim2 Then GoTo Error1
            If ci3 <> 0 And cim3 <> 0 And ci3 > cim3 Then GoTo Error1
            If ci4 <> 0 And ci3 <> 0 And ci4 > ci3 Then GoTo Error1
            If IsEmpty(Cells(i, 3)) And IsEmpty(Cells(i, 4)) Then GoTo Error1
            If ci3 <> 0 And cim4 <> 0 And ci3 > cim4 Then GoTo Error1
        End If
    Next i
    checkinput_stock = True
    Exit Function
Error1: If implementation = 1 Or implementation = 5 Then MsgBox ("Input Error: stocking")
    checkinput_stock = False
End Function



