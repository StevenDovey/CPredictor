Attribute VB_Name = "Module8"
'*************************************************************************************************************
'
' This module contains input error checking routines
'
'*************************************************************************************************************


Sub Error_checks_1()
    Dim thin As Integer, lift As Integer
    Worksheets("Inputs").Activate
    Error_flag = False
    If Cells(6, 5) < 1 Or Cells(6, 5) > 100 Then
        Error_flag = True
        MsgBox ("Rotation length outside allowed range")
    End If
    If Cells(5, 5) < 1 Or Cells(5, 5) > 10000 Then
        Error_flag = True
        MsgBox ("Stocking at planting not specified or outside allowed range")
    End If
    For thin = 1 To 4
        If Cells(11, thin + 4) <> 0 Then
            If Cells(11, thin + 4) < 1 Or Cells(11, thin + 4) > 100 Then
                Error_flag = True
                MsgBox ("Age of thinning outside allowed range")
            End If
            If Cells(12, thin + 4) < 1 Or Cells(12, thin + 4) > 10000 Then
                Error_flag = True
                MsgBox ("Stand density after thinning outside allowed range")
            End If
            If Cells(13, thin + 4) <> -999 And (Cells(13, thin + 4) < 0 Or Cells(13, thin + 4) > 10) Then
                Error_flag = True
                MsgBox ("Thinning coefficient outside allowed range")
            End If
        End If
    Next thin
    For lift = 1 To 4
        If Cells(9, lift + 12) <> 0 Then
            If Cells(9, lift + 12) < 1 Or Cells(11, lift + 12) > 100 Then
                Error_flag = True
                MsgBox ("Age of pruning outside allowed range")
            End If
            If Cells(10, lift + 12) < 1 Or Cells(10, lift + 12) > 10000 Then
                Error_flag = True
                MsgBox ("Number of stems pruned outside allowed range")
            End If
            If Cells(11, lift + 12) < 0 Or Cells(11, lift + 12) > 15 Then
                Error_flag = True
                MsgBox ("Pruning height outside allowed range")
            End If
        End If
    Next lift
End Sub

Sub Error_checks_2()
    Worksheets("Inputs").Activate
    If Cells(3, 5) < 1 Or Cells(3, 5) > 70 Then
        Error_flag = True
        MsgBox ("300 Index outside allowed range")
    End If
    If Cells(4, 5) < 1 Or Cells(4, 5) > 60 Then
        Error_flag = True
        MsgBox ("Site Index outside allowed range")
    End If
End Sub

Sub Error_checks_3()
    Worksheets("Inputs").Activate
    If Cells(20, 5) < 1 Or Cells(20, 5) > 200 Then
        Error_flag = True
        MsgBox ("Measurement age outside allowed range")
    End If
    If Cells(21, 5) < 1 Or Cells(21, 5) > 10000 Then
        Error_flag = True
        MsgBox ("Measurement stocking outside allowed range")
    End If
    If Cells(22, 5) < 1 Or Cells(22, 5) > 200 Then
        Error_flag = True
        MsgBox ("Measurement height outside allowed range")
    End If
    If Cells(23, 5) < 1 Or Cells(23, 5) > 500 Then
        Error_flag = True
        MsgBox ("Measurement DBH or BA outside allowed range")
    End If
    If Cells(24, 5) <> 0 And (Cells(24, 5) < 1 Or Cells(24, 5) > 200) Then
        Error_flag = True
        MsgBox ("Early height measurement age outside allowed range")
    End If
    If Cells(24, 5) <> 0 And (IsEmpty(Cells(25, 5)) Or Cells(25, 5) < 1 Or Cells(25, 5) > 60) Then
        Error_flag = True
        MsgBox ("Early height measurement not specified or outside allowed range")
    End If
End Sub

Sub Error_checks_4()
    Dim tree As Long, No_dbh As Long, No_ht As Long
    Worksheets("Starting tree list").Activate
    If Cells(3, 2) < 0.001 Or Cells(3, 2) > 100 Then
        Error_flag = True
        MsgBox ("Plot area of tree list not specified or outside allowed range")
    End If
    If Cells(4, 2) < 1 Or Cells(4, 2) > 200 Then
        Error_flag = True
        MsgBox ("Age of trees in tree list not specified or outside allowed range")
    End If
    nstems = Range(Cells(7, 1), Cells(7, 1).End(xlDown)).Rows.Count
    If nstems < 2 Or nstems > 1000 Then
        Error_flag = True
        MsgBox ("Number of stems in tree list outside allowed range")
    End If
    No_dbh = 0
    No_ht = 0
    For tree = 1 To nstems
        If Cells(tree + 6, 2) <> 0 Then No_dbh = No_dbh + 1
        If Cells(tree + 6, 3) <> 0 Then No_ht = No_ht + 1
    Next tree
    If No_dbh <> nstems Then
        Error_flag = True
        MsgBox ("At least one stem in tree list has a missing DBH")
    End If
    If No_ht < 3 Then
        Error_flag = True
        MsgBox ("At least 3 stems in tree list must have a measured height")
    End If
    Worksheets("Growth model").Activate
End Sub

Sub Error_checks_4b()
    Dim tree As Long, No_dbh As Long, No_ht As Long
    Worksheets("Starting tree list").Activate
    nstems = Range(Cells(7, 1), Cells(7, 1).End(xlDown)).Rows.Count
    If nstems < 2 Or nstems > 1000 Then
        Error_flag = True
        MsgBox ("Number of stems in tree list outside allowed range")
    End If
    No_dbh = 0
    For tree = 1 To nstems
        If Cells(tree + 6, 2) <> 0 Then No_dbh = No_dbh + 1
    Next tree
    If No_dbh <> nstems Then
        Error_flag = True
        MsgBox ("At least one stem in tree list has a missing DBH")
    End If
    Worksheets("Growth model").Activate
End Sub

Sub Error_checks_5()
    Dim T As Long
    For T = 1 To rotlength
        If N(T) > N(T - 1) Then MsgBox ("Warning: Yield table shows an increase in predicted stocking - check compatibility of stocking at planting, stocking after each thinning, and stocking of calibration measurement")
    Next T
End Sub


