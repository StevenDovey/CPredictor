Attribute VB_Name = "Module6"
'*************************************************************************************************************
'
'This module contains input/output procedures for C_Change
'
'*************************************************************************************************************

Option Explicit


Sub inout(i, v)
    If i = 5 Then
        Sheets("INPUT").Activate
        v = ActiveSheet.Cells(line5, col5)
        col5 = col5 + 1
    End If
    If i = 6 Then
        Sheets("LP1OUT").Activate
        ActiveSheet.Cells(line6, col6) = v
        col6 = col6 + 1
    End If
    If i = 10 Then
        Sheets("LP2OUT").Activate
        ActiveSheet.Cells(line10, col10) = v
        col10 = col10 + 1
    End If
    If i = 7 Then
        Sheets("Nitrogen").Activate
        ActiveSheet.Cells(line7, col7) = v
        col7 = col7 + 1
    End If
    If i = 8 Then
        Sheets("Phosphorus").Activate
        ActiveSheet.Cells(line8, col8) = v
        col8 = col8 + 1
    End If
End Sub

Sub newline(i)
    If i = 5 Then
        line5 = line5 + 1
        col5 = 1
    End If
    If i = 6 Then
        line6 = line6 + 1
        col6 = 1
    End If
    If i = 10 Then
        line10 = line10 + 1
        col10 = 1
    End If
    If i = 7 Then
        line7 = line7 + 1
        col7 = 1
    End If
    If i = 8 Then
        line8 = line8 + 1
        col8 = 1
    End If
End Sub

