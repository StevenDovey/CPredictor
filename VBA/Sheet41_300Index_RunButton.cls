VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    'Run button on 300 Index worksheet
    Dim mode As Integer
    mode = Cells(8, 6)
    If mode = 3 Or mode = 2 Then Call OutputGrowth
    If mode = 1 Then
        If Cells(15, 3) <> 0 Then Call siteIndex
        If Cells(8, 3) <> 0 Then
            Call Calc300Index
        Else
            Call OutputGrowth
        End If
    End If
End Sub
