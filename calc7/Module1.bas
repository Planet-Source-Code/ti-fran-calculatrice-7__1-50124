Attribute VB_Name = "Module1"
Public M, N, O As Integer
Public Ftrigo As Boolean
Sub sort()
With GraphPlot
    .Label1(0).ForeColor = .Label1(1).ForeColor
    .Label1(0).Caption = .Label1(1).Caption
    .Label1(1).ForeColor = .Label1(2).ForeColor
    .Label1(1).Caption = .Label1(2).Caption
    .Label1(2).ForeColor = .Label1(3).ForeColor
    .Label1(2).Caption = .Label1(3).Caption
    .Label1(3).ForeColor = .Label1(4).ForeColor
    .Label1(3).Caption = .Label1(4).Caption
    .Label1(4).ForeColor = .Label1(5).ForeColor
    .Label1(4).Caption = .Label1(5).Caption
    .Label1(5).ForeColor = .Label1(6).ForeColor
    .Label1(5).Caption = .Label1(6).Caption
    .Label1(6).ForeColor = .Label1(7).ForeColor
    .Label1(6).Caption = .Label1(7).Caption
    .Label1(7).ForeColor = .Label1(8).ForeColor
    .Label1(7).Caption = .Label1(8).Caption
    .Label1(8).ForeColor = .Label1(9).ForeColor
    .Label1(8).Caption = .Label1(9).Caption
    .Label1(9).ForeColor = .Label1(10).ForeColor
    .Label1(9).Caption = .Label1(10).Caption
    .Label1(10).ForeColor = .Label1(11).ForeColor
    .Label1(10).Caption = .Label1(11).Caption
    .Label1(11).ForeColor = .Label1(12).ForeColor
    .Label1(11).Caption = .Label1(12).Caption
End With
End Sub
Sub RndColor()
    M = Int(Rnd * 255)
    N = Int(Rnd * 255)
    O = Int(Rnd * 255)
End Sub
