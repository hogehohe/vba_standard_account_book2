VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "s_Master_2nd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub SetCellsHW(MyWidth As Integer)
'ピクセル単位でセルの高さと幅を指定する
    'Range(Columns(7), Columns(Cells(6, 7).End(xlToRight).Column)).ColumnWidth = MyWidth * 0.118
    Dim mojisu As Double
    mojisu = MyWidth * (Cells(6, 7).ColumnWidth / Cells(6, 7).Width)
    Range(Columns(7), Columns(Cells(6, 7).End(xlToRight).Column)).ColumnWidth = mojisu
    'MsgBox ActiveWindow.PointsToScreenPixelsX(Cells(6, 7).ColumnWidth)
End Sub

Function GetWidthPoints() As Integer
    'MsgBox Range(Columns(7), Columns(Cells(6, 7).End(xlToRight).Column)).ColumnWidth
    GetWidthPoints = Int(Cells(6, 7).ColumnWidth / (Cells(6, 7).ColumnWidth / Cells(6, 7).Width))
End Function

Function activeCells() As Integer
    'MsgBox ActiveWindow.ActiveCell.Column
    activeCells = Cells.Column
End Function

Private Sub CommandButton1_Click()

End Sub
