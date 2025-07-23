VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAjsSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub SetCellsHW(MyWidth As Integer, ws As Worksheet)
'ピクセル単位でセルの高さと幅を指定する
    'Range(Columns(7), Columns(Cells(6, 7).End(xlToRight).Column)).ColumnWidth = MyWidth * 0.118
    Dim mojisu As Double
    mojisu = MyWidth * (ws.Cells(6, 7).ColumnWidth / ws.Cells(6, 7).Width)
'    ws.Range(ws.Columns(7), ws.Columns(ws.Cells(6, 7).End(xlToRight).Column)).ColumnWidth = mojisu
    '列幅を引数で指定するようにした。
     ws.Range(ws.Columns(7), ws.Columns(ws.Cells(6, 7).End(xlToRight).Column)).ColumnWidth = MyWidth / 10
    '20221228_下里
    'シートを2回拡大させた後、縮小してもフォントサイズが変わらないバグ
    'そもそもサイズを変えないほうが見やすいためコメントアウト
    'Call fontSize(MyWidth * 1.5, ws)
End Sub

Function GetWidthPoints() As Integer
'    MsgBox Cells(6, 7).ColumnWidth
'    MsgBox Cells(6, 7).Width
'    GetWidthPoints = Int(Cells(6, 7).Width * (Cells(6, 7).ColumnWidth / Cells(6, 7).Width))
    GetWidthPoints = Cells(6, 7).Width
End Function

Function activeCells() As Integer
    'MsgBox ActiveWindow.ActiveCell.Column
    Dim objPane As Pane
    Set objPane = ActiveWindow.Panes(4)
'    objPane.VisibleRange.Select
'    Stop
'    ActiveWindow.ActivePane.VisibleRange.Select
'    Stop
    
    If objPane.Index = ActiveWindow.ActivePane.Index Then
        activeCells = Selection.Column
    Else
        activeCells = 0
    End If
    
'    activeCells = Selection.Column
'    activeCells = ActiveWindow.ActivePane
End Function


Sub fontSize(size As Integer, ws As Worksheet)
'最後閾値を決めるが、とりあえず5以下の数値が来たらスルー
    If size < 6 Then
        Exit Sub
    ElseIf 11 < size Then
        size = 12
    End If
    
    '現在のアクティブセルを保持する
    Dim actCell As Range
    'セル選択用
    Dim r As Range
    ws.Activate
    Set actCell = ActiveCell
'    Set r = ws.Range(ws.Cells(23, 7), ws.Cells(23, ws.Columns.Count))
    Set r = ws.Range(ws.Cells(23, 7), ws.Cells(23, ws.Cells(23, ws.Columns.Count).End(xlToLeft).Column))
    r.Select
    
    With Selection.Font
        .size = size
    End With
    actCell.Select
End Sub

