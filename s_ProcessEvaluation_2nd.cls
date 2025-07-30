VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "s_ProcessEvaluation_2nd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit '変数の宣言を強制


Const CAPTION_TRACK2_FILE_NAME_SOEJI           As String = "2" '字幕トラック２用のファイル名末尾につける添字
Const CAPTION_CHUKAN_FILE_NAME_SOEJI           As String = "tmp" 'リネーム時にファイル名カブリを防ぐために中間ファイルにする
'字幕切替のチェックボックスがクリックされた時
Private Sub CheckBox1_Click()

    Dim PosExt         As Long
    Dim StrFileName    As String


    StrFileName = ThisWorkbook.Name 'ファイル名を取得
    PosExt = InStrRev(StrFileName, ".") '.以前の文字数を取得



    '--- 拡張子を除いたパス（ファイル名）を格納する変数 ---'
    If (0 < PosExt) Then
        StrFileName = Left(StrFileName, PosExt - 1)
    End If



    '他のシートのチェックボックスと同期する
    If ThisWorkbook.Worksheets("姿勢重量点調査票").CheckBox1 = True Then    'チェックボックスにチェックが入った時
        ThisWorkbook.Worksheets("拳上腰曲げ膝曲げグラフ").CheckBox1 = True
    Else 'チェックボックスのチェックが外れた時
        ThisWorkbook.Worksheets("拳上腰曲げ膝曲げグラフ").CheckBox1 = False
    End If



    '---字幕ファイルをリネームする---
    'Name 変更前のファイル名 As 変更後のファイル名
    'トラック２用字幕を中間ファイルに変更
    Name ActiveWorkbook.Path & "\" & StrFileName & CAPTION_TRACK2_FILE_NAME_SOEJI & ".srt" _
      As ActiveWorkbook.Path & "\" & StrFileName & CAPTION_CHUKAN_FILE_NAME_SOEJI & ".srt"

    'トラック１用字幕をトラック２用字幕に変更
    Name ActiveWorkbook.Path & "\" & StrFileName & ".srt" _
      As ActiveWorkbook.Path & "\" & StrFileName & CAPTION_TRACK2_FILE_NAME_SOEJI & ".srt"

    '中間ファイルをトラック１用字幕に変更
    Name ActiveWorkbook.Path & "\" & StrFileName & CAPTION_CHUKAN_FILE_NAME_SOEJI & ".srt" _
      As ActiveWorkbook.Path & "\" & StrFileName & ".srt"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim currentCell As Range
    Dim row As Long
    Dim ajVal As Variant, ajAbove As Variant, ajBelow As Variant
    Dim aoVal As Variant, aqVal As Variant
    Dim ajTime As Double, ajAboveTime As Double, ajBelowTime As Double
    Dim aoTime As Double, aqTime As Double
    Static isEditing As Boolean

    If isEditing Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False
    isEditing = True

    For Each currentCell In Target
        row = currentCell.row

        ' === 1. AJ列：前後のセルと比較 ===
        If currentCell.Column = 36 Then
            ajVal = Cells(row, 36).Value
            If TryConvertToSeconds(ajVal, ajTime) Then

                ' 前の行と比較
                If row > 1 Then
                    ajAbove = Cells(row - 1, 36).Value
                    If TryConvertToSeconds(ajAbove, ajAboveTime) Then
                        If ajTime < ajAboveTime Then
                            MsgBox "作業開始時間は前の行よりも大きな値を入力してください", vbExclamation
                            currentCell.ClearContents
                            GoTo SafeExit
                        End If
                    End If
                End If

                ' 次の行と比較（空でない場合）
                If Not IsEmpty(Cells(row + 1, 36)) Then
                    ajBelow = Cells(row + 1, 36).Value
                    If TryConvertToSeconds(ajBelow, ajBelowTime) Then
                        If ajTime > ajBelowTime Then
                            MsgBox "作業開始時間は次の行よりも小さな値を入力してください", vbExclamation
                            currentCell.ClearContents
                            GoTo SafeExit
                        End If
                    End If
                End If

            End If
        End If

        ' === 2. AOとAQ：片方入力済なら比較 ===
        aoVal = Cells(row, 41).Value
        aqVal = Cells(row, 43).Value

        ' AOに入力 → AQがすでにあるなら比較
        If currentCell.Column = 41 And Not IsEmpty(aqVal) Then
            If TryConvertToSeconds(aoVal, aoTime) And TryConvertToSeconds(aqVal, aqTime) Then
                If aqTime <= aoTime Then
                    MsgBox "終了時間は開始時間より大きい値を入力してください", vbExclamation
                    currentCell.ClearContents
                    GoTo SafeExit
                End If
            End If
        End If

        ' AQに入力 → AOがすでにあるなら比較
        If currentCell.Column = 43 And Not IsEmpty(aoVal) Then
            If TryConvertToSeconds(aoVal, aoTime) And TryConvertToSeconds(aqVal, aqTime) Then
                If aqTime <= aoTime Then
                    MsgBox "終了時間は開始時間より大きい値を入力してください", vbExclamation
                    currentCell.ClearContents
                    GoTo SafeExit
                End If
            End If
        End If

    Next currentCell

SafeExit:
    isEditing = False
    Application.EnableEvents = True
End Sub


' === セルの値を秒に正確変換（時刻 or 数値対応）===
Private Function TryConvertToSeconds(val As Variant, ByRef result As Double) As Boolean
    On Error GoTo Fail
    If IsDate(val) Then
        result = CDbl(CDate(val)) * 86400 ' 時刻 → 秒（1日 = 86400秒）
        TryConvertToSeconds = True
        Exit Function
    ElseIf IsNumeric(val) Then
        result = CDbl(val) ' 数値は秒として直接扱う
        TryConvertToSeconds = True
        Exit Function
    End If
Fail:
    TryConvertToSeconds = False
End Function
