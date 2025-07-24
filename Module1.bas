Option Explicit

'---------------------------------------------
'姿勢素点修正シートで使う定数
'---------------------------------------------
'1マスの秒数を定義
Const UNIT_TIME                         As Double = 0.1

'1秒の列
Const COLUMN_ZERO_NUM                   As Long = 6

'行
'信頼性上端
Const ROW_RELIABILITY_TOP               As Long = 2
'信頼性下端
Const ROW_RELIABILITY_BOTTOM            As Long = 7
'姿勢点上端
'2023/12/19育成G追記（レイアウト変更により2行分追加）
Const ROW_POSTURE_SCORE_TOP             As Long = 12 + 2
'姿勢点下端
'2023/12/19育成G追記（レイアウト変更により2行分追加）
Const ROW_POSTURE_SCORE_BOTTOM          As Long = 21 + 2


'2023/12/08 育成G小杉追記
'拳上_姿勢点
Const ROW_POSTURE_SCORE_KOBUSHIAGE      As Long = 10 + 2 '一旦姿勢素点の下側に表示する

'---------------------------------------------
'ポイント計算シートの列
'---------------------------------------------
'姿勢点が保存されている列
Const COLUMN_DATA_RESULT_ORIGIN         As Long = 203
'姿勢点が保存されている列 2023/12/12 育成G追記
Const COLUMN_POSTURE_SCORE_ALL          As Long = 203

'2023/12/11 育成G小杉追記 条件A(拳上)が保存されている列
Const COLUMN_POSTURE_SCORE_KOBUSHIAGE   As Long = 245

'信頼性が保存されている列
'測定
Const COLUMN_MEASURE_SECTION            As Long = 204
'推定
Const COLUMN_PREDICT_SECTION            As Long = 205
'除外区間
Const COLUMN_REMOVE_SECTION             As Long = 206
'強制区間
Const COLUMN_FORCED_SECTION             As Long = 207
'強制区間 2023/12/12 育成G追記
Const COLUMN_FORCED_SECTION_TOTAL       As Long = 207
'元データ
Const COLUMN_DATA_RESULT_FIX            As Long = 208
'元データ 2023/12/12 育成G追記
Const COLUMN_BASE_SCORE                 As Long = 208
'姿勢素点緑色
Const COLUMN_POSTURE_GREEN              As Long = 209
'姿勢素点黄色
Const COLUMN_POSTURE_YELLOW             As Long = 210
'姿勢素点赤色
Const COLUMN_POSTURE_RED                As Long = 211
'欠損
Const COLUMN_MISSING_SECTION            As Long = 219
'拳上強制区間 2023/12/12 育成G追記
Const COLUMN_FORCED_SECTION_KOBUSHIAGE  As Long = 223

'---------------------------------------------
'姿勢素点修正シート　関連
'---------------------------------------------
'LIMIT_COLUMNの設定値は3の倍数とする必要がある
'30fps×60秒×9分＝16200
'姿勢素点修正シートは9分毎に次のシートを生成する
Const LIMIT_COLUMN             As Long = 16200

Const SHEET_LIMIT_COLUMN       As Long = LIMIT_COLUMN + COLUMN_ZERO_NUM

'時刻表示セルの幅
Const TIME_WIDTH               As Long = 30
'時刻表示セルが存在する行
'2023/12/19育成G追記（レイアウト変更により2行分追加）
Const TIME_ROW                 As Long = 25 + 2
'1つ目の時刻表示セルの左端
Const TIME_COLUMN_LEFT         As Long = 22
'1つ目の時刻表示セルの右端
Const TIME_COLUMN_RIGHT        As Long = 51
'データ調整用のテーブルの下端
'2023/12/19育成G追記（レイアウト変更により2行分追加）
Const BOTTOM_OF_TABLE          As Long = 26 + 2

'列幅用の列挙
Private Enum widthSize
    Small = 1
    Medium = 2
    Large = 4
    LL = 6
End Enum

'列幅調整ボタン名前
Const EXPANDBTN_NAME           As String = "expandBtn"
Const REDUCEBTN_NAME           As String = "reduceBtn"

'---------------------------------------------
'複数モジュールで使用する変数
'---------------------------------------------
'再生・停止ボタンで使用
'指定した時間が経過すると処理を実行する
Private ResTime     As Date
Private scrollTime  As Double

Public postureFlag(1 To 100) As Boolean 'ページ最大100として仮設定


'処理時間短縮のため、更新をストップ
' 引数1 ：なし
' 戻り値：なし
Function stopUpdate()
    '表示・更新をオフにする
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Function


'処理時間短縮のため、更新をリスタート
' 引数1 ：なし
' 戻り値：なし
Function restartUpdate()
    '表示・更新をオンに戻す
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Function


'------------------------------------------------------------
' 姿勢素点修正シートのテンプレート生成
'
' ・映像フレーム数とFPSをもとに作業時間などを算出
' ・後続処理で罫線や時刻のオートフィルを行う前準備
'------------------------------------------------------------
Sub autoFillTemplate()

    ' ラベル列と時間単位列数の初期値
    Dim startColumnNum      As Long      ' ラベル列の開始列番号
    Dim unit10SecColumnNum  As Long      ' 10秒分の列数（単位時間で割る）

    ' 作業時間やFPSなど時間関連の一時変数
    Dim workTime    As Double            ' 作業時間（秒）
    Dim fps         As Double            ' フレームレート
    Dim maxFrameNum As Long              ' 最大フレーム番号

    ' 罫線描画用の列番号・アルファベット表記
    Dim ruleLineColumnNum As Long
    Dim ruleLineColumnAlf As String

    ' ラベル列の初期位置（0列目の次）
    startColumnNum = COLUMN_ZERO_NUM + 1

    ' 10秒分の列数を単位時間で割って算出（例：UNIT_TIME = 0.5秒 → 20列）
    unit10SecColumnNum = 10 / UNIT_TIME

    ' 「ポイント計算シート」から映像情報を取得
    With ThisWorkbook.Sheets("ポイント計算シート")
        ' フレームレートの取得（例：30fps など）
        fps = getFps()

        ' 最終行のフレーム番号を取得（列Aの最下行）
        maxFrameNum = getLastRow()

        ' 作業時間（秒）＝フレーム数 ÷ FPS
        ' ※現在は未使用。必要になればコメントアウト解除
        ' workTime = CDbl(maxFrameNum / fps)
    End With

End Sub


'------------------------------------------------------------
' 罫線の複製処理
'
' 「G2:EZ26」のレイアウトを基準として、右方向へ罫線や装飾を複製
'
' 引数:
'   ws       - 対象のワークシート
'   endline  - 罫線を引く対象の最終列（上限補正あり）
'------------------------------------------------------------
Private Sub autoFillLine(ws As Worksheet, endline As Long)
    Dim ruleLineColumnNum   As Long         ' 実際に処理対象とする列番号
    Dim ruleLineColumnAlf   As String       ' 列番号をアルファベット表記に変換したもの
    Dim frame30Mod          As Long         ' フレーム調整用（未使用）

    ' 上限を超える列数の場合、制限値までに抑える
    ruleLineColumnNum = endline
    If ruleLineColumnNum > SHEET_LIMIT_COLUMN Then
        ruleLineColumnNum = SHEET_LIMIT_COLUMN
    End If

    ' 行数が24→26に変更されたことに伴い、余り計算（未使用）
    ' ※必要ならこの値を使ってruleLineColumnNumを30の倍数に調整する想定
    frame30Mod = (ruleLineColumnNum + 26) Mod 30

    ' オートフィル先の終了列（アルファベット表記）を取得
    ruleLineColumnAlf = Split(ws.Cells(1, ruleLineColumnNum).Address(True, False), "$")(0)

    ' 対象範囲をクリア（色や罫線含め全消去）
    Call clear(ws)

    ' レイアウトのベース範囲（G2:EZ26）を右方向へオートフィル
    ws.Range("G2:EZ26").AutoFill _
        Destination:=ws.Range("G2:" & ruleLineColumnAlf & 26), _
        Type:=xlFillDefault

    ' 不要な範囲の罫線を消去（右端からXFD列まで）
    ruleLineColumnAlf = Split(ws.Cells(1, ruleLineColumnNum + 1).Address(True, False), "$")(0)
    ws.Range(ruleLineColumnAlf & "2:XFD26").Borders.LineStyle = xlLineStyleNone

End Sub


'------------------------------------------------------------
' 時刻を時間セルに挿入する処理
'
' 引数:
'   ws      - 対象のワークシート
'   min     - 分単位（例: 5 → 00:05:01のように開始）
'   endclm  - 処理対象の最終列
'------------------------------------------------------------
Private Sub autoFillTime(ws As Worksheet, min As Long, endclm As Long)

    ' 変数定義
    Dim tmp         As Long
    Dim boldcnt     As Long: boldcnt = 0
    Dim r           As Range
    Dim timeStr     As String
    Dim frame30Mod  As Long
    Dim i           As Long

    ' 最終列の調整（上限制限を考慮）
    tmp = endclm
    If 30 <= tmp - TIME_COLUMN_LEFT Then
        If tmp > LIMIT_COLUMN Then
            tmp = LIMIT_COLUMN
        End If
    End If

    ' 結合セルがあるとオートフィル時にエラーになるため、事前に解除・クリア
    ws.Range(ws.Cells(TIME_ROW, 12), ws.Cells(TIME_ROW, 16384)).Clear

    ' 時間セルの書式設定と結合処理
    For i = TIME_COLUMN_LEFT To SHEET_LIMIT_COLUMN Step TIME_WIDTH
        Set r = ws.Range(ws.Cells(TIME_ROW, i), ws.Cells(TIME_ROW, i + TIME_WIDTH - 1))

        boldcnt = boldcnt + 1

        With r
            .Merge True                      ' セル結合（横方向）
            .Orientation = -90              ' 縦書き（90度回転）
            .ReadingOrder = xlContext       ' 文字方向：自動判定
            .HorizontalAlignment = xlCenter ' 横位置：中央
            .NumberFormatLocal = "hh:mm:ss" ' 時刻形式にする

            ' 5回に1回は太字にする
            If boldcnt = 5 Then
                .Font.FontStyle = "bold"
                boldcnt = 0
            End If
        End With
    Next i

    ' 初期の2つの時刻を直接入力（例: 00:05:01, 00:05:02）
    timeStr = "00:" & Format(min, "00") & ":01"
    ws.Range(ws.Cells(TIME_ROW, TIME_COLUMN_LEFT), _
             ws.Cells(TIME_ROW, TIME_COLUMN_RIGHT)).Value = timeStr

    timeStr = "00:" & Format(min, "00") & ":02"
    ws.Range(ws.Cells(TIME_ROW, TIME_COLUMN_LEFT + TIME_WIDTH), _
             ws.Cells(TIME_ROW, TIME_COLUMN_RIGHT + TIME_WIDTH)).Value = timeStr

    ' フレーム幅の余りを調整し、オートフィル範囲をフレーム単位に丸める
    frame30Mod = (tmp - TIME_COLUMN_LEFT) Mod TIME_WIDTH
    If frame30Mod Then
        tmp = tmp + TIME_WIDTH - frame30Mod
    End If

    ' 2つ目の時刻より右側が存在する場合にのみ、オートフィルを実行
    If (TIME_COLUMN_LEFT + TIME_WIDTH) < tmp Then
        ws.Range(ws.Cells(TIME_ROW, TIME_COLUMN_LEFT), _
                 ws.Cells(TIME_ROW, TIME_COLUMN_RIGHT + TIME_WIDTH)).AutoFill _
            Destination:=ws.Range(ws.Cells(TIME_ROW, TIME_COLUMN_LEFT), _
                                  ws.Cells(TIME_ROW, tmp - 1)), _
            Type:=xlFillValues
    End If

End Sub



'単位時間当たり最も多い姿勢点・信頼性を調べてセルに色を塗る
'processingRange　1:選択範囲（部分的に強制をキャンセル） 2:全体 else:特定の1セルごと
Sub paintPostureScore(processingRange As Long)
    '---------------------------------------------
    'RGBを指定するための変数を定義
    '---------------------------------------------

    '信頼性
    Dim colorMeasureSection    As String '水色
    Dim colorPredictSection    As String '黄色
    Dim colorMissingSection    As String 'ピンク
    Dim colorForcedSection     As String '青色
    Dim colorRemoveSection     As String 'グレー

    '姿勢点
    Dim colorResultGreen       As String '緑色
    Dim colorResultYellow      As String '黄色
    Dim colorResultRed         As String '赤色
    Dim colorResultGlay        As String 'グレー
    Dim colorResultWhite       As String '白色 20221219_下里
    Dim colorResultBrown       As String '茶色 20221222_下里
    Dim colorResultOFFGlay     As String 'グレー 20221222_下里

    '---------------------------------------------
    '変数に色をセット
    '---------------------------------------------
    '1:測定、2:推定、3:欠損、4:強制、5:除外
    '信頼性
    colorMeasureSection = RGB(0, 176, 240)   '水色
    colorPredictSection = RGB(252, 246, 0)   '黄色
    colorMissingSection = RGB(255, 124, 128) 'ピンク
    colorForcedSection = RGB(0, 51, 204)     '青色
    colorRemoveSection = RGB(191, 191, 191)  'グレー
    '姿勢点
    colorResultGreen = RGB(0, 176, 80)       '緑色
    colorResultYellow = RGB(255, 192, 0)     '黄色
    colorResultRed = RGB(192, 0, 0)          '赤色
    colorResultGlay = RGB(191, 191, 191)     'グレー
    colorResultWhite = RGB(255, 255, 255)    '白色
    colorResultBrown = RGB(64, 0, 0)         '茶色
    colorResultOFFGlay = RGB(217, 217, 217)  '判定オフ用のグレー


    '---------------------------------------------
    '配列
    '---------------------------------------------
    'ポイント計算シートの姿勢点を保管
    Dim postureScoreDataArray()           As Long
    '2023/12/11　育成G小杉追記-------------
    Dim postureScoreDataArray_A()  As Long
    '1 ~ 11点のフレーム数をそれぞれ合計
    Dim postureScoreCounterArray(11)      As Long
    ' 2023/12/11　育成G小杉追記 拳上げ点数--
    Dim postureScoreCounterArray_A(0 To 1) As Integer
    '---------------------------------------

    'ポイント計算シートの信頼性を保管
    '1:測定、2:推定、3:欠損
    Dim reliabilityDataArray()     As Long
    '信頼性1 ~ 3のフレーム数をそれぞれ合計
    Dim reliabilityCounterArray(3) As Long

    '---------------------------------------------
    'その他の変数
    '---------------------------------------------
    'ポイント計算シート最大行数の変数定義
    Dim RowNumCount As Long
    Dim maxRowNum      As Long

    '変数定義
    Dim wholeStartCount As Long
    Dim PointComp0       As Long
    Dim PointComp1       As Long
    Dim PointComp2      As Long

    Dim fps        As Double

    '単位時間の繰り返し処理の開始終了地点を定義
    Dim wholeStart As Long
    Dim wholeEnd   As Long

    '姿勢点一時記憶用の変数
    Dim postureScoreFlag      As Long
    ' 2023/12/11　育成G小杉追記--
    Dim postureScoreFlag_A      As Integer
    '---------------------------------------
    '単位時間の中で一番多い姿勢点を保管
    Dim mostOftenPostureScore As Long
    ' 2023/12/11　育成G小杉追記 拳上げ点数--
    Dim mostOftenPostureScore_A As Integer
    '---------------------------------------
    '信頼性一時記憶用の変数
    Dim reliabilityFlag       As Long
    '単位時間の中で一番多い信頼性を保管
    Dim mostOftenReliability  As Long

    '次ページにいく制限
    Dim thisPageLimit As Long
    thisPageLimit = LIMIT_COLUMN
    '前のページの最終列を保存する
    Dim preClm As Long
    preClm = 0
    Call stopUpdate

    Dim baseClm As Long
    Dim shtPage As Long

    '動画時間(秒)により列の初期幅を変更する

    Dim wSize     As widthSize
        '---------------------------------------------
    '変数、配列に値を入力
    '---------------------------------------------
    With ThisWorkbook.Sheets("ポイント計算シート")
        '最終行を取得
        maxRowNum = getLastRow()
        '配列の最後尾
'        余分を削除
        maxRowNum = maxRowNum - 1
        '配列を再定義
        ReDim postureScoreDataArray(maxRowNum, 0)
        '2023/12/11　育成G小杉追記-------------
        ReDim postureScoreDataArray_A(maxRowNum, 0) As Long
        '--------------------------------------
        '信頼性区間用
        ReDim reliabilityDataArray(maxRowNum, 0)

        '配列の中に値を入れる
        For RowNumCount = 1 To maxRowNum
'        For i = 1 To 10
            '姿勢点の列を配列に入れる
            '配列は0から始まるため+1、2行目から使用するため+1
            postureScoreDataArray(RowNumCount - 1, 0) = .Cells(RowNumCount + 1, COLUMN_DATA_RESULT_ORIGIN).Value
            ' 2023/12/11育成G小杉追記 オリジナル拳上げデータ参照----------
            postureScoreDataArray_A(RowNumCount - 1, 0) = .Cells(RowNumCount + 1, COLUMN_POSTURE_SCORE_KOBUSHIAGE - 1).Value
            '----------------------------------
            '信頼性を配列に入れる
            '1:測定、2:推定、3:欠損

            If .Cells(RowNumCount + 1, COLUMN_MEASURE_SECTION).Value > 0 Then
                reliabilityDataArray(RowNumCount, 0) = 1
            End If
            If .Cells(RowNumCount + 1, COLUMN_PREDICT_SECTION).Value > 0 Then
                reliabilityDataArray(RowNumCount, 0) = 2
            End If
            If .Cells(RowNumCount + 1, COLUMN_MISSING_SECTION).Value > 0 Then
                reliabilityDataArray(RowNumCount, 0) = 3
            End If
        Next
        'フレームレートを取得
        fps = getFps()
        Dim video_sec As Double: video_sec = wholeEnd / fps

    End With 'With ThisWorkbook.Sheets("ポイント計算シート")


    '---------------------------------------------
    '処理範囲を決める
    '---------------------------------------------
    'キャンセル(戻る)ボタンから呼ばれたとき


    If processingRange = 1 Then
        'アクティブセルの一番左が6列目以下の時
        'エラーメッセージを出して処理をやめる

        shtPage = calcSheetNamePlace(ThisWorkbook.ActiveSheet)
        baseClm = LIMIT_COLUMN * shtPage


        'pageLimitを次のページとなる閾値まで更新
        thisPageLimit = (shtPage + 1) * LIMIT_COLUMN
        preClm = (LIMIT_COLUMN * shtPage) * -1

        Dim lCol As Long, rCol As Long
        If Not CropSelectionToDataArea(lCol, rCol) Then
            MsgBox "範囲外です", vbCritical
            Exit Sub
        End If

        wholeStart = lCol - COLUMN_ZERO_NUM + baseClm
        wholeEnd = rCol - COLUMN_ZERO_NUM + baseClm

        If wholeStart < 1 Then
            wholeStart = 1
        End If

    'メインの処理から呼ばれたとき
    ElseIf processingRange = 2 Then

        'すでに計算シートがあるとテスト用関数からcreateSheetすると増殖するため
        Call DeleteSheet(0)
        Call createSheet(0)

        '先頭から
        wholeStart = 1
        '末尾まで
        wholeEnd = maxRowNum

        '基準のworkSheet、合わせて初期化
        ThisWorkbook.Sheets("姿勢素点修正シート").Activate
        preClm = 0

        '動画が短いとオートフィルでエラーが出るため、エラー処理を追加20231004早川
        If maxRowNum >= 150 Then
            Call autoFillLine(ActiveSheet, wholeEnd + COLUMN_ZERO_NUM) '230206 + COLUMN_ZERO_NUMを追加
            Call autoFillTime(Worksheets("姿勢素点修正シート"), 0, wholeEnd)
        End If

        Call addPageShape(ActiveSheet, False, True)

        '15秒以下を列幅2とする
        If video_sec <= 15 Then
            wSize = LL
            Call changeBtnState(EXPANDBTN_NAME, False)
            Call changeBtnState(REDUCEBTN_NAME, True)
        Else
            wSize = Small
            Call changeBtnState(REDUCEBTN_NAME, False)
            Call changeBtnState(EXPANDBTN_NAME, True)
        End If

        Call DataAjsSht.SetCellsHW(CInt(wSize), Worksheets("姿勢素点修正シート"))

    '除外があるフレームに強制を上書きしたとき（１セルずつ実行）
    Else
        shtPage = calcSheetNamePlace(ThisWorkbook.ActiveSheet)
        baseClm = LIMIT_COLUMN * shtPage


        'pageLimitを次のページとなる閾値まで更新
        thisPageLimit = (shtPage + 1) * LIMIT_COLUMN
        preClm = (LIMIT_COLUMN * shtPage) * -1

        '20230126_下里
        '選択範囲が6列以内の場合、データの左端になるように修正
'        If Selection.Column <= COLUMN_ZERO_NUM Then
'            wholeStart = 1
'        Else
        wholeStart = processingRange - COLUMN_ZERO_NUM + baseClm
'        End If
        '
        wholeEnd = wholeStart
    End If

    For wholeStartCount = wholeStart To wholeEnd
        '姿勢点のカウンターをリセット
        Erase postureScoreCounterArray
        '2023/12/11　育成G小杉追記 -----------
        Erase postureScoreCounterArray_A
        '-------------------------------------
        '信頼性のカウンターをリセット
        Erase reliabilityCounterArray

        '姿勢点を確認
        postureScoreFlag = postureScoreDataArray(wholeStartCount - 1, 0)
        '2023/12/11　育成G小杉追記 -----------
        postureScoreFlag_A = postureScoreDataArray_A(wholeStartCount - 1, 0)
        '-------------------------------------
        '姿勢点フラグを立てる
        postureScoreCounterArray(postureScoreFlag) = 1
        '2023/12/11　育成G小杉追記 -----------
        postureScoreCounterArray_A(postureScoreFlag_A) = 1
        '-------------------------------------
        '信頼性を確認
'        reliabilityFlag = reliabilityDataArray(i - 1, 0)230209
        reliabilityFlag = reliabilityDataArray(wholeStartCount, 0)
        '信頼性フラグを立てる
        reliabilityCounterArray(reliabilityFlag) = 1



        '---------------------------------------------
        'フレーム数が最も多いものを探す
        '---------------------------------------------
        '初期は1
        mostOftenPostureScore = 1

        '姿勢点1～10の先頭から順に比較
        For PointComp0 = 2 To 10
            'フレーム数の合計が多い姿勢点を選ぶ
            '合計が同じ場合は辛い姿勢を優先する
            If postureScoreCounterArray(mostOftenPostureScore) <= postureScoreCounterArray(PointComp0) Then
                mostOftenPostureScore = PointComp0
            End If
        Next

        '2023/12/11　育成G小杉追記 拳上げ箇所追加-----------
         '姿勢点0～1の先頭から順に比較
        For PointComp1 = 0 To 1
            'フレーム数の合計が多い姿勢点を選ぶ
            '合計が同じ場合は辛い姿勢を優先する

            '拳上げ
            If postureScoreCounterArray_A(mostOftenPostureScore_A) <= postureScoreCounterArray_A(PointComp1) Then
                 mostOftenPostureScore_A = PointComp1
            End If
        Next
        '--------------------------------------------------

        '初期は1
        mostOftenReliability = 1
            '信頼性1～3の先頭から順に比較
            '1:測定、2:推定、3:欠損
        For PointComp2 = 2 To 3
            'フレーム数の合計が多い姿勢点を選ぶ
            '合計が同じ場合は信頼性が低い方を優先する
            If reliabilityCounterArray(mostOftenReliability) <= reliabilityCounterArray(PointComp2) Then
                mostOftenReliability = PointComp2
            End If
        Next

        'active sheetを変更する基準
        If wholeStartCount <= thisPageLimit Then
            '何もしない
        Else
            ThisWorkbook.ActiveSheet.Next.Activate
            If InStr(ThisWorkbook.ActiveSheet.Name, "姿勢素点修正シート") > 0 Then
                '何もしない
            Else
                '戻る
                ThisWorkbook.ActiveSheet.Previous.Activate
                Call createSheet(0)
            End If
            '更新
            thisPageLimit = thisPageLimit + LIMIT_COLUMN
            preClm = preClm - LIMIT_COLUMN
            Call clear(ActiveSheet)
            Call autoFillLine(ActiveSheet, wholeEnd - COLUMN_ZERO_NUM)
            Call autoFillTime(ThisWorkbook.ActiveSheet, ((thisPageLimit / LIMIT_COLUMN) - 1) * 9, wholeEnd - wholeStartCount)
            Call addPageShape(ActiveSheet, True, True)
        End If
        '---------------------------------------------
        '姿勢素点修正シートのセルに色を塗る
        '---------------------------------------------
        With ThisWorkbook.ActiveSheet
            'カラーを保持する変数
            Dim colorStr As String
            '最も多かった姿勢点に応じて
            'セルの選択範囲、色を変更
            '1,2点の場合は緑
            If mostOftenPostureScore <= 2 Then
                colorStr = colorResultGreen

            '3～5点の場合は黄
            ElseIf mostOftenPostureScore >= 3 _
            And mostOftenPostureScore <= 5 Then
                colorStr = colorResultYellow

            '6～10点の場合は赤
            ElseIf mostOftenPostureScore >= 6 _
            And mostOftenPostureScore <= 10 Then
                colorStr = colorResultRed
            End If

            '2023/12/8　育成G小杉追記 拳上げ箇所追加-----------
            Dim colorStr1 As String '条件A
            '-------------条件A
            '0点の場合、白
            If mostOftenPostureScore_A = 0 Then
                colorStr1 = colorResultWhite


            '1点の場合、赤
            ElseIf mostOftenPostureScore_A = 1 Then
                colorStr1 = colorResultRed

            End If
            '------------------------------------------

            '色をクリア
            .Range _
            ( _
                .Cells(ROW_POSTURE_SCORE_BOTTOM, COLUMN_ZERO_NUM + wholeStartCount + preClm), _
                .Cells(ROW_POSTURE_SCORE_TOP, COLUMN_ZERO_NUM + wholeStartCount + preClm) _
            ) _
            .Interior.ColorIndex = 0


            '色を塗る
            .Range _
            ( _
                .Cells(ROW_POSTURE_SCORE_BOTTOM, COLUMN_ZERO_NUM + wholeStartCount + preClm), _
                .Cells(ROW_POSTURE_SCORE_BOTTOM - mostOftenPostureScore + 1, COLUMN_ZERO_NUM + wholeStartCount + preClm) _
            ) _
            .Interior.Color = colorStr

            ' データ信頼性・姿勢素点のセルに強制的に白を塗る
            .Range _
            ( _
                .Cells(ROW_RELIABILITY_TOP, COLUMN_ZERO_NUM), _
                .Cells(ROW_POSTURE_SCORE_TOP, COLUMN_ZERO_NUM) _
            ) _
            .Interior.Color = colorResultWhite

            '''''''''''''''''''''''''''''''''''''
            '▽拳上を一時的に除外
            '
            ''2023/12/8　育成G小杉追記 拳上げ箇所追加-----------
            '
            '    .Range _
            '    ( _
            '        .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE, COLUMN_ZERO_NUM + wholeStartCount + preClm), _
            '        .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE, COLUMN_ZERO_NUM + wholeStartCount + preClm) _
            '    ) _
            '    .Interior.Color = colorStr1
            ''--------------------------------------------------
            '▽END_拳上を一時的に除外

            '最も多かった信頼性に応じて
            '色を変更
            '1:測定、2:推定、3:欠損
            If mostOftenReliability = 1 Then
                colorStr = colorMeasureSection
            ElseIf mostOftenReliability = 2 Then
                colorStr = colorPredictSection
            ElseIf mostOftenReliability = 3 Then
                colorStr = colorMissingSection
            End If

            .Range _
            ( _
                .Cells(ROW_RELIABILITY_TOP, COLUMN_ZERO_NUM + wholeStartCount + preClm), _
                .Cells(ROW_RELIABILITY_BOTTOM, COLUMN_ZERO_NUM + wholeStartCount + preClm) _
            ) _
            .Interior.Color = colorStr

        End With 'With ThisWorkbook.Sheets("姿勢素点修正シート")
    Next 'i = wholeStart To wholeEnd

    ' キャンセルボタン以外からの処理の時
    If 1 < processingRange Then
        If calcSheetNamePlace(ThisWorkbook.ActiveSheet) = 0 Then ' 0 = Base sheet
            Call addPageShape(ActiveSheet, False, False)
        Else
            Call addPageShape(ActiveSheet, True, False)
        End If
    End If

    '各シートを更新
    Call checkReliabilityRatio
    Call restartUpdate

End Sub


'『全体を処理』ボタンが押されたとき
'全体の姿勢点を計算して、色を塗る
Sub paintAll()
    Call paintPostureScore(2)
End Sub


'『Cancel』ボタンが押されたとき
'選択範囲の姿勢点を計算して、色を塗る（強制ボタンのキャンセル）
Sub paintSelected()

    ' 選択範囲の左端の列が「0列目（＝姿勢点列）」以下なら処理をスキップ
    If DataAjsSht.activeCells <= COLUMN_ZERO_NUM Then
        Exit Sub
    End If

    ' 選択範囲のみ再描画
    paintPostureScore 1

End Sub


'塗りつぶしを全てクリア
Sub clear(ws As Worksheet)
    ws _
    .Range _
    ( _
        Cells(ROW_RELIABILITY_TOP, COLUMN_ZERO_NUM + 1), _
        Cells( _
            ROW_POSTURE_SCORE_BOTTOM, _
            Cells(ROW_POSTURE_SCORE_BOTTOM, COLUMN_ZERO_NUM + 1).End(xlToRight).Column _
        ) _
    ) _
    .Interior.ColorIndex = 0
End Sub


'------------------------------------------------------------
' 姿勢点の強制修正処理
'
' 引数:
'   postureScorebutton - 押されたボタンの点数（0〜10:強制, -1:リセット, 99:除外）
'------------------------------------------------------------
Sub forceResult(postureScorebutton As Long)

    ' 色設定：信頼性
    Dim colorMeasureSection As String
    Dim colorPredictSection As String
    Dim colorMissingSection As String
    Dim colorForcedSection  As String
    Dim colorResetSection   As String

    ' 色設定：姿勢点
    Dim colorResultGreen    As String
    Dim colorResultYellow   As String
    Dim colorResultRed      As String
    Dim colorResultGlay     As String
    Dim colorResultWhite    As String
    Dim colorResultOFFGlay  As String

    ' RGBの割り当て
    colorMeasureSection = RGB(0, 176, 240)
    colorPredictSection = RGB(252, 246, 0)
    colorMissingSection = RGB(255, 124, 128)
    colorForcedSection  = RGB(0, 51, 204)
    colorResetSection   = RGB(191, 191, 191)

    colorResultGreen    = RGB(0, 176, 80)
    colorResultYellow   = RGB(255, 192, 0)
    colorResultRed      = RGB(192, 0, 0)
    colorResultGlay     = RGB(191, 191, 191)
    colorResultWhite    = RGB(255, 255, 255)
    colorResultOFFGlay  = RGB(217, 217, 217)

    ' 現在のシート位置から列オフセットを計算
    Dim shtPage As Long: shtPage = calcSheetNamePlace(ThisWorkbook.ActiveSheet)
    Dim baseClm As Long: baseClm = LIMIT_COLUMN * shtPage

    Dim lCol As Long, rCol As Long
    Dim MinLeftCell As Variant, MaxRightCell As Variant

    If CropSelectionToDataArea(lCol, rCol) Then
        MinLeftCell = lCol
        MaxRightCell = rCol

        With ThisWorkbook.ActiveSheet
            ' 選択範囲をクリア
            .Range(.Cells(ROW_POSTURE_SCORE_TOP, MinLeftCell), _
                   .Cells(ROW_POSTURE_SCORE_BOTTOM, MaxRightCell)).Interior.ColorIndex = 0

            ' リセットボタン処理（postureScorebutton = -1）
            If postureScorebutton = -1 Then
                postureUpdate MinLeftCell + baseClm, MaxRightCell + baseClm, 0, CInt(postureScorebutton)
                paintPostureScore 1
                postureFlag(shtPage + 1) = False

            ' 強制処理（0〜10, 99）
            ElseIf postureScorebutton >= 0 Then
                postureUpdate MinLeftCell + baseClm, MaxRightCell + baseClm, 1, CInt(postureScorebutton)

                If postureScorebutton = 99 Then
                    If lCol <= 6 Then postureFlag(shtPage + 1) = True

                    ' 姿勢点 + 信頼性セルに除外色を適用
                    .Range(.Cells(ROW_RELIABILITY_TOP, MinLeftCell), _
                           .Cells(ROW_RELIABILITY_BOTTOM, MaxRightCell)).Interior.Color = colorResetSection
                    .Range(.Cells(ROW_POSTURE_SCORE_BOTTOM, MinLeftCell), _
                           .Cells(ROW_POSTURE_SCORE_TOP, MaxRightCell)).Interior.Color = colorResetSection

                Else
                    ' 姿勢点別の色分け
                    Select Case postureScorebutton
                        Case 1 To 2
                            .Range(.Cells(ROW_POSTURE_SCORE_BOTTOM, MinLeftCell), _
                                   .Cells(ROW_POSTURE_SCORE_BOTTOM - postureScorebutton + 1, MaxRightCell)).Interior.Color = colorResultGreen
                        Case 3 To 5
                            .Range(.Cells(ROW_POSTURE_SCORE_BOTTOM, MinLeftCell), _
                                   .Cells(ROW_POSTURE_SCORE_BOTTOM - postureScorebutton + 1, MaxRightCell)).Interior.Color = colorResultYellow
                        Case 6 To 10
                            .Range(.Cells(ROW_POSTURE_SCORE_BOTTOM, MinLeftCell), _
                                   .Cells(ROW_POSTURE_SCORE_BOTTOM - postureScorebutton + 1, MaxRightCell)).Interior.Color = colorResultRed
                    End Select

                    ' 信頼性セルに強制色
                    .Range(.Cells(ROW_RELIABILITY_TOP, MinLeftCell), _
                           .Cells(ROW_RELIABILITY_BOTTOM, MaxRightCell)).Interior.Color = colorForcedSection
                End If

                ' 信頼性チェック更新
                checkReliabilityRatio

                ' 信頼性・姿勢点の左端タイトルセルを白でリセット
                .Range(.Cells(ROW_RELIABILITY_TOP, COLUMN_ZERO_NUM), _
                       .Cells(ROW_POSTURE_SCORE_TOP, COLUMN_ZERO_NUM)).Interior.Color = colorResultWhite
            End If
        End With
    Else
        MsgBox "範囲はグラフ内から選択してください", vbOKOnly + vbCritical, "範囲選択エラー"
    End If
End Sub


'------------------------------------------------------------
' 拳上げ・膝曲げの姿勢点強制修正処理
'
' 引数:
'   postureScorebutton - 押されたボタンの点数（0/1:強制, -1:リセット, 99:除外）
'------------------------------------------------------------
Sub forceResult_Kobushiage(postureScorebutton As Integer)

    ' 色設定：信頼性
    Dim colorMeasureSection As String
    Dim colorPredictSection As String
    Dim colorMissingSection As String
    Dim colorForcedSection As String
    Dim colorRemoveSection As String

    ' 色設定：姿勢点
    Dim colorResultGreen As String
    Dim colorResultYellow As String
    Dim colorResultRed As String
    Dim colorResultGlay As String
    Dim colorResultWhite As String
    DIm colorResultBrown As String
    Dim colorResultOFFGlay As String

    ' RGBの割り当て
    colorMeasureSection = RGB(0, 176, 240)
    colorPredictSection = RGB(252, 246, 0)
    colorMissingSection = RGB(255, 124, 128)
    colorForcedSection  = RGB(0, 51, 204)
    colorRemoveSection  = RGB(191, 191, 191)

    colorResultGreen    = RGB(0, 176, 80)
    colorResultYellow   = RGB(255, 192, 0)
    colorResultRed      = RGB(192, 0, 0)
    colorResultGlay     = RGB(191, 191, 191)
    colorResultWhite    = RGB(255, 255, 255)
    colorResultBrown    = RGB(64, 0, 0)
    colorResultOFFGlay  = RGB(217, 217, 217)

    Dim baseClm As Long: baseClm = LIMIT_COLUMN * shtPage

    Dim lCol As Long, rCol As Long
    Dim MinLeftCell As Variant, MaxRightCell As Variant
    Dim k As Long, m As Long

    If CropSelectionToDataArea(lCol, rCol) Then
        MinLeftCell = lCol
        MaxRightCell = rCol

        With ThisWorkbook.ActiveSheet

            ' 選択範囲の値取得（未使用）
            Dim SelectCells As Variant
            SelectCells = .Range(.Cells(Selection.Row, Selection.Column), _
                                 .Cells(Selection.Row, Selection.Column + Selection.Columns.Count - 1)).Value

            ' リセット処理（戻るボタン）
            If postureScorebutton = -1 Then
                postureUpdate_Kobushiage MinLeftCell + baseClm, MaxRightCell + baseClm, 0, CInt(postureScorebutton)
                paintPostureScore 1

            ' 強制処理（0/1/99）
            ElseIf postureScorebutton >= 0 Then
                postureUpdate_Kobushiage MinLeftCell + baseClm, MaxRightCell + baseClm, 1, CInt(postureScorebutton)

                If postureScorebutton = 99 Then
                    ' 除外色適用
                    .Range(.Cells(ROW_RELIABILITY_TOP, MinLeftCell), _
                           .Cells(ROW_RELIABILITY_BOTTOM, MaxRightCell)).Interior.Color = colorRemoveSection

                    For k = 1 To 3
                        .Range(.Cells(ROW_POSTURE_SCORE_KOBUSHIAGE - 2 + 2 * k, MinLeftCell), _
                               .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE - 2 + 2 * k, MaxRightCell)).Interior.Color = colorResultGlay
                    Next

                Else
                    ' 除外状態だったセルを元の色に戻す
                    For m = MinLeftCell To MaxRightCell
                        If .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE, m).Interior.Color = colorResultGlay Then
                            paintPostureScore m
                        End If
                    Next

                    ' 強制入力後の信頼性セル色塗り
                    .Range(.Cells(ROW_RELIABILITY_TOP, MinLeftCell), _
                           .Cells(ROW_RELIABILITY_BOTTOM, MaxRightCell)).Interior.Color = colorForcedSection

                    ' ※拳上げセルの色塗り処理は一時的に除外中
                End If

                checkReliabilityRatio
            End If
        End With
    Else
        MsgBox "範囲はグラフ内から選択してください", vbOKOnly + vbCritical, "範囲選択エラー"
    End If

    checkReliabilityRatio
End Sub


'------------------------------------------------------------
' ポイント計算シートの姿勢点・信頼性を更新
'
' 引数:
'   sclm  - 選択範囲の左端列番号（実際の列位置）
'   fclm  - 選択範囲の右端列番号
'   bit   - 初期化フラグ（0:初期化 / 1:強制）※未使用？
'   score - 姿勢点（-1:初期化, 1〜10:強制, 99:除外）
'------------------------------------------------------------
Sub postureUpdate(sclm As Long, fclm As Long, bit As Long, score As Long)

    Dim s As Long, last As Long, i As Long
    Dim vle As Long

    ' データ列へのオフセット変換（データは2行目から）
    s = sclm - COLUMN_ZERO_NUM + 1
    last = fclm - COLUMN_ZERO_NUM + 1

    For i = s To last
        With ThisWorkbook.Sheets("ポイント計算シート")

            '-------------------------------
            ' 初期化処理
            '-------------------------------
            If score = -1 Then
                vle = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value

                ' 拳上げデータも元に戻す
                .Cells(i, COLUMN_POSTURE_SCORE_KOBUSHIAGE).Value = _
                    .Cells(i, COLUMN_POSTURE_SCORE_KOBUSHIAGE - 1).Value

            '-------------------------------
            ' 強制スコア（1～9）
            '-------------------------------
            ElseIf 1 <= score And score <= 9 Then
                vle = score

            '-------------------------------
            ' 除外（score = 99）またはその他
            '-------------------------------
            Else
                vle = score
            End If

            '-------------------------------
            ' 姿勢点の表示・集計用列にスコアを設定
            '-------------------------------
            Select Case vle
                Case 99 ' 除外
                    .Cells(i, COLUMN_POSTURE_GREEN).Value = 0
                    .Cells(i, COLUMN_POSTURE_YELLOW).Value = 0
                    .Cells(i, COLUMN_POSTURE_RED).Value = 0
                    .Cells(i, COLUMN_DATA_RESULT_FIX).Value = 0
                    .Cells(i, COLUMN_POSTURE_SCORE_KOBUSHIAGE).Value = 0 ' 拳上げ列も0に

                Case 1 To 2 ' 楽な姿勢
                    .Cells(i, COLUMN_POSTURE_GREEN).Value = vle
                    .Cells(i, COLUMN_POSTURE_YELLOW).Value = 0
                    .Cells(i, COLUMN_POSTURE_RED).Value = 0
                    .Cells(i, COLUMN_DATA_RESULT_FIX).Value = vle

                Case 3 To 5 ' やや辛い姿勢
                    .Cells(i, COLUMN_POSTURE_GREEN).Value = 0
                    .Cells(i, COLUMN_POSTURE_YELLOW).Value = vle
                    .Cells(i, COLUMN_POSTURE_RED).Value = 0
                    .Cells(i, COLUMN_DATA_RESULT_FIX).Value = vle

                Case Else ' 辛い姿勢（6以上）
                    .Cells(i, COLUMN_POSTURE_GREEN).Value = 0
                    .Cells(i, COLUMN_POSTURE_YELLOW).Value = 0
                    .Cells(i, COLUMN_POSTURE_RED).Value = vle
                    .Cells(i, COLUMN_DATA_RESULT_FIX).Value = vle
            End Select
        End With

        ' 信頼性の更新（個別フレーム）
        reliabilityUpdate i, score
    Next

End Sub


'------------------------------------------------------------
' ポイント計算シートの信頼性フラグを更新する処理
'
' 引数:
'   row  - 処理対象のデータ行番号（2行目以降）
'   vle  - 入力値（-1:初期化, 99:除外, それ以外:強制）
'------------------------------------------------------------
Sub reliabilityUpdate(row As Long, vle As Long)
    With ThisWorkbook.Sheets("ポイント計算シート")
        Select Case vle
            Case -1 ' 初期化処理
                .Cells(row, COLUMN_FORCED_SECTION).Value = 0
                .Cells(row, COLUMN_REMOVE_SECTION).Value = 0

            Case 99 ' 除外処理
                .Cells(row, COLUMN_REMOVE_SECTION).Value = 1
                .Cells(row, COLUMN_FORCED_SECTION).Value = 0

            Case Else ' 強制処理
                .Cells(row, COLUMN_REMOVE_SECTION).Value = 0
                .Cells(row, COLUMN_FORCED_SECTION).Value = 1
        End Select
    End With
End Sub


'------------------------------------------------------------
' ポイント計算シートの拳上げスコアを更新する処理
'
' 引数:
'   sclm  - 選択範囲の左端列番号（実際の列位置）
'   fclm  - 選択範囲の右端列番号
'   bit   - 処理種別（0:初期化/戻る, 1:強制）
'   score - スコア値（-1:戻る, 0:OFF, 1:ON, 99:除外）
'------------------------------------------------------------
Sub postureUpdate_Kobushiage(sclm As Long, fclm As Long, bit As Long, score As Long)

    Dim s As Long, last As Long, i As Long
    Dim fbit As Long, vle As Long
    Dim column_forced_num As Long

    ' 選択行が拳上げ行だった場合にのみ対象列をセット
    If Selection.Row = ROW_POSTURE_SCORE_KOBUSHIAGE Then
        column_forced_num = COLUMN_POSTURE_SCORE_KOBUSHIAGE
    End If

    ' データは2行目以降が対象
    s = sclm - COLUMN_ZERO_NUM + 2
    last = fclm - COLUMN_ZERO_NUM + 2

    For i = s To last
        With ThisWorkbook.Sheets("ポイント計算シート")

            ' 信頼性ビットを取得
            fbit = .Cells(i, COLUMN_FORCED_SECTION_TOTAL).Value

            '-------------------------------
            ' bit=0: 初期化時の値決定
            '-------------------------------
            If bit = 0 Then
                If fbit = 0 Then
                    vle = .Cells(i, COLUMN_POSTURE_SCORE_ALL).Value
                Else
                    vle = .Cells(i, COLUMN_BASE_SCORE).Value
                End If

                ' 除外フラグが立っている場合は常にベーススコアを使用
                If .Cells(i, COLUMN_REMOVE_SECTION).Value = 1 Then
                    vle = .Cells(i, COLUMN_BASE_SCORE).Value
                End If

            '-------------------------------
            ' bit=1: 強制入力時のスコア値
            '-------------------------------
            Else
                vle = score
            End If

            ' ベーススコアも更新
            baseScore i, bit

            ' 拳上げスコア列に値を代入（表示・字幕用）
            .Cells(i, COLUMN_POSTURE_SCORE_KOBUSHIAGE).Value = vle

            Debug.Print vle
        End With

        ' 信頼性更新処理（拳上げ用）
        reliabilityUpdate_Kobushiage i, bit, vle, column_forced_num
    Next
End Sub


'------------------------------------------------------------
' 拳上げスコアに対応する信頼性フラグを更新する処理
'
' 引数:
'   row                 - 処理対象のデータ行番号
'   bit                 - 処理種別（0:リセット, 1:強制）
'   vle                 - スコア値（-1:リセット, 0:OFF, 1:ON, 99:除外）
'   column_forced_num   - 処理対象の姿勢スコア列（拳上げ固定）
'------------------------------------------------------------
Sub reliabilityUpdate_Kobushiage(row As Long, bit As Long, vle As Long, column_forced_num As Long)
    With ThisWorkbook.Sheets("ポイント計算シート")

        Select Case True
            '------------------------------
            ' 除外処理（スコア99 + 強制）
            '------------------------------
            Case vle = 99 And bit = 1
                ' 除外ビットを立てる
                .Cells(row, COLUMN_REMOVE_SECTION).Value = bit
                ' 拳上げスコア初期化
                .Cells(row, COLUMN_POSTURE_SCORE_KOBUSHIAGE).Value = 0
                ' 拳上げ強制フラグ解除
                .Cells(row, COLUMN_FORCED_SECTION_KOBUSHIAGE).Value = 0

            '------------------------------
            ' リセット処理（bit = 0）
            '------------------------------
            Case bit = 0
                ' 全体フラグと除外フラグのクリア
                .Cells(row, COLUMN_FORCED_SECTION_TOTAL).Value = 0
                .Cells(row, COLUMN_REMOVE_SECTION).Value = 0
                ' 拳上げスコアを初期値に戻す
                .Cells(row, COLUMN_POSTURE_SCORE_KOBUSHIAGE).Value = .Cells(row, COLUMN_POSTURE_SCORE_KOBUSHIAGE - 1).Value
                ' 拳上げ強制フラグ解除
                .Cells(row, COLUMN_FORCED_SECTION_KOBUSHIAGE).Value = 0

            '------------------------------
            ' 強制入力（bit = 1）
            '------------------------------
            Case Else
                ' 除外フラグを解除
                .Cells(row, COLUMN_REMOVE_SECTION).Value = 0
                ' 拳上げ信頼性ビットを強制に設定
                .Cells(row, COLUMN_FORCED_SECTION_TOTAL).Value = 1
        End Select

    End With
End Sub


'------------------------------------------------------------
' 拳上げ用の元データ列（姿勢スコア）を更新する処理
'
' 引数:
'   row - データ対象の行番号（ポイント計算シート）
'   bit - 処理種別（0:戻るでリセット, 1:強制）
'------------------------------------------------------------
Sub baseScore(row As Long, bit As Long)
    With ThisWorkbook.Sheets("ポイント計算シート")
        If bit = 1 Then
            ' 強制入力時：元データが空なら、現スコアを記録
            If .Cells(row, COLUMN_BASE_SCORE).Value = "" Then
                .Cells(row, COLUMN_BASE_SCORE).Value = .Cells(row, COLUMN_POSTURE_SCORE_ALL).Value
            End If
        Else
            ' リセット時：保存していた元データをスコア列に復元
            .Cells(row, COLUMN_POSTURE_SCORE_ALL).Value = .Cells(row, COLUMN_BASE_SCORE).Value
        End If
    End With
End Sub

'『初期化』ボタンが押されたとき
Sub reset()
    Call forceResult(-1)
End Sub


'姿勢点『0』強制ボタンが押されたとき
Sub jogai()
    Call forceResult(99)
End Sub


'2023/12/11 育成G追記 拳上げ強制OFFボタンを押す
Sub force_kobushi_OFF()
    Call forceResult_Kobushiage(0)
End Sub


'2023/12/11 育成G追記 拳上げ強制ONボタンを押す
Sub force_kobushi_On()
    Call forceResult_Kobushiage(1)
End Sub


'姿勢点『1』強制ボタンが押されたとき
Sub force1()
    Call forceResult(1)
End Sub


'姿勢点『2』強制ボタンが押されたとき
Sub force2()
    Call forceResult(2)
End Sub


'姿勢点『3』強制ボタンが押されたとき
Sub force3()
    Call forceResult(3)
End Sub


'姿勢点『4』強制ボタンが押されたとき
Sub force4()
    Call forceResult(4)
End Sub


'姿勢点『5』強制ボタンが押されたとき
Sub force5()
    Call forceResult(5)
End Sub


'姿勢点『6』強制ボタンが押されたとき
Sub force6()
    Call forceResult(6)
End Sub


'姿勢点『7』強制ボタンが押されたとき
Sub force7()
    Call forceResult(7)
End Sub


'姿勢点『8』強制ボタンが押されたとき
Sub force8()
    Call forceResult(8)
End Sub


'姿勢点『9』強制ボタンが押されたとき
Sub force9()
    Call forceResult(9)
End Sub


'姿勢点『10』強制ボタンが押されたとき
Sub force10()
    Call forceResult(10)
End Sub


'------------------------------------------------------------
' 信頼性の割合を計算し、修正シートへ反映する処理
'
' 背景:
'   姿勢データにおける5種の区間（測定・推定・欠損・強制・除外）の割合を
'   計算し、各修正シートにパーセンテージで出力する。
'
' 備考:
'   RGB色は目視チェックや将来的な色変更に備えて定数化されている。
'------------------------------------------------------------
Sub checkReliabilityRatio()
    ' 変数定義
    Dim i                       As Long
    Dim fps                     As Double
    Dim maxRowNum               As Long
    Dim ColumnNum               As Long
    Dim maxArrayNum             As Long
    Dim reliabilityFlag         As Long
    Dim measurementSectionRatio As Double
    Dim predictSectionRatio     As Double
    Dim missingSectionRatio     As Double
    Dim coercionSectionRatio    As Double
    Dim exclusionSectionRatio   As Double
    Dim totalRatio              As Double

    Dim reliabilityColorDataArray()     As Long
    Dim reliabilityColorCounterArray(5) As Long

    ' 色定義（RGB）
    Dim colorMeasureSection    As String: colorMeasureSection = RGB(0, 176, 240)
    Dim colorPredictSection    As String: colorPredictSection = RGB(252, 246, 0)
    Dim colorMissingSection    As String: colorMissingSection = RGB(255, 124, 128)
    Dim colorForcedSection     As String: colorForcedSection = RGB(0, 51, 204)
    Dim colorRemoveSection     As String: colorRemoveSection = RGB(191, 191, 191)

    ' ポイント計算シートからFPSと最終行を取得
    With ThisWorkbook.Sheets("ポイント計算シート")
        fps = getFps()
        maxRowNum = getLastRow()
    End With

    ' 修正シート列数を初期化（全体列 - ラベル列）
    ColumnNum = 16206
    maxArrayNum = ColumnNum - 1
    ReDim reliabilityColorDataArray(maxArrayNum, 0)
    Erase reliabilityColorCounterArray

    ' 信頼性カウント処理
    For i = 2 To maxRowNum + 1
        With ThisWorkbook.Sheets("ポイント計算シート")
            Select Case True
                Case .Cells(i, COLUMN_REMOVE_SECTION).Value > 0
                    reliabilityColorCounterArray(5) = reliabilityColorCounterArray(5) + 1
                Case .Cells(i, COLUMN_FORCED_SECTION).Value > 0
                    reliabilityColorCounterArray(4) = reliabilityColorCounterArray(4) + 1
                Case .Cells(i, COLUMN_MISSING_SECTION).Value > 0
                    reliabilityColorCounterArray(3) = reliabilityColorCounterArray(3) + 1
                Case .Cells(i, COLUMN_PREDICT_SECTION).Value > 0
                    reliabilityColorCounterArray(2) = reliabilityColorCounterArray(2) + 1
                Case .Cells(i, COLUMN_MEASURE_SECTION).Value > 0
                    reliabilityColorCounterArray(1) = reliabilityColorCounterArray(1) + 1
            End Select
        End With
    Next

    ' 各信頼性の割合を計算
    predictSectionRatio     = reliabilityColorCounterArray(2) / maxRowNum * 100
    missingSectionRatio     = reliabilityColorCounterArray(3) / maxRowNum * 100
    exclusionSectionRatio   = reliabilityColorCounterArray(5) / maxRowNum * 100
    measurementSectionRatio = reliabilityColorCounterArray(1) / maxRowNum * 100
    coercionSectionRatio    = reliabilityColorCounterArray(4) / maxRowNum * 100

    ' 修正シートの一覧取得と結果反映
    Dim sName() As String
    Dim n As Long
    Dim actSheet As Worksheet
    Set actSheet = ActiveSheet
    sName = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢素点修正シート")

    For n = 1 To UBound(sName)
        With ThisWorkbook.Sheets(sName(n))
            .Cells(3, 4) = Round(measurementSectionRatio, 1) & "%" ' 測定
            .Cells(4, 4) = Round(coercionSectionRatio, 1) & "%"    ' 強制
            .Cells(5, 4) = Round(exclusionSectionRatio, 1) & "%"   ' 除外
            .Cells(6, 4) = Round(predictSectionRatio, 1) & "%"     ' 推定
            .Cells(7, 4) = Round(missingSectionRatio, 1) & "%"     ' 欠損
            .Cells(3, 5) = Round(measurementSectionRatio + coercionSectionRatio + exclusionSectionRatio, 1) & "%"
            .Cells(6, 5) = Round(predictSectionRatio + missingSectionRatio, 1) & "%"
        End With
    Next
End Sub


'------------------------------------------------------------
' 幅調整処理（拡大・縮小）
'
' 概要:
'   ボタン操作に応じて姿勢素点修正シートの列幅を拡大／縮小する。
'
' 引数:
'   expansionFlag - 拡大/縮小のフラグ
'     True  = 拡大
'     False = 縮小
'
' 備考:
'   初回呼び出し時に現在の幅サイズを基準として記録する。
'------------------------------------------------------------
Sub adjustWidth(expansionFlag As Boolean)
    Const EXPANSION_RATIO As Long = 100  ' 未使用定数だが保持
    Static initFin As Boolean            ' 初回呼び出しフラグ
    Static wSize As widthSize           ' 現在の幅サイズ

    ' 画面更新停止
    Call stopUpdate

    ' 初回のみ現在の列幅を取得してwSizeを初期化
    If Not initFin Then
        initFin = True
        Dim initSize As Long
        initSize = DataAjsSht.GetWidthPoints()

        Select Case initSize
            Case Is < widthSize.Medium
                wSize = Small
            Case Is < widthSize.Large
                wSize = Medium
            Case Is < widthSize.LL
                wSize = Large
            Case Else
                wSize = LL
        End Select
    End If

    ' フラグに応じて次のサイズへ移行
    wSize = sizeNext(wSize, expansionFlag)

    ' シート一覧取得と処理実行
    Dim sName() As String
    Dim n As Long
    Dim actSheet As Worksheet
    Set actSheet = ActiveSheet

    sName = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢素点修正シート")
    For n = 1 To UBound(sName)
        Call DataAjsSht.SetCellsHW(CInt(wSize), ThisWorkbook.Sheets(sName(n)))
    Next

    ' 元のシートへ戻す
    actSheet.Activate

    ' 画面更新再開
    Call restartUpdate
End Sub


'『幅拡大』ボタンが押されたとき
Sub expandWidth()
    '引数：expansionFlag As Long　幅の拡大or縮小を決める
    'False：縮小　True:拡大
    Call adjustWidth(True)
End Sub


'『幅縮小』ボタンが押されたとき
Sub reduceWidth()
    '引数：expansionFlag As Boolean　幅の拡大or縮小を決める
    'False：縮小　True:拡大
    Call adjustWidth(False)
End Sub


'1画面左へスクロール
Sub scrollToLeftPage()
        ActiveWindow.LargeScroll ToLeft:=1
End Sub


'1画面右へスクロール
Sub scrollToRightPage()
        If ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count <= _
        ActiveSheet.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column Then
            ActiveWindow.LargeScroll ToRight:=1
        End If
End Sub


'最も左へスクロール
Sub scrollToLeftEnd()
    Dim scrclm As Long
    If getClm(1) Then
        If getPageShapeState(ActiveSheet, "prevPage") Then
            Call prevPage_Click
        Else
            Call initCellPlace(ActiveSheet)
        End If
    Else
        Call initCellPlace(ActiveSheet)
    End If

End Sub


'------------------------------------------------------------
' 右端スクロール処理
'
' 概要:
'   姿勢素点修正シートの右端までスクロールする。
'   すでに右端なら次ページへ移動（nextPageボタン相当）する。
'
' 備考:
'   TIME_ROW：時刻行（列の終端判定に使用）
'   getClm, getPageShapeState, finCellPlace は補助関数
'------------------------------------------------------------
Sub scrollToRightEnd()
    Dim keepColumn As Long

    ' 右端に達しているかをチェック
    If getClm(ActiveSheet.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column) Then
        ' ページ移動が可能なら次のページへ
        If getPageShapeState(ActiveSheet, "nextPage") Then
            Call nextPage_Click
        End If
    Else
        ' 現在の最右列を取得（※必要ならkeepColumnに保持可能）
        keepColumn = ActiveSheet.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column

        ' スクロール位置を設定（見やすい位置まで29列戻す）
        ActiveWindow.Panes(2).ScrollColumn = keepColumn - 29

        ' 小スクロールでスクロール範囲を表示に収める
        ActiveWindow.SmallScroll ToLeft:=ActiveWindow.Panes(2).VisibleRange.Columns.Count

        ' 条件によって微調整
        If keepColumn = 16192 Then
            ' 特定列の場合は微調整（5列分右へ）
            ActiveWindow.SmallScroll ToRight:=5
        Else
            ' 通常は約3秒分（90列）スクロール
            ActiveWindow.SmallScroll ToRight:=90
        End If

        ' カーソル・セル位置の最終調整
        Call finCellPlace(ActiveSheet)
    End If
End Sub


'------------------------------------------------------------
' 同じ列に対して連続でスクロール処理が呼ばれたかを判定する
'
' 引数:
'   clm - 現在のカラム番号（列位置）
'
' 戻り値:
'   True  - 直前のカラムと同じ（連続呼び出し）
'   False - カラムが変わった（初回または位置変更）
'
' 備考:
'   Static変数 keepColumn によって、前回呼び出し時の列位置を記憶する。
'------------------------------------------------------------
Private Function getClm(clm As Long) As Boolean
    Static keepColumn As Long
    Dim isSameColumn As Boolean

    If keepColumn = clm Then
        isSameColumn = True
    Else
        keepColumn = clm
        isSameColumn = False
    End If

    getClm = isSameColumn
End Function


'表示倍率を画面にフィット
Sub fit()
    '見えている列範囲を取得
    Dim visibleColumn As String

    '見えている列範囲のうち左から7番目の列を取得（編集ボタンが置かれている1～6列を飛ばす）
    visibleColumn = Split(ActiveWindow.VisibleRange.Cells(7, 1).Address(True, False), "$")(0)
    '1～時刻の2行下までを選択
    Range(visibleColumn & "1:" & visibleColumn & BOTTOM_OF_TABLE + 2).Select
    '画面にフィット
    ActiveWindow.Zoom = True
    'A1セルを選択
    Range("A1").Select
    '画面を一番上までスクロール
    ActiveWindow.ScrollRow = 1

End Sub


'------------------------------------------------------------
' 再生ボタン処理
'
' 概要:
'   姿勢素点修正シート上で再生ボタンが押された際に、
'   時間カラムの自動選択処理を定期的に実行する。
'
' 引数:
'   なし（グローバル変数 ResTime を使用して再帰的に呼び出される）
'
' 備考:
'   - シートが「姿勢素点修正シート」でなければ再帰処理を停止する。
'   - 再生ボタンの非表示処理を実行する。
'------------------------------------------------------------
Sub RegularInterval1()
    Dim iend As Long, i As Long
    Dim dajsht() As String
    Dim currentColumn As Long
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 対象となる修正シート一覧を取得
    dajsht = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢素点修正シート")
    iend = UBound(dajsht)

    ' 全シートで再生ボタンを非表示にする
    For i = 1 To iend
        With Worksheets(dajsht(i))
            .Shapes("playBtn").Visible = False
        End With
    Next

    ' カラム位置確認
    currentColumn = ActiveCell.Column
    If currentColumn < TIME_COLUMN_LEFT Then
        ActiveSheet.Cells(BOTTOM_OF_TABLE, TIME_COLUMN_LEFT).Select
        ' 初期スタートに見せるため1秒待機
        Application.Wait Now() + TimeValue("00:00:01")
    End If

    ' 次回実行時刻（1秒後）を設定
    ResTime = Now + TimeValue("00:00:01")

    ' 自身を1秒後に再実行する設定
    Application.OnTime EarliestTime:=ResTime, Procedure:="RegularInterval1"

    ' シート名確認（姿勢素点修正シートのみ継続）
    If ActiveSheet.Name Like "姿勢素点修正シート*" Then
        Call nextTimeSelect
    Else
        Call Cancel1
    End If
End Sub


'------------------------------------------------------------
' 時刻選択処理
'
' 概要:
'   アクティブセルの次の時間セルへ移動し、
'   時刻が表示されていなければ次シートへ遷移する。
'
' 引数:
'   なし
'
' 備考:
'   - アクティブセルの行はそのままで、右方向へ時刻を進める。
'   - 時刻が存在しない場合は終了または次ページへ遷移。
'------------------------------------------------------------
Sub nextTimeSelect()
    ' アクティブセルの位置に応じて処理を行う

    ' 選択中の列の時刻行 (23行目) に移動
    Cells(TIME_ROW, Selection.Column).Select

    ' 一つ右のセル（次の時刻）に移動
    ActiveCell.Offset(0, 1).Select

    ' 表示を1秒分右へスクロール（30fps × 1秒）
    ActiveWindow.SmallScroll ToRight:=TIME_WIDTH

    ' 選択されたセルが空白なら処理継続 or 終了判定
    If IsEmpty(ActiveCell.Value) Then
        If getPageShapeState(ActiveSheet, "nextPage") Then
            ' 次のページが存在すれば遷移
            Call nextPage_Click
        Else
            ' 時刻が存在せず、次ページもなければ処理終了
            Call Cancel1
        End If
    End If
End Sub


'------------------------------------------------------------
' 停止ボタン処理
'
' 概要:
'   姿勢素点修正シート上で「停止ボタン」が押されたときに、
'   再生の自動実行を止め、各シートの再生ボタンを再表示する。
'
' 引数:
'   なし
'
' 備考:
'   - 再生制御用の ResTime を用いて Application.OnTime を解除。
'   - 姿勢素点修正シートすべての再生ボタンを再表示。
'------------------------------------------------------------
Sub Cancel1()
    Dim iend As Long, i As Long
    Dim dajsht() As String

    ' 姿勢素点修正シート一覧を取得
    dajsht = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢素点修正シート")
    iend = UBound(dajsht)

    ' 各シートの再生ボタンを再表示
    For i = 1 To iend
        With Worksheets(dajsht(i))
            .Shapes("playBtn").Visible = True
        End With
    Next

    ' Application.OnTime を使ってタイマーを停止
    On Error Resume Next
    Application.OnTime EarliestTime:=ResTime, _
                       Procedure:="RegularInterval1", _
                       Schedule:=False
    On Error GoTo 0
End Sub


'メッセージボックスの表示
'戻り値：メッセージボックス
Function YesorNo() As VbMsgBoxResult
    YesorNo = MsgBox("この場所に" & ActiveWorkbook.Name & _
                        "という名前のファイルが既にあります。置き換えますか？", _
                        vbInformation + vbYesNoCancel + vbDefaultButton2)
End Function


'ブック全体の保存
Sub Savebook()
    Dim dotPoint     As String
    Dim workbookName As String
    Dim fps          As Double

    'フレームレートを取得
    fps = getFps()
    If YesorNo() = vbYes Then


        Call stopUpdate
        Call Module2.fixSheetZensya(fps)

        dotPoint = InStrRev(ActiveWorkbook.Name, ".")
        workbookName = Left(ActiveWorkbook.Name, dotPoint - 1)
        Call Module2.outputCaption(workbookName)

        ActiveWorkbook.Save

        Call restartUpdate
    End If
End Sub


'sheetの左から何番に属するか判定する
'引数1：シート
'戻り値：シートが何番目に属しているか
Function calcSheetNamePlace(ws As Worksheet)
    Dim shNameArray() As String
    Dim i As Long
    Dim iend As Long
    Dim ret As Long: ret = 0

    shNameArray() = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢素点修正シート")
    iend = UBound(shNameArray)
    For i = 1 To iend
        If ws.Name = shNameArray(i) Then
            ret = i - 1
        End If
    Next
    calcSheetNamePlace = ret
End Function


'------------------------------------------------------------
' 対象名を含むワークシート名を配列として返す関数
'
' 概要:
'   指定された名前を含むワークシートを左から順に検索し、
'   一致したシート名を配列として返す。
'
' 引数:
'   wb  : Workbook オブジェクト
'   str : シート名に含まれる文字列（例: "姿勢素点修正シート"）
'
' 戻り値:
'   一致したワークシート名の配列（String型）
'------------------------------------------------------------
Function call_GetSheetNameToArrayspecific(wb As Workbook, str As String) As String()
    Dim tmp() As String
    Dim ws As Worksheet
    Dim i As Long: i = 0
    Dim sheetCount As Long
    sheetCount = wb.Worksheets.Count

    ' 全ワークシートを走査して対象名を含むものを追加
    For Each ws In wb.Worksheets
        If InStr(ws.Name, str) > 0 Then
            i = i + 1
            ReDim Preserve tmp(1 To i)
            tmp(i) = ws.Name
        End If
    Next

    call_GetSheetNameToArrayspecific = tmp
End Function


'簡易的なシート切替処理も兼ねた矢印の図形
'引数1：姿勢素点修正シート
'引数2：前ページに移動するアイコンを非表示にするかどうか（true or false）
'引数3：次ページに移動するアイコンを非表示にするかどうか（true or false）
Private Sub addPageShape(ws As Worksheet, pPageState As Boolean, nPageState As Boolean)
    Const pPage As String = "prevPage"
    Const nPage As String = "nextPage"
    Const pCover As String = "prevPage_Disabled"
    Const nCover As String = "nextPage_Disabled"

    Call initCellPlace(ws)

    ' 表示可能なページ送りボタン
    ws.Shapes(pPage).Visible = pPageState
    ws.Shapes(nPage).Visible = nPageState

    ' グレーのカバー画像（押せない状態の見た目）を逆に連動
    ws.Shapes(pCover).Visible = Not pPageState
    ws.Shapes(nCover).Visible = Not nPageState
End Sub


'図形がVisibleかどうか判定する
'引数1：ワークシート
'引数2：図形の名前
'戻り値:Visibleかどうか（0 or 1）
Private Function getPageShapeState(ws As Worksheet, shapeName As String)
    getPageShapeState = ws.Shapes(shapeName).Visible
End Function


'------------------------------------------------------------
' 姿勢素点修正シートを右隣に複製する処理
'
' 引数:
'   dupCount : 複製対象のインデックス（0の場合はオリジナル名）
'
' 備考:
'   - 既にシートが存在する場合は再帰的にインデックスを増加
'   - s_Master_2nd シートを複製し、名前を設定
------------------------------------------------------------
Sub createSheet(dupCount As Long)
    Dim ws As Worksheet
    Dim dupFlag As Boolean
    Dim wsName As String

    wsName = "姿勢素点修正シート"
    dupFlag = False

    ' シートの重複確認
    For Each ws In Worksheets
        If dupCount = 0 Then
            If ws.Name = wsName Then dupFlag = True
        Else
            If ws.Name = wsName & " (" & Replace(CStr(dupCount), " ", "") & ")" Then dupFlag = True
        End If
    Next ws

    ' 重複していたら再帰的に次の番号で作成を試みる
    If dupFlag Then
        Call createSheet(dupCount + 1)
    Else
        On Error GoTo ErrLabel
        ' シートの複製とリネーム処理
        s_Master_2nd.Visible = True
        Select Case dupCount
            Case 0
                s_Master_2nd.Copy After:=Worksheets(s_Graph_2nd.Name)
                ActiveSheet.Name = wsName
            Case 1
                s_Master_2nd.Copy After:=Worksheets(wsName)
                ActiveSheet.Name = wsName & " (" & Replace(CStr(dupCount), " ", "") & ")"
            Case Else
                s_Master_2nd.Copy After:=Worksheets(wsName & " (" & Replace(CStr(dupCount - 1), " ", "") & ")")
                ActiveSheet.Name = wsName & " (" & Replace(CStr(dupCount), " ", "") & ")"
        End Select

        s_Master_2nd.Visible = xlSheetVeryHidden
    End If

    Exit Sub

ErrLabel:
    MsgBox "存在しないシートです", vbCritical
End Sub


'ワークシートを消去
Sub DeleteSheet(dupCount As Long)

    Dim ws As Worksheet
    Dim dupflag As Boolean
    Dim wsName As String
    Dim time As Long
    time = 1000

    wsName = "姿勢素点修正シート"

    For Each ws In Worksheets
    'シート検索

        If dupCount = 0 Then
            If ws.Name = wsName Then
                dupflag = True
                Application.DisplayAlerts = False
                Worksheets(ws.Name).Delete
                Application.Wait [Now()] + time / 86400000
                Application.DisplayAlerts = True
            End If
        Else
            If ws.Name = wsName + " (" + Replace(str(dupCount), " ", "") + ")" Then
                Application.DisplayAlerts = False
                Worksheets(ws.Name).Delete
                Application.DisplayAlerts = True
                Application.Wait [Now()] + time / 86400000
                dupflag = True
            End If
        End If

    Next ws

    If dupflag = True Then
    'シートが存在するため再起する
        DeleteSheet (dupCount + 1)
    End If

End Sub


'ひとつ前のシートをアクティブにし、データの最後尾まで行く
Sub prevPage_Click()
    ThisWorkbook.ActiveSheet.Previous.Activate
    Call finCellPlace(ThisWorkbook.ActiveSheet)
End Sub


'ひとつ次のシートをアクティブにし、データの最初に行く
Sub nextPage_Click()
    ThisWorkbook.ActiveSheet.Next.Activate
    Call initCellPlace(ThisWorkbook.ActiveSheet)
End Sub


'セルの初期位置
Private Sub initCellPlace(ws As Worksheet)
    ws.Cells(TIME_ROW, TIME_COLUMN_LEFT).Select
End Sub


'セルの最終位置
Private Sub finCellPlace(ws As Worksheet)
    ws.Cells(TIME_ROW, ws.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column).Select
End Sub


'------------------------------------------------------------
' サイズ切り替えロジック
'
' 引数:
'   wSize       : 現在のサイズ（Small, Medium, Large, LL）
'   nextChange  : True なら次のサイズへ拡大、False なら縮小
'
' 戻り値:
'   次に設定すべきサイズ（Small = 1, Medium = 2, Large = 4, LL = 6）
'------------------------------------------------------------
Private Function sizeNext(wSize As widthSize, nextChange As Boolean) As widthSize
    Dim tmpSize As widthSize

    Select Case wSize
        Case widthSize.Small
            If nextChange Then
                tmpSize = widthSize.Medium
                Call changeBtnState(REDUCEBTN_NAME, True)
            Else
                tmpSize = widthSize.Small
                Call changeBtnState(EXPANDBTN_NAME, True)
                ' ベースファイルの保存が悪かった時用の保険処理
                Call changeBtnState(REDUCEBTN_NAME, False)
            End If

        Case widthSize.Medium
            If nextChange Then
                tmpSize = widthSize.Large
            Else
                tmpSize = widthSize.Small
                Call changeBtnState(EXPANDBTN_NAME, True)
                Call changeBtnState(REDUCEBTN_NAME, False)
            End If

        Case widthSize.Large
            If nextChange Then
                tmpSize = widthSize.LL
                Call changeBtnState(EXPANDBTN_NAME, False)
                Call changeBtnState(REDUCEBTN_NAME, True)
            Else
                tmpSize = widthSize.Medium
            End If

        Case widthSize.LL
            If Not nextChange Then
                tmpSize = widthSize.Large
                Call changeBtnState(EXPANDBTN_NAME, True)
            Else
                tmpSize = widthSize.LL
                Call changeBtnState(REDUCEBTN_NAME, True)
                ' ベースファイルの保存が悪かった時用の保険処理
                Call changeBtnState(EXPANDBTN_NAME, False)
            End If
    End Select

    sizeNext = tmpSize
End Function


'幅調整用のボタンに使う予定。実際名前さえ決めることができればなんとでもなる。

'引数1：ボタンの名前（EXPANDBTN_NAME or REDUCEBTN_NAME）
'引数2：ボタンを押せるかどうか
Private Sub changeBtnState(btnName As String, btnstate As Boolean)
    Dim iend, i As Long
    Dim dajsht() As String

    dajsht() = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢素点修正シート")
    iend = UBound(dajsht)
    For i = 1 To iend
        With Worksheets(dajsht(i))
            .Shapes(btnName).Visible = btnstate
        End With
    Next
End Sub

'------------------------------------------------------------
' シートを初期状態にリセットする処理
' ・拡大・縮小ボタンを表示
' ・ページ切替ボタンを非表示
' ・背景色と罫線をクリア
' ・一部範囲の内容をクリア
'------------------------------------------------------------
Sub resetSheet()
    Const pPage As String = "prevPage"
    Const nPage As String = "nextPage"
    Dim iend As Long
    Dim i As Long
    Dim dajsht() As String

    ' 姿勢素点修正シートの一覧を取得
    dajsht() = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢素点修正シート")
    iend = UBound(dajsht)

    ' 各シートに対して初期化処理を実行
    For i = 1 To iend
        With Worksheets(dajsht(i))
            ' ボタン表示制御
            .Shapes(EXPANDBTN_NAME).Visible = True
            .Shapes(REDUCEBTN_NAME).Visible = True
            .Shapes(pPage).Visible = False
            .Shapes(nPage).Visible = False

            ' 背景色クリア
            .Range("G2:G25").Select
            .Range(Selection, Selection.End(xlToRight)).Select
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

            ' 罫線クリア
            .Range("FN2:FN25").Select
            .Range(Selection, Selection.End(xlToRight)).Select
            With Selection.Borders
                .LineStyle = xlNone
            End With

            ' データクリア
            .Range("G24:XFD25").Select
            Selection.ClearContents
        End With
    Next i
End Sub


'------------------------------------------------------------
' 非表示の「名前の定義」を再表示するマクロ
' （例：シートコピー後のエラー対策）
' 実行後、対象の名前をユーザーに通知
'------------------------------------------------------------
Public Sub ShowInvisibleNames()
    Dim oName As Name

    ' ワークブック内すべての名前をチェック
    For Each oName In Names
        If oName.Visible = False Then
            oName.Visible = True
        End If
    Next

    MsgBox "非表示の名前の定義を表示しました。", vbOKOnly
End Sub


' 選択範囲をデータ有効域と交差させる
' 戻り値 : True → 交差あり（ leftCol/rightCol が返る ）
'          False → 交差なし（メッセージは呼び出し側で）
Public Function CropSelectionToDataArea(ByRef leftCol As Long, ByRef rightCol As Long) As Boolean
    Const PAGE_FRAME_MAX As Long = LIMIT_COLUMN '16200
    Const VALID_ROW_TOP As Long = 14
    Const VALID_ROW_BOTTOM As Long = 23

    Dim shtPage   As Long
    Dim baseClm   As Long
    Dim selR      As Long              '選択列の終わり
    Dim frmR      As Long              '選択フレーム
    Dim pageFrmR  As Long              'ページの有効フレーム
    Dim totalFrm  As Long
    Dim rowTop    As Long
    Dim rowBottom As Long

    ' ボタン列を選んだら無視
    If Selection.Column > Columns.Count Then
        Exit Function
    End If

    ' 選択列範囲の終わりを算出
    selR = Selection.Column + Selection.Columns.Count - 1

    ' 行範囲を取得
    rowTop = Selection.row
    rowBottom = Selection.row + Selection.Rows.Count - 1

    'rowTop: 選択範囲の最初の行番号
    'rowBottom: 選択範囲の最後の行番号
    'VALID_ROW_TOP: 有効範囲の最上部（14）
    'VALID_ROW_BOTTOM: 有効範囲の最下部（23）

    ' ↓ 選択範囲が完全に範囲外（上にも下にも交差しない）場合にだけ除外
    'If rowBottom < VALID_ROW_TOP Or rowTop > VALID_ROW_BOTTOM Then
        ' 選択範囲全体が有効行の上下に完全に外れているなら終了
        'CropSelectionToDataArea = False
        'Exit Function
    'End If

    ' ↓「選択範囲の上端が14より前 もしくは 下端が23より後ろなら」選択範囲は範囲外
    ' 行と交差しているかをチェック
    If rowTop < VALID_ROW_TOP Or rowBottom > VALID_ROW_BOTTOM Then
        CropSelectionToDataArea = False
        Exit Function
    End If

    ' ページ情報の取得
    shtPage = calcSheetNamePlace(ActiveSheet)
    baseClm = LIMIT_COLUMN * shtPage

    totalFrm = getLastRow() - 1

    ' フレーム右端
    frmR = selR - COLUMN_ZERO_NUM + baseClm
    pageFrmR = WorksheetFunction.min(baseClm + PAGE_FRAME_MAX, totalFrm)
    frmR = WorksheetFunction.min(frmR, pageFrmR)

    ' 有効列範囲内に絞る
    leftCol = WorksheetFunction.Max(Selection.Column, COLUMN_ZERO_NUM + 1)
    rightCol = frmR - baseClm + COLUMN_ZERO_NUM

    If leftCol > rightCol Then
        CropSelectionToDataArea = False   ' 列で交差していない
    Else
        CropSelectionToDataArea = True    ' 行・列とも交差
    End If
End Function


'fpsの値を取得する
'戻り値：fpsの値
Function getFps() As Double
    getFps = ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 199).Value
End Function


'最終行を取得する
'戻り値：最終行
Function getLastRow() As Long
    getLastRow = ThisWorkbook.Sheets("ポイント計算シート").Cells(1, 3).End(xlDown).row
End Function