Option Explicit

'======================================================================================
'条件設定シートの各データの行番号、列番号を定義 (拳上概要の定数もここで定義）
'======================================================================================
Const TEKUBI_SPEED_UPLIM_PREDICT                As Double = 10      '（km/h）手首z位置の変化量上限　遮蔽検知に使う
Const MEAGERE_TIME_MACROUPDATEDATA              As Boolean = True   'TrueのときMacroUpdateDataの処理時間を測定する

'makeGraph、outputCaption、fixGraphDataAndSheetモジュールの中に条件設定シートのセル内から値を読み出す部分あり

'======================================================================================
'ポイント計算シート上の各データの行番号、列番号を定義
'======================================================================================
Const COLUMN_POSE_NAME                              As Long = 1
Const COLUMN_POSE_KEEP_TIME                         As Long = 2
Const COLUMN_HIZA_R_ANGLE                           As Long = 6
Const COLUMN_HIZA_L_ANGLE                           As Long = 7
Const COLUMN_KOSHI_ANGLE                            As Long = 8
Const COLUMN_SHOOTING_DIRECTION                     As Long = 9

Const COLUMN_POS_KOSHI_Z                            As Long = 13

Const COLUMN_POS_AHIKUBI_R_Z                        As Long = 25
Const COLUMN_POS_AHIKUBI_L_Z                        As Long = 37

Const COLUMN_POS_KATA_R_Z                           As Long = 57
Const COLUMN_POS_KATA_L_Z                           As Long = 69

Const COLUMN_POS_HIJI_R_Z                           As Long = 61
Const COLUMN_POS_HIJI_L_Z                           As Long = 73

Const COLUMN_POS_TEKUBI_R_Z                         As Long = 65
Const COLUMN_POS_TEKUBI_L_Z                         As Long = 77

Const COLUMN_ROUGH_TIME                             As Long = 201
Const COLUMN_CAPTION_WORK_NAME                      As Long = 202
Const COLUMN_DATA_RESULT_ORIGIN                     As Long = 203
Const COLUMN_DATA_MEASURE_SECTION                   As Long = 204
Const COLUMN_DATA_PREDICT_SECTION                   As Long = 205
Const COLUMN_DATA_REMOVE_SECTION                    As Long = 206
Const COLUMN_DATA_FORCED_SECTION                    As Long = 207
Const COLUMN_DATA_RESULT_FIX                        As Long = 208
Const COLUMN_DATA_RESULT_GREEN                      As Long = 209
Const COLUMN_DATA_RESULT_YELLOW                     As Long = 210
Const COLUMN_DATA_RESULT_RED                        As Long = 211

Const COLUMN_DATA_MISSING_SECTION                   As Long = 219

Const COLUMN_DATA_KOSHIMAGE_MEASURE_SECTION         As Long = 225
Const COLUMN_DATA_KOSHIMAGE_PREDICT_SECTION         As Long = 226
Const COLUMN_DATA_KOSHIMAGE_MISSING_SECTION         As Long = 227
Const COLUMN_KOSHIMAGE_FORCED_SECTION               As Long = 228
Const COLUMN_KOSHIMAGE_RESULT                       As Long = 247
Const COLUMN_DATA_HIZAMAGE_MEASURE_SECTION          As Long = 230
Const COLUMN_DATA_HIZAMAGE_PREDICT_SECTION          As Long = 231
Const COLUMN_DATA_HIZAMAGE_MISSING_SECTION          As Long = 232
Const COLUMN_HIZAMAGE_FORCED_SECTION                As Long = 233
Const COLUMN_HIZAMAGE_RESULT                        As Long = 249

Const COLUMN_TEKUBI_RZ_SPEED                        As Long = 237    '右手首Ｚ位置の差
Const COLUMN_TEKUBI_LZ_SPEED                        As Long = 238    '左手首Ｚ位置の差
Const COLUMN_TEKUBI_Z_SPEED_OVER                    As Long = 239    '手首Ｚ位置の差 しきい値超えフラグ
Const COLUMN_MEAGERE_TIME_MACROUPDATEDATA           As Long = 242    'MacroUpdateDataの処理時間を測定結果を格納する

Const COLUMN_DATA_RESULT_GH_KOSHIMAGE               As Long = 247
Const COLUMN_DATA_RESULT_GH_HIZAMAGE                As Long = 249
Const COLUMN_DATA_RESULT_GH_SONKYO                  As Long = 251

Const COLUMN_GH_HIZA_L                              As Long = 252
Const COLUMN_GH_HIZA_R                              As Long = 253

Const COLUMN_MAX_NUMBER                             As Long = 256   '現在使用されている列番号の最大値


'======================================================================================
'姿勢重量点調査票シートの各データの行番号、列番号を定義
'======================================================================================
Const SHIJUTEN_SHEET_ROW_KOUTEI_NAME                As Long = 3
Const SHIJUTEN_SHEET_ROW_POSESTART_INDEX            As Long = 9
Const SHIJUTEN_SHEET_ROW_EXPAND_NUMBER_CHECK        As Long = 29

Const SHIJUTEN_SHEET_EXPAND_NUM_CHECK_WORD          As String = "その他の時間（定時稼働時間7.5H-Σ延べ時間）"

Const SHIJUTEN_SHEET_COLUMN_WORK_NUMBER             As Long = 2
Const SHIJUTEN_SHEET_COLUMN_WORK_NAME               As Long = 3
Const SHIJUTEN_SHEET_COLUMN_KOUTEI_NAME             As Long = 4
Const SHIJUTEN_SHEET_COLUMN_WORK_TIME               As Long = 9
Const SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX        As Long = 10

Const SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME          As Long = 36
Const SHIJUTEN_SHEET_COLUMN_WORKEND_TIME            As Long = 38

Const SHIJUTEN_SHEET_COLUMN_DATA_MISSING_SECTION    As Long = 46
Const SHIJUTEN_SHEET_COLUMN_DATA_PREDICT_SECTION    As Long = 47

Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME          As Long = 51 '腰曲げ時間
Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME           As Long = 53 '膝曲げ時間

Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_MISSING_TIME  As Long = 57 '腰曲げ欠損区間
Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_PREDICT_TIME  As Long = 58 '腰曲げ推定区間

Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_MISSING_TIME   As Long = 60 '膝曲げ欠損区間
Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_PREDICT_TIME   As Long = 61 '膝曲げ推定区間


'======================================================================================
'工程評価シートの各データの行番号、列番号を定義
'======================================================================================

Const GH_HYOUKA_SHEET_ROW_POSESTART                    As Long = 9
Const GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK          As Long = 115

Const GH_HYOUKA_SHEET_EXPAND_NUM_CHECK_WORD            As String = "合計"

Const GH_HYOUKA_SHEET_COLUMN_WORK_NUMBER               As Long = 2
Const GH_HYOUKA_SHEET_COLUMN_WORK_NAME                 As Long = 3
Const GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME            As Long = 36
Const GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME              As Long = 38
Const GH_HYOUKA_SHEET_COLUMN_WORK_TIME                 As Long = 16
Const GH_HYOUKA_SHEET_COLUMN_KOSHIMAGE_TIME            As Long = 18
Const GH_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME             As Long = 19

'======================================================================================
'外販用　姿勢判定のしきい値を定義
'======================================================================================

Const GH_ANGLE_KOSHIMAGE_MIN As Double = 30
Const GH_ANGLE_KOSHIMAGE_MAX As Double = 180
Const GH_ANGLE_HIZAMAGE_MIN  As Double = 60
Const GH_ANGLE_HIZAMAGE_MAX As Double = 180

'======================================================================================
'DataAdjustingSheet用
'======================================================================================

Const LIMIT_COLUMN           As Long = 16200

'======================================================================================
'字幕情報の定義
'======================================================================================
Const CAPTION_TRACK2_FILE_NAME_SOEJI           As String = "2" '字幕トラック2用のファイル名末尾につける添字
'各種字幕のフォントサイズ係数
'分母の値のため、値が小さいほど文字は大きい
'動画が縦の時
Const TRACK1_TATE_UPPER_COEF                   As Long = 22 'トラック1用：上段
Const TRACK1_TATE_LOWER_COEF                   As Long = 11 'トラック1用：下段
Const TRACK2_TATE_1ST_COEF                     As Long = 22 'トラック2用：1段目
Const TRACK2_TATE_2ND_COEF                     As Long = 22 'トラック2用：2段目
Const TRACK2_TATE_3RD_COEF                     As Long = 13 'トラック2用：3段目

'動画が横の時
Const TRACK1_YOKO_UPPER_COEF                   As Long = 30 'トラック1用：上段
Const TRACK1_YOKO_LOWER_COEF                   As Long = 15 'トラック1用：下段
Const TRACK2_YOKO_1ST_COEF                     As Long = 30 'トラック2用：1段目
Const TRACK2_YOKO_2ND_COEF                     As Long = 30 'トラック2用：2段目
Const TRACK2_YOKO_3RD_COEF                     As Long = 18 'トラック2用：3段目

'各種字幕の色
Const COLOR_DATA_REMOVE_SECTION                As String = "#bfbfbf" 'グレー
Const COLOR_DATA_FORCED_SECTION                As String = "#0033cc" '青色
Const COLOR_DATA_MISSING_SECTION               As String = "#ff7c80" '朱色
Const COLOR_DATA_PREDICT_SECTION               As String = "#fcf600" '黄色
Const COLOR_DATA_MEASURE_SECTION               As String = "#00b0f0" '水色
Const COLOR_DATA_RESULT_GREEN                  As String = "#00b050" '緑色
Const COLOR_DATA_RESULT_YELLOW                 As String = "#ffc000" '黄色
Const COLOR_DATA_RESULT_RED                    As String = "#c00000" '赤色
Const COLOR_DATA_RESULT_GLAY                   As String = "#bfbfbf" 'グレー

'帯グラフのデータ（信頼度）を示す字幕文字列（字幕トラック1用 上段右側に表示）
Const CAPTION_DATA_MEASURE_SECTION             As String = "【データ測定区間】"
Const CAPTION_DATA_PREDICT_SECTION             As String = "【データ推定区間】"
Const CAPTION_DATA_REMOVE_SECTION              As String = "【データ除外区間】"
Const CAPTION_DATA_FORCED_SECTION              As String = "【データ強制区間】"
Const CAPTION_DATA_MISSING_SECTION             As String = "【データ欠損区間】"

'帯グラフのデータ（信頼度）を示す字幕文字列（字幕トラック2用 2段目に表示）
Const CAPTION_DATA_TRACK2_MEASURE_SECTION      As String = "【データ測定区間】"
Const CAPTION_DATA_TRACK2_PREDICT_SECTION      As String = "【データ推定区間】"
Const CAPTION_DATA_TRACK2_REMOVE_SECTION       As String = "【データ除外区間】"
Const CAPTION_DATA_TRACK2_FORCED_SECTION       As String = "【データ強制区間】"
Const CAPTION_DATA_TRACK2_MISSING_SECTION      As String = "【データ欠損区間】"

'外販用の字幕文字列（字幕トラック2用 3段目に表示）
Const CAPTION_A_RESULT_NAME1  As String = "　　　　拳上"
Const CAPTION_B_RESULT_NAME1  As String = "  　　腰曲げ　 　"
Const CAPTION_C_RESULT_NAME1  As String = "膝曲げ"

'外販用の条件字幕文字列（字幕トラック2用 4段目に表示）
Const CAPTION_A_RESULT_NAME2  As String = "手首が肩より上"
Const CAPTION_B_RESULT_NAME2  As String = "30°以上"
Const CAPTION_C_RESULT_NAME2  As String = "60°以上"

'キャプションノイズ除去の閾値
Const CAPTION_REMOVE_NOISE_SECOND              As Double = 0.1 'キャプションノイズを除去する長さ(秒) （～未満なら除去）

'姿勢素点の値によって、緑／黄／赤を分ける際の境界条件
Const DATA_SEPARATION_GREEN_BOTTOM             As Long = 1
Const DATA_SEPARATION_GREEN_TOP                As Long = 2
Const DATA_SEPARATION_YELLOW_BOTTOM            As Long = 3
Const DATA_SEPARATION_YELLOW_TOP               As Long = 5
Const DATA_SEPARATION_RED_BOTTOM               As Long = 6
Const DATA_SEPARATION_RED_TOP                  As Long = 10


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
' 腰曲げ・膝曲げの角度判定を行い、ポイント計算シートに結果を反映する処理
'
' 引数:
'   なし
'
' 備考:
'   - 判定しきい値は定数 GH_ANGLE_〜 から取得
'   - 膝角度は外販用に (180 - 測定値) で変換して記録
'   - キャプション時刻は hh:mm:ss,ミリ秒 形式で生成
'   - シートの表示更新を一時停止し、処理後に再開する
'------------------------------------------------------------
Sub makeGraphJisya()

    Call stopUpdate

    ' 条件設定しきい値
    Dim AngleKoshiMin   As Double: AngleKoshiMin = GH_ANGLE_KOSHIMAGE_MIN
    Dim AngleKoshiMax   As Double: AngleKoshiMax = GH_ANGLE_KOSHIMAGE_MAX
    Dim AngleHizaMin    As Double: AngleHizaMin  = GH_ANGLE_HIZAMAGE_MIN
    Dim AngleHizaMax    As Double: AngleHizaMax  = GH_ANGLE_HIZAMAGE_MAX

    ' 角度・判定・時間用変数
    Dim ValAngleKoshi   As Double
    Dim ValAngleHizaL   As Double
    Dim ValAngleHizaR   As Double
    Dim mSeconds        As String
    Dim totalSecond     As Long
    Dim hour            As Long
    Dim min             As Long
    Dim sec             As Long
    Dim t               As Date

    ' 行数・配列
    Dim max_row_num As Long
    Dim max_array_num As Long
    Dim i As Long
    Dim PointCalcSheetArray As Variant
    Dim HizaAngleLArray() As Double
    Dim HizaAngleRArray() As Double

    With ThisWorkbook.Sheets("ポイント計算シート")

        ' 行数と配列初期化
        max_row_num = getLastRow()
        max_array_num = max_row_num - 2 ' 2行目開始・0ベース
        PointCalcSheetArray = .Range(.Cells(1, 1), .Cells(max_row_num, COLUMN_MAX_NUMBER))

        ReDim HizaAngleLArray(max_array_num, 0)
        ReDim HizaAngleRArray(max_array_num, 0)

        ' 膝角度（外販用）は180-実値で取得
        For i = 0 To max_array_num
            HizaAngleLArray(i, 0) = 180 - .Cells(i + 2, COLUMN_HIZA_L_ANGLE).Value
            HizaAngleRArray(i, 0) = 180 - .Cells(i + 2, COLUMN_HIZA_R_ANGLE).Value
        Next

        ' 姿勢判定ループ
        For i = 2 To max_row_num

            ' キャプション時刻生成（hh:mm:ss,ミリ秒）
            totalSecond = Int(PointCalcSheetArray(i, 2))
            mSeconds = Right(Format(Round(PointCalcSheetArray(i, 2) - totalSecond, 3), "0.000"), 3)
            hour = totalSecond \ 3600
            min = (totalSecond Mod 3600) \ 60
            sec = totalSecond Mod 60
            t = TimeSerial(hour, min, sec)
            PointCalcSheetArray(i, COLUMN_ROUGH_TIME) = Format(t, "hh:mm:ss") & "," & mSeconds

            ' 関節角度読み出し
            ValAngleKoshi = CDbl(PointCalcSheetArray(i, COLUMN_KOSHI_ANGLE))
            ValAngleHizaL = CDbl(PointCalcSheetArray(i, COLUMN_HIZA_L_ANGLE))
            ValAngleHizaR = CDbl(PointCalcSheetArray(i, COLUMN_HIZA_R_ANGLE))

            ' 腰曲げ判定
            If ValAngleKoshi > AngleKoshiMin And ValAngleKoshi <= AngleKoshiMax Then
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_KOSHIMAGE) = 1
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_KOSHIMAGE - 1) = 1
            Else
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_KOSHIMAGE) = 0
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_KOSHIMAGE - 1) = 0
            End If

            ' 膝曲げ判定（180-実測値で比較）
            If (180 - ValAngleHizaL > AngleHizaMin And 180 - ValAngleHizaL <= AngleHizaMax) Or _
               (180 - ValAngleHizaR > AngleHizaMin And 180 - ValAngleHizaR <= AngleHizaMax) Then
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_HIZAMAGE) = 1
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_HIZAMAGE - 1) = 1
            Else
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_HIZAMAGE) = 0
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_HIZAMAGE - 1) = 0
            End If

        Next

        ' 判定・角度の書き戻し
        .Range(.Cells(1, 1), .Cells(max_row_num, COLUMN_MAX_NUMBER)) = PointCalcSheetArray
        .Range(.Cells(2, COLUMN_GH_HIZA_L), .Cells(max_row_num, COLUMN_GH_HIZA_L)).Value = HizaAngleLArray
        .Range(.Cells(2, COLUMN_GH_HIZA_R), .Cells(max_row_num, COLUMN_GH_HIZA_R)).Value = HizaAngleRArray

    End With

    Call restartUpdate

End Sub


'------------------------------------------------------------
' 姿勢点の判定とグラフ描画のためのデータを生成する処理
'
' 概要:
'   条件設定シートから各点数に対応する姿勢のしきい値を取得し、
'   ポイント計算シート上の各行について、関節角度から姿勢点を判定。
'   判定結果は色分け用の列にも反映。
'
' 備考:
'   - 姿勢点は1点〜10点で評価され、最初に条件を満たした点数が優先される
'   - 判定は腰・膝の角度による範囲チェックで行う
'------------------------------------------------------------
Sub makeGraphZensya()

    Call stopUpdate

    ' 姿勢点の判定に使用するしきい値を格納する配列
    Dim KoshiMax(10)            As Double
    Dim KoshiMin(10)            As Double
    Dim HizaMax(10)             As Double
    Dim HizaMin(10)             As Double
    Dim CaptionName2(10)        As String

    ' 各種変数
    Dim i                       As Long
    Dim j                       As Long
    Dim data_no                 As Long
    Dim max_row_num             As Long
    Dim KoshiAngle              As Double
    Dim HizaAngleL              As Double
    Dim HizaAngleR              As Double
    Dim correctPose             As Boolean
    Dim totalSecond             As Long
    Dim hour                    As Long
    Dim min                     As Long
    Dim sec                     As Long
    Dim t                       As Date
    Dim mSeconds                As String

    ' 条件設定シートから各点数（10〜1点）の条件を取得
    With ThisWorkbook.Worksheets("条件設定シート")
        For j = 1 To 10
            ' 各点数の設定は18行おきに記載されているため、オフセットで取得
            Dim offset As Long: offset = 168 - (j - 1) * 18
            CaptionName2(j) = .Cells(offset, 2)
            KoshiMax(j) = .Cells(offset + 2, 7)
            KoshiMin(j) = .Cells(offset + 3, 7)
            HizaMax(j) = .Cells(offset + 5, 7)
            HizaMin(j) = .Cells(offset + 6, 7)
        Next j
    End With

    ' ポイント計算シートの行数を取得し、各行について判定
    With ThisWorkbook.Sheets("ポイント計算シート")
        max_row_num = getLastRow()

        For i = 2 To max_row_num
            ' 経過時間（秒）を時刻に変換し、書式付き文字列にする
            totalSecond = WorksheetFunction.RoundDown(.Cells(i, 2), 0)
            mSeconds = Right(Format(WorksheetFunction.RoundDown(.Cells(i, 2), 3), "0.000"), 3)
            hour = totalSecond \ 3600
            min = (totalSecond Mod 3600) \ 60
            sec = totalSecond Mod 60
            t = TimeSerial(hour, min, sec)

            ' 該当する姿勢点が見つかったかのフラグ
            correctPose = False

            ' 対象行の関節角度を読み取る（腰・左右膝）
            KoshiAngle = CDbl(.Cells(i, COLUMN_KOSHI_ANGLE).Value)
            HizaAngleL = CDbl(.Cells(i, COLUMN_HIZA_L_ANGLE).Value)
            HizaAngleR = CDbl(.Cells(i, COLUMN_HIZA_R_ANGLE).Value)

            ' 点数の条件をチェック（2点〜10点）
            For j = 2 To 10
                If (KoshiMin(j) >= KoshiAngle And KoshiAngle > KoshiMax(j)) And _
                    ((HizaMin(j) >= HizaAngleL And HizaAngleL > HizaMax(j)) Or _
                    (HizaMin(j) >= HizaAngleR And HizaAngleR > HizaMax(j))) Then

                    ' 条件に一致した場合：該当の点数を記録
                    correctPose = True
                    .Cells(i, COLUMN_ROUGH_TIME).Value = Format(t, "hh:mm:ss") & "," & mSeconds
                    .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value = j
                    .Cells(i, COLUMN_DATA_RESULT_FIX).Value = j
                    Exit For
                End If
            Next j

            ' 該当しなければ1点として扱う
            If Not correctPose Then
                .Cells(i, COLUMN_ROUGH_TIME).Value = Format(t, "hh:mm:ss") & "," & mSeconds
                .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value = 1
                .Cells(i, COLUMN_DATA_RESULT_FIX).Value = 1
            End If
        Next i

        ' 判定結果に応じて、色分け列に値を設定
        For i = 2 To max_row_num
            data_no = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value

            If data_no >= DATA_SEPARATION_GREEN_BOTTOM And data_no <= DATA_SEPARATION_GREEN_TOP Then
                .Cells(i, COLUMN_DATA_RESULT_GREEN).Value = data_no
                .Cells(i, COLUMN_DATA_RESULT_YELLOW).Value = 0
                .Cells(i, COLUMN_DATA_RESULT_RED).Value = 0

            ElseIf data_no >= DATA_SEPARATION_YELLOW_BOTTOM And data_no <= DATA_SEPARATION_YELLOW_TOP Then
                .Cells(i, COLUMN_DATA_RESULT_GREEN).Value = 0
                .Cells(i, COLUMN_DATA_RESULT_YELLOW).Value = data_no
                .Cells(i, COLUMN_DATA_RESULT_RED).Value = 0

            ElseIf data_no >= DATA_SEPARATION_RED_BOTTOM And data_no <= DATA_SEPARATION_RED_TOP Then
                .Cells(i, COLUMN_DATA_RESULT_GREEN).Value = 0
                .Cells(i, COLUMN_DATA_RESULT_YELLOW).Value = 0
                .Cells(i, COLUMN_DATA_RESULT_RED).Value = data_no
            End If
        Next i
    End With

    Call restartUpdate
End Sub


'------------------------------------------------------------
' 姿勢素点の字幕データに含まれるノイズを除去する処理
'
' 引数:
'   fps : フレームレート（frames per second）
'
' 備考:
'   - 判定結果の変化が短時間で元に戻るような場合を「ノイズ」とみなし、除去対象とする
'   - 閾値は CAPTION_REMOVE_NOISE_SECOND 秒に相当するフレーム数
'   - ノイズ判定後、ノイズと思われる行を前の行の内容で上書き
'------------------------------------------------------------
Function removeCaptionNoise(fps As Double)

    Dim max_row_num     As Long
    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim i_max           As Long
    Dim j_max           As Long
    Dim currentValue    As String
    Dim targetValue     As String
    Dim compareValue    As String
    Dim sameValueNum    As Long
    Dim noise_num       As Long

    ' フレーム数に応じてノイズ判定の閾値（最小でも2フレーム）
    noise_num = CAPTION_REMOVE_NOISE_SECOND * fps

    If noise_num < 2 Then noise_num = 2

    Call stopUpdate

    With ThisWorkbook.Sheets("ポイント計算シート")

        max_row_num = getLastRow()
        i_max = max_row_num - noise_num - 1 ' ノイズ判定を安全に行える範囲でループ

        ' 行ごとにキャプションの変化を確認
        For i = 2 To i_max

            currentValue = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value
            targetValue = .Cells(i + 1, COLUMN_DATA_RESULT_ORIGIN).Value

            ' 判定結果に変化がある場合にのみ処理
            If currentValue <> targetValue Then

                sameValueNum = 1
                j_max = i + 1 + noise_num - 1

                ' targetValue が連続して出現する数をカウント
                For j = i + 2 To j_max
                    compareValue = .Cells(j, COLUMN_DATA_RESULT_ORIGIN).Value
                    If targetValue = compareValue Then
                        sameValueNum = sameValueNum + 1
                    Else
                        Exit For
                    End If
                Next

                ' 閾値未満 → ノイズと判定し、前の値で上書き
                If sameValueNum < noise_num Then
                    For k = i + 1 To j - 1
                        If Not IsEmpty(.Cells(i, COLUMN_DATA_RESULT_ORIGIN)) Then
                            For tmp = 0 To 14
                                .Cells(k, COLUMN_DATA_RESULT_ORIGIN + tmp).Value = _
                                    .Cells(i, COLUMN_DATA_RESULT_ORIGIN + tmp).Value
                            Next
                        End If
                    Next
                End If
            End If
        Next
    End With

    Call restartUpdate

End Function


'------------------------------------------------------------
' 姿勢重量点調査票に評価除外・強制フラグの反映と作業時間等を計算・転記する
'
' 概要:
'   - フレームレートから時間を計算
'   - 除外フラグの有無に応じた開始/終了時刻の決定
'   - 姿勢素点や推定・欠損区間の時間を集計
'   - 作業名・作業時間を自動で補完
'
' 備考:
'   - 1回目はPythonからの入力、更新ボタンで再計算
'   - 動画末尾の除外も考慮
'------------------------------------------------------------
Sub fixSheetZensya()

    ' 画面更新を一時停止（パフォーマンス向上用）
    Call stopUpdate

    ' ===== 初期化セクション =====
    ' フレームレートを取得
    Dim fps         As Double: fps = getFps()
    ' 最終行番号を取得（分析対象の最終フレーム）
    Dim max_row_num As Long: max_row_num = getLastRow()
    ' シート拡張行数（自動検出）
    Dim expand_no   As Long: expand_no = 0
    ' 拡張チェックワードが出現するまでループ
    Do While s_ProcessEvaluation_2nd.Cells(29 + expand_no, 3) <> SHIJUTEN_SHEET_EXPAND_NUM_CHECK_WORD And expand_no < 979
        expand_no = expand_no + 1
    Loop

    ' ===== 作業情報用の変数定義と初期化 =====
    Dim i                   As Long
    Dim j                   As Long
    Dim bottom_jogai_start  As Long
    Dim format_time         As String
    Dim lastInputRow        As Long: lastInputRow = 1

    ' 時間関連の変数定義（t: 現在時間、t0: 前の作業終了時間）
    Dim t               As Double: t = s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME + 1).Value
    Dim t0              As Double: t0 = t

    ' フレーム位置
    Dim start_frame     As Long
    Dim end_frame       As Long

    ' 作業時間および除外時間のフレームカウント
    Dim workFrames      As Long
    Dim removeFrames    As Long

    ' 姿勢・状態別カウント格納配列（1〜15）
    Dim data_array(15)  As Long ' 姿勢10要素 + 欠損 + 推定 + 空欄 + 腰曲げ + 膝曲げ
    Dim data_no         As Long

    ' ===== 作業開始・終了時刻の初期設定 =====
    ' 開始時間が空欄の場合、初期値（00:00:00.00）を代入
    If IsEmpty(s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME)) Then
        s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value = "00:00:00.00"
    End If

    ' 終了時間が未入力なら動画全体から換算した終了時間を代入
    If IsEmpty(s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME)) Then
        Dim seconds As Double: seconds = max_row_num / fps
        Dim time_str As String: time_str = timeConvert(seconds)
        s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value = time_str
    End If

    ' ===== 除外フラグによる開始時間調整 =====
    ' 除外フラグを取得（冒頭が除外されているか）
    Dim remove_flag As Variant: remove_flag = s_PointCalc.Cells(2, COLUMN_DATA_REMOVE_SECTION).Value

    ' 除外なし：そのまま00:00:00.00からスタート
    If remove_flag = 0 Then
        s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value = "00:00:00.00"

    ' 除外あり：除外が終わるまでスキップして開始
    ElseIf remove_flag = 1 Then
        Dim top_jogai_end As Long: top_jogai_end = 0
        Do While s_PointCalc.Cells(2 + top_jogai_end, COLUMN_DATA_REMOVE_SECTION) = 1
            top_jogai_end = top_jogai_end + 1
        Loop
        seconds = top_jogai_end / fps
        time_str = timeConvert(seconds)
        s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value = time_str
    End If

    ' ===== 各作業行の処理（開始/終了/時間/名称） =====
    For i = 0 To SHIJUTEN_SHEET_ROW_EXPAND_NUMBER_CHECK - SHIJUTEN_SHEET_ROW_POSESTART_INDEX - 1 + expand_no
        With s_ProcessEvaluation_2nd
            Dim currentStartTime As Variant
            currentStartTime = .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value

            ' 開始時間が空 → 無効行とみなして各セルを初期化
            If IsEmpty(currentStartTime) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME).MergeArea.ClearContents
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).ClearContents
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).ClearContents
                For j = 0 To 10
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX + j).ClearContents
                Next j
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME).ClearContents
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME).ClearContents
            Else
                ' 作業名が空欄 → 自動補完（"作業n"）
                If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME).Value) Then
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME).Value = "作業" & i + 1
                End If

                ' 終了時間の自動補完：次行が空なら末尾から計算
                Dim nextRowStartTime As Variant
                nextRowStartTime = .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i + 1, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value
                If IsEmpty(nextRowStartTime) Then
                    bottom_jogai_start = 0
                    Do While s_PointCalc.Cells(max_row_num - bottom_jogai_start, COLUMN_DATA_REMOVE_SECTION) = 1
                        bottom_jogai_start = bottom_jogai_start + 1
                    Loop
                    seconds = (max_row_num - bottom_jogai_start - 1) / fps
                    format_time = timeConvert(seconds)
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value = format_time
                Else
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value = nextRowStartTime
                End If

                ' 再計算用の更新
                If Not IsEmpty(currentStartTime) Then
                    lastInputRow = SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i - 9
                End If

                Call restartUpdate
                Call stopUpdate

                ' 実作業時間（終了 - 開始）
                Dim startTimeSerial As Double
                Dim endTimeSerial   As Double
                startTimeSerial = .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME + 1).Value
                endTimeSerial = .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME + 1).Value

                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).Value = endTimeSerial - startTimeSerial
            End If
        End With
    Next i

    ' ===== 作業Noに通し番号を割り当てる処理 =====
    For i = 0 To 19 + expand_no
        s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NUMBER).Value = i + 1
    Next i

    ' ===== 各作業ブロックに対して姿勢・状態の集計処理 =====
    For i = 0 To 19 + expand_no
        Dim isWorkValid As Boolean
        isWorkValid = Not IsEmpty(s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME))

        If isWorkValid Then
            ' 対象時間の取得
            Dim separate_work_time As Double
            separate_work_time = s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).Value

            ' 開始/終了フレームの計算
            t0 = t
            t = t + separate_work_time
            start_frame = t0 * fps
            end_frame = t * fps - 1

            ' カウンタ初期化
            For j = 1 To 15
                data_array(j) = 0
            Next j
            removeFrames = 0

            ' フレームごとの状態カウント
            For j = start_frame To end_frame
                data_no = s_PointCalc.Cells(2 + j, COLUMN_DATA_RESULT_FIX).Value
                If data_no >= 1 And data_no <= 10 Then
                    data_array(data_no) = data_array(data_no) + 1
                End If

                If s_PointCalc.Cells(2 + j, COLUMN_DATA_MISSING_SECTION).Value = 1 Then
                    data_array(11) = data_array(11) + 1
                End If
                If s_PointCalc.Cells(2 + j, COLUMN_DATA_PREDICT_SECTION).Value >= 1 Then
                    data_array(12) = data_array(12) + 1
                End If
                If s_PointCalc.Cells(2 + j, COLUMN_DATA_RESULT_GH_KOSHIMAGE).Value = 1 Then
                    data_array(14) = data_array(14) + 1
                End If
                If s_PointCalc.Cells(2 + j, COLUMN_DATA_RESULT_GH_HIZAMAGE).Value = 1 Then
                    data_array(15) = data_array(15) + 1
                End If
                If s_PointCalc.Cells(2 + j, COLUMN_DATA_REMOVE_SECTION).Value = 1 Then
                    removeFrames = removeFrames + 1
                End If

                ' キャプション欄への出力（作業No + 作業名）
                Dim workNoLabel As String
                workNoLabel = "作業No." & s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NUMBER).Value & _
                              " " & s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME).Value
                s_ProcessEvaluation_2nd.Cells(2 + j, COLUMN_CAPTION_WORK_NAME).Value = workNoLabel
            Next j

            ' 除外を除いた実際の作業フレーム数 → 秒に変換
            workFrames = (end_frame + 1 - start_frame) - removeFrames
            s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).Value = workFrames / fps

            ' 姿勢データを秒換算で出力
            For j = 0 To 9
                Dim poseCount As Long
                poseCount = data_array(10 - j)
                If poseCount = 0 Then
                    s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX + j).Value = ""
                Else
                    s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX + j).Value = poseCount / fps
                End If
            Next j

            ' 腰曲げ・膝曲げも同様に出力
            s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME).Value = data_array(14) / fps
            s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME).Value = data_array(15) / fps
        End If
    Next i

End Sub


'------------------------------------------------------------
' 字幕ファイル出力処理
'
' 引数:
'   movieName : 入力動画のファイル名（拡張子なし）
'
' 説明:
'   - ポイント計算シートと姿勢重量点調査票シートの情報を元に、
'     字幕ファイル（.srt形式）を2トラック構成で出力する。
'   - 動画の縦横サイズに応じてフォントサイズを計算し、
'     字幕のレイアウトやカラーリングを調整。
'   - 姿勢素点や評価除外など、評価データを条件に従って出力。
'------------------------------------------------------------
Function outputCaption(movieName As String)
    Dim i                       As Long
    Dim j                       As Long
    Dim video_width             As Long
    Dim video_height            As Long
    Dim fps                     As Double
    Dim max_row_num             As Long
    Dim max_array_num           As Long
    Dim WorkName()              As String
    Dim CaptionName2(10)        As String
    Dim CaptionName2Koshimage   As String
    Dim CaptionName2Hizamage    As String
    Dim Track1FileName          As String
    Dim CaptionName0            As String
    Dim CaptionName1            As String
    Dim CaptionNo2              As Long
    Dim ColorName1              As String
    Dim ColorName2              As String
    Dim Track1OutputString1     As String
    Dim Track1OutputString2     As String
    Dim track1_font_size1       As Long
    Dim track1_font_size2       As Long
    Dim track1_coef_font_size1  As Long
    Dim track1_coef_font_size2  As Long

    ' 描画抑止
    stopUpdate

    ' 動画サイズ取得
    With ThisWorkbook.Sheets("ポイント計算シート")
        video_width = .Cells(2, 198)
        video_height = .Cells(2, 197)
    End With

    ' サイズに応じた係数設定
    If video_width < video_height Then
        track1_coef_font_size1 = TRACK1_TATE_UPPER_COEF
        track1_coef_font_size2 = TRACK1_TATE_LOWER_COEF
    Else
        track1_coef_font_size1 = TRACK1_YOKO_UPPER_COEF
        track1_coef_font_size2 = TRACK1_YOKO_LOWER_COEF
    End If

    ' フォントサイズ設定
    track1_font_size1 = video_width / track1_coef_font_size1
    track1_font_size2 = video_width / track1_coef_font_size2

    ' 姿勢素点キャプションの読み込み
    With ThisWorkbook.Sheets("条件設定シート")
        CaptionName2(10) = .Cells(6, 2)
        CaptionName2(9) = .Cells(24, 2)
        CaptionName2(8) = .Cells(42, 2)
        CaptionName2(7) = .Cells(60, 2)
        CaptionName2(6) = .Cells(78, 2)
        CaptionName2(5) = .Cells(96, 2)
        CaptionName2(4) = .Cells(114, 2)
        CaptionName2(3) = .Cells(132, 2)
        CaptionName2(2) = .Cells(150, 2)
        CaptionName2(1) = .Cells(168, 2)
        CaptionName2Koshimage = .Cells(210, 2)
        CaptionName2Hizamage = .Cells(228, 2)
    End With
    CaptionName2(0) = "0-姿勢評価なし"

    ' ファイル名設定
    Track1FileName = ActiveWorkbook.Path & "\" & movieName & ".srt"

    ' フレームレートと行数取得
    fps = getFps()
    max_row_num = getLastRow()
    max_array_num = max_row_num - 2
    ReDim WorkName(max_array_num, 0)

    ' 作業名の読み込み
    Call fillWorkNames(WorkName, fps, max_array_num)

    ' 字幕ファイル書き出し
    Open Track1FileName For Output As #1
    With ThisWorkbook.Sheets("ポイント計算シート")
        For i = 2 To max_row_num
            CaptionName0 = WorkName(i - 2, 0)

            ' 評価区間の条件
            Select Case True
                Case .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0
                    CaptionName1 = CAPTION_DATA_REMOVE_SECTION
                    ColorName1 = COLOR_DATA_REMOVE_SECTION
                Case .Cells(i, COLUMN_DATA_FORCED_SECTION).Value > 0
                    CaptionName1 = CAPTION_DATA_FORCED_SECTION
                    ColorName1 = COLOR_DATA_FORCED_SECTION
                Case .Cells(i, COLUMN_DATA_MISSING_SECTION).Value > 0
                    CaptionName1 = CAPTION_DATA_MISSING_SECTION
                    ColorName1 = COLOR_DATA_MISSING_SECTION
                Case .Cells(i, COLUMN_DATA_MEASURE_SECTION).Value > 0
                    CaptionName1 = CAPTION_DATA_MEASURE_SECTION
                    ColorName1 = COLOR_DATA_MEASURE_SECTION
                Case .Cells(i, COLUMN_DATA_PREDICT_SECTION).Value > 0
                    CaptionName1 = CAPTION_DATA_PREDICT_SECTION
                    ColorName1 = COLOR_DATA_PREDICT_SECTION
            End Select

            ' 姿勢素点の条件
            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
                CaptionNo2 = 0
                ColorName2 = COLOR_DATA_REMOVE_SECTION
            Else
                CaptionNo2 = .Cells(i, COLUMN_DATA_RESULT_FIX).Value
                If .Cells(i, COLUMN_DATA_RESULT_GREEN).Value > 0 Then
                    ColorName2 = COLOR_DATA_RESULT_GREEN
                ElseIf .Cells(i, COLUMN_DATA_RESULT_YELLOW).Value > 0 Then
                    ColorName2 = COLOR_DATA_RESULT_YELLOW
                ElseIf .Cells(i, COLUMN_DATA_RESULT_RED).Value > 0 Then
                    ColorName2 = COLOR_DATA_RESULT_RED
                End If
            End If

            ' 字幕構築
            Track1OutputString1 = "<font size=\"" & track1_font_size1 & "\" color =\"#ffffff\">" & CaptionName0 & "</font>" & _
                                  "<font size=\"" & track1_font_size1 & "\" color =\"" & ColorName1 & "\">" & CaptionName1 & "</font>"

            Track1OutputString2 = "<font size=\"" & track1_font_size2 & "\" color =\"" & ColorName2 & "\">" & CaptionName2(CaptionNo2) & "</font>"

            ' ファイル出力
            Print #1, " " & i - 1
            Print #1, .Cells(i, COLUMN_ROUGH_TIME).Value & " --> " & .Cells(i + 1, COLUMN_ROUGH_TIME).Value
            Print #1, Replace(Track1OutputString1, vbLf, vbCrLf)
            Print #1, Replace(Track1OutputString2, vbLf, vbCrLf)
            Print #1, ""
            Print #1, ""
            .Cells(i, COLUMN_CAPTION_WORK_NAME).Clear

            ' 初期化
            ColorName1 = "#ffffff"
            ColorName2 = "#ffffff"
        Next
    End With
    Close #1

    ' 描画再開
    restartUpdate
End Function


'帳票更新ボタンが押された時の処理
' 引数  ：なし
' 戻り値：なし
Function ClickUpdateData()
    Dim tstart_click As Double
    Dim dotPoint     As String
    Dim workbookName As String
    Dim fps          As Double

    tstart_click = Timer
    fps = getFps()

    'ノイズ除去
    Call removeCaptionNoise(fps)

    '作業分割、時間測定
    Call fixSheetZensya

    ' 拡張子を除いたブック名を取得
    dotPoint = InStrRev(ActiveWorkbook.Name, ".")
    workbookName = Left(ActiveWorkbook.Name, dotPoint - 1)

    ' キャプションを出力
    Call outputCaption(workbookName)
    Debug.Print " 更新時間" & Format$(Timer - tstart_click, "0.00") & " sec."

End Function


' 概要 : 関節角度、3dデータのcsvをコピー貼り付けする
' 呼び元のシート : マクロテスト
' 補足 : 本ファイルと同じディレクトリにcsvファイルを置いておく
' 引数1 ：フレームレート
' 引数2 ：動画横幅の値
' 引数3 ：csvファイル名
' 引数4 ：動画縦の値 動画の向きによって字幕文字サイズを調整するために使用
' 戻り値：なし
Sub MacroInput3dData(fps As Double, video_width As Long, csv_file_name As String, video_height As Double)

    Dim wb     As Variant
    Dim ws     As Variant
    Dim MaxRow As Long
    Dim i      As Long

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    s_PointCalc.Visible = True
    Sheets("ポイント計算シート").Select
    Range("D2").Select

    Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & csv_file_name)

    With wb
        Set ws = .Sheets(1)

        Range("B2").Select
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Copy

        'このブックの「貼付先」シートへ値貼り付け
        ThisWorkbook.Worksheets("ポイント計算シート").Range("D2").PasteSpecial _
            xlPasteValuesAndNumberFormats

        'コピー状態を解除
        Application.CutCopyMode = False

        '保存せず終了
        .Close False
    End With

    ' A から C の時間を表すセルを実体化させる
    ' angle.csvを張り付けたあとの最下行番号を取得する
    MaxRow = Range("D2").End(xlDown).row
    For i = 0 To MaxRow - 2
        Range("A" & i + 2).Value = i
        Range("B" & i + 2).Value = i * (1 / fps)
        Range("C" & i + 2).FormulaR1C1 = "=LEFT(TEXT(RC[-1]/(24*60*60), ""hh:mm:ss.000""), 8)"
    Next

    'fps値の保存
    ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 199) = fps
    'video_width値の保存
    ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 198) = video_width
    'video_height値の保存 動画の向きによって字幕文字サイズを調整するために使用
    ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 197) = video_height

    ThisWorkbook.Save

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    s_PointCalc.Visible = xlSheetVeryHidden
End Sub


'Pythonから呼び出しされる
' 引数1 ：動画名
' 引数2 ：フレームレート
' 戻り値：なし
Sub MacroUpdateData(movieName As String, fps As Double)

    Dim tstart_first As Double

    If MEAGERE_TIME_MACROUPDATEDATA = True Then 'MacroUpdateDataの処理時間を測定する
        tstart_first = Timer
    End If

    With ThisWorkbook.Sheets("ポイント計算シート")
        Dim max_row_num As Long
        Dim i As Long

        '処理する行数を取得（3列目の最終セル）
        max_row_num = .Cells(1, 3).End(xlDown).row

        '★★★本処理は、将来的にPythonコード側で行う予定★★★
        'フラグが入力されるセルに入力されているスペースを検索して消去する
        'メイン字幕の姿勢素点の色が全て緑になる不具合の暫定対策
        'セル範囲が広すぎてメモリ不足になるため、for文で処理を細分化
        For i = 4 To 253
            .Range(.Cells(2, i), .Cells(max_row_num, i)).Replace " ", ""
        Next

        'fps値の保存
        fps = getFps()

    End With

    '姿勢判定
    Call makeGraphJisya
    Call makeGraphZensya

    'ノイズ除去
    Call removeCaptionNoise(fps)

    '作業分割、時間測定
    Call fixSheetZensya

    '修正シートの更新
    Call Module1.paintAll

    '字幕生成
    Call outputCaption(movieName)

    'MacroUpdateDataの処理時間を測定する
    If MEAGERE_TIME_MACROUPDATEDATA = True Then
        ThisWorkbook.Sheets("ポイント計算シート").Cells(2, COLUMN_MEAGERE_TIME_MACROUPDATEDATA) = Format$(Timer - tstart_first, "0.00")
    End If

    '初回分析済みのフラグを立てる
    ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 196) = 1

End Sub


Sub OutputOtrs()

    Dim max_row_num    As Long
    Dim i              As Long

    Dim targetRowCount As Long
    Dim writePoseNum   As Long
    Dim lastPoseNum    As Long
    Dim currentTime    As Double
    Dim lastTime       As Double
    Dim ret            As Long
    Dim destFilePath   As String
    Dim sourceFilePath As String

    Dim ReturnBook     As Workbook
    Dim targetWorkbook As Workbook
    Dim strYYYYMMDD    As String
    Dim PosExt         As Long
    Dim StrFileName    As String

    StrFileName = ThisWorkbook.Name
    PosExt = InStrRev(StrFileName, ".")

    '--- 拡張子を除いたパス（ファイル名）を格納する変数 ---'
    Dim strFileExExt As String

    If (0 < PosExt) Then
        StrFileName = Left(StrFileName, PosExt - 1)
    End If

    'Now関数で取得した現在日付をFormatで整形して変数に格納
    strYYYYMMDD = Format(Now, "yyyymmdd_HHMMSS_")

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    Set ReturnBook = ActiveWorkbook
    destFilePath = ActiveWorkbook.Path & "\" & StrFileName & "_otrs.xlsx"

    'もしotrs用ファイルがあれば、一度削除しておく
    If Dir(destFilePath) <> "" Then
        Kill destFilePath
    End If

    '作業用のワークブックのインスタンスを作る

    If Dir(destFilePath) = "" Then
        '新しいファイルを作成
        Set targetWorkbook = Workbooks.Add
        '新しいファイルをVBAを実行したファイルと同じフォルダ保存
        targetWorkbook.SaveAs destFilePath
    Else
        Set targetWorkbook = Workbooks.Open(destFilePath)
    End If

    ReturnBook.Activate
    lastPoseNum = -1
    lastTime = 0

    Dim CaptionName2(10) As String

    With ThisWorkbook.Worksheets("条件設定シート")
        CaptionName2(10) = .Cells(6, 2)
        CaptionName2(9) = .Cells(24, 2)
        CaptionName2(8) = .Cells(42, 2)
        CaptionName2(7) = .Cells(60, 2)
        CaptionName2(6) = .Cells(78, 2)
        CaptionName2(5) = .Cells(96, 2)
        CaptionName2(4) = .Cells(114, 2)
        CaptionName2(3) = .Cells(132, 2)
        CaptionName2(2) = .Cells(150, 2)
        CaptionName2(1) = .Cells(168, 2)
    End With

    CaptionName2(0) = "データなし"
        '以下のパターン以外はその他とする。
        '(10) 膝を曲げ上半身前屈(30°～90°)
        '(9) 膝を曲げ上半身前屈(15°～30°)
        '(8) 上半身前屈(45°～90°)
        '(7) 上半身前屈(30°～45°)
        '(6) 上半身前屈(90°～180°)
        '(4) 蹲踞または片膝つき蹲踞
        '(2) 上半身前屈(15°～30°)
        '(1) 基本の立ち姿勢
        '(0) 他"

    With ThisWorkbook.Sheets("ポイント計算シート")
        max_row_num = getLastRow()
        targetRowCount = 1
        Dim lastI As Long

        For i = 2 To max_row_num
            'COLUMN_DATA_RESULT_ORIGINが空白の可能性があるため一旦その他を入れておく
            writePoseNum = 0
            On Error Resume Next
            writePoseNum = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value 'キャプション番号のセル代入

            '最初に別のポーズに変わった時が欲しいので一回目は同一にする。
            If i = 2 Then
                lastPoseNum = writePoseNum
                lastI = i - 2
            End If

            If lastPoseNum <> writePoseNum Then
                '同一ポーズを取っていた時間が必要（切り替わった一個前の時間）
                currentTime = .Cells(i - 1, 2).Value
                '書き込み処理
                targetWorkbook.Activate
                targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NAME).Value = CaptionName2(lastPoseNum)
                targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_KEEP_TIME).Value = Round(currentTime - lastTime, 5)
                lastI = i - 2
                targetRowCount = targetRowCount + 1

                lastTime = currentTime
                lastPoseNum = writePoseNum

                ReturnBook.Activate
            End If
        Next

        'ループ終了後に最後に取っていた姿勢が継続しているならそれを書き込む
        If lastPoseNum = writePoseNum Then
            currentTime = .Cells(i - 1, 2).Value
            '書き込み処理
            targetWorkbook.Activate
            targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NAME).Value = CaptionName2(writePoseNum)
            targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_KEEP_TIME).Value = Round(currentTime - lastTime, 5)
            ReturnBook.Activate
        End If
    End With
    s_ProcessEvaluation_2nd.Activate
    ThisWorkbook.Save
    targetWorkbook.Close savechanges:=True
End Sub


'------------------------------------------------------------
' 秒数を hh:mm:ss.ffff 形式の文字列に変換する関数
'
' 引数:
'   seconds : 変換対象の秒数（小数あり）
'
' 戻り値:
'   hh:mm:ss.ffff 形式の文字列（ミリ秒は4桁）
'
' 備考:
'   - 小数部はミリ秒（1/10000）として切り出し
'   - 時間・分・秒をゼロ埋めで整形
'------------------------------------------------------------
Function timeConvert(seconds As Double) As String

    Dim ms As Long
    Dim sec As Long
    Dim min As Long
    Dim hr As Long

    ' 小数部からミリ秒（1/10000）を計算（ずれ防止のため先に処理）
    ms = (seconds - Int(seconds)) * 10000

    ' 秒数を整数に変換（時・分・秒用）
    seconds = Int(seconds)

    ' 時・分・秒を算出
    sec = seconds Mod 60
    min = (seconds \ 60) Mod 60
    hr = seconds \ 3600

    ' フォーマットして返却（hh:mm:ss.ffff）
    timeConvert = Format(hr, "00") & ":" & _
                  Format(min, "00") & ":" & _
                  Format(sec, "00") & "." & _
                  Format(ms, "0000")

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