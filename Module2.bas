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
' 姿勢点の判定とグラフ用データ生成
'
' 引数:
'   なし
'
' 処理概要:
'   - 条件設定シートから各姿勢点（1〜10点）のしきい値と名称を読み込み
'   - 条件設定シートの構造に基づいて、逆順で読み出し
'   - 各フレームに対して姿勢点を判定し、出力列に記録
'   - グラフ色分け用のデータ列を生成
'
' 備考:
'   - 表示更新を一時停止して処理速度を改善
'------------------------------------------------------------
Sub makeGraphZensya()

    Call stopUpdate

    Dim KoshiMax(10)            As Double
    Dim KoshiMin(10)            As Double
    Dim HizaMax(10)             As Double
    Dim HizaMin(10)             As Double
    Dim CaptionName2(10)        As String
    Dim CaptionName3Koshimage   As String
    Dim CaptionName3Hizamage    As String
    Dim Koshimage               As Double
    Dim Hizamage                As Double
    Dim i                       As Long
    Dim j                       As Long
    Dim data_no                 As Long
    Dim max_row_num             As Long
    Dim KoshiAngle              As Double
    Dim HizaAngleL              As Double
    Dim HizaAngleR              As Double
    Dim correctPose             As Boolean
    Dim mSeconds                As String
    Dim totalSecond             As Long
    Dim hour                    As Long
    Dim min                     As Long
    Dim sec                     As Long
    Dim t                       As Date
    Dim ds                      As String

    ' 姿勢のしきい値と名称をまとめて読み出し
    With ThisWorkbook.Worksheets("条件設定シート")
        Dim rowOffsets As Variant
        rowOffsets = Array(168, 150, 132, 114, 96, 78, 60, 42, 24, 6) ' 1〜10点の起点行（逆順）

        For j = 1 To 10
            Dim baseRow As Long
            baseRow = rowOffsets(j - 1)
            CaptionName2(j) = .Cells(baseRow, 2)
            KoshiMax(j) = .Cells(baseRow + 2, 7)
            KoshiMin(j) = .Cells(baseRow + 3, 7)
            HizaMax(j) = .Cells(baseRow + 5, 7)
            HizaMin(j) = .Cells(baseRow + 6, 7)
        Next j

        CaptionName3Koshimage = .Cells(210, 2)
        Koshimage = .Cells(212, 7)
        CaptionName3Hizamage = .Cells(228, 2)
        Hizamage = .Cells(230, 7)
    End With

    With ThisWorkbook.Sheets("ポイント計算シート")
        max_row_num = getLastRow()

        For i = 2 To max_row_num
            KoshiAngle = CDbl(.Cells(i, COLUMN_KOSHI_ANGLE).Value)
            HizaAngleL = CDbl(.Cells(i, COLUMN_HIZA_L_ANGLE).Value)
            HizaAngleR = CDbl(.Cells(i, COLUMN_HIZA_R_ANGLE).Value)
            mSeconds = Right(Format(WorksheetFunction.RoundDown(.Cells(i, 2), 3), "0.000"), 3)
            totalSecond = WorksheetFunction.RoundDown(.Cells(i, 2), 0)
            hour = totalSecond \ 3600
            min = (totalSecond Mod 3600) \ 60
            sec = totalSecond Mod 60
            t = TimeSerial(hour, min, sec)
            ds = Format(t, "hh:mm:ss")
            correctPose = False

            For j = 2 To 10
                If KoshiMin(j) >= KoshiAngle And KoshiMax(j) < KoshiAngle And _
                   ((HizaMin(j) >= HizaAngleL And HizaAngleL > HizaMax(j)) Or _
                    (HizaMin(j) >= HizaAngleR And HizaAngleR > HizaMax(j))) Then

                    .Cells(i, COLUMN_ROUGH_TIME).Value = ds & "," & mSeconds
                    .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value = j
                    .Cells(i, COLUMN_DATA_RESULT_FIX).Value = j
                    correctPose = True
                    Exit For
                End If
            Next j

            If Not correctPose Then
                .Cells(i, COLUMN_ROUGH_TIME).Value = ds & "," & mSeconds
                .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value = 1
                .Cells(i, COLUMN_DATA_RESULT_FIX).Value = 1
            End If
        Next i

        ' グラフ描画用フラグ（色分け）
        For i = 2 To max_row_num
            data_no = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value
            .Cells(i, COLUMN_DATA_RESULT_GREEN).Value = IIf(data_no >= DATA_SEPARATION_GREEN_BOTTOM And data_no <= DATA_SEPARATION_GREEN_TOP, data_no, 0)
            .Cells(i, COLUMN_DATA_RESULT_YELLOW).Value = IIf(data_no >= DATA_SEPARATION_YELLOW_BOTTOM And data_no <= DATA_SEPARATION_YELLOW_TOP, data_no, 0)
            .Cells(i, COLUMN_DATA_RESULT_RED).Value = IIf(data_no >= DATA_SEPARATION_RED_BOTTOM And data_no <= DATA_SEPARATION_RED_TOP, data_no, 0)
        Next i
    End With

    Call restartUpdate

End Sub


'姿勢素点の字幕、フラグのノイズを消去する
' 引数1 ：フレームレート
' 戻り値：なし
Function removeCaptionNoise(fps As Double)

    Dim max_row_num   As Long
    Dim max_array_num As Long

    Dim i             As Long
    Dim j             As Long
    Dim k             As Long
    Dim tmp           As Long

    Dim i_max         As Long
    Dim j_max         As Long
    Dim k_max         As Long

    Dim currentValue  As String
    Dim targetValue   As String
    Dim compareValue  As String

    Dim sameValueNum  As Long
    Dim noise_num     As Long: noise_num = CAPTION_REMOVE_NOISE_SECOND * fps

    If noise_num < 2 Then
        noise_num = 2
    End If

    '表示・更新をオフにする
    Call stopUpdate

    With ThisWorkbook.Sheets("ポイント計算シート")

        '処理する行数を取得（3列目の最終セル）
        max_row_num = getLastRow()
        max_array_num = max_row_num - 1 - 1 '2行目からセルに値が入るため-1、配列は0から使うため-1

        '下方向へ探索する際の起点(i), 終点(i_max)
        i_max = max_row_num - noise_num - 1

        'キャプションのノイズ除去
        For i = 2 To i_max

            currentValue = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value
            targetValue = .Cells(i + 1, COLUMN_DATA_RESULT_ORIGIN).Value

            '判定結果が変わったとき
            If currentValue <> targetValue Then

                'ノイズかどうか探索する 起点(j), 終点(j_max)
                j_max = i + 1 + noise_num - 1
                sameValueNum = 1
                For j = i + 2 To j_max
                    compareValue = .Cells(j, COLUMN_DATA_RESULT_ORIGIN).Value
                    '判定結果が変わったらループを抜ける
                    If targetValue = compareValue Then
                        sameValueNum = sameValueNum + 1
                    Else
                        Exit For
                    End If
                Next

                'ノイズが見つかったときの処理
                If sameValueNum < noise_num Then
                    For k = i + 1 To j
                        If Not IsEmpty(.Cells(i, COLUMN_DATA_RESULT_ORIGIN)) Then
                            For tmp = 0 To 14
                                .Cells(k, COLUMN_DATA_RESULT_ORIGIN + tmp) = .Cells(i, COLUMN_DATA_RESULT_ORIGIN + tmp)
                            Next
                        End If
                    Next
                End If
            End If
        Next
    End With

    '表示・更新をオンに戻す
    Call restartUpdate
End Function


'秒をhh:mm:ss:msに変換する
Function timeConvert(seconds As Double) As String

    Dim milliseconds        As Long
    Dim remainingSeconds    As Long
    Dim minutes             As Long
    Dim hours               As Long

    'ずれ防止のために小数点以下を切り捨てミリ秒・秒から先に出す
    milliseconds = (seconds - Int(seconds)) * 10000
    seconds = Int(seconds)

    remainingSeconds = seconds Mod 60
    minutes = (seconds Mod 3600) \ 60
    hours = seconds \ 3600

    timeConvert = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(remainingSeconds, "00") & "." & Format(milliseconds, "0000")
End Function


'姿勢重量点調査票で指定された評価除外、評価強制をポイント計算シートに反映させる
'ポイント計算シートのフラグから時間を計算して、姿勢重量点調査票に転記する
'１回目はPythonプログラムから値をもらう
'更新ボタンを押されたときはポイント計算シートから値を読み取る
' 引数1 ：フレームレート
' 戻り値：なし
Sub fixSheetZensya()

    '表示・更新をオフにする
    Call stopUpdate
    Dim fps                         As Double
    Dim separate_work_time          As Double 'tとt0の差を取得する
    Dim t0                          As Double '1つ前のtを一時保存する
    Dim t                           As Double '作業時間

    Dim i                           As Long
    Dim j                           As Long
    Dim k                           As Long
    Dim l                           As Long

    Dim max_row_num                 As Long

    Dim expand_no                   As Long '処理行数拡張用
    Dim data_flag                   As Long '姿勢素点の データ除外（0） または データ強制（1～10）フラグ記憶用 左記に該当しない場合は-1を入れて使う

    Dim top_jogai_end               As Long
    Dim bottom_jogai_start          As Long

    Dim koshimage_flag              As Long '腰曲げの データ除外（0）または データ強制（1） フラグ記憶用 左記に該当しない場合は-1を入れて使う
    Dim hizamage_flag               As Long '膝曲げの データ除外（0）または データ強制（1） フラグ記憶用 左記に該当しない場合は-1を入れて使う

    Dim start_frame                 As Long
    Dim end_frame                   As Long
    Dim start_array_num             As Long
    Dim end_array_num               As Long

    Dim data_array(15)              As Long '姿勢重量点 1 ~ 10 点、欠損区間、推定区間、拳上、腰曲げ、膝曲げの時間を合計するために使用
    Dim data_no                     As Long  'data_arrayの配列番号。1～10:姿勢重量点 11:欠損区間 12:推定区間 13:拳上 14:腰曲げ 15:膝曲げ

    Dim removeFrames                As Long
    Dim separate_removeFrames       As Long
    Dim workFrames                  As Long

    Dim separate_koshimage_missing  As Double '作業分割後　腰曲げ欠損区間
    Dim separate_koshimage_predict  As Double '作業分割後　腰曲げ推定区間
    Dim separate_hizamage_missing   As Double '作業分割後　膝曲げ欠損区間
    Dim separate_hizamage_predict   As Double '作業分割後　膝曲げ推定区間

    Dim seconds                     As Double
    Dim hours                       As String
    Dim minutes                     As String
    Dim remainingSeconds            As String
    Dim milliseconds                As String
    Dim format_time                 As String
    Dim val                         As Double
    Dim lastInputRow                As Long: lastInputRow = 1
    Dim intermediateFlag            As Boolean
    Dim sheet_index                 As Long

    'フレームレートを取得
    fps = getFps()

    'ポイント計算シートの最終行を取得
    max_row_num = getLastRow()

    '処理する追加行数を取得する
    'その他（時間計7.5H）のセル位置の移動量を調べる  ※最大999行(<979)にする
    expand_no = 0
    Do While s_ProcessEvaluation_2nd.Cells(29 + expand_no, 3) <> SHIJUTEN_SHEET_EXPAND_NUM_CHECK_WORD And expand_no < 979
        expand_no = expand_no + 1
    Loop

    'ここから初回分析のための処理
    '作業開始時間が空の場合は、0.0を入力
    If IsEmpty(s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME)) = True Then
        s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value = "00:00:00.00"
    End If

    '作業終了時間が空の場合は、ポイント計算シート最終行から計算して入力
    If IsEmpty(s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME)) = True Then
        seconds = max_row_num / fps 'ここに変換したい秒数を入力してください

        format_time = timeConvert(seconds)

        s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value = format_time
    End If

    'ここから帳票更新のための処理
    '動画の先頭に除外がある場合、除外の末尾より一つ下のセルから１つ目の作業開始時間を計算する

    '除外フラグの先頭が0の時
    If s_PointCalc.Cells(2, COLUMN_DATA_REMOVE_SECTION) = 0 Then
        '0秒にする
        s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value = "00:00:00.00"

    '除外フラグの先頭が1の時
    ElseIf s_PointCalc.Cells(2, COLUMN_DATA_REMOVE_SECTION) = 1 Then
        'リセット
        top_jogai_end = 0
        '除外の末尾を調べる
        '除外フラグが1でなくなるまでループ
        Do While s_PointCalc.Cells(2 + top_jogai_end, COLUMN_DATA_REMOVE_SECTION) = 1
            top_jogai_end = top_jogai_end + 1
        Loop

        '除外の終了時間を計算して開始時間の１行目に入力
        seconds = top_jogai_end / fps 'ここに変換したい秒数を入力してください

        format_time = timeConvert(seconds)

        s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value = format_time
    End If


    'ここから作業分割に関する処理
    For i = 0 To SHIJUTEN_SHEET_ROW_EXPAND_NUMBER_CHECK - SHIJUTEN_SHEET_ROW_POSESTART_INDEX - 1 + expand_no
        '作業開始時間が空なら
        If IsEmpty(s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME)) Then
            '作業名、作業終了時間、作業時間、拳上、腰曲げ、膝曲げを空にする
            s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME).MergeArea.ClearContents 'セル結合があるため
            s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).ClearContents
            s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).ClearContents

            '姿勢素点とひねりを空にする
            For j = 0 To 10
                '姿勢要素時間（フレーム数）が0のときは、空白セルにする
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX + j).ClearContents
            Next

            'NG時間を空にする
            s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME).ClearContents
            s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME).ClearContents

        '作業開始時間が入力されているなら
        Else
            'ここから作業名の入力
            '作業名が空なら入力する
            If IsEmpty(s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME)) Then
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME) = "作業" & i + 1
            End If

            'ここから作業終了時間の入力
            '１つ先の行の作業開始時間が空の時
            If IsEmpty(s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i + 1, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME)) Then

                'カウントリセット
                bottom_jogai_start = 0
                'max_row_num行目から一つずつ上がって、除外の先頭位置を探す
                Do While s_PointCalc.Cells(max_row_num - bottom_jogai_start, COLUMN_DATA_REMOVE_SECTION) = 1
                    bottom_jogai_start = bottom_jogai_start + 1
                Loop

                '動画末尾にある除外の開始時間を計算して入力
                seconds = (max_row_num - bottom_jogai_start - 1) / fps

                format_time = timeConvert(seconds)

                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value = format_time


            '１つ先の行の作業開始時間に値がある時、その値を入れる
            Else
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value _
                    = s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i + 1, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value
            End If

            If Not IsEmpty(s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME)) Then
                lastInputRow = SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i - 9 '上部のカウントを-しておく
            End If

            '行程評価シートで計算式が入力されたセルの値を更新する
            Call restartUpdate
            Call stopUpdate

            '作業終了時間と作業開始時間から作業時間を計算してセルに入力
            s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).Value = _
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME + 1).Value _
                - s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME + 1).Value
        End If
    Next

    '作業No.の代入
    For i = 0 To 19 + expand_no
        s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NUMBER).Value = i + 1
    Next

    '最初の作業（除外後の開始時刻）をシートから読み取り、t0 = t = 実際の開始秒 で初期化する。
    t = s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME + 1).Value
    t0 = t

    For i = 0 To 19 + expand_no
        '黄色セルの時間を読み取る
        '（t0秒～t秒までの姿勢を求める）
        separate_work_time = s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).Value
        t0 = t
        t = t + separate_work_time '作業時間を単一で入力する場合

        '--------------------------------------------------
        'start_frameの考え方
        'start_frameはt0秒のフレーム数とする
        '例：t0 = 5秒、fps = 30の場合、start_frameは5 * 30 = 150
        '
        'end_frameの考え方
        '動画のフレーム数は0から始まるので、end_frameはt秒のフレーム数-1とする
        '例：t = 10秒、fps = 30の場合、end_frameは10 * 30 - 1 = 299
        '
        'vbaのfor文は、範囲の終わりの値を含むので、start_frame To end_frameで処理する
        'jishaやoutputCaptionなどでも同様の考え方で処理している
        '--------------------------------------------------

        '秒数からフレーム数へ変換
        start_frame = t0 * fps
        end_frame = t * fps - 1

        '姿勢要素時間を入れる変数の初期化
        For j = 1 To 15
            data_array(j) = 0
        Next

        '欠損推定区間をカウントする変数の初期化
        separate_koshimage_missing = 0
        separate_koshimage_predict = 0
        separate_hizamage_missing = 0
        separate_hizamage_predict = 0

        'start_frameフレーム(t0秒) から end_frameフレーム(t秒) までの処理
        '作業時間が入力されている行のみ処理を実行する
        If Not IsEmpty(s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME)) Then

            Debug.Print ("start_rnd:" & start_frame & " " & end_frame)

            '分割地の除外フレームカウントを初期化
            removeFrames = 0

            For j = start_frame To end_frame

                '姿勢素点の値をカウント
                data_no = s_PointCalc.Cells(2 + j, COLUMN_DATA_RESULT_FIX).Value
                data_array(data_no) = data_array(data_no) + 1

                'データ欠損区間をカウント
                data_no = s_PointCalc.Cells(2 + j, COLUMN_DATA_MISSING_SECTION).Value
                If data_no = 1 Then
                    data_array(11) = data_array(11) + 1
                End If

                'データ推定区間をカウント
                data_no = s_PointCalc.Cells(2 + j, COLUMN_DATA_PREDICT_SECTION).Value
                If data_no >= 1 And data_no <= 10 Then
                    data_array(12) = data_array(12) + 1
                End If

                '腰曲げフラグをカウント
                data_no = s_PointCalc.Cells(2 + j, COLUMN_DATA_RESULT_GH_KOSHIMAGE).Value
                If data_no = 1 Then
                    data_array(14) = data_array(14) + 1
                End If

                '膝曲げフラグをカウント
                data_no = s_PointCalc.Cells(2 + j, COLUMN_DATA_RESULT_GH_HIZAMAGE).Value
                If data_no = 1 Then
                    data_array(15) = data_array(15) + 1
                End If

                '腰曲げ欠損をカウント
                If s_PointCalc.Cells(2 + j, COLUMN_DATA_KOSHIMAGE_MISSING_SECTION).Value = 1 Then
                    separate_koshimage_missing = separate_koshimage_missing + 1
                End If

                '腰曲げ推定をカウント
                If s_PointCalc.Cells(2 + j, COLUMN_DATA_KOSHIMAGE_PREDICT_SECTION).Value = 1 Then
                    separate_koshimage_predict = separate_koshimage_predict + 1
                End If

                '膝曲げ欠損をカウント
                If s_PointCalc.Cells(2 + j, COLUMN_DATA_HIZAMAGE_MISSING_SECTION).Value = 1 Then
                    separate_hizamage_missing = separate_hizamage_missing + 1
                End If

                '膝曲げ推定をカウント
                If s_PointCalc.Cells(2 + j, COLUMN_DATA_HIZAMAGE_PREDICT_SECTION).Value = 1 Then
                    separate_hizamage_predict = separate_hizamage_predict + 1
                End If

                '除外区間をカウント
                If s_PointCalc.Cells(2 + j, COLUMN_DATA_REMOVE_SECTION).Value = 1 Then
                    removeFrames = removeFrames + 1
                End If

                'ポイント計算シートのキャプション列へ、姿勢重量点調査票の作業No,と作業名を読み取り、
                '"作業No.xxx_作業名 "として入れておく
                s_ProcessEvaluation_2nd.Cells(2 + j, COLUMN_CAPTION_WORK_NAME).Value = _
                "作業No." & _
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NUMBER).Value & _
                " " & _
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME).Value & _
                " "

            Next

            '作業時間合計値を算出
            workFrames = (end_frame + 1 - start_frame) - removeFrames
            s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).Value = workFrames / fps

            '除外フラグの個数がフレーム全体の時と一致していたときはのみ,全ての行の時間を0,開始終了を-にする
            If top_jogai_end + 1 = max_row_num Then
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).Value = 0
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value = "-"
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value = "-"
            End If

        End If

        '姿勢要素10～1に対する個別処理
        For j = 0 To 9

            If data_array(10 - j) = 0 Then
                '姿勢要素時間（フレーム数）が0のときは、空白セルにする
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX + j).Value = ""
            Else
                '姿勢要素時間（フレーム数）があれば代入する
                s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX + j).Value = data_array(SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX - j) / fps
            End If

CONTINUE:

        Next

    Next

    '表示・更新をオンに戻す
    Call restartUpdate

End Sub



'字幕ファイル出力
'引数1 ：動画名
'戻り値：なし
Function outputCaption(movieName As String)
    Dim max_array_num               As Long
    Dim i                           As Long
    Dim j                           As Long
    Dim max_row_num                 As Long

    '動画の縦横を比較して文字サイズ調整するため、幅・高さどちらも使用する
    Dim video_width                 As Long '入力動画の幅 ※3Dポーズが結合された幅ではないため注意
    Dim video_height                As Long '入力動画の高さ

    '※coefはcoefficient（係数、率）の略記
    Dim track1_coef_font_size1      As Long '字幕トラック1用  上段のサイズ調整用係数
    Dim track1_coef_font_size2      As Long '字幕トラック1用  下段のサイズ調整用係数
    Dim track1_font_size1           As Long '字幕トラック1用  上段のサイズ
    Dim track1_font_size2           As Long '字幕トラック1用  下段のサイズ

    Dim track2_coef_font_size1      As Long '字幕トラック2用 1段目のサイズ調整用係数
    Dim track2_coef_font_size2      As Long '字幕トラック2用 2段目のサイズ調整用係数
    Dim track2_coef_font_size3      As Long '字幕トラック2用 3段目のサイズ調整用係数
    Dim track2_font_size1           As Long '字幕トラック2用 1段目のサイズ
    Dim track2_font_size2           As Long '字幕トラック2用 2段目のサイズ
    Dim track2_font_size3           As Long '字幕トラック2用 3段目のサイズ

    Dim WorkName()                  As String

    Dim CaptionName0                As String  '字幕トラック1用 上段左側 作業名          の字幕文字列
    Dim CaptionName1                As String  '字幕トラック1用 上段右側 帯グラフのデータ（信頼度）の字幕文字列
    Dim CaptionName2(10)            As String  '字幕トラック1用 下段 評価除外(添え字0)+姿勢素点1～10(添え字1～10)の字幕文字列
    Dim CaptionNo2                  As Long 'CaptionName2(10)にアクセスする際の添え字格納用変数

    Dim CaptionName2Koshimage       As String '字幕トラック2用 2段目 腰曲げデータ区間の字幕文字列
    Dim CaptionName2Hizamage        As String '字幕トラック2用 2段目 膝曲げデータ区間の字幕文字列

    Dim CaptionName3Koshimage       As String '字幕トラック2用 ３段目 腰曲げの字幕文字列
    Dim CaptionName3Hizamage        As String '字幕トラック2用 ３段目 膝曲げの字幕文字列

    Dim ColorName1                  As String '字幕トラック1用 上段右側（信頼度 ）の色
    Dim ColorName2                  As String '字幕トラック1用 下段  （姿勢素点）の色
    Dim ColorName2Koshimage         As String '字幕トラック2用 2段目 （腰曲げデータ区間 ）の色
    Dim ColorName2Hizamage          As String '字幕トラック2用 2段目 （膝曲げデータ区間 ）の色
    Dim ColorName3Koshimage         As String '字幕トラック2用 3段目 （腰曲げ ）の色
    Dim ColorName3Hizamage          As String '字幕トラック2用 3段目 （膝曲げ ）の色

    Dim Track1OutputString1         As String '字幕トラック1用：上段文字列
    Dim Track1OutputString2         As String '字幕トラック1用：下段文字列

    Dim Track2OutputString1         As String '字幕トラック2用：1段目文字列
    Dim Track2OutputString2         As String '字幕トラック2用：2段目文字列
    Dim Track2OutputString3         As String '字幕トラック2用：3段目文字列

    Dim Track1FileName              As String '字幕トラック1用のファイル名
    Dim Track2FileName              As String '字幕トラック2用のファイル名

    Dim t                           As Double
    Dim t0                          As Double
    Dim separate_work_time          As Double
    Dim expand_no                   As Long
    Dim start_frame                 As Double
    Dim end_frame                   As Double
    Dim fps                         As Double

    '表示・更新をオフにする
    Call stopUpdate

    '動画の縦横サイズを取得
    video_width = ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 198)
    video_height = ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 197) '動画の縦横判定のために高さも取得

    '動画の縦横によって係数を変更する
    '動画が縦の時
    If video_width < video_height Then
        track1_coef_font_size1 = TRACK1_TATE_UPPER_COEF  '動画が縦のときのトラック1用：上段
        track1_coef_font_size2 = TRACK1_TATE_LOWER_COEF
        track2_coef_font_size1 = TRACK2_TATE_1ST_COEF    'トラック2用：1段目
        track2_coef_font_size2 = TRACK2_TATE_2ND_COEF    'トラック2用：2段目
        track2_coef_font_size3 = TRACK2_TATE_3RD_COEF    'トラック2用：3段目
    '動画が横の時
    Else
        track1_coef_font_size1 = TRACK1_YOKO_UPPER_COEF  '動画が縦のときのトラック1用：上段
        track1_coef_font_size2 = TRACK1_YOKO_LOWER_COEF
        track2_coef_font_size1 = TRACK2_YOKO_1ST_COEF    'トラック2用：1段目
        track2_coef_font_size2 = TRACK2_YOKO_2ND_COEF    'トラック2用：2段目
        track2_coef_font_size3 = TRACK2_YOKO_3RD_COEF    'トラック2用：3段目
    End If

    'フォントサイズを設定
    track1_font_size1 = video_width / track1_coef_font_size1 '動画の縦or横によって分母を変更することで、文字サイズが変わる
    track1_font_size2 = video_width / track1_coef_font_size2
    track2_font_size1 = video_width / track2_coef_font_size1
    track2_font_size2 = video_width / track2_coef_font_size2
    track2_font_size3 = video_width / track2_coef_font_size3

    '各姿勢の名前と条件の読み出し
    'MinとMaxが直感的でないので注意
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
        CaptionName2Koshimage = .Cells(210, 2)
        CaptionName2Hizamage = .Cells(228, 2)
    End With

    '評価除外用
    CaptionName2(0) = "0-姿勢評価なし" '下段のキャプション名を表示しない

    Track1FileName = ActiveWorkbook.Path & "\" & movieName & ".srt"
    Track2FileName = ActiveWorkbook.Path & "\" & movieName & CAPTION_TRACK2_FILE_NAME_SOEJI & ".srt"

    'フレームレートの読み出し
    fps = getFps()

    '処理する行数を取得（3列目の最終セル）
    max_row_num = getLastRow()

    max_array_num = max_row_num - 1 - 1 '2行目からセルに値が入るため-1、配列は0から使うため-1

    '配列数が決まったため配列を再定義
    ReDim WorkName(max_array_num, 0)

    With ThisWorkbook.Sheets("姿勢重量点調査票")

        '工程評価シートから作業手順と作業名を読み取って配列に書き込む
        '処理する追加行数を取得する
        '"要素数"のセル位置の移動量を調べる  ※最大999行(<1050)にする
        expand_no = 0
        Do While ThisWorkbook.Worksheets("姿勢重量点調査票").Cells(107 + expand_no, 13) <> "要素数" And expand_no < 1050
            expand_no = expand_no + 1
        Loop

        '時間を初期値に設定
        separate_work_time = 0
        t0 = 0
        '動画先頭を除外したときに評価のスタートが0.0秒ではなくなるため変更
        t = .Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME + 1).Value

        '最初の作業（除外後の開始時刻）をシートから読み取り、t0 = t = 実際の開始秒 で初期化する。
        t = s_ProcessEvaluation_2nd.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME + 1).Value
        t0 = t

        'ポイント計算シートのフラグをカウントして、各作業姿勢の時間を計算する
        For i = 0 To 89 + expand_no

            '作業開始時間が空なら作業名は入力しない
            If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME)) Then

            '作業開始時間が入力されているなら配列に作業名を入力する
            Else
                separate_work_time = .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME + 1).Value
                t0 = t
                t = separate_work_time '作業時間を単一で入力する場合
                '秒数からフレーム数へ変換
                start_frame = t0 * fps
                end_frame = t * fps - 1

                '作業終了時間はラウンドアップ関数を使用しているため、はみだし防止
                If end_frame > max_array_num Then
                    end_frame = max_array_num
                End If

                '2セット目以降で、前回のend_frameと今回のstart_frameが重なるのを防止する
                If start_frame > 0 Then
                    start_frame = start_frame + 1
                End If

                'start_frameフレーム(t0秒) から end_frameフレーム(t秒) までの処理
                If start_frame < end_frame Then
                    For j = start_frame To end_frame
                        WorkName(j, 0) = _
                        .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_NUMBER).Text & _
                        "." & _
                        .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_NAME).Text
                    Next
                End If
            End If
        Next

    End With

    With ThisWorkbook.Sheets("ポイント計算シート")

        'ファイルを開く
        Open Track1FileName For Output As #1

        '処理する行数を取得（3列目の最終セル）
        max_row_num = getLastRow()

        'ファイル出力
        For i = 2 To max_row_num

            'ポイント計算シートのキャプション列より、姿勢重量点調査票の作業名を先に読み取っておく
            CaptionName0 = WorkName(i - 2, 0)

            'データ区間の描画色、キャプション名を設定する
            '※はじめに評価除外、データ強制区間、データ不良区間の順に判定する（重複ビットON時、字幕表示の優先度が高い順）
            ' 今のところ最後のfillDataで測定ビットは同時に立つ仕様。
            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
                CaptionName1 = CAPTION_DATA_REMOVE_SECTION
                ColorName1 = COLOR_DATA_REMOVE_SECTION
            ElseIf .Cells(i, COLUMN_DATA_FORCED_SECTION).Value > 0 Then
                CaptionName1 = CAPTION_DATA_FORCED_SECTION
                ColorName1 = COLOR_DATA_FORCED_SECTION
            ElseIf .Cells(i, COLUMN_DATA_MISSING_SECTION).Value > 0 Then
                CaptionName1 = CAPTION_DATA_MISSING_SECTION
                ColorName1 = COLOR_DATA_MISSING_SECTION
            ElseIf .Cells(i, COLUMN_DATA_MEASURE_SECTION).Value > 0 Then
                CaptionName1 = CAPTION_DATA_MEASURE_SECTION
                ColorName1 = COLOR_DATA_MEASURE_SECTION
            ElseIf .Cells(i, COLUMN_DATA_PREDICT_SECTION).Value > 0 Then
                CaptionName1 = CAPTION_DATA_PREDICT_SECTION
                ColorName1 = COLOR_DATA_PREDICT_SECTION
            End If

            '姿勢素点の描画色、キャプション名を設定する
            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
                '評価除外のとき
                CaptionNo2 = 0
                ColorName2 = COLOR_DATA_REMOVE_SECTION
            Else
                '通常時
                CaptionNo2 = .Cells(i, COLUMN_DATA_RESULT_FIX).Value
                If .Cells(i, COLUMN_DATA_RESULT_GREEN).Value > 0 Then
                    ColorName2 = COLOR_DATA_RESULT_GREEN
                ElseIf .Cells(i, COLUMN_DATA_RESULT_YELLOW).Value > 0 Then
                    ColorName2 = COLOR_DATA_RESULT_YELLOW
                ElseIf .Cells(i, COLUMN_DATA_RESULT_RED).Value > 0 Then
                    ColorName2 = COLOR_DATA_RESULT_RED
                End If
            End If

            Track1OutputString1 = _
                "<font size=""" & track1_font_size1 & """ color =" & "#ffffff" & ">" & CaptionName0 & "</font>" & _
                "<font size=""" & track1_font_size1 & """ color =" & ColorName1 & ">" & CaptionName1 & "</font>"

            Track1OutputString2 = _
                "<font size=""" & track1_font_size2 & """ color =" & ColorName2 & ">" & CaptionName2(CaptionNo2) & "</font>"

            '字幕文字列をテキストファイルに書き出しする
            Print #1, " " & i - 1 '数字の両側に半角スペースを入れる。字幕トラック2と区別するため
            Print #1, .Cells(i, COLUMN_ROUGH_TIME).Value&; " --> " & .Cells(i + 1, COLUMN_ROUGH_TIME).Value '時刻を出力

            Print #1, Replace(Track1OutputString1, vbLf, vbCrLf) '改行コードを置き換え、キャプション出力
            Print #1, Replace(Track1OutputString2, vbLf, vbCrLf) '改行コードを置き換え、キャプション出力

            Print #1, ""
            Print #1, ""

            'ポイント計算シートの字幕文字列 作業No. - 作業名をクリア
            .Cells(i, COLUMN_CAPTION_WORK_NAME).clear


            'デバッグ時、判定されない条件が分かるように色名をリセットしておく
            ColorName1 = "#ffffff"
            ColorName2 = "#ffffff"

        Next

        'ファイルを閉じる
        Close #1
        Close #2

    End With

    '表示・更新をオンに戻す
    Call restartUpdate

End Function


'帳票更新ボタンが押された時の処理
' 引数  ：なし
' 戻り値：なし
Function ClickUpdateDataCore()
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

    dotPoint = InStrRev(ActiveWorkbook.Name, ".")
    workbookName = Left(ActiveWorkbook.Name, dotPoint - 1)

    Call outputCaption(workbookName)
    Debug.Print " 更新時間" & Format$(Timer - tstart_click, "0.00") & " sec."

End Function


'帳票更新ボタンが押された時の処理
' 引数１：なし
' 戻り値：なし
Sub ClickUpdateData()
    Call ClickUpdateDataCore
End Sub


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


Sub test() '230207
    Dim dotPoint     As String
    Dim workbookName As String

    dotPoint = InStrRev(ActiveWorkbook.Name, ".")
    workbookName = Left(ActiveWorkbook.Name, dotPoint - 1)
    Call MacroUpdateData(workbookName, ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 199))
End Sub


' 引数1 ：なし
' 引数2 ：なし
' 戻り値：なし
Sub VeryHiddenSheet()
    Sheets("ポイント計算シート").Visible = xlVeryHidden
    Sheets("条件設定シート").Visible = xlVeryHidden
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



'姿勢重量点調査票の選択と保存
' 引数1 ：動画名
' 戻り値：なし
Sub MacroSaveData(movieName As String)
    ThisWorkbook.Save
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