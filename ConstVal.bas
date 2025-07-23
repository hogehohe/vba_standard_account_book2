Attribute VB_Name = "ConstVal"
Option Explicit

'帳票1
#Const DEFAULT_1st = 1
'帳票2
#Const DEFAULT_2nd = 2
'トヨタ車体
#Const TMC_CAR_BODY = 3
'バリアント設定
#Const VALIANT = DEFAULT_2nd

'---------------------------------------------
'   ボタンナンバー
'---------------------------------------------
Public Const BUTSEL_REMOVS As Long = -1
Public Const BUTSEL_POSTURE_OFF As Long = 0
Public Const BUTSEL_POSTURE_ON As Long = 1
Public Const BUTSEL_POSTUR_NON As Long = 2
Public Const BUTSEL_EXCLUSION As Long = 99

'---------------------------------------------
'姿勢素点修正シートで使う定数
'---------------------------------------------
'１マスの秒数を定義
Const UNIT_TIME       As Double = 0.1
'０秒の列
Const COLUMN_ZERO_NUM As Long = 6

'行
'信頼性上端
Const ROW_RELIABILITY_TOP      As Long = 2

'信頼性下端
Const ROW_RELIABILITY_BOTTOM   As Long = 7

'姿勢点上端
#If VALIANT = DEFAULT_2nd Then
    '姿勢点上端
    Const ROW_POSTURE_SCORE_TOP    As Long = 12 + 2
    '姿勢点下端
    Const ROW_POSTURE_SCORE_BOTTOM As Long = 21 + 2
#ElseIf VALIANT = TMC_CAR_BODY Then
    Const ROW_POSTURE_SCORE_TOP    As Long = 9
    '姿勢点上端
    '姿勢点下端
    Const ROW_POSTURE_SCORE_BOTTOM As Long = 17
#End If

'拳上_姿勢点
Const ROW_POSTURE_SCORE_KOBUSHIAGE      As Long = 10 + 2 '一旦姿勢素点の下側に表示する

'=== トヨタ車体特有 ===
'A_姿勢点
Const ROW_POSTURE_SCORE_A      As Long = 12
'B_姿勢点
Const ROW_POSTURE_SCORE_B      As Long = 14
'C_姿勢点
Const ROW_POSTURE_SCORE_C      As Long = 16
'D_姿勢点
Const ROW_POSTURE_SCORE_D      As Long = 18
'E_姿勢点
Const ROW_POSTURE_SCORE_E      As Long = 20

'---------------------------------------------
'ポイント計算シートの列
'---------------------------------------------
'姿勢点が保存されている列 2023/12/12 育成G追記
Const COLUMN_POSTURE_SCORE_ALL As Long = 203

'2023/12/11 育成G小杉追記 条件A(拳上)が保存されている列
Const COLUMN_POSTURE_SCORE_KOBUSHIAGE As Long = 245

'=== トヨタ車体特有 ===
'条件Aが保存されている列
Const COLUMN_POSTURE_SCORE_A As Long = 245
'条件Aが保存されている列
Const COLUMN_POSTURE_SCORE_B As Long = 247
'条件Aが保存されている列
Const COLUMN_POSTURE_SCORE_C As Long = 249
'条件Aが保存されている列
Const COLUMN_POSTURE_SCORE_D As Long = 251
'条件Aが保存されている列
Const COLUMN_POSTURE_SCORE_E As Long = 253


'信頼性が保存されている列
'測定
Const COLUMN_MEASURE_SECTION    As Long = 204
'推定
Const COLUMN_PREDICT_SECTION    As Long = 205
'除外区間
Const COLUMN_REMOVE_SECTION     As Long = 206
'強制区間
Const COLUMN_FORCED_SECTION     As Long = 207
'強制区間 2023/12/12 育成G追記
Const COLUMN_FORCED_SECTION_TOTAL    As Long = 207
'元データ
Const COLUMN_DATA_RESULT_FIX    As Long = 208
'元データ 2023/12/12 育成G追記
Const COLUMN_BASE_SCORE        As Long = 208
'姿勢素点緑色
Const COLUMN_POSTURE_GREEN      As Long = 209
'姿勢素点黄色
Const COLUMN_POSTURE_YELLOW     As Long = 210
'姿勢素点赤色
Const COLUMN_POSTURE_RED        As Long = 211
'欠損
Const COLUMN_MISSING_SECTION    As Long = 219
'拳上強制区間 2023/12/12 育成G追記
Const COLUMN_FORCED_SECTION_KOBUSHIAGE As Long = 223

'---------------------------------------------
'姿勢素点修正シート　関連
'---------------------------------------------
'時刻表示セルの幅
Const TIME_WIDTH               As Long = 30
'時刻表示セルが存在する行

#If VALIANT = DEFAULT_2nd Then
    '姿勢点下端
    '2023/12/19育成G追記（レイアウト変更により2行分追加）
    Const TIME_ROW  As Long = 25 + 2
    'データ調整用のテーブルの下端
    '2023/12/19育成G追記（レイアウト変更により2行分追加）
    Const BOTTOM_OF_TABLE   As Long = 26 + 2
#ElseIf VALIANT = TMC_CAR_BODY Then
    '姿勢点下端
    Const TIME_ROW  As Long = 24
    'データ調整用のテーブルの下端
    Const BOTTOM_OF_TABLE   As Long = 24
#End If

'一つ目の時刻表示セルの左端
Const TIME_COLUMN_LEFT         As Long = 22
'一つ目の時刻表示セルの右端
Const TIME_COLUMN_RIGHT        As Long = 51

'======================================================================================
'条件設定シートの各データの行番号、列番号を定義 (拳上概要の定数もここで定義）
'======================================================================================
Const KOBUSHIAGE_MISSING_DOWNLIM_TIME       As Double = 1     '（秒） 拳上欠損ノイズ判定に使う
Const TEKUBI_SPEED_UPLIM_PREDICT            As Double = 10    '（km/h）手首z位置の変化量上限　遮蔽検知に使う
Const MEAGERE_TIME_MACROUPDATEDATA          As Boolean = True 'TrueのときMacroUpdateDataの処理時間を測定する
Const KOBUSHIAGE_TIME_HOSEI_COEF_WORK       As Double = 5 / 355 '拳上時間補正係数 対象工程の中で最も作業時間が長い工程の　確認必要な欠損区間数/作業時間
Const KOBUSHIAGE_MISSING_DILATION_SIZE      As Double = 0.33   '（秒）拳上欠損の膨張処理に使う窓の大きさ（片側）
Const KOBUSHIAGE_MISSING_EROSION_SIZE       As Double = 0.33   '（秒）拳上欠損の収縮処理に使う窓の大きさ（片側）
Const KOBUSHIAGE_TIME_HOSEI_COEF_MISSING    As Double = 0.2     '拳上時間補正係数 確認必要な欠損区間数１個あたり

'makeGraph、outputCaption、fixGraphDataAndSheetモジュールの中に条件設定シートのセル内から値を読み出す部分あり

'======================================================================================
'ポイント計算シート上の各データの行番号、列番号を定義
'======================================================================================
Const COLUMN_POSE_NAME                      As Long = 1
Const COLUMN_POSE_KEEP_TIME                 As Long = 2
Const COLUMN_HIZA_R_ANGLE                   As Long = 6
Const COLUMN_HIZA_L_ANGLE                   As Long = 7
Const COLUMN_KOSHI_ANGLE                    As Long = 8
Const COLUMN_SHOOTING_DIRECTION             As Long = 9

Const COLUMN_POS_KOSHI_Z                    As Long = 13

Const COLUMN_POS_AHIKUBI_R_Z                As Long = 25
Const COLUMN_POS_AHIKUBI_L_Z                As Long = 37

Const COLUMN_POS_KATA_R_Z                   As Long = 57
Const COLUMN_POS_KATA_L_Z                   As Long = 69

Const COLUMN_POS_HIJI_R_Z                   As Long = 61
Const COLUMN_POS_HIJI_L_Z                   As Long = 73

Const COLUMN_POS_TEKUBI_R_Z                 As Long = 65
Const COLUMN_POS_TEKUBI_L_Z                 As Long = 77

Const COLUMN_ROUGH_TIME                     As Long = 201
Const COLUMN_CAPTION_WORK_NAME              As Long = 202
Const COLUMN_DATA_RESULT_ORIGIN             As Long = 203
Const COLUMN_DATA_MEASURE_SECTION           As Long = 204
Const COLUMN_DATA_PREDICT_SECTION           As Long = 205
Const COLUMN_DATA_REMOVE_SECTION            As Long = 206
Const COLUMN_DATA_FORCED_SECTION            As Long = 207

Const COLUMN_DATA_RESULT_GREEN              As Long = 209
Const COLUMN_DATA_RESULT_YELLOW             As Long = 210
Const COLUMN_DATA_RESULT_RED                As Long = 211
Const COLUMN_CAPTION_TRACK1                 As Long = 212 'デバッグ用（普段は使わない）

'=== トヨタ車体特有 ===
Const COLUMN_DATA_RESULT_ALL    As Long = 203
Const COLUMN_UDEAGE_RESULT      As Long = 224

Const COLUMN_DATA_MISSING_SECTION           As Long = 219

Const COLUMN_DATA_KOBUSHIAGE_MEASURE_SECTION_ORG    As Long = 221
Const COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_ORG    As Long = 222
Const COLUMN_KOBUSHIAGE_FORCED_SECTION              As Long = 223 '拳上、腰曲げ、膝曲げの強制、判定フラグ、フラグの記憶
Const COLUMN_KOBUSHIAGE_RESULT                      As Long = 245
Const COLUMN_DATA_KOSHIMAGE_MEASURE_SECTION         As Long = 225
Const COLUMN_DATA_KOSHIMAGE_PREDICT_SECTION         As Long = 226
Const COLUMN_DATA_KOSHIMAGE_MISSING_SECTION         As Long = 227
Const COLUMN_KOSHIMAGE_FORCED_SECTION               As Long = 228

#If VALIANT = DEFAULT_2nd Then
    Const COLUMN_KOSHIMAGE_RESULT                   As Long = 247
#ElseIf VALIANT = TMC_CAR_BODY Then
    Const COLUMN_KOSHIMAGE_RESULT                   As Long = 229
#End If

Const COLUMN_DATA_HIZAMAGE_MEASURE_SECTION             As Long = 230
Const COLUMN_DATA_HIZAMAGE_PREDICT_SECTION             As Long = 231
Const COLUMN_DATA_HIZAMAGE_MISSING_SECTION             As Long = 232
Const COLUMN_HIZAMAGE_FORCED_SECTION                   As Long = 233

#If VALIANT = DEFAULT_2nd Then
    Const COLUMN_HIZAMAGE_RESULT                        As Long = 249
#ElseIf VALIANT = TMC_CAR_BODY Then
    Const COLUMN_HIZAMAGE_RESULT                        As Long = 234
#End If

Const COLUMN_CAPTION_TRACK2                            As Long = 235 'デバッグ用（普段は使わない）

Const COLUMN_TEKUBI_RZ_SPEED                           As Long = 237 '右手首Ｚ位置の差
Const COLUMN_TEKUBI_LZ_SPEED                           As Long = 238 '左手首Ｚ位置の差
Const COLUMN_TEKUBI_Z_SPEED_OVER                       As Long = 239 '手首Ｚ位置の差 しきい値超えフラグ
Const COLUMN_DATA_KOBUSHIAGE_MEASURE_SECTION_DST           As Long = 240 '拳上測定区間
Const COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_DST           As Long = 241 '拳上欠損区間
Const COLUMN_MEAGERE_TIME_MACROUPDATEDATA              As Long = 242 'MacroUpdateDataの処理時間を測定結果を格納する

Const COLUMN_DATA_RESULT_GH_KOBUSHIAGE      As Long = 245
Const COLUMN_DATA_RESULT_GH_KOSHIMAGE       As Long = 247
Const COLUMN_DATA_RESULT_GH_HIZAMAGE        As Long = 249
Const COLUMN_DATA_RESULT_GH_SONKYO          As Long = 251

Const COLUMN_GH_HIZA_L                      As Long = 252
Const COLUMN_GH_HIZA_R                      As Long = 253

Const COLUMN_MAX_NUMBER                                As Long = 256 '現在使用されている列番号の最大値

'======================================================================================
'(トヨタ車体特有)工程評価シートの各データの行番号、列番号を定義
'======================================================================================

Const TB_HYOUKA_SHEET_ROW_POSESTART                    As Long = 16
Const TB_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK          As Long = 107

Const TB_HYOUKA_SHEET_COLUMN_WORK_NUMBER               As Long = 3
Const TB_HYOUKA_SHEET_COLUMN_WORK_NAME                 As Long = 4
Const TB_HYOUKA_SHEET_COLUMN_WORKSTART_TIME            As Long = 10
Const TB_HYOUKA_SHEET_COLUMN_WORKEND_TIME              As Long = 12
Const TB_HYOUKA_SHEET_COLUMN_EXPAND_NUMBER_CHECK       As Long = 15
Const TB_HYOUKA_SHEET_COLUMN_WORK_TIME                 As Long = 17

Const TB_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME           As Long = 20
Const TB_HYOUKA_SHEET_COLUMN_ZENKUTSU_TIME             As Long = 22
Const TB_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME             As Long = 24
Const TB_HYOUKA_SHEET_COLUMN_SONKYO_TIME               As Long = 26
'こぜわしさに関する定数
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_CHECK         As Long = 38
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_WORKTYPE      As Long = 42
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_DIAMETER      As Long = 43
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_LENGTH        As Long = 45
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_COMBINATION   As Long = 47
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_DIRECTION     As Long = 49
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_NUMBER        As Long = 51

'見にくさに関する定数
Const TB_HYOUKA_SHEET_COLUMN_MINIKUSA_CHECK            As Long = 52
Const TB_HYOUKA_SHEET_COLUMN_MINIKUSA_SIZE             As Long = 56
Const TB_HYOUKA_SHEET_COLUMN_MINIKUSA_DISTANCE         As Long = 58
Const TB_HYOUKA_SHEET_COLUMN_MINIKUSA_CONTRAST         As Long = 60
Const TB_HYOUKA_SHEET_COLUMN_MINIKUSA_DIRECTION        As Long = 62

'======================================================================================
'姿勢重量点調査票シートの各データの行番号、列番号を定義
'======================================================================================
Const SHIJUTEN_SHEET_ROW_KOUTEI_NAME                            As Long = 3
Const SHIJUTEN_SHEET_ROW_POSESTART_INDEX                        As Long = 9
Const SHIJUTEN_SHEET_ROW_EXPAND_NUMBER_CHECK                    As Long = 29

Const SHIJUTEN_SHEET_EXPAND_NUM_CHECK_WORD                      As String = "その他の時間（定時稼働時間7.5H-Σ延べ時間）"


Const SHIJUTEN_SHEET_COLUMN_WORK_NUMBER                         As Long = 2
Const SHIJUTEN_SHEET_COLUMN_WORK_NAME                           As Long = 3
Const SHIJUTEN_SHEET_COLUMN_KOUTEI_NAME                         As Long = 4
Const SHIJUTEN_SHEET_COLUMN_WORK_TIME                           As Long = 9
Const SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX                    As Long = 10

Const SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME                      As Long = 36

#If VALIANT = DEFAULT_2nd Then
    Const SHIJUTEN_SHEET_COLUMN_WORKEND_TIME                        As Long = 38
    Const SHIJUTEN_SHEET_COLUMN_DATA_MISSING_SECTION                As Long = 46
    Const SHIJUTEN_SHEET_COLUMN_DATA_PREDICT_SECTION                As Long = 47
    Const SHIJUTEN_SHEET_COLUMN_WORKEND_SECOND                      As Long = 37
    Const SHIJUTEN_SHEET_COLUMN_DATA_REMOVE_OR_FORCED               As Long = 38
    Const SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_TIME                     As Long = 49 '拳上時間
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME                      As Long = 51 '腰曲げ時間
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME                       As Long = 53 '膝曲げ時間
    Const SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_MISSING_TIME             As Long = 55 '拳上欠損区間
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_MISSING_TIME              As Long = 57 '腰曲げ欠損区間
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_PREDICT_TIME              As Long = 58 '腰曲げ推定区間
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_MISSING_TIME               As Long = 60 '膝曲げ欠損区間
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_PREDICT_TIME               As Long = 61 '膝曲げ推定区間

#ElseIf VALIANT = TMC_CAR_BODY Then
    Const SHIJUTEN_SHEET_COLUMN_WORKEND_TIME                        As Long = 36
    Const SHIJUTEN_SHEET_COLUMN_DATA_MISSING_SECTION                As Long = 45
    Const SHIJUTEN_SHEET_COLUMN_DATA_PREDICT_SECTION                As Long = 46
    Const SHIJUTEN_SHEET_COLUMN_UDEAGE_TIME                         As Long = 48 '腕上げ時間
    Const SHIJUTEN_SHEET_COLUMN_UDEAGE_REMOVE_OR_FORCED             As Long = 49 '腕上げ評価強制
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME                      As Long = 51 '腰曲げ時間
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_REMOVE_OR_FORCED          As Long = 52 '腰曲げ評価強制
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME                       As Long = 54 '膝曲げ時間
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_REMOVE_OR_FORCED           As Long = 55 '膝曲げ評価強制
    Const SHIJUTEN_SHEET_COLUMN_UDEAGE_MISSING_TIME                 As Long = 57 '腕上げ欠損区間
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_MISSING_TIME              As Long = 59 '腰曲げ欠損区間
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_PREDICT_TIME              As Long = 60 '腰曲げ推定区間
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_MISSING_TIME               As Long = 62 '膝曲げ欠損区間
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_PREDICT_TIME               As Long = 63 '膝曲げ推定区間
#End If



'======================================================================================
'工程評価シートの各データの行番号、列番号を定義
'======================================================================================

Const GH_HYOUKA_SHEET_ROW_POSESTART                    As Long = 15
Const GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK          As Long = 115

Const GH_HYOUKA_SHEET_EXPAND_NUM_CHECK_WORD            As String = "合計"

Const GH_HYOUKA_SHEET_COLUMN_WORK_NUMBER               As Long = 2
Const GH_HYOUKA_SHEET_COLUMN_WORK_NAME                 As Long = 3
Const GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME            As Long = 12
Const GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME              As Long = 14
Const GH_HYOUKA_SHEET_COLUMN_WORK_TIME                 As Long = 16
Const GH_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME           As Long = 17
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
'字幕情報の定義
'======================================================================================
Const CAPTION_TRACK2_FILE_NAME_SOEJI           As String = "2" '字幕トラック２用のファイル名末尾につける添字
'各種字幕のフォントサイズ係数
'分母の値のため、値が小さいほど文字は大きい


#If VALIANT = DEFAULT_2nd Then

    '動画が縦の時
    Const TRACK1_TATE_UPPER_COEF                   As Long = 22 'トラック1用：上段
    Const TRACK1_TATE_LOWER_COEF                   As Long = 11 'トラック1用：下段
    Const TRACK2_TATE_1ST_COEF                     As Long = 22 'トラック2用：１段目
    Const TRACK2_TATE_2ND_COEF                     As Long = 22 'トラック2用：２段目
    Const TRACK2_TATE_3RD_COEF                     As Long = 13 'トラック2用：３段目
    
    '動画が横の時
    Const TRACK1_YOKO_UPPER_COEF                   As Long = 30 'トラック1用：上段
    Const TRACK1_YOKO_LOWER_COEF                   As Long = 15 'トラック1用：下段
    Const TRACK2_YOKO_1ST_COEF                     As Long = 30 'トラック2用：１段目
    Const TRACK2_YOKO_2ND_COEF                     As Long = 30 'トラック2用：２段目
    Const TRACK2_YOKO_3RD_COEF                     As Long = 18 'トラック2用：３段目

#ElseIf VALIANT = TMC_CAR_BODY Then
'動画が横の時

    Const TRACK2_TATE_1ST_COEF                     As Long = 20 'トラック2用：１段目
    Const TRACK2_TATE_2ND_COEF                     As Long = 5 'トラック2用：２段目
    Const TRACK2_TATE_3RD_COEF                     As Long = 16 'トラック2用：３段目
    Const TRACK2_TATE_4RD_COEF                     As Long = 20 'トラック2用：４段目

    '動画が縦の時
    Const TRACK2_YOKO_1ST_COEF                     As Long = 22 'トラック2用：１段目
    Const TRACK2_YOKO_2ND_COEF                     As Long = 5 'トラック2用：２段目
    Const TRACK2_YOKO_3RD_COEF                     As Long = 15 'トラック2用：３段目
    Const TRACK2_YOKO_4RD_COEF                     As Long = 20 'トラック2用：４段目

#End If


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


#If VALIANT = DEFAULT_2nd Then

    '外販用の字幕文字列（字幕トラック2用 3段目に表示）
    Const CAPTION_A_RESULT_NAME1  As String = "　　　　拳上"
    Const CAPTION_B_RESULT_NAME1  As String = "  　　腰曲げ　 　"
    Const CAPTION_C_RESULT_NAME1  As String = "膝曲げ"
    
    '外販用の条件字幕文字列（字幕トラック2用 4段目に表示）
    Const CAPTION_A_RESULT_NAME2  As String = "手首が肩より上"
    Const CAPTION_B_RESULT_NAME2  As String = "30°以上"
    Const CAPTION_C_RESULT_NAME2  As String = "60°以上"

#ElseIf VALIANT = TMC_CAR_BODY Then
'動画が横の時

    'トヨタ車体用の字幕文字列（字幕トラック2用 3段目に表示）
    Const CAPTION_A_RESULT_NAME  As String = "　　拳上 　"
    Const CAPTION_B_RESULT_NAME  As String = "  前屈　　"
    Const CAPTION_C_RESULT_NAME  As String = "膝曲げ　　"
    Const CAPTION_D_RESULT_NAME  As String = "  蹲踞　"

    'トヨタ車体用の条件字幕文字列（字幕トラック2用 4段目に表示）
    Const CAPTION_A_RESULT_NAME2  As String = "手首が肩より上"
    Const CAPTION_B_RESULT_NAME2  As String = "45°以上"
    Const CAPTION_C_RESULT_NAME2  As String = "90°以上"
    Const CAPTION_D_RESULT_NAME2  As String = "90°以上"

#End If

'キャプションノイズ除去の閾値
Const CAPTION_REMOVE_NOISE_SECOND              As Double = 0.1 'キャプションノイズを除去する長さ(秒) （～未満なら除去）

'姿勢素点の値によって、緑／黄／赤を分ける際の境界条件
Const DATA_SEPARATION_GREEN_BOTTOM             As Long = 1
Const DATA_SEPARATION_GREEN_TOP                As Long = 2
Const DATA_SEPARATION_YELLOW_BOTTOM            As Long = 3
Const DATA_SEPARATION_YELLOW_TOP               As Long = 5
Const DATA_SEPARATION_RED_BOTTOM               As Long = 6
Const DATA_SEPARATION_RED_TOP                  As Long = 10

'======================================================================================
'DataAdjustingSheet用
'======================================================================================
'debug
'Const LIMIT_COLUMN           As Long = 800
Const LIMIT_COLUMN           As Long = 16200

'---------------------------------------------
'姿勢素点修正シート　関連
'---------------------------------------------
'LIMIT_COLUMNの設定値は3の倍数とする必要がある
'30fps×60秒×9分＝16200
Const SHEET_LIMIT_COLUMN       As Long = LIMIT_COLUMN + COLUMN_ZERO_NUM

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

