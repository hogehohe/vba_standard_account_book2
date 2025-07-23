Attribute VB_Name = "ConstVal"
Option Explicit

'���[1
#Const DEFAULT_1st = 1
'���[2
#Const DEFAULT_2nd = 2
'�g���^�ԑ�
#Const TMC_CAR_BODY = 3
'�o���A���g�ݒ�
#Const VALIANT = DEFAULT_2nd

'---------------------------------------------
'   �{�^���i���o�[
'---------------------------------------------
Public Const BUTSEL_REMOVS As Long = -1
Public Const BUTSEL_POSTURE_OFF As Long = 0
Public Const BUTSEL_POSTURE_ON As Long = 1
Public Const BUTSEL_POSTUR_NON As Long = 2
Public Const BUTSEL_EXCLUSION As Long = 99

'---------------------------------------------
'�p���f�_�C���V�[�g�Ŏg���萔
'---------------------------------------------
'�P�}�X�̕b�����`
Const UNIT_TIME       As Double = 0.1
'�O�b�̗�
Const COLUMN_ZERO_NUM As Long = 6

'�s
'�M������[
Const ROW_RELIABILITY_TOP      As Long = 2

'�M�������[
Const ROW_RELIABILITY_BOTTOM   As Long = 7

'�p���_��[
#If VALIANT = DEFAULT_2nd Then
    '�p���_��[
    Const ROW_POSTURE_SCORE_TOP    As Long = 12 + 2
    '�p���_���[
    Const ROW_POSTURE_SCORE_BOTTOM As Long = 21 + 2
#ElseIf VALIANT = TMC_CAR_BODY Then
    Const ROW_POSTURE_SCORE_TOP    As Long = 9
    '�p���_��[
    '�p���_���[
    Const ROW_POSTURE_SCORE_BOTTOM As Long = 17
#End If

'����_�p���_
Const ROW_POSTURE_SCORE_KOBUSHIAGE      As Long = 10 + 2 '��U�p���f�_�̉����ɕ\������

'=== �g���^�ԑ̓��L ===
'A_�p���_
Const ROW_POSTURE_SCORE_A      As Long = 12
'B_�p���_
Const ROW_POSTURE_SCORE_B      As Long = 14
'C_�p���_
Const ROW_POSTURE_SCORE_C      As Long = 16
'D_�p���_
Const ROW_POSTURE_SCORE_D      As Long = 18
'E_�p���_
Const ROW_POSTURE_SCORE_E      As Long = 20

'---------------------------------------------
'�|�C���g�v�Z�V�[�g�̗�
'---------------------------------------------
'�p���_���ۑ�����Ă���� 2023/12/12 �琬G�ǋL
Const COLUMN_POSTURE_SCORE_ALL As Long = 203

'2023/12/11 �琬G�����ǋL ����A(����)���ۑ�����Ă����
Const COLUMN_POSTURE_SCORE_KOBUSHIAGE As Long = 245

'=== �g���^�ԑ̓��L ===
'����A���ۑ�����Ă����
Const COLUMN_POSTURE_SCORE_A As Long = 245
'����A���ۑ�����Ă����
Const COLUMN_POSTURE_SCORE_B As Long = 247
'����A���ۑ�����Ă����
Const COLUMN_POSTURE_SCORE_C As Long = 249
'����A���ۑ�����Ă����
Const COLUMN_POSTURE_SCORE_D As Long = 251
'����A���ۑ�����Ă����
Const COLUMN_POSTURE_SCORE_E As Long = 253


'�M�������ۑ�����Ă����
'����
Const COLUMN_MEASURE_SECTION    As Long = 204
'����
Const COLUMN_PREDICT_SECTION    As Long = 205
'���O���
Const COLUMN_REMOVE_SECTION     As Long = 206
'�������
Const COLUMN_FORCED_SECTION     As Long = 207
'������� 2023/12/12 �琬G�ǋL
Const COLUMN_FORCED_SECTION_TOTAL    As Long = 207
'���f�[�^
Const COLUMN_DATA_RESULT_FIX    As Long = 208
'���f�[�^ 2023/12/12 �琬G�ǋL
Const COLUMN_BASE_SCORE        As Long = 208
'�p���f�_�ΐF
Const COLUMN_POSTURE_GREEN      As Long = 209
'�p���f�_���F
Const COLUMN_POSTURE_YELLOW     As Long = 210
'�p���f�_�ԐF
Const COLUMN_POSTURE_RED        As Long = 211
'����
Const COLUMN_MISSING_SECTION    As Long = 219
'���㋭����� 2023/12/12 �琬G�ǋL
Const COLUMN_FORCED_SECTION_KOBUSHIAGE As Long = 223

'---------------------------------------------
'�p���f�_�C���V�[�g�@�֘A
'---------------------------------------------
'�����\���Z���̕�
Const TIME_WIDTH               As Long = 30
'�����\���Z�������݂���s

#If VALIANT = DEFAULT_2nd Then
    '�p���_���[
    '2023/12/19�琬G�ǋL�i���C�A�E�g�ύX�ɂ��2�s���ǉ��j
    Const TIME_ROW  As Long = 25 + 2
    '�f�[�^�����p�̃e�[�u���̉��[
    '2023/12/19�琬G�ǋL�i���C�A�E�g�ύX�ɂ��2�s���ǉ��j
    Const BOTTOM_OF_TABLE   As Long = 26 + 2
#ElseIf VALIANT = TMC_CAR_BODY Then
    '�p���_���[
    Const TIME_ROW  As Long = 24
    '�f�[�^�����p�̃e�[�u���̉��[
    Const BOTTOM_OF_TABLE   As Long = 24
#End If

'��ڂ̎����\���Z���̍��[
Const TIME_COLUMN_LEFT         As Long = 22
'��ڂ̎����\���Z���̉E�[
Const TIME_COLUMN_RIGHT        As Long = 51

'======================================================================================
'�����ݒ�V�[�g�̊e�f�[�^�̍s�ԍ��A��ԍ����` (����T�v�̒萔�������Œ�`�j
'======================================================================================
Const KOBUSHIAGE_MISSING_DOWNLIM_TIME       As Double = 1     '�i�b�j ���㌇���m�C�Y����Ɏg��
Const TEKUBI_SPEED_UPLIM_PREDICT            As Double = 10    '�ikm/h�j���z�ʒu�̕ω��ʏ���@�Օ����m�Ɏg��
Const MEAGERE_TIME_MACROUPDATEDATA          As Boolean = True 'True�̂Ƃ�MacroUpdateData�̏������Ԃ𑪒肷��
Const KOBUSHIAGE_TIME_HOSEI_COEF_WORK       As Double = 5 / 355 '���㎞�ԕ␳�W�� �ΏۍH���̒��ōł���Ǝ��Ԃ������H���́@�m�F�K�v�Ȍ�����Ԑ�/��Ǝ���
Const KOBUSHIAGE_MISSING_DILATION_SIZE      As Double = 0.33   '�i�b�j���㌇���̖c�������Ɏg�����̑傫���i�Б��j
Const KOBUSHIAGE_MISSING_EROSION_SIZE       As Double = 0.33   '�i�b�j���㌇���̎��k�����Ɏg�����̑傫���i�Б��j
Const KOBUSHIAGE_TIME_HOSEI_COEF_MISSING    As Double = 0.2     '���㎞�ԕ␳�W�� �m�F�K�v�Ȍ�����Ԑ��P������

'makeGraph�AoutputCaption�AfixGraphDataAndSheet���W���[���̒��ɏ����ݒ�V�[�g�̃Z��������l��ǂݏo����������

'======================================================================================
'�|�C���g�v�Z�V�[�g��̊e�f�[�^�̍s�ԍ��A��ԍ����`
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
Const COLUMN_CAPTION_TRACK1                 As Long = 212 '�f�o�b�O�p�i���i�͎g��Ȃ��j

'=== �g���^�ԑ̓��L ===
Const COLUMN_DATA_RESULT_ALL    As Long = 203
Const COLUMN_UDEAGE_RESULT      As Long = 224

Const COLUMN_DATA_MISSING_SECTION           As Long = 219

Const COLUMN_DATA_KOBUSHIAGE_MEASURE_SECTION_ORG    As Long = 221
Const COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_ORG    As Long = 222
Const COLUMN_KOBUSHIAGE_FORCED_SECTION              As Long = 223 '����A���Ȃ��A�G�Ȃ��̋����A����t���O�A�t���O�̋L��
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

Const COLUMN_CAPTION_TRACK2                            As Long = 235 '�f�o�b�O�p�i���i�͎g��Ȃ��j

Const COLUMN_TEKUBI_RZ_SPEED                           As Long = 237 '�E���y�ʒu�̍�
Const COLUMN_TEKUBI_LZ_SPEED                           As Long = 238 '�����y�ʒu�̍�
Const COLUMN_TEKUBI_Z_SPEED_OVER                       As Long = 239 '���y�ʒu�̍� �������l�����t���O
Const COLUMN_DATA_KOBUSHIAGE_MEASURE_SECTION_DST           As Long = 240 '���㑪����
Const COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_DST           As Long = 241 '���㌇�����
Const COLUMN_MEAGERE_TIME_MACROUPDATEDATA              As Long = 242 'MacroUpdateData�̏������Ԃ𑪒茋�ʂ��i�[����

Const COLUMN_DATA_RESULT_GH_KOBUSHIAGE      As Long = 245
Const COLUMN_DATA_RESULT_GH_KOSHIMAGE       As Long = 247
Const COLUMN_DATA_RESULT_GH_HIZAMAGE        As Long = 249
Const COLUMN_DATA_RESULT_GH_SONKYO          As Long = 251

Const COLUMN_GH_HIZA_L                      As Long = 252
Const COLUMN_GH_HIZA_R                      As Long = 253

Const COLUMN_MAX_NUMBER                                As Long = 256 '���ݎg�p����Ă����ԍ��̍ő�l

'======================================================================================
'(�g���^�ԑ̓��L)�H���]���V�[�g�̊e�f�[�^�̍s�ԍ��A��ԍ����`
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
'�����킵���Ɋւ���萔
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_CHECK         As Long = 38
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_WORKTYPE      As Long = 42
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_DIAMETER      As Long = 43
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_LENGTH        As Long = 45
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_COMBINATION   As Long = 47
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_DIRECTION     As Long = 49
Const TB_HYOUKA_SHEET_COLUMN_KOZEWASHISA_NUMBER        As Long = 51

'���ɂ����Ɋւ���萔
Const TB_HYOUKA_SHEET_COLUMN_MINIKUSA_CHECK            As Long = 52
Const TB_HYOUKA_SHEET_COLUMN_MINIKUSA_SIZE             As Long = 56
Const TB_HYOUKA_SHEET_COLUMN_MINIKUSA_DISTANCE         As Long = 58
Const TB_HYOUKA_SHEET_COLUMN_MINIKUSA_CONTRAST         As Long = 60
Const TB_HYOUKA_SHEET_COLUMN_MINIKUSA_DIRECTION        As Long = 62

'======================================================================================
'�p���d�ʓ_�����[�V�[�g�̊e�f�[�^�̍s�ԍ��A��ԍ����`
'======================================================================================
Const SHIJUTEN_SHEET_ROW_KOUTEI_NAME                            As Long = 3
Const SHIJUTEN_SHEET_ROW_POSESTART_INDEX                        As Long = 9
Const SHIJUTEN_SHEET_ROW_EXPAND_NUMBER_CHECK                    As Long = 29

Const SHIJUTEN_SHEET_EXPAND_NUM_CHECK_WORD                      As String = "���̑��̎��ԁi�莞�ғ�����7.5H-�����׎��ԁj"


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
    Const SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_TIME                     As Long = 49 '���㎞��
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME                      As Long = 51 '���Ȃ�����
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME                       As Long = 53 '�G�Ȃ�����
    Const SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_MISSING_TIME             As Long = 55 '���㌇�����
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_MISSING_TIME              As Long = 57 '���Ȃ��������
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_PREDICT_TIME              As Long = 58 '���Ȃ�������
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_MISSING_TIME               As Long = 60 '�G�Ȃ��������
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_PREDICT_TIME               As Long = 61 '�G�Ȃ�������

#ElseIf VALIANT = TMC_CAR_BODY Then
    Const SHIJUTEN_SHEET_COLUMN_WORKEND_TIME                        As Long = 36
    Const SHIJUTEN_SHEET_COLUMN_DATA_MISSING_SECTION                As Long = 45
    Const SHIJUTEN_SHEET_COLUMN_DATA_PREDICT_SECTION                As Long = 46
    Const SHIJUTEN_SHEET_COLUMN_UDEAGE_TIME                         As Long = 48 '�r�グ����
    Const SHIJUTEN_SHEET_COLUMN_UDEAGE_REMOVE_OR_FORCED             As Long = 49 '�r�グ�]������
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME                      As Long = 51 '���Ȃ�����
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_REMOVE_OR_FORCED          As Long = 52 '���Ȃ��]������
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME                       As Long = 54 '�G�Ȃ�����
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_REMOVE_OR_FORCED           As Long = 55 '�G�Ȃ��]������
    Const SHIJUTEN_SHEET_COLUMN_UDEAGE_MISSING_TIME                 As Long = 57 '�r�グ�������
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_MISSING_TIME              As Long = 59 '���Ȃ��������
    Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_PREDICT_TIME              As Long = 60 '���Ȃ�������
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_MISSING_TIME               As Long = 62 '�G�Ȃ��������
    Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_PREDICT_TIME               As Long = 63 '�G�Ȃ�������
#End If



'======================================================================================
'�H���]���V�[�g�̊e�f�[�^�̍s�ԍ��A��ԍ����`
'======================================================================================

Const GH_HYOUKA_SHEET_ROW_POSESTART                    As Long = 15
Const GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK          As Long = 115

Const GH_HYOUKA_SHEET_EXPAND_NUM_CHECK_WORD            As String = "���v"

Const GH_HYOUKA_SHEET_COLUMN_WORK_NUMBER               As Long = 2
Const GH_HYOUKA_SHEET_COLUMN_WORK_NAME                 As Long = 3
Const GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME            As Long = 12
Const GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME              As Long = 14
Const GH_HYOUKA_SHEET_COLUMN_WORK_TIME                 As Long = 16
Const GH_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME           As Long = 17
Const GH_HYOUKA_SHEET_COLUMN_KOSHIMAGE_TIME            As Long = 18
Const GH_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME             As Long = 19

'======================================================================================
'�O�̗p�@�p������̂������l���`
'======================================================================================

Const GH_ANGLE_KOSHIMAGE_MIN As Double = 30
Const GH_ANGLE_KOSHIMAGE_MAX As Double = 180
Const GH_ANGLE_HIZAMAGE_MIN  As Double = 60
Const GH_ANGLE_HIZAMAGE_MAX As Double = 180

'======================================================================================
'�������̒�`
'======================================================================================
Const CAPTION_TRACK2_FILE_NAME_SOEJI           As String = "2" '�����g���b�N�Q�p�̃t�@�C���������ɂ���Y��
'�e�펚���̃t�H���g�T�C�Y�W��
'����̒l�̂��߁A�l���������قǕ����͑傫��


#If VALIANT = DEFAULT_2nd Then

    '���悪�c�̎�
    Const TRACK1_TATE_UPPER_COEF                   As Long = 22 '�g���b�N1�p�F��i
    Const TRACK1_TATE_LOWER_COEF                   As Long = 11 '�g���b�N1�p�F���i
    Const TRACK2_TATE_1ST_COEF                     As Long = 22 '�g���b�N2�p�F�P�i��
    Const TRACK2_TATE_2ND_COEF                     As Long = 22 '�g���b�N2�p�F�Q�i��
    Const TRACK2_TATE_3RD_COEF                     As Long = 13 '�g���b�N2�p�F�R�i��
    
    '���悪���̎�
    Const TRACK1_YOKO_UPPER_COEF                   As Long = 30 '�g���b�N1�p�F��i
    Const TRACK1_YOKO_LOWER_COEF                   As Long = 15 '�g���b�N1�p�F���i
    Const TRACK2_YOKO_1ST_COEF                     As Long = 30 '�g���b�N2�p�F�P�i��
    Const TRACK2_YOKO_2ND_COEF                     As Long = 30 '�g���b�N2�p�F�Q�i��
    Const TRACK2_YOKO_3RD_COEF                     As Long = 18 '�g���b�N2�p�F�R�i��

#ElseIf VALIANT = TMC_CAR_BODY Then
'���悪���̎�

    Const TRACK2_TATE_1ST_COEF                     As Long = 20 '�g���b�N2�p�F�P�i��
    Const TRACK2_TATE_2ND_COEF                     As Long = 5 '�g���b�N2�p�F�Q�i��
    Const TRACK2_TATE_3RD_COEF                     As Long = 16 '�g���b�N2�p�F�R�i��
    Const TRACK2_TATE_4RD_COEF                     As Long = 20 '�g���b�N2�p�F�S�i��

    '���悪�c�̎�
    Const TRACK2_YOKO_1ST_COEF                     As Long = 22 '�g���b�N2�p�F�P�i��
    Const TRACK2_YOKO_2ND_COEF                     As Long = 5 '�g���b�N2�p�F�Q�i��
    Const TRACK2_YOKO_3RD_COEF                     As Long = 15 '�g���b�N2�p�F�R�i��
    Const TRACK2_YOKO_4RD_COEF                     As Long = 20 '�g���b�N2�p�F�S�i��

#End If


'�e�펚���̐F
Const COLOR_DATA_REMOVE_SECTION                As String = "#bfbfbf" '�O���[
Const COLOR_DATA_FORCED_SECTION                As String = "#0033cc" '�F
Const COLOR_DATA_MISSING_SECTION               As String = "#ff7c80" '��F
Const COLOR_DATA_PREDICT_SECTION               As String = "#fcf600" '���F
Const COLOR_DATA_MEASURE_SECTION               As String = "#00b0f0" '���F
Const COLOR_DATA_RESULT_GREEN                  As String = "#00b050" '�ΐF
Const COLOR_DATA_RESULT_YELLOW                 As String = "#ffc000" '���F
Const COLOR_DATA_RESULT_RED                    As String = "#c00000" '�ԐF
Const COLOR_DATA_RESULT_GLAY                   As String = "#bfbfbf" '�O���[

'�уO���t�̃f�[�^�i�M���x�j����������������i�����g���b�N1�p ��i�E���ɕ\���j
Const CAPTION_DATA_MEASURE_SECTION             As String = "�y�f�[�^�����ԁz"
Const CAPTION_DATA_PREDICT_SECTION             As String = "�y�f�[�^�����ԁz"
Const CAPTION_DATA_REMOVE_SECTION              As String = "�y�f�[�^���O��ԁz"
Const CAPTION_DATA_FORCED_SECTION              As String = "�y�f�[�^������ԁz"
Const CAPTION_DATA_MISSING_SECTION             As String = "�y�f�[�^������ԁz"

'�уO���t�̃f�[�^�i�M���x�j����������������i�����g���b�N2�p 2�i�ڂɕ\���j
Const CAPTION_DATA_TRACK2_MEASURE_SECTION      As String = "�y�f�[�^�����ԁz"
Const CAPTION_DATA_TRACK2_PREDICT_SECTION      As String = "�y�f�[�^�����ԁz"
Const CAPTION_DATA_TRACK2_REMOVE_SECTION       As String = "�y�f�[�^���O��ԁz"
Const CAPTION_DATA_TRACK2_FORCED_SECTION       As String = "�y�f�[�^������ԁz"
Const CAPTION_DATA_TRACK2_MISSING_SECTION      As String = "�y�f�[�^������ԁz"


#If VALIANT = DEFAULT_2nd Then

    '�O�̗p�̎���������i�����g���b�N2�p 3�i�ڂɕ\���j
    Const CAPTION_A_RESULT_NAME1  As String = "�@�@�@�@����"
    Const CAPTION_B_RESULT_NAME1  As String = "  �@�@���Ȃ��@ �@"
    Const CAPTION_C_RESULT_NAME1  As String = "�G�Ȃ�"
    
    '�O�̗p�̏�������������i�����g���b�N2�p 4�i�ڂɕ\���j
    Const CAPTION_A_RESULT_NAME2  As String = "��񂪌�����"
    Const CAPTION_B_RESULT_NAME2  As String = "30���ȏ�"
    Const CAPTION_C_RESULT_NAME2  As String = "60���ȏ�"

#ElseIf VALIANT = TMC_CAR_BODY Then
'���悪���̎�

    '�g���^�ԑ̗p�̎���������i�����g���b�N2�p 3�i�ڂɕ\���j
    Const CAPTION_A_RESULT_NAME  As String = "�@�@���� �@"
    Const CAPTION_B_RESULT_NAME  As String = "  �O���@�@"
    Const CAPTION_C_RESULT_NAME  As String = "�G�Ȃ��@�@"
    Const CAPTION_D_RESULT_NAME  As String = "  �L���@"

    '�g���^�ԑ̗p�̏�������������i�����g���b�N2�p 4�i�ڂɕ\���j
    Const CAPTION_A_RESULT_NAME2  As String = "��񂪌�����"
    Const CAPTION_B_RESULT_NAME2  As String = "45���ȏ�"
    Const CAPTION_C_RESULT_NAME2  As String = "90���ȏ�"
    Const CAPTION_D_RESULT_NAME2  As String = "90���ȏ�"

#End If

'�L���v�V�����m�C�Y������臒l
Const CAPTION_REMOVE_NOISE_SECOND              As Double = 0.1 '�L���v�V�����m�C�Y���������钷��(�b) �i�`�����Ȃ珜���j

'�p���f�_�̒l�ɂ���āA�΁^���^�Ԃ𕪂���ۂ̋��E����
Const DATA_SEPARATION_GREEN_BOTTOM             As Long = 1
Const DATA_SEPARATION_GREEN_TOP                As Long = 2
Const DATA_SEPARATION_YELLOW_BOTTOM            As Long = 3
Const DATA_SEPARATION_YELLOW_TOP               As Long = 5
Const DATA_SEPARATION_RED_BOTTOM               As Long = 6
Const DATA_SEPARATION_RED_TOP                  As Long = 10

'======================================================================================
'DataAdjustingSheet�p
'======================================================================================
'debug
'Const LIMIT_COLUMN           As Long = 800
Const LIMIT_COLUMN           As Long = 16200

'---------------------------------------------
'�p���f�_�C���V�[�g�@�֘A
'---------------------------------------------
'LIMIT_COLUMN�̐ݒ�l��3�̔{���Ƃ���K�v������
'30fps�~60�b�~9����16200
Const SHEET_LIMIT_COLUMN       As Long = LIMIT_COLUMN + COLUMN_ZERO_NUM

'�񕝗p�̗�
Private Enum widthSize
    Small = 1
    Medium = 2
    Large = 4
    LL = 6
End Enum

'�񕝒����{�^�����O
Const EXPANDBTN_NAME           As String = "expandBtn"
Const REDUCEBTN_NAME           As String = "reduceBtn"

