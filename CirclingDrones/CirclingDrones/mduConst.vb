'*******************************************************
'      定数・変数定義用モジュール
'*******************************************************
Module mduConst

    '*******************************************************
    '      定数定義
    '*******************************************************

    '画面用
    Public Const pcAppName As String = "Circling Drones"            'アプリ名
    Public Const pcMax_d As Integer = 5                             'ドローン台数の最大値
    Public Const pcMax_v As Integer = 5                             '周回の最大値

    'モジュール用
    Public Const pcDoNothing As String = "DoNothing"                'DoNothingファイル名
    Public Const pcCalcProp As String = "CalcProp"                   'CalcPropファイル名
    Public Const pc3d_pot_test_plus As String = "3d_pot_test_plus"  '3d_pot_test_plusファイル名

    'INIファイル用
    Public Const pcIniFileName As String = "RKA_con.ini"

    Public Const pcSec_Route As String = "Route"
    Public Const pcKey_v As String = "V"
    Public Const pcKey_t As String = "T"
    Public Const pcSec_FileFlg As String = "FileFlg"
    Public Const pcKey_sw As String = "SW"
    Public Const pcKey_Exe1 As String = "Exe1_Path"
    Public Const pcKey_Exe2 As String = "Exe2_Path"
    Public Const pcKey_Exe22 As String = "Exe22"

    Public Const pcSec_Set As String = "Set"
    Public Const pcKey_11 As String = "PjName"
    Public Const pcKey_12 As String = "GnuPath"
    Public Const pcKey_13 As String = "PjPath"
    Public Const pcKey_14 As String = "P2mFile"

    Public Const pcSec_d1 As String = "d1"
    Public Const pcSec_d2 As String = "d2"
    Public Const pcSec_d3 As String = "d3"
    Public Const pcSec_d4 As String = "d4"
    Public Const pcSec_d5 As String = "d5"
    Public Const pcKey_X As String = "X"
    Public Const pcKey_Y As String = "Y"
    Public Const pcKey_Z As String = "Z"

    Public Const pcSec_Command As String = "Command"
    Public Const pcKey_1 As String = "1"
    Public Const pcKey_2 As String = "2"
    Public Const pcKey_3 As String = "3"
    Public Const pcKey_4 As String = "4"
    Public Const pcKey_5 As String = "5"

    '経路計算用
    Public Const pc130_Grid As String = "130.1K2490GD2route_grid.csv"
    Public Const pc130_Meter As String = "130.1K249GD2route_meter.csv"
    Public Const pcP2mFileCnt As Integer = 9

    '*******************************************************
    '      変数定義
    '*******************************************************
    'テーブル構造体
    Structure M_Iti_Def
        Dim sIdo As String     '緯度
        Dim sKeido As String   '経度
        Dim sTakasa As String  '高さ
    End Structure
    Structure M_P2m_Def
        Dim sRx As String       'Rx
        Dim sX As String        'X
        Dim sY As String        'Y
        Dim sZ As String        'Z
        Dim sDistance As String 'Distance
        Dim sPower As String    'Power
        Dim sPhase As String    'Phase
    End Structure
    Structure M_Txrx_Def
        Dim sdv As String       'ドローン情報
        Dim sIdo As String      '緯度
        Dim sKeido As String    '経度
        Dim sTakasa As String   '高さ
    End Structure
    Structure M_OkNg_Def
        Dim iLen As Integer     'カーソル位置
        Dim iColor As Integer   '色
    End Structure
    Structure M_130_Def
        Dim sX As String        'X
        Dim sY As String        'Y
        Dim sZ As String        'Z
        Dim sAA As String       'AA
        Dim dCal1 As Double     '計算結果1
        Dim dCal2 As Double     '計算結果2
        Dim iX As Integer       'X
        Dim iY As Integer       'Y
        Dim iZ As Integer       'Z
    End Structure
    Structure M_Sabun_Def
        Dim sX As String        'x
        Dim sY As String        'y
        Dim sSabun As String    'Sabun
    End Structure

    '経路計算用
    Public pIti(,) As M_Iti_Def         '位置情報
    Public pOrder_d() As String         'ドローン順番
    Public pP2m() As M_P2m_Def          'P2m情報
    Public pTxrx() As M_Txrx_Def        'Txrx情報
    Public p130() As M_130_Def          '経路計算情報
    Public pSabun(4888) As M_Sabun_Def  'sabun情報
    Public pPower(43992) As String    'Power更新情報

    'ファイル切替用
    '=========================================
    '0:ワイヤレスインサイトがない環境
    Public pFileFlg As Integer
    '=========================================

    '画面用
    Public pSet_d As Integer        'ドローン設定台数
    Public pSet_Interval As Integer '周回の間隔(ミリ秒)
    Public pSet_v As Integer        '周回設定
    Public pOkNg_No As Integer      'OKorNG表示No
    Public pOkNg() As M_OkNg_Def    'OKorNG表示色設定
    '
    'INIファイル用
    Public pV As String             'ドローン速度
    Public pT As String             '周回毎経過時間

    Public pIniPath As String       'INIファイルパス
    Public pPjName As String        'プロジェクト名
    Public pP2mDir As String        'p2m用フォルダ
    Public pGnuPath As String       'gnuプロットパス
    Public pPjPath As String        'pjフォルダパス
    Public pExePath As String       'exeフォルダパス
    Public pExe1_Path As String     'DoNothingパス
    Public pExe2_Path As String     'calcpropパス
    Public pExe22 As String         'calcprop引数

    'Public pX_d1 As String          'd1のX設定値
    'Public pY_d1 As String          'd1のY設定値
    'Public pZ_d1 As String          'd1のZ設定値
    'Public pX_d2 As String          'd2のX設定値
    'Public pY_d2 As String          'd2のY設定値
    'Public pZ_d2 As String          'd2のZ設定値
    'Public pX_d3 As String          'd3のX設定値
    'Public pY_d3 As String          'd3のY設定値
    'Public pZ_d3 As String          'd3のZ設定値
    'Public pX_d4 As String          'd4のX設定値
    'Public pY_d4 As String          'd4のY設定値
    'Public pZ_d4 As String          'd4のZ設定値
    'Public pX_d5 As String          'd5のX設定値
    'Public pY_d5 As String          'd5のY設定値
    'Public pZ_d5 As String          'd5のZ設定値

    Public pCommand(5) As String        'プロットコマンド値

    Public pEnc As System.Text.Encoding = New System.Text.UTF8Encoding(False)
    'Public pEnc As System.Text.Encoding = System.Text.Encoding.Default
    'Public pEnc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")
End Module
