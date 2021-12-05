'*******************************************************
'      メインモジュール
'*******************************************************
Module mduMain
    '*******************************************************
    '      定数定義
    '*******************************************************
    '画面用
    Public Const pcAppName As String = "Circling Drones"     'アプリ名
    Public Const pcMax_d As Integer = 5                      'ドローン台数の最大値
    Public Const pcMax_v As Integer = 1000                   '周回の最大値

    'モジュール用
    Public Const pSetup As String = "DoNothing"
    Public Const p3d_pot_test_plus As String = "3d_pot_test_plus"

    '制御用
    Public Const pcIniFileName As String = "RKA_con.ini"
    Public Const pcSec_Set As String = "Set"
    Public Const pcKey_11 As String = "PjName"
    Public Const pcKey_12 As String = "GnuPath"
    Public Const pcKey_13 As String = "PjPath"
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

    Public Const pcExt_Out_Grid As String = "130.1K2490GD2route_grid.csv"
    Public Const pcExt_Out_Meter As String = "130.1K249GD2route_meter.csv"
    Public Const pcExt_Out_Height As String = "true_height.csv"
    Public Const pcExt_Out_Power As String = "received_power.txt"


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

    '制御用
    Public pIti(,) As M_Iti_Def         '位置情報
    Public pOrder_d() As String         'ドローン順番
    Public pP2m() As M_P2m_Def          'P2m情報
    Public pTxrx() As M_Txrx_Def        'Txrx情報

    '画面用
    Public pSet_d As Integer            'ドローン設定台数
    Public pSet_Interval As Integer     '周回の間隔(ミリ秒)
    Public pSet_v As Integer            '周回設定
    Public pOkNg_No As Integer          'OKorNG表示No
    Public pOkNg() As M_OkNg_Def        'OKorNG表示色設定
    '
    Public pIniPath As String           'iniファイルパス
    Public pPjName As String            'プロジェクト名
    Public pGnuPath As String           'gnuプロットパス
    Public pPjPath As String            'pjフォルダパス
    Public pTempPath As String          'tempフォルダパス

    Public pX_d1 As String              'd1のX設定値
    Public pY_d1 As String              'd1のY設定値
    Public pZ_d1 As String              'd1のZ設定値
    Public pX_d2 As String              'd2のX設定値
    Public pY_d2 As String              'd2のY設定値
    Public pZ_d2 As String              'd2のZ設定値
    Public pX_d3 As String              'd3のX設定値
    Public pY_d3 As String              'd3のY設定値
    Public pZ_d3 As String              'd3のZ設定値
    Public pX_d4 As String              'd4のX設定値
    Public pY_d4 As String              'd4のY設定値
    Public pZ_d4 As String              'd4のZ設定値
    Public pX_d5 As String              'd5のX設定値
    Public pY_d5 As String              'd5のY設定値
    Public pZ_d5 As String              'd5のZ設定値

    Public pCommand(5) As String        'プロットコマンド値

    '*******************************************************
    '      前処理
    '******************************************************* 
    Public Sub subPreProc()
        On Error GoTo Error_Rtn

        'ドローン毎の周回用ファイルを削除
        If fncFileDelete_d() Then
            Call subLogOutput("> " & "ドローン毎の周回用ファイルを削除=>OK") '画面ログ出力
            Call subOkNg_Color(0) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Else
            Call subLogOutput("> " & "ドローン毎の周回用ファイルを削除=>NG") '画面ログ出力
            Call subOkNg_Color(1) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        End If

        'Itiファイルを作成する
        For m = 1 To pSet_d
            If fncItiFile(m, 0) Then
                Call subLogOutput("> " & "d" & m & "_Itiファイルの作成=>OK") '画面ログ出力
                Call subOkNg_Color(0) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
            Else
                Call subLogOutput("> " & "d" & m & "_Itiファイルの作成=>OK") '画面ログ出力
                Call subOkNg_Color(1) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
            End If
        Next

        'スタート前にtxrxファイルを削除
        If fncFileDelete_txrx() Then
            Call subLogOutput("> " & "Txrxファイルを削除=>OK") '画面ログ出力
            Call subOkNg_Color(0) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Else
            Call subLogOutput("> " & "Txrxファイルを削除=>NG") '画面ログ出力
            Call subOkNg_Color(1) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '      周回処理
    '******************************************************* 
    Public Sub subProc(ByVal m As Integer, ByVal n As Integer)
        On Error GoTo Error_Rtn

        Dim sSetup As String = pSetup & ".exe"
        Dim s3d_pot_test_plus As String = p3d_pot_test_plus & ".exe"

        Call subLogOutput("*** 周回処理(d" & m & "v" & n & ") ***") '画面ログ出力

        'プロジェクト名.setup 起動
        System.Diagnostics.Process.Start(sSetup, pPjName & ".setup 起動")
        Call subLogOutput("> " & pPjName & ".setup 起動") '画面ログ出力
        Call subOkNg_Color(2) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Call subiInterval(500) 'インターバル設定

        'プロジェクト名.p2m取込み
        If fncReadP2m(pPjPath + "\" & pPjName & ".p2m") Then
            Call subLogOutput("> " & pPjName & ".p2mファイル取込み=>OK") '画面ログ出力
            Call subOkNg_Color(0) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Else
            Call subLogOutput("> " & pPjName & ".p2mファイル取込み=>NG") '画面ログ出力
            Call subOkNg_Color(1) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        End If

        'power.txt作成
        If fncMem2Power(pPjPath) Then
            Call subLogOutput("> power.txt作成=>OK") '画面ログ出力
            Call subOkNg_Color(0) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Else
            Call subLogOutput("> power.txt作成=>NG") '画面ログ出力
            Call subOkNg_Color(1) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        End If

        '3d_pot_test_plus.exe 起動
        System.Diagnostics.Process.Start(s3d_pot_test_plus)
        Call subLogOutput("> 3d_pot_test_plus.exe 起動") '画面ログ出力
        Call subOkNg_Color(2) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Call subiInterval(500) 'インターバル設定

        '出力ファイル リネーム
        If fncExtOutRename(m, n) Then
            Call subLogOutput("> 出力ファイル リネーム=>OK") '画面ログ出力
            Call subOkNg_Color(0) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Else
            Call subLogOutput("> 出力ファイル リネーム=>NG") '画面ログ出力
            Call subOkNg_Color(1) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        End If

        '周回の間隔
        Call subLogOutput("> 周回の間隔") '画面ログ出力
        Call subiInterval(pSet_Interval) 'インターバル設定
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '      後処理
    '*******************************************************
    Public Sub subAfterProc(ByVal m As Integer, ByVal n As Integer)
        On Error GoTo Error_Rtn

        Call subLogOutput("*** 後処理(d" & m & "v" & n & ") ***") '画面ログ出力

        '経路計算は「fncItiFile」関数の中で行う
        Call subLogOutput("> 経路計算") '画面ログ出力

        'Itiファイルを作成する
        If fncItiFile(m, n) Then
            Call subLogOutput("> " & "d" & m & "_Itiファイルの作成=>OK") '画面ログ出力
            Call subOkNg_Color(0) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Else
            Call subLogOutput("> " & "d" & m & "_Itiファイルの作成=>NG") '画面ログ出力
            Call subOkNg_Color(1) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        End If

        'プロジェクト名.txrxファイルを削除
        If fncFileDelete_pj_txrx() Then
            Call subLogOutput("> " & "Txrxファイルを削除=>OK") '画面ログ出力
            Call subOkNg_Color(0) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Else
            Call subLogOutput("> " & "Txrxファイルを削除=>NG") '画面ログ出力
            Call subOkNg_Color(1) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        End If

        'Itiファイルからtxrx用配列を作成
        If fncItiFile2Txrx(m, n) Then
            Call subLogOutput("> Itiファイルからtxrx用配列作成=>OK") '画面ログ出力
            Call subOkNg_Color(0) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Else
            Call subLogOutput("> Itiファイルからtxrx用配列作成=>NG") '画面ログ出力
            Call subOkNg_Color(1) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        End If

        'プロジェクト名.txrxファイルを作成
        If fncMem2txrx() Then
            Call subLogOutput("> txrxファイル 作成=>OK") '画面ログ出力
            Call subOkNg_Color(0) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        Else
            Call subLogOutput("> txrxファイル 作成=>NG") '画面ログ出力
            Call subOkNg_Color(1) 'pOkNg設定（0：緑色、1：赤色、2：黄色）
        End If

Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub
End Module
