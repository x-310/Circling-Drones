'*******************************************************
'      共通関数用モジュール
'*******************************************************
Module mduCommon
    '*******************************************************
    '      画面コントロール初期化
    '*******************************************************
    Public Sub subCtlReset()
        On Error GoTo Error_Rtn

        frmMain.cmbSet_d.Text = "2"
        frmMain.cmbSet_Interval.Text = "1000"
        frmMain.cmbSet_v.Text = "2"

        frmMain.txtNow_m.Text = ""      '現在ローンNo
        frmMain.txtNow_n.Text = ""      '現在周目
        frmMain.txtLogDisp.Text = ""    '画面ログ出力クリア

        frmMain.txtIniPath.Text = ""     'iniファイルパス
        frmMain.txtPjName.Text = ""      'プロジェクト名
        frmMain.txtGnuPath.Text = ""     'gnuプロットパス
        frmMain.txtPjPath.Text = ""      'pjフォルダパス

        frmMain.txtX_d1.Text = ""        'X設定値
        frmMain.txtY_d1.Text = ""        'Y設定値
        frmMain.txtZ_d1.Text = ""        'Z設定値
        frmMain.txtX_d2.Text = ""        'X設定値
        frmMain.txtY_d2.Text = ""        'Y設定値
        frmMain.txtZ_d2.Text = ""        'Z設定値
        frmMain.txtX_d3.Text = ""        'X設定値
        frmMain.txtY_d3.Text = ""        'Y設定値
        frmMain.txtZ_d3.Text = ""        'Z設定値
        frmMain.txtX_d4.Text = ""        'X設定値
        frmMain.txtY_d4.Text = ""        'Y設定値
        frmMain.txtZ_d4.Text = ""        'Z設定値
        frmMain.txtX_d5.Text = ""        'X設定値
        frmMain.txtY_d5.Text = ""        'Y設定値
        frmMain.txtZ_d5.Text = ""        'Z設定値

        frmMain.txtCom1.Text = ""        'コマンド
        frmMain.txtCom2.Text = ""        'コマンド
        frmMain.txtCom3.Text = ""        'コマンド
        frmMain.txtCom4.Text = ""        'コマンド
        frmMain.txtCom5.Text = ""        'コマンド

        frmMain.txtStopFlg.Text = "0"   '周回終了フラグ=1:終了
        frmMain.txtNG.Text = "" 'NGカウント表示
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '      インターバル設定
    '*******************************************************
    Public Sub subiInterval(ByVal iInterval As Integer)
        On Error GoTo Error_Rtn

        'インターバル設定
        System.Threading.Thread.Sleep(iInterval)
        'イベント発生
        Application.DoEvents()
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '      画面ボタン制御
    '*******************************************************
    Public Sub subBtnOnoff(ByVal bFlg As Boolean)
        On Error GoTo Error_Rtn

        frmMain.cmbSet_v.Enabled = bFlg '周回
        frmMain.btnStart.Enabled = bFlg 'オート・スタート

        frmMain.pnlSet.Enabled = bFlg   '設定エリア
        frmMain.pnlPath.Enabled = bFlg  'Pathエリア
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '      pOkNg設定（0：緑色、1：赤色、2：黄色）
    '*******************************************************
    Public Sub subOkNg_Color(ByVal iColor As Integer)
        On Error GoTo Error_Rtn

        ReDim Preserve pOkNg(pOkNg_No)                              'OKorNG表示色変更
        pOkNg(pOkNg_No).iLen = frmMain.txtLogDisp.Text.Length - 3   'OKorNG表示色設定
        pOkNg(pOkNg_No).iColor = iColor
        pOkNg_No = pOkNg_No + 1
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '      画面ログ出力（0：緑色、1：赤色、2：黄色）
    '*******************************************************
    Public Sub subLogOutput_Color(ByVal iLen As Integer, ByVal iColor As Integer)
        On Error GoTo Error_Rtn

        'カレットを該当文字目に移動
        frmMain.txtLogDisp.Select(iLen, 0)
        'frmMain.txtLogDisp.Select(frmMain.txtLogDisp.Text.Length - 3, 0)
        '2文字選択する：OK or NG
        frmMain.txtLogDisp.SelectionLength = 2
        '色を変更する
        Select Case iColor
            Case 0
                frmMain.txtLogDisp.SelectionColor = Color.Lime
            Case 1
                frmMain.txtLogDisp.SelectionColor = Color.Red
            Case 2
                frmMain.txtLogDisp.SelectionColor = Color.Yellow
            Case Else
                frmMain.txtLogDisp.SelectionColor = Color.White
        End Select
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '      画面ログ出力
    '*******************************************************
    Public Sub subLogOutput(ByVal sMessage As String)
        On Error GoTo Error_Rtn

        frmMain.txtLogDisp.Text = frmMain.txtLogDisp.Text + sMessage + Chr(13) + Chr(10)
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '      画面設定情報を取得
    '*******************************************************
    Public Sub subSetGamen()
        On Error GoTo Error_Rtn

        pSet_d = frmMain.cmbSet_d.Text                  'ドローン設定台数
        pSet_Interval = frmMain.cmbSet_Interval.Text    '周回の間隔(ミリ秒)
        pSet_v = frmMain.cmbSet_v.Text                  '周回

        frmMain.txtStopFlg.Text = "0"                   '周回ループ
        frmMain.txtNG.Text = "" 'NGカウント表示
        frmMain.txtLogDisp.Text = ""                    '画面ログ出力
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '      テキストボックスのカーソル制御
    '*******************************************************
    Public Sub subSetRow()
        On Error GoTo Error_Rtn

        'カレット位置を末尾に移動
        frmMain.txtLogDisp.SelectionStart = frmMain.txtLogDisp.Text.Length
        'テキストボックスにフォーカスを移動
        frmMain.txtLogDisp.Focus()
        'カレット位置までスクロール
        frmMain.txtLogDisp.ScrollToCaret()
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '      メッセージボックス
    '*******************************************************
    Public Function fncMsgBox(ByVal NAME) As Long
        On Error GoTo Error_Rtn

        Dim MSG As String = NAME

        MessageBox.Show(MSG, pcAppName, MessageBoxButtons.OK, MessageBoxIcon.Information)
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '      エラーメッセージ
    '*******************************************************
    Public Function fncErrors(ByVal NAME) As Long
        On Error GoTo Error_Rtn

        Dim MSG As String

        Select Case Err()
            Case Else
                MSG = NAME & Chr(13) & Err.Description
                MessageBox.Show(MSG, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Select
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '      プロセスを終了する
    '*******************************************************
    Public Sub subKillProc(ByVal NAME)
        On Error Resume Next

        Dim ps As System.Diagnostics.Process() =
        System.Diagnostics.Process.GetProcessesByName(NAME)

        For Each p As System.Diagnostics.Process In ps
            'プロセスを強制的に終了させる
            p.Kill()
        Next
    End Sub
End Module
