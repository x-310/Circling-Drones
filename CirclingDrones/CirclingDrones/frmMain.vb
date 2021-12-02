'*******************************************************
'      画面表示用モジュール
'*******************************************************
Imports System.ComponentModel

Public Class frmMain
    Dim sw As New System.Diagnostics.Stopwatch()
    Dim n As Integer '周回カウント

    '*******************************************************
    '   メイン
    '*******************************************************
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error GoTo Error_Rtn

        Dim sSetup As String = pSetup & ".exe"
        Dim s3d_pot_test_plus As String = p3d_pot_test_plus & ".exe"

        If fncFileCheck(sSetup) = False Or fncFileCheck(s3d_pot_test_plus) = False Then
            fncMsgBox("実行ファイルを確認して下さい")
            End
        End If

        '画面コントロール初期化
        Call subCtlReset()

        'Iniファイル読込み
        Call subSetIni()
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   スタート
    '*******************************************************
    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        On Error GoTo Error_Rtn

        Dim sStartTime As DateTime  'スタート時間
        Dim sStopTime As DateTime   'ストップ時間
        Dim sTime As String         '経過時間
        Dim m As Integer            'ドローン毎ループ
        Dim iRow As Integer         '行No

        Call subSetGamen()          '画面設定情報を取得
        Call subBtnOnoff(False)     '画面ボタン制御
        ' 計測開始
        sw.Start()                  'ストップウォッチ
        sStartTime = DateTime.Now   'スタート時間
        Call subLogOutput("スタート--> " + sStartTime) '画面ログ出力

        'プロセス削除
        Call subKillProc(pSetup)

        '前処理
        Call subLogOutput("*** 前処理 ***") '画面ログ出力
        ReDim Preserve pOrder_d(-1)         'ドローン順番用配列
        ReDim Preserve pIti(pMax_d, -1)     '位置情報
        ReDim Preserve pTxrx(-1)            'Txrx用配列
        pOkNg_No = 0                        'OKorNG表示No
        ReDim pOkNg(-1)                     'OKorNG表示色設定
        Call subPreProc() '前処理

        'ドローン毎周回処理
        n = 1       '周回
        iRow = 0    '行No
        Do While txtStopFlg.Text = "0"  '周回ループ
            For m = 1 To pSet_d     'ドローン毎ループ
                txtNow_m.Text = CInt(m)  '現在ドローン
                txtNow_n.Text = CInt(n)  '現在周回

                ReDim Preserve pOrder_d(iRow) 'ドローン順番
                pOrder_d(iRow) = "d" & txtNow_m.Text & "v" & CInt(n).ToString("00")

                '周回処理
                Call subProc(CInt(txtNow_m.Text), CInt(txtNow_n.Text)) '周回処理

                '後処理
                Call subAfterProc(CInt(txtNow_m.Text), CInt(txtNow_n.Text)) '後処理

                '緯度・経度・高さ情報を画面ログに出力する
                Call subLogOutput("*** 緯度,経度,高さ(pIti) ***") '画面ログ出力
                Call subLogOutput("d" & m & "," &
                                  "v" & n & "," &
                                  pIti(m, n).sIdo & "," &
                                  pIti(m, n).sKeido & "," &
                                  pIti(m, n).sTakasa) '画面ログ出力
                'Txrx情報を画面ログに出力する
                Call subLogOutput("*** ドローン情報,緯度,経度,高さ(pTxrx) ***") '画面ログ出力
                For iLoop = 0 To pTxrx.Length - 1
                    Call subLogOutput(pTxrx(iLoop).sdv & "," &
                                      pTxrx(iLoop).sIdo & "," &
                                      pTxrx(iLoop).sKeido & "," &
                                      pTxrx(iLoop).sTakasa) '画面ログ出力
                Next
                iRow = iRow + 1 '行No
            Next

            '周回終了判定
            If n >= pSet_v Then
                '終了
                txtStopFlg.Text = "1"

                'プロセス削除
                Call subKillProc(pSetup)
            Else
                '周回カウントアップ
                n = n + 1
            End If
        Loop

        ' 計測停止
        sw.Stop()
        sStopTime = DateTime.Now
        Call subLogOutput("エンド--> " + sStopTime) '画面ログ出力

        sTime = CLng(sw.ElapsedMilliseconds)
        Call subLogOutput("-------------------") '画面ログ出力
        Call subLogOutput("経過時間(ms)> " & sTime) '画面ログ出力
        Call subLogOutput("-------------------") '画面ログ出力
        'ドローン順を表示する
        Call subLogOutput("*** ドローン順番(pOrder_d) ***") '画面ログ出力
        For iLoop = 0 To pOrder_d.Length - 1
            Call subLogOutput(iLoop + 1 & "." & pOrder_d(iLoop)) '画面ログ出力
        Next
        '画面ログ出力（0：緑色、1：赤色）
        Dim iCnt As Integer = 0
        For iLoop = 0 To pOkNg_No - 1
            Call subLogOutput_Color(pOkNg(iLoop).iLen, pOkNg(iLoop).iColor)
            If pOkNg(iLoop).iColor = 1 Then iCnt = iCnt + 1
        Next
        txtNG.Text = CInt(iCnt) 'NGカウント表示

        Call subSetRow() 'テキストボックスのカーソル制御
        Call subBtnOnoff(True) '画面ボタン制御
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   gnuプロット
    '*******************************************************
    Private Sub btmGnu_Click(sender As Object, e As EventArgs) Handles btmGnu.Click
        On Error GoTo Error_Rtn

        Dim p As System.Diagnostics.Process =
            System.Diagnostics.Process.Start(pGnuPath)
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   終了
    '*******************************************************
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        On Error GoTo Error_Rtn

        Me.Close()
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   iniファイル読込
    '*******************************************************

    Private Sub btnIniGet_Click(sender As Object, e As EventArgs) Handles btnIniGet.Click
        On Error GoTo Error_Rtn

        Dim sFileName As String = txtIniPath.Text & "\" & pIniFileName

        If fncFileCheck(sFileName) Then
            Call subSetIni()

            fncMsgBox("読込しました")
        Else
            fncErrors("iniファイルのパスを確認して下さい")
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   iniファイル保存
    '*******************************************************
    Private Sub btnIniSave_Click(sender As Object, e As EventArgs) Handles btnIniSave.Click
        On Error GoTo Error_Rtn

        Dim sFileName As String = txtIniPath.Text & "\" & pIniFileName

        If fncFileCheck(sFileName) Then
            Call subPutIni()
            Call subSetIni()

            fncMsgBox("保存しました")
        Else
            fncErrors("iniファイルのパスを確認して下さい")
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   iniファイル読込
    '*******************************************************
    Private Sub btnIniGet2_Click(sender As Object, e As EventArgs) Handles btnIniGet2.Click
        On Error GoTo Error_Rtn

        Dim sFileName As String = txtIniPath.Text & "\" & pIniFileName

        If fncFileCheck(sFileName) Then
            Call subSetIni()

            fncMsgBox("読込しました")
        Else
            fncErrors("iniファイルのパスを確認して下さい")
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   iniファイル保存
    '*******************************************************
    Private Sub btnIniSave2_Click(sender As Object, e As EventArgs) Handles btnIniSave2.Click
        On Error GoTo Error_Rtn

        Dim sFileName As String = txtIniPath.Text & "\" & pIniFileName

        If fncFileCheck(sFileName) Then
            Call subPutIni()
            Call subSetIni()

            fncMsgBox("保存しました")
        Else
            fncErrors("iniファイルのパスを確認して下さい")
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   iniファイル
    '*******************************************************
    Private Sub btnIni_Click_1(sender As Object, e As EventArgs) Handles btnIni.Click
        On Error GoTo Error_Rtn

        Dim p As System.Diagnostics.Process =
            System.Diagnostics.Process.Start(pIniPath + "\" + "RKA_con.ini")
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   フォームクローズ
    '*******************************************************
    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        On Error GoTo Error_Rtn

        'プロセス削除
        Call subKillProc(pSetup)
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   P2mファイル取込みテスト用
    '*******************************************************
    Private Sub btnGetP2m_Click(sender As Object, e As EventArgs) Handles btnGetP2m.Click
        If fncReadP2m(pPjPath + "\b2.p2m") Then
            fncMsgBox("OK")
        Else
            fncMsgBox("NG")
        End If
    End Sub

    '*******************************************************
    '   Power.txt作成テスト用
    '*******************************************************
    Private Sub btnPutPower_Click(sender As Object, e As EventArgs) Handles btnPutPower.Click
        If fncMem2Power(pPjPath) Then
            fncMsgBox("OK")
        Else
            fncMsgBox("NG")
        End If
    End Sub

    '*******************************************************
    '   pjフォルダ選択
    '*******************************************************
    Private Sub btnPj_Click(sender As Object, e As EventArgs) Handles btnPj.Click
        On Error GoTo Error_Rtn

        System.Diagnostics.Process.Start(txtPjPath.Text)

Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   exeフォルダ選択
    '*******************************************************
    Private Sub btnExe_Click(sender As Object, e As EventArgs) Handles btnExe.Click

        On Error GoTo Error_Rtn

        System.Diagnostics.Process.Start(CurDir())

Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub
End Class