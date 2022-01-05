'*******************************************************
'      画面用モジュール
'*******************************************************
Imports System.ComponentModel

Public Class frmMain

    Dim sw As New System.Diagnostics.Stopwatch()    'ストップウォッチ
    Dim n As Integer                                '周回カウント

    '*******************************************************
    '   メイン画面起動後処理
    '*******************************************************
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error GoTo Error_Rtn

        Dim sExeFile As String = "" 'ワイヤレスインサイト
        Dim s3d_pot_test_plus As String = pc3d_pot_test_plus & ".exe"

        'ワイヤレスインサイト
        If pFileFlg = 0 Then
            sExeFile = pcDoNothing & ".exe"
        Else
            sExeFile = pcCalcProp & ".exe"
        End If

        '3d_pot_test_plus
        If fncFileCheck(sExeFile) = False Or fncFileCheck(s3d_pot_test_plus) = False Then
            fncMsgBox("実行ファイルを確認して下さい")
            End
        End If

        '画面コントロール初期化
        Call subCtlReset()

        'INIファイル読込み
        Call subSetIni()

        'exeパス
        pExePath = System.Environment.CurrentDirectory

        'Debugフラグ
        cmbDebug.Text = "OFF"
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   スタートボタン押下後処理
    '*******************************************************
    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        On Error GoTo Error_Rtn

        Dim sStartTime As DateTime  'スタート時間
        Dim sStopTime As DateTime   'ストップ時間
        Dim sTime As String         '経過時間
        Dim m As Integer            'ドローン毎ループ
        Dim iRow As Integer         '行No
        Dim sExeFile As String = "" 'ワイヤレスインサイト
        Dim s3d_pot_test_plus As String = pc3d_pot_test_plus

        'ワイヤレスインサイト
        If pFileFlg = 0 Then
            sExeFile = pcDoNothing
        Else
            sExeFile = pcCalcProp
        End If

        '画面設定情報を取得
        Call subSetGamen()

        'INIファイル読込み
        Call subSetIni()

        '画面ボタン制御
        Call subBtnOnoff(False)

        ' 計測開始
        sw.Start()
        sStartTime = DateTime.Now
        Call subLogOutput("スタート--> " + sStartTime)

        'プロセス削除
        Call subKillProc(sExeFile)

        ReDim Preserve pOrder_d(-1)         'ドローン順番用配列
        ReDim Preserve pIti(pcMax_d, -1)     '位置情報
        ReDim Preserve pTxrx(-1)            'Txrx用配列
        ReDim pOkNg(-1)                     'OKorNG表示色設定
        pOkNg_No = 0                        'OKorNG表示No

        '*** 前処理 *** 
        Call subPreProc()

        'ドローン毎周回処理
        n = 1       '周回
        iRow = 0    '行No
        Do While txtStopFlg.Text = "0"      '周回ループ
            For m = 1 To pSet_d             'ドローン毎ループ
                '現在ドローン
                txtNow_m.Text = CInt(m)

                '現在周回
                txtNow_n.Text = CInt(n)

                'ドローン順番
                ReDim Preserve pOrder_d(iRow)
                pOrder_d(iRow) = "d" & txtNow_m.Text & "v" & CInt(n).ToString("0000")

                '*** 周回処理 *** 
                Call subProc(CInt(txtNow_m.Text), CInt(txtNow_n.Text))

                '*** 後処理 *** 
                Call subAfterProc(CInt(txtNow_m.Text), CInt(txtNow_n.Text))

                '*** Logファイル出力処理 ***
                Call subLogFile(CInt(txtNow_m.Text), CInt(txtNow_n.Text))

                iRow = iRow + 1 '行No
            Next

            '周回終了判定
            If n >= pSet_v Then
                '終了
                txtStopFlg.Text = "1"

                'プロセス削除
                Call subKillProc(sExeFile)
            Else
                '周回カウントアップ
                n = n + 1
            End If
        Loop

        ' 計測停止
        sw.Stop()
        sStopTime = DateTime.Now
        Call subLogOutput("")
        Call subLogOutput("エンド--> " + sStopTime)

        '経過時間
        sTime = CLng(sw.ElapsedMilliseconds)
        Call subLogOutput("")
        Call subLogOutput("-------------------")
        Call subLogOutput("経過時間(ms)> " & sTime)
        Call subLogOutput("-------------------")

        'ドローン順を表示する
        Call subLogOutput("")
        Call subLogOutput("*** ドローン順番(pOrder_d) ***")
        For iLoop = 0 To pOrder_d.Length - 1
            Call subLogOutput(iLoop + 1 & "." & pOrder_d(iLoop))
        Next

        'NGカウント表示
        Dim iCnt As Integer = 0
        For iLoop = 0 To pOkNg_No - 1
            Call subLogOutput_Color(pOkNg(iLoop).iLen, pOkNg(iLoop).iColor)
            If pOkNg(iLoop).iColor = 1 Then iCnt = iCnt + 1
        Next
        txtNG.Text = CInt(iCnt)

        'テキストボックスのカーソル制御
        Call subSetRow()

        '画面ボタン制御
        Call subBtnOnoff(True)
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   実行フォルダボタン押下後処理
    '*******************************************************
    Private Sub btnExe_Click(sender As Object, e As EventArgs) Handles btnExe.Click

        On Error GoTo Error_Rtn

        'exeフォルダを開く
        System.Diagnostics.Process.Start(CurDir())
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   PJフォルダボタン押下後処理
    '*******************************************************
    Private Sub btnPj_Click(sender As Object, e As EventArgs) Handles btnPj.Click
        On Error GoTo Error_Rtn

        'pjフォルダを開く
        System.Diagnostics.Process.Start(txtPjPath.Text)
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   INIファイルボタン押下後処理
    '*******************************************************
    Private Sub btnIni_Click_1(sender As Object, e As EventArgs) Handles btnIni.Click
        On Error GoTo Error_Rtn

        'INIファイルを起動
        Dim p As System.Diagnostics.Process =
            System.Diagnostics.Process.Start(pIniPath + "\" + "RKA_con.ini")
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   gnuプロットボタン押下後処理
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
    '   終了ボタン押下後処理
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
    '   タブコントロール1選択処理
    '*******************************************************
    Private Sub TabControl1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles TabControl1.DrawItem
        '対象のTabControlを取得
        Dim tab As TabControl = CType(sender, TabControl)
        'タブページのテキストを取得
        Dim txt As String = tab.TabPages(e.Index).Text
        'タブのテキストと背景を描画するためのブラシを決定する
        Dim foreBrush, backBrush As Brush

        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            '選択されているタブのテキストを白、背景を青とする
            foreBrush = Brushes.White
            backBrush = Brushes.DarkBlue
        Else
            '選択されていないタブのテキストは灰色、背景を白とする
            If (e.Index = 0) Then
                foreBrush = Brushes.Gray
                backBrush = Brushes.GhostWhite
            ElseIf (e.Index = 1) Then
                foreBrush = Brushes.Gray
                backBrush = Brushes.GhostWhite
            Else
                foreBrush = Brushes.Gray
                backBrush = Brushes.GhostWhite
            End If
        End If

        'StringFormatを作成
        Dim sf As New StringFormat
        '中央に表示する
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center

        '背景の描画
        e.Graphics.FillRectangle(backBrush, e.Bounds)
        'Textの描画
        e.Graphics.DrawString(txt, e.Font, foreBrush, RectangleF.op_Implicit(e.Bounds), sf)
    End Sub

    '*******************************************************
    '   INIファイル読込ボタン押下後処理
    '*******************************************************

    Private Sub btnIniGet_Click(sender As Object, e As EventArgs) Handles btnIniGet.Click
        On Error GoTo Error_Rtn

        Dim sFileName As String = txtIniPath.Text & "\" & pcIniFileName 'INIファイル名

        If fncFileCheck(sFileName) Then
            Call subSetIni()

            fncMsgBox("読込しました")
        Else
            fncErrors("INIファイルのパスを確認して下さい")
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   INIファイル読込ボタン押下後処理
    '*******************************************************

    Private Sub btnIniGet2_Click_1(sender As Object, e As EventArgs) Handles btnIniGet2.Click
        On Error GoTo Error_Rtn

        Dim sFileName As String = txtIniPath.Text & "\" & pcIniFileName 'INIファイル名

        If fncFileCheck(sFileName) Then
            Call subSetIni()

            fncMsgBox("読込しました")
        Else
            fncErrors("INIファイルのパスを確認して下さい")
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   INIファイル保存ボタン押下後処理
    '*******************************************************
    Private Sub btnIniSave2_Click_1(sender As Object, e As EventArgs) Handles btnIniSave2.Click
        On Error GoTo Error_Rtn

        Dim sFileName As String = txtIniPath.Text & "\" & pcIniFileName 'INIファイル名

        If fncFileCheck(sFileName) Then
            Call subPutIni()
            Call subSetIni()

            fncMsgBox("保存しました")
        Else
            fncErrors("INIファイルのパスを確認して下さい")
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   INIファイル保存ボタン押下後処理
    '*******************************************************
    Private Sub btnIniSave_Click(sender As Object, e As EventArgs) Handles btnIniSave.Click
        On Error GoTo Error_Rtn

        Dim sFileName As String = txtIniPath.Text & "\" & pcIniFileName 'INIファイル名

        If fncFileCheck(sFileName) Then
            Call subPutIni()
            Call subSetIni()

            fncMsgBox("保存しました")
        Else
            fncErrors("INIファイルのパスを確認して下さい")
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   メイン画面終了処理
    '*******************************************************
    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        On Error GoTo Error_Rtn

        Dim sExeFile As String = "" 'ワイヤレスインサイト
        Dim s3d_pot_test_plus As String = pc3d_pot_test_plus

        'ワイヤレスインサイト
        If pFileFlg = 0 Then
            sExeFile = pcDoNothing
        Else
            sExeFile = pcCalcProp
        End If

        'プロセス削除
        Call subKillProc(sExeFile)
        Call subKillProc(s3d_pot_test_plus)
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   起動確認
    '*******************************************************
    Private Sub btnKidou_Click(sender As Object, e As EventArgs) Handles btnKidou.Click
        On Error GoTo Error_Rtn

        Dim sFile As String
        Dim sFile2 As String

        If cmbFileFlg.Text = "0" Then
            sFile = txtExe1.Text
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                    System.Diagnostics.Process.Start(sFile)
            p.WaitForExit()
        Else
            sFile = txtExe2.Text
            sFile2 = txtExe22.Text
            'ファイルを開いて終了まで待機する
            Dim p2 As System.Diagnostics.Process =
                    System.Diagnostics.Process.Start(sFile, sFile2)
            p2.WaitForExit()
        End If

        MessageBox.Show("終了しました")
Error_Exit:
        Exit Sub
Error_Rtn:
        fncErrors("異常終了")
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   経路計算テスト
    '*******************************************************
    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        On Error GoTo Error_Rtn

        Dim result As DialogResult = MessageBox.Show("経路計算のテストを行いますか？",
                                             "経路計算テスト",
                                             MessageBoxButtons.YesNo,
                                             MessageBoxIcon.Exclamation,
                                             MessageBoxDefaultButton.Button2)
        If result = DialogResult.Yes Then
            '経路計算テスト
            Call subTest()

            Call fncMsgBox("終了")
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        fncErrors("異常終了")
        GoTo Error_Exit
    End Sub

    Private Sub btnTxrx_Click(sender As Object, e As EventArgs) Handles btnTxrx.Click
        On Error GoTo Error_Rtn

        Dim m As Integer = 2
        Dim n As Integer = 1

        pSet_d = 2

        ReDim Preserve pIti(pcMax_d, 1)     '位置情報
        ReDim Preserve pTag_route(pSet_d - 1)

        pIti(1, 0).sIdo = "11"
        pIti(1, 0).sKeido = "12"
        pIti(1, 0).sTakasa = "13"
        pIti(1, 1).sIdo = "14"
        pIti(1, 1).sKeido = "15"
        pIti(1, 1).sTakasa = "16"
        pIti(2, 0).sIdo = "21"
        pIti(2, 0).sKeido = "22"
        pIti(2, 0).sTakasa = "23"
        pIti(2, 1).sIdo = "24"
        pIti(2, 1).sKeido = "25"
        pIti(2, 1).sTakasa = "26"

        'txrxファイルを削除
        fncFileDelete_pj_txrx()

        'Iti用配列からtxrx用配列を作成する
        fncIti2Txrx(m, n)
        'タグデータを作成する
        fncTagCreate(m, n)
        'txrxファイルを作成する
        fncMem2txrx(m)

        Call fncMsgBox("作成しました")
Error_Exit:
        Exit Sub
Error_Rtn:
        fncErrors("異常終了")
        GoTo Error_Exit
    End Sub
End Class