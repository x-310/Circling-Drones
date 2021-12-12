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

        Dim sSetup As String = pSetup & ".exe"
        Dim s3d_pot_test_plus As String = p3d_pot_test_plus & ".exe"

        If fncFileCheck(sSetup) = False Or fncFileCheck(s3d_pot_test_plus) = False Then
            fncMsgBox("実行ファイルを確認して下さい")
            End
        End If

        '画面コントロール初期化
        Call subCtlReset()

        'INIファイル読込み
        Call subSetIni()
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
        Call subKillProc(pSetup)

        Call subLogOutput("*** 前処理 ***")
        ReDim Preserve pOrder_d(-1)         'ドローン順番用配列
        ReDim Preserve pIti(pcMax_d, -1)     '位置情報
        ReDim Preserve pTxrx(-1)            'Txrx用配列
        ReDim pOkNg(-1)                     'OKorNG表示色設定
        pOkNg_No = 0                        'OKorNG表示No

        '前処理
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
                pOrder_d(iRow) = "d" & txtNow_m.Text & "v" & CInt(n).ToString("00")

                '周回処理
                Call subProc(CInt(txtNow_m.Text), CInt(txtNow_n.Text))

                '後処理
                Call subAfterProc(CInt(txtNow_m.Text), CInt(txtNow_n.Text))

                '緯度・経度・高さ情報を画面ログに出力する
                Call subLogOutput("*** 緯度,経度,高さ(pIti) ***")
                Call subLogOutput("d" & m & "," &
                                  "v" & n & "," &
                                  pIti(m, n).sIdo & "," &
                                  pIti(m, n).sKeido & "," &
                                  pIti(m, n).sTakasa)
                'Txrx情報を画面ログに出力する
                Call subLogOutput("*** ドローン情報,緯度,経度,高さ(pTxrx) ***")
                For iLoop = 0 To pTxrx.Length - 1
                    Call subLogOutput(pTxrx(iLoop).sdv & "," &
                                      pTxrx(iLoop).sIdo & "," &
                                      pTxrx(iLoop).sKeido & "," &
                                      pTxrx(iLoop).sTakasa)
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
        Call subLogOutput("エンド--> " + sStopTime)

        sTime = CLng(sw.ElapsedMilliseconds)
        Call subLogOutput("-------------------")
        Call subLogOutput("経過時間(ms)> " & sTime)
        Call subLogOutput("-------------------")

        'ドローン順を表示する
        Call subLogOutput("*** ドローン順番(pOrder_d) ***")
        For iLoop = 0 To pOrder_d.Length - 1
            Call subLogOutput(iLoop + 1 & "." & pOrder_d(iLoop))
        Next

        '画面ログ出力（0：緑色、1：赤色）
        Dim iCnt As Integer = 0
        For iLoop = 0 To pOkNg_No - 1
            Call subLogOutput_Color(pOkNg(iLoop).iLen, pOkNg(iLoop).iColor)
            If pOkNg(iLoop).iColor = 1 Then iCnt = iCnt + 1
        Next

        'NGカウント表示
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
            '選択されているタブのテキストを赤、背景を青とする
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
    '   タブコントロール2選択処理
    '*******************************************************
    Private Sub TabControl2_DrawItem(sender As Object, e As DrawItemEventArgs) Handles TabControl2.DrawItem
        '対象のTabControlを取得
        Dim tab As TabControl = CType(sender, TabControl)
        'タブページのテキストを取得
        Dim txt As String = tab.TabPages(e.Index).Text
        'タブのテキストと背景を描画するためのブラシを決定する
        Dim foreBrush, backBrush As Brush

        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            '選択されているタブのテキストを赤、背景を青とする
            foreBrush = Brushes.White
            backBrush = Brushes.DarkBlue
        Else
            '選択されていないタブのテキストは灰色、背景を白とする
            If (e.Index = 0) Then
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
    '   INIファイル読込ボタン押下後処理
    '*******************************************************
    Private Sub btnIniGet2_Click(sender As Object, e As EventArgs) Handles btnIniGet3.Click
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
    Private Sub btnIniSave2_Click(sender As Object, e As EventArgs) Handles btnIniSave3.Click
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

        'プロセス削除
        Call subKillProc(pSetup)
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '130.CSVファイルを配列にセットする
        Dim arrCsv()() As String = ReadCsv("130.1K2490GD2route_grid.csv", False, False)
        Dim iRow As Integer = 0

        ReDim Preserve p130(-1)
        If arrCsv.Length > 0 Then
            For iRow = 0 To arrCsv.Length - 1
                ReDim Preserve p130(iRow)       'p130 に新規行を追加
                If iRow = 0 Then
                    p130(iRow).sX = arrCsv(iRow)(0).ToString
                    p130(iRow).sY = arrCsv(iRow)(1).ToString
                    p130(iRow).sZ = arrCsv(iRow)(2).ToString
                Else
                    p130(iRow).sX = arrCsv(iRow)(0).ToString
                    p130(iRow).sY = arrCsv(iRow)(1).ToString
                    p130(iRow).sZ = arrCsv(iRow)(2).ToString
                    p130(iRow).sAA = arrCsv(iRow)(3).ToString
                End If
            Next
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim iRow As Integer = 0

        fncFile2Grid(pc130_Grid)


        Dim iA1 As Integer = 0
        Dim iA2 As Integer = 0
        Dim iB1 As Integer = 0
        Dim iB2 As Integer = 0
        Dim iC1 As Integer = 0
        Dim iC2 As Integer = 0

        For iRow = 0 To p130.Length - 1
            If iRow = 0 Then
                iA1 = CInt(p130(1).sX)
                iA2 = CInt(p130(0).sX)
                iB1 = CInt(p130(1).sY)
                iB2 = CInt(p130(0).sY)
                iC1 = CInt(p130(1).sZ)
                iC2 = CInt(p130(0).sZ)
            Else
                iA1 = CInt(p130(iRow).sX)
                iA2 = CInt(p130(iRow - 1).sX)
                iB1 = CInt(p130(iRow).sY)
                iB2 = CInt(p130(iRow - 1).sY)
                iC1 = CInt(p130(iRow).sZ)
                iC2 = CInt(p130(iRow - 1).sZ)
            End If
            p130(iRow).dCal1 = Math.Sqrt((iA2 - iA1) ^ 2 + (iB2 - iB1) ^ 2 + (iC2 - iC1) ^ 2)
            p130(iRow).dCal2 = CInt(txtVT.Text)
        Next
    End Sub
End Class