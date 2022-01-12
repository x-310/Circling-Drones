'*******************************************************
'      メインモジュール
'*******************************************************
Module mduMain

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 前処理
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subPreProc()
        On Error GoTo Error_Rtn

        Call subLogOutput("")
        Call subLogOutput("*** 前処理 ***")

        'Logファイルを削除
        Call subFileDel_w(pPjPath + "\*.log")
        'Csvファイルを削除
        Call subFileDel_w(pPjPath + "\*.csv")
        'Txrxファイルを削除
        Call subFileDel_w(pPjPath + "\*.txrx")
        'p2mファイルを削除
        'Call subFileDel_w(pPjPath + pP2mFile)
        Call subLogOutput("> " & "ファイル(*.log,.csv,.txrx)削除=>OK")
        Call subOkNg_Color(0)

        If frmMain.cmbDebug.Text = "ON" Then
            '***************************************************
            'DebugモードON
            'b2\studyarea\Debug\dm_vn\サブフォルダを削除
            Dim sDir As String = pPjPath + "\Debug"
            If fncDirDel(sDir) Then
                Call subLogOutput("> " & "Debugフォルダを削除=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> " & "Debugフォルダを削除=>NG")
                Call subOkNg_Color(1)
            End If
            '***************************************************
        End If

        ReDim Preserve p130(-1)
        ReDim Preserve pTag_route(pSet_d - 1)
        'Itiファイルを作成する
        For m = pSet_d To 1 Step -1
            Call subLogOutput("> [d" & m & "_v0]")
            '130.CSVファイルから130配列にセット
            If fncFile2Grid(pc130_Grid) Then
                'fncMsgBox("130xxx.CSVファイル取込み=>OK")
                Call subLogOutput("> 　" & "130.CSVファイルから130配列にセット=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> 　" & "130.CSVファイルから130配列にセット=>NG")
                Call subOkNg_Color(1)
            End If

            '130配列から飛距離計算
            If fnc130Calc() Then
                Call subLogOutput("> 　" & "130配列から飛距離計算=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> 　" & "130配列から飛距離計算=>NG")
                Call subOkNg_Color(1)
            End If

            '経路計算(n=0)
            If fncItiCalc(m, 0) Then
                Call subLogOutput("> 　経路計算(n=0)=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> 　経路計算(n=0)=>NG")
                Call subOkNg_Color(1)
            End If

            'Itiファイル(n=0)を作成
            If fncItiFile(m, 0) Then
                Call subLogOutput("> 　" & "Itiファイル(n=0)作成=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> 　" & "Itiファイル(n=0)作成=>NG")
                Call subOkNg_Color(1)
            End If

            'New_Itiファイルを作成する
            If fncNewItiFile(m, 0) Then
                Call subLogOutput("> 　" & "New_Itiファイル作成=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> 　" & "New_Itiファイル作成=>NG")
                Call subOkNg_Color(1)
            End If
        Next
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 周回処理
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subProc(ByVal m As Integer, ByVal n As Integer)
        On Error GoTo Error_Rtn

        Dim sExeFile As String = ""
        Dim s3d_pot_test_plus As String = pc3d_pot_test_plus & ".exe"

        Call subLogOutput("")
        Call subLogOutput("*** 周回処理 [d" & m & "_v" & n & "] ***")

        '***************************************************

        'DoNothing / CalcProp起動
        If pFileFlg = 0 Then
            sExeFile = frmMain.txtExe1.Text
        Else
            sExeFile = frmMain.txtExe2.Text
        End If

        Dim sFile As String
        Dim sFile2 As String
        Dim sFile3 As String
        Dim sDir1 As String
        Dim sDir2 As String

        If frmMain.cmbFileFlg.Text = "0" Then
            sFile = frmMain.txtExe1.Text
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                    System.Diagnostics.Process.Start(sFile)
            p.WaitForExit()
        Else
            sFile = frmMain.txtExe2.Text
            sFile2 = frmMain.txtExe22.Text
            'ファイルを開いて終了まで待機する
            Dim p2 As System.Diagnostics.Process =
                    System.Diagnostics.Process.Start(sFile, sFile2)
            p2.WaitForExit()
        End If

        Call subLogOutput("> DoNothing / CalcProp起動")
        Call subOkNg_Color(2)
        Call subiInterval(500)

        If frmMain.cmbDebug.Text = "ON" Then
            '***************************************************
            'DebugモードON
            'b2\studyareaフォルダをdm_vnサブフォルダにコピー
            sDir1 = pPjPath & "\studyarea"
            sDir2 = pPjPath & "\Debug\d" & m & "_v" & n
            If fncDirCopy(sDir1, sDir2) Then
                Call subLogOutput("> " & "Debugフォルダにコピー=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> " & "Debugフォルダにコピー=>NG")
                Call subOkNg_Color(1)
            End If
            '***************************************************
        End If

        Dim sFileName() As String
        Dim iLoop As Integer
        Dim iRowCnt As Integer = 0

        ReDim Preserve pP2m(-1)

        'P2mファイル数確認
        If fncGetDir_P2m(pPjPath + pP2mDir) = pcP2mFileCnt Then
            'P2mファイル名取得
            sFileName = fncGetDirName_P2m(pPjPath + pP2mDir)
            'P2mファイルをソート
            sFileName = fncFileSort(sFileName)
        Else
            fncMsgBox("P2mファイル数エラー")
            End
        End If

        For iLoop = 0 To sFileName.Length - 1
            'P2mファイルを配列にセット
            iRowCnt = fncFile2P2m(iRowCnt, sFileName(iLoop))
            If iRowCnt <> 0 Then
                Call subLogOutput("> " & pPjName & "P2mファイルを配列にセット:" & iLoop & "=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> " & pPjName & "P2mファイルを配列にセット=>NG")
                Call subOkNg_Color(1)
            End If

            Application.DoEvents()
        Next

        'power.txt作成
        If fncMem2Power() Then
            Call subLogOutput("> power.txt作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> power.txt作成=>NG")
            Call subOkNg_Color(1)
        End If

        '***************************************************

        '3d_pot_test_plus.exe 起動
        sFile3 = pc3d_pot_test_plus & ".exe"
        'ファイルを開いて終了まで待機する
        Dim p3 As System.Diagnostics.Process =
                    System.Diagnostics.Process.Start(sFile3)
        p3.WaitForExit()

        Call subLogOutput("> " & pc3d_pot_test_plus & ".exe起動")
        Call subOkNg_Color(2)
        Call subiInterval(500)

        '履歴用130ファイルにコピー
        If fncCopy130(m, n) Then
            Call subLogOutput("> 履歴用130ファイルにコピー=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 履歴用130ファイルにコピー=>NG")
            Call subOkNg_Color(1)
        End If

        '***************************************************

        '周回の間隔
        Call subLogOutput("> 周回の間隔")
        Call subiInterval(pSet_Interval)
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 後処理
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subAfterProc(ByVal m As Integer, ByVal n As Integer)
        On Error GoTo Error_Rtn

        Call subLogOutput("")
        Call subLogOutput("*** 後処理 [d" & m & "_v" & n & "] ***")

        '***************************************************

        'txrxファイルを削除
        If fncFileDelete_pj_txrx() Then
            Call subLogOutput("> " & "Txrxファイル削除=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "Txrxファイル削除=>NG")
            Call subOkNg_Color(1)
        End If

        'New_Itiファイルを削除
        If fncFileDel(pExePath & "\New_Iti.csv") Then
            Call subLogOutput("> " & "New_Itiファイル削除=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "New_Itiファイル削除=>NG")
            Call subOkNg_Color(1)
        End If

        '***************************************************

        '130ファイルを配列にセット
        If fncFile2Grid(pc130_Grid) Then
            Call subLogOutput("> 130ファイルを配列にセット=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 130ファイルを配列にセット=>NG")
            Call subOkNg_Color(1)
        End If

        '130配列から飛距離計算する
        If fnc130Calc() Then
            Call subLogOutput("> 130配列から飛距離計算=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 130配列から飛距離計算=>NG")
            Call subOkNg_Color(1)
        End If

        '経路計算する
        If fncItiCalc(m, n) Then
            Call subLogOutput("> 経路計算=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 経路計算=>NG")
            Call subOkNg_Color(1)
        End If

        'Itiファイルを作成する
        If fncItiFile(m, n) Then
            Call subLogOutput("> " & "d" & m & "_Itiファイル作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "d" & m & "_Itiファイル作成=>NG")
            Call subOkNg_Color(1)
        End If

        'New_Itiファイルを作成する
        If fncNewItiFile(m, n) Then
            Call subLogOutput("> " & "New_Itiファイル作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "New_Itiファイル作成=>NG")
            Call subOkNg_Color(1)
        End If

        '***************************************************

        'Iti用配列からtxrx用配列を作成する
        If fncIti2Txrx(m, n) Then
            Call subLogOutput("> Iti用配列からtxrx用配列作成する=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> Iti用配列からtxrx用配列作成する=>NG")
            Call subOkNg_Color(1)
        End If

        'タグデータを作成する
        If fncTagCreate(m, n) Then
            Call subLogOutput("> タグデータ作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> タグデータ作成=>NG")
            Call subOkNg_Color(1)
        End If

        'txrxファイルを作成する
        If fncMem2txrx(m) Then
            Call subLogOutput("> txrxファイル作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> txrxファイル作成=>NG")
            Call subOkNg_Color(1)
        End If

        ''txrxファイルを作成する(テスト用)
        'If fncMem2txrx_test() Then
        '    Call subLogOutput("> txrxファイル作成(テスト用)=>OK")
        '    Call subOkNg_Color(0)
        'Else
        '    Call subLogOutput("> txrxファイル作成(テスト用)=>NG")
        '    Call subOkNg_Color(1)
        'End If

        '履歴用Txrxファイルにコピーする
        If fncCopyTxrx(m, n) Then
            Call subLogOutput("> 履歴用Txrxファイルにコピー=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 履歴用Txrxファイルにコピー=>NG")
            Call subOkNg_Color(1)
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Logファイル出力処理
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subLogFile(ByVal m As Integer, ByVal n As Integer)
        On Error GoTo Error_Rtn

        Call subLogOutput("")
        Call subLogOutput("*** Logファイル出力処理 ***")

        '***************************************************

        Dim sFile As String = "d" & m & "_v" & n & "_p130.log"
        Call subLogOutput("> " & sFile)

        Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\" & sFile, True, System.Text.Encoding.UTF8)
        Dim sData As String
        Dim iLoop As Integer

        ''経路計算情報用配列
        For iLoop = 0 To p130.Length - 1
            sData = p130(iLoop).sX & "," &
                    p130(iLoop).sY & "," &
                    p130(iLoop).sZ & "," &
                    p130(iLoop).sAA & "," &
                    p130(iLoop).dCal1 & "," &
                    p130(iLoop).dCal2 & "," &
                    p130(iLoop).iX & "," &
                    p130(iLoop).iY & "," &
                    p130(iLoop).iZ
            oFileWrite.WriteLine(sData)
        Next

        'クローズ
        oFileWrite.Dispose()
        oFileWrite.Close()
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub
End Module
