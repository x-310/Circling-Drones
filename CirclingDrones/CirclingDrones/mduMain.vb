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

        Call subLogOutput("")
        Call subLogOutput("*** 前処理 ***")

        'Csvファイルを削除
        Call subFileDel_w(pPjPath + "\*.csv")
        'Txrxファイルを削除
        Call subFileDel_w(pPjPath + "\*.txrx")
        'p2mファイルを削除
        'Call subFileDel_w(pPjPath + pP2mFile)
        Call subLogOutput("> " & "ファイル(.csv,.txrx)削除=>OK")
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
        ReDim Preserve p130_2(-1)

        'Itiファイルを作成する
        For m = 1 To pSet_d
            Call subLogOutput("> [d" & m & "_v0]")

            'Itiファイル(n=0)を作成
            ReDim Preserve pIti(pcMax_d, 1)   '位置情報
            ReDim Preserve p130(0)
            ReDim Preserve p130_2(0)

            Dim sX As String = ""
            Dim sY As String = ""
            Dim sZ As String = ""

            Select Case m
                Case 1
                    sX = frmMain.txtX_d1.Text
                    sY = frmMain.txtY_d1.Text
                    sZ = frmMain.txtZ_d1.Text
                Case 2
                    sX = frmMain.txtX_d2.Text
                    sY = frmMain.txtY_d2.Text
                    sZ = frmMain.txtZ_d2.Text
                Case 3
                    sX = frmMain.txtX_d3.Text
                    sY = frmMain.txtY_d3.Text
                    sZ = frmMain.txtZ_d3.Text
                Case 4
                    sX = frmMain.txtX_d4.Text
                    sY = frmMain.txtY_d4.Text
                    sZ = frmMain.txtZ_d4.Text
                Case 5
                    sX = frmMain.txtX_d5.Text
                    sY = frmMain.txtY_d5.Text
                    sZ = frmMain.txtZ_d5.Text
            End Select

            pIti(m, 0).sIdo = sX
            pIti(m, 0).sKeido = sY
            pIti(m, 0).sTakasa = sZ

            pIti(m, 1).sIdo = sX
            pIti(m, 1).sKeido = sY
            pIti(m, 1).sTakasa = sZ

            If fncItiFile(m, 0) Then
                Call subLogOutput("> 　" & "Itiファイル(n=0)作成=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> 　" & "Itiファイル(n=0)作成=>NG")
                Call subOkNg_Color(1)
            End If

            '130配列を作成
            p130(0).dCal1 = 0.0
            p130(0).dCal2 = CDbl(frmMain.txtV.Text)            'p130(0).sX = sX
            p130(0).sY = sY
            p130(0).sZ = sZ

            p130_2(0).dCal1 = 0.0
            p130_2(0).dCal2 = CDbl(frmMain.txtV.Text)            'p130(0).sX = sX
            p130_2(0).sX = sX
            p130_2(0).sY = sY
            p130_2(0).sZ = sZ
        Next

        'New_Itiファイルを作成する
        If fncNewItiFile(1, 0) Then
            Call subLogOutput("> 　" & "New_Itiファイル作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 　" & "New_Itiファイル作成=>NG")
            Call subOkNg_Color(1)
        End If

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

        Dim sFileName() As String
        Dim iLoop As Integer
        Dim iRowCnt As Integer = 0

        ReDim Preserve pTag_route(pSet_d - 1)
        ReDim Preserve pP2m(-1)

        If frmMain.cmbPowerOff.Text = "" Then
            'P2mファイル数確認
            If fncGetDir_P2m(pPjPath + pP2mDir) >= pcP2mFileCnt Then
                'P2mファイル名取得
                sFileName = fncGetDirName_P2m(pPjPath + pP2mDir)
                If sFileName.Length = pcP2mFileCnt Then
                    'P2mファイルをソート
                    sFileName = fncFileSort(sFileName)
                Else
                    fncMsgBox("P2mファイル数エラー")
                    End
                End If
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
        End If

        Application.DoEvents()

        '***************************************************

        '3d_pot_test_plus.exe 起動
        sFile3 = pc3d_pot_test_plus & ".exe"
        'ファイルを開いて終了まで待機する
        Dim p3 As System.Diagnostics.Process =
                    System.Diagnostics.Process.Start(sFile3)
        p3.WaitForExit()

        Call subLogOutput("> " & pc3d_pot_test_plus & ".exe起動")
        Call subOkNg_Color(2)

        '履歴用130ファイルにコピー
        If fncCopy130(m, n) Then
            Call subLogOutput("> 履歴用130ファイルにコピー=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 履歴用130ファイルにコピー=>NG")
            Call subOkNg_Color(1)
        End If

        Application.DoEvents()

        '***************************************************

        '周回の間隔
        Call subLogOutput("> 周回の間隔")
        Call subiInterval(pSet_Interval)

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
        Dim sDir1 As String
        Dim sDir2 As String

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
        If fncCalc130() Then
            Call subLogOutput("> 130配列から飛距離計算=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 130配列から飛距離計算=>NG")
            Call subOkNg_Color(1)
        End If

        Application.DoEvents()

        '経路計算する
        If fnc130Calc(m, n) Then
            Call subLogOutput("> 経路計算=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 経路計算=>NG")
            Call subOkNg_Color(1)
        End If

        Application.DoEvents()
        '************
        'tが倍用
        '130.CSVファイルから130配列にセット
        If fncFile2Grid_2() Then
            Call subLogOutput("> 　" & "130.CSVから130配列にセット(tが倍)=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 　" & "130.CSVから130配列にセット(tが倍)=>NG")
            Call subOkNg_Color(1)
        End If

        Application.DoEvents()

        '130配列から飛距離計算
        If fncCalc130_2() Then
            Call subLogOutput("> 　" & "130配列から飛距離計算(tが倍)=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 　" & "130配列から飛距離計算(tが倍)=>NG")
            Call subOkNg_Color(1)
        End If

        Application.DoEvents()

        '経路計算
        If fnc130Calc_2(m, n) Then
            Call subLogOutput("> 　経路計算 n=0 (tが倍)=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 　経路計算 n=0 (tが倍)=>NG")
            Call subOkNg_Color(1)
        End If

        Application.DoEvents()
        '************

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

        If frmMain.cmbDebug.Text = "ON" Then
            '***************************************************
            'DebugモードON
            'b2\studyareaフォルダをdm_vnサブフォルダにコピー
            sDir1 = pPjPath & "\studyarea"
            sDir2 = pPjPath & "\Debug\d" & m & "_v" & n
            If fncDirCopy(sDir1, sDir2) Then
                Call subLogOutput("> " & "studyareaﾌｫﾙﾀﾞをDebugﾌｫﾙﾀﾞにｺﾋﾟｰ=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> " & "studyareaﾌｫﾙﾀﾞをDebugﾌｫﾙﾀﾞにｺﾋﾟｰ=>NG")
                Call subOkNg_Color(1)
            End If

            'Prgフォルダをdm_vnサブフォルダにコピー
            sDir1 = pExePath
            sDir2 = pPjPath & "\Debug\d" & m & "_v" & n
            If fncDirCopy(sDir1, sDir2) Then
                Call subLogOutput("> " & "exeフォルダをDebugﾌｫﾙﾀﾞにｺﾋﾟ=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> " & "exeフォルダをDebugﾌｫﾙﾀﾞにｺﾋﾟ=>NG")
                Call subOkNg_Color(1)
            End If
            '***************************************************
        End If

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

        Call subLogOutput("")
        Call subLogOutput("*** Logファイル出力処理 ***")

        '***************************************************

        Dim sFile As String = "d" & m & "_v" & n & "_p130_log.csv"
        Call subLogOutput("> " & sFile)

        'Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\" & sFile, True, System.Text.Encoding.UTF8)
        Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\" & sFile, True, pEnc)
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

        '****************************

        Dim sFile2 As String = "d" & m & "_v" & n & "_p130_2_log.csv"
        Call subLogOutput("> " & sFile2)

        'Dim oFileWrite2 As New System.IO.StreamWriter(pPjPath & "\" & sFile2, True, System.Text.Encoding.UTF8)
        Dim oFileWrite2 As New System.IO.StreamWriter(pPjPath & "\" & sFile2, True, pEnc)

        sData = ""
        ''経路計算情報用配列
        For iLoop = 0 To p130_2.Length - 1
            sData = p130_2(iLoop).sX & "," &
                    p130_2(iLoop).sY & "," &
                    p130_2(iLoop).sZ & "," &
                    p130_2(iLoop).sAA & "," &
                    p130_2(iLoop).dCal1 & "," &
                    p130_2(iLoop).dCal2 & "," &
                    p130_2(iLoop).iX & "," &
                    p130_2(iLoop).iY & "," &
                    p130_2(iLoop).iZ
            oFileWrite2.WriteLine(sData)
        Next

        'クローズ
        oFileWrite2.Dispose()
        oFileWrite2.Close()

    End Sub
End Module
