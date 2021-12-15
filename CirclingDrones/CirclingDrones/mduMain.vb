﻿'*******************************************************
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

        'Logファイルを削除する
        Call subFileDel_w(pPjPath & "\*.log")
        Call subLogOutput("> " & "Logファイル(*.log)削除=>OK")
        Call subOkNg_Color(0)

        '周回用ファイルを削除する
        Call subFileDel_w(pPjPath & "\d*.*")
        Call subLogOutput("> " & "周回用ファイル(d*.*)削除=>OK")
        Call subOkNg_Color(0)

        'New_Itiファイルを削除
        If fncFileDel(pPjPath & "\New_Iti.csv") Then
            'fncMsgBox("New_Itiファイル削除=>OK")
        Else
            fncErrors("New_Itiファイル削除=>NG")
        End If
        '***************************************************

        'Itiファイルを作成する
        For m = 1 To pSet_d
            '経路計算(n=0)
            If fncItiCalc(m, 0) Then
                Call subLogOutput("> 経路計算(n=0)=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> 経路計算(n=0)=>NG")
                Call subOkNg_Color(1)
            End If

            '130.CSVファイルから130配列にセット
            If fncFile2Grid(pc130_Grid) Then
                'fncMsgBox("130xxx.CSVファイル取込み=>OK")
                Call subLogOutput("> " & "130.CSVファイルから130配列にセット=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> " & "130.CSVファイルから130配列にセット=>NG")
                Call subOkNg_Color(1)
            End If

            '130配列から飛距離計算
            If fnc130Calc() Then
                Call subLogOutput("> " & "130配列から飛距離計算=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> " & "130配列から飛距離計算=>NG")
                Call subOkNg_Color(1)
            End If

            'Itiファイル(n=0)を作成
            If fncItiFile(m, 0) Then
                Call subLogOutput("> " & "Itiファイル(n=0)作成=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> " & "Itiファイル(n=0)作成=>NG")
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

        Dim sSetup As String = pSetup & ".exe"                          'setupファイル
        Dim s3d_pot_test_plus As String = p3d_pot_test_plus & ".exe"    'p3d_pot_test_plusファイル

        Call subLogOutput("")
        Call subLogOutput("*** 周回処理 [d" & m & "_v" & n & "] ***")

        '***************************************************

        'プロジェクト名.setup 起動
        System.Diagnostics.Process.Start(sSetup, pPjName & ".setup起動")
        Call subLogOutput("> " & pPjName & ".setup起動")
        Call subOkNg_Color(2)
        Call subiInterval(500)

        'プロジェクト名.p2m取込み
        If fncFile2P2m(pPjPath + "\" & pPjName & ".p2m") Then
            Call subLogOutput("> " & pPjName & ".p2mファイル取込み=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & pPjName & ".p2mファイル取込み=>NG")
            Call subOkNg_Color(1)
        End If

        'power.txt作成
        If fncMem2Power(pPjPath) Then
            Call subLogOutput("> power.txt作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> power.txt作成=>NG")
            Call subOkNg_Color(1)
        End If

        '***************************************************

        '3d_pot_test_plus.exe 起動
        System.Diagnostics.Process.Start(s3d_pot_test_plus)
        Call subLogOutput("> 3d_pot_test_plus.exe起動")
        Call subOkNg_Color(2)
        Call subiInterval(500)

        '出力ファイル リネーム
        If fnc130Rename(m, n) Then
            Call subLogOutput("> 出力ファイル・リネーム=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 出力ファイル・リネーム=>NG")
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

        'プロジェクト名.txrxファイルを削除
        If fncFileDelete_pj_txrx() Then
            Call subLogOutput("> " & "Txrxファイル削除=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "Txrxファイル削除=>NG")
            Call subOkNg_Color(1)
        End If

        'New_Itiファイルを削除
        If fncFileDel(pPjPath & "\New_Iti.csv") Then
            Call subLogOutput("> " & "New_Itiファイル削除=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "New_Itiファイル削除=>NG")
            Call subOkNg_Color(1)
        End If

        '***************************************************

        '経路計算する
        If fncItiCalc(m, n) Then
            Call subLogOutput("> 経路計算=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 経路計算=>NG")
            Call subOkNg_Color(1)
        End If

        '130.CSVファイルを配列にセット
        If fncFile2Grid(pc130_Grid) Then
            Call subLogOutput("> 130xxx.CSVファイル取込み=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 130xxx.CSVファイル取込み=>NG")
            Call subOkNg_Color(1)
        End If

        '130配列から飛距離計算する
        If fnc130Calc() Then
            Call subLogOutput("> 130xxx.CSVファイル取込み=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 130xxx.CSVファイル取込み=>NG")
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

        'タグデータを(仮)作成する
        If fncTagCreate() Then
            Call subLogOutput("> タグデータ(仮)作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> タグデータ(仮)作成=>NG")
            Call subOkNg_Color(1)
        End If

        'プロジェクト名.txrxファイルを作成する
        If fncMem2txrx() Then
            Call subLogOutput(">プロジェクト名.txrxファイル作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> プロジェクト名.txrxファイル作成=>NG")
            Call subOkNg_Color(1)
        End If

        'プロジェクト名.Txrxファイルを周回用.Txrxファイルにコピーする
        If fncCopyTxrx(m, n) Then
            Call subLogOutput("> 周回用.Txrxファイルにコピー=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 周回用.Txrxファイルにコピー=>NG")
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
                    p130(iLoop).dCal2
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
