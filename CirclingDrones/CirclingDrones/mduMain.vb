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

        '周回用ファイルを削除する
        If fncFileDelete_d() Then
            Call subLogOutput("> " & "ドローン毎の周回用ファイルを削除=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "ドローン毎の周回用ファイルを削除=>NG")
            Call subOkNg_Color(1)
        End If

        'Itiファイルを作成する
        For m = 1 To pSet_d
            If fncItiFile(m, 0) Then
                Call subLogOutput("> " & "d" & m & "_Itiファイルの作成=>OK")
                Call subOkNg_Color(0)
            Else
                Call subLogOutput("> " & "d" & m & "_Itiファイルの作成=>OK")
                Call subOkNg_Color(1)
            End If
        Next

        'txrxファイルを削除する
        If fncFileDelete_txrx() Then
            Call subLogOutput("> " & "Txrxファイルを削除=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "Txrxファイルを削除=>NG")
            Call subOkNg_Color(1)
        End If
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

        Call subLogOutput("*** 周回処理(d" & m & "v" & n & ") ***")

        'プロジェクト名.setup 起動
        System.Diagnostics.Process.Start(sSetup, pPjName & ".setup 起動")
        Call subLogOutput("> " & pPjName & ".setup 起動")
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

        '3d_pot_test_plus.exe 起動
        System.Diagnostics.Process.Start(s3d_pot_test_plus)
        Call subLogOutput("> 3d_pot_test_plus.exe 起動")
        Call subOkNg_Color(2)
        Call subiInterval(500)

        '出力ファイル リネーム
        If fnc130Rename(m, n) Then
            Call subLogOutput("> 出力ファイル リネーム=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 出力ファイル リネーム=>NG")
            Call subOkNg_Color(1)
        End If

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

        Call subLogOutput("*** 後処理(d" & m & "v" & n & ") ***")

        'プロジェクト名.txrxファイルを削除
        If fncFileDelete_pj_txrx() Then
            Call subLogOutput("> " & "Txrxファイルを削除=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "Txrxファイルを削除=>NG")
            Call subOkNg_Color(1)
        End If

        '130.CSVファイルを配列にセット
        If fncFile2Grid(pc130_Grid) Then
            Call subLogOutput("> 130.CSVファイルを配列にセット=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 130.CSVファイルを配列にセット=>NG")
            Call subOkNg_Color(1)
        End If

        '経路計算する
        If fncItiCalc() Then
            Call subLogOutput("> 経路計算=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> 経路計算=>NG")
            Call subOkNg_Color(1)
        End If

        'new_itiファイルを削除
        If fncFileDel(pPjPath & "\new_iti.csv") Then
            Call subLogOutput("> " & "Txrxファイルを削除=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "Txrxファイルを削除=>NG")
            Call subOkNg_Color(1)
        End If

        'Itiファイルを作成する
        If fncItiFile(m, n) Then
            Call subLogOutput("> " & "d" & m & "_Itiファイルの作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> " & "d" & m & "_Itiファイルの作成=>NG")
            Call subOkNg_Color(1)
        End If

        'Itiファイルからtxrx用配列を作成
        If fncItiFile2Txrx(m, n) Then
            Call subLogOutput("> Itiファイルからtxrx用配列作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> Itiファイルからtxrx用配列作成=>NG")
            Call subOkNg_Color(1)
        End If

        'プロジェクト名.txrxファイルを作成
        If fncMem2txrx() Then
            Call subLogOutput("> txrxファイル 作成=>OK")
            Call subOkNg_Color(0)
        Else
            Call subLogOutput("> txrxファイル 作成=>NG")
            Call subOkNg_Color(1)
        End If
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub
End Module
