'*******************************************************
'      単体テスト用モジュール
'*******************************************************
Module mduTest

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 経路計算テスト
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subTest()

        Dim m As Integer = 1
        Dim n As Integer = 1
        pSet_d = 1

        '周回用ファイルを削除する
        Call subFileDel_w(pPjPath & "\d*.*")

        '***************************************************
        'New_Itiファイルを削除
        If fncFileDel(pPjPath & "\New_Iti.csv") Then
            'fncMsgBox("New_Itiファイル削除=>OK")
        Else
            fncErrors("New_Itiファイル削除=>NG")
        End If
        '***************************************************
        '経路計算(n=0)する
        If fncItiCalc(m, 0) Then
            'fncMsgBox("経路計算=>OK")
        Else
            fncErrors("経路計算=>NG")
        End If

        '130.CSVファイルから130配列にセット
        If fncFile2Grid(pc130_Grid) Then
            'fncMsgBox("130xxx.CSVファイル取込み=>OK")
        Else
            fncErrors("130xxx.CSVファイル取込み=>NG")
        End If

        '130配列から飛距離計算する
        If fnc130Calc() Then
            'fncMsgBox("130xxx.CSVファイル取込み=>OK")
        Else
            fncErrors("130xxx.CSVファイル取込み=>NG")
        End If

        'Itiファイル(n=0)を作成する
        If fncItiFile(m, 0) Then
            'fncMsgBox("d" & m & "_Itiファイル作成=>OK")
        Else
            fncErrors("d" & m & "_Itiファイル作成=>NG")
        End If

        '***************************************************
        '経路計算する
        If fncItiCalc(m, n) Then
            'fncMsgBox("経路計算=>OK")
        Else
            fncErrors("経路計算=>NG")
        End If

        '130.CSVファイルから130配列にセット
        If fncFile2Grid(pc130_Grid) Then
            'fncMsgBox("130xxx.CSVファイル取込み=>OK")
        Else
            fncErrors("130xxx.CSVファイル取込み=>NG")
        End If

        '130配列から飛距離計算する
        If fnc130Calc() Then
            'fncMsgBox("130xxx.CSVファイル取込み=>OK")
        Else
            fncErrors("130xxx.CSVファイル取込み=>NG")
        End If

        'Itiファイルを作成する
        If fncItiFile(m, n) Then
            'fncMsgBox("d" & m & "_Itiファイル作成=>OK")
        Else
            fncErrors("d" & m & "_Itiファイル作成=>NG")
        End If

        'New_Itiファイルを作成する
        If fncNewItiFile(m, n) Then
            'fncMsgBox("New_Itiファイル作成=>OK")
        Else
            fncErrors("New_Itiファイル作成=>NG")
        End If
    End Sub
End Module
