'*******************************************************
'      制御用モジュール
'*******************************************************
Module mduProc

    ''' <summary>
    ''' Itiファイルを作成する
    ''' </summary>
    ''' <param name="m">ドローン</param>
    ''' <param name="n">周回</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    Public Function fncItiFile(ByVal m As Integer, ByVal n As Integer) As Boolean
        On Error GoTo Error_Rtn
        fncItiFile = False

        Dim sData As String = ""            'ファイル書込み用データ
        ReDim Preserve pIti(pcMax_d, n)     '位置情報

        'n=0の場合、画面設定値から取込む
        If n <> 0 Then
            '***************************************
            '***************************************
            '***************************************
            '経路計算(仮設定：値を固定値で設定する)
            '下記変数に値をセットする
            'd1
            pX_d1 = "0.11"
            pY_d1 = "0.12"
            pZ_d1 = "13"
            'd2
            pX_d2 = "0.21"
            pY_d2 = "0.22"
            pZ_d2 = "23"
            'd3
            pX_d3 = "0.31"
            pY_d3 = "0.32"
            pZ_d3 = "33"
            'd4
            pX_d4 = "0.41"
            pY_d4 = "0.42"
            pZ_d4 = "43"
            'd5
            pX_d5 = "0.51"
            pY_d5 = "0.52"
            pZ_d5 = "53"
            '***************************************
            '***************************************
            '***************************************
        End If

        'm=1～、n=0～
        Select Case m
            Case 1
                'm=1
                pIti(m, n).sIdo = pX_d1
                pIti(m, n).sKeido = pY_d1
                pIti(m, n).sTakasa = pZ_d1

                Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d1_iti.csv", True, System.Text.Encoding.UTF8)
                sData = pX_d1 & "," & pY_d1 & "," & pZ_d1
                oFileWrite.WriteLine(sData)
                'クローズ
                oFileWrite.Dispose()
                oFileWrite.Close()
            Case 2
                'm=2
                pIti(m, n).sIdo = pX_d2
                pIti(m, n).sKeido = pY_d2
                pIti(m, n).sTakasa = pZ_d2

                Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d2_iti.csv", True, System.Text.Encoding.UTF8)
                sData = pX_d2 & "," & pY_d2 & "," & pZ_d2
                oFileWrite.WriteLine(sData)
                'クローズ
                oFileWrite.Dispose()
                oFileWrite.Close()
            Case 3
                'm=3
                pIti(m, n).sIdo = pX_d3
                pIti(m, n).sKeido = pY_d3
                pIti(m, n).sTakasa = pZ_d3

                Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d3_iti.csv", True, System.Text.Encoding.UTF8)
                sData = pX_d3 & "," & pY_d3 & "," & pZ_d3
                oFileWrite.WriteLine(sData)
                'クローズ
                oFileWrite.Dispose()
                oFileWrite.Close()
            Case 4
                'm=4
                pIti(m, n).sIdo = pX_d4
                pIti(m, n).sKeido = pY_d4
                pIti(m, n).sTakasa = pZ_d4

                Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d4_iti.csv", True, System.Text.Encoding.UTF8)
                sData = pX_d4 & "," & pY_d4 & "," & pZ_d4
                oFileWrite.WriteLine(sData)
                'クローズ
                oFileWrite.Dispose()
                oFileWrite.Close()
            Case 5
                'm=5
                pIti(m, n).sIdo = pX_d5
                pIti(m, n).sKeido = pY_d5
                pIti(m, n).sTakasa = pZ_d5

                Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d5_iti.csv", True, System.Text.Encoding.UTF8)
                sData = pX_d5 & "," & pY_d5 & "," & pZ_d5
                oFileWrite.WriteLine(sData)
                'クローズ
                oFileWrite.Dispose()
                oFileWrite.Close()
        End Select

        fncItiFile = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function
End Module
