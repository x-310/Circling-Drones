'*******************************************************
'      制御用モジュール
'*******************************************************
Module mduProc

    '*******************************************************
    '   Itiファイルを作成する
    '*******************************************************
    Public Function fncItiFile(ByVal m As Integer, ByVal n As Integer) As Boolean
        On Error GoTo Error_Rtn
        fncItiFile = False

        Dim sData As String = ""
        ReDim Preserve pIti(pMax_d, n)      '位置情報

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

    '*******************************************************
    '   Itiファイルからtxrxファイルを作成する
    '*******************************************************
    Public Function fncItiFile2Txrx(ByVal m As Integer, ByVal n As Integer) As Boolean
        On Error GoTo Error_Rtn
        fncItiFile2Txrx = False

        Dim iLoop As Integer
        Dim jLoop As Integer
        Dim iRow As Integer = 0
        Dim id As Integer
        Dim iv As Integer

        ReDim Preserve pTxrx(-1) 'Txrx用配列
        If n = 1 Then
            For iLoop = 1 To pSet_d
                id = iLoop
                iv = 0
                If id <> m Then
                    ReDim Preserve pTxrx(iRow)             'pTxrx に新規行を追加

                    pTxrx(iRow).sdv = "d" & id & "v" & iv
                    pTxrx(iRow).sIdo = pIti(id, iv).sIdo
                    pTxrx(iRow).sKeido = pIti(id, iv).sKeido
                    pTxrx(iRow).sTakasa = pIti(id, iv).sTakasa

                    'txrxファイル作成
                    Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\" & pPjName & ".txrx", True, System.Text.Encoding.UTF8)
                    oFileWrite.WriteLine("d" & id & "v" & iv)
                    oFileWrite.WriteLine(pIti(id, iv).sIdo)
                    oFileWrite.WriteLine(pIti(id, iv).sKeido)
                    oFileWrite.WriteLine(pIti(id, iv).sTakasa)
                    'クローズ
                    oFileWrite.Dispose()
                    oFileWrite.Close()

                    ''dm_vn_txrxファイル作成
                    'Dim oFileWrite_2 As New System.IO.StreamWriter(pPjPath & "\" & "d" & m & "_v" & n & ".txrx", True, System.Text.Encoding.UTF8)
                    'oFileWrite_2.WriteLine("d" & id & "v" & iv)
                    'oFileWrite_2.WriteLine(pIti(id, iv).sIdo)
                    'oFileWrite_2.WriteLine(pIti(id, iv).sKeido)
                    'oFileWrite_2.WriteLine(pIti(id, iv).sTakasa)
                    ''クローズ
                    'oFileWrite_2.Dispose()
                    'oFileWrite_2.Close()

                    iRow = iRow + 1
                End If
            Next
        Else
            'ドローン最新順
            For iLoop = pOrder_d.Length - 1 To 1 Step -1
                id = pOrder_d(iLoop).Substring(1, 1)
                iv = pOrder_d(iLoop).Substring(3, 2)
                If id <> m Then
                    ReDim Preserve pTxrx(iRow)             'pTxrx に新規行を追加

                    pTxrx(iRow).sdv = "d" & id & "v" & iv
                    pTxrx(iRow).sIdo = pIti(id, iv).sIdo
                    pTxrx(iRow).sKeido = pIti(id, iv).sKeido
                    pTxrx(iRow).sTakasa = pIti(id, iv).sTakasa

                    'txrxファイル作成
                    Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\" & pPjName & ".txrx", True, System.Text.Encoding.UTF8)
                    oFileWrite.WriteLine("d" & id & "v" & iv)
                    oFileWrite.WriteLine(pIti(id, iv).sIdo)
                    oFileWrite.WriteLine(pIti(id, iv).sKeido)
                    oFileWrite.WriteLine(pIti(id, iv).sTakasa)
                    'クローズ
                    oFileWrite.Dispose()
                    oFileWrite.Close()

                    ''dm_vn_txrxファイル作成
                    'Dim oFileWrite_2 As New System.IO.StreamWriter(pPjPath & "\" & "d" & m & "_v" & n & ".txrx", True, System.Text.Encoding.UTF8)
                    'oFileWrite_2.WriteLine("d" & id & "v" & iv)
                    'oFileWrite_2.WriteLine(pIti(id, iv).sIdo)
                    'oFileWrite_2.WriteLine(pIti(id, iv).sKeido)
                    'oFileWrite_2.WriteLine(pIti(id, iv).sTakasa)
                    ''クローズ
                    'oFileWrite_2.Dispose()
                    'oFileWrite_2.Close()

                    If (pSet_d - 1) = (iRow + 1) Then Exit For
                    iRow = iRow + 1
                End If
            Next
        End If

        'コピー
        fncFileCopy(pPjPath & "\" & pPjName & ".txrx", pPjPath & "\d" & m & "_v" & n & "_" & pPjName & ".txrx")

        fncItiFile2Txrx = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function
End Module
