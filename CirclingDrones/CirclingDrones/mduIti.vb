'*******************************************************
'      経路計算用モジュール
'*******************************************************
Module mduIti

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Itiファイルを作成する
    ''' </summary>
    ''' <param name="m">ドローン</param>
    ''' <param name="n">周回</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncItiFile(ByVal m As Integer, ByVal n As Integer) As Boolean
        On Error GoTo Error_Rtn
        fncItiFile = False

        Dim sData As String = ""            'ファイル書込み用データ
        Dim iA1 As Integer = 0
        Dim iA2 As Integer = 0
        Dim iB1 As Integer = 0
        Dim iB2 As Integer = 0
        Dim iC1 As Integer = 0
        Dim iC2 As Integer = 0
        Dim dE As Double = 0.0
        Dim dF As Double = 0.0

        ReDim Preserve pIti(pcMax_d, n)     '位置情報

        'n=0の場合、画面設定値から取込む
        If n <> 0 Then
            iA1 = CInt(p130(n - 1).sX)
            iA2 = CInt(p130(n).sX)
            iB1 = CInt(p130(n - 1).sY)
            iB2 = CInt(p130(n).sY)
            iC1 = CInt(p130(n - 1).sZ)
            iC2 = CInt(p130(n).sZ)
            dE = p130(n).dCal1
            dF = p130(n - 1).dCal2
            Select Case m
                Case 1
                    If dE > dF Then
                        pX_d1 = dF / dE * (iA2 - iA1) + iA1
                        pY_d1 = dF / dE * (iB2 - iB1) + iB1
                        pZ_d1 = dF / dE * (iC2 - iC1) + iC1
                    End If
                Case 2
                    If dE > dF Then
                        pX_d2 = dF / dE * (iA2 - iA1) + iA1
                        pY_d2 = dF / dE * (iB2 - iB1) + iB1
                        pZ_d2 = dF / dE * (iC2 - iC1) + iC1
                    End If
                Case 3
                    If dE > dF Then
                        pX_d3 = dF / dE * (iA2 - iA1) + iA1
                        pY_d3 = dF / dE * (iB2 - iB1) + iB1
                        pZ_d3 = dF / dE * (iC2 - iC1) + iC1
                    End If
                Case 4
                    If dE > dF Then
                        pX_d4 = dF / dE * (iA2 - iA1) + iA1
                        pY_d4 = dF / dE * (iB2 - iB1) + iB1
                        pZ_d4 = dF / dE * (iC2 - iC1) + iC1
                    End If
                Case 5
                    If dE > dF Then
                        pX_d5 = dF / dE * (iA2 - iA1) + iA1
                        pY_d5 = dF / dE * (iB2 - iB1) + iB1
                        pZ_d5 = dF / dE * (iC2 - iC1) + iC1
                    End If
            End Select
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

                Dim oFileWrite2 As New System.IO.StreamWriter(pPjPath & "\new_iti.csv", True, System.Text.Encoding.UTF8)
                sData = pX_d1 & "," & pY_d1 & "," & pZ_d1
                oFileWrite2.WriteLine(sData)
                'クローズ
                oFileWrite2.Dispose()
                oFileWrite2.Close()
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

                Dim oFileWrite2 As New System.IO.StreamWriter(pPjPath & "\new_iti.csv", True, System.Text.Encoding.UTF8)
                sData = pX_d2 & "," & pY_d2 & "," & pZ_d2
                oFileWrite2.WriteLine(sData)
                'クローズ
                oFileWrite2.Dispose()
                oFileWrite2.Close()
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

                Dim oFileWrite2 As New System.IO.StreamWriter(pPjPath & "\new_iti.csv", True, System.Text.Encoding.UTF8)
                sData = pX_d3 & "," & pY_d3 & "," & pZ_d3
                oFileWrite2.WriteLine(sData)
                'クローズ
                oFileWrite2.Dispose()
                oFileWrite2.Close()
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

                Dim oFileWrite2 As New System.IO.StreamWriter(pPjPath & "\new_iti.csv", True, System.Text.Encoding.UTF8)
                sData = pX_d4 & "," & pY_d4 & "," & pZ_d4
                oFileWrite2.WriteLine(sData)
                'クローズ
                oFileWrite2.Dispose()
                oFileWrite2.Close()
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

                Dim oFileWrite2 As New System.IO.StreamWriter(pPjPath & "\new_iti.csv", True, System.Text.Encoding.UTF8)
                sData = pX_d5 & "," & pY_d5 & "," & pZ_d5
                oFileWrite2.WriteLine(sData)
                'クローズ
                oFileWrite2.Dispose()
                oFileWrite2.Close()
        End Select

        fncItiFile = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 経路計算する
    ''' </summary>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncItiCalc() As Boolean
        On Error GoTo Error_Rtn
        fncItiCalc = False

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
            p130(iRow).dCal2 = CInt(frmMain.txtVT.Text)
        Next

        fncItiCalc = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 130.CSVファイルを配列にセットする
    ''' </summary>
    ''' <param name="sFileName">130Gridファイルのパス</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFile2Grid(ByVal sFileName As String) As Boolean
        On Error GoTo Error_Rtn
        fncFile2Grid = False

        '130.CSVファイルを配列にセットする
        Dim arrCsv()() As String = ReadCsv(sFileName, False, False)
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

        fncFile2Grid = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function
End Module
