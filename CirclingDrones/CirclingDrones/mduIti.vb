'*******************************************************
'      経路計算用モジュール
'*******************************************************
Module mduIti

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 130配列から飛距離計算する
    ''' </summary>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fnc130Calc() As Boolean
        On Error GoTo Error_Rtn
        fnc130Calc = False

        Dim iA1 As Integer = 0
        Dim iA2 As Integer = 0
        Dim iB1 As Integer = 0
        Dim iB2 As Integer = 0
        Dim iC1 As Integer = 0
        Dim iC2 As Integer = 0
        Dim dE As Double = 0.0
        Dim dF As Double = 0.0

        For iRow = 0 To p130.Length - 1
            If iRow = 0 Then
                iA1 = CInt(p130(0).sX)
                iA2 = CInt(p130(0).sX)
                iB1 = CInt(p130(0).sY)
                iB2 = CInt(p130(0).sY)
                iC1 = CInt(p130(0).sZ)
                iC2 = CInt(p130(0).sZ)
                dE = 0.0
                dF = CInt(frmMain.txtVT.Text)
            Else
                iA1 = CInt(p130(iRow - 1).sX)
                iA2 = CInt(p130(iRow).sX)
                iB1 = CInt(p130(iRow - 1).sY)
                iB2 = CInt(p130(iRow).sY)
                iC1 = CInt(p130(iRow - 1).sZ)
                iC2 = CInt(p130(iRow).sZ)
                dE = Math.Sqrt((iA2 - iA1) ^ 2 + (iB2 - iB1) ^ 2 + (iC2 - iC1) ^ 2)
                dF = p130(iRow - 1).dCal2 - dE
            End If

            p130(iRow).dCal1 = dE
            p130(iRow).dCal2 = dF
        Next

        fnc130Calc = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 経路計算する
    ''' </summary>
    ''' <param name="m">ドローン</param>
    ''' <param name="n">周回</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncItiCalc(ByVal m As Integer, ByVal n As Integer) As Boolean
        On Error GoTo Error_Rtn
        fncItiCalc = False

        Dim iRow As Integer = 0
        Dim iA1, iA2 As Integer
        Dim iB1, iB2 As Integer
        Dim iC1, iC2 As Integer
        Dim dE, dF As Double
        Dim dX, dY, dZ As Double
        Dim iX, iY, iZ As Integer

        iX = 0 : iY = 0 : iZ = 0
        dX = 0.0 : dY = 0.0 : dZ = 0.0

        pX_d1 = 0 : pY_d1 = 0 : pZ_d1 = 0
        pX_d2 = 0 : pY_d2 = 0 : pZ_d2 = 0
        pX_d3 = 0 : pY_d3 = 0 : pZ_d3 = 0
        pX_d4 = 0 : pY_d4 = 0 : pZ_d4 = 0
        pX_d5 = 0 : pY_d5 = 0 : pZ_d5 = 0
        'n=0の場合、画面設定値から取込む
        If n = 0 Then
            ReDim Preserve p130(0)

            p130(0).dCal1 = 0.0
            p130(0).dCal2 = CDbl(frmMain.txtV.Text)
            p130(0).iX = 0
            p130(0).iY = 0
            p130(0).iZ = 0
        Else
            For iRow = 1 To p130.Length - 1
                iA1 = CInt(p130(iRow - 1).sX)
                iA2 = CInt(p130(iRow).sX)
                iB1 = CInt(p130(iRow - 1).sY)
                iB2 = CInt(p130(iRow).sY)
                iC1 = CInt(p130(iRow - 1).sZ)
                iC2 = CInt(p130(iRow).sZ)
                dE = p130(iRow).dCal1
                dF = p130(iRow - 1).dCal2

                dX = dF / dE * (iA2 - iA1) + iA1
                dY = dF / dE * (iB2 - iB1) + iB1
                dZ = dF / dE * (iC2 - iC1) + iC1

                '四捨五入
                iX = Math.Round(dX, MidpointRounding.AwayFromZero)
                iY = Math.Round(dY, MidpointRounding.AwayFromZero)
                iZ = Math.Round(dZ, MidpointRounding.AwayFromZero)

                p130(iRow).iX = iX
                p130(iRow).iY = iY
                p130(iRow).iZ = iZ

                Select Case m
                    Case 1
                        pX_d1 = iX
                        pY_d1 = iY
                        pZ_d1 = iZ
                    Case 2
                        pX_d2 = iX
                        pY_d2 = iY
                        pZ_d2 = iZ
                    Case 3
                        pX_d3 = iX
                        pY_d3 = iY
                        pZ_d3 = iZ
                    Case 4
                        pX_d4 = iX
                        pY_d4 = iY
                        pZ_d4 = iZ
                    Case 5
                        pX_d5 = iX
                        pY_d5 = iY
                        pZ_d5 = iZ
                End Select

                If dF < dE Then
                    Exit For
                Else
                    p130(iRow).dCal2 = dF - dE
                End If
            Next
        End If

        fncItiCalc = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

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
        ReDim Preserve pIti(pcMax_d, n)     '位置情報

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

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' New_Itiファイルを作成する
    ''' </summary>
    ''' <param name="m">ドローン</param>
    ''' <param name="n">周回</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncNewItiFile(ByVal m As Integer, ByVal n As Integer) As Boolean
        On Error GoTo Error_Rtn
        fncNewItiFile = False

        Dim sData As String = ""            'ファイル書込み用データ

        Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\New_Iti.csv", True, System.Text.Encoding.UTF8)
        If n = 0 Then
            sData = pIti(m, 0).sIdo & "," & pIti(m, 0).sKeido & "," & pIti(m, 0).sTakasa
        Else
            sData = pIti(m, n).sIdo & "," & pIti(m, n).sKeido & "," & pIti(m, n).sTakasa
        End If
        oFileWrite.WriteLine(sData)
        'クローズ
        oFileWrite.Dispose()
        oFileWrite.Close()

        fncNewItiFile = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 130ファイルを配列にセットする
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
