'*******************************************************
'      経路計算用モジュール
'*******************************************************
Module mduIti

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

        'n=0の場合、画面設定値から取込む
        If n = 0 Then
            Select Case m
                Case 1
                    pX_d1 = frmMain.txtX_d1.Text
                    pY_d1 = frmMain.txtY_d1.Text
                    pZ_d1 = frmMain.txtZ_d1.Text
                Case 2
                    pX_d2 = frmMain.txtX_d2.Text
                    pY_d2 = frmMain.txtY_d2.Text
                    pZ_d2 = frmMain.txtZ_d2.Text
                Case 3
                    pX_d3 = frmMain.txtX_d3.Text
                    pY_d3 = frmMain.txtY_d3.Text
                    pZ_d3 = frmMain.txtZ_d3.Text
                Case 4
                    pX_d4 = frmMain.txtX_d4.Text
                    pY_d4 = frmMain.txtY_d4.Text
                    pZ_d4 = frmMain.txtZ_d4.Text
                Case 5
                    pX_d5 = frmMain.txtX_d5.Text
                    pY_d5 = frmMain.txtY_d5.Text
                    pZ_d5 = frmMain.txtZ_d5.Text
            End Select
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

                iX = CInt(p130(iRow).sX)
                iY = CInt(p130(iRow).sY)
                iZ = CInt(p130(iRow).sZ)

                dX = 0.0 : dY = 0.0 : dZ = 0.0
                If dE > dF Then
                    dX = dF / dE * (iA2 - iA1) + iA1
                    dY = dF / dE * (iB2 - iB1) + iB1
                    dZ = dF / dE * (iC2 - iC1) + iC1
                    Exit For
                End If
            Next

            '四捨五入
            iX = Math.Round(dX, MidpointRounding.AwayFromZero)
            iY = Math.Round(dY, MidpointRounding.AwayFromZero)
            iZ = Math.Round(dZ, MidpointRounding.AwayFromZero)

            Select Case m
                Case 1
                    If dE > dF Then
                        pX_d1 = iX
                        pY_d1 = iY
                        pZ_d1 = iZ
                    End If
                Case 2
                    If dE > dF Then
                        pX_d2 = iX
                        pY_d2 = iY
                        pZ_d2 = iZ
                    End If
                Case 3
                    If dE > dF Then
                        pX_d3 = iX
                        pY_d3 = iY
                        pZ_d3 = iZ
                    End If
                Case 4
                    If dE > dF Then
                        pX_d4 = iX
                        pY_d4 = iY
                        pZ_d4 = iZ
                    End If
                Case 5
                    If dE > dF Then
                        pX_d5 = iX
                        pY_d5 = iY
                        pZ_d5 = iZ
                    End If
            End Select
        End If

        fncItiCalc = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

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
            If iRow = 0 Then
                p130(iRow).dCal2 = CInt(frmMain.txtVT.Text)
            Else
                p130(iRow).dCal2 = p130(iRow).dCal1 - p130(iRow).dCal2
            End If
        Next

        fnc130Calc = True
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
            sData = p130(0).sX & "," & p130(0).sY & "," & p130(0).sZ
        Else
            sData = p130(n - 1).sX & "," & p130(n - 1).sY & "," & p130(n - 1).sZ
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
