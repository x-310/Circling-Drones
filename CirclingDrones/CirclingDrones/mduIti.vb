'*******************************************************
'      経路計算用モジュール
'*******************************************************
Module mduIti

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
    Public Function fncCalc130() As Boolean

        fncCalc130 = False

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

        fncCalc130 = True

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
    Public Function fnc130Calc(ByVal m As Integer, ByVal n As Integer) As Boolean

        fnc130Calc = False

        Dim iRow As Integer = 0
        Dim iA1, iA2 As Integer
        Dim iB1, iB2 As Integer
        Dim iC1, iC2 As Integer
        Dim dE, dF As Double
        Dim dX As Double = 0
        Dim dY As Double = 0
        Dim dZ As Double = 0
        Dim iX As Integer = 0.0
        Dim iY As Integer = 0.0
        Dim iZ As Integer = 0.0

        'pX_d1 = 0 : pY_d1 = 0 : pZ_d1 = 0
        'pX_d2 = 0 : pY_d2 = 0 : pZ_d2 = 0
        'pX_d3 = 0 : pY_d3 = 0 : pZ_d3 = 0
        'pX_d4 = 0 : pY_d4 = 0 : pZ_d4 = 0
        'pX_d5 = 0 : pY_d5 = 0 : pZ_d5 = 0

        ''n=0の場合、画面設定値から取込む
        'If n = 0 Then
        '   p130(0).dCal1 = 0.0
        '   p130(0).dCal2 = CDbl(frmMain.txtV.Text)
        '   p130(0).iX = CInt(pIti(m, 0).sIdo)
        '   p130(0).iY = CInt(pIti(m, 0).sKeido)
        '   p130(0).iZ = CInt(pIti(m, 0).sTakasa)
        'End If

        iX = 0
        iY = 0
        iZ = 0
        Select Case m
            Case 1
                '四捨五入
                iX = Math.Round(CDbl(frmMain.txtX_d1.Text), MidpointRounding.AwayFromZero)
                iY = Math.Round(CDbl(frmMain.txtY_d1.Text), MidpointRounding.AwayFromZero)
                iZ = Math.Round(CDbl(frmMain.txtZ_d1.Text), MidpointRounding.AwayFromZero)

                'pX_d1 = iX
                'pY_d1 = iY
                'pZ_d1 = iZ
            Case 2
                '四捨五入
                iX = Math.Round(CDbl(frmMain.txtX_d2.Text), MidpointRounding.AwayFromZero)
                iY = Math.Round(CDbl(frmMain.txtY_d2.Text), MidpointRounding.AwayFromZero)
                iZ = Math.Round(CDbl(frmMain.txtZ_d2.Text), MidpointRounding.AwayFromZero)

                'pX_d2 = iX
                'pY_d2 = iY
                'pZ_d2 = iZ
            Case 3
                '四捨五入
                iX = Math.Round(CDbl(frmMain.txtX_d3.Text), MidpointRounding.AwayFromZero)
                iY = Math.Round(CDbl(frmMain.txtY_d3.Text), MidpointRounding.AwayFromZero)
                iZ = Math.Round(CDbl(frmMain.txtZ_d3.Text), MidpointRounding.AwayFromZero)

                'pX_d3 = iX
                'pY_d3 = iY
                'pZ_d3 = iZ
            Case 4
                '四捨五入
                iX = Math.Round(CDbl(frmMain.txtX_d4.Text), MidpointRounding.AwayFromZero)
                iY = Math.Round(CDbl(frmMain.txtY_d4.Text), MidpointRounding.AwayFromZero)
                iZ = Math.Round(CDbl(frmMain.txtZ_d4.Text), MidpointRounding.AwayFromZero)

                'pX_d4 = iX
                'pY_d4 = iY
                'pZ_d4 = iZ
            Case 5
                '四捨五入
                iX = Math.Round(CDbl(frmMain.txtX_d5.Text), MidpointRounding.AwayFromZero)
                iY = Math.Round(CDbl(frmMain.txtY_d5.Text), MidpointRounding.AwayFromZero)
                iZ = Math.Round(CDbl(frmMain.txtZ_d5.Text), MidpointRounding.AwayFromZero)

                'pX_d5 = iX
                'pY_d5 = iY
                'pZ_d5 = iZ
        End Select

        For iRow = 1 To p130.Length - 1
            iA1 = CInt(p130(iRow - 1).sX)
            iA2 = CInt(p130(iRow).sX)
            iB1 = CInt(p130(iRow - 1).sY)
            iB2 = CInt(p130(iRow).sY)
            iC1 = CInt(p130(iRow - 1).sZ)
            iC2 = CInt(p130(iRow).sZ)
            dE = p130(iRow).dCal1
            dF = p130(iRow - 1).dCal2

            If dE = 0 Then
                dX = 0
                dY = 0
                dZ = 0
            Else
                dX = dF / dE * (iA2 - iA1) + iA1
                dY = dF / dE * (iB2 - iB1) + iB1
                dZ = dF / dE * (iC2 - iC1) + iC1
            End If

            '四捨五入
            iX = Math.Round(dX, MidpointRounding.AwayFromZero)
            iY = Math.Round(dY, MidpointRounding.AwayFromZero)
            iZ = Math.Round(dZ, MidpointRounding.AwayFromZero)

            p130(iRow).iX = iX
            p130(iRow).iY = iY
            p130(iRow).iZ = iZ

            'Select Case m
            '    Case 1
            '        pX_d1 = iX
            '        pY_d1 = iY
            '        pZ_d1 = iZ
            '    Case 2
            '        pX_d2 = iX
            '        pY_d2 = iY
            '        pZ_d2 = iZ
            '    Case 3
            '        pX_d3 = iX
            '        pY_d3 = iY
            '        pZ_d3 = iZ
            '    Case 4
            '        pX_d4 = iX
            '        pY_d4 = iY
            '        pZ_d4 = iZ
            '    Case 5
            '        pX_d5 = iX
            '        pY_d5 = iY
            '        pZ_d5 = iZ
            'End Select

            If dF < dE Then
                Exit For
            Else
                p130(iRow).dCal2 = dF - dE
            End If
        Next

        fnc130Calc = True

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

        fncItiFile = False

        Dim sData As String = ""            'ファイル書込み用データ
        Dim iLoop As Integer = 0

        'm=1～、n=0～
        If n <> 0 Then
            ReDim Preserve pIti(pcMax_d, n + 1)   '位置情報

            pIti(m, n).sIdo = p130(n).iX
            pIti(m, n).sKeido = p130(n).iY
            pIti(m, n).sTakasa = p130(n).iZ

            pIti(m, n + 1).sIdo = p130_2(n).iX
            pIti(m, n + 1).sKeido = p130_2(n).iY
            pIti(m, n + 1).sTakasa = p130_2(n).iZ
        End If

        Select Case m
            Case 1
                'm=1
                'Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d1_iti.csv", False, System.Text.Encoding.UTF8)
                Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d1_iti.csv", False, pEnc)

                sData = ""
                For iLoop = 0 To n
                    sData = sData & pIti(m, iLoop).sIdo & "," & pIti(m, iLoop).sKeido & "," & pIti(m, iLoop).sTakasa & vbCrLf
                Next
                oFileWrite.WriteLine(sData)

                'クローズ
                oFileWrite.Dispose()
                oFileWrite.Close()
            Case 2
                'm=2
                'Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d2_iti.csv", False, System.Text.Encoding.UTF8)
                Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d2_iti.csv", False, pEnc)

                For iLoop = 0 To n
                    sData = sData & pIti(m, iLoop).sIdo & "," & pIti(m, iLoop).sKeido & "," & pIti(m, iLoop).sTakasa & vbCrLf
                Next
                oFileWrite.WriteLine(sData)

                'クローズ
                oFileWrite.Dispose()
                oFileWrite.Close()
            Case 3
                'm=3
                'Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d3_iti.csv", False, System.Text.Encoding.UTF8)
                Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d3_iti.csv", False, pEnc)

                For iLoop = 0 To n
                    sData = sData & pIti(m, iLoop).sIdo & "," & pIti(m, iLoop).sKeido & "," & pIti(m, iLoop).sTakasa & vbCrLf
                Next
                oFileWrite.WriteLine(sData)

                'クローズ
                oFileWrite.Dispose()
                oFileWrite.Close()
            Case 4
                'm=4
                'Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d4_iti.csv", False, System.Text.Encoding.UTF8)
                Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d4_iti.csv", False, pEnc)

                For iLoop = 0 To n
                    sData = sData & pIti(m, iLoop).sIdo & "," & pIti(m, iLoop).sKeido & "," & pIti(m, iLoop).sTakasa & vbCrLf
                Next
                oFileWrite.WriteLine(sData)

                'クローズ
                oFileWrite.Dispose()
                oFileWrite.Close()
            Case 5
                'm=5
                'Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d5_iti.csv", False, System.Text.Encoding.UTF8)
                Dim oFileWrite As New System.IO.StreamWriter(pPjPath & "\d5_iti.csv", False, pEnc)

                For iLoop = 0 To n
                    sData = sData & pIti(m, iLoop).sIdo & "," & pIti(m, iLoop).sKeido & "," & pIti(m, iLoop).sTakasa & vbCrLf
                Next
                oFileWrite.WriteLine(sData)

                'クローズ
                oFileWrite.Dispose()
                oFileWrite.Close()
        End Select

        fncItiFile = True

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

        fncNewItiFile = False

        Dim sData As String = ""            'ファイル書込み用データ
        Dim dX As Double = 0.0
        Dim dY As Double = 0.0
        Dim dZ As Double = 0.0
        Dim sX As String = ""
        Dim sY As String = ""
        Dim sZ As String = ""

        'Dim oFileWrite As New System.IO.StreamWriter(pExePath & "\New_Iti.csv", False, System.Text.Encoding.UTF8)
        Dim oFileWrite As New System.IO.StreamWriter(pExePath & "\New_Iti.csv", False, pEnc)

        If n = 0 Then
            dX = Integer.Parse(pIti(m, 0).sIdo)
            dY = Integer.Parse(pIti(m, 0).sKeido)
            dZ = Integer.Parse(pIti(m, 0).sTakasa)

            sX = (dX / 5).ToString("0")
            sY = (dY / 5).ToString("0")
            sZ = ((dZ - 10) / 5).ToString("0")

            sData = sX & "," & sY & "," & sZ
        Else
            dX = Integer.Parse(pIti(m, n - 1).sIdo)
            dY = Integer.Parse(pIti(m, n - 1).sKeido)
            dZ = Integer.Parse(pIti(m, n - 1).sTakasa)

            sX = (dX / 5).ToString("0")
            sY = (dY / 5).ToString("0")
            sZ = ((dZ - 10) / 5).ToString("0")

            sData = sX & "," & sY & "," & sZ
            'sData = pIti(m, n).sIdo & "," & pIti(m, n).sKeido & "," & pIti(m, n).sTakasa
        End If
        oFileWrite.WriteLine(sData)
        'クローズ
        oFileWrite.Dispose()
        oFileWrite.Close()

        fncNewItiFile = True

    End Function

End Module
