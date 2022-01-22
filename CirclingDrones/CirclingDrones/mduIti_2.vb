'*******************************************************
'      経路計算用モジュール(tが倍用)
'*******************************************************
Module mduIti_2

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 130ファイルを配列にセットする
    ''' </summary>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFile2Grid_2() As Boolean

        fncFile2Grid_2 = False

        ReDim Preserve p130_2(p130.Length - 1)

        Array.Copy(p130, p130_2, p130.Length)

        '130.CSVファイルを配列にセットする
        'Dim arrCsv()() As String = ReadCsv(sFileName, False, False)
        'Dim iRow As Integer = 0

        'If arrCsv.Length > 0 Then
        '    For iRow = 0 To arrCsv.Length - 1
        '        ReDim Preserve p130_2(iRow)       'p130_2 に新規行を追加
        '        If iRow = 0 Then
        '            p130_2(iRow).sX = arrCsv(iRow)(0).ToString
        '            p130_2(iRow).sY = arrCsv(iRow)(1).ToString
        '            p130_2(iRow).sZ = arrCsv(iRow)(2).ToString
        '        Else
        '            p130_2(iRow).sX = arrCsv(iRow)(0).ToString
        '            p130_2(iRow).sY = arrCsv(iRow)(1).ToString
        '            p130_2(iRow).sZ = arrCsv(iRow)(2).ToString
        '            p130_2(iRow).sAA = arrCsv(iRow)(3).ToString
        '        End If
        '    Next
        'End If

        fncFile2Grid_2 = True

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
    Public Function fnc130Calc_2() As Boolean

        fnc130Calc_2 = False

        Dim iA1 As Integer = 0
        Dim iA2 As Integer = 0
        Dim iB1 As Integer = 0
        Dim iB2 As Integer = 0
        Dim iC1 As Integer = 0
        Dim iC2 As Integer = 0
        Dim dE As Double = 0.0
        Dim dF As Double = 0.0

        For iRow = 0 To p130_2.Length - 1
            If iRow = 0 Then
                iA1 = CInt(p130_2(0).sX)
                iA2 = CInt(p130_2(0).sX)
                iB1 = CInt(p130_2(0).sY)
                iB2 = CInt(p130_2(0).sY)
                iC1 = CInt(p130_2(0).sZ)
                iC2 = CInt(p130_2(0).sZ)
                dE = 0.0
                dF = CInt(frmMain.txtVT2.Text)
            Else
                iA1 = CInt(p130_2(iRow - 1).sX)
                iA2 = CInt(p130_2(iRow).sX)
                iB1 = CInt(p130_2(iRow - 1).sY)
                iB2 = CInt(p130_2(iRow).sY)
                iC1 = CInt(p130_2(iRow - 1).sZ)
                iC2 = CInt(p130_2(iRow).sZ)
                dE = Math.Sqrt((iA2 - iA1) ^ 2 + (iB2 - iB1) ^ 2 + (iC2 - iC1) ^ 2)
                dF = p130_2(iRow - 1).dCal2 - dE
            End If

            p130_2(iRow).dCal1 = dE
            p130_2(iRow).dCal2 = dF
        Next

        fnc130Calc_2 = True

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
    Public Function fncItiCalc_2(ByVal m As Integer, ByVal n As Integer) As Boolean

        fncItiCalc_2 = False

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

        'n=0の場合、画面設定値から取込む
        If n = 0 Then
            ReDim Preserve p130_2(0)

            iX = 0
            iY = 0
            iZ = 0
            Select Case m
                Case 1
                    '四捨五入
                    iX = Math.Round(CDbl(frmMain.txtX_d1.Text), MidpointRounding.AwayFromZero)
                    iY = Math.Round(CDbl(frmMain.txtY_d1.Text), MidpointRounding.AwayFromZero)
                    iZ = Math.Round(CDbl(frmMain.txtZ_d1.Text), MidpointRounding.AwayFromZero)
                Case 2
                    '四捨五入
                    iX = Math.Round(CDbl(frmMain.txtX_d2.Text), MidpointRounding.AwayFromZero)
                    iY = Math.Round(CDbl(frmMain.txtY_d2.Text), MidpointRounding.AwayFromZero)
                    iZ = Math.Round(CDbl(frmMain.txtZ_d2.Text), MidpointRounding.AwayFromZero)
                Case 3
                    '四捨五入
                    iX = Math.Round(CDbl(frmMain.txtX_d3.Text), MidpointRounding.AwayFromZero)
                    iY = Math.Round(CDbl(frmMain.txtY_d3.Text), MidpointRounding.AwayFromZero)
                    iZ = Math.Round(CDbl(frmMain.txtZ_d3.Text), MidpointRounding.AwayFromZero)
                Case 4
                    '四捨五入
                    iX = Math.Round(CDbl(frmMain.txtX_d4.Text), MidpointRounding.AwayFromZero)
                    iY = Math.Round(CDbl(frmMain.txtY_d4.Text), MidpointRounding.AwayFromZero)
                    iZ = Math.Round(CDbl(frmMain.txtZ_d4.Text), MidpointRounding.AwayFromZero)
                Case 5
                    '四捨五入
                    iX = Math.Round(CDbl(frmMain.txtX_d5.Text), MidpointRounding.AwayFromZero)
                    iY = Math.Round(CDbl(frmMain.txtY_d5.Text), MidpointRounding.AwayFromZero)
                    iZ = Math.Round(CDbl(frmMain.txtZ_d5.Text), MidpointRounding.AwayFromZero)
            End Select

            p130_2(0).dCal1 = 0.0
            p130_2(0).dCal2 = CDbl(frmMain.txtV2.Text)
            p130_2(0).iX = iX
            p130_2(0).iY = iY
            p130_2(0).iZ = iZ
        Else
            'p130_2(0).iX = CInt(pIti(m, 0).sIdo)
            'p130_2(0).iY = CInt(pIti(m, 0).sKeido)
            'p130_2(0).iZ = CInt(pIti(m, 0).sTakasa)

            For iRow = 1 To p130_2.Length - 1
                iA1 = CInt(p130_2(iRow - 1).sX)
                iA2 = CInt(p130_2(iRow).sX)
                iB1 = CInt(p130_2(iRow - 1).sY)
                iB2 = CInt(p130_2(iRow).sY)
                iC1 = CInt(p130_2(iRow - 1).sZ)
                iC2 = CInt(p130_2(iRow).sZ)
                dE = p130_2(iRow).dCal1
                dF = p130_2(iRow - 1).dCal2

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

                p130_2(iRow).iX = iX
                p130_2(iRow).iY = iY
                p130_2(iRow).iZ = iZ

                If dF < dE Then
                    Exit For
                Else
                    p130_2(iRow).dCal2 = dF - dE
                End If
            Next
        End If

        fncItiCalc_2 = True

    End Function

End Module
