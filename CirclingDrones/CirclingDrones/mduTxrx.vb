'*******************************************************
'      Txrx作成用モジュール
'*******************************************************
Module mduTxrx

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' タグデータを作成する
    ''' </summary>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncTagCreate(ByVal m As Integer, ByVal n As Integer) As Boolean

        fncTagCreate = False

        Dim sFile = pPjPath & "\" & pPjName & ".txrx"   'txrxファイルパス
        Dim iLoop As Integer = 0
        Dim iCnt As Integer = 0
        Dim sValue_route(1) As String                   'route用nVerticesデータ用配列
        Dim dX_1 As Double
        Dim dY_1 As Double
        Dim dZ_1 As Double
        Dim sId As String = "45"

        '**********************************************
        '配列を更新
        For iLoop = 1 To pSet_d
            If iLoop <> m Then
                If n = 1 Then
                    dX_1 = CDbl(pIti(iLoop, 0).sIdo) + 1079
                    dY_1 = CDbl(pIti(iLoop, 0).sKeido) + 537
                    dZ_1 = fncGetSabun(CDbl(pIti(iLoop, 0).sIdo), CDbl(pIti(iLoop, 0).sKeido))
                    dZ_1 = dZ_1 + CDbl(pIti(iLoop, 0).sTakasa) + 10
                    'routeタグ
                    sValue_route(0) = dX_1.ToString("0.000000000000000") & " " & dY_1.ToString("0.000000000000000") & " " & dZ_1.ToString("0.000000000000000")
                Else
                    dX_1 = CDbl(pIti(iLoop, n - 1).sIdo) + 1079
                    dY_1 = CDbl(pIti(iLoop, n - 1).sKeido) + 537
                    dZ_1 = fncGetSabun(CDbl(pIti(iLoop, n - 1).sIdo), CDbl(pIti(iLoop, n - 1).sKeido))
                    dZ_1 = dZ_1 + CDbl(pIti(iLoop, n - 1).sTakasa) + 10
                    'routeタグ
                    sValue_route(0) = dX_1.ToString("0.000000000000000") & " " & dY_1.ToString("0.000000000000000") & " " & dZ_1.ToString("0.000000000000000")
                End If

                pTag_point(iCnt) = fncTagKeyAdd(pcTag_point, "nVertices", 1, sValue_route)

                pTag_point(iCnt) = fncTagKeyUpdate(pTag_point(iCnt), "begin_<points> another drone tx", iCnt + 1)

                iCnt = iCnt + 1
            End If
        Next

        fncTagCreate = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' プロジェクト名.txrxファイルを作成する
    ''' </summary>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncMem2txrx(ByVal m As Integer) As Boolean

        fncMem2txrx = False

        Dim sFile = pPjPath & "\" & pPjName & ".txrx"   'txrxファイルパス

        'txrxファイル作成
        'Dim oFileWrite As New System.IO.StreamWriter(sFile, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'Dim enc As System.Text.Encoding = New System.Text.UTF8Encoding(False)
        Dim oFileWrite As New System.IO.StreamWriter(sFile, False, pEnc)

        'pointタグ
        Dim iCnt As Integer = 0
        Dim sTag_Point As String = ""

        'gridタグ
        oFileWrite.WriteLine(pTag_grid)

        'pointタグ 1
        sTag_Point = pTag_point(0)
        oFileWrite.WriteLine(sTag_Point)

        If pSet_d = 3 Then
            'pointタグ 2
            sTag_Point = fncTagKeyUpdate(pTag_point(1), "project_id", "42")
            oFileWrite.WriteLine(sTag_Point)
        ElseIf pSet_d = 4 Then
            'pointタグ 2
            sTag_Point = fncTagKeyUpdate(pTag_point(1), "project_id", "42")
            oFileWrite.WriteLine(sTag_Point)

            System.Threading.Thread.Sleep(500)

            'pointタグ 3
            sTag_Point = fncTagKeyUpdate(pTag_point(2), "project_id", "43")
            oFileWrite.WriteLine(sTag_Point)
        ElseIf pSet_d = 5 Then
            'pointタグ 2
            sTag_Point = fncTagKeyUpdate(pTag_point(1), "project_id", "42")
            oFileWrite.WriteLine(sTag_Point)

            System.Threading.Thread.Sleep(500)

            'pointタグ 3
            sTag_Point = fncTagKeyUpdate(pTag_point(2), "project_id", "43")
            oFileWrite.WriteLine(sTag_Point)

            System.Threading.Thread.Sleep(500)

            'pointタグ 4
            sTag_Point = fncTagKeyUpdate(pTag_point(3), "project_id", "44")
            oFileWrite.WriteLine(sTag_Point)
        End If

        'クローズ
        oFileWrite.Dispose()
        oFileWrite.Close()

        fncMem2txrx = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 履歴用Txrxファイルにコピーする
    ''' </summary>
    ''' <param name="m">ドローン</param>
    ''' <param name="n">周回</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncCopyTxrx(ByVal m As Integer, ByVal n As Integer) As Boolean

        fncCopyTxrx = False

        fncFileCopy(pPjPath & "\" & pPjName & ".txrx", pPjPath & "\d" & m & "_v" & n & "_" & pPjName & ".txrx")

        fncCopyTxrx = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' タグの項目の下にキー値を追加する
    ''' </summary>
    ''' <param name="sBlockData">ブロックデータ</param>
    ''' <param name="sTagName">タグ名</param>
    ''' <param name="iAddCnt">追加したい行数</param>
    ''' <param name="sValue">追加データ</param>
    ''' <returns>追加する項目データ</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncTagKeyAdd(ByVal sBlockData As String, ByVal sTagName As String, ByVal iAddCnt As Integer, ByVal sValue() As String) As String

        fncTagKeyAdd = ""

        Dim sMemData() As String
        Dim sRowData As String = ""     '行データ
        Dim sColData As String = ""     '1文字データ
        Dim sData As String = ""        'ブロックデータ
        Dim iLOOP As Integer            'ブロックデータループ
        Dim iRowCnt As Integer = 0      '行カウント
        Dim jRowCnt As Integer = 0      '行カウント
        Dim iColNo As Integer = 0
        Dim sColData_Mae As String = ""

        If sBlockData.Length >= 1 Then
            ReDim Preserve sMemData(-1)
            For iLOOP = 1 To sBlockData.Length
                If Mid(sBlockData, iLOOP, 1) = vbLf Or Mid(sBlockData, iLOOP, 1) = vbCr Then
                    If sData <> "" Then
                        ReDim Preserve sMemData(iRowCnt)

                        sMemData(iRowCnt) = sData & vbCrLf
                        sData = ""
                        '行カウント
                        iRowCnt = iRowCnt + 1
                    End If
                Else
                    sData = sData + Mid(sBlockData, iLOOP, 1)
                End If
            Next

            ReDim Preserve sMemData(iRowCnt)
            sMemData(iRowCnt) = sData

            For iLOOP = 0 To iRowCnt
                '行データ取得
                iColNo = sMemData(iLOOP).IndexOf(sTagName)
                If iColNo <> -1 Then
                    sMemData(iLOOP + 1) = sValue(0) & vbCrLf
                    'sMemData(iLOOP + 2) = sValue(1) & vbCrLf
                End If
            Next

            sData = ""
            For iLOOP = 0 To iRowCnt
                '行データ取得
                sData = sData + sMemData(iLOOP)
            Next
        End If

        fncTagKeyAdd = sData

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' タグの項目でスペース以降の値を更新する
    ''' </summary>
    ''' <param name="sBlockData">ブロックデータ</param>
    ''' <param name="sTagName">タグ名</param>
    ''' <param name="sValue">変更データ</param>
    ''' <returns>更新する項目データ</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncTagKeyUpdate(ByVal sBlockData As String, ByVal sTagName As String, sValue As String) As String

        fncTagKeyUpdate = ""

        Dim sMemData() As String
        Dim sRowData As String = ""     '行データ
        Dim sColData As String = ""     '1文字データ
        Dim sData As String = ""        'ブロックデータ
        Dim iLOOP As Integer            'ブロックデータループ
        Dim iRowCnt As Integer = 0      '行カウント
        Dim jRowCnt As Integer = 0      '行カウント
        Dim iColNo As Integer = 0
        Dim sColData_Mae As String = ""

        If sBlockData.Length >= 1 Then
            ReDim Preserve sMemData(-1)
            For iLOOP = 1 To sBlockData.Length
                If Mid(sBlockData, iLOOP, 1) = vbLf Or Mid(sBlockData, iLOOP, 1) = vbCr Then
                    If sData <> "" Then
                        ReDim Preserve sMemData(iRowCnt)

                        sMemData(iRowCnt) = sData & vbCrLf
                        sData = ""
                        '行カウント
                        iRowCnt = iRowCnt + 1
                    End If
                Else
                    sData = sData + Mid(sBlockData, iLOOP, 1)
                End If
            Next

            ReDim Preserve sMemData(iRowCnt)
            sMemData(iRowCnt) = sData

            For iLOOP = 0 To iRowCnt
                '行データ取得
                iColNo = sMemData(iLOOP).IndexOf(sTagName)
                If iColNo <> -1 Then
                    sColData_Mae = Mid(sMemData(iLOOP), 1, iColNo + sTagName.Length)
                    sMemData(iLOOP) = sColData_Mae + " " & sValue & vbCrLf
                End If
            Next

            sData = ""
            For iLOOP = 0 To iRowCnt
                '行データ取得
                sData = sData + sMemData(iLOOP)
            Next
        End If

        fncTagKeyUpdate = sData

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 行末をCrLfに更新する
    ''' </summary>
    ''' <param name="sBlockData">ブロックデータ</param>
    ''' <returns>ブロックデータ</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncTagLf(ByVal sBlockData As String) As String

        fncTagLf = ""

        Dim sData As String = ""        'ブロックデータ

        sData = sBlockData.Replace(vbCrLf, vbLf)

        fncTagLf = sData

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Iti用配列からtxrx用配列を作成する
    ''' </summary>
    ''' <param name="m">ドローン</param>
    ''' <param name="n">周回</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncIti2Txrx(ByVal m As Integer, ByVal n As Integer) As Boolean

        fncIti2Txrx = False

        Dim iLoop As Integer        'ドローンループ
        Dim iRow As Integer = 0     '行No
        Dim id As Integer           'ドローンNo
        Dim iv As Integer           '周回No

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

                    iRow = iRow + 1
                End If
            Next
        Else
            'ドローン最新順
            For iLoop = pSet_d To 1 Step -1
                id = pOrder_d(iLoop).Substring(1, 1) 'pcMax_d
                iv = pOrder_d(iLoop).Substring(3, 1) 'pcMax_v
                If id <> m Then
                    ReDim Preserve pTxrx(iRow)             'pTxrx に新規行を追加

                    pTxrx(iRow).sdv = "d" & id & "v" & iv
                    pTxrx(iRow).sIdo = pIti(id, iv).sIdo
                    pTxrx(iRow).sKeido = pIti(id, iv).sKeido
                    pTxrx(iRow).sTakasa = pIti(id, iv).sTakasa

                    If (pSet_d - 1) = (iRow + 1) Then Exit For
                    iRow = iRow + 1
                End If
            Next
        End If

        fncIti2Txrx = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' プロジェクト名.txrxファイルを削除する
    ''' </summary>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFileDelete_pj_txrx() As Boolean

        fncFileDelete_pj_txrx = False

        Dim sFile As String = ""    'ファイル名

        sFile = pPjPath & "\" & pPjName & ".txrx"
        fncFileDel(sFile)

        fncFileDelete_pj_txrx = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' txrxファイルを削除する
    ''' </summary>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFileDelete_txrx() As Boolean

        fncFileDelete_txrx = False

        Dim sFile As String = ""    'ファイル名

        sFile = pPjPath & "\" & "*.txrx"
        fncFileDel(sFile)

        fncFileDelete_txrx = True

    End Function
End Module