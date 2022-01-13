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

        '**********************************************
        '配列を更新
        For iLoop = 1 To pSet_d
            If iLoop <> m Then
                'routeタグ
                sValue_route(0) = Format(CDbl(pIti(m, n - 1).sIdo), "#.000000000000000") & " " & Format(CDbl(pIti(m, n - 1).sKeido), "#.000000000000000") & " " & Format(CDbl(pIti(m, n - 1).sTakasa), "#.000000000000000")
                sValue_route(1) = Format(CDbl(pIti(m, n).sIdo), "#.000000000000000") & " " & Format(CDbl(pIti(m, n).sKeido), "#.000000000000000") & " " & Format(CDbl(pIti(m, n).sTakasa), "#.000000000000000")

                pTag_route(iCnt) = fncTagKeyAdd(pcTag_route, "nVertices", 2, sValue_route)
                iCnt = iCnt + 1

                'pTag_route(iCnt) = fncTagKeyUpdate(pcTag_route, "begin_<route>", "d" & m & " Route")
                'pTag_route(iCnt) = pTag_route(iCnt) & vbCrLf
                'pTag_route(iCnt) = fncTagKeyUpdate(pTag_route(iCnt), "project_id", m)
                'pTag_route(iCnt) = pTag_route(iCnt) & vbCrLf
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
        Dim iLoop As Integer

        'txrxファイル作成
        Dim enc As System.Text.Encoding = New System.Text.UTF8Encoding(False)
        Dim oFileWrite As New System.IO.StreamWriter(sFile, False, enc)

        'routeタグ
        Dim iCnt As Integer = 0
        Dim sTag_Route As String = ""
        For iLoop = 1 To pSet_d
            If iLoop <> m Then
                'routeタグ
                sTag_Route = pTag_route(iCnt)
                oFileWrite.WriteLine(sTag_Route)
                'gridタグ
                oFileWrite.WriteLine(pTag_grid)

                iCnt = iCnt + 1
            End If
        Next

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

        Dim sRowData As String = ""     '行データ
        Dim sColData As String = ""     '1文字データ
        Dim sData As String = ""        'ブロックデータ
        Dim iLOOP As Integer = 0        'ブロックデータループ
        Dim jLOOP As Integer = 0        '行データループ
        Dim iRowCnt As Integer = 0      '行カウント
        Dim jRowCnt As Integer = 0      '行カウント
        Dim iAddFlg As Integer = 0      '追加フラグ

        If iAddCnt <= sValue.Length Then
            If sBlockData.Length >= 1 Then
                For iLOOP = 1 To sBlockData.Length
                    If Mid(sBlockData, iLOOP, 1) = vbLf Then
                        '行カウント
                        iRowCnt = iRowCnt + 1
                    End If
                Next

                For iLOOP = 1 To sBlockData.Length
                    '行データ取得
                    If Mid(sBlockData, iLOOP, 1) = vbLf Or Mid(sBlockData, iLOOP, 1) = vbCr Then
                        If sRowData.Contains(sTagName) Then
                            '行データに該当タグがある
                            sData = sData & sRowData & vbCrLf
                            '行カウント
                            jRowCnt = jRowCnt + 1
                            '追加フラグ
                            iAddFlg = 1

                            '行データクリア
                            sRowData = ""
                        Else
                            If iAddFlg = 1 Then
                                For jLOOP = 0 To iAddCnt - 1
                                    '行追加
                                    sData = sData & sValue(jLOOP) & vbCrLf
                                    '行カウント
                                    jRowCnt = jRowCnt + 1
                                Next
                                '追加フラグ
                                iAddFlg = 0
                            End If

                            '行データに該当タグはない
                            If sRowData <> "" Then
                                '行カウント
                                jRowCnt = jRowCnt + 1

                                'ブロックデータにセットする
                                If jRowCnt >= (iRowCnt + iAddCnt) Then
                                    '最終行は改行しない
                                    sData = sData & sRowData
                                Else
                                    sData = sData & sRowData & vbCrLf
                                End If
                                '行データクリア
                                sRowData = ""
                            End If
                        End If
                    Else
                        '行データ
                        sRowData = sRowData & Mid(sBlockData, iLOOP, 1)
                    End If
                Next
            End If
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

        Dim sRowData As String = ""     '行データ
        Dim sColData As String = ""     '1文字データ
        Dim sData As String = ""        'ブロックデータ
        Dim iLOOP As Integer            'ブロックデータループ
        Dim jLOOP As Integer            '行データループ
        Dim iRowCnt As Integer = 0      '行カウント
        Dim jRowCnt As Integer = 0      '行カウント

        If sBlockData.Length >= 1 Then
            For iLOOP = 1 To sBlockData.Length
                If Mid(sBlockData, iLOOP, 1) = vbLf Then
                    '行カウント
                    iRowCnt = iRowCnt + 1
                End If
            Next

            For iLOOP = 1 To sBlockData.Length
                '行データ取得
                If Mid(sBlockData, iLOOP, 1) = vbLf Or Mid(sBlockData, iLOOP, 1) = vbCr Then
                    If sRowData.Contains(sTagName) Then
                        '行データに該当タグがある
                        sColData = ""
                        For jLOOP = 1 To sRowData.Length
                            'スペース後に値を更新する
                            If Mid(sRowData, jLOOP, 1) = " " Then
                                'ブロックデータにセットする
                                sData = sData & (sColData & " " & sValue) & vbCrLf
                                Exit For
                            Else
                                sColData = sColData & Mid(sRowData, jLOOP, 1)
                            End If
                        Next
                        '行カウント
                        jRowCnt = jRowCnt + 1

                        '行データクリア
                        sRowData = ""
                    Else
                        '行データに該当タグはない
                        If sRowData <> "" Then
                            '行カウント
                            jRowCnt = jRowCnt + 1

                            'ブロックデータにセットする
                            If jRowCnt = iRowCnt Then
                                '最終行は改行しない
                                sData = sData & sRowData
                            Else
                                sData = sData & sRowData & vbCrLf
                            End If
                            '行データクリア
                            sRowData = ""
                        End If
                    End If
                Else
                    '行データ
                    sRowData = sRowData & Mid(sBlockData, iLOOP, 1)
                End If
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
    Public Function fncTagCrLf(ByVal sBlockData As String) As String

        fncTagCrLf = ""

        Dim sData As String = ""        'ブロックデータ

        sData = sBlockData.Replace(vbLf, vbCrLf)

        fncTagCrLf = sData

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
                id = pOrder_d(iLoop).Substring(1, 1)
                iv = pOrder_d(iLoop).Substring(3, 4)
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