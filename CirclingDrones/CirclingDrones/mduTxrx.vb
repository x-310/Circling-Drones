'*******************************************************
'      Txrx作成用モジュール
'*******************************************************
Module mduTxrx

    '*******************************************************
    '   プロジェクト名.txrxファイルを作成する
    '*******************************************************
    Public Function fncMem2txrx() As Boolean
        On Error GoTo Error_Rtn
        fncMem2txrx = False

        Dim sFile = pPjPath & "\" & pPjName & ".txrx"   'txrxファイルパス

        Dim inVertices_points As Integer = 0            'points用nVertices作成個数
        Dim sValue_points() As String                   'points用nVerticesデータ用配列
        Dim inVertices_route As Integer = 0             'route用nVertices作成個数
        Dim sValue_route() As String                    'route用nVerticesデータ用配列
        Dim inVertices_grid As Integer = 0              'grid用nVertices作成個数
        Dim sValue_grid() As String                     'grid用nVerticesデータ用配列

        'ファイル存在チェック
        If fncFileCheck(sFile) Then
            '存在すればファイル削除
            fncFileDel(sFile)
        End If

        'データ用配列初期化
        ReDim Preserve sValue_points(-1)                'nVerticesデータ用配列初期化
        ReDim Preserve sValue_route(-1)                 'nVerticesデータ用配列初期化
        ReDim Preserve sValue_grid(-1)                  'nVerticesデータ用配列初期化

        ' **********(仮)**********
        'points用nVerticesデータ作成
        inVertices_points = 1                                                                   'points用nVertices作成個数
        ReDim Preserve sValue_points(inVertices_points - 1)                                     'points用nVerticesデータ用配列
        sValue_points(0) = "1111.247570000000000 111.187590000000000 1.000000000000000"
        '配列を更新
        pTag_points = fncTagKeyUpdate(pcTag_points, "begin_<points>", "*** Points ***")         'pointsタグ
        pTag_points = pTag_points & vbCrLf
        pTag_points = fncTagKeyUpdate(pTag_points, "nVertices", CInt(inVertices_points))        'pointsタグ
        pTag_points = pTag_points & vbCrLf
        pTag_points = fncTagKeyAdd(pTag_points, "nVertices", inVertices_points, sValue_points)  'pointsタグ

        'route用nVerticesデータ作成
        inVertices_route = 3                                                                'route用nVertices作成個数
        ReDim Preserve sValue_route(inVertices_route - 1)                                  'route用nVerticesデータ用配列
        sValue_route(0) = "1111.247570000000000 111.187590000000000 1.000000000000000"
        sValue_route(1) = "2222.247570000000000 222.187590000000000 2.000000000000000"
        sValue_route(2) = "3333.247570000000000 333.187590000000000 3.000000000000000"
        '配列を更新
        pTag_route = fncTagKeyUpdate(pcTag_route, "begin_<route>", "*** route ***")         'routeタグ
        pTag_route = pTag_route & vbCrLf
        pTag_route = fncTagKeyUpdate(pTag_route, "nVertices", CInt(inVertices_route))       'routeタグ
        pTag_route = pTag_route & vbCrLf
        pTag_route = fncTagKeyAdd(pTag_route, "nVertices", inVertices_route, sValue_route)  'routeタグ

        'grid用nVerticesデータ作成
        inVertices_grid = 5                                                                 'grid用nVertices作成個数
        ReDim Preserve sValue_grid(inVertices_grid - 1)                                     'grid用nVerticesデータ用配列
        sValue_grid(0) = "1111.247570000000000 111.187590000000000 1.000000000000000"
        sValue_grid(1) = "2222.247570000000000 222.187590000000000 2.000000000000000"
        sValue_grid(2) = "3333.247570000000000 333.187590000000000 3.000000000000000"
        sValue_grid(3) = "4444.247570000000000 444.187590000000000 4.000000000000000"
        sValue_grid(4) = "5555.247570000000000 555.187590000000000 5.000000000000000"
        '配列を更新
        pTag_grid = fncTagKeyUpdate(pcTag_grid, "begin_<grid>", "*** grid ***")             'gridタグ
        pTag_grid = pTag_grid & vbCrLf
        pTag_grid = fncTagKeyUpdate(pTag_grid, "nVertices", CInt(inVertices_grid))          'gridタグ
        pTag_grid = pTag_grid & vbCrLf
        pTag_grid = fncTagKeyAdd(pTag_grid, "nVertices", inVertices_grid, sValue_grid)      'gridタグ

        'txrxファイル作成
        Dim oFileWrite As New System.IO.StreamWriter(sFile, True, System.Text.Encoding.UTF8)

        oFileWrite.WriteLine(pTag_points)  'pointsタグ
        'ループにする
        oFileWrite.WriteLine(pTag_route)   'routeタグ
        oFileWrite.WriteLine(pTag_route)   'routeタグ
        oFileWrite.WriteLine(pTag_route)   'routeタグ

        oFileWrite.WriteLine(pTag_grid)    'gridタグ

        'クローズ
        oFileWrite.Dispose()
        oFileWrite.Close()

        fncMem2txrx = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '   タグの項目の下にキー値を追加する
    '*******************************************************
    Public Function fncTagKeyAdd(ByVal sBlockData As String, ByVal sTagName As String, ByVal iAddCnt As Integer, ByVal sValue() As String) As String
        On Error GoTo Error_Rtn
        fncTagKeyAdd = ""

        Dim sRowData As String = ""                 '行データ
        Dim sColData As String = ""                 '1文字データ
        Dim sData As String = ""                    'ブロックデータ
        Dim iLOOP As Integer = 0
        Dim jLOOP As Integer = 0
        Dim kLOOP As Integer = 0
        Dim iRowCnt As Integer = 0
        Dim jRowCnt As Integer = 0
        Dim iAddFlg As Integer = 0

        If iAddCnt > sValue.Length Then
            GoTo Error_Exit
        End If

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
                            For klooop = 0 To iAddCnt - 1
                                '行追加
                                sData = sData & sValue(klooop) & vbCrLf
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

        fncTagKeyAdd = sData
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '   タグの項目でスペース以降の値を更新する
    '*******************************************************
    Public Function fncTagKeyUpdate(ByVal sBlockData As String, ByVal sTagName As String, sValue As String) As String
        On Error GoTo Error_Rtn
        fncTagKeyUpdate = ""

        Dim sRowData As String = ""                 '行データ
        Dim sColData As String = ""                 '1文字データ
        Dim sData As String = ""                    'ブロックデータ
        Dim iLOOP As Integer
        Dim jLOOP As Integer
        Dim iRowCnt As Integer = 0
        Dim jRowCnt As Integer = 0

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
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '   Itiファイルからtxrx用配列を作成する
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

    '*******************************************************
    '      'プロジェクト名.txrxファイルを削除する
    '*******************************************************
    Public Function fncFileDelete_pj_txrx() As Boolean
        fncFileDelete_pj_txrx = False

        Dim sFile As String = ""

        sFile = pPjPath & "\" & pPjName & ".txrx"
        fncFileDel(sFile)

        fncFileDelete_pj_txrx = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '      'スタート前にtxrxファイルを削除する
    '*******************************************************
    Public Function fncFileDelete_txrx() As Boolean
        fncFileDelete_txrx = False

        Dim sFile As String = ""

        sFile = pPjPath & "\" & "*.txrx"
        fncFileDel(sFile)

        fncFileDelete_txrx = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function
End Module