'*******************************************************
'      ファイル用モジュール
'*******************************************************
Module mduFile
    Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (
              ByVal lpApplicationName As String,
              ByVal lpKeyName As String,
              ByVal nDefault As Integer,
              ByVal lpFileName As String) As Integer

    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (
              ByVal lpApplicationName As String,
              ByVal lpKeyName As String,
              ByVal lpDefault As String,
              ByVal lpReturnedString As String,
              ByVal nSize As Integer,
              ByVal lpFileName As String) As Integer

    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (
        ByVal lpApplicationName As String,
              ByVal lpKeyName As String,
              ByVal lpString As String,
              ByVal lpFileName As String) As Integer

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' P2mファイルを配列にセット
    ''' </summary>
    ''' <param name="sFileName">P2mファイルのパス</param>
    ''' <returns>取得した行数</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFileSort(ByVal sFileName() As String) As String()
        Dim sFileName_2() As String

        ' StreamReader の新しいインスタンスを生成する
        Dim oReader_1 As New System.IO.StreamReader(sFileName(0), System.Text.Encoding.Default)     'ファイル読込用
        Dim oReader_2 As New System.IO.StreamReader(sFileName(1), System.Text.Encoding.Default)     'ファイル読込用
        Dim oReader_3 As New System.IO.StreamReader(sFileName(2), System.Text.Encoding.Default)     'ファイル読込用
        Dim oReader_4 As New System.IO.StreamReader(sFileName(3), System.Text.Encoding.Default)     'ファイル読込用
        Dim oReader_5 As New System.IO.StreamReader(sFileName(4), System.Text.Encoding.Default)     'ファイル読込用
        Dim oReader_6 As New System.IO.StreamReader(sFileName(5), System.Text.Encoding.Default)     'ファイル読込用
        Dim oReader_7 As New System.IO.StreamReader(sFileName(6), System.Text.Encoding.Default)     'ファイル読込用
        Dim oReader_8 As New System.IO.StreamReader(sFileName(7), System.Text.Encoding.Default)     'ファイル読込用
        Dim oReader_9 As New System.IO.StreamReader(sFileName(8), System.Text.Encoding.Default)     'ファイル読込用
        Dim sRowData_1 As String = ""                                                               '行データ
        Dim sRowData_2 As String = ""                                                               '行データ
        Dim sRowData_3 As String = ""                                                               '行データ
        Dim sRowData_4 As String = ""                                                               '行データ
        Dim sRowData_5 As String = ""                                                               '行データ
        Dim sRowData_6 As String = ""                                                               '行データ
        Dim sRowData_7 As String = ""                                                               '行データ
        Dim sRowData_8 As String = ""                                                               '行データ
        Dim sRowData_9 As String = ""                                                               '行データ
        Dim iColData(9) As Integer                                                                  'カラムデータ
        Dim iSort(9) As Integer                                                                  'カラムデータ
        Dim iLoop As Integer

        sFileName_2 = sFileName.Clone

        sRowData_1 = oReader_1.ReadLine()   '1行を文字型配列に格納
        sRowData_2 = oReader_2.ReadLine()   '1行を文字型配列に格納
        sRowData_3 = oReader_3.ReadLine()   '1行を文字型配列に格納
        sRowData_4 = oReader_4.ReadLine()   '1行を文字型配列に格納
        sRowData_5 = oReader_5.ReadLine()   '1行を文字型配列に格納
        sRowData_6 = oReader_6.ReadLine()   '1行を文字型配列に格納
        sRowData_7 = oReader_7.ReadLine()   '1行を文字型配列に格納
        sRowData_8 = oReader_8.ReadLine()   '1行を文字型配列に格納
        sRowData_9 = oReader_9.ReadLine()   '1行を文字型配列に格納

        iColData(0) = Mid(sRowData_1, sRowData_1.IndexOf("m") - 2, 3)
        iColData(1) = Mid(sRowData_2, sRowData_1.IndexOf("m") - 2, 3)
        iColData(2) = Mid(sRowData_3, sRowData_1.IndexOf("m") - 2, 3)
        iColData(3) = Mid(sRowData_4, sRowData_1.IndexOf("m") - 2, 3)
        iColData(4) = Mid(sRowData_5, sRowData_1.IndexOf("m") - 2, 3)
        iColData(5) = Mid(sRowData_6, sRowData_1.IndexOf("m") - 2, 3)
        iColData(6) = Mid(sRowData_7, sRowData_1.IndexOf("m") - 2, 3)
        iColData(7) = Mid(sRowData_8, sRowData_1.IndexOf("m") - 2, 3)
        iColData(8) = Mid(sRowData_9, sRowData_1.IndexOf("m") - 2, 3)

        For iLoop = 0 To 8
            If iColData(iLoop) = 10 Then
                sFileName(0) = sFileName_2(iLoop)
            ElseIf iColData(iLoop) = 15 Then
                sFileName(1) = sFileName_2(iLoop)
            ElseIf iColData(iLoop) = 20 Then
                sFileName(2) = sFileName_2(iLoop)
            ElseIf iColData(iLoop) = 25 Then
                sFileName(3) = sFileName_2(iLoop)
            ElseIf iColData(iLoop) = 30 Then
                sFileName(4) = sFileName_2(iLoop)
            ElseIf iColData(iLoop) = 35 Then
                sFileName(5) = sFileName_2(iLoop)
            ElseIf iColData(iLoop) = 40 Then
                sFileName(6) = sFileName_2(iLoop)
            ElseIf iColData(iLoop) = 45 Then
                sFileName(7) = sFileName_2(iLoop)
            ElseIf iColData(iLoop) = 50 Then
                sFileName(8) = sFileName_2(iLoop)
            End If
        Next

        oReader_1.Close()
        oReader_1.Dispose()
        oReader_2.Close()
        oReader_2.Dispose()
        oReader_3.Close()
        oReader_3.Dispose()
        oReader_4.Close()
        oReader_4.Dispose()
        oReader_5.Close()
        oReader_5.Dispose()
        oReader_6.Close()
        oReader_6.Dispose()
        oReader_7.Close()
        oReader_7.Dispose()
        oReader_8.Close()
        oReader_8.Dispose()
        oReader_9.Close()
        oReader_9.Dispose()

        fncFileSort = sFileName

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' フォルダ内P2mファイル名を取得する
    ''' </summary>
    ''' <param name="sDirName">P2mファイルのパス</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncGetDirName_P2m(ByVal sDirName As String) As String()
        Dim sFiles2 As String() = System.IO.Directory.GetFiles(
            sDirName, "*.p2m", System.IO.SearchOption.AllDirectories)

        Dim sFiles() As String
        Dim iLoop As Integer
        Dim iCnt As Integer = 0

        ReDim Preserve sFiles(-1)

        For iLoop = 0 To sFiles2.Length - 1
            If sFiles2(iLoop).Contains("power") Then
                ReDim Preserve sFiles(iCnt)
                sFiles(iLoop) = sFiles2(iLoop)

                iCnt = iCnt + 1
            End If
        Next

        fncGetDirName_P2m = sFiles

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' フォルダ内P2mファイル数を取得する
    ''' </summary>
    ''' <param name="sDirName">P2mファイルのパス</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncGetDir_P2m(ByVal sDirName As String) As Integer
        Dim fileCount As Integer = System.IO.Directory.GetFiles(sDirName,
                                                                "*.p2m",
                                                                System.IO.SearchOption.TopDirectoryOnly).Length

        fncGetDir_P2m = fileCount

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Power_P2mファイルを配列にセット
    ''' </summary>
    ''' <returns>OK:True NG:False</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFile2Power() As Boolean

        fncFile2Power = False

        ' StreamReader の新しいインスタンスを生成する
        Dim oReader As New System.IO.StreamReader("Power.txt", System.Text.Encoding.Default)  'ファイル読込用
        Dim iRow As Integer = 0
        Dim sRowData As String = ""

        '行ループ
        While (oReader.Peek() >= 0)
            sRowData = oReader.ReadLine()   '1行を文字型配列に格納
            '列ループ：１行分の列を埋める
            pPower(iRow) = sRowData
            iRow += 1
        End While '行ループ

        oReader.Close()
        oReader.Dispose()

        fncFile2Power = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sabunファイルを配列にセット
    ''' </summary>
    ''' <returns>OK:True NG:False</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFile2Sabun() As Boolean

        fncFile2Sabun = False

        ' StreamReader の新しいインスタンスを生成する
        Dim oReader As New System.IO.StreamReader("x_y_sabun.csv", System.Text.Encoding.Default)  'ファイル読込用
        Dim sRowData As String = ""                                                         '行データ
        Dim sColData As String = ""                                                         'カラムデータ
        Dim iRow As Integer = 0                                                             '行
        Dim iCol As Integer = 0                                                             '列
        Dim iColNo As Integer = 0                                                           '列No
        Dim iNowFlg As String = 0                                                           '現在行
        Dim iLastFlg As String = 0                                                          '最終行

        '行ループ
        While (oReader.Peek() >= 0)
            sRowData = oReader.ReadLine()   '1行を文字型配列に格納
            '列ループ：１行分の列を埋める
            iColNo = 0
            sColData = ""
            For iCol = 0 To sRowData.Length - 1
                If sRowData.Substring(iCol, 1) = "," Then
                    iNowFlg = 1
                    iColNo = iColNo + 1
                Else
                    iNowFlg = 0
                End If

                If iNowFlg = 1 Then
                    Select Case iColNo
                        Case 1
                            pSabun(iRow).sX = Trim(sColData)
                        Case 2
                            pSabun(iRow).sY = Trim(sColData)
                    End Select
                    iNowFlg = 0
                    sColData = ""
                Else
                    sColData = sColData + sRowData.Substring(iCol, 1)
                End If
            Next '列ループ
            pSabun(iRow).sSabun = Trim(sColData)
            iRow += 1
        End While '行ループ

        oReader.Close()
        oReader.Dispose()

        fncFile2Sabun = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' P2mファイルを配列にセット
    ''' </summary>
    ''' <param name="iRowCnt">取得した行数</param>
    ''' <param name="sFileName">P2mファイルのパス</param>
    ''' <returns>取得した行数</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFile2P2m(ByVal iRowCnt As Integer, ByVal sFileName As String) As Integer

        fncFile2P2m = 0

        ' StreamReader の新しいインスタンスを生成する
        Dim oReader As New System.IO.StreamReader(sFileName, System.Text.Encoding.Default)  'ファイル読込用
        Dim sRowData As String = ""                                                         '行データ
        Dim sColData As String = ""                                                         'カラムデータ
        Dim iRow As Integer = iRowCnt                                                             '行
        Dim iCol As Integer = 0                                                             '列
        Dim iColNo As Integer = 0                                                           '列No
        Dim iNowFlg As String = 0                                                           '現在行
        Dim iLastFlg As String = 0                                                          '最終行

        '行ループ
        While (oReader.Peek() >= 0)
            ReDim Preserve pP2m(iRow)       'pP2m に新規行を追加
            sRowData = oReader.ReadLine()   '1行を文字型配列に格納
            If Left(sRowData, 1) <> "#" Then
                '列ループ：１行分の列を埋める
                iColNo = 0
                sColData = ""
                For iCol = 0 To sRowData.Length - 1
                    If iColNo < 6 Then
                        iLastFlg = iNowFlg
                        If sRowData.Substring(iCol, 1) = " " Then
                            iNowFlg = 0
                        Else
                            iNowFlg = 1
                        End If

                        If (iLastFlg = 1 AndAlso iNowFlg = 0) Then
                            Select Case iColNo
                                Case 0
                                    pP2m(iRow).sRx = Trim(sColData)
                                Case 1
                                    pP2m(iRow).sX = Trim(sColData)
                                Case 2
                                    pP2m(iRow).sY = Trim(sColData)
                                Case 3
                                    pP2m(iRow).sZ = Trim(sColData)
                                Case 4
                                    pP2m(iRow).sDistance = Trim(sColData)
                                Case 5
                                    pP2m(iRow).sPower = Trim(sColData)
                            End Select
                            iColNo = iColNo + 1
                            sColData = ""
                        Else
                            sColData = sColData + sRowData.Substring(iCol, 1)
                        End If
                    Else
                        'Case 6 行が終了するため
                        sColData = sColData + sRowData.Substring(iCol, 1)
                        pP2m(iRow).sPhase = Trim(sColData)
                    End If
                Next '列ループ
                iRow += 1
            End If
        End While '行ループ

        oReader.Close()
        oReader.Dispose()

        fncFile2P2m = iRow

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Power.txtを作成する
    ''' </summary>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncMem2Power() As Boolean

        fncMem2Power = False

        Dim iRow As Integer '行ループ
        '存在すればファイル削除
        fncFileDel(pExePath & "\Power.txt")

        If pP2m.Length >= 1 Then
            'Power.txtファイル作成
            'Dim oFileWrite As New System.IO.StreamWriter(pExePath & "\Power.txt", True, System.Text.Encoding.UTF8)
            Dim oFileWrite As New System.IO.StreamWriter(pExePath & "\Power.txt", True, pEnc)

            pP2m(iRow).sPower = fncTagLf(pP2m(iRow).sPower)

            For iRow = 0 To pP2m.Length - 1
                oFileWrite.WriteLine(pP2m(iRow).sPower)
            Next

            'クローズ
            oFileWrite.Dispose()
            oFileWrite.Close()

            fncMem2Power = True
        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 履歴用130ファイルにコピー
    ''' </summary>
    ''' <param name="m">ドローン</param>
    ''' <param name="n">周回</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncCopy130(ByVal m As Integer, ByVal n As Integer) As Boolean

        fncCopy130 = False

        Dim iErrCnt As Integer = 0  'エラーカウント
        Dim sFile As String = ""    'ファイル名

        '出力元ファイル存在チェック
        If fncFileCheck(pc130_Grid) AndAlso
            fncFileCheck(pc130_Meter) Then
            'リネームして指定フォルダにコピー
            sFile = pPjPath & "\" & "d" & m & "_v" & n & "_" & pc130_Grid
            If fncFileCopy(pc130_Grid, sFile) = False Then
                'コピーエラー
                iErrCnt = iErrCnt + 1
            End If

            'リネームして指定フォルダにコピー
            sFile = pPjPath & "\" & "d" & m & "_v" & n & "_" & pc130_Meter
            If fncFileCopy(pc130_Meter, sFile) = False Then
                'コピーエラー
                iErrCnt = iErrCnt + 1
            End If
        End If

        If iErrCnt = 0 Then
            fncCopy130 = True
        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' フォルダを削除する
    ''' </summary>
    ''' <param name="sDirName">削除フォルダ名(\は付けない)</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncDirDel(ByVal sDirName As String) As Boolean

        fncDirDel = False

        If System.IO.Directory.Exists(sDirName) Then
            'フォルダ削除
            System.IO.Directory.Delete(sDirName, True)
        End If

        fncDirDel = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' フォルダをコピーする
    ''' </summary>
    ''' <param name="sDirName1">コピー元フォルダ名(\は付けない)</param>
    ''' <param name="sDirName2">コピー先フォルダ名(\は付けない)</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncDirCopy(ByVal sDirName1 As String, ByVal sDirName2 As String) As Boolean

        fncDirCopy = False

        'コピー先のディレクトリがないときは作る
        If Not System.IO.Directory.Exists(sDirName2) Then
            System.IO.Directory.CreateDirectory(sDirName2)
            '属性もコピー
            System.IO.File.SetAttributes(sDirName2,
            System.IO.File.GetAttributes(sDirName1))
        End If
        'フォルダコピー
        My.Computer.FileSystem.CopyDirectory(sDirName1, sDirName2,
            FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)

        fncDirCopy = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ファイルの存在をチェックする
    ''' </summary>
    ''' <param name="sFileName">ファイル名</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFileCheck(ByVal sFileName As String) As Boolean

        fncFileCheck = False

        fncFileCheck = System.IO.File.Exists(sFileName)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ファイルを削除する
    ''' </summary>
    ''' <param name="sFileName">削除ファイル名</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFileDel(ByVal sFileName As String) As Boolean

        fncFileDel = False

        'ファイル存在チェック
        If fncFileCheck(sFileName) Then
            FileSystem.Kill(sFileName)
        End If

        fncFileDel = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ワイルドカードでファイルを削除する
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subFileDel_w(ByVal sFileName As String)
        On Error Resume Next

        FileSystem.Kill(sFileName)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ファイルを上書きコピーする
    ''' </summary>
    ''' <param name="sFileName1">コピー元ファイル名</param>
    ''' <param name="sFileName2">コピー先ファイル名</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFileCopy(ByVal sFileName1 As String, ByVal sFileName2 As String) As Boolean

        fncFileCopy = False

        System.IO.File.Copy(sFileName1, sFileName2, True)

        fncFileCopy = True

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' INIファイルを読込する
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subSetIni()
        Const DEF_STR As String = vbNullString  'Null設定
        Dim iRET As Integer = 0
        Dim sBuf As String = New String(" ", 1024) 'Spaceが1024文字

        'INIファイルのパスはカレントフォルダ固定
        pIniPath = ""
        pIniPath = System.IO.Directory.GetCurrentDirectory()

        'v：ドローン速度
        pV = ""
        iRET = GetPrivateProfileString(pcSec_Route, pcKey_v, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pV = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))

        't：周回毎経過時間
        pT = ""
        iRET = GetPrivateProfileString(pcSec_Route, pcKey_t, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pT = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))

        'ファイル切替
        pFileFlg = 0
        iRET = GetPrivateProfileString(pcSec_FileFlg, pcKey_sw, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pFileFlg = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        'Exe1_Path
        pExe1_Path = ""
        iRET = GetPrivateProfileString(pcSec_FileFlg, pcKey_Exe1, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pExe1_Path = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        'Exe2_Path
        pExe2_Path = ""
        iRET = GetPrivateProfileString(pcSec_FileFlg, pcKey_Exe2, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pExe2_Path = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        'Exe22
        pExe22 = ""
        iRET = GetPrivateProfileString(pcSec_FileFlg, pcKey_Exe22, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pExe22 = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))

        'PjName
        pPjName = ""
        iRET = GetPrivateProfileString(pcSec_Set, pcKey_11, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pPjName = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        'GnuPath
        pGnuPath = ""
        iRET = GetPrivateProfileString(pcSec_Set, pcKey_12, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pGnuPath = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        'PjPath
        pPjPath = ""
        iRET = GetPrivateProfileString(pcSec_Set, pcKey_13, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pPjPath = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        'p2mDir
        pP2mDir = ""
        iRET = GetPrivateProfileString(pcSec_Set, pcKey_14, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pP2mDir = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))

        frmMain.txtV.Text = pV                                  'ドローン速度
        frmMain.txtT.Text = pT                                  '周回毎経過時間
        frmMain.txtVT.Text = CDbl(pV) * CDbl(pT)                '飛距離FD

        frmMain.cmbFileFlg.Text = pFileFlg      'ファイル切替フラグ
        frmMain.txtExe1.Text = pExe1_Path       'Exe1_Path
        frmMain.txtExe2.Text = pExe2_Path       'Exe2_Path
        frmMain.txtExe22.Text = pExe22          'Exe22

        frmMain.txtIniPath.Text = pIniPath      'INIファイルパス
        frmMain.txtPjName.Text = pPjName        'プロジェクト名
        frmMain.txtGnuPath.Text = pGnuPath      'gnuプロットパス
        frmMain.txtPjPath.Text = pPjPath        'pjフォルダパス
        frmMain.txtP2mDir.Text = pP2mDir        'プロジェクト名

        Dim iCnt As Integer 'ドローンループ
        Dim sSec As String  'セクション名
        Dim sX As String    'X値
        Dim sY As String    'Y値
        Dim sZ As String    'Z値

        For iCnt = 1 To 5
            sX = ""
            sY = ""
            sZ = ""

            sSec = "d" & CInt(iCnt)
            'X
            iRET = GetPrivateProfileString(sSec, pcKey_X, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
            sX = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
            'Y
            iRET = GetPrivateProfileString(sSec, pcKey_Y, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
            sY = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
            'Z
            iRET = GetPrivateProfileString(sSec, pcKey_Z, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
            sZ = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))

            Select Case iCnt
                Case 1
                    frmMain.txtX_d1.Text = sX            'X設定値
                    frmMain.txtY_d1.Text = sY            'Y設定値
                    frmMain.txtZ_d1.Text = sZ            'Z設定値

                    'pX_d1 = sX
                    'pY_d1 = sY
                    'pZ_d1 = sZ
                Case 2
                    frmMain.txtX_d2.Text = sX            'X設定値
                    frmMain.txtY_d2.Text = sY            'Y設定値
                    frmMain.txtZ_d2.Text = sZ            'Z設定値

                    'pX_d2 = sX
                    'pY_d2 = sY
                    'pZ_d2 = sZ
                Case 3
                    frmMain.txtX_d3.Text = sX            'X設定値
                    frmMain.txtY_d3.Text = sY            'Y設定値
                    frmMain.txtZ_d3.Text = sZ            'Z設定値

                    'pX_d3 = sX
                    'pY_d3 = sY
                    'pZ_d3 = sZ
                Case 4
                    frmMain.txtX_d4.Text = sX            'X設定値
                    frmMain.txtY_d4.Text = sY            'Y設定値
                    frmMain.txtZ_d4.Text = sZ            'Z設定値

                    'pX_d4 = sX
                    'pY_d4 = sY
                    'pZ_d4 = sZ
                Case 5
                    frmMain.txtX_d5.Text = sX            'X設定値
                    frmMain.txtY_d5.Text = sY            'Y設定値
                    frmMain.txtZ_d5.Text = sZ            'Z設定値

                    'pX_d5 = sX
                    'pY_d5 = sY
                    'pZ_d5 = sZ
            End Select
        Next

        'Command1-5
        For iCnt = 0 To 4
            pCommand(iCnt) = ""
        Next

        iRET = GetPrivateProfileString(pcSec_Command, pcKey_1, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pCommand(0) = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        iRET = GetPrivateProfileString(pcSec_Command, pcKey_2, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pCommand(1) = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        iRET = GetPrivateProfileString(pcSec_Command, pcKey_3, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pCommand(2) = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        iRET = GetPrivateProfileString(pcSec_Command, pcKey_4, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pCommand(3) = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        iRET = GetPrivateProfileString(pcSec_Command, pcKey_5, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pcIniFileName)
        pCommand(4) = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))

        frmMain.txtCom1.Text = pCommand(0)            'コマンド
        frmMain.txtCom2.Text = pCommand(1)            'コマンド
        frmMain.txtCom3.Text = pCommand(2)            'コマンド
        frmMain.txtCom4.Text = pCommand(3)            'コマンド
        frmMain.txtCom5.Text = pCommand(4)            'コマンド

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' INIファイルを保存する
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subPutIni()
        Dim iRET As Integer = 0 'リターン値

        iRET = WritePrivateProfileString(pcSec_Route, pcKey_v, frmMain.txtV.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Route, pcKey_t, frmMain.txtT.Text, pIniPath & "\" & pcIniFileName)

        iRET = WritePrivateProfileString(pcSec_FileFlg, pcKey_sw, frmMain.cmbFileFlg.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_FileFlg, pcKey_Exe1, frmMain.txtExe1.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_FileFlg, pcKey_Exe2, frmMain.txtExe2.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_FileFlg, pcKey_Exe22, frmMain.txtExe22.Text, pIniPath & "\" & pcIniFileName)

        iRET = WritePrivateProfileString(pcSec_Set, pcKey_11, frmMain.txtPjName.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Set, pcKey_12, frmMain.txtGnuPath.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Set, pcKey_13, frmMain.txtPjPath.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Set, pcKey_14, frmMain.txtP2mDir.Text, pIniPath & "\" & pcIniFileName)

        iRET = WritePrivateProfileString(pcSec_d1, pcKey_X, frmMain.txtX_d1.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d1, pcKey_Y, frmMain.txtY_d1.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d1, pcKey_Z, frmMain.txtZ_d1.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d2, pcKey_X, frmMain.txtX_d2.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d2, pcKey_Y, frmMain.txtY_d2.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d2, pcKey_Z, frmMain.txtZ_d2.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d3, pcKey_X, frmMain.txtX_d3.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d3, pcKey_Y, frmMain.txtY_d3.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d3, pcKey_Z, frmMain.txtZ_d3.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d4, pcKey_X, frmMain.txtX_d4.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d4, pcKey_Y, frmMain.txtY_d4.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d4, pcKey_Z, frmMain.txtZ_d4.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d5, pcKey_X, frmMain.txtX_d5.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d5, pcKey_Y, frmMain.txtY_d5.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_d5, pcKey_Z, frmMain.txtZ_d5.Text, pIniPath & "\" & pcIniFileName)

        iRET = WritePrivateProfileString(pcSec_Command, pcKey_1, frmMain.txtCom1.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Command, pcKey_2, frmMain.txtCom2.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Command, pcKey_3, frmMain.txtCom3.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Command, pcKey_4, frmMain.txtCom4.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Command, pcKey_5, frmMain.txtCom5.Text, pIniPath & "\" & pcIniFileName)

    End Sub
End Module
