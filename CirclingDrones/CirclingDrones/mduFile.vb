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
    ''' P2mファイルを取込みする
    ''' </summary>
    ''' <param name="sFileName">P2mファイルのパス</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFile2P2m(ByVal sFileName As String) As Boolean
        On Error GoTo Error_Rtn
        fncFile2P2m = False

        ' StreamReader の新しいインスタンスを生成する
        Dim oReader As New System.IO.StreamReader(sFileName, System.Text.Encoding.Default)  'ファイル読込用
        Dim sRowData As String = ""                                                         '行データ
        Dim sColData As String = ""                                                         'カラムデータ
        Dim iRow As Integer = 0                                                             '行
        Dim iCol As Integer = 0                                                             '列
        Dim iColNo As Integer = 0                                                           '列No
        Dim iNowFlg As String = 0                                                           '現在行
        Dim iLastFlg As String = 0                                                          '最終行

        '行ループ
        ReDim Preserve pP2m(-1)
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

        fncFile2P2m = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Power.txtを作成する
    ''' </summary>
    ''' <param name="sPath">Power.txtのパス</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncMem2Power(ByVal sPath As String) As Boolean
        On Error GoTo Error_Rtn
        fncMem2Power = False

        Dim iRow As Integer '行ループ
        'ファイル存在チェック
        If fncFileCheck(sPath & "\Power.txt") Then
            '存在すればファイル削除
            fncFileDel(sPath & "\Power.txt")
        End If

        If pP2m.Length >= 1 Then
            'Power.txtファイル作成
            Dim oFileWrite As New System.IO.StreamWriter(sPath & "\Power.txt", True, System.Text.Encoding.UTF8)

            For iRow = 0 To pP2m.Length - 1
                oFileWrite.WriteLine(pP2m(iRow).sPower)
            Next

            'クローズ
            oFileWrite.Dispose()
            oFileWrite.Close()

            fncMem2Power = True
        End If
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ドローン毎の周回用ファイルを削除する
    ''' </summary>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncFileDelete_d() As Boolean
        fncFileDelete_d = False

        Dim sFile As String = ""    'ファイル名

        sFile = pPjPath & "\" & "d*.*"
        fncFileDel(sFile)

        fncFileDelete_d = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 出力ファイルをリネームして指定フォルダにコピーする
    ''' </summary>
    ''' <param name="m">ドローン</param>
    ''' <param name="n">周回</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fnc130Rename(ByVal m As Integer, ByVal n As Integer) As Boolean
        On Error GoTo Error_Rtn
        fnc130Rename = False

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

        If iErrCnt <> 0 Then GoTo Error_Exit
        fnc130Rename = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
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
        On Error GoTo Error_Rtn
        fncFileCheck = False

        fncFileCheck = System.IO.File.Exists(sFileName)
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
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
        On Error GoTo Error_Rtn
        fncFileDel = False

        FileSystem.Kill(sFileName)

        fncFileDel = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

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
        On Error GoTo Error_Rtn
        fncFileCopy = False

        System.IO.File.Copy(sFileName1, sFileName2, True)

        fncFileCopy = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
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
        On Error GoTo Error_Rtn

        Const DEF_STR As String = vbNullString  'Null設定

        Dim iRET As Integer = 0
        Dim sBuf As String = New String(" ", 1024) 'Spaceが1024文字
        Dim sFileName As String

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

        'vt：飛距離FD
        pVT = CInt(pV) * CInt(pT)

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

        frmMain.txtV.Text = pV                  'ドローン速度
        frmMain.txtT.Text = pT                  '周回毎経過時間
        frmMain.txtVT.Text = pVT                '飛距離FD

        frmMain.txtIniPath.Text = pIniPath      'INIファイルパス
        frmMain.txtPjName.Text = pPjName        'プロジェクト名
        frmMain.txtGnuPath.Text = pGnuPath      'gnuプロットパス
        frmMain.txtPjPath.Text = pPjPath        'pjフォルダパス

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

                    pX_d1 = sX
                    pY_d1 = sY
                    pZ_d1 = sZ
                Case 2
                    frmMain.txtX_d2.Text = sX            'X設定値
                    frmMain.txtY_d2.Text = sY            'Y設定値
                    frmMain.txtZ_d2.Text = sZ            'Z設定値

                    pX_d2 = sX
                    pY_d2 = sY
                    pZ_d2 = sZ
                Case 3
                    frmMain.txtX_d3.Text = sX            'X設定値
                    frmMain.txtY_d3.Text = sY            'Y設定値
                    frmMain.txtZ_d3.Text = sZ            'Z設定値

                    pX_d3 = sX
                    pY_d3 = sY
                    pZ_d3 = sZ
                Case 4
                    frmMain.txtX_d4.Text = sX            'X設定値
                    frmMain.txtY_d4.Text = sY            'Y設定値
                    frmMain.txtZ_d4.Text = sZ            'Z設定値

                    pX_d4 = sX
                    pY_d4 = sY
                    pZ_d4 = sZ
                Case 5
                    frmMain.txtX_d5.Text = sX            'X設定値
                    frmMain.txtY_d5.Text = sY            'Y設定値
                    frmMain.txtZ_d5.Text = sZ            'Z設定値

                    pX_d5 = sX
                    pY_d5 = sY
                    pZ_d5 = sZ
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
Error_Exit:
        Exit Sub
Error_Rtn:
        fncMsgBox("INIファイルを確認して下さい")
        GoTo Error_Exit
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
        On Error GoTo Error_Rtn

        Dim iRET As Integer = 0 'リターン値

        iRET = WritePrivateProfileString(pcSec_Route, pcKey_v, frmMain.txtV.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Route, pcKey_t, frmMain.txtT.Text, pIniPath & "\" & pcIniFileName)

        iRET = WritePrivateProfileString(pcSec_Set, pcKey_11, frmMain.txtPjName.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Set, pcKey_12, frmMain.txtGnuPath.Text, pIniPath & "\" & pcIniFileName)
        iRET = WritePrivateProfileString(pcSec_Set, pcKey_13, frmMain.txtPjPath.Text, pIniPath & "\" & pcIniFileName)

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
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub
End Module
