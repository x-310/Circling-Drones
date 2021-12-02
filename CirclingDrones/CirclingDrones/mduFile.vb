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

    '*******************************************************
    '   P2mファイル取込み
    '*******************************************************
    Public Function fncReadP2m(ByVal sFileName As String)
        On Error GoTo Error_Rtn
        fncReadP2m = False

        ' StreamReader の新しいインスタンスを生成する
        Dim oReader As New System.IO.StreamReader(sFileName, System.Text.Encoding.Default)
        Dim sRowData As String = ""
        Dim sColData As String = ""
        Dim iRow As Integer = 0
        Dim iCol As Integer = 0
        Dim iColNo As Integer = 0
        Dim iNowFlg As String = 0
        Dim iLastFlg As String = 0
        '▼▼ 行ループ
        ReDim Preserve pP2m(-1)
        While (oReader.Peek() >= 0)
            ReDim Preserve pP2m(iRow)             'pP2m に新規行を追加
            sRowData = oReader.ReadLine()         '1行を文字型配列に格納
            If Left(sRowData, 1) <> "#" Then
                '▼ 列ループ：１行分の列を埋める
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
                Next '▲ 列ループ
                iRow += 1
            End If
        End While '▲▲ 行ループ

        oReader.Close()
        oReader.Dispose()

        fncReadP2m = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '   メモリからPower.txtを作成する
    '*******************************************************
    Public Function fncMem2Power(ByVal sPath As String) As Boolean
        On Error GoTo Error_Rtn
        fncMem2Power = False

        Dim iRow As Integer
        'ファイル存在チェック
        If fncFileCheck(sPath & "\Power.txt") Then
            '存在すればファイル削除
            fncFileDel(sPath & "\Power.txt")
        End If

        If Not IsNothing(pP2m) Then
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
    '      'txrxファイルを削除する
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

    '*******************************************************
    '      'ドローン毎の周回用ファイルを削除する
    '*******************************************************
    Public Function fncFileDelete_d() As Boolean
        fncFileDelete_d = False

        Dim sFile As String = ""

        sFile = pPjPath & "\" & "d*.*"
        fncFileDel(sFile)

        fncFileDelete_d = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '      3d_pot_test_plus.exeの出力ファイルを
    '      リネームして指定フォルダにコピーする
    '*******************************************************
    Public Function fncExtOutRename(ByVal m As Integer, ByVal n As Integer) As Boolean
        On Error GoTo Error_Rtn
        fncExtOutRename = False

        Dim iErrCnt As Integer = 0
        Dim sFile As String = ""

        '      End If
        '出力元ファイル存在チェック
        If fncFileCheck(pExt_Out_Grid) AndAlso
            fncFileCheck(pExt_Out_Meter) Then
            'fncFileCheck(pExt_Out_Height) AndAlso
            'fncFileCheck(pExt_Out_Power) Then
            '①リネームして指定フォルダにコピー
            sFile = pPjPath & "\" & "d" & m & "_v" & n & "_" & pExt_Out_Grid
            If fncFileCopy(pExt_Out_Grid, sFile) = False Then
                'コピーエラー
                iErrCnt = iErrCnt + 1
            End If
            '②リネームして指定フォルダにコピー
            sFile = pPjPath & "\" & "d" & m & "_v" & n & "_" & pExt_Out_Meter
            If fncFileCopy(pExt_Out_Meter, sFile) = False Then
                'コピーエラー
                iErrCnt = iErrCnt + 1
            End If
            '③リネームして指定フォルダにコピー
            'sFile = pPjPath & "\" & "d" & m & "_v" & n & "_" & pExt_Out_Height
            'If fncFileCopy(pExt_Out_Height, sFile) = False Then
            '    'コピーエラー
            '    iErrCnt = iErrCnt + 1
            'End If
            '④リネームして指定フォルダにコピー
            'sFile = pPjPath & "\" & "d" & m & "_v" & n & "_" & pExt_Out_Power
            'If fncFileCopy(pExt_Out_Power, sFile) = False Then
            '    'コピーエラー
            '    iErrCnt = iErrCnt + 1
            'End If
        End If

        If iErrCnt <> 0 Then GoTo Error_Exit
        fncExtOutRename = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '      ファイルの存在チェック
    '*******************************************************
    Public Function fncFileCheck(ByVal fileName As String) As Boolean
        On Error GoTo Error_Rtn
        fncFileCheck = False

        fncFileCheck = System.IO.File.Exists(fileName)
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '      ファイルの削除
    '*******************************************************
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

    '*******************************************************
    '      ファイルの上書きコピー
    '*******************************************************
    Public Function fncFileCopy(ByVal fileName1 As String, ByVal fileName2 As String) As Boolean
        On Error GoTo Error_Rtn
        fncFileCopy = False

        System.IO.File.Copy(fileName1, fileName2, True)

        fncFileCopy = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    '*******************************************************
    '   INIファイル読込
    '*******************************************************
    Public Sub subSetIni()
        On Error GoTo Error_Rtn

        Const DEF_STR As String = vbNullString
        Const DEF_VAL As Long = 0

        Dim iRET As Integer = 0
        Dim sBuf As String = New String(" ", 1024) 'Spaceが1024文字
        Dim sFileName As String

        'iniファイルのパスはカレントフォルダ固定
        pIniPath = ""
        pIniPath = System.IO.Directory.GetCurrentDirectory()

        'PjName
        pPjName = ""
        iRET = GetPrivateProfileString(pSec_Set, pKey_11, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
        pPjName = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        'GnuPath
        pGnuPath = ""
        iRET = GetPrivateProfileString(pSec_Set, pKey_12, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
        pGnuPath = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        'PjPath
        pPjPath = ""
        iRET = GetPrivateProfileString(pSec_Set, pKey_13, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
        pPjPath = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))

        frmMain.txtIniPath.Text = pIniPath      'iniファイルパス
        frmMain.txtPjName.Text = pPjName        'プロジェクト名
        frmMain.txtGnuPath.Text = pGnuPath      'gnuプロットパス
        frmMain.txtPjPath.Text = pPjPath        'pjフォルダパス

        Dim iCnt As Integer
        Dim sX As String
        Dim sY As String
        Dim sZ As String
        Dim sSec As String
        For iCnt = 1 To 5
            sX = ""
            sY = ""
            sZ = ""

            sSec = "d" & CInt(iCnt)
            'X
            iRET = GetPrivateProfileString(sSec, pKey_X, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
            sX = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
            'Y
            iRET = GetPrivateProfileString(sSec, pKey_Y, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
            sY = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
            'Z
            iRET = GetPrivateProfileString(sSec, pKey_Z, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
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

        iRET = GetPrivateProfileString(pSec_Command, pKey_1, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
        pCommand(0) = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        iRET = GetPrivateProfileString(pSec_Command, pKey_2, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
        pCommand(1) = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        iRET = GetPrivateProfileString(pSec_Command, pKey_3, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
        pCommand(2) = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        iRET = GetPrivateProfileString(pSec_Command, pKey_4, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
        pCommand(3) = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))
        iRET = GetPrivateProfileString(pSec_Command, pKey_5, DEF_STR, sBuf, sBuf.Length, pIniPath & "\" & pIniFileName)
        pCommand(4) = sBuf.Substring(0, sBuf.IndexOf(vbNullChar))

        frmMain.txtCom1.Text = pCommand(0)            'コマンド
        frmMain.txtCom2.Text = pCommand(1)            'コマンド
        frmMain.txtCom3.Text = pCommand(2)            'コマンド
        frmMain.txtCom4.Text = pCommand(3)            'コマンド
        frmMain.txtCom5.Text = pCommand(4)            'コマンド
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   INIファイル保存
    '*******************************************************
    Public Sub subPutIni()
        On Error GoTo Error_Rtn

        Dim iRET As Integer = 0

        iRET = WritePrivateProfileString(pSec_Set, pKey_11, frmMain.txtPjName.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_Set, pKey_12, frmMain.txtGnuPath.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_Set, pKey_13, frmMain.txtPjPath.Text, pIniPath & "\" & pIniFileName)

        iRET = WritePrivateProfileString(pSec_d1, pKey_X, frmMain.txtX_d1.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d1, pKey_Y, frmMain.txtY_d1.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d1, pKey_Z, frmMain.txtZ_d1.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d2, pKey_X, frmMain.txtX_d2.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d2, pKey_Y, frmMain.txtY_d2.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d2, pKey_Z, frmMain.txtZ_d2.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d3, pKey_X, frmMain.txtX_d3.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d3, pKey_Y, frmMain.txtY_d3.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d3, pKey_Z, frmMain.txtZ_d3.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d4, pKey_X, frmMain.txtX_d4.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d4, pKey_Y, frmMain.txtY_d4.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d4, pKey_Z, frmMain.txtZ_d4.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d5, pKey_X, frmMain.txtX_d5.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d5, pKey_Y, frmMain.txtY_d5.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_d5, pKey_Z, frmMain.txtZ_d5.Text, pIniPath & "\" & pIniFileName)

        iRET = WritePrivateProfileString(pSec_Command, pKey_1, frmMain.txtCom1.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_Command, pKey_2, frmMain.txtCom2.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_Command, pKey_3, frmMain.txtCom3.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_Command, pKey_4, frmMain.txtCom4.Text, pIniPath & "\" & pIniFileName)
        iRET = WritePrivateProfileString(pSec_Command, pKey_5, frmMain.txtCom5.Text, pIniPath & "\" & pIniFileName)
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub
End Module
