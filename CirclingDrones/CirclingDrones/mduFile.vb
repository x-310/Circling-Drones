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
    '   Power.txtを作成する
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
        If fncFileCheck(pcExt_Out_Grid) AndAlso
            fncFileCheck(pcExt_Out_Meter) Then
            'fncFileCheck(pcExt_Out_Height) AndAlso
            'fncFileCheck(pcExt_Out_Power) Then
            '①リネームして指定フォルダにコピー
            sFile = pPjPath & "\" & "d" & m & "_v" & n & "_" & pcExt_Out_Grid
            If fncFileCopy(pcExt_Out_Grid, sFile) = False Then
                'コピーエラー
                iErrCnt = iErrCnt + 1
            End If
            '②リネームして指定フォルダにコピー
            sFile = pPjPath & "\" & "d" & m & "_v" & n & "_" & pcExt_Out_Meter
            If fncFileCopy(pcExt_Out_Meter, sFile) = False Then
                'コピーエラー
                iErrCnt = iErrCnt + 1
            End If
            '③リネームして指定フォルダにコピー
            'sFile = pPjPath & "\" & "d" & m & "_v" & n & "_" & pcExt_Out_Height
            'If fncFileCopy(pcExt_Out_Height, sFile) = False Then
            '    'コピーエラー
            '    iErrCnt = iErrCnt + 1
            'End If
            '④リネームして指定フォルダにコピー
            'sFile = pPjPath & "\" & "d" & m & "_v" & n & "_" & pcExt_Out_Power
            'If fncFileCopy(pcExt_Out_Power, sFile) = False Then
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
        GoTo Error_Exit
    End Sub

    '*******************************************************
    '   INIファイル保存
    '*******************************************************
    Public Sub subPutIni()
        On Error GoTo Error_Rtn

        Dim iRET As Integer = 0

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
