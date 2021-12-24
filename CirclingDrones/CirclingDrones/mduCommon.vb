'*******************************************************
'      共通関数用モジュール
'*******************************************************
Imports Microsoft.VisualBasic.FileIO

Module mduCommon

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 画面コントロールを初期化する
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subCtlReset()
        On Error GoTo Error_Rtn

        frmMain.cmbSet_d.Text = "2"
        frmMain.cmbSet_Interval.Text = "1000"
        frmMain.cmbSet_v.Text = "2"

        frmMain.txtNow_m.Text = ""      '現在ローンNo
        frmMain.txtNow_n.Text = ""      '現在周目
        frmMain.txtLogDisp.Text = ""    '画面ログ出力クリア

        frmMain.txtV.Text = ""          'ドローン速度
        frmMain.txtT.Text = ""          '周回毎経過時間
        frmMain.txtVT.Text = ""         '飛距離FD

        frmMain.txtIniPath.Text = ""    'INIファイルパス
        frmMain.txtPjName.Text = ""     'プロジェクト名
        frmMain.txtGnuPath.Text = ""    'gnuプロットパス
        frmMain.txtPjPath.Text = ""     'pjフォルダパス

        frmMain.txtX_d1.Text = ""        'X設定値
        frmMain.txtY_d1.Text = ""        'Y設定値
        frmMain.txtZ_d1.Text = ""        'Z設定値
        frmMain.txtX_d2.Text = ""        'X設定値
        frmMain.txtY_d2.Text = ""        'Y設定値
        frmMain.txtZ_d2.Text = ""        'Z設定値
        frmMain.txtX_d3.Text = ""        'X設定値
        frmMain.txtY_d3.Text = ""        'Y設定値
        frmMain.txtZ_d3.Text = ""        'Z設定値
        frmMain.txtX_d4.Text = ""        'X設定値
        frmMain.txtY_d4.Text = ""        'Y設定値
        frmMain.txtZ_d4.Text = ""        'Z設定値
        frmMain.txtX_d5.Text = ""        'X設定値
        frmMain.txtY_d5.Text = ""        'Y設定値
        frmMain.txtZ_d5.Text = ""        'Z設定値

        frmMain.txtCom1.Text = ""        'コマンド
        frmMain.txtCom2.Text = ""        'コマンド
        frmMain.txtCom3.Text = ""        'コマンド
        frmMain.txtCom4.Text = ""        'コマンド
        frmMain.txtCom5.Text = ""        'コマンド

        frmMain.txtStopFlg.Text = "0"   '周回終了フラグ=1:終了
        frmMain.txtNG.Text = "" 'NGカウント表示

        'タブのサイズを固定する
        frmMain.TabControl1.SizeMode = TabSizeMode.Fixed
        frmMain.TabControl1.ItemSize = New Size(90, 30)

        'TabControlをオーナードローする
        frmMain.TabControl1.DrawMode = TabDrawMode.OwnerDrawFixed
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' インターバルを設定する
    ''' </summary>
    ''' <param name="iInterval">インターバル値</param>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subiInterval(ByVal iInterval As Integer)
        On Error GoTo Error_Rtn


        System.Threading.Thread.Sleep(iInterval)
        'イベント発生
        Application.DoEvents()
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 画面ボタンを制御する
    ''' </summary>
    ''' <param name="bFlg">True:ON False:OFF</param>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history> 
    ''' -----------------------------------------------------------------------------
    Public Sub subBtnOnoff(ByVal bFlg As Boolean)
        On Error GoTo Error_Rtn

        frmMain.cmbSet_v.Enabled = bFlg     '周回
        frmMain.btnStart.Enabled = bFlg     'オート・スタート

        frmMain.pnlSet.Enabled = bFlg       '設定エリア
        frmMain.pnlPath.Enabled = bFlg      'Pathエリア

        frmMain.btnIniGet.Enabled = bFlg    'INIファイル読込
        frmMain.btnIniSave.Enabled = bFlg   'INIファイル保存
        frmMain.btnIniGet3.Enabled = bFlg   'INIファイル読込
        frmMain.btnIniSave3.Enabled = bFlg  'INIファイル保存
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 画面ログ「OK/NG」用配列を設定する
    ''' </summary>
    ''' <param name="iColor">0：緑色、1：赤色、2：黄色</param>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history> 
    ''' -----------------------------------------------------------------------------
    Public Sub subOkNg_Color(ByVal iColor As Integer)
        On Error GoTo Error_Rtn

        ReDim Preserve pOkNg(pOkNg_No)                              'OKorNG表示色変更
        pOkNg(pOkNg_No).iLen = frmMain.txtLogDisp.Text.Length - 3   'OKorNG表示色設定
        pOkNg(pOkNg_No).iColor = iColor
        pOkNg_No = pOkNg_No + 1
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 画面ログ「OK/NG」の表示色をを出力する
    ''' </summary>
    ''' <param name="iLen">該当文字目</param>
    ''' <param name="iColor">0：緑色、1：赤色、2：黄色</param>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>   
    ''' -----------------------------------------------------------------------------
    Public Sub subLogOutput_Color(ByVal iLen As Integer, ByVal iColor As Integer)
        On Error GoTo Error_Rtn

        'カレットを該当文字目に移動
        frmMain.txtLogDisp.Select(iLen, 0)
        '2文字選択する：OK or NG
        frmMain.txtLogDisp.SelectionLength = 2
        '色を変更する
        Select Case iColor
            Case 0
                frmMain.txtLogDisp.SelectionColor = Color.Lime
            Case 1
                frmMain.txtLogDisp.SelectionColor = Color.Red
            Case 2
                frmMain.txtLogDisp.SelectionColor = Color.Yellow
            Case Else
                frmMain.txtLogDisp.SelectionColor = Color.White
        End Select
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 画面ログを出力する
    ''' </summary>
    ''' <param name="sMessage">出力文字</param>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subLogOutput(ByVal sMessage As String)
        On Error GoTo Error_Rtn

        frmMain.txtLogDisp.Text = frmMain.txtLogDisp.Text + sMessage + Chr(13) + Chr(10)
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 画面設定情報を取得する
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subSetGamen()
        On Error GoTo Error_Rtn

        'ドローン設定台数
        pSet_d = frmMain.cmbSet_d.Text

        '周回の間隔(ミリ秒)
        pSet_Interval = frmMain.cmbSet_Interval.Text

        '周回
        pSet_v = frmMain.cmbSet_v.Text

        '周回ループ
        frmMain.txtStopFlg.Text = "0"

        'NGカウント表示
        frmMain.txtNG.Text = ""

        '画面ログ出力
        frmMain.txtLogDisp.Text = ""
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' テキストボックスのカーソルを制御する
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subSetRow()
        On Error GoTo Error_Rtn

        'カレット位置を末尾に移動
        frmMain.txtLogDisp.SelectionStart = frmMain.txtLogDisp.Text.Length

        'テキストボックスにフォーカスを移動
        frmMain.txtLogDisp.Focus()

        'カレット位置までスクロール
        frmMain.txtLogDisp.ScrollToCaret()
Error_Exit:
        Exit Sub
Error_Rtn:
        GoTo Error_Exit
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' メッセージボックスを表示する
    ''' </summary>
    ''' <param name="sName">表示文字</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncMsgBox(ByVal sName) As Boolean
        On Error GoTo Error_Rtn
        fncMsgBox = False

        Dim sMSG As String = sName
        MessageBox.Show(sMSG, pcAppName, MessageBoxButtons.OK, MessageBoxIcon.Information)

        fncMsgBox = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' エラーメッセージを表示する
    ''' </summary>
    ''' <param name="sName">表示文字</param>
    ''' <returns>True:OK False:NG</returns>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Function fncErrors(ByVal sName) As Boolean
        On Error GoTo Error_Rtn
        fncErrors = False

        Dim sMSG As String  'メッセージ文

        Select Case Err()
            Case Else
                sMSG = sName & Chr(13) & Err.Description
                MessageBox.Show(sMSG, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Select

        fncErrors = True
Error_Exit:
        Exit Function
Error_Rtn:
        GoTo Error_Exit
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' プロセスを終了する
    ''' </summary>
    ''' <param name="sName">終了したいプロセス名</param>
    ''' <remarks></remarks>
    ''' <author>RKA</author>
    ''' <history></history>
    ''' -----------------------------------------------------------------------------
    Public Sub subKillProc(ByVal sName As String)
        On Error Resume Next

        Dim ps As System.Diagnostics.Process() =
        System.Diagnostics.Process.GetProcessesByName(sName)

        For Each p As System.Diagnostics.Process In ps
            'プロセスを強制的に終了させる
            p.Kill()
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' CSVファイルの読込処理
    ''' </summary>
    ''' <param name="astrFileName">ファイル名</param>
    ''' <param name="ablnTab">区切りの指定(True:タブ区切り, False:カンマ区切り)</param>
    ''' <param name="ablnQuote">引用符フラグ(True:引用符で囲まれている, False:囲まれていない)</param>
    ''' <returns>読込結果の文字列の2次元配列</returns>
    ''' -----------------------------------------------------------------------------
    Public Function ReadCsv(ByVal astrFileName As String,
                         ByVal ablnTab As Boolean,
                         ByVal ablnQuote As Boolean) As String()()
        ReadCsv = Nothing
        'ファイルStreamReader
        Dim parser As Microsoft.VisualBasic.FileIO.TextFieldParser = Nothing
        Try
            'Shift-JISエンコードで変換できない場合は「?」文字の設定
            Dim encFallBack As System.Text.DecoderReplacementFallback = New System.Text.DecoderReplacementFallback("?")
            Dim enc As System.Text.Encoding =
            System.Text.Encoding.GetEncoding("shift_jis", System.Text.EncoderFallback.ReplacementFallback, encFallBack)

            'TextFieldParserクラス
            parser = New TextFieldParser(astrFileName, enc)

            '区切りの指定
            parser.TextFieldType = FieldType.Delimited
            If ablnTab = False Then
                'カンマ区切り
                parser.SetDelimiters(",")
            Else
                'タブ区切り
                parser.SetDelimiters(vbTab)
            End If

            If ablnQuote = True Then
                'フィールドが引用符で囲まれているか
                parser.HasFieldsEnclosedInQuotes = True
            End If

            'フィールドの空白トリム設定
            parser.TrimWhiteSpace = False

            Dim strArr()() As String = Nothing
            Dim nLine As Integer = 0
            'ファイルの終端までループ
            While Not parser.EndOfData
                'フィールドを読込
                Dim strDataArr As String() = parser.ReadFields()

                '戻り値領域の拡張
                ReDim Preserve strArr(nLine)

                '退避
                strArr(nLine) = strDataArr
                nLine += 1
            End While

            '正常終了
            Return strArr

        Catch ex As Exception
            'エラー
            MsgBox(ex.Message)
        Finally
            '閉じる
            If parser IsNot Nothing Then
                parser.Close()
            End If
        End Try
    End Function
End Module
