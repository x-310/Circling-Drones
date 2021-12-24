Module Module1
    Public pFileName As String = ""

    Public Class Program
        Public Shared Sub Main()
            Try
                ChkParam()

                Dim NewForm1 As New frmMain
                NewForm1.ShowDialog()

            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show( _
                   "エラーが発生しました" & vbCrLf & vbCrLf & _
                        ex.Message, _
                    "IniCntUp", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Sub

        Public Shared Sub ChkParam()
            'コマンドライン引数を取得する
            Dim args As String() = System.Environment.GetCommandLineArgs()

            If args.Length > 1 AndAlso args(1) <> "" Then
                pFileName = args(1)
            End If
        End Sub
    End Class
End Module
