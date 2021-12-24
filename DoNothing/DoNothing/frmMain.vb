Public Class frmMain

    Private Sub frmMain_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        'イベント発生
        Application.DoEvents()

        System.Threading.Thread.Sleep(3000)

        Me.Close()
    End Sub
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If pFileName <> "" Then
            txtMsg.Text = pFileName
        End If
    End Sub
End Class