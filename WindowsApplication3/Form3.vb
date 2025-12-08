Public Class Form3
    Private Sub Form3_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Form1.StopFlag = True
        If Form1.BackgroundWorker1.IsBusy Then
            Form1.BackgroundWorker1.CancelAsync()
        End If
        Form1.Show()
    End Sub
    Private Sub Form3_Load(sender As System.Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = "Start..." & Environment.NewLine
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As EventArgs) Handles Button1.Click
        Label1.Text = "Stopping..."
        TextBox1.AppendText("Stopping requested..." & Environment.NewLine)
        Form1.BackgroundWorker1.CancelAsync()
        Form1.StopFlag = True
    End Sub

End Class