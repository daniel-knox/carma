Public Class Form4

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click




        TextBox1.Text = " " & Chr(162) & Chr(162) & Chr(162) & Chr(162) & Chr(162) & Chr(162) & Chr(162) & Chr(162)
        TextBox1.Font = New Font("Wingdings 2", 33.0, FontStyle.Regular)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Form4_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        TextBox1.Text = vbNewLine & "            " & "KWH" & vbNewLine & "0"

    End Sub
    Public Sub Form4_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

       e.Cancel = True


    End Sub





End Class