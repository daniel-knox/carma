Public Class Form2

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = FormatDateTime(DateString, DateFormat.LongDate) & "      " & Format(Now, "hh.mm.ss.fff tt")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form3.Show()
        Form3.TopMost = True
        Form3.Label1.Text = "Meter 1"
    End Sub
End Class