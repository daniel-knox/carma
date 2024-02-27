Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.IO
Imports System.Drawing.Printing
Imports System.Windows.Forms


Public Class Form8
    Inherits Form
    Private printButton As Button
    Private printDocument2 As New PrintDocument()
    Private stringToPrint As String






    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim date1 As String
        Dim Bench As String
        Dim FileToDelete As String
        TextBox1.Text = ""
        FileToDelete = "C:\temp\hold.txt"

        ListBox2.SelectedIndex = -1
        ListBox3.SelectedIndex = -1
        ListBox4.SelectedIndex = -1




        If System.IO.File.Exists(FileToDelete) = True Then

            System.IO.File.Delete(FileToDelete)

        End If



        If ListBox1.SelectedItems.Count < 1 Then
            MsgBox("Please Select BENCH")
            Exit Sub
        Else
            Bench = ListBox1.SelectedItem.ToString()
            date1 = DateTimePicker1.Value.ToString("yyyy-MM-dd 00:00:00 ")
        End If


        Call SqlBlob2File("C:\temp\hold.txt", date1, Bench)

        If System.IO.File.Exists("C:\temp\hold.txt") = True Then
            For Each s As String In System.IO.File.ReadAllLines("C:\temp\hold.txt")
                TextBox1.AppendText(s + vbNewLine)
            Next
        Else
            TextBox1.AppendText("No Data Avaiable for those selections" + vbNewLine)
        End If


    End Sub


    Private Sub SqlBlob2File(ByVal DestFilePath As String, ByVal date1 As String, ByVal Bench As String)
        Dim Content As Integer = 0 ' the column # of the BLOB field
        Dim cn As New SqlConnection("server=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
        Dim cmd As New SqlCommand("Select Top 1 [MCTEST].[dbo].[TestTable].[Content] FROM  [MCTEST].[dbo].TestTable  where date > " & "'" & date1 & "'" & "and [MCTEST].[dbo].[TestTable].[FileName] like '%" & Bench & "%' order by ID DESC", cn)
        cn.Open()
        Dim dr As SqlDataReader = cmd.ExecuteReader()
        dr.Read()



        If dr.HasRows Then

            Dim b(dr.GetBytes(Content, 0, Nothing, 0, Integer.MaxValue) - 1) As Byte
            dr.GetBytes(Content, 0, b, 0, b.Length)
            dr.Close()
            cn.Close()
            Dim fs As New System.IO.FileStream(DestFilePath, IO.FileMode.Create, IO.FileAccess.Write)
            fs.Write(b, 0, b.Length)
            fs.Close()
        Else
        End If
        End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PrintFile("C:\temp\hold.txt")

    End Sub
    Sub PrintFile(ByVal fileName As String)
        Dim myFile As New ProcessStartInfo
        With myFile
            .UseShellExecute = True
            .WindowStyle = ProcessWindowStyle.Hidden
            .FileName = fileName
            .Verb = "Print"
        End With
        Process.Start(myFile)
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim Voltage As String
        Dim Vtap As String
        Dim Bench As String
        Dim FileToDelete As String
        TextBox1.Text = ""
        FileToDelete = "C:\temp\hold.txt"
        ListBox1.SelectedIndex = -1



        If System.IO.File.Exists(FileToDelete) = True Then

            System.IO.File.Delete(FileToDelete)

        End If

        If ListBox3.SelectedItems.Count < 1 Then
            MsgBox("Please Select Voltage")
            Exit Sub
        Else
            Voltage = ListBox3.SelectedItem.ToString()
        End If

        If ListBox2.SelectedItems.Count < 1 Then
            MsgBox("Please Select Tap")
            Exit Sub
        Else
            Vtap = ListBox2.SelectedItem.ToString()
        End If

        If ListBox4.SelectedItems.Count < 1 Then
            MsgBox("Please Select Bench")
            Exit Sub
        Else
            Bench = ListBox4.SelectedItem.ToString()
        End If

        If System.IO.File.Exists(FileToDelete) = True Then
            System.IO.File.Delete("C:\temp\hold.txt")
        End If
        Call SqlBlob2File2("C:\temp\hold.txt", Voltage, Vtap, Bench)


        If System.IO.File.Exists("C:\temp\hold.txt") = True Then

            For Each s As String In System.IO.File.ReadAllLines("C:\temp\hold.txt")
                TextBox1.AppendText(s + vbNewLine)
            Next
        Else
            TextBox1.AppendText("No Data Avaiable for those selections" + vbNewLine)

        End If


    End Sub
    Private Sub SqlBlob2File2(ByVal DestFilePath As String, ByVal Voltage As String, ByVal Vtap As String, ByVal Bench As String)

        Dim year As String

        year = ListBox6.SelectedItem



        Dim Content As Integer = 0 ' the column # of the BLOB field
        Dim cn As New SqlConnection("server=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
        Dim cmd As New SqlCommand("Select Top 1 [MCTEST].[dbo].[TestTable].[Content] FROM  [MCTEST].[dbo].TestTable  join [MCTEST].[dbo].[TestResults] on [MCTEST].[dbo].[TestTable].[ID] = [MCTEST].[dbo].[TestResults].[UniqID_ text]  where[MCTEST].[dbo].[TestResults].date like" & "'%" & year & "%'" & " and [MCTEST].[dbo].[TestResults].VOLTAGE = " & "'" & Voltage & "' and [MCTEST].[dbo].[TestResults].VTap = " & "'" & Vtap & "' and  [MCTEST].[dbo].[TestTable].[FileName] like " & "'%" & Bench & "%'" & " Order by [MCTEST].[dbo].[TestTable].[date] desc", cn)
        cn.Open()
        Dim dr As SqlDataReader = cmd.ExecuteReader()
        dr.Read()

        If dr.HasRows Then
            Dim b(dr.GetBytes(Content, 0, Nothing, 0, Integer.MaxValue) - 1) As Byte
            dr.GetBytes(Content, 0, b, 0, b.Length)
            dr.Close()
            cn.Close()
            Dim fs As New System.IO.FileStream(DestFilePath, IO.FileMode.Create, IO.FileAccess.Write)
            fs.Write(b, 0, b.Length)
            fs.Close()

            Dim cmd1 As New SqlCommand("Select Top 1 [MCTEST].[dbo].[TestTable].[date] FROM  [MCTEST].[dbo].TestTable  join [MCTEST].[dbo].[TestResults] on [MCTEST].[dbo].[TestTable].[ID] = [MCTEST].[dbo].[TestResults].[UniqID_ text]  where [MCTEST].[dbo].[TestResults].VOLTAGE = " & "'" & Voltage & "' and [MCTEST].[dbo].[TestResults].VTap = " & "'" & Vtap & "' and  [MCTEST].[dbo].[TestTable].[FileName] like " & "'%" & Bench & "%'" & " Order by [MCTEST].[dbo].[TestTable].[date] desc", cn)
            cn.Open()
            Dim dr1 As SqlDataReader = cmd1.ExecuteReader()
            dr1.Read()

            If dr1.HasRows Then

                Dim timeFormat As String = "yyyy-MM-dd HH:mm:ss"
                Dim mydate As DateTime
                mydate = dr1.GetSqlDateTime(0)
                TextBox2.Text = mydate.ToString(timeFormat)
                dr1.Close()
                cn.Close()


            End If



        Else



        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim rws As Integer = 0
        Dim year As String
        Dim SBench As String = ListBox7.SelectedItem
        SBench = Mid(SBench, 6, 1)
        year = ListBox5.SelectedItem


            TextBox1.Text = "" & vbNewLine
        Dim cn3 As New SqlConnection("server=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
        'Dim cmd3 As New SqlCommand("SELECT distinct [Voltage],[corrected_vtap]FROM [MCTEST].[dbo].[INIList]where( ([MCTEST].[dbo].[INIList].[VTap] is not null)and ([MCTEST].[dbo].[INIList].[BenchNum] = " & "'" & SBench & "'" & ")  and NOT EXISTS (SELECT [UniqID],[UniqID_ text],[Units],[VOLTAGE],[Load],[POWERFACTOR],[VTap],[percent_error],[WHexternal],[WHinternal],[operator],[Date] FROM [MCTEST].[dbo].[TestResults] where [MCTEST].[dbo].[TestResults].date like" & "'%" & year & "%'" & " and  ([MCTEST].[dbo].[INIList].[Type] = [MCTEST].[dbo].[TestResults].[units]) and ([MCTEST].[dbo].[INIList].[Corrected_Vtap] = [MCTEST].[dbo].[TestResults].[VTap]) and ([MCTEST].[dbo].[INIList].[Voltage] = [MCTEST].[dbo].[TestResults].[Voltage])and([MCTEST].[dbo].[INIList].[Load] = [MCTEST].[dbo].[TestResults].[Load]) )) order by Corrected_Vtap", cn3)
        Dim cmd3 As New SqlCommand("SELECT distinct [Voltage],[corrected_vtap]FROM [MCTEST].[dbo].[INIList]where ([MCTEST].[dbo].[INIList].[VTap] is not null)and ([MCTEST].[dbo].[INIList].[BenchNum] = " & "'" & SBench & "'" & ")and not EXISTS (SELECT[ID],[TestTable].[date],[FileName],VOLTAGE,TestResults.VTap FROM [MCTEST].[dbo].[TestTable] Join TestResults on TestTable.ID = TestResults.[UniqID_ text] where [MCTEST].[dbo].[TestResults].date like'%2016%' and TestTable.FileName like '%Bench" & SBench & "%' and ([MCTEST].[dbo].[INIList].[Voltage] = [MCTEST].[dbo].[Testresults].VOLTAGE)and ([MCTEST].[dbo].[INIList].Corrected_Vtap = [MCTEST].[dbo].[Testresults].VTap))", cn3)

        cn3.Close()
            cn3.Open()

            Dim dr3 As SqlDataReader = cmd3.ExecuteReader()
            While dr3.Read()
            TextBox3.Text = dr3.GetString(1)
                If dr3.HasRows Then
                    TextBox1.Text = "   " & TextBox1.Text & dr3.GetString(1) & "   " & dr3.GetValue(0).ToString & vbNewLine
                    rws = rws + 1
                End If

            End While

            TextBox3.Text = " " & (rws * 4).ToString & " points yet to be tested this year "
















    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click


        Call FTestApp.CompleteaccuracyMA()

















    End Sub

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

 
End Class