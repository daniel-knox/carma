Imports System.ComponentModel
Imports System.IO.Ports
Imports System
Imports System.Threading
Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports AdvDIOLib
Imports AxAdvDIOLib
Imports AxDSOFramer
Imports SourceGrid2





Public Class Form1
    Public Event DoWork As DoWorkEventHandler
    Public SerialPort2 As SerialPort
    Public Property x1Wbook As Object
    Public Property x1App As Object



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox22.Text = "LLS – Light load series unity " + Environment.NewLine + Environment.NewLine + "FLS – Full load series unity" + Environment.NewLine + Environment.NewLine + "FLOC – Full load outer coil unity " _
             + Environment.NewLine + Environment.NewLine + "FLOP – Full load outer coil PF" + Environment.NewLine + Environment.NewLine + "FLMC – Full load middle coil unity" + Environment.NewLine + Environment.NewLine + _
            "FLIC – Full load inner coil unity" + Environment.NewLine + Environment.NewLine + "FLIP – Full load inner coil PF" + Environment.NewLine + Environment.NewLine + "kW Demand" _
        + Environment.NewLine + Environment.NewLine + "kVA Demand" + Environment.NewLine + Environment.NewLine + "Creep" + Environment.NewLine + Environment.NewLine + "Data retention test" + Environment.NewLine + Environment.NewLine + "Detent test" _
         + Environment.NewLine + Environment.NewLine + "Set Field Configuration"
        Button2.BackColor = Color.Lime
        Button1.BackColor = Color.Red

        '        a.	LLS – Light load series unity – 2 min
        'b.	FLS – Full load series unity – 2 min
        'c.	FLOC – Full load outer coil unity – 2 min
        'd.	FLOP – Full load outer coil PF – 2 min
        'e.	FLMC – Full load middle coil unity – 2 min
        'f.	FLMP – Full load middle coil PF – 2 min
        'g.	FLIC – Full load inner coil unity – 2 min
        'h.	FLIP – Full load inner coil PF – 2 min
        'i.	kW Demand – 15 min
        'j.	kVA Demand – 5 min
        'k.	Creep – 1 min
        'l.	Data retention test – 1 min
        'm.	TOU set
        '        n.TOU verified
        'o.	Time Set
        '        p.Time verified
        'q.	Clear kWh registers
        'r.	kWh registers verified cleared


    End Sub

    Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox22.Text = ""
        Button2.BackColor = SystemColors.GradientInactiveCaption
        Button1.BackColor = SystemColors.GradientInactiveCaption


    End Sub

    Private Sub NewEMPConfigToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewEMPConfigToolStripMenuItem.Click

        System.Diagnostics.Process.Start("C:\")

    End Sub

    Private Sub NonEMPConfigToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NonEMPConfigToolStripMenuItem.Click
        System.Diagnostics.Process.Start("C:\")
    End Sub

    Private Sub AccuracyCheckToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AccuracyCheckToolStripMenuItem.Click
        'System.Diagnostics.Process.Start("C:\")
    End Sub

    Private Sub EndTestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EndTestToolStripMenuItem.Click
        Button1_Click(Me, EventArgs.Empty)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub GroupBox24_Enter(sender As Object, e As EventArgs) Handles GroupBox24.Enter
        Form2.Show()
        Form2.TopMost = True
        Form2.Label1.Text = "Meter 1"

    End Sub

    Private Sub GroupBox25_Enter(sender As Object, e As EventArgs) Handles GroupBox25.Enter
        Form2.Show()
        Form2.TopMost = True
        Form2.Label1.Text = "Meter 2"
    End Sub

    Private Sub GroupBox26_Enter(sender As Object, e As EventArgs) Handles GroupBox26.Enter
        Form2.Show()
        Form2.TopMost = True
        Form2.Label1.Text = "Meter 3"
    End Sub

    Private Sub GroupBox27_Enter(sender As Object, e As EventArgs) Handles GroupBox27.Enter
        Form2.Show()
        Form2.TopMost = True
        Form2.Label1.Text = "Meter 4"
    End Sub

    Private Sub GroupBox28_Enter(sender As Object, e As EventArgs) Handles GroupBox28.Enter
        Form2.Show()
        Form2.TopMost = True
        Form2.Label1.Text = "Meter 5"
    End Sub

    Private Sub GroupBox29_Enter(sender As Object, e As EventArgs) Handles GroupBox29.Enter
        Form2.Show()
        Form2.TopMost = True
        Form2.Label1.Text = "Meter 6"
    End Sub

    Private Sub GroupBox30_Enter(sender As Object, e As EventArgs) Handles GroupBox30.Enter
        Form2.Show()
        Form2.TopMost = True
        Form2.Label1.Text = "Meter 7"
    End Sub

    Private Sub GroupBox31_Enter(sender As Object, e As EventArgs) Handles GroupBox31.Enter
        Form2.Show()
        Form2.TopMost = True
        Form2.Label1.Text = "Meter 8"
    End Sub

    Private Sub GroupBox32_Enter(sender As Object, e As EventArgs) Handles GroupBox32.Enter
        Form2.Show()
        Form2.TopMost = True
        Form2.Label1.Text = "Meter 9"
    End Sub

    Private Sub GroupBox33_Enter(sender As Object, e As EventArgs) Handles GroupBox33.Enter
        Form2.Show()
        Form2.TopMost = True
        Form2.Label1.Text = "Meter 10"
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click

        If Button7.BackColor = Color.Lime Then
            Button7.BackColor = SystemColors.GradientInactiveCaption
            Call Button3_Click(0, e)
            Call Button4_Click(0, e)

        Else

            Button10.BackColor = SystemColors.GradientInactiveCaption
            Button9.BackColor = SystemColors.GradientInactiveCaption
            Button8.BackColor = SystemColors.GradientInactiveCaption
            Button6.BackColor = SystemColors.GradientInactiveCaption
            Button7.BackColor = Color.Lime
            Call Button3_Click(0, e)
            Call Button4_Click(0, e)

        End If


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Button10.BackColor = SystemColors.GradientInactiveCaption
        Button9.BackColor = SystemColors.GradientInactiveCaption
        Button8.BackColor = SystemColors.GradientInactiveCaption
        Button7.BackColor = SystemColors.GradientInactiveCaption
        Button6.BackColor = Color.Lime
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Button10.BackColor = SystemColors.GradientInactiveCaption
        Button9.BackColor = SystemColors.GradientInactiveCaption
        Button6.BackColor = SystemColors.GradientInactiveCaption
        Button7.BackColor = SystemColors.GradientInactiveCaption
        Button8.BackColor = Color.Lime
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Button9.BackColor = Color.Lime
        Button10.BackColor = SystemColors.GradientInactiveCaption
        Button8.BackColor = SystemColors.GradientInactiveCaption
        Button6.BackColor = SystemColors.GradientInactiveCaption
        Button7.BackColor = SystemColors.GradientInactiveCaption
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Button10.BackColor = Color.Lime
        Button9.BackColor = SystemColors.GradientInactiveCaption
        Button8.BackColor = SystemColors.GradientInactiveCaption
        Button6.BackColor = SystemColors.GradientInactiveCaption
        Button7.BackColor = SystemColors.GradientInactiveCaption



    End Sub
    Friend Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim pList() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("C:\Program Files (x86)\Radian\RR-Kit\RRKit_VBNETSample.exe")
        Dim resp As MsgBoxResult

        pList = System.Diagnostics.Process.GetProcessesByName("RRKit_VBNETSample")

        If Button3.BackColor <> Color.Lime Then
            Button3.BackColor = Color.Lime
            Process.Start("C:\Program Files (x86)\Radian\RR-Kit\RRKit_VBNETSample.exe")
        ElseIf Button3.BackColor = Color.Lime And pList.Rank <> 0 Then
            Button3.BackColor = SystemColors.GradientInactiveCaption
            For Each proc As System.Diagnostics.Process In pList
                resp = 1
                If resp = 1 Then
                    proc.Kill()
                    Button3.BackColor = SystemColors.GradientInactiveCaption
                End If
            Next

        End If





        If Button3.BackColor = Color.Lime Then
            Button4.BackColor = SystemColors.GradientInactiveCaption

            For Each proc As System.Diagnostics.Process In pList
                resp = 1
                If resp = 1 Then
                    proc.Kill()
                    Button3.BackColor = SystemColors.GradientInactiveCaption
                End If
            Next





        End If










100:














    End Sub

    Friend Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim pList() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("C:\Program Files (x86)\Radian\RR-Kit\RRKit_VBNETSample.exe")
        Dim resp As MsgBoxResult




        On Error Resume Next
        If Button4.BackColor <> Color.Lime Then
            Form4.Show()
            Form4.Visible = True
            Form4.Button1.Focus()
            Button4.BackColor = Color.Lime
            GoTo 100
            If Button4.BackColor <> Color.Lime Then
                Form4.Visible = False
                Button4.BackColor = Color.Lime
                GoTo 100
            End If

        End If

        If Button4.BackColor = Color.Lime And Form4.Visible = False Then
            Button4.BackColor = SystemColors.GradientInactiveCaption
            GoTo 100


            If Button4.BackColor <> Color.Lime Then
                Form4.Visible = False
                Button4.BackColor = Color.Lime
                GoTo 100
            End If

        End If







        If Button3.BackColor = Color.Lime Then
            pList = System.Diagnostics.Process.GetProcessesByName("RRKit_VBNETSample")
            For Each proc As System.Diagnostics.Process In pList
                resp = 1
                If resp = 1 Then
                    proc.Kill()
                    Button3.BackColor = SystemColors.GradientInactiveCaption





                End If
            Next

        End If






200:    On Error Resume Next
        If Form4.Visible = False And Button4.BackColor = Color.Lime Then

            Button4.BackColor = SystemColors.GradientInactiveCaption
            Form4.TextBox1.Text = ""
            GoTo 100
        End If

        If Form4.Visible = True And Button4.BackColor = Color.Lime Then
            Form4.Visible = False
            Button4.BackColor = SystemColors.GradientInactiveCaption
            Form4.TextBox1.Text = ""

        End If







100:
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        BackgroundWorker1.RunWorkerAsync()
        BackgroundWorker1.WorkerReportsProgress = True
        BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        BackgroundWorker2.WorkerReportsProgress = True
        BackgroundWorker2.WorkerSupportsCancellation = True
      
     
    End Sub

    Public Sub Form1_CallScans(e As Integer)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        BackgroundWorker1.WorkerSupportsCancellation = True
        BackgroundWorker1.RunWorkerAsync(e)

    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim returnValue As String
        Dim EMPPort As Integer
        Dim Portnames As String() = System.IO.Ports.SerialPort.GetPortNames
        Dim frmCollection As New FormCollection()
        Dim buffer(3) As Byte
        Dim byteEnd(2) As Char
        Dim send As String
        Dim PT1A As String
        Dim PT1B As String
        Dim PT1C As String
        Dim x As Long
        Dim z As String
        Dim y As String
        Dim a As String
        Dim hexrep1A As String
        Dim hexrep2A As String
        Dim u As String
        Dim v As String
        Dim zz As String
        Dim s As String
        Dim t As String
        Dim zzz As String
        Dim q As String
        Dim r As String
        Dim zzzz As String
        Dim CT1A As String
        Dim k As String
        Dim l As String
        Dim m As String
        Dim n As String
        Dim o As String
        Dim zzzzz As String
        Dim Degreesint As Int32
        Dim Degrees As Double
        Dim num As Double
        Dim watts120 As Double
        Dim PF As Decimal
        Dim pi As Double
        Dim zzzzzz As String
        Dim watts360 As Double
        Dim zzzzzzz As String
        Dim watts600 As Double
        Dim va120 As Double
        Dim va360 As Double






        BackgroundWorker1.WorkerSupportsCancellation = True

            zzzz = ""
            SerialPort3 = New SerialPort
            If Portnames Is Nothing Then
                MsgBox("There are no Com Ports detected!")
                Me.Close()
            End If
            With SerialPort3
                Try
                    EMPPort = Integer.Parse(TextBox_EMPPort.Text)
                Catch eg As Exception
                    EMPPort = 3
                End Try
                .PortName = "COM" + EMPPort.ToString()
                .ParityReplace = &H3B                    ' replace ";" when parity error occurs 
                .BaudRate = "9600"
                .Parity = IO.Ports.Parity.None
                .DataBits = 8
                .StopBits = IO.Ports.StopBits.One
                .Handshake = IO.Ports.Handshake.None
                .RtsEnable = True
                .ReceivedBytesThreshold = 3             'threshold: one byte in buffer > event is fired 
                ' CR must be the last char in frame. This terminates the SerialPort.readLine 
                .ReadTimeout = 1000
                .WriteTimeout = 1000
                .NewLine = Chr(13)
            End With

            Try
            SerialPort3.Open()
            Catch
            MessageBox.Show("Cannot open COM so cannot pulse the pulse output")
            End Try

        If (e.Argument = 5) Then
            If SerialPort3.IsOpen Then

                Try
                    'send = "XF0HS0100002884"
                    'SerialPort3.WriteLine(send)
                    send = ""
                    send = "XF0o"  'Pulse the output pin (I.e. start test)
                    SerialPort3.WriteLine(send)

                Catch ex As Exception
                    MessageBox.Show("Cannot pulse the pulse output to start - the WriteLine command failed!")
                    SerialPort3.Close()
                End Try

            End If
        End If

        'Run NSCANS command for the number of scans indicated by e.Argument




        frmCollection = Application.OpenForms()
        Do While frmCollection.Item("Form1").IsHandleCreated And BackgroundWorker1.CancellationPending = False

            If (Me.BackgroundWorker1.CancellationPending = True) Then
                SerialPort3.Close()
                e.Cancel = True
                Exit Do
            End If

            Try
                If (SerialPort3.IsOpen = False) Then
                SerialPort3.Open()
                End If

                If SerialPort3.IsOpen Then

                    Try
                        send = "XF0HS0100002884"
                        SerialPort3.WriteLine(send)
                        Threading.Thread.Sleep(400)

                    Catch ex As Exception
                        MsgBox("Read " & ex.Message)
                    End Try

                End If

                If SerialPort3.IsOpen Then

                    Try
                        returnValue = SerialPort3.ReadLine()


                    Catch ex As Exception
                        MsgBox("Error #2")
                    End Try

                End If
            Catch ex As Exception

                ' MessageBox.Show(ex.Message)
            Finally

                If (SerialPort3.IsOpen = False) Then
                    SerialPort3.Open()
                End If

                If SerialPort3.IsOpen Then

                    Try
                        send = "NF070"
                        SerialPort3.WriteLine(send)
                        Threading.Thread.Sleep(400)
                        returnValue = SerialPort3.ReadLine()

                        x = Mid(returnValue, 6, 2)
                        y = Mid(returnValue, 8, 2)
                        a = Mid(returnValue, 10, 2)
                        hexrep1A = Chr(CInt("&H" & x))
                        hexrep2A = Chr(CInt("&H" & y))
                        z = (a & hexrep1A & hexrep2A)
                        PT1A = (Convert.ToInt64(z, 16) / 100)
                        If PT1A = "0" Then
                            PT1A = "0.00"
                            zzzz = ""



                            If (SerialPort3.IsOpen = False) Then
                                SerialPort3.Open()
                            End If

                            If SerialPort3.IsOpen Then

                                If (e.Argument = 5) Then
                                    Try
                                        send = "XF0o"  'Pulse the output pin (end test)
                                        SerialPort3.WriteLine(send)
                                        Threading.Thread.Sleep(200)
                                        send = ""
                                        send = "XF0HS0100002884"
                                        SerialPort3.WriteLine(send)


                                        Threading.Thread.Sleep(500)

                                    Catch ex As Exception
                                        MsgBox("Error #3")
                                        SerialPort3.Close()
                                    End Try
                                End If
                            End If

                        End If
                        If Len(PT1A) < 6 Then
                            PT1A = PT1A & "0"
                        End If





                        s = Mid(returnValue, 14, 2)
                        t = Mid(returnValue, 12, 2)
                        zzz = (s & t)
                        PT1B = (Convert.ToInt64(zzz, 16) / 100)
                        If PT1B = "0" Then
                            PT1B = "0.00"
                        End If
                        If Len(PT1B) < 6 Then
                            PT1B = PT1B & "0"
                        End If






                        u = Mid(returnValue, 18, 2)
                        v = Mid(returnValue, 16, 2)
                        zz = (u & v)
                        PT1C = (Convert.ToInt64(zz, 16) / 100)
                        If PT1C = "0" Then
                            PT1C = "0.00"
                        End If
                        If Len(PT1C) < 6 Then
                            PT1C = PT1C & "0"
                        End If

                        q = Mid(returnValue, 146, 2)
                        r = Mid(returnValue, 144, 2)
                        zzzz = (q & r)
                        CT1A = (Convert.ToInt64(zzzz, 16) / 100).ToString("####0.00")
                        If CT1A = "0" Then
                            CT1A = "0.00"
                        End If








                        If (Me.IsHandleCreated = True) Then
                            Me.Invoke(New MethodInvoker(Sub()
                                                            Me.TextBox10.Text = PT1A
                                                            Me.TextBox5.BackColor = Color.LimeGreen
                                                            Me.Label6.Text = "Updating"
                                                            Me.Update()
                                                            Me.TextBox9.Text = PT1B
                                                            Me.TextBox8.Text = PT1C
                                                            Me.TextBox6.Text = CT1A
                                                            If Me.TextBox10.Text = "0.000" Then
                                                                Me.TextBox6.Text = "0.00"
                                                            End If

                                                        End Sub))

                        End If




                    Catch
                    End Try

                End If

            End Try

            If (Me.BackgroundWorker1.CancellationPending = True) Then
                e.Cancel = True
                SerialPort3.Close()
                Exit Do
            End If

            If (SerialPort3.IsOpen = False) Then
                SerialPort3.Open()
            End If


            If SerialPort3.IsOpen Then

                Try
                    Thread.Sleep(500)
                    send = "XF0HE0100BA"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(500)
                    returnValue = SerialPort3.ReadLine()
                    send = "NF00D"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(1000)
                    returnValue = SerialPort3.ReadLine()
                    n = Mid(returnValue, 24, 2)
                    o = Mid(returnValue, 22, 2)
                    l = Mid(returnValue, 28, 2)
                    m = Mid(returnValue, 26, 2)
                    zzzzz = (n & o)
                    Degreesint = Int32.Parse(zzzzz, System.Globalization.NumberStyles.HexNumber)
                    Degrees = Math.Round(Degreesint / 100, 4)
                    pi = 3.14159265
                    num = Degrees * pi / 180
                    PF = (Math.Cos(num))
                    zzzzzz = (l & m)
                    watts120 = (Convert.ToInt64(zzzzzz, 16) / 10)

                    If (Me.IsHandleCreated = True) Then
                        Me.Invoke(New MethodInvoker(Sub()
                                                        If Me.RadioButton1.Checked = True Then
                                                            Me.TextBox7.Text = Degrees.ToString("####0.00")
                                                        ElseIf Me.RadioButton2.Checked = True Then
                                                            Me.TextBox7.Text = PF.ToString("####0.000")
                                                        End If
                                                        If Me.TextBox10.Text = "0.000" Then
                                                            Me.TextBox7.Text = "0.000"
                                                        End If
                                                        Me.TextBox11.Text = watts120.ToString("####0.00")
                                                    End Sub))
                    End If
                Catch
                End Try
            End If

            If (Me.BackgroundWorker1.CancellationPending = True) Then
                e.Cancel = True
                SerialPort3.Close()
                Exit Do
            End If

            If SerialPort3.IsOpen Then

                Try
                    send = "XF0HE0101B9"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(400)
                    returnValue = SerialPort3.ReadLine()
                    send = "NF00D"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(200)
                    returnValue = SerialPort3.ReadLine()
                    k = Mid(returnValue, 30, 2)
                    l = Mid(returnValue, 28, 2)
                    m = Mid(returnValue, 26, 2)
                    zzzzzz = (k & l & m)
                    watts360 = (Convert.ToInt64(zzzzzz, 16) / 10)
                    If (Me.IsHandleCreated = True) Then
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox12.Text = watts360.ToString("####0.00")

                                                    End Sub))
                    End If
                Catch
                End Try
            End If

            If (Me.BackgroundWorker1.CancellationPending = True) Then
                e.Cancel = True
                SerialPort3.Close()
                Exit Do
            End If

            If SerialPort3.IsOpen Then

                Try
                    send = "XF0HE0102B8"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(200)
                    returnValue = SerialPort3.ReadLine()
                    send = "NF00D"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(200)
                    returnValue = SerialPort3.ReadLine()
                    k = Mid(returnValue, 30, 2)
                    l = Mid(returnValue, 28, 2)
                    m = Mid(returnValue, 26, 2)
                    zzzzzzz = (k & l & m)
                    watts600 = (Convert.ToInt64(zzzzzzz, 16) / 10)
                    If (Me.IsHandleCreated = True) Then
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox13.Text = watts600.ToString("####0.00")

                                                    End Sub))
                    End If
                Catch
                End Try
            End If

            If (Me.BackgroundWorker1.CancellationPending = True) Then
                e.Cancel = True
                Exit Do
            End If


            If SerialPort3.IsOpen Then

                Try
                    send = "XF0HP0400AC"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(200)
                    returnValue = SerialPort3.ReadLine()
                    send = "NF010"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(200)
                    returnValue = SerialPort3.ReadLine()
                    va120 = (Convert.ToInt64((Mid(returnValue, 26, 2) & (Mid(returnValue, 24, 2))), 16) / 10)


                    If (Me.IsHandleCreated = True) Then
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox14.Text = va120.ToString("####0.00")

                                                    End Sub))

                    End If


                Catch
                End Try
            End If

            If (Me.BackgroundWorker1.CancellationPending = True) Then
                e.Cancel = True
                Exit Do
            End If

            If SerialPort3.IsOpen Then

                Try
                    send = "XF0HP0401AB"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(200)
                    returnValue = SerialPort3.ReadLine()
                    send = "NF010"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(200)
                    returnValue = SerialPort3.ReadLine()
                    va360 = (Convert.ToInt64((Mid(returnValue, 28, 2) & (Mid(returnValue, 26, 2) & (Mid(returnValue, 24, 2)))), 16) / 10)


                    If (Me.IsHandleCreated = True) Then
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox15.Text = va360.ToString("####0.00")

                                                    End Sub))

                    End If


                Catch
                End Try
            End If

            If (Me.BackgroundWorker1.CancellationPending = True) Then
                e.Cancel = True
                SerialPort3.Close()
                Exit Do
            End If

            If SerialPort3.IsOpen Then

                Try
                    send = "XF0HP0402AA"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(200)
                    returnValue = SerialPort3.ReadLine()
                    send = "NF010"
                    SerialPort3.WriteLine(send)
                    Threading.Thread.Sleep(200)
                    returnValue = SerialPort3.ReadLine()
                    va360 = (Convert.ToInt64((Mid(returnValue, 28, 2) & (Mid(returnValue, 26, 2) & (Mid(returnValue, 24, 2)))), 16) / 10)


                    If (Me.IsHandleCreated = True) Then
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox16.Text = va360.ToString("####0.00")

                                                    End Sub))

                    End If


                Catch
                End Try
            End If


            If (Me.BackgroundWorker1.CancellationPending = True) Then
                e.Cancel = True
                Exit Do
            End If



            If (Me.IsHandleCreated = True) Then
                Me.Invoke(New MethodInvoker(Sub()
                                                Me.TextBox5.BackColor = Color.White
                                                Me.Label6.Text = "Idle"
                                            End Sub))
            End If

            'SerialPort3.Close()
            returnValue = ""


        Loop
        e.Cancel = True
        SerialPort3.Close()
    End Sub
   
    Private Sub MCMKAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MCMKAToolStripMenuItem.Click
        Dim MC_Cert_dialog_result As DialogResult = MessageBox.Show("This is the Measurement Canada Certification process. Are you sure you want to run this process?", _
                                                                    "MC Accuracy Check", _
                                                                    MessageBoxButtons.OKCancel)
        If (MC_Cert_dialog_result = DialogResult.OK) Then
            FTestApp.Visible = True  ' Form5.Visible = True
        Else
            'other
        End If

    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Form6.Show()

    End Sub

    Public Sub MAAccuracy(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Call FTestApp.accuracyMA()

    End Sub

    Private Sub MAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MAToolStripMenuItem.Click
        Form8.Show()
        Form8.ListBox1.Items.Add("Bench1")
        Form8.ListBox1.Items.Add("Bench2")
        Form8.ListBox1.Items.Add("Bench3")

        Form8.ListBox2.Items.Add("M02")
        Form8.ListBox2.Items.Add("M03")
        Form8.ListBox2.Items.Add("M04")
        Form8.ListBox2.Items.Add("M05")
        Form8.ListBox2.Items.Add("M06")
        Form8.ListBox2.Items.Add("M07")
        Form8.ListBox2.Items.Add("M08")
        Form8.ListBox2.Items.Add("M09")
        Form8.ListBox2.Items.Add("M10")
        Form8.ListBox2.Items.Add("M11")
        Form8.ListBox2.Items.Add("M12")
        Form8.ListBox2.Items.Add("M13")
        Form8.ListBox2.Items.Add("M14")
        Form8.ListBox2.Items.Add("M15")
        Form8.ListBox2.Items.Add("M16")
        Form8.ListBox2.Items.Add("M17")
        Form8.ListBox2.Items.Add("M18")
        Form8.ListBox2.Items.Add("M19")
        Form8.ListBox2.Items.Add("M20")
        Form8.ListBox2.Items.Add("M21")
        Form8.ListBox2.Items.Add("M22")
        Form8.ListBox2.Items.Add("M23")
        Form8.ListBox2.Items.Add("M24")
        Form8.ListBox2.Items.Add("M25")
        Form8.ListBox2.Items.Add("M26")
        Form8.ListBox2.Items.Add("M27")
        Form8.ListBox2.Items.Add("M28")
        Form8.ListBox2.Items.Add("M29")
        Form8.ListBox2.Items.Add("M30")
        Form8.ListBox2.Items.Add("M31")
        Form8.ListBox2.Items.Add("M32")
        Form8.ListBox2.Items.Add("M33")
        Form8.ListBox2.Items.Add("M34")
        Form8.ListBox2.Items.Add("M35")
        Form8.ListBox2.Items.Add("M36")
        Form8.ListBox2.Items.Add("M37")
        Form8.ListBox2.Items.Add("M38")
        Form8.ListBox2.Items.Add("M39")
        Form8.ListBox2.Items.Add("M40")

        Form8.ListBox3.Items.Add("120")
        Form8.ListBox3.Items.Add("208")
        Form8.ListBox3.Items.Add("240")
        Form8.ListBox3.Items.Add("277")
        Form8.ListBox3.Items.Add("347")
        Form8.ListBox3.Items.Add("416")
        Form8.ListBox3.Items.Add("480")
        Form8.ListBox3.Items.Add("600")

        Form8.ListBox4.Items.Add("Bench1")
        Form8.ListBox4.Items.Add("Bench2")
        Form8.ListBox4.Items.Add("Bench3")

        Form8.ListBox5.Items.Add("2016")
        Form8.ListBox5.Items.Add("2017")
        Form8.ListBox5.Items.Add("2018")
        Form8.ListBox5.Items.Add("2019")
        Form8.ListBox5.Items.Add("2020")
        Form8.ListBox5.Items.Add("2021")
        Form8.ListBox5.Items.Add("2022")
        Form8.ListBox5.Items.Add("2023")
        Form8.ListBox5.Items.Add("2024")
        Form8.ListBox5.Items.Add("2025")

        Form8.ListBox6.Items.Add("2016")
        Form8.ListBox6.Items.Add("2017")
        Form8.ListBox6.Items.Add("2018")
        Form8.ListBox6.Items.Add("2019")
        Form8.ListBox6.Items.Add("2020")
        Form8.ListBox6.Items.Add("2021")
        Form8.ListBox6.Items.Add("2022")
        Form8.ListBox6.Items.Add("2023")
        Form8.ListBox6.Items.Add("2024")
        Form8.ListBox6.Items.Add("2025")

        Form8.ListBox7.Items.Add("Bench1")
        Form8.ListBox7.Items.Add("Bench2")
        Form8.ListBox7.Items.Add("Bench3")



        Dim year As String
        Dim time As DateTime = DateTime.Now
        Dim x As Integer

        year = time.ToString("yyyy")
        x = CInt(year) - 2016
        Form8.ListBox5.SelectedIndex = x
        Form8.ListBox6.SelectedIndex = x

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If FTestApp.Enabled Then
            FTestApp.Close()
        End If
    End Sub
End Class
