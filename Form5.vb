Imports System.IO
Imports NationalInstruments.VisaNS
Imports System.ComponentModel
Imports System.IO.Ports
Imports System
Imports System.Threading
Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports AdvDIOLib
Imports AxAdvDIOLib
Imports SourceGrid2
Imports AxDSOFramer
Imports AxMSTSCLib
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices









Public Class FTestApp

    Inherits System.Windows.Forms.Form
    Public EMPSerialPort As New SerialPort("COM1", 9600)
    Public comDevice As Integer 'Global handle to RD/RS device connected
    Public Status As Integer    'Global status variable to communicate with DLL
    Public PhaseCount As Short  'Global phase count for device connected
    Public TestType As String   'Global test type for time/pulse test timer
    Public TapPhase As Byte     'Global variable for tap change phase selection
    Public oBook As Object
    Public oSheet As Object
    Public oExcel As Object
    Public repx As Integer
    Public stopflag As Integer
    Public CBNUM As String
    Public CTNUM As String
    Public tboxprn As String
    Public ACCTESTFLAG As Integer = 0
    Public Daily_Accuracy_Test As Integer = 0 'Global flag that is set by default to 0 but is set to 1 when called from Daily Accuracy Check Menu Item
    Public initials As String = ""
    Public strFileName() As String '// String Array.
    Public dic As New Dictionary(Of String, String)
    Public xstr As String
    Public checkboxarray() As String = {"CheckBox2"}
    Public Uniq_ID_Flag As Integer = 0
    Dim UniqID_text As Integer







#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.

    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents tbcDocsCont As System.Windows.Forms.TabControl
    Friend WithEvents OFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents tbDoc1 As System.Windows.Forms.TabPage
    Friend WithEvents tbDoc2 As System.Windows.Forms.TabPage
    Friend WithEvents tbDoc3 As System.Windows.Forms.TabPage
    Friend WithEvents tbDoc4 As System.Windows.Forms.TabPage

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FTestApp))
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.tbcDocsCont = New System.Windows.Forms.TabControl()
        Me.tbMain = New System.Windows.Forms.TabPage()
        Me.AxInstantDoCtrl3 = New AxBDaqOcxLib.AxInstantDoCtrl()
        Me.AxInstantDoCtrl2 = New AxBDaqOcxLib.AxInstantDoCtrl()
        Me.AxInstantDoCtrl1 = New AxBDaqOcxLib.AxInstantDoCtrl()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Label50 = New System.Windows.Forms.Label()
        Me.TextBox53 = New System.Windows.Forms.TextBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox52 = New System.Windows.Forms.TextBox()
        Me.TextBox51 = New System.Windows.Forms.TextBox()
        Me.TextBox50 = New System.Windows.Forms.TextBox()
        Me.TextBox49 = New System.Windows.Forms.TextBox()
        Me.AxAdvDIO2 = New AxAdvDIOLib.AxAdvDIO()
        Me.AxAdvDIO1 = New AxAdvDIOLib.AxAdvDIO()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button12 = New System.Windows.Forms.Button()
        Me.TextBox21 = New System.Windows.Forms.TextBox()
        Me.TextBox20 = New System.Windows.Forms.TextBox()
        Me.TextBox19 = New System.Windows.Forms.TextBox()
        Me.TextBox18 = New System.Windows.Forms.TextBox()
        Me.TextBox16 = New System.Windows.Forms.TextBox()
        Me.TextBox15 = New System.Windows.Forms.TextBox()
        Me.TextBox10 = New System.Windows.Forms.TextBox()
        Me.TextBox9 = New System.Windows.Forms.TextBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox8 = New System.Windows.Forms.TextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.SETRELAYS = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnOpenFile = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label49 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.TextBox41 = New System.Windows.Forms.TextBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.TextBox40 = New System.Windows.Forms.TextBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.TextBox39 = New System.Windows.Forms.TextBox()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.TextBox38 = New System.Windows.Forms.TextBox()
        Me.TextBox36 = New System.Windows.Forms.TextBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.TextBox35 = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.TextBox34 = New System.Windows.Forms.TextBox()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.TextBox33 = New System.Windows.Forms.TextBox()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.TextBox32 = New System.Windows.Forms.TextBox()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.TextBox31 = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.TextBox30 = New System.Windows.Forms.TextBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.TextBox29 = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.TextBox42 = New System.Windows.Forms.TextBox()
        Me.TextBox44 = New System.Windows.Forms.TextBox()
        Me.TextBox45 = New System.Windows.Forms.TextBox()
        Me.TextBox46 = New System.Windows.Forms.TextBox()
        Me.TextBox37 = New System.Windows.Forms.TextBox()
        Me.TextBox47 = New System.Windows.Forms.TextBox()
        Me.TextBox48 = New System.Windows.Forms.TextBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.TextBox43 = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.TextBox17 = New System.Windows.Forms.TextBox()
        Me.TextBox25 = New System.Windows.Forms.TextBox()
        Me.TextBox22 = New System.Windows.Forms.TextBox()
        Me.TextBox28 = New System.Windows.Forms.TextBox()
        Me.TextBox27 = New System.Windows.Forms.TextBox()
        Me.TextBox26 = New System.Windows.Forms.TextBox()
        Me.TextBox24 = New System.Windows.Forms.TextBox()
        Me.TextBox23 = New System.Windows.Forms.TextBox()
        Me.TextBox11 = New System.Windows.Forms.TextBox()
        Me.TextBox12 = New System.Windows.Forms.TextBox()
        Me.TextBox13 = New System.Windows.Forms.TextBox()
        Me.TextBox14 = New System.Windows.Forms.TextBox()
        Me.TextBox54 = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton40 = New System.Windows.Forms.RadioButton()
        Me.RadioButton39 = New System.Windows.Forms.RadioButton()
        Me.RadioButton38 = New System.Windows.Forms.RadioButton()
        Me.RadioButton37 = New System.Windows.Forms.RadioButton()
        Me.RadioButton36 = New System.Windows.Forms.RadioButton()
        Me.RadioButton35 = New System.Windows.Forms.RadioButton()
        Me.RadioButton34 = New System.Windows.Forms.RadioButton()
        Me.RadioButton33 = New System.Windows.Forms.RadioButton()
        Me.RadioButton32 = New System.Windows.Forms.RadioButton()
        Me.RadioButton31 = New System.Windows.Forms.RadioButton()
        Me.RadioButton30 = New System.Windows.Forms.RadioButton()
        Me.RadioButton29 = New System.Windows.Forms.RadioButton()
        Me.RadioButton28 = New System.Windows.Forms.RadioButton()
        Me.RadioButton27 = New System.Windows.Forms.RadioButton()
        Me.RadioButton26 = New System.Windows.Forms.RadioButton()
        Me.RadioButton25 = New System.Windows.Forms.RadioButton()
        Me.RadioButton24 = New System.Windows.Forms.RadioButton()
        Me.RadioButton23 = New System.Windows.Forms.RadioButton()
        Me.RadioButton22 = New System.Windows.Forms.RadioButton()
        Me.RadioButton21 = New System.Windows.Forms.RadioButton()
        Me.RadioButton20 = New System.Windows.Forms.RadioButton()
        Me.RadioButton19 = New System.Windows.Forms.RadioButton()
        Me.RadioButton18 = New System.Windows.Forms.RadioButton()
        Me.RadioButton17 = New System.Windows.Forms.RadioButton()
        Me.RadioButton16 = New System.Windows.Forms.RadioButton()
        Me.RadioButton15 = New System.Windows.Forms.RadioButton()
        Me.RadioButton14 = New System.Windows.Forms.RadioButton()
        Me.RadioButton13 = New System.Windows.Forms.RadioButton()
        Me.RadioButton12 = New System.Windows.Forms.RadioButton()
        Me.RadioButton11 = New System.Windows.Forms.RadioButton()
        Me.RadioButton10 = New System.Windows.Forms.RadioButton()
        Me.RadioButton9 = New System.Windows.Forms.RadioButton()
        Me.RadioButton8 = New System.Windows.Forms.RadioButton()
        Me.RadioButton7 = New System.Windows.Forms.RadioButton()
        Me.RadioButton6 = New System.Windows.Forms.RadioButton()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.CheckBox40 = New System.Windows.Forms.CheckBox()
        Me.CheckBox39 = New System.Windows.Forms.CheckBox()
        Me.CheckBox38 = New System.Windows.Forms.CheckBox()
        Me.CheckBox37 = New System.Windows.Forms.CheckBox()
        Me.CheckBox36 = New System.Windows.Forms.CheckBox()
        Me.CheckBox35 = New System.Windows.Forms.CheckBox()
        Me.CheckBox34 = New System.Windows.Forms.CheckBox()
        Me.CheckBox33 = New System.Windows.Forms.CheckBox()
        Me.CheckBox32 = New System.Windows.Forms.CheckBox()
        Me.CheckBox31 = New System.Windows.Forms.CheckBox()
        Me.CheckBox30 = New System.Windows.Forms.CheckBox()
        Me.CheckBox29 = New System.Windows.Forms.CheckBox()
        Me.CheckBox28 = New System.Windows.Forms.CheckBox()
        Me.CheckBox27 = New System.Windows.Forms.CheckBox()
        Me.CheckBox26 = New System.Windows.Forms.CheckBox()
        Me.CheckBox25 = New System.Windows.Forms.CheckBox()
        Me.CheckBox24 = New System.Windows.Forms.CheckBox()
        Me.CheckBox23 = New System.Windows.Forms.CheckBox()
        Me.CheckBox22 = New System.Windows.Forms.CheckBox()
        Me.CheckBox21 = New System.Windows.Forms.CheckBox()
        Me.CheckBox20 = New System.Windows.Forms.CheckBox()
        Me.CheckBox19 = New System.Windows.Forms.CheckBox()
        Me.CheckBox18 = New System.Windows.Forms.CheckBox()
        Me.CheckBox17 = New System.Windows.Forms.CheckBox()
        Me.CheckBox16 = New System.Windows.Forms.CheckBox()
        Me.CheckBox15 = New System.Windows.Forms.CheckBox()
        Me.CheckBox14 = New System.Windows.Forms.CheckBox()
        Me.CheckBox13 = New System.Windows.Forms.CheckBox()
        Me.CheckBox12 = New System.Windows.Forms.CheckBox()
        Me.CheckBox11 = New System.Windows.Forms.CheckBox()
        Me.CheckBox10 = New System.Windows.Forms.CheckBox()
        Me.CheckBox9 = New System.Windows.Forms.CheckBox()
        Me.CheckBox8 = New System.Windows.Forms.CheckBox()
        Me.CheckBox7 = New System.Windows.Forms.CheckBox()
        Me.CheckBox6 = New System.Windows.Forms.CheckBox()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.tbDoc4 = New System.Windows.Forms.TabPage()
        Me.axFramer4 = New AxDSOFramer.AxFramerControl()
        Me.tbDoc1 = New System.Windows.Forms.TabPage()
        Me.axFramer1 = New AxDSOFramer.AxFramerControl()
        Me.tbDoc2 = New System.Windows.Forms.TabPage()
        Me.axFramer2 = New AxDSOFramer.AxFramerControl()
        Me.tbDoc3 = New System.Windows.Forms.TabPage()
        Me.axFramer3 = New AxDSOFramer.AxFramerControl()
        Me.OFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker()
        Me.AxInstantDoCtrl4 = New AxBDaqOcxLib.AxInstantDoCtrl()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbcDocsCont.SuspendLayout()
        Me.tbMain.SuspendLayout()
        CType(Me.AxInstantDoCtrl3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxInstantDoCtrl2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxInstantDoCtrl1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxAdvDIO2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxAdvDIO1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.axFramer4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.axFramer1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.axFramer2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.axFramer3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxInstantDoCtrl4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picLogo
        '
        Me.picLogo.Location = New System.Drawing.Point(0, 0)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(224, 56)
        Me.picLogo.TabIndex = 0
        Me.picLogo.TabStop = False
        '
        'tbcDocsCont
        '
        Me.tbcDocsCont.Controls.Add(Me.tbMain)
        Me.tbcDocsCont.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbcDocsCont.Location = New System.Drawing.Point(0, 0)
        Me.tbcDocsCont.Name = "tbcDocsCont"
        Me.tbcDocsCont.SelectedIndex = 0
        Me.tbcDocsCont.Size = New System.Drawing.Size(1486, 998)
        Me.tbcDocsCont.TabIndex = 2
        '
        'tbMain
        '
        Me.tbMain.BackColor = System.Drawing.Color.LightSteelBlue
        Me.tbMain.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.tbMain.Controls.Add(Me.AxInstantDoCtrl3)
        Me.tbMain.Controls.Add(Me.AxInstantDoCtrl2)
        Me.tbMain.Controls.Add(Me.AxInstantDoCtrl1)
        Me.tbMain.Controls.Add(Me.AxInstantDoCtrl4)
        Me.tbMain.Controls.Add(Me.Button13)
        Me.tbMain.Controls.Add(Me.Label36)
        Me.tbMain.Controls.Add(Me.Label50)
        Me.tbMain.Controls.Add(Me.TextBox53)
        Me.tbMain.Controls.Add(Me.PictureBox1)
        Me.tbMain.Controls.Add(Me.Label9)
        Me.tbMain.Controls.Add(Me.Label8)
        Me.tbMain.Controls.Add(Me.TextBox52)
        Me.tbMain.Controls.Add(Me.TextBox51)
        Me.tbMain.Controls.Add(Me.TextBox50)
        Me.tbMain.Controls.Add(Me.TextBox49)
        Me.tbMain.Controls.Add(Me.AxAdvDIO2)
        Me.tbMain.Controls.Add(Me.AxAdvDIO1)
        Me.tbMain.Controls.Add(Me.Button3)
        Me.tbMain.Controls.Add(Me.Button12)
        Me.tbMain.Controls.Add(Me.TextBox21)
        Me.tbMain.Controls.Add(Me.TextBox20)
        Me.tbMain.Controls.Add(Me.TextBox19)
        Me.tbMain.Controls.Add(Me.TextBox18)
        Me.tbMain.Controls.Add(Me.TextBox16)
        Me.tbMain.Controls.Add(Me.TextBox15)
        Me.tbMain.Controls.Add(Me.TextBox10)
        Me.tbMain.Controls.Add(Me.TextBox9)
        Me.tbMain.Controls.Add(Me.ComboBox1)
        Me.tbMain.Controls.Add(Me.Button11)
        Me.tbMain.Controls.Add(Me.Button10)
        Me.tbMain.Controls.Add(Me.Button9)
        Me.tbMain.Controls.Add(Me.Button8)
        Me.tbMain.Controls.Add(Me.Button7)
        Me.tbMain.Controls.Add(Me.Button5)
        Me.tbMain.Controls.Add(Me.Label3)
        Me.tbMain.Controls.Add(Me.TextBox4)
        Me.tbMain.Controls.Add(Me.Label7)
        Me.tbMain.Controls.Add(Me.TextBox1)
        Me.tbMain.Controls.Add(Me.Button6)
        Me.tbMain.Controls.Add(Me.Button1)
        Me.tbMain.Controls.Add(Me.TextBox8)
        Me.tbMain.Controls.Add(Me.Button4)
        Me.tbMain.Controls.Add(Me.Label6)
        Me.tbMain.Controls.Add(Me.Label5)
        Me.tbMain.Controls.Add(Me.Label4)
        Me.tbMain.Controls.Add(Me.Label2)
        Me.tbMain.Controls.Add(Me.Label1)
        Me.tbMain.Controls.Add(Me.TextBox7)
        Me.tbMain.Controls.Add(Me.TextBox6)
        Me.tbMain.Controls.Add(Me.TextBox5)
        Me.tbMain.Controls.Add(Me.TextBox3)
        Me.tbMain.Controls.Add(Me.TextBox2)
        Me.tbMain.Controls.Add(Me.SETRELAYS)
        Me.tbMain.Controls.Add(Me.Button2)
        Me.tbMain.Controls.Add(Me.btnOpenFile)
        Me.tbMain.Controls.Add(Me.GroupBox2)
        Me.tbMain.Controls.Add(Me.TextBox54)
        Me.tbMain.Controls.Add(Me.GroupBox1)
        Me.tbMain.Controls.Add(Me.GroupBox3)
        Me.tbMain.Location = New System.Drawing.Point(4, 22)
        Me.tbMain.Name = "tbMain"
        Me.tbMain.Size = New System.Drawing.Size(1478, 972)
        Me.tbMain.TabIndex = 1
        Me.tbMain.Text = "Main"
        '
        'AxInstantDoCtrl3
        '
        Me.AxInstantDoCtrl3.Enabled = True
        Me.AxInstantDoCtrl3.Location = New System.Drawing.Point(345, 852)
        Me.AxInstantDoCtrl3.Name = "AxInstantDoCtrl3"
        Me.AxInstantDoCtrl3.OcxState = CType(resources.GetObject("AxInstantDoCtrl3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxInstantDoCtrl3.Size = New System.Drawing.Size(17, 17)
        Me.AxInstantDoCtrl3.TabIndex = 212
        '
        'AxInstantDoCtrl2
        '
        Me.AxInstantDoCtrl2.Enabled = True
        Me.AxInstantDoCtrl2.Location = New System.Drawing.Point(205, 852)
        Me.AxInstantDoCtrl2.Name = "AxInstantDoCtrl2"
        Me.AxInstantDoCtrl2.OcxState = CType(resources.GetObject("AxInstantDoCtrl2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxInstantDoCtrl2.Size = New System.Drawing.Size(17, 17)
        Me.AxInstantDoCtrl2.TabIndex = 211
        '
        'AxInstantDoCtrl1
        '
        Me.AxInstantDoCtrl1.Enabled = True
        Me.AxInstantDoCtrl1.Location = New System.Drawing.Point(21, 852)
        Me.AxInstantDoCtrl1.Name = "AxInstantDoCtrl1"
        Me.AxInstantDoCtrl1.OcxState = CType(resources.GetObject("AxInstantDoCtrl1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxInstantDoCtrl1.Size = New System.Drawing.Size(17, 17)
        Me.AxInstantDoCtrl1.TabIndex = 210
        '
        'Button13
        '
        Me.Button13.Location = New System.Drawing.Point(705, 770)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(187, 58)
        Me.Button13.TabIndex = 212
        Me.Button13.Text = "Manual Test Posittions"
        Me.Button13.UseVisualStyleBackColor = True
        Me.Button13.Visible = False
        '
        'Label36
        '
        Me.Label36.Location = New System.Drawing.Point(646, 647)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(324, 85)
        Me.Label36.TabIndex = 210
        '
        'Label50
        '
        Me.Label50.AutoSize = True
        Me.Label50.Location = New System.Drawing.Point(684, 592)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(111, 13)
        Me.Label50.TabIndex = 209
        Me.Label50.Text = "Repeat Cycle Number"
        '
        'TextBox53
        '
        Me.TextBox53.Location = New System.Drawing.Point(841, 589)
        Me.TextBox53.Name = "TextBox53"
        Me.TextBox53.Size = New System.Drawing.Size(100, 20)
        Me.TextBox53.TabIndex = 208
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.PictureBox1.InitialImage = CType(resources.GetObject("PictureBox1.InitialImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(30, 770)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(440, 95)
        Me.PictureBox1.TabIndex = 187
        Me.PictureBox1.TabStop = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(853, 359)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 13)
        Me.Label9.TabIndex = 181
        Me.Label9.Text = "% Error"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(703, 359)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(66, 13)
        Me.Label8.TabIndex = 180
        Me.Label8.Text = "Current Step"
        '
        'TextBox52
        '
        Me.TextBox52.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox52.Location = New System.Drawing.Point(841, 526)
        Me.TextBox52.Multiline = True
        Me.TextBox52.Name = "TextBox52"
        Me.TextBox52.Size = New System.Drawing.Size(108, 29)
        Me.TextBox52.TabIndex = 179
        '
        'TextBox51
        '
        Me.TextBox51.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox51.Location = New System.Drawing.Point(841, 484)
        Me.TextBox51.Multiline = True
        Me.TextBox51.Name = "TextBox51"
        Me.TextBox51.Size = New System.Drawing.Size(108, 30)
        Me.TextBox51.TabIndex = 178
        '
        'TextBox50
        '
        Me.TextBox50.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox50.Location = New System.Drawing.Point(841, 434)
        Me.TextBox50.Multiline = True
        Me.TextBox50.Name = "TextBox50"
        Me.TextBox50.Size = New System.Drawing.Size(108, 34)
        Me.TextBox50.TabIndex = 177
        '
        'TextBox49
        '
        Me.TextBox49.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox49.Location = New System.Drawing.Point(841, 387)
        Me.TextBox49.Multiline = True
        Me.TextBox49.Name = "TextBox49"
        Me.TextBox49.Size = New System.Drawing.Size(108, 31)
        Me.TextBox49.TabIndex = 176
        '
        'AxAdvDIO2
        '
        Me.AxAdvDIO2.Enabled = True
        Me.AxAdvDIO2.Location = New System.Drawing.Point(1487, 750)
        Me.AxAdvDIO2.Name = "AxAdvDIO2"
        Me.AxAdvDIO2.OcxState = CType(resources.GetObject("AxAdvDIO2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxAdvDIO2.Size = New System.Drawing.Size(33, 33)
        Me.AxAdvDIO2.TabIndex = 175
        '
        'AxAdvDIO1
        '
        Me.AxAdvDIO1.Enabled = True
        Me.AxAdvDIO1.Location = New System.Drawing.Point(1487, 21)
        Me.AxAdvDIO1.Name = "AxAdvDIO1"
        Me.AxAdvDIO1.OcxState = CType(resources.GetObject("AxAdvDIO1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxAdvDIO1.Size = New System.Drawing.Size(33, 33)
        Me.AxAdvDIO1.TabIndex = 174
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(124, 622)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(96, 23)
        Me.Button3.TabIndex = 172
        Me.Button3.Text = "Repeatability"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button12
        '
        Me.Button12.Location = New System.Drawing.Point(1193, 686)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(75, 24)
        Me.Button12.TabIndex = 171
        Me.Button12.Text = "Read EMP"
        Me.Button12.UseVisualStyleBackColor = True
        '
        'TextBox21
        '
        Me.TextBox21.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox21.Location = New System.Drawing.Point(1053, 353)
        Me.TextBox21.Name = "TextBox21"
        Me.TextBox21.Size = New System.Drawing.Size(100, 21)
        Me.TextBox21.TabIndex = 143
        '
        'TextBox20
        '
        Me.TextBox20.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox20.Location = New System.Drawing.Point(1053, 330)
        Me.TextBox20.Name = "TextBox20"
        Me.TextBox20.Size = New System.Drawing.Size(100, 21)
        Me.TextBox20.TabIndex = 142
        '
        'TextBox19
        '
        Me.TextBox19.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox19.Location = New System.Drawing.Point(1053, 307)
        Me.TextBox19.Name = "TextBox19"
        Me.TextBox19.Size = New System.Drawing.Size(100, 21)
        Me.TextBox19.TabIndex = 141
        '
        'TextBox18
        '
        Me.TextBox18.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox18.Location = New System.Drawing.Point(1053, 284)
        Me.TextBox18.Name = "TextBox18"
        Me.TextBox18.Size = New System.Drawing.Size(100, 21)
        Me.TextBox18.TabIndex = 140
        '
        'TextBox16
        '
        Me.TextBox16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox16.Location = New System.Drawing.Point(1053, 238)
        Me.TextBox16.Name = "TextBox16"
        Me.TextBox16.Size = New System.Drawing.Size(100, 21)
        Me.TextBox16.TabIndex = 138
        '
        'TextBox15
        '
        Me.TextBox15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox15.Location = New System.Drawing.Point(1053, 215)
        Me.TextBox15.Name = "TextBox15"
        Me.TextBox15.Size = New System.Drawing.Size(100, 21)
        Me.TextBox15.TabIndex = 137
        '
        'TextBox10
        '
        Me.TextBox10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox10.Location = New System.Drawing.Point(1053, 103)
        Me.TextBox10.Name = "TextBox10"
        Me.TextBox10.Size = New System.Drawing.Size(100, 21)
        Me.TextBox10.TabIndex = 132
        '
        'TextBox9
        '
        Me.TextBox9.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox9.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox9.Location = New System.Drawing.Point(1053, 77)
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.Size = New System.Drawing.Size(100, 20)
        Me.TextBox9.TabIndex = 131
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(119, 651)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(101, 21)
        Me.ComboBox1.TabIndex = 130
        '
        'Button11
        '
        Me.Button11.Location = New System.Drawing.Point(822, 273)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(75, 23)
        Me.Button11.TabIndex = 127
        Me.Button11.Text = "-"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(720, 273)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(75, 23)
        Me.Button10.TabIndex = 126
        Me.Button10.Text = "+"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(822, 175)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(75, 23)
        Me.Button9.TabIndex = 125
        Me.Button9.Text = "-"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(720, 175)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(75, 23)
        Me.Button8.TabIndex = 124
        Me.Button8.Text = "+"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(822, 83)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(75, 23)
        Me.Button7.TabIndex = 123
        Me.Button7.Text = "-"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(720, 83)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(75, 23)
        Me.Button5.TabIndex = 122
        Me.Button5.Text = "+"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(684, 230)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(70, 13)
        Me.Label3.TabIndex = 121
        Me.Label3.Text = "Power Factor"
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(768, 227)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(100, 20)
        Me.TextBox4.TabIndex = 120
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(703, 132)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(41, 13)
        Me.Label7.TabIndex = 119
        Me.Label7.Text = "Current"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(767, 132)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 118
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(124, 589)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(96, 23)
        Me.Button6.TabIndex = 117
        Me.Button6.Text = "Turn Off Relays"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(13, 618)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 23)
        Me.Button1.TabIndex = 115
        Me.Button1.Text = "Pause"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox8
        '
        Me.TextBox8.Location = New System.Drawing.Point(267, 40)
        Me.TextBox8.Multiline = True
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(356, 605)
        Me.TextBox8.TabIndex = 114
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(267, 687)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(122, 23)
        Me.Button4.TabIndex = 111
        Me.Button4.Text = "Open PRN File"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(629, 525)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(61, 13)
        Me.Label6.TabIndex = 109
        Me.Label6.Text = "VAh Check"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(632, 481)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 13)
        Me.Label5.TabIndex = 108
        Me.Label5.Text = "Wh Check"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(629, 436)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 13)
        Me.Label4.TabIndex = 107
        Me.Label4.Text = "Wh Check"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(629, 390)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 101
        Me.Label2.Text = "Wh Check"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(702, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "Voltage"
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(696, 525)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(100, 20)
        Me.TextBox7.TabIndex = 99
        Me.TextBox7.Text = "25% @ PF"
        Me.TextBox7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(695, 481)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(100, 20)
        Me.TextBox6.TabIndex = 98
        Me.TextBox6.Text = "25% @ PF"
        Me.TextBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(695, 436)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(100, 20)
        Me.TextBox5.TabIndex = 97
        Me.TextBox5.Text = "25% Unity"
        Me.TextBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(695, 390)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(100, 20)
        Me.TextBox3.TabIndex = 95
        Me.TextBox3.Text = "2.5% Unity"
        Me.TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(767, 40)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(100, 20)
        Me.TextBox2.TabIndex = 94
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'SETRELAYS
        '
        Me.SETRELAYS.Location = New System.Drawing.Point(13, 589)
        Me.SETRELAYS.Name = "SETRELAYS"
        Me.SETRELAYS.Size = New System.Drawing.Size(81, 23)
        Me.SETRELAYS.TabIndex = 50
        Me.SETRELAYS.Text = "Start Test"
        Me.SETRELAYS.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(14, 647)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(80, 23)
        Me.Button2.TabIndex = 46
        Me.Button2.Text = "Stop Test"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'btnOpenFile
        '
        Me.btnOpenFile.Location = New System.Drawing.Point(470, 687)
        Me.btnOpenFile.Name = "btnOpenFile"
        Me.btnOpenFile.Size = New System.Drawing.Size(153, 23)
        Me.btnOpenFile.TabIndex = 1
        Me.btnOpenFile.Text = "Get Result File"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.SystemColors.ControlLight
        Me.GroupBox2.Controls.Add(Me.Label49)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.Label48)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.Label47)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Label46)
        Me.GroupBox2.Controls.Add(Me.Label45)
        Me.GroupBox2.Controls.Add(Me.Label44)
        Me.GroupBox2.Controls.Add(Me.Label43)
        Me.GroupBox2.Controls.Add(Me.Label42)
        Me.GroupBox2.Controls.Add(Me.Label41)
        Me.GroupBox2.Controls.Add(Me.Label40)
        Me.GroupBox2.Controls.Add(Me.TextBox41)
        Me.GroupBox2.Controls.Add(Me.Label39)
        Me.GroupBox2.Controls.Add(Me.TextBox40)
        Me.GroupBox2.Controls.Add(Me.Label38)
        Me.GroupBox2.Controls.Add(Me.TextBox39)
        Me.GroupBox2.Controls.Add(Me.Label37)
        Me.GroupBox2.Controls.Add(Me.TextBox38)
        Me.GroupBox2.Controls.Add(Me.TextBox36)
        Me.GroupBox2.Controls.Add(Me.Label35)
        Me.GroupBox2.Controls.Add(Me.TextBox35)
        Me.GroupBox2.Controls.Add(Me.Label34)
        Me.GroupBox2.Controls.Add(Me.TextBox34)
        Me.GroupBox2.Controls.Add(Me.Label33)
        Me.GroupBox2.Controls.Add(Me.TextBox33)
        Me.GroupBox2.Controls.Add(Me.Label32)
        Me.GroupBox2.Controls.Add(Me.TextBox32)
        Me.GroupBox2.Controls.Add(Me.Label31)
        Me.GroupBox2.Controls.Add(Me.TextBox31)
        Me.GroupBox2.Controls.Add(Me.Label30)
        Me.GroupBox2.Controls.Add(Me.TextBox30)
        Me.GroupBox2.Controls.Add(Me.Label29)
        Me.GroupBox2.Controls.Add(Me.TextBox29)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.Label28)
        Me.GroupBox2.Controls.Add(Me.Label27)
        Me.GroupBox2.Controls.Add(Me.Label26)
        Me.GroupBox2.Controls.Add(Me.TextBox42)
        Me.GroupBox2.Controls.Add(Me.TextBox44)
        Me.GroupBox2.Controls.Add(Me.TextBox45)
        Me.GroupBox2.Controls.Add(Me.TextBox46)
        Me.GroupBox2.Controls.Add(Me.TextBox37)
        Me.GroupBox2.Controls.Add(Me.TextBox47)
        Me.GroupBox2.Controls.Add(Me.TextBox48)
        Me.GroupBox2.Controls.Add(Me.Label25)
        Me.GroupBox2.Controls.Add(Me.Label24)
        Me.GroupBox2.Controls.Add(Me.Label23)
        Me.GroupBox2.Controls.Add(Me.Label22)
        Me.GroupBox2.Controls.Add(Me.Label21)
        Me.GroupBox2.Controls.Add(Me.TextBox43)
        Me.GroupBox2.Controls.Add(Me.Label20)
        Me.GroupBox2.Controls.Add(Me.Label19)
        Me.GroupBox2.Controls.Add(Me.Label18)
        Me.GroupBox2.Controls.Add(Me.Label17)
        Me.GroupBox2.Controls.Add(Me.Label16)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Controls.Add(Me.TextBox17)
        Me.GroupBox2.Controls.Add(Me.TextBox25)
        Me.GroupBox2.Controls.Add(Me.TextBox22)
        Me.GroupBox2.Controls.Add(Me.TextBox28)
        Me.GroupBox2.Controls.Add(Me.TextBox27)
        Me.GroupBox2.Controls.Add(Me.TextBox26)
        Me.GroupBox2.Controls.Add(Me.TextBox24)
        Me.GroupBox2.Controls.Add(Me.TextBox23)
        Me.GroupBox2.Controls.Add(Me.TextBox11)
        Me.GroupBox2.Controls.Add(Me.TextBox12)
        Me.GroupBox2.Controls.Add(Me.TextBox13)
        Me.GroupBox2.Controls.Add(Me.TextBox14)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(988, 57)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(418, 588)
        Me.GroupBox2.TabIndex = 185
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Relay Indicators"
        '
        'Label49
        '
        Me.Label49.AutoSize = True
        Me.Label49.Location = New System.Drawing.Point(234, 187)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(63, 15)
        Me.Label49.TabIndex = 220
        Me.Label49.Text = "Relay 28"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(4, 95)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(55, 15)
        Me.Label12.TabIndex = 184
        Me.Label12.Text = "Relay 4"
        '
        'Label48
        '
        Me.Label48.AutoSize = True
        Me.Label48.Location = New System.Drawing.Point(234, 472)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(63, 15)
        Me.Label48.TabIndex = 219
        Me.Label48.Text = "Relay 40"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(4, 71)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(55, 15)
        Me.Label11.TabIndex = 183
        Me.Label11.Text = "Relay 3"
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.Location = New System.Drawing.Point(236, 449)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(63, 15)
        Me.Label47.TabIndex = 218
        Me.Label47.Text = "Relay 39"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(4, 47)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(55, 15)
        Me.Label10.TabIndex = 182
        Me.Label10.Text = "Relay 2"
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Location = New System.Drawing.Point(234, 426)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(63, 15)
        Me.Label46.TabIndex = 217
        Me.Label46.Text = "Relay 38"
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Location = New System.Drawing.Point(234, 403)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(63, 15)
        Me.Label45.TabIndex = 216
        Me.Label45.Text = "Relay 37"
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(234, 378)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(63, 15)
        Me.Label44.TabIndex = 215
        Me.Label44.Text = "Relay 36"
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(234, 356)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(63, 15)
        Me.Label43.TabIndex = 214
        Me.Label43.Text = "Relay 35"
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(236, 329)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(63, 15)
        Me.Label42.TabIndex = 213
        Me.Label42.Text = "Relay 34"
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(234, 303)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(63, 15)
        Me.Label41.TabIndex = 212
        Me.Label41.Text = "Relay 33"
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(234, 282)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(63, 15)
        Me.Label40.TabIndex = 211
        Me.Label40.Text = "Relay 32"
        '
        'TextBox41
        '
        Me.TextBox41.Location = New System.Drawing.Point(303, 296)
        Me.TextBox41.Name = "TextBox41"
        Me.TextBox41.Size = New System.Drawing.Size(100, 21)
        Me.TextBox41.TabIndex = 163
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(234, 259)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(63, 15)
        Me.Label39.TabIndex = 210
        Me.Label39.Text = "Relay 31"
        '
        'TextBox40
        '
        Me.TextBox40.Location = New System.Drawing.Point(303, 277)
        Me.TextBox40.Name = "TextBox40"
        Me.TextBox40.Size = New System.Drawing.Size(100, 21)
        Me.TextBox40.TabIndex = 162
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Location = New System.Drawing.Point(234, 236)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(63, 15)
        Me.Label38.TabIndex = 209
        Me.Label38.Text = "Relay 30"
        '
        'TextBox39
        '
        Me.TextBox39.Location = New System.Drawing.Point(303, 254)
        Me.TextBox39.Name = "TextBox39"
        Me.TextBox39.Size = New System.Drawing.Size(100, 21)
        Me.TextBox39.TabIndex = 161
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(236, 211)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(63, 15)
        Me.Label37.TabIndex = 208
        Me.Label37.Text = "Relay 29"
        '
        'TextBox38
        '
        Me.TextBox38.Location = New System.Drawing.Point(303, 233)
        Me.TextBox38.Name = "TextBox38"
        Me.TextBox38.Size = New System.Drawing.Size(100, 21)
        Me.TextBox38.TabIndex = 160
        '
        'TextBox36
        '
        Me.TextBox36.Location = New System.Drawing.Point(303, 185)
        Me.TextBox36.Name = "TextBox36"
        Me.TextBox36.Size = New System.Drawing.Size(100, 21)
        Me.TextBox36.TabIndex = 158
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(234, 167)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(63, 15)
        Me.Label35.TabIndex = 206
        Me.Label35.Text = "Relay 27"
        '
        'TextBox35
        '
        Me.TextBox35.Location = New System.Drawing.Point(303, 163)
        Me.TextBox35.Name = "TextBox35"
        Me.TextBox35.Size = New System.Drawing.Size(100, 21)
        Me.TextBox35.TabIndex = 157
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(234, 147)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(63, 15)
        Me.Label34.TabIndex = 205
        Me.Label34.Text = "Relay 26"
        '
        'TextBox34
        '
        Me.TextBox34.Location = New System.Drawing.Point(303, 145)
        Me.TextBox34.Name = "TextBox34"
        Me.TextBox34.Size = New System.Drawing.Size(100, 21)
        Me.TextBox34.TabIndex = 156
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(234, 127)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(63, 15)
        Me.Label33.TabIndex = 204
        Me.Label33.Text = "Relay 25"
        '
        'TextBox33
        '
        Me.TextBox33.Location = New System.Drawing.Point(303, 121)
        Me.TextBox33.Name = "TextBox33"
        Me.TextBox33.Size = New System.Drawing.Size(100, 21)
        Me.TextBox33.TabIndex = 155
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(234, 103)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(63, 15)
        Me.Label32.TabIndex = 203
        Me.Label32.Text = "Relay 24"
        '
        'TextBox32
        '
        Me.TextBox32.Location = New System.Drawing.Point(303, 95)
        Me.TextBox32.Name = "TextBox32"
        Me.TextBox32.Size = New System.Drawing.Size(100, 21)
        Me.TextBox32.TabIndex = 154
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(234, 80)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(63, 15)
        Me.Label31.TabIndex = 202
        Me.Label31.Text = "Relay 23"
        '
        'TextBox31
        '
        Me.TextBox31.Location = New System.Drawing.Point(303, 71)
        Me.TextBox31.Name = "TextBox31"
        Me.TextBox31.Size = New System.Drawing.Size(100, 21)
        Me.TextBox31.TabIndex = 153
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(234, 55)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(63, 15)
        Me.Label30.TabIndex = 201
        Me.Label30.Text = "Relay 22"
        '
        'TextBox30
        '
        Me.TextBox30.Location = New System.Drawing.Point(303, 47)
        Me.TextBox30.Name = "TextBox30"
        Me.TextBox30.Size = New System.Drawing.Size(100, 21)
        Me.TextBox30.TabIndex = 152
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(234, 29)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(63, 15)
        Me.Label29.TabIndex = 200
        Me.Label29.Text = "Relay 21"
        '
        'TextBox29
        '
        Me.TextBox29.Location = New System.Drawing.Point(303, 22)
        Me.TextBox29.Name = "TextBox29"
        Me.TextBox29.Size = New System.Drawing.Size(100, 21)
        Me.TextBox29.TabIndex = 151
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(4, 163)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(55, 15)
        Me.Label15.TabIndex = 186
        Me.Label15.Text = "Relay 7"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(-1, 471)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(63, 15)
        Me.Label28.TabIndex = 199
        Me.Label28.Text = "Relay 20"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(-1, 447)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(63, 15)
        Me.Label27.TabIndex = 198
        Me.Label27.Text = "Relay 19"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(-1, 426)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(63, 15)
        Me.Label26.TabIndex = 197
        Me.Label26.Text = "Relay 18"
        '
        'TextBox42
        '
        Me.TextBox42.Location = New System.Drawing.Point(303, 322)
        Me.TextBox42.Name = "TextBox42"
        Me.TextBox42.Size = New System.Drawing.Size(100, 21)
        Me.TextBox42.TabIndex = 164
        '
        'TextBox44
        '
        Me.TextBox44.Location = New System.Drawing.Point(303, 371)
        Me.TextBox44.Name = "TextBox44"
        Me.TextBox44.Size = New System.Drawing.Size(100, 21)
        Me.TextBox44.TabIndex = 166
        '
        'TextBox45
        '
        Me.TextBox45.Location = New System.Drawing.Point(303, 396)
        Me.TextBox45.Name = "TextBox45"
        Me.TextBox45.Size = New System.Drawing.Size(100, 21)
        Me.TextBox45.TabIndex = 167
        '
        'TextBox46
        '
        Me.TextBox46.Location = New System.Drawing.Point(303, 419)
        Me.TextBox46.Name = "TextBox46"
        Me.TextBox46.Size = New System.Drawing.Size(100, 21)
        Me.TextBox46.TabIndex = 168
        '
        'TextBox37
        '
        Me.TextBox37.Location = New System.Drawing.Point(303, 208)
        Me.TextBox37.Name = "TextBox37"
        Me.TextBox37.Size = New System.Drawing.Size(100, 21)
        Me.TextBox37.TabIndex = 159
        '
        'TextBox47
        '
        Me.TextBox47.Location = New System.Drawing.Point(303, 445)
        Me.TextBox47.Name = "TextBox47"
        Me.TextBox47.Size = New System.Drawing.Size(100, 21)
        Me.TextBox47.TabIndex = 169
        '
        'TextBox48
        '
        Me.TextBox48.Location = New System.Drawing.Point(303, 472)
        Me.TextBox48.Name = "TextBox48"
        Me.TextBox48.Size = New System.Drawing.Size(100, 21)
        Me.TextBox48.TabIndex = 170
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(-1, 403)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(63, 15)
        Me.Label25.TabIndex = 196
        Me.Label25.Text = "Relay 17"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(-3, 377)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(63, 15)
        Me.Label24.TabIndex = 195
        Me.Label24.Text = "Relay 16"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(-2, 351)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(63, 15)
        Me.Label23.TabIndex = 194
        Me.Label23.Text = "Relay 15"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(-3, 325)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(63, 15)
        Me.Label22.TabIndex = 193
        Me.Label22.Text = "Relay 14"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(-3, 303)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(63, 15)
        Me.Label21.TabIndex = 192
        Me.Label21.Text = "Relay 13"
        '
        'TextBox43
        '
        Me.TextBox43.Location = New System.Drawing.Point(303, 346)
        Me.TextBox43.Name = "TextBox43"
        Me.TextBox43.Size = New System.Drawing.Size(100, 21)
        Me.TextBox43.TabIndex = 165
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(-3, 282)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(63, 15)
        Me.Label20.TabIndex = 191
        Me.Label20.Text = "Relay 12"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(-3, 259)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(63, 15)
        Me.Label19.TabIndex = 190
        Me.Label19.Text = "Relay 11"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(-3, 234)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(63, 15)
        Me.Label18.TabIndex = 189
        Me.Label18.Text = "Relay 10"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(4, 211)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(55, 15)
        Me.Label17.TabIndex = 188
        Me.Label17.Text = "Relay 9"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(4, 188)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(55, 15)
        Me.Label16.TabIndex = 187
        Me.Label16.Text = "Relay 8"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(4, 141)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(55, 15)
        Me.Label14.TabIndex = 186
        Me.Label14.Text = "Relay 6"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(4, 118)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(55, 15)
        Me.Label13.TabIndex = 186
        Me.Label13.Text = "Relay 5"
        '
        'TextBox17
        '
        Me.TextBox17.Location = New System.Drawing.Point(65, 204)
        Me.TextBox17.Name = "TextBox17"
        Me.TextBox17.Size = New System.Drawing.Size(100, 21)
        Me.TextBox17.TabIndex = 139
        '
        'TextBox25
        '
        Me.TextBox25.Location = New System.Drawing.Point(65, 395)
        Me.TextBox25.Name = "TextBox25"
        Me.TextBox25.Size = New System.Drawing.Size(100, 21)
        Me.TextBox25.TabIndex = 147
        '
        'TextBox22
        '
        Me.TextBox22.Location = New System.Drawing.Point(65, 320)
        Me.TextBox22.Name = "TextBox22"
        Me.TextBox22.Size = New System.Drawing.Size(100, 21)
        Me.TextBox22.TabIndex = 144
        '
        'TextBox28
        '
        Me.TextBox28.Location = New System.Drawing.Point(65, 465)
        Me.TextBox28.Name = "TextBox28"
        Me.TextBox28.Size = New System.Drawing.Size(100, 21)
        Me.TextBox28.TabIndex = 150
        '
        'TextBox27
        '
        Me.TextBox27.Location = New System.Drawing.Point(65, 441)
        Me.TextBox27.Name = "TextBox27"
        Me.TextBox27.Size = New System.Drawing.Size(100, 21)
        Me.TextBox27.TabIndex = 149
        '
        'TextBox26
        '
        Me.TextBox26.Location = New System.Drawing.Point(65, 419)
        Me.TextBox26.Name = "TextBox26"
        Me.TextBox26.Size = New System.Drawing.Size(100, 21)
        Me.TextBox26.TabIndex = 148
        '
        'TextBox24
        '
        Me.TextBox24.Location = New System.Drawing.Point(65, 371)
        Me.TextBox24.Name = "TextBox24"
        Me.TextBox24.Size = New System.Drawing.Size(100, 21)
        Me.TextBox24.TabIndex = 146
        '
        'TextBox23
        '
        Me.TextBox23.Location = New System.Drawing.Point(65, 346)
        Me.TextBox23.Name = "TextBox23"
        Me.TextBox23.Size = New System.Drawing.Size(100, 21)
        Me.TextBox23.TabIndex = 145
        '
        'TextBox11
        '
        Me.TextBox11.Location = New System.Drawing.Point(65, 68)
        Me.TextBox11.Name = "TextBox11"
        Me.TextBox11.Size = New System.Drawing.Size(100, 21)
        Me.TextBox11.TabIndex = 133
        '
        'TextBox12
        '
        Me.TextBox12.Location = New System.Drawing.Point(65, 89)
        Me.TextBox12.Name = "TextBox12"
        Me.TextBox12.Size = New System.Drawing.Size(100, 21)
        Me.TextBox12.TabIndex = 134
        '
        'TextBox13
        '
        Me.TextBox13.Location = New System.Drawing.Point(65, 112)
        Me.TextBox13.Name = "TextBox13"
        Me.TextBox13.Size = New System.Drawing.Size(100, 21)
        Me.TextBox13.TabIndex = 135
        '
        'TextBox14
        '
        Me.TextBox14.Location = New System.Drawing.Point(65, 135)
        Me.TextBox14.Name = "TextBox14"
        Me.TextBox14.Size = New System.Drawing.Size(100, 21)
        Me.TextBox14.TabIndex = 136
        '
        'TextBox54
        '
        Me.TextBox54.Location = New System.Drawing.Point(988, 57)
        Me.TextBox54.Multiline = True
        Me.TextBox54.Name = "TextBox54"
        Me.TextBox54.Size = New System.Drawing.Size(418, 588)
        Me.TextBox54.TabIndex = 211
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.GroupBox1.Controls.Add(Me.RadioButton40)
        Me.GroupBox1.Controls.Add(Me.RadioButton39)
        Me.GroupBox1.Controls.Add(Me.RadioButton38)
        Me.GroupBox1.Controls.Add(Me.RadioButton37)
        Me.GroupBox1.Controls.Add(Me.RadioButton36)
        Me.GroupBox1.Controls.Add(Me.RadioButton35)
        Me.GroupBox1.Controls.Add(Me.RadioButton34)
        Me.GroupBox1.Controls.Add(Me.RadioButton33)
        Me.GroupBox1.Controls.Add(Me.RadioButton32)
        Me.GroupBox1.Controls.Add(Me.RadioButton31)
        Me.GroupBox1.Controls.Add(Me.RadioButton30)
        Me.GroupBox1.Controls.Add(Me.RadioButton29)
        Me.GroupBox1.Controls.Add(Me.RadioButton28)
        Me.GroupBox1.Controls.Add(Me.RadioButton27)
        Me.GroupBox1.Controls.Add(Me.RadioButton26)
        Me.GroupBox1.Controls.Add(Me.RadioButton25)
        Me.GroupBox1.Controls.Add(Me.RadioButton24)
        Me.GroupBox1.Controls.Add(Me.RadioButton23)
        Me.GroupBox1.Controls.Add(Me.RadioButton22)
        Me.GroupBox1.Controls.Add(Me.RadioButton21)
        Me.GroupBox1.Controls.Add(Me.RadioButton20)
        Me.GroupBox1.Controls.Add(Me.RadioButton19)
        Me.GroupBox1.Controls.Add(Me.RadioButton18)
        Me.GroupBox1.Controls.Add(Me.RadioButton17)
        Me.GroupBox1.Controls.Add(Me.RadioButton16)
        Me.GroupBox1.Controls.Add(Me.RadioButton15)
        Me.GroupBox1.Controls.Add(Me.RadioButton14)
        Me.GroupBox1.Controls.Add(Me.RadioButton13)
        Me.GroupBox1.Controls.Add(Me.RadioButton12)
        Me.GroupBox1.Controls.Add(Me.RadioButton11)
        Me.GroupBox1.Controls.Add(Me.RadioButton10)
        Me.GroupBox1.Controls.Add(Me.RadioButton9)
        Me.GroupBox1.Controls.Add(Me.RadioButton8)
        Me.GroupBox1.Controls.Add(Me.RadioButton7)
        Me.GroupBox1.Controls.Add(Me.RadioButton6)
        Me.GroupBox1.Controls.Add(Me.RadioButton5)
        Me.GroupBox1.Controls.Add(Me.RadioButton4)
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(13, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(240, 543)
        Me.GroupBox1.TabIndex = 93
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Automated Relay Test  Positions"
        '
        'RadioButton40
        '
        Me.RadioButton40.AutoSize = True
        Me.RadioButton40.Location = New System.Drawing.Point(128, 488)
        Me.RadioButton40.Name = "RadioButton40"
        Me.RadioButton40.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton40.TabIndex = 132
        Me.RadioButton40.TabStop = True
        Me.RadioButton40.Text = "Relay 40"
        Me.RadioButton40.UseVisualStyleBackColor = True
        '
        'RadioButton39
        '
        Me.RadioButton39.AutoSize = True
        Me.RadioButton39.Location = New System.Drawing.Point(128, 465)
        Me.RadioButton39.Name = "RadioButton39"
        Me.RadioButton39.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton39.TabIndex = 131
        Me.RadioButton39.TabStop = True
        Me.RadioButton39.Text = "Relay 39"
        Me.RadioButton39.UseVisualStyleBackColor = True
        '
        'RadioButton38
        '
        Me.RadioButton38.AutoSize = True
        Me.RadioButton38.Location = New System.Drawing.Point(128, 442)
        Me.RadioButton38.Name = "RadioButton38"
        Me.RadioButton38.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton38.TabIndex = 130
        Me.RadioButton38.TabStop = True
        Me.RadioButton38.Text = "Relay 38"
        Me.RadioButton38.UseVisualStyleBackColor = True
        '
        'RadioButton37
        '
        Me.RadioButton37.AutoSize = True
        Me.RadioButton37.Location = New System.Drawing.Point(128, 419)
        Me.RadioButton37.Name = "RadioButton37"
        Me.RadioButton37.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton37.TabIndex = 129
        Me.RadioButton37.TabStop = True
        Me.RadioButton37.Text = "Relay 37"
        Me.RadioButton37.UseVisualStyleBackColor = True
        '
        'RadioButton36
        '
        Me.RadioButton36.AutoSize = True
        Me.RadioButton36.Location = New System.Drawing.Point(128, 396)
        Me.RadioButton36.Name = "RadioButton36"
        Me.RadioButton36.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton36.TabIndex = 128
        Me.RadioButton36.TabStop = True
        Me.RadioButton36.Text = "Relay 36"
        Me.RadioButton36.UseVisualStyleBackColor = True
        '
        'RadioButton35
        '
        Me.RadioButton35.AutoSize = True
        Me.RadioButton35.Location = New System.Drawing.Point(128, 373)
        Me.RadioButton35.Name = "RadioButton35"
        Me.RadioButton35.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton35.TabIndex = 127
        Me.RadioButton35.TabStop = True
        Me.RadioButton35.Text = "Relay 35"
        Me.RadioButton35.UseVisualStyleBackColor = True
        '
        'RadioButton34
        '
        Me.RadioButton34.AutoSize = True
        Me.RadioButton34.Location = New System.Drawing.Point(128, 350)
        Me.RadioButton34.Name = "RadioButton34"
        Me.RadioButton34.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton34.TabIndex = 126
        Me.RadioButton34.TabStop = True
        Me.RadioButton34.Text = "Relay 34"
        Me.RadioButton34.UseVisualStyleBackColor = True
        '
        'RadioButton33
        '
        Me.RadioButton33.AutoSize = True
        Me.RadioButton33.Location = New System.Drawing.Point(128, 327)
        Me.RadioButton33.Name = "RadioButton33"
        Me.RadioButton33.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton33.TabIndex = 125
        Me.RadioButton33.TabStop = True
        Me.RadioButton33.Text = "Relay 33"
        Me.RadioButton33.UseVisualStyleBackColor = True
        '
        'RadioButton32
        '
        Me.RadioButton32.AutoSize = True
        Me.RadioButton32.Location = New System.Drawing.Point(128, 304)
        Me.RadioButton32.Name = "RadioButton32"
        Me.RadioButton32.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton32.TabIndex = 124
        Me.RadioButton32.TabStop = True
        Me.RadioButton32.Text = "Relay 32"
        Me.RadioButton32.UseVisualStyleBackColor = True
        '
        'RadioButton31
        '
        Me.RadioButton31.AutoSize = True
        Me.RadioButton31.Location = New System.Drawing.Point(128, 281)
        Me.RadioButton31.Name = "RadioButton31"
        Me.RadioButton31.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton31.TabIndex = 123
        Me.RadioButton31.TabStop = True
        Me.RadioButton31.Text = "Relay 31"
        Me.RadioButton31.UseVisualStyleBackColor = True
        '
        'RadioButton30
        '
        Me.RadioButton30.AutoSize = True
        Me.RadioButton30.Location = New System.Drawing.Point(128, 258)
        Me.RadioButton30.Name = "RadioButton30"
        Me.RadioButton30.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton30.TabIndex = 122
        Me.RadioButton30.TabStop = True
        Me.RadioButton30.Text = "Relay 30"
        Me.RadioButton30.UseVisualStyleBackColor = True
        '
        'RadioButton29
        '
        Me.RadioButton29.AutoSize = True
        Me.RadioButton29.Location = New System.Drawing.Point(128, 235)
        Me.RadioButton29.Name = "RadioButton29"
        Me.RadioButton29.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton29.TabIndex = 121
        Me.RadioButton29.TabStop = True
        Me.RadioButton29.Text = "Relay 29"
        Me.RadioButton29.UseVisualStyleBackColor = True
        '
        'RadioButton28
        '
        Me.RadioButton28.AutoSize = True
        Me.RadioButton28.Location = New System.Drawing.Point(128, 212)
        Me.RadioButton28.Name = "RadioButton28"
        Me.RadioButton28.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton28.TabIndex = 120
        Me.RadioButton28.TabStop = True
        Me.RadioButton28.Text = "Relay 28"
        Me.RadioButton28.UseVisualStyleBackColor = True
        '
        'RadioButton27
        '
        Me.RadioButton27.AutoSize = True
        Me.RadioButton27.Location = New System.Drawing.Point(128, 189)
        Me.RadioButton27.Name = "RadioButton27"
        Me.RadioButton27.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton27.TabIndex = 119
        Me.RadioButton27.TabStop = True
        Me.RadioButton27.Text = "Relay 27"
        Me.RadioButton27.UseVisualStyleBackColor = True
        '
        'RadioButton26
        '
        Me.RadioButton26.AutoSize = True
        Me.RadioButton26.Location = New System.Drawing.Point(128, 166)
        Me.RadioButton26.Name = "RadioButton26"
        Me.RadioButton26.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton26.TabIndex = 118
        Me.RadioButton26.TabStop = True
        Me.RadioButton26.Text = "Relay 26"
        Me.RadioButton26.UseVisualStyleBackColor = True
        '
        'RadioButton25
        '
        Me.RadioButton25.AutoSize = True
        Me.RadioButton25.Location = New System.Drawing.Point(128, 143)
        Me.RadioButton25.Name = "RadioButton25"
        Me.RadioButton25.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton25.TabIndex = 117
        Me.RadioButton25.TabStop = True
        Me.RadioButton25.Text = "Relay 25"
        Me.RadioButton25.UseVisualStyleBackColor = True
        '
        'RadioButton24
        '
        Me.RadioButton24.AutoSize = True
        Me.RadioButton24.Location = New System.Drawing.Point(128, 120)
        Me.RadioButton24.Name = "RadioButton24"
        Me.RadioButton24.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton24.TabIndex = 116
        Me.RadioButton24.TabStop = True
        Me.RadioButton24.Text = "Relay 24"
        Me.RadioButton24.UseVisualStyleBackColor = True
        '
        'RadioButton23
        '
        Me.RadioButton23.AutoSize = True
        Me.RadioButton23.Location = New System.Drawing.Point(128, 97)
        Me.RadioButton23.Name = "RadioButton23"
        Me.RadioButton23.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton23.TabIndex = 115
        Me.RadioButton23.TabStop = True
        Me.RadioButton23.Text = "Relay 23"
        Me.RadioButton23.UseVisualStyleBackColor = True
        '
        'RadioButton22
        '
        Me.RadioButton22.AutoSize = True
        Me.RadioButton22.Location = New System.Drawing.Point(128, 74)
        Me.RadioButton22.Name = "RadioButton22"
        Me.RadioButton22.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton22.TabIndex = 114
        Me.RadioButton22.TabStop = True
        Me.RadioButton22.Text = "Relay 22"
        Me.RadioButton22.UseVisualStyleBackColor = True
        '
        'RadioButton21
        '
        Me.RadioButton21.AutoSize = True
        Me.RadioButton21.Location = New System.Drawing.Point(128, 51)
        Me.RadioButton21.Name = "RadioButton21"
        Me.RadioButton21.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton21.TabIndex = 113
        Me.RadioButton21.TabStop = True
        Me.RadioButton21.Text = "Relay 21"
        Me.RadioButton21.UseVisualStyleBackColor = True
        '
        'RadioButton20
        '
        Me.RadioButton20.AutoSize = True
        Me.RadioButton20.Location = New System.Drawing.Point(8, 488)
        Me.RadioButton20.Name = "RadioButton20"
        Me.RadioButton20.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton20.TabIndex = 112
        Me.RadioButton20.TabStop = True
        Me.RadioButton20.Text = "Relay 20"
        Me.RadioButton20.UseVisualStyleBackColor = True
        '
        'RadioButton19
        '
        Me.RadioButton19.AutoSize = True
        Me.RadioButton19.Location = New System.Drawing.Point(8, 465)
        Me.RadioButton19.Name = "RadioButton19"
        Me.RadioButton19.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton19.TabIndex = 111
        Me.RadioButton19.TabStop = True
        Me.RadioButton19.Text = "Relay 19"
        Me.RadioButton19.UseVisualStyleBackColor = True
        '
        'RadioButton18
        '
        Me.RadioButton18.AutoSize = True
        Me.RadioButton18.Location = New System.Drawing.Point(7, 442)
        Me.RadioButton18.Name = "RadioButton18"
        Me.RadioButton18.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton18.TabIndex = 110
        Me.RadioButton18.TabStop = True
        Me.RadioButton18.Text = "Relay 18"
        Me.RadioButton18.UseVisualStyleBackColor = True
        '
        'RadioButton17
        '
        Me.RadioButton17.AutoSize = True
        Me.RadioButton17.Location = New System.Drawing.Point(8, 419)
        Me.RadioButton17.Name = "RadioButton17"
        Me.RadioButton17.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton17.TabIndex = 109
        Me.RadioButton17.TabStop = True
        Me.RadioButton17.Text = "Relay 17"
        Me.RadioButton17.UseVisualStyleBackColor = True
        '
        'RadioButton16
        '
        Me.RadioButton16.AutoSize = True
        Me.RadioButton16.Location = New System.Drawing.Point(8, 396)
        Me.RadioButton16.Name = "RadioButton16"
        Me.RadioButton16.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton16.TabIndex = 108
        Me.RadioButton16.TabStop = True
        Me.RadioButton16.Text = "Relay 16"
        Me.RadioButton16.UseVisualStyleBackColor = True
        '
        'RadioButton15
        '
        Me.RadioButton15.AutoSize = True
        Me.RadioButton15.Location = New System.Drawing.Point(8, 373)
        Me.RadioButton15.Name = "RadioButton15"
        Me.RadioButton15.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton15.TabIndex = 107
        Me.RadioButton15.TabStop = True
        Me.RadioButton15.Text = "Relay 15"
        Me.RadioButton15.UseVisualStyleBackColor = True
        '
        'RadioButton14
        '
        Me.RadioButton14.AutoSize = True
        Me.RadioButton14.Location = New System.Drawing.Point(7, 350)
        Me.RadioButton14.Name = "RadioButton14"
        Me.RadioButton14.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton14.TabIndex = 106
        Me.RadioButton14.TabStop = True
        Me.RadioButton14.Text = "Relay 14"
        Me.RadioButton14.UseVisualStyleBackColor = True
        '
        'RadioButton13
        '
        Me.RadioButton13.AutoSize = True
        Me.RadioButton13.Location = New System.Drawing.Point(8, 327)
        Me.RadioButton13.Name = "RadioButton13"
        Me.RadioButton13.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton13.TabIndex = 105
        Me.RadioButton13.TabStop = True
        Me.RadioButton13.Text = "Relay 13"
        Me.RadioButton13.UseVisualStyleBackColor = True
        '
        'RadioButton12
        '
        Me.RadioButton12.AutoSize = True
        Me.RadioButton12.Location = New System.Drawing.Point(7, 304)
        Me.RadioButton12.Name = "RadioButton12"
        Me.RadioButton12.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton12.TabIndex = 104
        Me.RadioButton12.TabStop = True
        Me.RadioButton12.Text = "Relay 12"
        Me.RadioButton12.UseVisualStyleBackColor = True
        '
        'RadioButton11
        '
        Me.RadioButton11.AutoSize = True
        Me.RadioButton11.Location = New System.Drawing.Point(8, 281)
        Me.RadioButton11.Name = "RadioButton11"
        Me.RadioButton11.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton11.TabIndex = 103
        Me.RadioButton11.TabStop = True
        Me.RadioButton11.Text = "Relay 11"
        Me.RadioButton11.UseVisualStyleBackColor = True
        '
        'RadioButton10
        '
        Me.RadioButton10.AutoSize = True
        Me.RadioButton10.Location = New System.Drawing.Point(7, 258)
        Me.RadioButton10.Name = "RadioButton10"
        Me.RadioButton10.Size = New System.Drawing.Size(81, 19)
        Me.RadioButton10.TabIndex = 102
        Me.RadioButton10.TabStop = True
        Me.RadioButton10.Text = "Relay 10"
        Me.RadioButton10.UseVisualStyleBackColor = True
        '
        'RadioButton9
        '
        Me.RadioButton9.AutoSize = True
        Me.RadioButton9.Location = New System.Drawing.Point(8, 235)
        Me.RadioButton9.Name = "RadioButton9"
        Me.RadioButton9.Size = New System.Drawing.Size(73, 19)
        Me.RadioButton9.TabIndex = 101
        Me.RadioButton9.TabStop = True
        Me.RadioButton9.Text = "Relay 9"
        Me.RadioButton9.UseVisualStyleBackColor = True
        '
        'RadioButton8
        '
        Me.RadioButton8.AutoSize = True
        Me.RadioButton8.Location = New System.Drawing.Point(7, 212)
        Me.RadioButton8.Name = "RadioButton8"
        Me.RadioButton8.Size = New System.Drawing.Size(73, 19)
        Me.RadioButton8.TabIndex = 100
        Me.RadioButton8.TabStop = True
        Me.RadioButton8.Text = "Relay 8"
        Me.RadioButton8.UseVisualStyleBackColor = True
        '
        'RadioButton7
        '
        Me.RadioButton7.AutoSize = True
        Me.RadioButton7.Location = New System.Drawing.Point(8, 190)
        Me.RadioButton7.Name = "RadioButton7"
        Me.RadioButton7.Size = New System.Drawing.Size(73, 19)
        Me.RadioButton7.TabIndex = 99
        Me.RadioButton7.TabStop = True
        Me.RadioButton7.Text = "Relay 7"
        Me.RadioButton7.UseVisualStyleBackColor = True
        '
        'RadioButton6
        '
        Me.RadioButton6.AutoSize = True
        Me.RadioButton6.Location = New System.Drawing.Point(7, 166)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.Size = New System.Drawing.Size(73, 19)
        Me.RadioButton6.TabIndex = 98
        Me.RadioButton6.TabStop = True
        Me.RadioButton6.Text = "Relay 6"
        Me.RadioButton6.UseVisualStyleBackColor = True
        '
        'RadioButton5
        '
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Location = New System.Drawing.Point(8, 143)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(73, 19)
        Me.RadioButton5.TabIndex = 97
        Me.RadioButton5.TabStop = True
        Me.RadioButton5.Text = "Relay 5"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Location = New System.Drawing.Point(8, 120)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(73, 19)
        Me.RadioButton4.TabIndex = 96
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.Text = "Relay 4"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(8, 97)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(73, 19)
        Me.RadioButton3.TabIndex = 95
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "Relay 3"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(7, 74)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(73, 19)
        Me.RadioButton2.TabIndex = 94
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Relay 2"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(7, 51)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton1.TabIndex = 93
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.CheckBox40)
        Me.GroupBox3.Controls.Add(Me.CheckBox39)
        Me.GroupBox3.Controls.Add(Me.CheckBox38)
        Me.GroupBox3.Controls.Add(Me.CheckBox37)
        Me.GroupBox3.Controls.Add(Me.CheckBox36)
        Me.GroupBox3.Controls.Add(Me.CheckBox35)
        Me.GroupBox3.Controls.Add(Me.CheckBox34)
        Me.GroupBox3.Controls.Add(Me.CheckBox33)
        Me.GroupBox3.Controls.Add(Me.CheckBox32)
        Me.GroupBox3.Controls.Add(Me.CheckBox31)
        Me.GroupBox3.Controls.Add(Me.CheckBox30)
        Me.GroupBox3.Controls.Add(Me.CheckBox29)
        Me.GroupBox3.Controls.Add(Me.CheckBox28)
        Me.GroupBox3.Controls.Add(Me.CheckBox27)
        Me.GroupBox3.Controls.Add(Me.CheckBox26)
        Me.GroupBox3.Controls.Add(Me.CheckBox25)
        Me.GroupBox3.Controls.Add(Me.CheckBox24)
        Me.GroupBox3.Controls.Add(Me.CheckBox23)
        Me.GroupBox3.Controls.Add(Me.CheckBox22)
        Me.GroupBox3.Controls.Add(Me.CheckBox21)
        Me.GroupBox3.Controls.Add(Me.CheckBox20)
        Me.GroupBox3.Controls.Add(Me.CheckBox19)
        Me.GroupBox3.Controls.Add(Me.CheckBox18)
        Me.GroupBox3.Controls.Add(Me.CheckBox17)
        Me.GroupBox3.Controls.Add(Me.CheckBox16)
        Me.GroupBox3.Controls.Add(Me.CheckBox15)
        Me.GroupBox3.Controls.Add(Me.CheckBox14)
        Me.GroupBox3.Controls.Add(Me.CheckBox13)
        Me.GroupBox3.Controls.Add(Me.CheckBox12)
        Me.GroupBox3.Controls.Add(Me.CheckBox11)
        Me.GroupBox3.Controls.Add(Me.CheckBox10)
        Me.GroupBox3.Controls.Add(Me.CheckBox9)
        Me.GroupBox3.Controls.Add(Me.CheckBox8)
        Me.GroupBox3.Controls.Add(Me.CheckBox7)
        Me.GroupBox3.Controls.Add(Me.CheckBox6)
        Me.GroupBox3.Controls.Add(Me.CheckBox5)
        Me.GroupBox3.Controls.Add(Me.CheckBox4)
        Me.GroupBox3.Controls.Add(Me.CheckBox3)
        Me.GroupBox3.Controls.Add(Me.CheckBox2)
        Me.GroupBox3.Controls.Add(Me.CheckBox1)
        Me.GroupBox3.Location = New System.Drawing.Point(14, 40)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(240, 530)
        Me.GroupBox3.TabIndex = 213
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Manual Relay Check Positions"
        '
        'CheckBox40
        '
        Me.CheckBox40.AutoSize = True
        Me.CheckBox40.Location = New System.Drawing.Point(132, 489)
        Me.CheckBox40.Name = "CheckBox40"
        Me.CheckBox40.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox40.TabIndex = 40
        Me.CheckBox40.Text = "Relay 40"
        Me.CheckBox40.UseVisualStyleBackColor = True
        '
        'CheckBox39
        '
        Me.CheckBox39.AutoSize = True
        Me.CheckBox39.Location = New System.Drawing.Point(132, 466)
        Me.CheckBox39.Name = "CheckBox39"
        Me.CheckBox39.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox39.TabIndex = 39
        Me.CheckBox39.Text = "Relay 39"
        Me.CheckBox39.UseVisualStyleBackColor = True
        '
        'CheckBox38
        '
        Me.CheckBox38.AutoSize = True
        Me.CheckBox38.Location = New System.Drawing.Point(132, 443)
        Me.CheckBox38.Name = "CheckBox38"
        Me.CheckBox38.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox38.TabIndex = 38
        Me.CheckBox38.Text = "Relay 38"
        Me.CheckBox38.UseVisualStyleBackColor = True
        '
        'CheckBox37
        '
        Me.CheckBox37.AutoSize = True
        Me.CheckBox37.Location = New System.Drawing.Point(132, 420)
        Me.CheckBox37.Name = "CheckBox37"
        Me.CheckBox37.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox37.TabIndex = 37
        Me.CheckBox37.Text = "Relay 37"
        Me.CheckBox37.UseVisualStyleBackColor = True
        '
        'CheckBox36
        '
        Me.CheckBox36.AutoSize = True
        Me.CheckBox36.Location = New System.Drawing.Point(132, 397)
        Me.CheckBox36.Name = "CheckBox36"
        Me.CheckBox36.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox36.TabIndex = 36
        Me.CheckBox36.Text = "Relay 36"
        Me.CheckBox36.UseVisualStyleBackColor = True
        '
        'CheckBox35
        '
        Me.CheckBox35.AutoSize = True
        Me.CheckBox35.Location = New System.Drawing.Point(132, 374)
        Me.CheckBox35.Name = "CheckBox35"
        Me.CheckBox35.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox35.TabIndex = 35
        Me.CheckBox35.Text = "Relay 35"
        Me.CheckBox35.UseVisualStyleBackColor = True
        '
        'CheckBox34
        '
        Me.CheckBox34.AutoSize = True
        Me.CheckBox34.Location = New System.Drawing.Point(132, 353)
        Me.CheckBox34.Name = "CheckBox34"
        Me.CheckBox34.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox34.TabIndex = 34
        Me.CheckBox34.Text = "Relay 34"
        Me.CheckBox34.UseVisualStyleBackColor = True
        '
        'CheckBox33
        '
        Me.CheckBox33.AutoSize = True
        Me.CheckBox33.Location = New System.Drawing.Point(132, 328)
        Me.CheckBox33.Name = "CheckBox33"
        Me.CheckBox33.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox33.TabIndex = 33
        Me.CheckBox33.Text = "Relay 33"
        Me.CheckBox33.UseVisualStyleBackColor = True
        '
        'CheckBox32
        '
        Me.CheckBox32.AutoSize = True
        Me.CheckBox32.Location = New System.Drawing.Point(132, 305)
        Me.CheckBox32.Name = "CheckBox32"
        Me.CheckBox32.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox32.TabIndex = 32
        Me.CheckBox32.Text = "Relay 32"
        Me.CheckBox32.UseVisualStyleBackColor = True
        '
        'CheckBox31
        '
        Me.CheckBox31.AutoSize = True
        Me.CheckBox31.Location = New System.Drawing.Point(132, 282)
        Me.CheckBox31.Name = "CheckBox31"
        Me.CheckBox31.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox31.TabIndex = 31
        Me.CheckBox31.Text = "Relay 31"
        Me.CheckBox31.UseVisualStyleBackColor = True
        '
        'CheckBox30
        '
        Me.CheckBox30.AutoSize = True
        Me.CheckBox30.Location = New System.Drawing.Point(132, 259)
        Me.CheckBox30.Name = "CheckBox30"
        Me.CheckBox30.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox30.TabIndex = 30
        Me.CheckBox30.Text = "Relay 30"
        Me.CheckBox30.UseVisualStyleBackColor = True
        '
        'CheckBox29
        '
        Me.CheckBox29.AutoSize = True
        Me.CheckBox29.Location = New System.Drawing.Point(132, 236)
        Me.CheckBox29.Name = "CheckBox29"
        Me.CheckBox29.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox29.TabIndex = 29
        Me.CheckBox29.Text = "Relay 29"
        Me.CheckBox29.UseVisualStyleBackColor = True
        '
        'CheckBox28
        '
        Me.CheckBox28.AutoSize = True
        Me.CheckBox28.Location = New System.Drawing.Point(132, 213)
        Me.CheckBox28.Name = "CheckBox28"
        Me.CheckBox28.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox28.TabIndex = 28
        Me.CheckBox28.Text = "Relay 28"
        Me.CheckBox28.UseVisualStyleBackColor = True
        '
        'CheckBox27
        '
        Me.CheckBox27.AutoSize = True
        Me.CheckBox27.Location = New System.Drawing.Point(132, 191)
        Me.CheckBox27.Name = "CheckBox27"
        Me.CheckBox27.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox27.TabIndex = 27
        Me.CheckBox27.Text = "Relay 27"
        Me.CheckBox27.UseVisualStyleBackColor = True
        '
        'CheckBox26
        '
        Me.CheckBox26.AutoSize = True
        Me.CheckBox26.Location = New System.Drawing.Point(132, 165)
        Me.CheckBox26.Name = "CheckBox26"
        Me.CheckBox26.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox26.TabIndex = 26
        Me.CheckBox26.Text = "Relay 26"
        Me.CheckBox26.UseVisualStyleBackColor = True
        '
        'CheckBox25
        '
        Me.CheckBox25.AutoSize = True
        Me.CheckBox25.Location = New System.Drawing.Point(132, 144)
        Me.CheckBox25.Name = "CheckBox25"
        Me.CheckBox25.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox25.TabIndex = 25
        Me.CheckBox25.Text = "Relay 25"
        Me.CheckBox25.UseVisualStyleBackColor = True
        '
        'CheckBox24
        '
        Me.CheckBox24.AutoSize = True
        Me.CheckBox24.Location = New System.Drawing.Point(132, 121)
        Me.CheckBox24.Name = "CheckBox24"
        Me.CheckBox24.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox24.TabIndex = 24
        Me.CheckBox24.Text = "Relay 24"
        Me.CheckBox24.UseVisualStyleBackColor = True
        '
        'CheckBox23
        '
        Me.CheckBox23.AutoSize = True
        Me.CheckBox23.Location = New System.Drawing.Point(132, 96)
        Me.CheckBox23.Name = "CheckBox23"
        Me.CheckBox23.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox23.TabIndex = 23
        Me.CheckBox23.Text = "Relay 23"
        Me.CheckBox23.UseVisualStyleBackColor = True
        '
        'CheckBox22
        '
        Me.CheckBox22.AutoSize = True
        Me.CheckBox22.Location = New System.Drawing.Point(132, 75)
        Me.CheckBox22.Name = "CheckBox22"
        Me.CheckBox22.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox22.TabIndex = 22
        Me.CheckBox22.Text = "Relay 22"
        Me.CheckBox22.UseVisualStyleBackColor = True
        '
        'CheckBox21
        '
        Me.CheckBox21.AutoSize = True
        Me.CheckBox21.Location = New System.Drawing.Point(132, 52)
        Me.CheckBox21.Name = "CheckBox21"
        Me.CheckBox21.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox21.TabIndex = 21
        Me.CheckBox21.Text = "Relay 21"
        Me.CheckBox21.UseVisualStyleBackColor = True
        '
        'CheckBox20
        '
        Me.CheckBox20.AutoSize = True
        Me.CheckBox20.Location = New System.Drawing.Point(16, 489)
        Me.CheckBox20.Name = "CheckBox20"
        Me.CheckBox20.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox20.TabIndex = 20
        Me.CheckBox20.Text = "Relay 20"
        Me.CheckBox20.UseVisualStyleBackColor = True
        '
        'CheckBox19
        '
        Me.CheckBox19.AutoSize = True
        Me.CheckBox19.Location = New System.Drawing.Point(16, 466)
        Me.CheckBox19.Name = "CheckBox19"
        Me.CheckBox19.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox19.TabIndex = 19
        Me.CheckBox19.Text = "Relay 19"
        Me.CheckBox19.UseVisualStyleBackColor = True
        '
        'CheckBox18
        '
        Me.CheckBox18.AutoSize = True
        Me.CheckBox18.Location = New System.Drawing.Point(16, 443)
        Me.CheckBox18.Name = "CheckBox18"
        Me.CheckBox18.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox18.TabIndex = 18
        Me.CheckBox18.Text = "Relay 18"
        Me.CheckBox18.UseVisualStyleBackColor = True
        '
        'CheckBox17
        '
        Me.CheckBox17.AutoSize = True
        Me.CheckBox17.Location = New System.Drawing.Point(16, 420)
        Me.CheckBox17.Name = "CheckBox17"
        Me.CheckBox17.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox17.TabIndex = 17
        Me.CheckBox17.Text = "Relay 17"
        Me.CheckBox17.UseVisualStyleBackColor = True
        '
        'CheckBox16
        '
        Me.CheckBox16.AutoSize = True
        Me.CheckBox16.Location = New System.Drawing.Point(16, 397)
        Me.CheckBox16.Name = "CheckBox16"
        Me.CheckBox16.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox16.TabIndex = 16
        Me.CheckBox16.Text = "Relay 16"
        Me.CheckBox16.UseVisualStyleBackColor = True
        '
        'CheckBox15
        '
        Me.CheckBox15.AutoSize = True
        Me.CheckBox15.Location = New System.Drawing.Point(16, 374)
        Me.CheckBox15.Name = "CheckBox15"
        Me.CheckBox15.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox15.TabIndex = 15
        Me.CheckBox15.Text = "Relay 15"
        Me.CheckBox15.UseVisualStyleBackColor = True
        '
        'CheckBox14
        '
        Me.CheckBox14.AutoSize = True
        Me.CheckBox14.Location = New System.Drawing.Point(16, 349)
        Me.CheckBox14.Name = "CheckBox14"
        Me.CheckBox14.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox14.TabIndex = 14
        Me.CheckBox14.Text = "Relay 14"
        Me.CheckBox14.UseVisualStyleBackColor = True
        '
        'CheckBox13
        '
        Me.CheckBox13.AutoSize = True
        Me.CheckBox13.Location = New System.Drawing.Point(16, 328)
        Me.CheckBox13.Name = "CheckBox13"
        Me.CheckBox13.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox13.TabIndex = 13
        Me.CheckBox13.Text = "Relay 13"
        Me.CheckBox13.UseVisualStyleBackColor = True
        '
        'CheckBox12
        '
        Me.CheckBox12.AutoSize = True
        Me.CheckBox12.Location = New System.Drawing.Point(16, 305)
        Me.CheckBox12.Name = "CheckBox12"
        Me.CheckBox12.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox12.TabIndex = 12
        Me.CheckBox12.Text = "Relay 12"
        Me.CheckBox12.UseVisualStyleBackColor = True
        '
        'CheckBox11
        '
        Me.CheckBox11.AutoSize = True
        Me.CheckBox11.Location = New System.Drawing.Point(16, 282)
        Me.CheckBox11.Name = "CheckBox11"
        Me.CheckBox11.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox11.TabIndex = 11
        Me.CheckBox11.Text = "Relay 11"
        Me.CheckBox11.UseVisualStyleBackColor = True
        '
        'CheckBox10
        '
        Me.CheckBox10.AutoSize = True
        Me.CheckBox10.Location = New System.Drawing.Point(16, 259)
        Me.CheckBox10.Name = "CheckBox10"
        Me.CheckBox10.Size = New System.Drawing.Size(68, 17)
        Me.CheckBox10.TabIndex = 10
        Me.CheckBox10.Text = "Relay 10"
        Me.CheckBox10.UseVisualStyleBackColor = True
        '
        'CheckBox9
        '
        Me.CheckBox9.AutoSize = True
        Me.CheckBox9.Location = New System.Drawing.Point(16, 236)
        Me.CheckBox9.Name = "CheckBox9"
        Me.CheckBox9.Size = New System.Drawing.Size(62, 17)
        Me.CheckBox9.TabIndex = 9
        Me.CheckBox9.Text = "Relay 9"
        Me.CheckBox9.UseVisualStyleBackColor = True
        '
        'CheckBox8
        '
        Me.CheckBox8.AutoSize = True
        Me.CheckBox8.Location = New System.Drawing.Point(16, 213)
        Me.CheckBox8.Name = "CheckBox8"
        Me.CheckBox8.Size = New System.Drawing.Size(62, 17)
        Me.CheckBox8.TabIndex = 8
        Me.CheckBox8.Text = "Relay 8"
        Me.CheckBox8.UseVisualStyleBackColor = True
        '
        'CheckBox7
        '
        Me.CheckBox7.AutoSize = True
        Me.CheckBox7.Location = New System.Drawing.Point(16, 190)
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.Size = New System.Drawing.Size(62, 17)
        Me.CheckBox7.TabIndex = 7
        Me.CheckBox7.Text = "Relay 7"
        Me.CheckBox7.UseVisualStyleBackColor = True
        '
        'CheckBox6
        '
        Me.CheckBox6.AutoSize = True
        Me.CheckBox6.Location = New System.Drawing.Point(16, 167)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(62, 17)
        Me.CheckBox6.TabIndex = 6
        Me.CheckBox6.Text = "Relay 6"
        Me.CheckBox6.UseVisualStyleBackColor = True
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Location = New System.Drawing.Point(16, 144)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(62, 17)
        Me.CheckBox5.TabIndex = 5
        Me.CheckBox5.Text = "Relay 5"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Location = New System.Drawing.Point(16, 121)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(62, 17)
        Me.CheckBox4.TabIndex = 4
        Me.CheckBox4.Text = "Relay 4"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(16, 98)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(62, 17)
        Me.CheckBox3.TabIndex = 3
        Me.CheckBox3.Text = "Relay 3"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(16, 75)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(62, 17)
        Me.CheckBox2.TabIndex = 2
        Me.CheckBox2.Text = "Relay 2"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(16, 52)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(15, 14)
        Me.CheckBox1.TabIndex = 1
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'tbDoc4
        '
        Me.tbDoc4.Location = New System.Drawing.Point(4, 22)
        Me.tbDoc4.Name = "tbDoc4"
        Me.tbDoc4.Size = New System.Drawing.Size(564, 451)
        Me.tbDoc4.TabIndex = 0
        Me.tbDoc4.Tag = 4
        Me.tbDoc4.Text = "Document4"
        '
        'axFramer4
        '
        Me.axFramer4.Enabled = True
        Me.axFramer4.Location = New System.Drawing.Point(0, 0)
        Me.axFramer4.Name = "axFramer4"
        Me.axFramer4.TabIndex = 0
        '
        'tbDoc1
        '
        Me.tbDoc1.Location = New System.Drawing.Point(4, 22)
        Me.tbDoc1.Name = "tbDoc1"
        Me.tbDoc1.Size = New System.Drawing.Size(564, 451)
        Me.tbDoc1.TabIndex = 0
        Me.tbDoc1.Tag = 1
        Me.tbDoc1.Text = "Document1"
        '
        'axFramer1
        '
        Me.axFramer1.Enabled = True
        Me.axFramer1.Location = New System.Drawing.Point(0, 0)
        Me.axFramer1.Name = "axFramer1"
        Me.axFramer1.TabIndex = 0
        '
        'tbDoc2
        '
        Me.tbDoc2.Location = New System.Drawing.Point(4, 22)
        Me.tbDoc2.Name = "tbDoc2"
        Me.tbDoc2.Size = New System.Drawing.Size(564, 451)
        Me.tbDoc2.TabIndex = 0
        Me.tbDoc2.Tag = 2
        Me.tbDoc2.Text = "Document2"
        '
        'axFramer2
        '
        Me.axFramer2.Enabled = True
        Me.axFramer2.Location = New System.Drawing.Point(0, 0)
        Me.axFramer2.Name = "axFramer2"
        Me.axFramer2.TabIndex = 0
        '
        'tbDoc3
        '
        Me.tbDoc3.Location = New System.Drawing.Point(4, 22)
        Me.tbDoc3.Name = "tbDoc3"
        Me.tbDoc3.Size = New System.Drawing.Size(564, 451)
        Me.tbDoc3.TabIndex = 0
        Me.tbDoc3.Tag = 3
        Me.tbDoc3.Text = "Document3"
        '
        'axFramer3
        '
        Me.axFramer3.Enabled = True
        Me.axFramer3.Location = New System.Drawing.Point(0, 0)
        Me.axFramer3.Name = "axFramer3"
        Me.axFramer3.TabIndex = 0
        '
        'OFileDialog
        '
        Me.OFileDialog.Filter = "Microsoft Office Files|*.doc;*.docx;*.docm;*.xls;*.xlsx;*.xlsm;*.xlsb;*.ppt;*.ppt" & _
    "x;*.pptm|All Files|*.*"
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        Me.BackgroundWorker2.WorkerSupportsCancellation = True
        '
        'AxInstantDoCtrl4
        '
        Me.AxInstantDoCtrl4.Enabled = True
        Me.AxInstantDoCtrl4.Location = New System.Drawing.Point(1192, 4)
        Me.AxInstantDoCtrl4.Name = "AxInstantDoCtrl4"
        Me.AxInstantDoCtrl4.OcxState = CType(resources.GetObject("AxInstantDoCtrl4.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxInstantDoCtrl4.Size = New System.Drawing.Size(17, 17)
        Me.AxInstantDoCtrl4.TabIndex = 214
        '
        'FTestApp
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1486, 998)
        Me.Controls.Add(Me.tbcDocsCont)
        Me.Controls.Add(Me.picLogo)
        Me.MinimumSize = New System.Drawing.Size(580, 500)
        Me.Name = "FTestApp"
        Me.Text = "MC Test 39 Taps"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbcDocsCont.ResumeLayout(False)
        Me.tbMain.ResumeLayout(False)
        Me.tbMain.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxAdvDIO2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxAdvDIO1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.axFramer4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.axFramer1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.axFramer2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.axFramer3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxInstantDoCtrl4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Dim m_iCounter As Integer
    'Dim m_bFilesOpen(4) As Boolean
    'Dim TriggerPaused As Boolean
    'Dim wh As New System.Threading.EventWaitHandle(False, Threading.EventResetMode.AutoReset)
    '' ============================================================================
    ''  GetOpenSlot - Helper Function
    ''
    ''   Returns free slot for tab page and framer control. You could implement a
    ''   a dynamic control array and support up to the max number of framer controls,
    ''   but we are keeping this simple and limiting to just 4 open docs at a time.
    ''
    '' ============================================================================
    'Private Function GetOpenSlot() As Integer
    '    Dim i As Integer
    '    For i = 0 To 3
    '        If m_bFilesOpen(i) = False Then
    '            Return i + 1
    '        End If
    '    Next
    '    Return 0
    'End Function

    '' ============================================================================
    ''  GetTabPageFromIdx - Helper Function
    ''
    ''   Returns tab page control for the given slot.
    ''
    '' ============================================================================
    'Private Function GetTabPageFromIdx(ByVal idx As Integer) As TabPage
    '    Select Case idx
    '        Case 1
    '            Return Me.tbDoc1
    '        Case 2
    '            Return Me.tbDoc2
    '        Case 3
    '            Return Me.tbDoc3
    '        Case 4
    '            Return Me.tbDoc4
    '        Case Else
    '            Throw New Exception("Invalid Index")
    '    End Select
    'End Function

    '' ============================================================================
    ''  GetFramerCtlFromIdx - Helper Function
    ''
    ''   Returns DsoFramer control for the given slot.
    ''
    '' ============================================================================
    'Private Function GetFramerCtlFromIdx(ByVal idx As Integer) As AxDSOFramer.AxFramerControl
    '    Select Case idx
    '        Case 1
    '            Return Me.axFramer1
    '        Case 2
    '            Return Me.axFramer2
    '        Case 3
    '            Return Me.axFramer3
    '        Case 4
    '            Return Me.axFramer4
    '        Case Else
    '            Throw New Exception("Invalid Index")
    '    End Select
    'End Function

    '' ============================================================================
    ''  AddTabAndActivate - Helper Subroutine
    ''
    ''   Adds tab page to the TabControl on our main form, and sets a default name
    ''   for the tab (assuming this is new document). Then selects the tab to activate.
    ''
    '' ============================================================================
    'Private Sub AddTabAndActivate(ByVal idx As Integer)
    '    Dim tab As TabPage
    '    Dim axControl As AxDSOFramer.AxFramerControl
    '    m_iCounter = m_iCounter + 1

    '    ' Get the tab control and add it to the collection...
    '    tab = GetTabPageFromIdx(idx)
    '    tab.Text = "New Document " & m_iCounter
    '    Me.tbcDocsCont.Controls.Add(tab)
    '    Me.tbcDocsCont.SelectedTab = tab

    '    ' Get the Framer control and set some default properties since
    '    ' we don't want user to open/new without going to our main tab.
    '    axControl = GetFramerCtlFromIdx(idx)
    '    axControl.set_EnableFileCommand(DSOFramer.dsoFileCommandType.dsoFileNew, False)
    '    axControl.set_EnableFileCommand(DSOFramer.dsoFileCommandType.dsoFileOpen, False)
    '    ' We need to explicitly enable the event sinks. Due to a strange bug in .NET
    '    ' when the control is sited to tab page and not the main form, it is told to 
    '    ' freeze events (IOleControl) but never told to unfreeze. So events don't get
    '    ' fired correctly from tab strip. This sets the flag to re-enable the events.
    '    ' axControl.EventsEnabled = True
    '    axControl.Select()

    '    ' Since we just activated, mark this slot as occupied...
    '    m_bFilesOpen(idx - 1) = True
    'End Sub

    '' ============================================================================
    ''  RemoveTabAndSelectMain - Helper Subroutine
    ''
    ''   Removes the tab from the collection when the document is closed by user.
    ''
    '' ============================================================================
    'Private Sub RemoveTabAndSelectMain(ByVal idx As Integer)
    '    Me.tbcDocsCont.Controls.Remove(GetTabPageFromIdx(idx))
    '    Me.tbcDocsCont.SelectedIndex = 1
    '    m_bFilesOpen(idx - 1) = False
    'End Sub


    '' ============================================================================
    ''  btnOpenFile_Click - File Open Button Click Handler
    ''
    ''   Opens the file in DsoFramer control and appends it to a new tab in the
    ''   tab strip control. We allow user to pick the file using File Open dialog.
    ''
    '' ============================================================================
    'Private Sub btnOpenFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenFile.Click
    '    ' Ensure we have a free slot...
    '    Dim idx As Integer = GetOpenSlot()
    '    If idx = 0 Then
    '        MessageBox.Show("You can only have four documents open at a time. Close will need to close one to continue.", "Open File", _
    '            MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '        Exit Sub
    '    End If

    '    ' Temporarily disable the buttons so we don't re-enter
    '    'btnCreateNew.Enabled = False
    '    'btnOpenFile.Enabled = False

    '    ' Ask user for the file to open...
    '    Dim r As DialogResult = OFileDialog.ShowDialog()
    '    If r = DialogResult.OK Then

    '        ' Add the tab to the collection and switch to the tab...
    '        AddTabAndActivate(idx)

    '        ' Get the framer control for that slot...
    '        Dim ctl As AxDSOFramer.AxFramerControl
    '        ctl = GetFramerCtlFromIdx(idx)
    '        Try
    '            ' Ask it to open the file...
    '            ctl.Open(OFileDialog.FileName)

    '            ' Get the tab page and change the title to the name of the file opened...
    '            Dim tp As TabPage
    '            tp = GetTabPageFromIdx(idx)
    '            tp.Text = ctl.DocumentName

    '        Catch ex As Exception
    '            ' Show the error to user...
    '            MessageBox.Show("Unable to open the file. " & ex.Message, "File Open", _
    '                MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            ' If we fail, remove the tab and take us back to the main page
    '            RemoveTabAndSelectMain(idx)
    '        End Try
    '    End If

    '    'btnCreateNew.Enabled = True
    '    ' btnOpenFile.Enabled = True
    'End Sub

    ' ============================================================================
    '  btnCreateNew_Click - Create New Button Click Handler
    '
    '   Creates new blank document of one of the types selected in Radio buttons.
    '
    ' ============================================================================
    'Private Sub btnCreateNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    ' Ensure we have a free slot...
    '    Dim idx As Integer = GetOpenSlot()
    '    If idx = 0 Then
    '        MessageBox.Show("You can only have four documents open at a time. Close will need to close one to continue.", "Create New", _
    '            MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '        Exit Sub
    '    End If

    '    ' Temporarily disable the buttons so we don't re-enter
    '    ' btnCreateNew.Enabled = False
    '    ' btnOpenFile.Enabled = False

    '    ' Pick the ProgID from the Radio button selected...
    '    'Dim sProgID As String
    '    'If rbtnNewWord.Checked Then
    '    '    sProgID = "Word.Document.8"
    '    'ElseIf rbtnNewExcel.Checked Then
    '    '    sProgID = "Excel.Sheet.8"
    '    'Else
    '    '    sProgID = "PowerPoint.Show.8"
    '    'End If

    '    ' Add the tab page to tab and make it visible...
    '    AddTabAndActivate(idx)

    '    ' Get the framer control for that free slot...
    '    Dim ctl As AxDSOFramer.AxFramerControl
    '    ctl = GetFramerCtlFromIdx(idx)

    '    Try
    '        ' Ask it to create the new object...
    '        'ctl.CreateNew(sProgID)

    '    Catch ex As Exception
    '        ' Show error to user...
    '        MessageBox.Show("Unable to create the new document. " & ex.Message, "Create New", _
    '            MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        ' If we fail, remove the tab and take us back to the main page
    '        RemoveTabAndSelectMain(idx)
    '    End Try

    '    'btnCreateNew.Enabled = True
    '    ' btnOpenFile.Enabled = True
    'End Sub

    '' ============================================================================
    ''  tbcDocsCont_SelectedIndexChanged 
    ''
    ''   When switching tabs, activate the framer control associated with that tab.
    ''
    '' ============================================================================
    'Private Sub tbcDocsCont_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbcDocsCont.SelectedIndexChanged
    '    Dim tab As TabPage = tbcDocsCont.SelectedTab
    '    Dim framer As AxDSOFramer.AxFramerControl
    '    Dim idx As Integer = tab.Tag
    '    If (idx >= 1 And idx <= 4) Then
    '        framer = GetFramerCtlFromIdx(idx)
    '        framer.Activate()
    '    End If
    'End Sub

    '' ============================================================================
    ''  axFramerX_OnDocumentClosed 
    ''
    ''  Control event handlers to remove the tab when the document(s) are closed.
    ''
    '' ============================================================================
    'Private Sub axFramer1_OnDocumentClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles axFramer1.OnDocumentClosed
    '    RemoveTabAndSelectMain(1)
    'End Sub

    'Private Sub axFramer2_OnDocumentClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles axFramer2.OnDocumentClosed
    '    RemoveTabAndSelectMain(2)
    'End Sub

    'Private Sub axFramer3_OnDocumentClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles axFramer3.OnDocumentClosed
    '    RemoveTabAndSelectMain(3)
    'End Sub

    'Private Sub axFramer4_OnDocumentClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles axFramer4.OnDocumentClosed
    '    RemoveTabAndSelectMain(4)
    'End Sub

    '' ============================================================================
    ''  axFramerX_OnSaveCompleted 
    ''
    ''  Control event handlers to rename tabs if document is saved by a new name.
    ''
    '' ============================================================================
    'Private Sub axFramer1_OnSaveCompleted(ByVal sender As Object, ByVal e As AxDSOFramer._DFramerCtlEvents_OnSaveCompletedEvent) Handles axFramer1.OnSaveCompleted
    '    Dim s As String = e.docName
    '    If s.Length > 1 Then tbDoc1.Text = e.docName
    'End Sub

    'Private Sub axFramer2_OnSaveCompleted(ByVal sender As Object, ByVal e As AxDSOFramer._DFramerCtlEvents_OnSaveCompletedEvent) Handles axFramer2.OnSaveCompleted
    '    Dim s As String = e.docName
    '    If s.Length > 1 Then tbDoc2.Text = e.docName
    'End Sub

    'Private Sub axFramer3_OnSaveCompleted(ByVal sender As Object, ByVal e As AxDSOFramer._DFramerCtlEvents_OnSaveCompletedEvent) Handles axFramer3.OnSaveCompleted
    '    Dim s As String = e.docName
    '    If s.Length > 1 Then tbDoc3.Text = e.docName
    'End Sub

    'Private Sub axFramer4_OnSaveCompleted(ByVal sender As Object, ByVal e As AxDSOFramer._DFramerCtlEvents_OnSaveCompletedEvent) Handles axFramer4.OnSaveCompleted
    '    Dim s As String = e.docName
    '    If s.Length > 1 Then tbDoc4.Text = e.docName
    'End Sub

    Public Sub SETRELAYS_Click(sender As Object, e As EventArgs) Handles SETRELAYS.Click
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)
        Dim IntCount(256) As Long
        Dim checki As Integer = 0
        Dim checkx As Integer = 0
        Dim hold As String = ""


        stopflag = 0

        Call Find_check(0, e)
        If Button13.BackColor = Color.Green Then

            Me.Label36.Text = "DAILY ACCURACY TEST  "
            Me.Label36.Font = New Drawing.Font("Times New Roman", 20, FontStyle.Bold)
            Me.Label36.BackColor = Color.Red



            Dim MyObject As Control

            For Each MyObject In Me.GroupBox3.Controls

                If TypeOf MyObject Is CheckBox Then

                    MyObject.Enabled = False

                End If

            Next


            Call Find_check2(0, e)
            If checkboxarray.Length = 1 Then
                'MsgBox("Please Select Taps and Start")
                For Each MyObject In Me.GroupBox3.Controls

                    If TypeOf MyObject Is CheckBox Then

                        MyObject.Enabled = True

                    End If

                Next





                Exit Sub
            End If


            checki = checkboxarray.Length
            For checkx = 1 To checki - 1
                hold = Mid(checkboxarray(checkx), 9, 2)
                CBNUM = "RadioButton" & hold

                If stopflag = 1 Then
                    Exit Sub
                End If

                Select Case hold

                    Case "40"
                        CheckBox40.BackColor = Color.Green
                    Case "39"
                        CheckBox39.BackColor = Color.Green
                    Case "38"
                        CheckBox38.BackColor = Color.Green
                    Case "37"
                        CheckBox37.BackColor = Color.Green
                    Case "36"
                        CheckBox36.BackColor = Color.Green
                    Case "35"
                        CheckBox35.BackColor = Color.Green
                    Case "34"
                        CheckBox34.BackColor = Color.Green
                    Case "33"
                        CheckBox33.BackColor = Color.Green
                    Case "32"
                        CheckBox32.BackColor = Color.Green
                    Case "31"
                        CheckBox31.BackColor = Color.Green
                    Case "30"
                        CheckBox30.BackColor = Color.Green
                    Case "29"
                        CheckBox29.BackColor = Color.Green
                    Case "28"
                        CheckBox28.BackColor = Color.Green
                    Case "27"
                        CheckBox27.BackColor = Color.Green
                    Case "26"
                        CheckBox26.BackColor = Color.Green
                    Case "25"
                        CheckBox25.BackColor = Color.Green
                    Case "24"
                        CheckBox24.BackColor = Color.Green
                    Case "23"
                        CheckBox23.BackColor = Color.Green
                    Case "22"
                        CheckBox22.BackColor = Color.Green
                    Case "21"
                        CheckBox21.BackColor = Color.Green
                    Case "20"
                        CheckBox20.BackColor = Color.Green
                    Case "19"
                        CheckBox19.BackColor = Color.Green
                    Case "18"
                        CheckBox18.BackColor = Color.Green
                    Case "17"
                        CheckBox17.BackColor = Color.Green
                    Case "16"
                        CheckBox16.BackColor = Color.Green
                    Case "15"
                        CheckBox15.BackColor = Color.Green
                    Case "14"
                        CheckBox14.BackColor = Color.Green
                    Case "13"
                        CheckBox13.BackColor = Color.Green
                    Case "12"
                        CheckBox12.BackColor = Color.Green
                    Case "11"
                        CheckBox11.BackColor = Color.Green
                    Case "10"
                        CheckBox10.BackColor = Color.Green
                    Case "9"
                        CheckBox9.BackColor = Color.Green
                    Case "8"
                        CheckBox8.BackColor = Color.Green
                    Case "7"
                        CheckBox7.BackColor = Color.Green
                    Case "6"
                        CheckBox6.BackColor = Color.Green
                    Case "5"
                        CheckBox5.BackColor = Color.Green
                    Case "4"
                        CheckBox4.BackColor = Color.Green
                    Case "3"
                        CheckBox3.BackColor = Color.Green
                    Case "2"
                        CheckBox2.BackColor = Color.Green
                End Select

                If CBNUM = "" Then
                    RadioButton2.Checked = True
                End If
                stopflag = 0



                'Start a new workbook in Excel.

                oExcel = CreateObject("Excel.Application")
                oBook = oExcel.Workbooks.Add
                oSheet = oBook.Worksheets(1)






                RadioButton1.AutoCheck = True
                RadioButton2.AutoCheck = True
                RadioButton3.AutoCheck = True
                RadioButton4.AutoCheck = True
                RadioButton5.AutoCheck = True
                RadioButton6.AutoCheck = True
                RadioButton7.AutoCheck = True
                RadioButton8.AutoCheck = True
                RadioButton9.AutoCheck = True
                RadioButton10.AutoCheck = True
                RadioButton11.AutoCheck = True
                RadioButton12.AutoCheck = True
                RadioButton13.AutoCheck = True
                RadioButton14.AutoCheck = True
                RadioButton15.AutoCheck = True
                RadioButton16.AutoCheck = True
                RadioButton17.AutoCheck = True
                RadioButton18.AutoCheck = True
                RadioButton19.AutoCheck = True
                RadioButton20.AutoCheck = True
                RadioButton21.AutoCheck = True
                RadioButton22.AutoCheck = True
                RadioButton23.AutoCheck = True
                RadioButton24.AutoCheck = True
                RadioButton25.AutoCheck = True
                RadioButton26.AutoCheck = True
                RadioButton27.AutoCheck = True
                RadioButton28.AutoCheck = True
                RadioButton29.AutoCheck = True
                RadioButton30.AutoCheck = True
                RadioButton31.AutoCheck = True
                RadioButton32.AutoCheck = True
                RadioButton33.AutoCheck = True
                RadioButton34.AutoCheck = True
                RadioButton35.AutoCheck = True
                RadioButton36.AutoCheck = True
                RadioButton37.AutoCheck = True
                RadioButton38.AutoCheck = True
                RadioButton39.AutoCheck = True
                RadioButton40.AutoCheck = True



                '---------------------------------------------------Change for each CT-----------------------------------------------------------------------------------------------------------------------
                Call Find_check(0, e)
                If CBNUM = "RadioButton2" Then
                    Call RLYOFF()

                    ''''''' close relay 2 ******************************************************
                    WriteDoChannel2(1, 22)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 22)
                    '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                    Call Body()
                    If Not Button13.BackColor = Color.Green Then

                        RadioButton3.Checked = True

                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If

                '--------------------------------------------------------------------
                ''''''' close relay 3 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton3" Then


                    Call RLYOFF()
                    WriteDoChannel2(1, 21)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 21)
                    '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton4.Checked = True

                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 4 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton4" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 20)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 20)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                        RadioButton5.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If

                '--------------------------------------------------------------------
                ''''''' close relay 5 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton5" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 19)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 19)

                    Call Body()
                    '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton6.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 6 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton6" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 18)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 18)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                        RadioButton7.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 7 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton7" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 17)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 17)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                        RadioButton8.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If

                '--------------------------------------------------------------------
                ''''''' close relay 8 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton8" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 16)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 16)

                    Call Body()
                    '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton9.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If

                '--------------------------------------------------------------------
                ''''''' close relay 9 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton9" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 15)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 15)

                    Call Body()
                    '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton10.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If

                '--------------------------------------------------------------------
                ''''''' close relay 10 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton10" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 14)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 14)

                    Call Body()

                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------


                        RadioButton11.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If

                '--------------------------------------------------------------------
                ''''''' close relay 11 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton11" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 13)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 13)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton12.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '-----
                ' --------------------------------------------------------------------------------------------------


                '--------------------------------------------------------------------
                ''''''' close relay 12 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton12" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 12)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 12)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------
                        RadioButton13.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 13 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton13" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 11)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 11)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------
                        RadioButton14.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 14 ******************************************************

                Call Find_check(0, e)
                If CBNUM = "RadioButton14" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 10)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 10)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton15.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 15 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton15" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 9)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 9)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton16.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 16 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton16" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 8)
                    ' AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 8)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------
                        RadioButton17.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 17 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton17" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 7)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 7)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton18.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 18 ******************************************************

                Call Find_check(0, e)
                If CBNUM = "RadioButton18" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 6)
                    ' AxAdvDIO2.DeviceNumber = 2
                    ' AxAdvDIO2.WriteDoChannel(1, 6)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------
                        RadioButton19.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If

                '--------------------------------------------------------------------
                ''''''' close relay 19 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton19" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 5)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 5)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton20.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 20 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton20" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 4)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 4)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------
                        RadioButton21.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                ''''''' close relay 21 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton21" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 3)
                    ' AxAdvDIO2.DeviceNumber = 2
                    ' AxAdvDIO2.WriteDoChannel(1, 3)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton22.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If

                '--------------------------------------------------------------------
                ''''''' close relay 22 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton22" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 2)
                    'AxAdvDIO2.DeviceNumber = 2
                    ' AxAdvDIO2.WriteDoChannel(1, 2)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------
                        RadioButton23.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 23 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton23" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 1)
                    'AxAdvDIO2.DeviceNumber = 2
                    ' AxAdvDIO2.WriteDoChannel(1, 1)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------
                        RadioButton24.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 24 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton24" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 0)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 0)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------
                        RadioButton25.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If

                '--------------------------------------------------------------------
                ''''''' close relay 25 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton25" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 47)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 47)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton26.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 26 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton26" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 46)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 46)

                    Call Body()

                    If Not Button13.BackColor = Color.Green Then

                        '-------------------------------------------------------------------------------------------------------

                        RadioButton27.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 27 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton27" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 45)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 45)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------
                        RadioButton28.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 28 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton28" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 44)
                    ' AxAdvDIO2.DeviceNumber = 2
                    ' AxAdvDIO2.WriteDoChannel(1, 44)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------

                        RadioButton29.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 29 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton29" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 43)
                    ' AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 43)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton30.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 30 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton30" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 42)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 42)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton31.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 31 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton31" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 41)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 41)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton32.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 32 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton32" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 40)
                    ' AxAdvDIO2.DeviceNumber = 2
                    ' AxAdvDIO2.WriteDoChannel(1, 40)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton33.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 33 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton33" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 39)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 39)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton34.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 34 ******************************************************

                Call Find_check(0, e)
                If CBNUM = "RadioButton34" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 38)
                    ' AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 38)

                    Call Body()
                    If Not Button13.BackColor = Color.Green Then
                        '-------------------------------------------------------------------------------------------------------

                        RadioButton35.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 35 ******************************************************

                Call Find_check(0, e)
                If CBNUM = "RadioButton35" Then
                    Call RLYOFF()
                    WriteDoChannel2(1, 37)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 37)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton36.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 36 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton36" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 36)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 36)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton37.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 37 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton37" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 35)
                    ' AxAdvDIO2.DeviceNumber = 2
                    ' AxAdvDIO2.WriteDoChannel(1, 35)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton38.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 38 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton38" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 34)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 34)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton39.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 39 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton39" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 33)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 33)

                    Call Body()
                    '-------------------------------------------------------------------------------------------------------
                    If Not Button13.BackColor = Color.Green Then
                        RadioButton40.Checked = True
                        If stopflag = 1 Then
                            Exit Sub
                        End If
                    End If
                End If
                '--------------------------------------------------------------------
                ''''''' close relay 40 ******************************************************
                Call Find_check(0, e)
                If CBNUM = "RadioButton40" Then

                    Call RLYOFF()
                    WriteDoChannel2(1, 32)
                    'AxAdvDIO2.DeviceNumber = 2
                    'AxAdvDIO2.WriteDoChannel(1, 32)
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                    Call Body()

                    If Not Button13.BackColor = Color.Green Then
                        Exit Sub

                    End If
                End If
                '-------------------------------------------------------------------------------------------------------
                On Error Resume Next
                If Not ACCTESTFLAG = 1 Then
                    If repx = 0 Then

                        oExcel.DisplayAlerts = False
                        Dim dt As New DateTime
                        dt = DateTime.Now
                        Dim mt As String
                        mt = String.Format("C:\MCMKARESULTS_{0:MMM_dd_yyyy_HH-mm}.xlsx", dt)
                        oBook.Worksheets(1).SaveAs(mt.ToString())

                    End If
                End If
                If repx = 1 Then
                    oExcel.DisplayAlerts = False
                    oBook.Worksheets(1).SaveAs("C:\MCMKARepeat1.xlsx")
                    oSheet = Nothing
                    oBook = Nothing
                    oExcel.Quit()
                    oExcel = Nothing
                    GC.Collect()

                End If

                If repx = 2 Then
                    oExcel.DisplayAlerts = False
                    oBook.Worksheets(1).SaveAs("C:\MCMKARepeat2.xlsx")
                    oSheet = Nothing
                    oBook = Nothing
                    oExcel.Quit()
                    oExcel = Nothing
                    GC.Collect()

                End If
                If repx = 3 Then
                    oExcel.DisplayAlerts = False
                    oBook.Worksheets(1).SaveAs("C:\MCMKARepeat3.xlsx")
                    oSheet = Nothing
                    oBook = Nothing
                    oExcel.Quit()
                    oExcel = Nothing
                    GC.Collect()

                End If
                If repx = 4 Then
                    oExcel.DisplayAlerts = False
                    oBook.Worksheets(1).SaveAs("C:\MCMKARepeat4.xlsx")
                    oSheet = Nothing
                    oBook = Nothing
                    oExcel.Quit()
                    oExcel = Nothing
                    GC.Collect()

                End If

            Next checkx

            If Not stopflag = 1 Then


                Me.Label36.Text = "DAILY ACCURACY TEST   Complete"
                Me.Label36.Font = New Drawing.Font("Times New Roman", 20, FontStyle.Bold)
                Me.Label36.BackColor = Color.LimeGreen

                For Each MyObject In Me.GroupBox3.Controls

                    If TypeOf MyObject Is CheckBox Then

                        MyObject.Enabled = True

                    End If

                Next

            End If




        Else
            If CBNUM = "" Then
                RadioButton2.Checked = True
            End If
            stopflag = 0



            'Start a new workbook in Excel.
            Dim dt As DialogResult
            dt = MessageBox.Show("This program will automatically close down ALL Excel applications.\r\nPlease ensure you do not have any excel files open or your data will be lost!", "Warning - All Excel instances will be closed", MessageBoxButtons.OKCancel)
            If repx <> 0 Then
                If dt = Windows.Forms.DialogResult.Cancel Then
                    Exit Sub
                End If
            End If

            Dim processes() As Process = Process.GetProcesses
            For p As Integer = processes.Count - 1 To 0 Step -1
                If processes(p).ProcessName = "EXCEL" Then
                    processes(p).Kill()
                End If
            Next


            'Start a new workbook in Excel.
            Thread.Sleep(200)

            File.Delete("C:\\MCTemp.xlsx")


            oExcel = CreateObject("Excel.Application")
            oBook = oExcel.Workbooks.Add
            oSheet = oBook.Worksheets(1)




            RadioButton1.AutoCheck = True
            RadioButton2.AutoCheck = True
            RadioButton3.AutoCheck = True
            RadioButton4.AutoCheck = True
            RadioButton5.AutoCheck = True
            RadioButton6.AutoCheck = True
            RadioButton7.AutoCheck = True
            RadioButton8.AutoCheck = True
            RadioButton9.AutoCheck = True
            RadioButton10.AutoCheck = True
            RadioButton11.AutoCheck = True
            RadioButton12.AutoCheck = True
            RadioButton13.AutoCheck = True
            RadioButton14.AutoCheck = True
            RadioButton15.AutoCheck = True
            RadioButton16.AutoCheck = True
            RadioButton17.AutoCheck = True
            RadioButton18.AutoCheck = True
            RadioButton19.AutoCheck = True
            RadioButton20.AutoCheck = True
            RadioButton21.AutoCheck = True
            RadioButton22.AutoCheck = True
            RadioButton23.AutoCheck = True
            RadioButton24.AutoCheck = True
            RadioButton25.AutoCheck = True
            RadioButton26.AutoCheck = True
            RadioButton27.AutoCheck = True
            RadioButton28.AutoCheck = True
            RadioButton29.AutoCheck = True
            RadioButton30.AutoCheck = True
            RadioButton31.AutoCheck = True
            RadioButton32.AutoCheck = True
            RadioButton33.AutoCheck = True
            RadioButton34.AutoCheck = True
            RadioButton35.AutoCheck = True
            RadioButton36.AutoCheck = True
            RadioButton37.AutoCheck = True
            RadioButton38.AutoCheck = True
            RadioButton39.AutoCheck = True
            RadioButton40.AutoCheck = True



            '---------------------------------------------------Change for each CT-----------------------------------------------------------------------------------------------------------------------
            Call Find_check(0, e)
            If CBNUM = "RadioButton2" Then
                Call RLYOFF()

                ''''''' close relay 2 ******************************************************
                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 22)
                WriteDoChannel2(1, 22)


                '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                Call Body()
                If Not Button13.BackColor = Color.Green Then

                    RadioButton3.Checked = True

                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If

            '--------------------------------------------------------------------
            ''''''' close relay 3 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton3" Then


                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 21)
                WriteDoChannel2(1, 21)
                '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    RadioButton4.Checked = True

                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 4 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton4" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 20)
                WriteDoChannel2(1, 20)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                    RadioButton5.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If

            '--------------------------------------------------------------------
            ''''''' close relay 5 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton5" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 19)
                WriteDoChannel2(1, 19)
                Call Body()
                '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                If Not Button13.BackColor = Color.Green Then
                    RadioButton6.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 6 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton6" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 18)
                WriteDoChannel2(1, 18)

                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                    RadioButton7.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 7 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton7" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 17)
                WriteDoChannel2(1, 17)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                    RadioButton8.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If

            '--------------------------------------------------------------------
            ''''''' close relay 8 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton8" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 16)
                WriteDoChannel2(1, 16)
                Call Body()
                '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                If Not Button13.BackColor = Color.Green Then
                    RadioButton9.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If

            '--------------------------------------------------------------------
            ''''''' close relay 9 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton9" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 15)
                WriteDoChannel2(1, 15)
                Call Body()
                '----------------------------------------------------------------------------------------------------------------------------------       '*****        ''''''' clo
                If Not Button13.BackColor = Color.Green Then
                    RadioButton10.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If

            '--------------------------------------------------------------------
            ''''''' close relay 10 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton10" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 14)
                WriteDoChannel2(1, 14)
                Call Body()

                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------


                    RadioButton11.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If

            '--------------------------------------------------------------------
            ''''''' close relay 11 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton11" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 13)
                WriteDoChannel2(1, 13)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    RadioButton12.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '-----
            ' --------------------------------------------------------------------------------------------------


            '--------------------------------------------------------------------
            ''''''' close relay 12 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton12" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 12)
                WriteDoChannel2(1, 12)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------
                    RadioButton13.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 13 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton13" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 11)
                WriteDoChannel2(1, 11)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------
                    RadioButton14.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 14 ******************************************************

            Call Find_check(0, e)
            If CBNUM = "RadioButton14" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 10)
                WriteDoChannel2(1, 10)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton15.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 15 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton15" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 9)
                WriteDoChannel2(1, 9)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton16.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 16 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton16" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 8)
                WriteDoChannel2(1, 8)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------
                    RadioButton17.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 17 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton17" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 7)
                WriteDoChannel2(1, 7)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton18.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 18 ******************************************************

            Call Find_check(0, e)
            If CBNUM = "RadioButton18" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 6)
                WriteDoChannel2(1, 6)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------
                    RadioButton19.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If

            '--------------------------------------------------------------------
            ''''''' close relay 19 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton19" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 5)
                WriteDoChannel2(1, 5)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton20.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 20 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton20" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 4)
                WriteDoChannel2(1, 4)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------
                    RadioButton21.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            ''''''' close relay 21 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton21" Then

                Call RLYOFF()

                ' AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 3)
                WriteDoChannel2(1, 3)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton22.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If

            '--------------------------------------------------------------------
            ''''''' close relay 22 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton22" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 2)
                WriteDoChannel2(1, 2)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------
                    RadioButton23.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 23 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton23" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 1)
                WriteDoChannel2(1, 1)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------
                    RadioButton24.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 24 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton24" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 0)
                WriteDoChannel2(1, 0)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------
                    RadioButton25.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If

            '--------------------------------------------------------------------
            ''''''' close relay 25 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton25" Then
                Call RLYOFF()

                ' AxAdvDIO2.DeviceNumber = 2
                ' AxAdvDIO2.WriteDoChannel(1, 47)
                WriteDoChannel2(1, 47)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton26.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 26 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton26" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 46)
                WriteDoChannel2(1, 46)
                Call Body()

                If Not Button13.BackColor = Color.Green Then

                    '-------------------------------------------------------------------------------------------------------

                    RadioButton27.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 27 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton27" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 45)
                WriteDoChannel2(1, 45)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------
                    RadioButton28.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 28 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton28" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 44)
                WriteDoChannel2(1, 44)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------

                    RadioButton29.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 29 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton29" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 43)
                WriteDoChannel2(1, 43)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton30.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 30 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton30" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 42)
                WriteDoChannel2(1, 42)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton31.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 31 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton31" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 41)
                WriteDoChannel2(1, 41)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton32.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 32 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton32" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 40)
                WriteDoChannel2(1, 40)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton33.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 33 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton33" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 39)
                WriteDoChannel2(1, 39)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton34.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 34 ******************************************************

            Call Find_check(0, e)
            If CBNUM = "RadioButton34" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 38)
                WriteDoChannel2(1, 38)
                Call Body()
                If Not Button13.BackColor = Color.Green Then
                    '-------------------------------------------------------------------------------------------------------

                    RadioButton35.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 35 ******************************************************

            Call Find_check(0, e)
            If CBNUM = "RadioButton35" Then
                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 37)
                WriteDoChannel2(1, 37)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton36.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 36 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton36" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 36)
                WriteDoChannel2(1, 36)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton37.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 37 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton37" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 35)
                WriteDoChannel2(1, 35)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton38.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 38 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton38" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 34)
                WriteDoChannel2(1, 34)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton39.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 39 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton39" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 33)
                WriteDoChannel2(1, 33)
                Call Body()
                '-------------------------------------------------------------------------------------------------------
                If Not Button13.BackColor = Color.Green Then
                    RadioButton40.Checked = True
                    If stopflag = 1 Then
                        Exit Sub
                    End If
                End If
            End If
            '--------------------------------------------------------------------
            ''''''' close relay 40 ******************************************************
            Call Find_check(0, e)
            If CBNUM = "RadioButton40" Then

                Call RLYOFF()

                'AxAdvDIO2.DeviceNumber = 2
                'AxAdvDIO2.WriteDoChannel(1, 32)
                WriteDoChannel2(1, 32)
                If stopflag = 1 Then
                    Exit Sub
                End If
                Call Body()
                stopflag = 1
                If stopflag = 1 Then
                    Exit Sub
                End If
            End If

            '-------------------------------------------------------------------------------------------------------
            On Error Resume Next

            If repx = 0 Then
                If Not ACCTESTFLAG = 1 Then
                    oExcel.DisplayAlerts = False
                    Dim dt1 As New DateTime
                    dt1 = DateTime.Now
                    Dim mt As String
                    mt = String.Format("C:\MCMKARESULTS_{0:MMM_dd_yyyy_HH-mm}.xlsx", dt)
                    oBook.Worksheets(1).SaveAs(mt.ToString())

                End If
            End If

            If repx = 1 Then
                oExcel.DisplayAlerts = False
                oBook.Worksheets(1).SaveAs("C:\MCMKARepeat1.xlsx")
                oSheet = Nothing
                oBook = Nothing
                oExcel.Quit()
                oExcel = Nothing
                GC.Collect()

            End If

            If repx = 2 Then
                oExcel.DisplayAlerts = False
                oBook.Worksheets(1).SaveAs("C:\MCMKARepeat2.xlsx")
                oSheet = Nothing
                oBook = Nothing
                oExcel.Quit()
                oExcel = Nothing
                GC.Collect()

            End If
            If repx = 3 Then
                oExcel.DisplayAlerts = False
                oBook.Worksheets(1).SaveAs("C:\MCMKARepeat3.xlsx")
                oSheet = Nothing
                oBook = Nothing
                oExcel.Quit()
                oExcel = Nothing
                GC.Collect()

            End If
            If repx = 4 Then
                oExcel.DisplayAlerts = False
                oBook.Worksheets(1).SaveAs("C:\MCMKARepeat4.xlsx")
                oSheet = Nothing
                oBook = Nothing
                oExcel.Quit()
                oExcel = Nothing
                GC.Collect()

            End If
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Call RLYON()

        RadioButton1.AutoCheck = False
        RadioButton2.AutoCheck = False
        RadioButton3.AutoCheck = False
        RadioButton4.AutoCheck = False
        RadioButton5.AutoCheck = False
        RadioButton6.AutoCheck = False
        RadioButton7.AutoCheck = False
        RadioButton8.AutoCheck = False
        RadioButton9.AutoCheck = False
        RadioButton10.AutoCheck = False
        RadioButton11.AutoCheck = False
        RadioButton12.AutoCheck = False
        RadioButton13.AutoCheck = False
        RadioButton14.AutoCheck = False
        RadioButton15.AutoCheck = False
        RadioButton16.AutoCheck = False
        RadioButton17.AutoCheck = False
        RadioButton18.AutoCheck = False
        RadioButton19.AutoCheck = False
        RadioButton20.AutoCheck = False
        RadioButton21.AutoCheck = False
        RadioButton22.AutoCheck = False
        RadioButton23.AutoCheck = False
        RadioButton24.AutoCheck = False
        RadioButton25.AutoCheck = False
        RadioButton26.AutoCheck = False
        RadioButton27.AutoCheck = False
        RadioButton28.AutoCheck = False
        RadioButton29.AutoCheck = False
        RadioButton30.AutoCheck = False
        RadioButton31.AutoCheck = False
        RadioButton32.AutoCheck = False
        RadioButton33.AutoCheck = False
        RadioButton34.AutoCheck = False
        RadioButton35.AutoCheck = False
        RadioButton36.AutoCheck = False
        RadioButton37.AutoCheck = False
        RadioButton38.AutoCheck = False
        RadioButton39.AutoCheck = False
        RadioButton40.AutoCheck = False
        RadioButton1.Checked = True
        RadioButton2.Checked = True
        RadioButton3.Checked = True
        RadioButton4.Checked = True
        RadioButton5.Checked = True
        RadioButton6.Checked = True
        RadioButton7.Checked = True
        RadioButton8.Checked = True
        RadioButton9.Checked = True
        RadioButton10.Checked = True
        RadioButton11.Checked = True
        RadioButton12.Checked = True
        RadioButton13.Checked = True
        RadioButton14.Checked = True
        RadioButton15.Checked = True
        RadioButton16.Checked = True
        RadioButton17.Checked = True
        RadioButton18.Checked = True
        RadioButton19.Checked = True
        RadioButton20.Checked = True
        RadioButton21.Checked = True
        RadioButton22.Checked = True
        RadioButton23.Checked = True
        RadioButton24.Checked = True
        RadioButton25.Checked = True
        RadioButton26.Checked = True
        RadioButton27.Checked = True
        RadioButton28.Checked = True
        RadioButton29.Checked = True
        RadioButton30.Checked = True
        RadioButton31.Checked = True
        RadioButton32.Checked = True
        RadioButton33.Checked = True
        RadioButton34.Checked = True
        RadioButton35.Checked = True
        RadioButton36.Checked = True
        RadioButton37.Checked = True
        RadioButton38.Checked = True
        RadioButton39.Checked = True
        RadioButton40.Checked = True
    End Sub

    Private Sub RadioButton17_CheckedChanged(sender As Object, e As EventArgs)

    End Sub
    Friend WithEvents tbMain As System.Windows.Forms.TabPage
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents SETRELAYS As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents btnOpenFile As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button4 As System.Windows.Forms.Button





    Private mbSession As MessageBasedSession






    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim Voltage As String

        Dim result As DialogResult = OpenFileDialog1.ShowDialog()


        If result = Windows.Forms.DialogResult.OK Then

            ' Get the file name.
            Dim path As String = OpenFileDialog1.FileName()
            Dim text As String = File.ReadAllText(path)
            TextBox8.Text = File.ReadAllText(path)
            tboxprn = TextBox8.Text
        End If
        Dim FirstCharacter As Integer = TextBox8.Text.IndexOf("1    P")
        Voltage = Mid(TextBox8.Text, (FirstCharacter + 7), 3)
        TextBox2.Text = Voltage

    End Sub
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox

    Private Sub FTestApp_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call RLYON()
        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While
        BackgroundWorker2.RunWorkerAsync(5)

    End Sub

    Private Sub StopRadian()
        'Dim RadianStatus As String = ""
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)


        Threading.Thread.Sleep(5000)
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If
        Threading.Thread.Sleep(1000)
        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        'Return RadianStatus
    End Sub



    Private Sub StopTest()
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)





        Dim MyObject As Control

        For Each MyObject In Me.GroupBox3.Controls

            If TypeOf MyObject Is CheckBox Then
                MyObject.Enabled = True
                DirectCast(MyObject, CheckBox).Checked = False
                MyObject.BackColor = Color.LightSteelBlue
            End If

        Next

        repx = 0
        stopflag = 1









        mbSession = CType(ResourceManager.GetLocalManager.Open("GPIB0::0::INSTR"), MessageBasedSession)
        mbSession.Write("CHAN 2 Rang 10V AMR 0.0001")
        mbSession.Write("CHAN 2 OUTP OFF ")
        Call closeSession()
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(0, 33)
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(0, 34)
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(0, 39)
        'AxAdvDIO1.DeviceNumber = 0
        'AxAdvDIO1.WriteDoChannel(0, 22)
        'AxAdvDIO1.WriteDoChannel(0, 23)
        WriteDoChannel1(0, 33)
        WriteDoChannel1(0, 34)
        WriteDoChannel1(0, 39)
        WriteDoChannel0(0, 22)
        WriteDoChannel0(0, 23)
        If Not ACCTESTFLAG = 1 Then

            TextBox8.Text = ""
            TextBox2.Text = ""
        End If



        TextBox2.BackColor = Color.White
        TextBox3.BackColor = Color.White
        TextBox5.BackColor = Color.White
        TextBox6.BackColor = Color.White
        TextBox7.BackColor = Color.White
        TextBox10.BackColor = Color.White
        TextBox10.Text = ""
        TextBox11.BackColor = Color.White
        TextBox11.Text = ""
        TextBox12.BackColor = Color.White
        TextBox12.Text = ""
        TextBox13.BackColor = Color.White
        TextBox13.Text = ""
        TextBox14.BackColor = Color.White
        TextBox14.Text = ""
        TextBox15.BackColor = Color.White
        TextBox15.Text = ""
        TextBox16.BackColor = Color.White
        TextBox16.Text = ""
        TextBox17.BackColor = Color.White
        TextBox17.Text = ""
        TextBox18.BackColor = Color.White
        TextBox18.Text = ""
        TextBox19.BackColor = Color.White
        TextBox19.Text = ""
        TextBox20.BackColor = Color.White
        TextBox20.Text = ""
        TextBox21.BackColor = Color.White
        TextBox21.Text = ""
        TextBox22.BackColor = Color.White
        TextBox22.Text = ""
        TextBox23.BackColor = Color.White
        TextBox23.Text = ""
        TextBox24.BackColor = Color.White
        TextBox24.Text = ""
        TextBox25.BackColor = Color.White
        TextBox25.Text = ""
        TextBox26.BackColor = Color.White
        TextBox26.Text = ""
        TextBox27.BackColor = Color.White
        TextBox27.Text = ""
        TextBox28.BackColor = Color.White
        TextBox28.Text = ""
        TextBox29.BackColor = Color.White
        TextBox29.Text = ""
        TextBox30.BackColor = Color.White
        TextBox30.Text = ""
        TextBox31.BackColor = Color.White
        TextBox31.Text = ""
        TextBox32.BackColor = Color.White
        TextBox32.Text = ""
        TextBox33.BackColor = Color.White
        TextBox33.Text = ""
        TextBox34.BackColor = Color.White
        TextBox34.Text = ""
        TextBox35.BackColor = Color.White
        TextBox35.Text = ""
        TextBox36.BackColor = Color.White
        TextBox36.Text = ""
        TextBox37.BackColor = Color.White
        TextBox37.Text = ""
        TextBox38.BackColor = Color.White
        TextBox38.Text = ""
        TextBox39.BackColor = Color.White
        TextBox39.Text = ""
        TextBox40.BackColor = Color.White
        TextBox40.Text = ""
        TextBox41.BackColor = Color.White
        TextBox41.Text = ""
        TextBox42.BackColor = Color.White
        TextBox42.Text = ""
        TextBox43.BackColor = Color.White
        TextBox43.Text = ""
        TextBox44.BackColor = Color.White
        TextBox44.Text = ""
        TextBox45.BackColor = Color.White
        TextBox45.Text = ""
        TextBox46.BackColor = Color.White
        TextBox46.Text = ""
        TextBox47.BackColor = Color.White
        TextBox47.Text = ""
        TextBox48.BackColor = Color.White
        TextBox48.Text = ""
        TextBox53.Text = ""
        Me.TextBox1.BackColor = Color.White
        Me.TextBox1.Text = ""
        Me.TextBox4.BackColor = Color.White
        Me.TextBox4.Text = ""


        tboxprn = ""




        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If
        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        Call RLYOFF()



        ''Save the Workbook and quit Excel.


        If repx = 0 Then
            On Error Resume Next
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If

        If repx = 1 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat1temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If

        If repx = 2 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat2temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If
        If repx = 3 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat3temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If
        If repx = 4 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat4temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If



        repx = 0





    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call StopTest()

        'Dim commport1 As Byte = 4
        'Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        'Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        'Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        'Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        'Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)





        'Dim MyObject As Control

        'For Each MyObject In Me.GroupBox3.Controls

        '    If TypeOf MyObject Is CheckBox Then
        '        MyObject.Enabled = True
        '        DirectCast(MyObject, CheckBox).Checked = False
        '        MyObject.BackColor = Color.LightSteelBlue
        '    End If

        'Next

        'repx = 0
        'stopflag = 1









        'mbSession = CType(ResourceManager.GetLocalManager.Open("GPIB0::0::INSTR"), MessageBasedSession)
        'mbSession.Write("CHAN 2 Rang 10V AMR 0.0001")
        'mbSession.Write("CHAN 2 OUTP OFF ")
        'Call closeSession()
        ''AxAdvDIO1.DeviceNumber = 1
        ''AxAdvDIO1.WriteDoChannel(0, 33)
        ''AxAdvDIO1.DeviceNumber = 1
        ''AxAdvDIO1.WriteDoChannel(0, 34)
        ''AxAdvDIO1.DeviceNumber = 1
        ''AxAdvDIO1.WriteDoChannel(0, 39)
        ''AxAdvDIO1.DeviceNumber = 0
        ''AxAdvDIO1.WriteDoChannel(0, 22)
        ''AxAdvDIO1.WriteDoChannel(0, 23)
        'WriteDoChannel1(0, 33)
        'WriteDoChannel1(0, 34)
        'WriteDoChannel1(0, 39)
        'WriteDoChannel0(0, 22)
        'WriteDoChannel0(0, 23)
        'If Not ACCTESTFLAG = 1 Then

        '    TextBox8.Text = ""
        '    TextBox2.Text = ""
        'End If



        'TextBox2.BackColor = Color.White
        'TextBox3.BackColor = Color.White
        'TextBox5.BackColor = Color.White
        'TextBox6.BackColor = Color.White
        'TextBox7.BackColor = Color.White
        'TextBox10.BackColor = Color.White
        'TextBox10.Text = ""
        'TextBox11.BackColor = Color.White
        'TextBox11.Text = ""
        'TextBox12.BackColor = Color.White
        'TextBox12.Text = ""
        'TextBox13.BackColor = Color.White
        'TextBox13.Text = ""
        'TextBox14.BackColor = Color.White
        'TextBox14.Text = ""
        'TextBox15.BackColor = Color.White
        'TextBox15.Text = ""
        'TextBox16.BackColor = Color.White
        'TextBox16.Text = ""
        'TextBox17.BackColor = Color.White
        'TextBox17.Text = ""
        'TextBox18.BackColor = Color.White
        'TextBox18.Text = ""
        'TextBox19.BackColor = Color.White
        'TextBox19.Text = ""
        'TextBox20.BackColor = Color.White
        'TextBox20.Text = ""
        'TextBox21.BackColor = Color.White
        'TextBox21.Text = ""
        'TextBox22.BackColor = Color.White
        'TextBox22.Text = ""
        'TextBox23.BackColor = Color.White
        'TextBox23.Text = ""
        'TextBox24.BackColor = Color.White
        'TextBox24.Text = ""
        'TextBox25.BackColor = Color.White
        'TextBox25.Text = ""
        'TextBox26.BackColor = Color.White
        'TextBox26.Text = ""
        'TextBox27.BackColor = Color.White
        'TextBox27.Text = ""
        'TextBox28.BackColor = Color.White
        'TextBox28.Text = ""
        'TextBox29.BackColor = Color.White
        'TextBox29.Text = ""
        'TextBox30.BackColor = Color.White
        'TextBox30.Text = ""
        'TextBox31.BackColor = Color.White
        'TextBox31.Text = ""
        'TextBox32.BackColor = Color.White
        'TextBox32.Text = ""
        'TextBox33.BackColor = Color.White
        'TextBox33.Text = ""
        'TextBox34.BackColor = Color.White
        'TextBox34.Text = ""
        'TextBox35.BackColor = Color.White
        'TextBox35.Text = ""
        'TextBox36.BackColor = Color.White
        'TextBox36.Text = ""
        'TextBox37.BackColor = Color.White
        'TextBox37.Text = ""
        'TextBox38.BackColor = Color.White
        'TextBox38.Text = ""
        'TextBox39.BackColor = Color.White
        'TextBox39.Text = ""
        'TextBox40.BackColor = Color.White
        'TextBox40.Text = ""
        'TextBox41.BackColor = Color.White
        'TextBox41.Text = ""
        'TextBox42.BackColor = Color.White
        'TextBox42.Text = ""
        'TextBox43.BackColor = Color.White
        'TextBox43.Text = ""
        'TextBox44.BackColor = Color.White
        'TextBox44.Text = ""
        'TextBox45.BackColor = Color.White
        'TextBox45.Text = ""
        'TextBox46.BackColor = Color.White
        'TextBox46.Text = ""
        'TextBox47.BackColor = Color.White
        'TextBox47.Text = ""
        'TextBox48.BackColor = Color.White
        'TextBox48.Text = ""
        'TextBox53.Text = ""
        'Me.TextBox1.BackColor = Color.White
        'Me.TextBox1.Text = ""
        'Me.TextBox4.BackColor = Color.White
        'Me.TextBox4.Text = ""


        'tboxprn = ""




        'Status = RadRDAssignDevice(CShort(commport1), comDevice)

        'If Status = 0 Then
        '    'Successfully connected
        '    'Get unit information and populate status bar
        '    Status = RadRDModel(comDevice, Model)
        '    Status = RadRDSerial(comDevice, Serial)
        '    Status = RadRDVersion(comDevice, Version)
        '    Status = RadRDName(comDevice, DeviceName)
        '    RadRDAccumStop(comDevice)
        '    RadRDAccumReset(comDevice, 0)
        'End If
        'If comDevice <> 0 Then
        '    RadRDReleaseDevice(comDevice)
        '    comDevice = 0
        'End If
        'Status = RadRDAssignDevice(CShort(commport), comDevice)

        'If Status = 0 Then
        '    'Successfully connected
        '    'Get unit information and populate status bar
        '    Status = RadRDModel(comDevice, Model)
        '    Status = RadRDSerial(comDevice, Serial)
        '    Status = RadRDVersion(comDevice, Version)
        '    Status = RadRDName(comDevice, DeviceName)
        '    RadRDAccumStop(comDevice)
        '    RadRDAccumReset(comDevice, 0)
        'End If

        'If comDevice <> 0 Then
        '    RadRDReleaseDevice(comDevice)
        '    comDevice = 0
        'End If


        'Call RLYOFF()



        ' ''Save the Workbook and quit Excel.


        'If repx = 0 Then
        '    On Error Resume Next
        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")
        '    oSheet = Nothing
        '    oBook = Nothing
        '    oExcel.Quit()
        '    oExcel = Nothing
        '    GC.Collect()

        'End If

        'If repx = 1 Then
        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCMKARepeat1temp.xlsx")
        '    oSheet = Nothing
        '    oBook = Nothing
        '    oExcel.Quit()
        '    oExcel = Nothing
        '    GC.Collect()

        'End If

        'If repx = 2 Then
        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCMKARepeat2temp.xlsx")
        '    oSheet = Nothing
        '    oBook = Nothing
        '    oExcel.Quit()
        '    oExcel = Nothing
        '    GC.Collect()

        'End If
        'If repx = 3 Then
        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCMKARepeat3temp.xlsx")
        '    oSheet = Nothing
        '    oBook = Nothing
        '    oExcel.Quit()
        '    oExcel = Nothing
        '    GC.Collect()

        'End If
        'If repx = 4 Then
        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCMKARepeat4temp.xlsx")
        '    oSheet = Nothing
        '    oBook = Nothing
        '    oExcel.Quit()
        '    oExcel = Nothing
        '    GC.Collect()

        'End If



        'repx = 0





    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Call RLYOFF()


        RadioButton1.AutoCheck = True
        RadioButton2.AutoCheck = True
        RadioButton3.AutoCheck = True
        RadioButton4.AutoCheck = True
        RadioButton5.AutoCheck = True
        RadioButton6.AutoCheck = True
        RadioButton7.AutoCheck = True
        RadioButton8.AutoCheck = True
        RadioButton9.AutoCheck = True
        RadioButton10.AutoCheck = True
        RadioButton11.AutoCheck = True
        RadioButton12.AutoCheck = True
        RadioButton13.AutoCheck = True
        RadioButton14.AutoCheck = True
        RadioButton15.AutoCheck = True
        RadioButton16.AutoCheck = True
        RadioButton17.AutoCheck = True
        RadioButton18.AutoCheck = True
        RadioButton19.AutoCheck = True
        RadioButton20.AutoCheck = True
        RadioButton21.AutoCheck = True
        RadioButton22.AutoCheck = True
        RadioButton23.AutoCheck = True
        RadioButton24.AutoCheck = True
        RadioButton25.AutoCheck = True
        RadioButton26.AutoCheck = True
        RadioButton27.AutoCheck = True
        RadioButton28.AutoCheck = True
        RadioButton29.AutoCheck = True
        RadioButton30.AutoCheck = True
        RadioButton31.AutoCheck = True
        RadioButton32.AutoCheck = True
        RadioButton33.AutoCheck = True
        RadioButton34.AutoCheck = True
        RadioButton35.AutoCheck = True
        RadioButton36.AutoCheck = True
        RadioButton37.AutoCheck = True
        RadioButton38.AutoCheck = True
        RadioButton39.AutoCheck = True
        RadioButton40.AutoCheck = True












    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox("Paused")
    End Sub

    Private Function InsertCommonEscapeSequences(ByVal s As String) As String
        Return s.Replace(vbLf, "\n").Replace(vbCr, "\r")
    End Function
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button11 As System.Windows.Forms.Button
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox

    Public Sub teststep1()
        Dim CURinternal As Single
        Dim IntCount(256) As Long
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)
        Dim test As String
        Dim x As Single
        Dim Y As Single
        Dim PFinternal As Single

        'Call StopRadian()
        'Threading.Thread.Sleep(1000)

        Me.TextBox1.BackColor = Color.White
        Me.TextBox1.Text = ""

        TextBox3.BackColor = Color.LimeGreen
        TextBox7.BackColor = Color.White
        Application.DoEvents()
        mbSession.Write("IDEN?")
        test = mbSession.ReadString()

        If Len(test) > 1 Then
            mbSession.Write("CHAN 2  FREQ 60.0000") 'Set Frequency
            mbSession.Write("CHAN 1  FREQ 60.0000")

            mbSession.Write("CHAN 1 MODE GATE")
            mbSession.Write("CHAN 1 GATE OFF")

            mbSession.Write("CHAN 2 OUTP OFF")
            ' mbSession.Write("CHAN 2 Rang 1V;AMR 0.0020") '0.1258")
            mbSession.Write("CHAN 2 PHAS 0.00")
            mbSession.Write("CHAN 2 FNC")


            mbSession.Write("CHAN 1 OUTP OFF")
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.0020") '0.1258")
            mbSession.Write("CHAN 1 PHAS 0.00")
            mbSession.Write("CHAN 1 FNC")


        mbSession.Write("CHAN 2 OUTP ON ")
            mbSession.Write("CHAN 1 OUTP ON")
            ' mbSession.Write("CHAN 2 Rang 1V;AMR 0.0020;OFFS 0") '0.1258")
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.0020;OFFS 0") '0.1258")

            'mbSession.Write("CHAN 2 SPH +1")         'Set Phase
            'mbSession.Write("CHAN 1 SPH +6.853")
            'mbSession.Write("CHAN 2 SPH +0.4")         'Set Phase
            'mbSession.Write("CHAN 2 SPH +358.6") ' Changed November 21, 2022
            'mbSession.Write("CHAN 2 SPH +0.0200") ' Changed November 22, 2022
            mbSession.Write("CHAN 2 SPH +0.2500") ' Changed May 3rd 2023 - KW
            mbSession.Write("CHAN 1 SPH +1.0") 'Changed November 16, 2022
            mbSession.Write("CHAN 1 OFFS 0") '0.1258")
            'mbSession.Write("CHAN 2 Rang 10V;AMR 5.208") '0.1258")
            mbSession.Write("CHAN 2 FNO 1")

            'mbSession.Write("CHAN 1 Rang 1V;AMR 0.0126") '0.1258") ' Changed on November 17, 2022 - CH
            'mbSession.Write("CHAN 1 Rang 1V;AMR 0.0131") '0.1258") 'Changed May 3rd 2023 KW
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.0125") '0.1258")
            mbSession.Write("CHAN 1 FNO 1")


            Threading.Thread.Sleep(500)
        'AxAdvDIO1.DeviceNumber = 1
            ' AxAdvDIO1.WriteDoChannel(1, 33)
            'TURNS ON RELAY ON VALHALLA - JS'
        WriteDoChannel1(1, 33)
            Threading.Thread.Sleep(100)

            mbSession.Write("CHAN 1 GATE ON")
            Threading.Thread.Sleep(400)

            'Threading.Thread.Sleep(1000)
            'Call StopRadian()
            'Threading.Thread.Sleep(500)
            'Call StopRadian()

        End If




        'Threading.Thread.Sleep(1000)

        Threading.Thread.Sleep(300)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If
        '**********************************************************************************************************************************
        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            Threading.Thread.Sleep(1000)
            RadRDInstMetric(comDevice, 0, RAD_INST_A, CURinternal)
            x = CURinternal
            Me.TextBox1.Text = x.ToString
            Me.TextBox1.BackColor = Color.LimeGreen

        End If
        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            Threading.Thread.Sleep(1000)
            RadRDInstMetric(comDevice, 0, RAD_INST_PF, PFinternal)
            Y = PFinternal
            Me.TextBox4.Text = Y.ToString
            Me.TextBox4.BackColor = Color.LimeGreen

        End If


        Application.DoEvents()






    End Sub
    Public Sub teststep2()
        Dim test As String
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)
        Dim x As Single
        Dim CURinternal As Single
        Dim Y As Single
        Dim PFinternal As Single
        TextBox7.BackColor = Color.White

        'Call StopRadian()

        'Threading.Thread.Sleep(1000)

        Me.TextBox1.BackColor = Color.White
        Me.TextBox1.Text = ""
        TextBox3.BackColor = Color.White
        TextBox5.BackColor = Color.LimeGreen
        Application.DoEvents()
        mbSession.Write("IDEN?")
        test = mbSession.ReadString()
        If Len(test) > 1 Then
            mbSession.Write("CHAN 2  FREQ 60.0000") 'Set Frequency
            mbSession.Write("CHAN 1  FREQ 60.0000")

            mbSession.Write("CHAN 1 MODE GATE")
            mbSession.Write("CHAN 1 GATE OFF")

            mbSession.Write("CHAN 2 OUTP OFF")
            ' mbSession.Write("CHAN 2 Rang 1V;AMR 0.0020") '0.1258")
            mbSession.Write("CHAN 2 PHAS 0.00")
            mbSession.Write("CHAN 2 FNC")


            mbSession.Write("CHAN 1 OUTP OFF")
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.0020") '0.1258")
            mbSession.Write("CHAN 1 PHAS 0.00")
            mbSession.Write("CHAN 1 FNC")


        mbSession.Write("CHAN 2 OUTP ON ")
            mbSession.Write("CHAN 1 OUTP ON")
            ' mbSession.Write("CHAN 2 Rang 1V;AMR 0.0020;OFFS 0") '0.1258")
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.0020;OFFS 0") '0.1258")

            'mbSession.Write("CHAN 2 SPH +1")         'Set Phase
            'mbSession.Write("CHAN 1 SPH +6.853")
            'mbSession.Write("CHAN 2 SPH +0.4")         'Set Phase
            'mbSession.Write("CHAN 2 SPH +358.60") 'Changed November 21, 2022
            'mbSession.Write("CHAN 2 SPH +0.0200") 'Changed November 22, 2022
            mbSession.Write("CHAN 2 SPH +0.2500") 'Changed May 3rd 2023 - KW
            mbSession.Write("CHAN 1 SPH +1.0")  'Changed November 16, 2022
            mbSession.Write("CHAN 1 OFFS 0") '0.1258")
            ' mbSession.Write("CHAN 2 Rang 10V;AMR 5.208") '0.1258")
            mbSession.Write("CHAN 2 FNO 1")

            'mbSession.Write("CHAN 1 Rang 1V;AMR 0.1258") '0.1258") ' Changed on November 17, 2022 - CH
            'mbSession.Write("CHAN 1 Rang 1V;AMR 0.1313") '0.1258") ' Changed May 3rd 2023 - KW
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.1249") '0.1258")
            mbSession.Write("CHAN 1 FNO 1")
            

            Threading.Thread.Sleep(100)
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(1, 33)
        WriteDoChannel1(1, 33)
            Threading.Thread.Sleep(200)

            mbSession.Write("CHAN 1 GATE ON")
            Threading.Thread.Sleep(200)
            'Call StopRadian()
            'Threading.Thread.Sleep(500)
            'Call StopRadian()

        End If

        'Threading.Thread.Sleep(1000)

        Threading.Thread.Sleep(300)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If
        '**********************************************************************************************************************************
        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            Threading.Thread.Sleep(1000)
            RadRDInstMetric(comDevice, 0, RAD_INST_A, CURinternal)
            x = CURinternal
            Me.TextBox1.Text = x.ToString
            Me.TextBox1.BackColor = Color.LimeGreen
        End If
        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            'Threading.Thread.Sleep(1000)
            RadRDInstMetric(comDevice, 0, RAD_INST_PF, PFinternal)
            Y = PFinternal
            Me.TextBox4.Text = Y.ToString
            Me.TextBox4.BackColor = Color.LimeGreen

        End If
        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDInstMetric(comDevice, 0, RAD_INST_PF, PFinternal)
            Y = PFinternal
            Me.TextBox4.Text = Y.ToString
            Me.TextBox4.BackColor = Color.LimeGreen

        End If


        Application.DoEvents()

    End Sub
    Public Sub teststep3()
        Dim test As String
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)
        Dim x As Single
        Dim CURinternal As Single
        Dim Y As Single
        Dim PFinternal As Single

        'Call StopRadian()
        'Threading.Thread.Sleep(1000)


        Me.TextBox1.BackColor = Color.White
        Me.TextBox1.Text = ""
        TextBox5.BackColor = Color.White
        TextBox6.BackColor = Color.LimeGreen
        Application.DoEvents()
        mbSession.Write("IDEN?")
        test = mbSession.ReadString()

        If Len(test) > 1 Then
            mbSession.Write("CHAN 2  FREQ 60.0000") 'Set Frequency
            mbSession.Write("CHAN 1  FREQ 60.0000")

            mbSession.Write("CHAN 1 MODE GATE")
            mbSession.Write("CHAN 1 GATE OFF")

            mbSession.Write("CHAN 2 OUTP OFF")
            ' mbSession.Write("CHAN 2 Rang 1V;AMR 0.0020") '0.1258")
            mbSession.Write("CHAN 2 PHAS 0.00")
            mbSession.Write("CHAN 2 FNC")


            mbSession.Write("CHAN 1 OUTP OFF")
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.0020") '0.1258")
            mbSession.Write("CHAN 1 PHAS 0.00")
            mbSession.Write("CHAN 1 FNC")
            '   mbSession.Write("CHAN 1 MOD 0")


        mbSession.Write("CHAN 2 OUTP ON ")
            mbSession.Write("CHAN 1 OUTP ON")
            'mbSession.Write("CHAN 2 Rang 1V;AMR 0.0020;OFFS 0") '0.1258")
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.0020;OFFS 0") '0.1258")

            'mbSession.Write("CHAN 2 SPH +1")         'Set Phase
            'mbSession.Write("CHAN 1 SPH +55.44")
            'mbSession.Write("CHAN 2 SPH +60.4")         'Set Phase
            'mbSession.Write("CHAN 2 SPH +58.6")  'Changed November 21, 2022       'Set Phase
            'mbSession.Write("CHAN 2 SPH +60.0200")  'Changed November 22, 2022 
            mbSession.Write("CHAN 2 SPH +60.2500")  'Changed May 2, 2023 - CH 
            mbSession.Write("CHAN 1 SPH +1.0") ' Changed November 16, 2022
            mbSession.Write("CHAN 1 OFFS 0") '0.1258")
            ' mbSession.Write("CHAN 2 Rang 10V;AMR 5.208") '0.1258")
            mbSession.Write("CHAN 2 FNO 1")

            'mbSession.Write("CHAN 1 Rang 1V;AMR 0.1313") '0.1258") changed May 3rd - KW
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.1249") '0.1258")
            mbSession.Write("CHAN 1 FNO 1")


            Threading.Thread.Sleep(100)
            'AxAdvDIO1.DeviceNumber = 1
            ' AxAdvDIO1.WriteDoChannel(1, 33)
        WriteDoChannel1(1, 33)
            Threading.Thread.Sleep(200)

            mbSession.Write("CHAN 1 GATE ON")
            Threading.Thread.Sleep(200)

            'Threading.Thread.Sleep(1000)
            'Call StopRadian()
            'Threading.Thread.Sleep(500)
            'Call StopRadian()
        End If

        'Threading.Thread.Sleep(1000)

        'AxAdvDIO1.DeviceNumber = 0
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If
        '**********************************************************************************************************************************
        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            Threading.Thread.Sleep(1000)
            RadRDInstMetric(comDevice, 0, RAD_INST_A, CURinternal)
            x = CURinternal
            Me.TextBox1.Text = x.ToString
            Me.TextBox1.BackColor = Color.LimeGreen

        End If

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            'Threading.Thread.Sleep(1000)
            RadRDInstMetric(comDevice, 0, RAD_INST_PF, PFinternal)
            Y = PFinternal
            Me.TextBox4.Text = Y.ToString
            Me.TextBox4.BackColor = Color.LimeGreen

        End If

        Application.DoEvents()


    End Sub
    Public Sub teststep4()
        Dim test As String
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)
        Dim x As Single
        Dim CURinternal As Single
        Dim Y As Single
        Dim PFinternal As Single

        Me.TextBox1.BackColor = Color.White
        Me.TextBox1.Text = ""

        'Call StopRadian()
        'Threading.Thread.Sleep(1000)

        TextBox6.BackColor = Color.White
        TextBox7.BackColor = Color.LimeGreen
        Application.DoEvents()
        mbSession.Write("IDEN?")
        test = mbSession.ReadString()

        If Len(test) > 1 Then
            mbSession.Write("CHAN 2  FREQ 60.0000") 'Set Frequency
            mbSession.Write("CHAN 1  FREQ 60.0000")

            mbSession.Write("CHAN 1 MODE GATE")
            mbSession.Write("CHAN 1 GATE OFF")

            mbSession.Write("CHAN 2 OUTP OFF")
            'mbSession.Write("CHAN 2 Rang 1V;AMR 0.0020") '0.1258")
            mbSession.Write("CHAN 2 PHAS 0.00")
            mbSession.Write("CHAN 2 FNC")


            mbSession.Write("CHAN 1 OUTP OFF")
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.0020") '0.1258")
            mbSession.Write("CHAN 1 PHAS 0.00")
            mbSession.Write("CHAN 1 FNC")


        mbSession.Write("CHAN 2 OUTP ON ")
            mbSession.Write("CHAN 1 OUTP ON")
            ' mbSession.Write("CHAN 2 Rang 1V;AMR 0.0020;OFFS 0") '0.1258")
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.0020;OFFS 0") '0.1258")

            'mbSession.Write("CHAN 2 SPH +1")         'Set Phase
            'mbSession.Write("CHAN 1 SPH +55.44")
            'mbSession.Write("CHAN 2 SPH +60.4")         'Set Phase
            'mbSession.Write("CHAN 2 SPH +58.6")  'Changed November 21, 2022       'Set Phase
            'mbSession.Write("CHAN 2 SPH +60.0200")  'Changed November 21, 2022 
            mbSession.Write("CHAN 2 SPH +60.2500")  'Changed May 2, 2023 - CH 
            mbSession.Write("CHAN 1 SPH +1.0")  'Changed November 16, 2022
            mbSession.Write("CHAN 1 OFFS 0") '0.1258")
            ' mbSession.Write("CHAN 2 Rang 10V;AMR 5.208") '0.1258")
            mbSession.Write("CHAN 2 FNO 1")

            'mbSession.Write("CHAN 1 Rang 1V;AMR 0.1313") '0.1258") Changed may 3rd - KW
            mbSession.Write("CHAN 1 Rang 1V;AMR 0.1249") '0.1258")
            mbSession.Write("CHAN 1 FNO 1")
 

            Threading.Thread.Sleep(100)
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(1, 33)
        WriteDoChannel1(1, 33)
            Threading.Thread.Sleep(200)

            mbSession.Write("CHAN 1 GATE ON")
            Threading.Thread.Sleep(200)
            'Threading.Thread.Sleep(1000)
            'Call StopRadian()
            'Threading.Thread.Sleep(500)
            'Call StopRadian()

        End If

        'Threading.Thread.Sleep(1000)
        Threading.Thread.Sleep(500)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If
        '**********************************************************************************************************************************
        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            Threading.Thread.Sleep(1000)
            RadRDInstMetric(comDevice, 0, RAD_INST_A, CURinternal)
            x = CURinternal
            Me.TextBox1.Text = x.ToString
            Me.TextBox1.BackColor = Color.LimeGreen

        End If
        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            'Threading.Thread.Sleep(1000)
            RadRDInstMetric(comDevice, 0, RAD_INST_PF, PFinternal)
            Y = PFinternal
            Me.TextBox4.Text = Y.ToString
            Me.TextBox4.BackColor = Color.LimeGreen

        End If
        Application.DoEvents()

    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox

    Private Sub SetChan2(ByVal voltage As String)
        ' Private Sub Routine to Set the Voltage on Channel 2
        If voltage = "120" Then
            '*120 to PT1
            WriteDoChannel0(1, 22)
            WriteDoChannel0(0, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.1952") ''bench2 Old setting
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.1992") ''bench2 changed to this on November 11, 2022 by Colin Hughes as per details from Kyle Wagner
            mbSession.Write("CHAN 2 Rang 10V AMR 5.2072") ''bench2 changed to this on May 2, 2023 by Colin Hughes as per details from Kyle Wagner
        End If

        If voltage = "208" Then
            '*208 to PT1
            WriteDoChannel0(0, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.0015") ''Bench2 Old Setting
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.0038") ''Bench2 Changed to this on November 11, 2022 by Colin Hughes as per details from Kyle Wagner
            mbSession.Write("CHAN 2 Rang 10V AMR 3.0085") ''Bench2 Changed to this on May 2, 2023 by Colin Hughes as per details from Kyle Wagner
        End If

        If voltage = "240" Then
            '*240 to PT1
            WriteDoChannel0(0, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.4635") 'bench2 Old Settings
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.4661") 'bench2 Changed to this on November 11, 2022 by Colin Hughes as per details from Kyle Wagner
            mbSession.Write("CHAN 2 Rang 10V AMR 3.4715") 'bench2 Changed to this on May 2, 2023 by Colin Hughes as per details from Kyle Wagner
        End If

        If voltage = "277" Then
            '*277 to PT1
            WriteDoChannel0(0, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.9974") ''bench2 old Settings
            'mbSession.Write("CHAN 2 Rang 10V AMR 4.0005") ''bench2 Changed to this on November 11, 2022 by Colin Hughes as per details from Kyle Wagner
            mbSession.Write("CHAN 2 Rang 10V AMR 4.0066") ''bench2 Changed to this on May 2, 2023 by Colin Hughes as per details from Kyle Wagner
        End If


        If voltage = "347" Then
            '*347 to PT1
            WriteDoChannel0(0, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.0076") ''bench2 Old Settings
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.0115") ''bench2 Changed to this on November 11, 2022 by Colin Hughes as per details from Kyle Wagner
            mbSession.Write("CHAN 2 Rang 10V AMR 5.0192") ''bench2 Changed to this on May 2, 2023 by Colin Hughes as per details from Kyle Wagner
        End If
        If voltage = "416" Then
            '*416 to PT1
            WriteDoChannel0(1, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.6020") ''Bench2 Old Settings
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.6048") ''Bench2 Changed to this on November 11, 2022 by Colin Hughes as per details from Kyle Wagner
            mbSession.Write("CHAN 2 Rang 10V AMR 3.6103") ''Bench2 Changed to this on May 2, 2023 by Colin Hughes as per details from Kyle Wagner
        End If
        If voltage = "480" Then
            '*480 to PT1
            WriteDoChannel0(1, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 4.1562 ") ''Bench2 Old Settings
            'mbSession.Write("CHAN 2 Rang 10V AMR 4.1594 ") ''Bench2 Changed to this on November 11, 2022 by Colin Hughes as per details from Kyle Wagner
            mbSession.Write("CHAN 2 Rang 10V AMR 4.1658") ''Bench2 Changed to this on May 2, 2023 by Colin Hughes as per details from Kyle Wagner

        End If
        If voltage = "600" Then
            '*600 to PT1
            WriteDoChannel0(1, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.1952 ") ''Bench2 Old Settings
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.1992 ") ''Bench2 Changed to this on November 11, 2022 by Colin Hughes as per details from Kyle Wagner
            mbSession.Write("CHAN 2 Rang 10V AMR 5.2072") ''Bench2 Changed to this on May 2, 2023 by Colin Hughes as per details from Kyle Wagner
        End If
    End Sub

    Private Function ReadAllTextFromINI(vtap As String) As Double
        Dim appPath As String = System.IO.Directory.GetCurrentDirectory()
        ' Reader to read from the file
        Dim sr As New System.IO.StreamReader(appPath + "\\MA_Settings.ini")
        ' Hold the amount of lines already read in a 'counter-variable' 
        Dim placeholder As Integer = 0
        Do While sr.Peek <> -1 ' Is -1 when no data exists on the next line of the CSV file
            Dim readData As String() = sr.ReadLine.Split("=")
            If readData(0) = vtap Then
                Return Double.Parse(readData(1).ToString())
            End If
        Loop
        ' NOTE: If you wish to parse the data inside the array,
        ' loop through all elements, use the String.Split() method
        ' delimited by a comma (for CSV).

        Return 0
    End Function


    Public Sub SetAfterPrn_Daily_Accuracy_Check()
        Dim WHinternal As Single
        Dim WHexternal As Single
        Dim IntCount(256) As Long
        Dim cbox1 As String
        Dim voltage As String
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)
        On Error Resume Next

        Call SetTxtBoxtoEmpty()

        mbSession = CType(ResourceManager.GetLocalManager().Open("GPIB0::0::INSTR"), MessageBasedSession)
        mbSession.Write("CHAN 1 MODE GATE")
        mbSession.Write("CHAN 1 GATE OFF")
        voltage = TextBox2.Text
        Call SetChan2(voltage)
        Threading.Thread.Sleep(1000)
        Call teststep1()

        '88888888888888888888888888888888888888888888888888888888888888888888 After Step 1 8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If


        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================
        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(150)
        End While
        Threading.Thread.Sleep(150)
        Application.DoEvents()

        Me.BackgroundWorker2.RunWorkerAsync()
        Threading.Thread.Sleep(150)

        Form1.Form1_CallScans(5)
        Thread.Sleep(500)


        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        Threading.Thread.Sleep(150)
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If


        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999

        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If

        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        'Threading.Thread.Sleep(250)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(120)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        Thread.Sleep(100)
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")
        TextBox52.Text = ""
        TextBox52.BackColor = Color.White
        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            Threading.Thread.Sleep(150)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            Threading.Thread.Sleep(150)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If



        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.

        'Dim LastRow As Long
        'With oSheet
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row


        '    '''''''''''''''''''' Top Row Only'''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    If .cells(1, 1).Value = "" Then


        '        .cells(1, 1).Value = "TAP"
        '        .cells(1, 2).Value = "Console Radian"
        '        .cells(1, 3).Value = "MC Radian"
        '        .cells(1, 4).Value = "Current Step"
        '        .cells(1, 5).Value = "V Mult"
        '        .cells(1, 6).Value = "Rcons"
        '        .cells(1, 7).Value = "C Mult"
        '        .cells(1, 8).Value = "Rmc"
        '        .cells(1, 9).Value = "% Err"

        '    End If

        '    '''''''''''''''''''''''''''''''''CT and Votage '''''''''''''''''''''''''''''''''''''''''''''''''''
        '    '''''''add CT here
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    .cells(LastRow + 1, 1).value = "CT"
        '    .cells(LastRow + 1, 2).Value = voltage



        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row

        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "WH @ 2.5% Unity"

        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage

        '    ' Formula => (Rcons) = (CT) * (V mult)
        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value

        '    ' Setting C Mult = 1000
        '    .cells(LastRow + 1, 7).Value = 1000

        '    ' Formula => (Rmc) = (C Mult) * (MC Radian)
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value

        '    ' Formula  = ((Rmc - Rcons)/Rmc) * 100
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100

        '    TextBox49.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox49.BackColor = Color.Red
        '    Else
        '        TextBox49.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        '    Application.DoEvents()
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With
        '88888888888888888888888888888888888888888888888888888888888888 step 1 error calc 88888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
        If ACCTESTFLAG = 1 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""
            
            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double


            Dim Error_Calc_Meter As Double = 0
            Dim Error_Calc_Corrected_Meter As Double = 0
            Dim Scaled_WattHourMeter As Double = 0
            Dim Scaled_WattHourConsole As Double = 0
            Dim Accuracy_Test_Results As Double = 0
            Dim rnd_Accuracy_Test_Results As Double = 0


            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[WH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(117) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 4)
                        WH100array(ii) = strFileName(iii)
                        Dim fstring As String = "1.25_1.0_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Radian = Math.Round(CDbl(WHinternal), 3)
            Dim Percent2 As String

            Dim tuple_CalcResults As Tuple(Of Double, Double, Double, Double, Double, Double)
            tuple_CalcResults = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResults.Item1
            Scaled_WattHourMeter = tuple_CalcResults.Item2
            Scaled_WattHourConsole = tuple_CalcResults.Item3
            Error_Calc_Meter = tuple_CalcResults.Item4
            Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteDAC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                     , Error_Calc_Meter, Accuracy_Test_Results, voltage, "WH @ 2.5% Unity", oSheet, 1)

            tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)


            Call DAC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
                                     , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
                                     , rnd_Accuracy_Test_Results.ToString, 1)
            ' Write results to DB

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            external = WHexternal.ToString
            Load = "1.25"
            powerfactor = "1.0"
            Unit = "WH"


            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If

            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If

            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()






        End If

        Call teststep2() ' Step 2 Begins

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(200)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If

        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If
        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(100)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While

        BackgroundWorker2.RunWorkerAsync()

        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        Thread.Sleep(100)
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")
        'MsgBox("set Yokogawa Voltage , Current, Phase   OFF")
        'MsgBox("Read MC MKA to excel")
        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.
100:

        'With oSheet
        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "WH @ 25% Unity"

        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage



        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
        '    .cells(LastRow + 1, 7).Value = 1000
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
        '    TextBox50.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox50.BackColor = Color.Red
        '    Else
        '        TextBox50.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        '    Application.DoEvents()
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With


        '88888888888888888888888888888888888888888888888888888888888888 step 2 error calc 88888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
        If ACCTESTFLAG = 1 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""

            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double


            Dim Error_Calc_Meter As Double = 0
            Dim Error_Calc_Corrected_Meter As Double = 0
            Dim Scaled_WattHourMeter As Double = 0
            Dim Scaled_WattHourConsole As Double = 0
            Dim Accuracy_Test_Results As Double = 0
            Dim rnd_Accuracy_Test_Results As Double = 0


            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[WH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(116) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 4)
                        WH100array(ii) = strFileName(iii)
                        'Dim fstring As String = "1.25_1.0_M" & numtap
                        Dim fstring As String = "12.5_1.0_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Radian = Math.Round(CDbl(WHinternal), 3)
            Dim Percent2 As String

            Dim tuple_CalcResults As Tuple(Of Double, Double, Double, Double, Double, Double)
            tuple_CalcResults = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResults.Item1
            Scaled_WattHourMeter = tuple_CalcResults.Item2
            Scaled_WattHourConsole = tuple_CalcResults.Item3
            Error_Calc_Meter = tuple_CalcResults.Item4
            Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteDAC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                , Error_Calc_Meter, Accuracy_Test_Results, voltage, "WH @ 25% Unity", oSheet, 2)

            tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)



            Call DAC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
                                     , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
                                     , rnd_Accuracy_Test_Results.ToString, 2)
            ' Write results to DB

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            external = WHexternal.ToString
            Load = "12.5"
            powerfactor = "1.0"
            Unit = "WH"


            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If

            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If

            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()

        End If

DQDQ2:
        ' *************************** DAC Step 3 Testing **********************
        Call teststep3()

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(110)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999


        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If
        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(100)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(120)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")

        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.


        'With oSheet

        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "WH @ 25% @ PF"


        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage



        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
        '    .cells(LastRow + 1, 7).Value = 1000
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
        '    TextBox51.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox51.BackColor = Color.Red
        '    Else
        '        TextBox51.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        '    Application.DoEvents()
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With

        '************************************** DAC Step3 Error Calculation ********************************************
        If ACCTESTFLAG = 1 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""

            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double


            Dim Error_Calc_Meter As Double = 0
            Dim Error_Calc_Corrected_Meter As Double = 0
            Dim Scaled_WattHourMeter As Double = 0
            Dim Scaled_WattHourConsole As Double = 0
            Dim Accuracy_Test_Results As Double = 0
            Dim rnd_Accuracy_Test_Results As Double = 0


            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[WH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(116) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 4)
                        WH100array(ii) = strFileName(iii)
                        Dim fstring As String = "12.5_0.5_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Radian = Math.Round(CDbl(WHinternal), 3)
            Dim Percent2 As String

            Dim tuple_CalcResults As Tuple(Of Double, Double, Double, Double, Double, Double)
            tuple_CalcResults = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResults.Item1
            Scaled_WattHourMeter = tuple_CalcResults.Item2
            Scaled_WattHourConsole = tuple_CalcResults.Item3
            Error_Calc_Meter = tuple_CalcResults.Item4
            Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteDAC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                , Error_Calc_Meter, Accuracy_Test_Results, voltage, "WH @ 25% @ PF", oSheet, 3)


            tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)


            Call DAC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
                                     , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
                                     , rnd_Accuracy_Test_Results.ToString, 3)
            ' Write results to DB

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            external = WHexternal.ToString
            Load = "12.5"
            powerfactor = "0.5"
            Unit = "WH"


            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If

            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If

            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()

        End If

        ' ******************************** DAC Step4 Test ***************************************
        Call teststep4()

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(200)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If
        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999

        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If

        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(250)
        ComboBox1.Text = ""
        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")

        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_VAH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_VAH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.

        'With oSheet
        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "VAH @ 25% @ PF"

        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage

        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
        '    .cells(LastRow + 1, 7).Value = 1000
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
        '    TextBox52.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox52.BackColor = Color.Red
        '    Else
        '        TextBox52.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        '    Application.DoEvents()
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With
        '************************************** DAC Step4 Error Calculation ********************************************
        If ACCTESTFLAG = 1 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""

            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double


            Dim Error_Calc_Meter As Double = 0
            Dim Error_Calc_Corrected_Meter As Double = 0
            Dim Scaled_WattHourMeter As Double = 0
            Dim Scaled_WattHourConsole As Double = 0
            Dim Accuracy_Test_Results As Double = 0
            Dim rnd_Accuracy_Test_Results As Double = 0


            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[VAH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(40) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 2)
                        WH100array(ii) = strFileName(iii)
                        Dim fstring As String = "12.5_0.5_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Radian = Math.Round(CDbl(WHinternal), 3)
            Dim Percent2 As String

            Dim tuple_CalcResults As Tuple(Of Double, Double, Double, Double, Double, Double)
            tuple_CalcResults = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResults.Item1
            Scaled_WattHourMeter = tuple_CalcResults.Item2
            Scaled_WattHourConsole = tuple_CalcResults.Item3
            Error_Calc_Meter = tuple_CalcResults.Item4
            Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteDAC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
              , Error_Calc_Meter, Accuracy_Test_Results, voltage, "VAH @ 25% @ PF", oSheet, 4)

            tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)


            ' **** Update Accuracy Check Window *************
            Call DAC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
                                     , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
                                     , rnd_Accuracy_Test_Results.ToString, 4)
            ' Write results to DB

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            external = WHexternal.ToString
            Load = "12.5"
            powerfactor = "0.5"
            Unit = "Vah"


            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If

            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If

            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()

        End If
        If ACCTESTFLAG = 1 And RadioButton40.Checked = True And TextBox2.Text = "600" Then
            Call Button2_Click(0, System.EventArgs.Empty)
            'Exit Sub

        End If


    End Sub

    Public Sub WriteMC_toExcel(ByVal tapNumber As Integer, ByVal WHExternal As Double, ByVal WHInternal As Double, ByVal Scaled_WHInternal As Double _
                                , ByVal Scaled_WHExternal As Double, ByVal ConsoleError As Double, ByVal Multiplier As Double, ByVal ErrorbeforeCorrection As Double _
                                , ByVal voltage As String, ByVal Current_Step As String, ByRef oSheet As Object, ByVal testNumber As Integer)
        Dim LastRow As Long
        With oSheet
            LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row


            '''''''''''''''''''' Top Row Only'''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .cells(1, 1).Value = "" Then


                .cells(1, 1).Value = "TAP"
                .cells(1, 2).Value = "Console Radian"
                .cells(1, 3).Value = "MC Radian"
                .cells(1, 4).Value = "Current Step"
                .cells(1, 5).Value = "V Mult"
                .cells(1, 6).Value = "Rcons"
                .cells(1, 7).Value = "C Mult"
                .cells(1, 8).Value = "Rmc"
                .cells(1, 9).Value = "% Err"

            End If
            '''''''''''''''''''''''''''''''''CT and Votage '''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''add CT here
            If (testNumber = 1) Then
                LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
                .cells(LastRow + 1, 1).value = "CT"
                .cells(LastRow + 1, 2).Value = voltage
            End If

            '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
            LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row

            .cells(LastRow + 1, 1).Value = tapNumber
            .cells(LastRow + 1, 2).Value = WHInternal
            .cells(LastRow + 1, 3).Value = WHExternal
            .cells(LastRow + 1, 4).Value = Current_Step
            .cells(LastRow + 1, 5).Value = Multiplier
            .cells(LastRow + 1, 6).Value = Scaled_WHInternal
            .cells(LastRow + 1, 7).Value = 1000 ' MilliAmpp Multiplier
            .cells(LastRow + 1, 8).Value = Scaled_WHExternal
            .cells(LastRow + 1, 9).Value = ErrorbeforeCorrection ' Has console error correction applied to it

            If (testNumber = 1) Then
                TextBox49.Text = .cells(LastRow + 1, 9).Text
                If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                    TextBox49.BackColor = Color.Red
                Else
                    TextBox49.BackColor = Color.Lime
                End If
            ElseIf (testNumber = 2) Then
                TextBox50.Text = .cells(LastRow + 1, 9).Text
                If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                    TextBox50.BackColor = Color.Red
                Else
                    TextBox50.BackColor = Color.Lime
                End If
            ElseIf (testNumber = 3) Then
                TextBox51.Text = .cells(LastRow + 1, 9).Text
                If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                    TextBox51.BackColor = Color.Red
                Else
                    TextBox51.BackColor = Color.Lime
                End If
            ElseIf (testNumber = 4) Then
                TextBox52.Text = .cells(LastRow + 1, 9).Text
                If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                    TextBox52.BackColor = Color.Red
                Else
                    TextBox52.BackColor = Color.Lime
                End If
            End If


            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

            Application.DoEvents()
        End With

    End Sub

    Public Sub WriteDAC_toExcel(ByVal tapNumber As Integer, ByVal WHExternal As Double, ByVal WHInternal As Double, ByVal Scaled_WHInternal As Double _
                               , ByVal Scaled_WHExternal As Double, ByVal ConsoleError As Double, ByVal Multiplier As Double, ByVal ErrorbeforeCorrection As Double _
                               , ByVal Corrected_Error_fromConsole As Double, ByVal voltage As String, ByVal Current_Step As String, ByRef oSheet As Object, ByVal testNumber As Integer)
        Dim LastRow As Long
        With oSheet
            LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row


            '''''''''''''''''''' Top Row Only'''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .cells(1, 1).Value = "" Then


                .cells(1, 1).Value = "TAP"
                .cells(1, 2).Value = "Console Radian"
                .cells(1, 3).Value = "MC Radian"
                .cells(1, 4).Value = "Current Step"
                .cells(1, 5).Value = "V Mult"
                .cells(1, 6).Value = "Rcons"
                .cells(1, 7).Value = "C Mult"
                .cells(1, 8).Value = "Rmc"
                .cells(1, 9).Value = "% Err"

            End If
            '''''''''''''''''''''''''''''''''CT and Votage '''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''add CT here
            If (testNumber = 1) Then
                LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
                .cells(LastRow + 1, 1).value = "CT"
                .cells(LastRow + 1, 2).Value = voltage
            End If
            
            '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
            LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row

            .cells(LastRow + 1, 1).Value = tapNumber
            .cells(LastRow + 1, 2).Value = WHInternal
            .cells(LastRow + 1, 3).Value = WHExternal
            .cells(LastRow + 1, 4).Value = Current_Step
            .cells(LastRow + 1, 5).Value = Multiplier
            .cells(LastRow + 1, 6).Value = Scaled_WHInternal
            .cells(LastRow + 1, 7).Value = 1000 ' MilliAmpp Multiplier
            .cells(LastRow + 1, 8).Value = Scaled_WHExternal
            .cells(LastRow + 1, 9).Value = Corrected_Error_fromConsole ' Has console error correction applied to it


            If (testNumber = 1) Then
                TextBox49.Text = .cells(LastRow + 1, 9).Text
                If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                    TextBox49.BackColor = Color.Red
                Else
                    TextBox49.BackColor = Color.Lime
                End If
            ElseIf (testNumber = 2) Then
                TextBox50.Text = .cells(LastRow + 1, 9).Text
                If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                    TextBox50.BackColor = Color.Red
                Else
                    TextBox50.BackColor = Color.Lime
                End If
            ElseIf (testNumber = 3) Then
                TextBox51.Text = .cells(LastRow + 1, 9).Text
                If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                    TextBox51.BackColor = Color.Red
                Else
                    TextBox51.BackColor = Color.Lime
                End If
            ElseIf (testNumber = 4) Then
                TextBox52.Text = .cells(LastRow + 1, 9).Text
                If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                    TextBox52.BackColor = Color.Red
                Else
                    TextBox52.BackColor = Color.Lime
                End If
            End If


            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

            Application.DoEvents()
        End With

    End Sub

    


    Public Sub SetAfterPrn_MC_Accuracy_Check()
        strFileName = IO.File.ReadAllLines("C:\Bench2\Console\mmc_err.INI") '// add each line as String Array.
        Dim WHinternal As Single
        Dim WHexternal As Single
        Dim IntCount(256) As Long
        Dim cbox1 As String
        Dim voltage As String
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)
        On Error Resume Next

        Call SetTxtBoxtoEmpty()

        mbSession = CType(ResourceManager.GetLocalManager().Open("GPIB0::0::INSTR"), MessageBasedSession)
        mbSession.Write("CHAN 1 MODE GATE")
        mbSession.Write("CHAN 1 GATE OFF")
        voltage = TextBox2.Text
        Call SetChan2(voltage)

        Call teststep1()

        '88888888888888888888888888888888888888888888888888888888888888888888 After Step 1 8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
        '==================================================stop accumulation clear both radians==========================================================================
        Threading.Thread.Sleep(2000)
        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            ' RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If


        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================
        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(150)
        End While

        Application.DoEvents()

        Me.BackgroundWorker2.RunWorkerAsync()
        Threading.Thread.Sleep(150)

        Form1.Form1_CallScans(5)
        Thread.Sleep(500)


        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If

        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If


        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999

        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If

        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        'Threading.Thread.Sleep(250)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(60)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        Thread.Sleep(100)
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")
        TextBox52.Text = ""
        TextBox52.BackColor = Color.White
        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''
        Threading.Thread.Sleep(2000)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If



        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.

        'Dim LastRow As Long
        'With oSheet
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row


        '    '''''''''''''''''''' Top Row Only'''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    If .cells(1, 1).Value = "" Then


        '        .cells(1, 1).Value = "TAP"
        '        .cells(1, 2).Value = "Console Radian"
        '        .cells(1, 3).Value = "MC Radian"
        '        .cells(1, 4).Value = "Current Step"
        '        .cells(1, 5).Value = "V Mult"
        '        .cells(1, 6).Value = "Rcons"
        '        .cells(1, 7).Value = "C Mult"
        '        .cells(1, 8).Value = "Rmc"
        '        .cells(1, 9).Value = "% Err"

        '    End If

        '    '''''''''''''''''''''''''''''''''CT and Votage '''''''''''''''''''''''''''''''''''''''''''''''''''
        '    '''''''add CT here
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    .cells(LastRow + 1, 1).value = "CT"
        '    .cells(LastRow + 1, 2).Value = voltage



        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row

        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "WH @ 2.5% Unity"

        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage

        '    ' Formula => (Rcons) = (CT) * (V mult)
        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value

        '    ' Setting C Mult = 1000
        '    .cells(LastRow + 1, 7).Value = 1000

        '    ' Formula => (Rmc) = (C Mult) * (MC Radian)
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value

        '    ' Formula  = ((Rmc - Rcons)/Rmc) * 100
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100

        '    TextBox49.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox49.BackColor = Color.Red
        '    Else
        '        TextBox49.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        '    Application.DoEvents()
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With
        '88888888888888888888888888888888888888888888888888888888888888 step 1 error calc 88888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
        If ACCTESTFLAG = 0 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""

            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double


            Dim Error_Calc_Meter As Double = 0
            Dim Error_Calc_Corrected_Meter As Double = 0
            Dim Scaled_WattHourMeter As Double = 0
            Dim Scaled_WattHourConsole As Double = 0
            Dim Accuracy_Test_Results As Double = 0
            Dim rnd_Accuracy_Test_Results As Double = 0


            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[WH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(117) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 4)
                        WH100array(ii) = strFileName(iii)
                        Dim fstring As String = "1.25_1.0_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Radian = Math.Round(CDbl(WHinternal), 3)
            Dim Percent2 As String

            Dim tuple_CalcResults As Tuple(Of Double, Double, Double, Double)
            'tuple_CalcResults = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            tuple_CalcResults = AccuracyCheckCalcFunctions.MC_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResults.Item1
            Scaled_WattHourMeter = tuple_CalcResults.Item2
            Scaled_WattHourConsole = tuple_CalcResults.Item3
            Error_Calc_Meter = tuple_CalcResults.Item4
            'Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            'Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteMC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(Estd), CDbl(mulitplier) _
                                 , Error_Calc_Meter, voltage, "WH @ 2.5% Unity", oSheet, 1)

            'tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)



            'Call MC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
            '                         , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
            '                         , rnd_Accuracy_Test_Results.ToString)

            'Call DAC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
            '                         , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
            '                         , rnd_Accuracy_Test_Results.ToString)
            ' Write results to DB

            'Dim myconnection As New ADODB.Connection
            'Dim mycommand As New ADODB.Command
            'Dim ra As Integer
            'Dim Load As String
            'Dim powerfactor As String
            'Dim connt As ADODB.Connection
            'Dim connectionString As String
            'Dim external As String
            'Dim Recset As New ADODB.Recordset
            'Dim Recset1 As New ADODB.Recordset
            'Dim Recset2 As New ADODB.Recordset
            'Dim Mdate As String
            'Dim mdDate As DateTime
            'Dim Unit As String = ""
            'external = WHexternal.ToString
            'Load = "1.25"
            'powerfactor = "1.0"
            'Unit = "WH"


            'If Not Button13.BackColor = Color.Green Then
            '    Vtap = CDbl(numtap) + 1
            'Else
            '    Vtap = CDbl(numtap)

            'End If

            'Vtap = Vtap.ToString
            'If Len(Vtap) < 2 Then
            '    Vtap = "M0" & Vtap
            'Else
            '    Vtap = "M" & Vtap
            'End If

            'myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            'myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
            '           "'" & Unit & "', " & _
            '           "'" & outp & "', " & _
            '           "'" & Load & "', " & _
            '           "'" & powerfactor & "', " & _
            '           "'" & Vtap & "', " & _
            '           "'" & _Error & "'," & _
            '           "'" & WHexternal & "'," & _
            '           "'" & WHinternal & "'," & _
            '           "'" & opertr & "'," & _
            '           "'" & tdate & "'" & _
            '           ")")

            'Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            'If Not Uniq_ID_Flag > 0 Then
            '    If Not Recset.EOF Then
            '        Mdate = Recset.GetString
            '        mdDate = DateTime.Parse(Mdate)
            '        myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
            '                                            "'" & mdDate & "'" & _
            '                  "))")

            '        Uniq_ID_Flag = 1
            '    End If
            '    Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

            '    If Not Recset1.EOF Then
            '        UniqID_text = Recset1.GetString
            '        UniqID_text = CInt(UniqID_text)
            '        myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


            '    End If

            'Else

            '    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            'End If

            'myconnection.Close()






        End If

        Call teststep2() ' Step 2 Begins

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(200)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If

        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If
        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(100)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While

        BackgroundWorker2.RunWorkerAsync()

        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        Thread.Sleep(100)
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")
        'MsgBox("set Yokogawa Voltage , Current, Phase   OFF")
        'MsgBox("Read MC MKA to excel")
        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.
100:
        'Dim LastRow As Long
        'With oSheet
        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "WH @ 25% Unity"

        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage



        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
        '    .cells(LastRow + 1, 7).Value = 1000
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
        '    TextBox50.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox50.BackColor = Color.Red
        '    Else
        '        TextBox50.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        '    Application.DoEvents()
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With


        '88888888888888888888888888888888888888888888888888888888888888 step 2 error calc 88888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
        If ACCTESTFLAG = 0 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""

            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double


            Dim Error_Calc_Meter As Double = 0
            Dim Error_Calc_Corrected_Meter As Double = 0
            Dim Scaled_WattHourMeter As Double = 0
            Dim Scaled_WattHourConsole As Double = 0
            Dim Accuracy_Test_Results As Double = 0
            Dim rnd_Accuracy_Test_Results As Double = 0


            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[WH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(116) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 4)
                        WH100array(ii) = strFileName(iii)
                        'Dim fstring As String = "1.25_1.0_M" & numtap
                        Dim fstring As String = "12.5_1.0_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Radian = Math.Round(CDbl(WHinternal), 3)
            Dim Percent2 As String

            Dim tuple_CalcResults As Tuple(Of Double, Double, Double, Double)
            tuple_CalcResults = AccuracyCheckCalcFunctions.MC_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResults.Item1
            Scaled_WattHourMeter = tuple_CalcResults.Item2
            Scaled_WattHourConsole = tuple_CalcResults.Item3
            Error_Calc_Meter = tuple_CalcResults.Item4
            'Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            'Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteMC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                     , Error_Calc_Meter, voltage, "WH @ 25% Unity", oSheet, 2)

            'tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)



            'Call MC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
            '                         , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
            '                         , rnd_Accuracy_Test_Results.ToString)
            '' Write results to DB

            'Dim myconnection As New ADODB.Connection
            'Dim mycommand As New ADODB.Command
            'Dim ra As Integer
            'Dim Load As String
            'Dim powerfactor As String
            'Dim connt As ADODB.Connection
            'Dim connectionString As String
            'Dim external As String
            'Dim Recset As New ADODB.Recordset
            'Dim Recset1 As New ADODB.Recordset
            'Dim Recset2 As New ADODB.Recordset
            'Dim Mdate As String
            'Dim mdDate As DateTime
            'Dim Unit As String = ""
            'external = WHexternal.ToString
            'Load = "12.5"
            'powerfactor = "1.0"
            'Unit = "WH"


            'If Not Button13.BackColor = Color.Green Then
            '    Vtap = CDbl(numtap) + 1
            'Else
            '    Vtap = CDbl(numtap)

            'End If

            'Vtap = Vtap.ToString
            'If Len(Vtap) < 2 Then
            '    Vtap = "M0" & Vtap
            'Else
            '    Vtap = "M" & Vtap
            'End If

            'myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            'myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
            '           "'" & Unit & "', " & _
            '           "'" & outp & "', " & _
            '           "'" & Load & "', " & _
            '           "'" & powerfactor & "', " & _
            '           "'" & Vtap & "', " & _
            '           "'" & _Error & "'," & _
            '           "'" & WHexternal & "'," & _
            '           "'" & WHinternal & "'," & _
            '           "'" & opertr & "'," & _
            '           "'" & tdate & "'" & _
            '           ")")

            'Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            'If Not Uniq_ID_Flag > 0 Then
            '    If Not Recset.EOF Then
            '        Mdate = Recset.GetString
            '        mdDate = DateTime.Parse(Mdate)
            '        myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
            '                                            "'" & mdDate & "'" & _
            '                  "))")

            '        Uniq_ID_Flag = 1
            '    End If
            '    Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

            '    If Not Recset1.EOF Then
            '        UniqID_text = Recset1.GetString
            '        UniqID_text = CInt(UniqID_text)
            '        myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


            '    End If

            'Else

            '    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            'End If

            'myconnection.Close()

        End If

DQDQ2:
        ' *************************** DAC Step 3 Testing **********************
        Call teststep3()

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(110)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999


        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If
        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(100)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")

        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.


        'With oSheet

        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "WH @ 25% @ PF"


        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage



        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
        '    .cells(LastRow + 1, 7).Value = 1000
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
        '    TextBox51.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox51.BackColor = Color.Red
        '    Else
        '        TextBox51.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        '    Application.DoEvents()
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With

        '************************************** DAC Step3 Error Calculation ********************************************
        If ACCTESTFLAG = 0 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""

            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double


            Dim Error_Calc_Meter As Double = 0
            Dim Error_Calc_Corrected_Meter As Double = 0
            Dim Scaled_WattHourMeter As Double = 0
            Dim Scaled_WattHourConsole As Double = 0
            Dim Accuracy_Test_Results As Double = 0
            Dim rnd_Accuracy_Test_Results As Double = 0


            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[WH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(116) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 4)
                        WH100array(ii) = strFileName(iii)
                        Dim fstring As String = "12.5_0.5_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Radian = Math.Round(CDbl(WHinternal), 3)
            Dim Percent2 As String

            Dim tuple_CalcResults As Tuple(Of Double, Double, Double, Double)
            'tuple_CalcResults = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            tuple_CalcResults = AccuracyCheckCalcFunctions.MC_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            rnd_Accuracy_Test_Results = tuple_CalcResults.Item1
            Scaled_WattHourMeter = tuple_CalcResults.Item2
            Scaled_WattHourConsole = tuple_CalcResults.Item3
            Error_Calc_Meter = tuple_CalcResults.Item4
            'Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            'Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteMC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                     , Error_Calc_Meter, voltage, "WH @ 25% @ PF", oSheet, 3)

            'tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)


            'Call MC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
            '                         , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
            '                         , rnd_Accuracy_Test_Results.ToString)
            '' Write results to DB

            'Dim myconnection As New ADODB.Connection
            'Dim mycommand As New ADODB.Command
            'Dim ra As Integer
            'Dim Load As String
            'Dim powerfactor As String
            'Dim connt As ADODB.Connection
            'Dim connectionString As String
            'Dim external As String
            'Dim Recset As New ADODB.Recordset
            'Dim Recset1 As New ADODB.Recordset
            'Dim Recset2 As New ADODB.Recordset
            'Dim Mdate As String
            'Dim mdDate As DateTime
            'Dim Unit As String = ""
            'external = WHexternal.ToString
            'Load = "12.5"
            'powerfactor = "0.5"
            'Unit = "WH"


            'If Not Button13.BackColor = Color.Green Then
            '    Vtap = CDbl(numtap) + 1
            'Else
            '    Vtap = CDbl(numtap)

            'End If

            'Vtap = Vtap.ToString
            'If Len(Vtap) < 2 Then
            '    Vtap = "M0" & Vtap
            'Else
            '    Vtap = "M" & Vtap
            'End If

            'myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            'myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
            '           "'" & Unit & "', " & _
            '           "'" & outp & "', " & _
            '           "'" & Load & "', " & _
            '           "'" & powerfactor & "', " & _
            '           "'" & Vtap & "', " & _
            '           "'" & _Error & "'," & _
            '           "'" & WHexternal & "'," & _
            '           "'" & WHinternal & "'," & _
            '           "'" & opertr & "'," & _
            '           "'" & tdate & "'" & _
            '           ")")

            'Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            'If Not Uniq_ID_Flag > 0 Then
            '    If Not Recset.EOF Then
            '        Mdate = Recset.GetString
            '        mdDate = DateTime.Parse(Mdate)
            '        myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
            '                                            "'" & mdDate & "'" & _
            '                  "))")

            '        Uniq_ID_Flag = 1
            '    End If
            '    Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

            '    If Not Recset1.EOF Then
            '        UniqID_text = Recset1.GetString
            '        UniqID_text = CInt(UniqID_text)
            '        myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


            '    End If

            'Else

            '    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            'End If

            'myconnection.Close()

        End If

        ' ******************************** DAC Step4 Test ***************************************
        Call teststep4()

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(200)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If
        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999

        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If

        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(250)
        ComboBox1.Text = ""
        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")

        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_VAH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_VAH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.

        'With oSheet
        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "VAH @ 25% @ PF"

        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage

        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
        '    .cells(LastRow + 1, 7).Value = 1000
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
        '    TextBox52.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox52.BackColor = Color.Red
        '    Else
        '        TextBox52.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        '    Application.DoEvents()
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With
        '************************************** DAC Step4 Error Calculation ********************************************
        If ACCTESTFLAG = 0 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""

            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double


            Dim Error_Calc_Meter As Double = 0
            Dim Error_Calc_Corrected_Meter As Double = 0
            Dim Scaled_WattHourMeter As Double = 0
            Dim Scaled_WattHourConsole As Double = 0
            Dim Accuracy_Test_Results As Double = 0
            Dim rnd_Accuracy_Test_Results As Double = 0


            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[VAH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(40) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 2)
                        WH100array(ii) = strFileName(iii)
                        Dim fstring As String = "12.5_0.5_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Radian = Math.Round(CDbl(WHinternal), 3)
            Dim Percent2 As String

            Dim tuple_CalcResults As Tuple(Of Double, Double, Double, Double)
            'tuple_CalcResults = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            tuple_CalcResults = AccuracyCheckCalcFunctions.MC_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            rnd_Accuracy_Test_Results = tuple_CalcResults.Item1
            Scaled_WattHourMeter = tuple_CalcResults.Item2
            Scaled_WattHourConsole = tuple_CalcResults.Item3
            Error_Calc_Meter = tuple_CalcResults.Item4
            'Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            'Accuracy_Test_Results = tuple_CalcResults.Item6


            Call WriteMC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                     , Error_Calc_Meter, voltage, "VAH @ 25% @ PF", oSheet, 4)


            'tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)


            '' **** Update Accuracy Check Window *************
            'Call MC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
            '                         , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
            '                         , rnd_Accuracy_Test_Results.ToString)
            '' Write results to DB

            'Dim myconnection As New ADODB.Connection
            'Dim mycommand As New ADODB.Command
            'Dim ra As Integer
            'Dim Load As String
            'Dim powerfactor As String
            'Dim connt As ADODB.Connection
            'Dim connectionString As String
            'Dim external As String
            'Dim Recset As New ADODB.Recordset
            'Dim Recset1 As New ADODB.Recordset
            'Dim Recset2 As New ADODB.Recordset
            'Dim Mdate As String
            'Dim mdDate As DateTime
            'Dim Unit As String = ""
            'external = WHexternal.ToString
            'Load = "12.5"
            'powerfactor = "0.5"
            'Unit = "Vah"


            'If Not Button13.BackColor = Color.Green Then
            '    Vtap = CDbl(numtap) + 1
            'Else
            '    Vtap = CDbl(numtap)

            'End If

            'Vtap = Vtap.ToString
            'If Len(Vtap) < 2 Then
            '    Vtap = "M0" & Vtap
            'Else
            '    Vtap = "M" & Vtap
            'End If

            'myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            'myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
            '           "'" & Unit & "', " & _
            '           "'" & outp & "', " & _
            '           "'" & Load & "', " & _
            '           "'" & powerfactor & "', " & _
            '           "'" & Vtap & "', " & _
            '           "'" & _Error & "'," & _
            '           "'" & WHexternal & "'," & _
            '           "'" & WHinternal & "'," & _
            '           "'" & opertr & "'," & _
            '           "'" & tdate & "'" & _
            '           ")")

            'Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            'If Not Uniq_ID_Flag > 0 Then
            '    If Not Recset.EOF Then
            '        Mdate = Recset.GetString
            '        mdDate = DateTime.Parse(Mdate)
            '        myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
            '                                            "'" & mdDate & "'" & _
            '                  "))")

            '        Uniq_ID_Flag = 1
            '    End If
            '    Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

            '    If Not Recset1.EOF Then
            '        UniqID_text = Recset1.GetString
            '        UniqID_text = CInt(UniqID_text)
            '        myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


            '    End If

            'Else

            '    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            'End If

            'myconnection.Close()

        End If
        If ACCTESTFLAG = 1 And RadioButton40.Checked = True And TextBox2.Text = "600" Then
            Call Button2_Click(0, System.EventArgs.Empty)
            'Exit Sub

        End If


    End Sub

    Private Sub DAC_UpdateTextbox54(ByVal Testing_DateTime As DateTime, ByVal str_Technician As String, ByVal outp As String, ByVal tap As String _
                                , ByVal Meter_Type As String, ByVal SerialNumber As String, ByVal WHinternal As String, ByVal WHexternal As String _
                                , ByVal multiplier As String, ByVal Scaled_WHExternal As Double, ByVal Error_WHInternal As String _
                                , ByVal Error_WHExternal As String, ByVal Scaled_WHInternal As Double, ByVal Error_Checked_Device As Double _
                                , ByVal Corrected_Error_External As Double, ByVal DAC_results As String, ByVal test_number As Integer)
        Dim str_Units As String = "wh"
        Dim str_Current As String = "25.0 mA"
        Dim str_degree As String = "0 Deg ," & "15 Sec"
        If (test_number = 1) Then
            str_Current = "2.50 mA"
        ElseIf (test_number = 3) Then
            str_degree = "60 Deg ," & "15 Sec"
        ElseIf (test_number = 4) Then
            str_Units = "Vah"
            str_degree = "60 Deg ," & "15 Sec"
        End If


        TextBox54.Text = TextBox54.Text & Environment.NewLine & Environment.NewLine & Testing_DateTime & ":  Accuracy Check  Bench 2 " & Environment.NewLine & Environment.NewLine & "Operator:    " & str_Technician & Environment.NewLine & "Output:      " & outp & "   " & str_Current & "   " & tap & str_degree _
        & Environment.NewLine & Environment.NewLine & "Meter Type :  " & Meter_Type & "         S/N:  " & SerialNumber _
        & Environment.NewLine & Environment.NewLine & "Console Reading = " & CDbl(WHinternal).ToString & " " & str_Units & "   Voltage Multiplier " & multiplier _
        & Environment.NewLine & Environment.NewLine & "Checked Device :  " & CDbl(WHexternal) & " X 1000  = " & Scaled_WHExternal.ToString & " " & str_Units _
        & Environment.NewLine & Environment.NewLine & "Known errors (%)      Console Error: " & Error_WHInternal & "     Checked Device Error: " & Error_WHExternal _
        & Environment.NewLine & Environment.NewLine & "Checked Device Error (from Console): ((" & Math.Round(Scaled_WHExternal, 4) & " - " & Math.Round(Scaled_WHInternal, 4) & ") / " & Math.Round(Scaled_WHInternal, 4) & ")  X   100 = " & Math.Round(Error_Checked_Device, 3) _
        & Environment.NewLine & "Corrected Error of Checked Device: " & Math.Round(Error_Checked_Device, 3) & " - " & Error_WHInternal & " = " & Math.Round(Corrected_Error_External, 3).ToString _
        & Environment.NewLine & Environment.NewLine & "Accuracy Test Result: " & Math.Round(Corrected_Error_External, 3) & " - " & Error_WHExternal & " = " & DAC_results



        If CDbl(DAC_results) > 0.1 Then

            TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Failed ***********************************"
            Dim Failed_Result As DialogResult = MessageBox.Show("The Meter failed the Accuracy Check. Do you wish to continue testing?", _
                                                                    "FAILED METER - Daily Accuracy Check", _
                                                                    MessageBoxButtons.YesNo)
            If (Failed_Result = Windows.Forms.DialogResult.No) Then
                TextBox54.Text = TextBox54.Text & Environment.NewLine & "******************************* User Stopped Test ******************************"
                Call StopTest()
            End If
        Else
            TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Passed ***********************************"
        End If

        TextBox54.Refresh()
        'Set the cursor to the end of the textbox.
        TextBox54.SelectionStart = TextBox54.TextLength
        'Scroll down to the cursor position.
        TextBox54.ScrollToCaret()
    End Sub

    Private Sub MC_UpdateTextbox54(ByVal Testing_DateTime As DateTime, ByVal str_Technician As String, ByVal outp As String, ByVal tap As String _
                               , ByVal Meter_Type As String, ByVal SerialNumber As String, ByVal WHinternal As String, ByVal WHexternal As String _
                               , ByVal multiplier As String, ByVal Scaled_WHExternal As Double, ByVal Error_WHInternal As String _
                               , ByVal Error_WHExternal As String, ByVal Scaled_WHInternal As Double, ByVal Error_Checked_Device As Double _
                               , ByVal Corrected_Error_External As Double, ByVal DAC_results As String)
        TextBox54.Text = TextBox54.Text & Environment.NewLine & Environment.NewLine & Testing_DateTime & ":  MC Accuracy Check  Bench 2 " & Environment.NewLine & Environment.NewLine & "Operator:    " & str_Technician & Environment.NewLine & "Output:      " & outp & "   2.50  mA   " & tap & "0 Deg ," & "15 Sec" _
        & Environment.NewLine & Environment.NewLine & "Meter Type :  " & Meter_Type & "         S/N:  " & SerialNumber _
        & Environment.NewLine & Environment.NewLine & "Console Reading = " & CDbl(WHinternal).ToString & " wh" & "   Voltage Multiplier " & multiplier _
        & Environment.NewLine & Environment.NewLine & "MC Radian Device :  " & CDbl(WHexternal) & " X 1000  = " & Scaled_WHExternal.ToString & " wh" _
        & Environment.NewLine & Environment.NewLine & "Known errors (%)      Console Error: " & Error_WHInternal _
        & Environment.NewLine & Environment.NewLine & "Checked Device Error (from Console): ((" & Math.Round(Scaled_WHInternal, 4) & " - " & Math.Round(Scaled_WHExternal, 4) & ") / " & Math.Round(Scaled_WHExternal, 4) & ")  X   100 = " & Math.Round(Error_Checked_Device, 3) _
        & Environment.NewLine & "Corrected Error of Checked Device: " & Math.Round(Error_Checked_Device, 3) & " - " & Error_WHInternal & " = " & Math.Round(Corrected_Error_External, 3).ToString _
        & Environment.NewLine & Environment.NewLine & "Accuracy Test Result: " & Math.Round(Corrected_Error_External, 3) & " - " & Error_WHExternal & " = " & DAC_results



        If CDbl(DAC_results) > 0.1 Then

            TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Failed ***********************************"

        Else
            TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Passed ***********************************"
        End If

        TextBox54.Refresh()
        'Set the cursor to the end of the textbox.
        TextBox54.SelectionStart = TextBox54.TextLength
        'Scroll down to the cursor position.
        TextBox54.ScrollToCaret()
    End Sub


    Private Sub SetTxtBoxtoEmpty()


        TextBox49.Text = ""
        TextBox49.BackColor = Color.White
        TextBox50.Text = ""
        TextBox50.BackColor = Color.White
        TextBox51.Text = ""
        TextBox51.BackColor = Color.White
        TextBox10.BackColor = Color.White
        TextBox10.Text = ""
        TextBox11.BackColor = Color.White
        TextBox11.Text = ""
        TextBox12.BackColor = Color.White
        TextBox12.Text = ""
        TextBox13.BackColor = Color.White
        TextBox13.Text = ""
        TextBox14.BackColor = Color.White
        TextBox14.Text = ""
        TextBox15.BackColor = Color.White
        TextBox15.Text = ""
        TextBox16.BackColor = Color.White
        TextBox16.Text = ""
        TextBox17.BackColor = Color.White
        TextBox17.Text = ""
        TextBox18.BackColor = Color.White
        TextBox18.Text = ""
        TextBox19.BackColor = Color.White
        TextBox19.Text = ""
        TextBox20.BackColor = Color.White
        TextBox20.Text = ""
        TextBox21.BackColor = Color.White
        TextBox21.Text = ""
        TextBox22.BackColor = Color.White
        TextBox22.Text = ""
        TextBox23.BackColor = Color.White
        TextBox23.Text = ""
        TextBox24.BackColor = Color.White
        TextBox24.Text = ""
        TextBox25.BackColor = Color.White
        TextBox25.Text = ""
        TextBox26.BackColor = Color.White
        TextBox26.Text = ""
        TextBox27.BackColor = Color.White
        TextBox27.Text = ""
        TextBox28.BackColor = Color.White
        TextBox28.Text = ""
        TextBox29.BackColor = Color.White
        TextBox29.Text = ""
        TextBox30.BackColor = Color.White
        TextBox30.Text = ""
        TextBox31.BackColor = Color.White
        TextBox31.Text = ""
        TextBox32.BackColor = Color.White
        TextBox32.Text = ""
        TextBox33.BackColor = Color.White
        TextBox33.Text = ""
        TextBox34.BackColor = Color.White
        TextBox34.Text = ""
        TextBox35.BackColor = Color.White
        TextBox35.Text = ""
        TextBox36.BackColor = Color.White
        TextBox36.Text = ""
        TextBox37.BackColor = Color.White
        TextBox37.Text = ""
        TextBox38.BackColor = Color.White
        TextBox38.Text = ""
        TextBox39.BackColor = Color.White
        TextBox39.Text = ""
        TextBox40.BackColor = Color.White
        TextBox40.Text = ""
        TextBox41.BackColor = Color.White
        TextBox41.Text = ""
        TextBox42.BackColor = Color.White
        TextBox42.Text = ""
        TextBox43.BackColor = Color.White
        TextBox43.Text = ""
        TextBox44.BackColor = Color.White
        TextBox44.Text = ""
        TextBox45.BackColor = Color.White
        TextBox45.Text = ""
        TextBox46.BackColor = Color.White
        TextBox46.Text = ""
        TextBox47.BackColor = Color.White
        TextBox47.Text = ""
        TextBox48.BackColor = Color.White
        TextBox48.Text = ""
    End Sub

    Private Function getTapNumber(ByVal numTap As String, ByVal txtbxY_BackgroundColour As Color, ByRef rButtonName As String) As String
        Dim tap As String = ""

        If Not Button13.BackColor = Color.Green Then
            Select Case rButtonName

                Case "RadioButton40"
                    tap = "( mA Tap 40 ) , "
                Case "RadioButton39"
                    tap = "( mA Tap 39 ) , "
                Case "RadioButton38"
                    tap = "( mA Tap 38 ) , "
                Case "RadioButton37"
                    tap = "( mA Tap 37 ) , "
                Case "RadioButton36"
                    tap = "( mA Tap 36 ) , "
                Case "RadioButton35"
                    tap = "( mA Tap 35 ) , "
                Case "RadioButton34"
                    tap = "( mA Tap 34 ) , "
                Case "RadioButton33"
                    tap = "( mA Tap 33 ) , "
                Case "RadioButton32"
                    tap = "( mA Tap 32 ) , "
                Case "RadioButton31"
                    tap = "( mA Tap 31 ) , "
                Case "RadioButton30"
                    tap = "( mA Tap 30 ) , "
                Case "RadioButton29"
                    tap = "( mA Tap 29 ) , "
                Case "RadioButton28"
                    tap = "( mA Tap 28 ) , "
                Case "RadioButton27"
                    tap = "( mA Tap 27 ) , "
                Case "RadioButton26"
                    tap = "( mA Tap 26 ) , "
                Case "RadioButton25"
                    tap = "( mA Tap 25 ) , "
                Case "RadioButton24"
                    tap = "( mA Tap 24 ) , "
                Case "RadioButton23"
                    tap = "( mA Tap 23 ) , "
                Case "RadioButton22"
                    tap = "( mA Tap 22 ) , "
                Case "RadioButton21"
                    tap = "( mA Tap 21 ) , "
                Case "RadioButton20"
                    tap = "( mA Tap 20 ) , "
                Case "RadioButton19"
                    tap = "( mA Tap 19 ) , "
                Case "RadioButton18"
                    tap = "( mA Tap 18 ) , "
                Case "RadioButton17"
                    tap = "( mA Tap 17 ) , "
                Case "RadioButton16"
                    tap = "( mA Tap 16 ) , "
                Case "RadioButton15"
                    tap = "( mA Tap 15 ) , "
                Case "RadioButton14"
                    tap = "( mA Tap 14 ) , "
                Case "RadioButton13"
                    tap = "( mA Tap 13 ) , "
                Case "RadioButton12"
                    tap = "( mA Tap 12 ) , "
                Case "RadioButton11"
                    tap = "( mA Tap 11 ) , "
                Case "RadioButton10"
                    tap = "( mA Tap 30 ) , "
                Case "RadioButton9"
                    tap = "( mA Tap 9 ) , "
                Case "RadioButton8"
                    tap = "( mA Tap 8 ) , "
                Case "RadioButton7"
                    tap = "( mA Tap 7 ) , "
                Case "RadioButton6"
                    tap = "( mA Tap 6 ) , "
                Case "RadioButton5"
                    tap = "( mA Tap 5 ) , "
                Case "RadioButton4"
                    tap = "( mA Tap 4 ) , "
                Case "RadioButton3"
                    tap = "( mA Tap 3 ) , "
                Case "RadioButton2"
                    tap = "( mA Tap 2 ) , "
            End Select

        Else

            numTap = CInt(numTap) + 1
            numTap = numTap.ToString
            If Len(numTap) = 1 Then
                numTap = "0" & numTap
            End If





            Select Case numTap

                Case "40"
                    tap = "( mA Tap 40 ) , "
                Case "39"
                    tap = "( mA Tap 39 ) , "
                Case "38"
                    tap = "( mA Tap 38 ) , "
                Case "37"
                    tap = "( mA Tap 37 ) , "
                Case "36"
                    tap = "( mA Tap 36 ) , "
                Case "35"
                    tap = "( mA Tap 35 ) , "
                Case "34"
                    tap = "( mA Tap 34 ) , "
                Case "33"
                    tap = "( mA Tap 33 ) , "
                Case "32"
                    tap = "( mA Tap 32 ) , "
                Case "31"
                    tap = "( mA Tap 31 ) , "
                Case "30"
                    tap = "( mA Tap 30 ) , "
                Case "29"
                    tap = "( mA Tap 29 ) , "
                Case "28"
                    tap = "( mA Tap 28 ) , "
                Case "27"
                    tap = "( mA Tap 27 ) , "
                Case "26"
                    tap = "( mA Tap 26 ) , "
                Case "25"
                    tap = "( mA Tap 25 ) , "
                Case "24"
                    tap = "( mA Tap 24 ) , "
                Case "23"
                    tap = "( mA Tap 23 ) , "
                Case "22"
                    tap = "( mA Tap 22 ) , "
                Case "21"
                    tap = "( mA Tap 21 ) , "
                Case "20"
                    tap = "( mA Tap 20 ) , "
                Case "19"
                    tap = "( mA Tap 19 ) , "
                Case "18"
                    tap = "( mA Tap 18 ) , "
                Case "17"
                    tap = "( mA Tap 17 ) , "
                Case "16"
                    tap = "( mA Tap 16 ) , "
                Case "15"
                    tap = "( mA Tap 15 ) , "
                Case "14"
                    tap = "( mA Tap 14 ) , "
                Case "13"
                    tap = "( mA Tap 13 ) , "
                Case "12"
                    tap = "( mA Tap 12 ) , "
                Case "11"
                    tap = "( mA Tap 11 ) , "
                Case "10"
                    tap = "( mA Tap 30 ) , "
                Case "09"
                    tap = "( mA Tap 9 ) , "
                Case "08"
                    tap = "( mA Tap 8 ) , "
                Case "07"
                    tap = "( mA Tap 7 ) , "
                Case "06"
                    tap = "( mA Tap 6 ) , "
                Case "05"
                    tap = "( mA Tap 5 ) , "
                Case "04"
                    tap = "( mA Tap 4 ) , "
                Case "03"
                    tap = "( mA Tap 3 ) , "
                Case "02"
                    tap = "( mA Tap 2 ) , "
            End Select

        End If
        Return tap
    End Function

    Public Sub SetAfterPrn()
        Dim WHinternal As Single
        Dim WHexternal As Single
        Dim IntCount(256) As Long
        Dim cbox1 As String
        Dim voltage As String
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)
        On Error Resume Next

        Call SetTxtBoxtoEmpty()
        'TextBox49.Text = ""
        'TextBox49.BackColor = Color.White
        'TextBox50.Text = ""
        'TextBox50.BackColor = Color.White
        'TextBox51.Text = ""
        'TextBox51.BackColor = Color.White
        'TextBox10.BackColor = Color.White
        'TextBox10.Text = ""
        'TextBox11.BackColor = Color.White
        'TextBox11.Text = ""
        'TextBox12.BackColor = Color.White
        'TextBox12.Text = ""
        'TextBox13.BackColor = Color.White
        'TextBox13.Text = ""
        'TextBox14.BackColor = Color.White
        'TextBox14.Text = ""
        'TextBox15.BackColor = Color.White
        'TextBox15.Text = ""
        'TextBox16.BackColor = Color.White
        'TextBox16.Text = ""
        'TextBox17.BackColor = Color.White
        'TextBox17.Text = ""
        'TextBox18.BackColor = Color.White
        'TextBox18.Text = ""
        'TextBox19.BackColor = Color.White
        'TextBox19.Text = ""
        'TextBox20.BackColor = Color.White
        'TextBox20.Text = ""
        'TextBox21.BackColor = Color.White
        'TextBox21.Text = ""
        'TextBox22.BackColor = Color.White
        'TextBox22.Text = ""
        'TextBox23.BackColor = Color.White
        'TextBox23.Text = ""
        'TextBox24.BackColor = Color.White
        'TextBox24.Text = ""
        'TextBox25.BackColor = Color.White
        'TextBox25.Text = ""
        'TextBox26.BackColor = Color.White
        'TextBox26.Text = ""
        'TextBox27.BackColor = Color.White
        'TextBox27.Text = ""
        'TextBox28.BackColor = Color.White
        'TextBox28.Text = ""
        'TextBox29.BackColor = Color.White
        'TextBox29.Text = ""
        'TextBox30.BackColor = Color.White
        'TextBox30.Text = ""
        'TextBox31.BackColor = Color.White
        'TextBox31.Text = ""
        'TextBox32.BackColor = Color.White
        'TextBox32.Text = ""
        'TextBox33.BackColor = Color.White
        'TextBox33.Text = ""
        'TextBox34.BackColor = Color.White
        'TextBox34.Text = ""
        'TextBox35.BackColor = Color.White
        'TextBox35.Text = ""
        'TextBox36.BackColor = Color.White
        'TextBox36.Text = ""
        'TextBox37.BackColor = Color.White
        'TextBox37.Text = ""
        'TextBox38.BackColor = Color.White
        'TextBox38.Text = ""
        'TextBox39.BackColor = Color.White
        'TextBox39.Text = ""
        'TextBox40.BackColor = Color.White
        'TextBox40.Text = ""
        'TextBox41.BackColor = Color.White
        'TextBox41.Text = ""
        'TextBox42.BackColor = Color.White
        'TextBox42.Text = ""
        'TextBox43.BackColor = Color.White
        'TextBox43.Text = ""
        'TextBox44.BackColor = Color.White
        'TextBox44.Text = ""
        'TextBox45.BackColor = Color.White
        'TextBox45.Text = ""
        'TextBox46.BackColor = Color.White
        'TextBox46.Text = ""
        'TextBox47.BackColor = Color.White
        'TextBox47.Text = ""
        'TextBox48.BackColor = Color.White
        'TextBox48.Text = ""

        mbSession = CType(ResourceManager.GetLocalManager().Open("GPIB0::0::INSTR"), MessageBasedSession)
        mbSession.Write("CHAN 1 MODE GATE")
        mbSession.Write("CHAN 1 GATE OFF")
        voltage = TextBox2.Text
        Call SetChan2(voltage)

        'If voltage = "120" Then
        '    '*120 to PT1
        '    'AxAdvDIO1.DeviceNumber = 0
        '    'AxAdvDIO1.WriteDoChannel(1, 22)
        '    'AxAdvDIO1.WriteDoChannel(0, 23)
        '    ' mbSession.Write("CHAN 2 Rang 10V AMR 3.4944") 'bench3
        '    WriteDoChannel0(1, 22)
        '    WriteDoChannel0(0, 23)
        '    mbSession.Write("CHAN 2 Rang 10V AMR 5.1952") ''bench2
        'End If

        'If voltage = "208" Then
        '    '*208 to PT1
        '    'AxAdvDIO1.DeviceNumber = 0
        '    'AxAdvDIO1.WriteDoChannel(0, 22)
        '    'AxAdvDIO1.WriteDoChannel(1, 23)
        '    'mbSession.Write("CHAN 2 Rang 10V AMR 2.0189") ''Bench3
        '    WriteDoChannel0(0, 22)
        '    WriteDoChannel0(1, 23)
        '    mbSession.Write("CHAN 2 Rang 10V AMR 3.0015") ''Bench2
        'End If


        'If voltage = "240" Then
        '    '*240 to PT1
        '    'AxAdvDIO1.DeviceNumber = 0
        '    'AxAdvDIO1.WriteDoChannel(0, 22)
        '    'AxAdvDIO1.WriteDoChannel(1, 23)
        '    ' mbSession.Write("CHAN 2 Rang 10V AMR 2.3296") 'bench3
        '    WriteDoChannel0(0, 22)
        '    WriteDoChannel0(1, 23)
        '    mbSession.Write("CHAN 2 Rang 10V AMR 3.4635") 'bench2
        'End If

        'If voltage = "277" Then
        '    '*277 to PT1
        '    'AxAdvDIO1.DeviceNumber = 0
        '    'AxAdvDIO1.WriteDoChannel(0, 22)
        '    'AxAdvDIO1.WriteDoChannel(1, 23)
        '    'mbSession.Write("CHAN 2 Rang 10V AMR 2.6887") ''bench3
        '    WriteDoChannel0(0, 22)
        '    WriteDoChannel0(1, 23)
        '    mbSession.Write("CHAN 2 Rang 10V AMR 3.9974") ''bench2
        'End If


        'If voltage = "347" Then
        '    '*347 to PT1
        '    'AxAdvDIO1.DeviceNumber = 0
        '    'AxAdvDIO1.WriteDoChannel(0, 22)
        '    'AxAdvDIO1.WriteDoChannel(1, 23)
        '    'mbSession.Write("CHAN 2 Rang 10V AMR 3.3682") ''bench3
        '    WriteDoChannel0(0, 22)
        '    WriteDoChannel0(1, 23)
        '    mbSession.Write("CHAN 2 Rang 10V AMR 5.0076") ''bench2
        'End If
        'If voltage = "416" Then
        '    '*416 to PT1
        '    ' AxAdvDIO1.DeviceNumber = 0
        '    ' AxAdvDIO1.WriteDoChannel(1, 22)
        '    ' AxAdvDIO1.WriteDoChannel(1, 23)
        '    ' mbSession.Write("CHAN 2 Rang 10V AMR 2.4228") ''bench3
        '    WriteDoChannel0(1, 22)
        '    WriteDoChannel0(1, 23)
        '    mbSession.Write("CHAN 2 Rang 10V AMR 3.6020") ''Bench2
        'End If
        'If voltage = "480" Then
        '    '*480 to PT1
        '    ' AxAdvDIO1.DeviceNumber = 0
        '    'AxAdvDIO1.WriteDoChannel(1, 22)
        '    ' AxAdvDIO1.WriteDoChannel(1, 23)
        '    ' mbSession.Write("CHAN 2 Rang 10V AMR 2.7955") ''Bench3
        '    WriteDoChannel0(1, 22)
        '    WriteDoChannel0(1, 23)
        '    mbSession.Write("CHAN 2 Rang 10V AMR 4.1562 ") ''Bench2

        'End If
        'If voltage = "600" Then
        '    '*600 to PT1
        '    'AxAdvDIO1.DeviceNumber = 0
        '    ' AxAdvDIO1.WriteDoChannel(1, 22)
        '    ' AxAdvDIO1.WriteDoChannel(1, 23)
        '    'mbSession.Write("CHAN 2 Rang 10V AMR 3.4944") ''bench3
        '    WriteDoChannel0(1, 22)
        '    WriteDoChannel0(1, 23)
        '    mbSession.Write("CHAN 2 Rang 10V AMR 5.1952 ") ''Bench2
        'End If
        Call teststep1()

        '88888888888888888888888888888888888888888888888888888888888888888888 After Step 1 8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            ' RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If


        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================
        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(150)
        End While

        Application.DoEvents()

        Me.BackgroundWorker2.RunWorkerAsync()
        Threading.Thread.Sleep(150)

        Form1.Form1_CallScans(5)
        Thread.Sleep(500)


        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If

        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If


        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999

        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If

        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        'Threading.Thread.Sleep(250)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(60)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        Thread.Sleep(100)
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")
        TextBox52.Text = ""
        TextBox52.BackColor = Color.White
        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If



        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.

        Dim LastRow As Long
        With oSheet
            LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row


            '''''''''''''''''''' Top Row Only'''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .cells(1, 1).Value = "" Then


                .cells(1, 1).Value = "TAP"
                .cells(1, 2).Value = "Console Radian"
                .cells(1, 3).Value = "MC Radian"
                .cells(1, 4).Value = "Current Step"
                .cells(1, 5).Value = "V Mult"
                .cells(1, 6).Value = "Rcons"
                .cells(1, 7).Value = "C Mult"
                .cells(1, 8).Value = "Rmc"
                .cells(1, 9).Value = "% Err"

            End If

            '''''''''''''''''''''''''''''''''CT and Votage '''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''add CT here
            LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
            .cells(LastRow + 1, 1).value = "CT"
            .cells(LastRow + 1, 2).Value = voltage



            '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
            LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row

            If Len(CBNUM) = 13 Then
                CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
            Else
                CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
            End If

            .cells(LastRow + 1, 1).Value = CTNUM
            .cells(LastRow + 1, 2).Value = WHinternal
            .cells(LastRow + 1, 3).Value = WHexternal
            .cells(LastRow + 1, 4).Value = "WH @ 2.5% Unity"

            voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
            .cells(LastRow + 1, 5).Value = voltage

            .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
            .cells(LastRow + 1, 7).Value = 1000
            .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
            .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
            TextBox49.Text = .cells(LastRow + 1, 9).Text
            If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                TextBox49.BackColor = Color.Red
            Else
                TextBox49.BackColor = Color.Lime
            End If

            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

            Application.DoEvents()
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End With
        '88888888888888888888888888888888888888888888888888888888888888 step 1 error calc 88888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
        If ACCTESTFLAG = 1 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""
            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double


            Dim Error_checked_Device As Double = 0
            Dim Corrected_error_checked_device As Double = 0
            Dim Accuracy_test_result As Double = 0
            Dim XX As Double
            Dim Y As Double

            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[WH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(117) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 4)
                        WH100array(ii) = strFileName(iii)
                        Dim fstring As String = "1.25_1.0_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()

            'Select Case Vtap
            '    Case "1"
            '        mulitplier = "1.000"
            '    Case "3"
            '        mulitplier = "2.999"
            '    Case "5"
            '        mulitplier = "5.002"
            'End Select

            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Radian = Math.Round(CDbl(WHinternal), 3)
            Dim Percent2 As String
            Dim Percent3 As String




            '*** Checked Device Error determined by Console  = (Check Device Reading X 1000 (convert from milliamp to amps) – Console Reading X the Voltage Multiplier(based on the voltage from the PRN) / Console Reading X the Voltage Multiplier(based on the voltage from the PRN)) X 100 
            XX = CDbl(WHexternal) * 1000
            Y = CDbl(WHinternal) * CDbl(mulitplier)

            Error_checked_Device = ((Y - XX) / XX) * 100

            '*** Corrected Error of the Checked device =  Checked Device Error determined by Console - %Econ(from MC Spreadsheet)
            Corrected_error_checked_device = Error_checked_Device - _Error

            '*** Accuracy Test Result = Corrected Error of the Checked device - P-E-01 6.1.1.2(1C) reference meter option results (to be determined by Steve)

            Accuracy_test_result = Corrected_error_checked_device - Estd

            percent = Math.Abs(Math.Round(Accuracy_test_result, 3)).ToString
            'Percent2 = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            'Percent3 = AccuracyCheckCalcFunctions.MC_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))


            tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)
            'If Not Button13.BackColor = Color.Green Then
            '    Select Case rButton.Name

            '        Case "RadioButton40"
            '            tap = "( mA Tap 40 ) , "
            '        Case "RadioButton39"
            '            tap = "( mA Tap 39 ) , "
            '        Case "RadioButton38"
            '            tap = "( mA Tap 38 ) , "
            '        Case "RadioButton37"
            '            tap = "( mA Tap 37 ) , "
            '        Case "RadioButton36"
            '            tap = "( mA Tap 36 ) , "
            '        Case "RadioButton35"
            '            tap = "( mA Tap 35 ) , "
            '        Case "RadioButton34"
            '            tap = "( mA Tap 34 ) , "
            '        Case "RadioButton33"
            '            tap = "( mA Tap 33 ) , "
            '        Case "RadioButton32"
            '            tap = "( mA Tap 32 ) , "
            '        Case "RadioButton31"
            '            tap = "( mA Tap 31 ) , "
            '        Case "RadioButton30"
            '            tap = "( mA Tap 30 ) , "
            '        Case "RadioButton29"
            '            tap = "( mA Tap 29 ) , "
            '        Case "RadioButton28"
            '            tap = "( mA Tap 28 ) , "
            '        Case "RadioButton27"
            '            tap = "( mA Tap 27 ) , "
            '        Case "RadioButton26"
            '            tap = "( mA Tap 26 ) , "
            '        Case "RadioButton25"
            '            tap = "( mA Tap 25 ) , "
            '        Case "RadioButton24"
            '            tap = "( mA Tap 24 ) , "
            '        Case "RadioButton23"
            '            tap = "( mA Tap 23 ) , "
            '        Case "RadioButton22"
            '            tap = "( mA Tap 22 ) , "
            '        Case "RadioButton21"
            '            tap = "( mA Tap 21 ) , "
            '        Case "RadioButton20"
            '            tap = "( mA Tap 20 ) , "
            '        Case "RadioButton19"
            '            tap = "( mA Tap 19 ) , "
            '        Case "RadioButton18"
            '            tap = "( mA Tap 18 ) , "
            '        Case "RadioButton17"
            '            tap = "( mA Tap 17 ) , "
            '        Case "RadioButton16"
            '            tap = "( mA Tap 16 ) , "
            '        Case "RadioButton15"
            '            tap = "( mA Tap 15 ) , "
            '        Case "RadioButton14"
            '            tap = "( mA Tap 14 ) , "
            '        Case "RadioButton13"
            '            tap = "( mA Tap 13 ) , "
            '        Case "RadioButton12"
            '            tap = "( mA Tap 12 ) , "
            '        Case "RadioButton11"
            '            tap = "( mA Tap 11 ) , "
            '        Case "RadioButton10"
            '            tap = "( mA Tap 30 ) , "
            '        Case "RadioButton9"
            '            tap = "( mA Tap 9 ) , "
            '        Case "RadioButton8"
            '            tap = "( mA Tap 8 ) , "
            '        Case "RadioButton7"
            '            tap = "( mA Tap 7 ) , "
            '        Case "RadioButton6"
            '            tap = "( mA Tap 6 ) , "
            '        Case "RadioButton5"
            '            tap = "( mA Tap 5 ) , "
            '        Case "RadioButton4"
            '            tap = "( mA Tap 4 ) , "
            '        Case "RadioButton3"
            '            tap = "( mA Tap 3 ) , "
            '        Case "RadioButton2"
            '            tap = "( mA Tap 2 ) , "
            '    End Select

            'Else

            '    numtap = CInt(numtap) + 1
            '    numtap = numtap.ToString
            '    If Len(numtap) = 1 Then
            '        numtap = "0" & numtap
            '    End If





            '    Select Case numtap

            '        Case "40"
            '            tap = "( mA Tap 40 ) , "
            '        Case "39"
            '            tap = "( mA Tap 39 ) , "
            '        Case "38"
            '            tap = "( mA Tap 38 ) , "
            '        Case "37"
            '            tap = "( mA Tap 37 ) , "
            '        Case "36"
            '            tap = "( mA Tap 36 ) , "
            '        Case "35"
            '            tap = "( mA Tap 35 ) , "
            '        Case "34"
            '            tap = "( mA Tap 34 ) , "
            '        Case "33"
            '            tap = "( mA Tap 33 ) , "
            '        Case "32"
            '            tap = "( mA Tap 32 ) , "
            '        Case "31"
            '            tap = "( mA Tap 31 ) , "
            '        Case "30"
            '            tap = "( mA Tap 30 ) , "
            '        Case "29"
            '            tap = "( mA Tap 29 ) , "
            '        Case "28"
            '            tap = "( mA Tap 28 ) , "
            '        Case "27"
            '            tap = "( mA Tap 27 ) , "
            '        Case "26"
            '            tap = "( mA Tap 26 ) , "
            '        Case "25"
            '            tap = "( mA Tap 25 ) , "
            '        Case "24"
            '            tap = "( mA Tap 24 ) , "
            '        Case "23"
            '            tap = "( mA Tap 23 ) , "
            '        Case "22"
            '            tap = "( mA Tap 22 ) , "
            '        Case "21"
            '            tap = "( mA Tap 21 ) , "
            '        Case "20"
            '            tap = "( mA Tap 20 ) , "
            '        Case "19"
            '            tap = "( mA Tap 19 ) , "
            '        Case "18"
            '            tap = "( mA Tap 18 ) , "
            '        Case "17"
            '            tap = "( mA Tap 17 ) , "
            '        Case "16"
            '            tap = "( mA Tap 16 ) , "
            '        Case "15"
            '            tap = "( mA Tap 15 ) , "
            '        Case "14"
            '            tap = "( mA Tap 14 ) , "
            '        Case "13"
            '            tap = "( mA Tap 13 ) , "
            '        Case "12"
            '            tap = "( mA Tap 12 ) , "
            '        Case "11"
            '            tap = "( mA Tap 11 ) , "
            '        Case "10"
            '            tap = "( mA Tap 30 ) , "
            '        Case "09"
            '            tap = "( mA Tap 9 ) , "
            '        Case "08"
            '            tap = "( mA Tap 8 ) , "
            '        Case "07"
            '            tap = "( mA Tap 7 ) , "
            '        Case "06"
            '            tap = "( mA Tap 6 ) , "
            '        Case "05"
            '            tap = "( mA Tap 5 ) , "
            '        Case "04"
            '            tap = "( mA Tap 4 ) , "
            '        Case "03"
            '            tap = "( mA Tap 3 ) , "
            '        Case "02"
            '            tap = "( mA Tap 2 ) , "
            '    End Select

            'End If

            TextBox54.Text = TextBox54.Text & Environment.NewLine & Environment.NewLine & tdate & ":  Accuracy Check  Bench 2 " & Environment.NewLine & Environment.NewLine & "Operator:    " & opertr & Environment.NewLine & "Output:      " & outp & "   2.50  mA   " & tap & "0 Deg ," & "15 Sec" _
                & Environment.NewLine & Environment.NewLine & "Meter Type :  " & modl & "         S/N:  " & seral _
                & Environment.NewLine & Environment.NewLine & "Console Reading = " & CDbl(WHinternal).ToString & " wh" & "   Voltage Multiplier " & mulitplier _
                & Environment.NewLine & Environment.NewLine & "Checked Device :  " & CDbl(WHexternal) & " X 1000  = " & XX & " wh" _
                & Environment.NewLine & Environment.NewLine & "Known errors (%)      Console Error: " & _Error & "     Checked Device Error: " & Estd _
                & Environment.NewLine & Environment.NewLine & "Checked Device Error (from Console): ((" & Math.Round(Y, 4) & " - " & Math.Round(XX, 4) & ") / " & Math.Round(XX, 4) & ")  X   100 = " & Math.Round(Error_checked_Device, 3) _
                & Environment.NewLine & "Corrected Error of Checked Device: " & Math.Round(Error_checked_Device, 3) & " - " & _Error & " = " & Math.Round(Corrected_error_checked_device, 3).ToString _
                & Environment.NewLine & Environment.NewLine & "Accuracy Test Result: " & Math.Round(Corrected_error_checked_device, 3) & " - " & Estd & " = " & percent _
                & Environment.NewLine & Environment.NewLine & "Daily Accuracy Current Test Result: " & percent
            '& Environment.NewLine & Environment.NewLine & "Daily Accuracy Function Percent Test Result: " & Percent2 _
            '& Environment.NewLine & Environment.NewLine & "MC Accuracy Function Percent Test Result: " & Percent3


            If percent > 0.1 Then

                TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Failed ***********************************"
            Else
                TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Passed ***********************************"

            End If

            TextBox54.Refresh()
            'Set the cursor to the end of the textbox.
            TextBox54.SelectionStart = TextBox54.TextLength
            'Scroll down to the cursor position.
            TextBox54.ScrollToCaret()

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            external = WHexternal.ToString
            Load = "1.25"
            powerfactor = "1.0"
            Unit = "WH"


            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If

            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If

            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()






        End If



        Call teststep2()

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(200)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        '    Status = RadRDAssignDevice(CShort(commport1), comDevice)

        '    If Status = 0 Then
        '        'Successfully connected
        '        'Get unit information and populate status bar
        '        Status = RadRDModel(comDevice, Model)
        '        Status = RadRDSerial(comDevice, Serial)
        '        Status = RadRDVersion(comDevice, Version)
        '        Status = RadRDName(comDevice, DeviceName)
        '    End If
        'RadRDAccumStart(comDevice)
        '    If comDevice <> 0 Then
        '        RadRDReleaseDevice(comDevice)
        '        comDevice = 0
        '    End If




        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999







        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If
        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(100)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While

        BackgroundWorker2.RunWorkerAsync()

        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        Thread.Sleep(100)
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")
        'MsgBox("set Yokogawa Voltage , Current, Phase   OFF")
        'MsgBox("Read MC MKA to excel")
        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.
100:

        With oSheet
            '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
            LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
            If Len(CBNUM) = 13 Then
                CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
            Else
                CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
            End If

            .cells(LastRow + 1, 1).Value = CTNUM
            .cells(LastRow + 1, 2).Value = WHinternal
            .cells(LastRow + 1, 3).Value = WHexternal
            .cells(LastRow + 1, 4).Value = "WH @ 25% Unity"

            voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
            .cells(LastRow + 1, 5).Value = voltage



            .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
            .cells(LastRow + 1, 7).Value = 1000
            .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
            .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
            TextBox50.Text = .cells(LastRow + 1, 9).Text
            If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                TextBox50.BackColor = Color.Red
            Else
                TextBox50.BackColor = Color.Lime
            End If

            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

            Application.DoEvents()
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End With
        If ACCTESTFLAG = 1 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""
            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double

            Dim Error_checked_Device As Double = 0
            Dim Corrected_error_checked_device As Double = 0
            Dim Accuracy_test_result As Double = 0
            Dim XX As Double
            Dim Y As Double

            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[WH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(116) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 4)
                        WH100array(ii) = strFileName(iii)
                        Dim fstring As String = "12.5_1.0_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next

            '****************************************accuracy math *********************************************************************

            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Etrue = xpercent
            Change = Etrue - Emeter
            Total = Change + econsole
            Radian = Math.Round(CDbl(WHinternal), 3)

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()

            'Select Case Vtap
            '    Case "1"
            '        mulitplier = "1.000"
            '    Case "3"
            '        mulitplier = "2.999"
            '    Case "5"
            '        mulitplier = "5.002"
            'End Select

            '*** Checked Device Error determined by Console  = (Check Device Reading X 1000 (convert from milliamp to amps) – Console Reading X the Voltage Multiplier(based on the voltage from the PRN) / Console Reading X the Voltage Multiplier(based on the voltage from the PRN)) X 100 
            XX = CDbl(WHexternal) * 1000
            Y = CDbl(WHinternal) * CDbl(mulitplier)

            Error_checked_Device = ((Y - XX) / XX) * 100

            '*** Corrected Error of the Checked device =  Checked Device Error determined by Console - %Econ(from MC Spreadsheet)
            Corrected_error_checked_device = Error_checked_Device - _Error

            '*** Accuracy Test Result = Corrected Error of the Checked device - P-E-01 6.1.1.2(1C) reference meter option results (to be determined by Steve)

            Accuracy_test_result = Corrected_error_checked_device - Estd

            percent = Math.Abs(Math.Round(Accuracy_test_result, 3)).ToString



            If Not Button13.BackColor = Color.Green Then
                Select Case rButton.Name

                    Case "RadioButton40"
                        tap = "( mA Tap 40 ) , "
                    Case "RadioButton39"
                        tap = "( mA Tap 39 ) , "
                    Case "RadioButton38"
                        tap = "( mA Tap 38 ) , "
                    Case "RadioButton37"
                        tap = "( mA Tap 37 ) , "
                    Case "RadioButton36"
                        tap = "( mA Tap 36 ) , "
                    Case "RadioButton35"
                        tap = "( mA Tap 35 ) , "
                    Case "RadioButton34"
                        tap = "( mA Tap 34 ) , "
                    Case "RadioButton33"
                        tap = "( mA Tap 33 ) , "
                    Case "RadioButton32"
                        tap = "( mA Tap 32 ) , "
                    Case "RadioButton31"
                        tap = "( mA Tap 31 ) , "
                    Case "RadioButton30"
                        tap = "( mA Tap 30 ) , "
                    Case "RadioButton29"
                        tap = "( mA Tap 29 ) , "
                    Case "RadioButton28"
                        tap = "( mA Tap 28 ) , "
                    Case "RadioButton27"
                        tap = "( mA Tap 27 ) , "
                    Case "RadioButton26"
                        tap = "( mA Tap 26 ) , "
                    Case "RadioButton25"
                        tap = "( mA Tap 25 ) , "
                    Case "RadioButton24"
                        tap = "( mA Tap 24 ) , "
                    Case "RadioButton23"
                        tap = "( mA Tap 23 ) , "
                    Case "RadioButton22"
                        tap = "( mA Tap 22 ) , "
                    Case "RadioButton21"
                        tap = "( mA Tap 21 ) , "
                    Case "RadioButton20"
                        tap = "( mA Tap 20 ) , "
                    Case "RadioButton19"
                        tap = "( mA Tap 19 ) , "
                    Case "RadioButton18"
                        tap = "( mA Tap 18 ) , "
                    Case "RadioButton17"
                        tap = "( mA Tap 17 ) , "
                    Case "RadioButton16"
                        tap = "( mA Tap 16 ) , "
                    Case "RadioButton15"
                        tap = "( mA Tap 15 ) , "
                    Case "RadioButton14"
                        tap = "( mA Tap 14 ) , "
                    Case "RadioButton13"
                        tap = "( mA Tap 13 ) , "
                    Case "RadioButton12"
                        tap = "( mA Tap 12 ) , "
                    Case "RadioButton11"
                        tap = "( mA Tap 11 ) , "
                    Case "RadioButton10"
                        tap = "( mA Tap 30 ) , "
                    Case "RadioButton9"
                        tap = "( mA Tap 9 ) , "
                    Case "RadioButton8"
                        tap = "( mA Tap 8 ) , "
                    Case "RadioButton7"
                        tap = "( mA Tap 7 ) , "
                    Case "RadioButton6"
                        tap = "( mA Tap 6 ) , "
                    Case "RadioButton5"
                        tap = "( mA Tap 5 ) , "
                    Case "RadioButton4"
                        tap = "( mA Tap 4 ) , "
                    Case "RadioButton3"
                        tap = "( mA Tap 3 ) , "
                    Case "RadioButton2"
                        tap = "( mA Tap 2 ) , "
                End Select

            Else
                numtap = CInt(numtap) + 1
                numtap = numtap.ToString
                If Len(numtap) = 1 Then
                    numtap = "0" & numtap
                End If

                Select Case numtap

                    Case "40"
                        tap = "( mA Tap 40 ) , "
                    Case "39"
                        tap = "( mA Tap 39 ) , "
                    Case "38"
                        tap = "( mA Tap 38 ) , "
                    Case "37"
                        tap = "( mA Tap 37 ) , "
                    Case "36"
                        tap = "( mA Tap 36 ) , "
                    Case "35"
                        tap = "( mA Tap 35 ) , "
                    Case "34"
                        tap = "( mA Tap 34 ) , "
                    Case "33"
                        tap = "( mA Tap 33 ) , "
                    Case "32"
                        tap = "( mA Tap 32 ) , "
                    Case "31"
                        tap = "( mA Tap 31 ) , "
                    Case "30"
                        tap = "( mA Tap 30 ) , "
                    Case "29"
                        tap = "( mA Tap 29 ) , "
                    Case "28"
                        tap = "( mA Tap 28 ) , "
                    Case "27"
                        tap = "( mA Tap 27 ) , "
                    Case "26"
                        tap = "( mA Tap 26 ) , "
                    Case "25"
                        tap = "( mA Tap 25 ) , "
                    Case "24"
                        tap = "( mA Tap 24 ) , "
                    Case "23"
                        tap = "( mA Tap 23 ) , "
                    Case "22"
                        tap = "( mA Tap 22 ) , "
                    Case "21"
                        tap = "( mA Tap 21 ) , "
                    Case "20"
                        tap = "( mA Tap 20 ) , "
                    Case "19"
                        tap = "( mA Tap 19 ) , "
                    Case "18"
                        tap = "( mA Tap 18 ) , "
                    Case "17"
                        tap = "( mA Tap 17 ) , "
                    Case "16"
                        tap = "( mA Tap 16 ) , "
                    Case "15"
                        tap = "( mA Tap 15 ) , "
                    Case "14"
                        tap = "( mA Tap 14 ) , "
                    Case "13"
                        tap = "( mA Tap 13 ) , "
                    Case "12"
                        tap = "( mA Tap 12 ) , "
                    Case "11"
                        tap = "( mA Tap 11 ) , "
                    Case "10"
                        tap = "( mA Tap 30 ) , "
                    Case "09"
                        tap = "( mA Tap 9 ) , "
                    Case "08"
                        tap = "( mA Tap 8 ) , "
                    Case "07"
                        tap = "( mA Tap 7 ) , "
                    Case "06"
                        tap = "( mA Tap 6 ) , "
                    Case "05"
                        tap = "( mA Tap 5 ) , "
                    Case "04"
                        tap = "( mA Tap 4 ) , "
                    Case "03"
                        tap = "( mA Tap 3 ) , "
                    Case "02"
                        tap = "( mA Tap 2 ) , "
                End Select

            End If



            TextBox54.Text = TextBox54.Text & Environment.NewLine & Environment.NewLine & tdate & ":  Accuracy Check  Bench 2 " & Environment.NewLine & Environment.NewLine & "Operator:    " & opertr & Environment.NewLine & "Output:      " & outp & "   25.0  mA   " & tap & "0 Deg ," & "15 Sec" _
                 & Environment.NewLine & Environment.NewLine & "Meter Type :  " & modl & "         S/N:  " & seral _
                 & Environment.NewLine & Environment.NewLine & "Console Reading = " & CDbl(WHinternal).ToString & " wh" & "   Voltage Multiplier " & mulitplier _
                 & Environment.NewLine & Environment.NewLine & "Checked Device :  " & CDbl(WHexternal) & " X 1000  = " & XX & " wh" _
                 & Environment.NewLine & Environment.NewLine & "Known errors (%)      Console Error: " & _Error & "     Checked Device Error: " & Estd _
                 & Environment.NewLine & Environment.NewLine & "Checked Device Error (from Console): ((" & Math.Round(Y, 4) & " - " & Math.Round(XX, 4) & ") / " & Math.Round(XX, 4) & ")  X   100 = " & Math.Round(Error_checked_Device, 3) _
                 & Environment.NewLine & "Corrected Error of Checked Device: " & Math.Round(Error_checked_Device, 3) & " - " & _Error & " = " & Math.Round(Corrected_error_checked_device, 3).ToString _
                 & Environment.NewLine & Environment.NewLine & "Accuracy Test Result: " & Math.Round(Corrected_error_checked_device, 3) & " - " & Estd & " = " & percent.ToString

            If percent > 0.1 Then

                TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Failed ***********************************"
            Else
                TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Passed ***********************************"

            End If
            TextBox54.Refresh()
            'Set the cursor to the end of the textbox.
            TextBox54.SelectionStart = TextBox54.TextLength
            'Scroll down to the cursor position.
            TextBox54.ScrollToCaret()

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""

            Unit = "WH"
            external = WHexternal.ToString
            Load = "12.5"
            powerfactor = "1.0"
            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If
            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If

            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  (" & _
                        "'" & Unit & "', " & _
                        "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()






        End If

DQDQ:

        Call teststep3()

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(110)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If




        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999









        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If
        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(100)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")

        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.


        With oSheet

            '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
            LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
            If Len(CBNUM) = 13 Then
                CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
            Else
                CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
            End If

            .cells(LastRow + 1, 1).Value = CTNUM
            .cells(LastRow + 1, 2).Value = WHinternal
            .cells(LastRow + 1, 3).Value = WHexternal
            .cells(LastRow + 1, 4).Value = "WH @ 25% @ PF"


            voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
            .cells(LastRow + 1, 5).Value = voltage



            .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
            .cells(LastRow + 1, 7).Value = 1000
            .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
            .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
            TextBox51.Text = .cells(LastRow + 1, 9).Text
            If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                TextBox51.BackColor = Color.Red
            Else
                TextBox51.BackColor = Color.Lime
            End If

            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

            Application.DoEvents()
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End With
        If ACCTESTFLAG = 1 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""
            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double

            Dim Error_checked_Device As Double = 0
            Dim Corrected_error_checked_device As Double = 0
            Dim Accuracy_test_result As Double = 0
            Dim XX As Double
            Dim Y As Double



            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[WH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(116) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 4)
                        WH100array(ii) = strFileName(iii)
                        Dim fstring As String = "12.5_0.5_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next
            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Etrue = xpercent
            Change = Etrue - Emeter
            Total = Change + econsole
            Radian = Math.Round(CDbl(WHinternal), 3)

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()

            'Select Case Vtap.Substring(1, 1)
            '    Case "1"
            '        mulitplier = "1.000"
            '    Case "3"
            '        mulitplier = "2.999"
            '    Case "5"
            '        mulitplier = "5.002"
            '    Case Else  'This should not happen I am putting in a large multiplier in case it does
            '        mulitplier = "1000"
            'End Select

            '*** Checked Device Error determined by Console  = (Check Device Reading X 1000 (convert from milliamp to amps) – Console Reading X the Voltage Multiplier(based on the voltage from the PRN) / Console Reading X the Voltage Multiplier(based on the voltage from the PRN)) X 100 
            XX = CDbl(WHexternal) * 1000
            Y = CDbl(WHinternal) * CDbl(mulitplier)

            Error_checked_Device = ((Y - XX) / XX) * 100

            '*** Corrected Error of the Checked device =  Checked Device Error determined by Console - %Econ(from MC Spreadsheet)
            Corrected_error_checked_device = Error_checked_Device - _Error

            '*** Accuracy Test Result = Corrected Error of the Checked device - P-E-01 6.1.1.2(1C) reference meter option results (to be determined by Steve)

            Accuracy_test_result = Corrected_error_checked_device - Estd

            percent = Math.Abs(Math.Round(Accuracy_test_result, 3)).ToString



            If Not Button13.BackColor = Color.Green Then
                Select Case rButton.Name

                    Case "RadioButton40"
                        tap = "( mA Tap 40 ) , "
                    Case "RadioButton39"
                        tap = "( mA Tap 39 ) , "
                    Case "RadioButton38"
                        tap = "( mA Tap 38 ) , "
                    Case "RadioButton37"
                        tap = "( mA Tap 37 ) , "
                    Case "RadioButton36"
                        tap = "( mA Tap 36 ) , "
                    Case "RadioButton35"
                        tap = "( mA Tap 35 ) , "
                    Case "RadioButton34"
                        tap = "( mA Tap 34 ) , "
                    Case "RadioButton33"
                        tap = "( mA Tap 33 ) , "
                    Case "RadioButton32"
                        tap = "( mA Tap 32 ) , "
                    Case "RadioButton31"
                        tap = "( mA Tap 31 ) , "
                    Case "RadioButton30"
                        tap = "( mA Tap 30 ) , "
                    Case "RadioButton29"
                        tap = "( mA Tap 29 ) , "
                    Case "RadioButton28"
                        tap = "( mA Tap 28 ) , "
                    Case "RadioButton27"
                        tap = "( mA Tap 27 ) , "
                    Case "RadioButton26"
                        tap = "( mA Tap 26 ) , "
                    Case "RadioButton25"
                        tap = "( mA Tap 25 ) , "
                    Case "RadioButton24"
                        tap = "( mA Tap 24 ) , "
                    Case "RadioButton23"
                        tap = "( mA Tap 23 ) , "
                    Case "RadioButton22"
                        tap = "( mA Tap 22 ) , "
                    Case "RadioButton21"
                        tap = "( mA Tap 21 ) , "
                    Case "RadioButton20"
                        tap = "( mA Tap 20 ) , "
                    Case "RadioButton19"
                        tap = "( mA Tap 19 ) , "
                    Case "RadioButton18"
                        tap = "( mA Tap 18 ) , "
                    Case "RadioButton17"
                        tap = "( mA Tap 17 ) , "
                    Case "RadioButton16"
                        tap = "( mA Tap 16 ) , "
                    Case "RadioButton15"
                        tap = "( mA Tap 15 ) , "
                    Case "RadioButton14"
                        tap = "( mA Tap 14 ) , "
                    Case "RadioButton13"
                        tap = "( mA Tap 13 ) , "
                    Case "RadioButton12"
                        tap = "( mA Tap 12 ) , "
                    Case "RadioButton11"
                        tap = "( mA Tap 11 ) , "
                    Case "RadioButton10"
                        tap = "( mA Tap 30 ) , "
                    Case "RadioButton9"
                        tap = "( mA Tap 9 ) , "
                    Case "RadioButton8"
                        tap = "( mA Tap 8 ) , "
                    Case "RadioButton7"
                        tap = "( mA Tap 7 ) , "
                    Case "RadioButton6"
                        tap = "( mA Tap 6 ) , "
                    Case "RadioButton5"
                        tap = "( mA Tap 5 ) , "
                    Case "RadioButton4"
                        tap = "( mA Tap 4 ) , "
                    Case "RadioButton3"
                        tap = "( mA Tap 3 ) , "
                    Case "RadioButton2"
                        tap = "( mA Tap 2 ) , "
                End Select

            Else
                numtap = CInt(numtap) + 1
                numtap = numtap.ToString
                If Len(numtap) = 1 Then
                    numtap = "0" & numtap
                End If
                Select Case numtap

                    Case "40"
                        tap = "( mA Tap 40 ) , "
                    Case "39"
                        tap = "( mA Tap 39 ) , "
                    Case "38"
                        tap = "( mA Tap 38 ) , "
                    Case "37"
                        tap = "( mA Tap 37 ) , "
                    Case "36"
                        tap = "( mA Tap 36 ) , "
                    Case "35"
                        tap = "( mA Tap 35 ) , "
                    Case "34"
                        tap = "( mA Tap 34 ) , "
                    Case "33"
                        tap = "( mA Tap 33 ) , "
                    Case "32"
                        tap = "( mA Tap 32 ) , "
                    Case "31"
                        tap = "( mA Tap 31 ) , "
                    Case "30"
                        tap = "( mA Tap 30 ) , "
                    Case "29"
                        tap = "( mA Tap 29 ) , "
                    Case "28"
                        tap = "( mA Tap 28 ) , "
                    Case "27"
                        tap = "( mA Tap 27 ) , "
                    Case "26"
                        tap = "( mA Tap 26 ) , "
                    Case "25"
                        tap = "( mA Tap 25 ) , "
                    Case "24"
                        tap = "( mA Tap 24 ) , "
                    Case "23"
                        tap = "( mA Tap 23 ) , "
                    Case "22"
                        tap = "( mA Tap 22 ) , "
                    Case "21"
                        tap = "( mA Tap 21 ) , "
                    Case "20"
                        tap = "( mA Tap 20 ) , "
                    Case "19"
                        tap = "( mA Tap 19 ) , "
                    Case "18"
                        tap = "( mA Tap 18 ) , "
                    Case "17"
                        tap = "( mA Tap 17 ) , "
                    Case "16"
                        tap = "( mA Tap 16 ) , "
                    Case "15"
                        tap = "( mA Tap 15 ) , "
                    Case "14"
                        tap = "( mA Tap 14 ) , "
                    Case "13"
                        tap = "( mA Tap 13 ) , "
                    Case "12"
                        tap = "( mA Tap 12 ) , "
                    Case "11"
                        tap = "( mA Tap 11 ) , "
                    Case "10"
                        tap = "( mA Tap 30 ) , "
                    Case "09"
                        tap = "( mA Tap 9 ) , "
                    Case "08"
                        tap = "( mA Tap 8 ) , "
                    Case "07"
                        tap = "( mA Tap 7 ) , "
                    Case "06"
                        tap = "( mA Tap 6 ) , "
                    Case "05"
                        tap = "( mA Tap 5 ) , "
                    Case "04"
                        tap = "( mA Tap 4 ) , "
                    Case "03"
                        tap = "( mA Tap 3 ) , "
                    Case "02"
                        tap = "( mA Tap 2 ) , "
                End Select

            End If

            TextBox54.Text = TextBox54.Text & Environment.NewLine & Environment.NewLine & tdate & ":  Accuracy Check  Bench 2 " & Environment.NewLine & Environment.NewLine & "Operator:    " & opertr & Environment.NewLine & "Output:      " & outp & "   25.0  mA   " & tap & "60 Deg ," & "15 Sec" _
             & Environment.NewLine & Environment.NewLine & "Meter Type :  " & modl & "         S/N:  " & seral _
             & Environment.NewLine & Environment.NewLine & "Console Reading = " & CDbl(WHinternal).ToString & " wh" & "   Voltage Multiplier " & mulitplier _
             & Environment.NewLine & Environment.NewLine & "Checked Device :  " & CDbl(WHexternal) & " X 1000  = " & XX & " wh" _
             & Environment.NewLine & Environment.NewLine & "Known errors (%)      Console Error: " & _Error & "     Checked Device Error: " & Estd _
             & Environment.NewLine & Environment.NewLine & "Checked Device Error (from Console): ((" & Math.Round(Y, 4) & " - " & Math.Round(XX, 4) & ") / " & Math.Round(XX, 4) & ")  X   100 = " & Math.Round(Error_checked_Device, 3) _
             & Environment.NewLine & "Corrected Error of Checked Device: " & Math.Round(Error_checked_Device, 3) & " - " & _Error & " = " & Math.Round(Corrected_error_checked_device, 3).ToString _
             & Environment.NewLine & Environment.NewLine & "Accuracy Test Result: " & Math.Round(Corrected_error_checked_device, 3) & " - " & Estd & " = " & percent.ToString

            If percent > 0.1 Then

                TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Failed ***********************************"
            Else
                TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Passed ***********************************"

            End If

            TextBox54.Refresh()
            'Set the cursor to the end of the textbox.
            TextBox54.SelectionStart = TextBox54.TextLength
            'Scroll down to the cursor position.
            TextBox54.ScrollToCaret()


            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            Unit = "WH"
            external = WHexternal.ToString
            Load = "12.5"
            powerfactor = "0.5"
            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If
            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If



            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  (" & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()

        End If

        Call teststep4()

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(200)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If
        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999

        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If

        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(250)
        ComboBox1.Text = ""
        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")

        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_VAH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_VAH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.

        With oSheet
            '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
            LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
            If Len(CBNUM) = 13 Then
                CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
            Else
                CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
            End If

            .cells(LastRow + 1, 1).Value = CTNUM
            .cells(LastRow + 1, 2).Value = WHinternal
            .cells(LastRow + 1, 3).Value = WHexternal
            .cells(LastRow + 1, 4).Value = "VAH @ 25% @ PF"

            voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
            .cells(LastRow + 1, 5).Value = voltage

            .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
            .cells(LastRow + 1, 7).Value = 1000
            .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
            .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
            TextBox52.Text = .cells(LastRow + 1, 9).Text
            If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
                TextBox52.BackColor = Color.Red
            Else
                TextBox52.BackColor = Color.Lime
            End If

            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

            Application.DoEvents()
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End With


        If ACCTESTFLAG = 1 Then
            Dim tdate As DateTime
            Dim outp As String
            Dim opertr As String
            Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
            Dim tap As String = ""
            Dim reading As String = Math.Round(CDbl(WHinternal))
            Dim modl As String = Model
            Dim seral As String = Serial
            Dim xpercent As Double = 0
            Dim percent As String = ""
            Dim xeconsole As Double = 0
            Dim econsole As String = ""
            Dim i As Integer = 0
            Dim numtap As String = ""
            Dim Vtap As String = ""
            Dim Emeas As String = ""
            Dim Estd As String = ""
            Dim _Error As String = ""
            Dim mulitplier As String = ""
            Dim xEtrue As Double = 0
            Dim Emeter As Double = 0
            Dim Etrue As Double = 0
            Dim Change As Double = 0
            Dim Total As Double = 0
            Dim Radian As Double

            Dim Error_checked_Device As Double = 0
            Dim Corrected_error_checked_device As Double = 0
            Dim Accuracy_test_result As Double = 0
            Dim XX As Double
            Dim Y As Double




            tdate = DateTime.Now
            outp = TextBox2.Text
            opertr = initials
            modl = modl.Replace(ControlChars.NullChar, "")
            seral = seral.Replace(ControlChars.NullChar, "")

            If Button13.BackColor = Color.Green Then
                numtap = Mid(CBNUM, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            Else
                numtap = Mid(rButton.Name, 12, 2)
                numtap = Trim(numtap)
                numtap = CInt(numtap) - 1
            End If
            numtap = numtap.ToString
            If Len(numtap) = 1 Then
                numtap = "0" & numtap
            End If
            For Each x As String In strFileName
                If x.Equals("[VAH100_" & outp & "]") Then
                    Dim index1 As Integer = Array.IndexOf(strFileName, x)
                    Dim WH100array(40) As String
                    For ii As Integer = 0 To WH100array.Count - 1
                        Dim iii As Integer = ii + (index1 + 4)
                        WH100array(ii) = strFileName(iii)
                        Dim fstring As String = "12.5_0.5_M" & numtap
                        Dim sndhalf As String
                        Dim spresult() As String
                        If WH100array(ii).Contains(fstring) Then
                            sndhalf = WH100array(ii)
                            spresult = sndhalf.Split("=")
                            Dim results() As String
                            results = spresult(1).Split(",")
                            Vtap = Trim(results(0))
                            Emeas = results(1)
                            Estd = results(2)
                            _Error = results(3)
                        End If

                    Next

                End If
            Next
            _Error = CDbl(_Error)
            _Error = Math.Round(CDbl(_Error), 3)
            econsole = _Error
            Emeter = CDbl(Estd)
            Etrue = xpercent
            Change = Etrue - Emeter
            Total = Change + econsole
            Radian = Math.Round(CDbl(WHinternal), 3)

            mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()

            'Select Case Vtap
            '    Case "1"
            '        mulitplier = "1.000"
            '    Case "3"
            '        mulitplier = "2.999"
            '    Case "5"
            '        mulitplier = "5.002"
            'End Select

            '*** Checked Device Error determined by Console  = (Check Device Reading X 1000 (convert from milliamp to amps) – Console Reading X the Voltage Multiplier(based on the voltage from the PRN) / Console Reading X the Voltage Multiplier(based on the voltage from the PRN)) X 100 
            XX = CDbl(WHexternal) * 1000
            Y = CDbl(WHinternal) * CDbl(mulitplier)

            Error_checked_Device = ((Y - XX) / XX) * 100

            '*** Corrected Error of the Checked device =  Checked Device Error determined by Console - %Econ(from MC Spreadsheet)
            Corrected_error_checked_device = Error_checked_Device - _Error

            '*** Accuracy Test Result = Corrected Error of the Checked device - P-E-01 6.1.1.2(1C) reference meter option results (to be determined by Steve)

            Accuracy_test_result = Corrected_error_checked_device - Estd

            percent = Math.Abs(Math.Round(Accuracy_test_result, 3)).ToString

            If Not Button13.BackColor = Color.Green Then
                Select Case rButton.Name

                    Case "RadioButton40"
                        tap = "( mA Tap 40 ) , "
                    Case "RadioButton39"
                        tap = "( mA Tap 39 ) , "
                    Case "RadioButton38"
                        tap = "( mA Tap 38 ) , "
                    Case "RadioButton37"
                        tap = "( mA Tap 37 ) , "
                    Case "RadioButton36"
                        tap = "( mA Tap 36 ) , "
                    Case "RadioButton35"
                        tap = "( mA Tap 35 ) , "
                    Case "RadioButton34"
                        tap = "( mA Tap 34 ) , "
                    Case "RadioButton33"
                        tap = "( mA Tap 33 ) , "
                    Case "RadioButton32"
                        tap = "( mA Tap 32 ) , "
                    Case "RadioButton31"
                        tap = "( mA Tap 31 ) , "
                    Case "RadioButton30"
                        tap = "( mA Tap 30 ) , "
                    Case "RadioButton29"
                        tap = "( mA Tap 29 ) , "
                    Case "RadioButton28"
                        tap = "( mA Tap 28 ) , "
                    Case "RadioButton27"
                        tap = "( mA Tap 27 ) , "
                    Case "RadioButton26"
                        tap = "( mA Tap 26 ) , "
                    Case "RadioButton25"
                        tap = "( mA Tap 25 ) , "
                    Case "RadioButton24"
                        tap = "( mA Tap 24 ) , "
                    Case "RadioButton23"
                        tap = "( mA Tap 23 ) , "
                    Case "RadioButton22"
                        tap = "( mA Tap 22 ) , "
                    Case "RadioButton21"
                        tap = "( mA Tap 21 ) , "
                    Case "RadioButton20"
                        tap = "( mA Tap 20 ) , "
                    Case "RadioButton19"
                        tap = "( mA Tap 19 ) , "
                    Case "RadioButton18"
                        tap = "( mA Tap 18 ) , "
                    Case "RadioButton17"
                        tap = "( mA Tap 17 ) , "
                    Case "RadioButton16"
                        tap = "( mA Tap 16 ) , "
                    Case "RadioButton15"
                        tap = "( mA Tap 15 ) , "
                    Case "RadioButton14"
                        tap = "( mA Tap 14 ) , "
                    Case "RadioButton13"
                        tap = "( mA Tap 13 ) , "
                    Case "RadioButton12"
                        tap = "( mA Tap 12 ) , "
                    Case "RadioButton11"
                        tap = "( mA Tap 11 ) , "
                    Case "RadioButton10"
                        tap = "( mA Tap 30 ) , "
                    Case "RadioButton9"
                        tap = "( mA Tap 9 ) , "
                    Case "RadioButton8"
                        tap = "( mA Tap 8 ) , "
                    Case "RadioButton7"
                        tap = "( mA Tap 7 ) , "
                    Case "RadioButton6"
                        tap = "( mA Tap 6 ) , "
                    Case "RadioButton5"
                        tap = "( mA Tap 5 ) , "
                    Case "RadioButton4"
                        tap = "( mA Tap 4 ) , "
                    Case "RadioButton3"
                        tap = "( mA Tap 3 ) , "
                    Case "RadioButton2"
                        tap = "( mA Tap 2 ) , "
                End Select

            Else
                numtap = CInt(numtap) + 1
                numtap = numtap.ToString
                If Len(numtap) = 1 Then
                    numtap = "0" & numtap
                End If
                Select Case numtap

                    Case "40"
                        tap = "( mA Tap 40 ) , "
                    Case "39"
                        tap = "( mA Tap 39 ) , "
                    Case "38"
                        tap = "( mA Tap 38 ) , "
                    Case "37"
                        tap = "( mA Tap 37 ) , "
                    Case "36"
                        tap = "( mA Tap 36 ) , "
                    Case "35"
                        tap = "( mA Tap 35 ) , "
                    Case "34"
                        tap = "( mA Tap 34 ) , "
                    Case "33"
                        tap = "( mA Tap 33 ) , "
                    Case "32"
                        tap = "( mA Tap 32 ) , "
                    Case "31"
                        tap = "( mA Tap 31 ) , "
                    Case "30"
                        tap = "( mA Tap 30 ) , "
                    Case "29"
                        tap = "( mA Tap 29 ) , "
                    Case "28"
                        tap = "( mA Tap 28 ) , "
                    Case "27"
                        tap = "( mA Tap 27 ) , "
                    Case "26"
                        tap = "( mA Tap 26 ) , "
                    Case "25"
                        tap = "( mA Tap 25 ) , "
                    Case "24"
                        tap = "( mA Tap 24 ) , "
                    Case "23"
                        tap = "( mA Tap 23 ) , "
                    Case "22"
                        tap = "( mA Tap 22 ) , "
                    Case "21"
                        tap = "( mA Tap 21 ) , "
                    Case "20"
                        tap = "( mA Tap 20 ) , "
                    Case "19"
                        tap = "( mA Tap 19 ) , "
                    Case "18"
                        tap = "( mA Tap 18 ) , "
                    Case "17"
                        tap = "( mA Tap 17 ) , "
                    Case "16"
                        tap = "( mA Tap 16 ) , "
                    Case "15"
                        tap = "( mA Tap 15 ) , "
                    Case "14"
                        tap = "( mA Tap 14 ) , "
                    Case "13"
                        tap = "( mA Tap 13 ) , "
                    Case "12"
                        tap = "( mA Tap 12 ) , "
                    Case "11"
                        tap = "( mA Tap 11 ) , "
                    Case "10"
                        tap = "( mA Tap 30 ) , "
                    Case "09"
                        tap = "( mA Tap 9 ) , "
                    Case "08"
                        tap = "( mA Tap 8 ) , "
                    Case "07"
                        tap = "( mA Tap 7 ) , "
                    Case "06"
                        tap = "( mA Tap 6 ) , "
                    Case "05"
                        tap = "( mA Tap 5 ) , "
                    Case "04"
                        tap = "( mA Tap 4 ) , "
                    Case "03"
                        tap = "( mA Tap 3 ) , "
                    Case "02"
                        tap = "( mA Tap 2 ) , "
                End Select
            End If



            TextBox54.Text = TextBox54.Text & Environment.NewLine & Environment.NewLine & tdate & ":  Accuracy Check  Bench 2 " & Environment.NewLine & Environment.NewLine & "Operator:    " & opertr & Environment.NewLine & "Output:      " & outp & "   25.0  mA   " & tap & "60 Deg ," & "15 Sec" _
                            & Environment.NewLine & Environment.NewLine & "Meter Type :  " & modl & "         S/N:  " & seral _
                            & Environment.NewLine & Environment.NewLine & "Console Reading = " & Mid(CDbl(WHinternal), 1, 5).ToString & " Vah" & "   Voltage Multiplier " & mulitplier _
                            & Environment.NewLine & Environment.NewLine & "Checked Device :  " & CDbl(WHexternal) & " X 1000  = " & XX & " Vah" _
                            & Environment.NewLine & Environment.NewLine & "Known errors (%)      Console Error: " & _Error & "     Checked Device Error: " & Estd _
                            & Environment.NewLine & Environment.NewLine & "Checked Device Error (from Console): ((" & Math.Round(Y, 4) & " - " & Math.Round(XX, 4) & ") / " & Math.Round(XX, 4) & ")  X   100 = " & Math.Round(Error_checked_Device, 3) _
                            & Environment.NewLine & "Corrected Error of Checked Device: " & Math.Round(Error_checked_Device, 3) & " - " & _Error & " = " & Math.Round(Corrected_error_checked_device, 3).ToString _
                            & Environment.NewLine & Environment.NewLine & "Accuracy Test Result: " & Math.Round(Corrected_error_checked_device, 3) & " - " & Estd & " = " & percent.ToString

            If percent > 0.1 Then

                TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Failed ***********************************"
            Else
                TextBox54.Text = TextBox54.Text & Environment.NewLine & "************************************ Passed ***********************************"

            End If


            TextBox54.Refresh()
            'Set the cursor to the end of the textbox.
            TextBox54.SelectionStart = TextBox54.TextLength
            'Scroll down to the cursor position.
            TextBox54.ScrollToCaret()

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            Unit = "VAh"
            external = WHexternal.ToString
            Load = "12.5"
            powerfactor = "0.5"
            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If
            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If



            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  (" & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()

        End If
        If ACCTESTFLAG = 1 And RadioButton40.Checked = True And TextBox2.Text = "600" Then
            Call Button2_Click(0, System.EventArgs.Empty)
            'Exit Sub

        End If





        Application.DoEvents()
    End Sub
    Friend WithEvents Button12 As System.Windows.Forms.Button
    Friend WithEvents TextBox48 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox47 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox46 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox45 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox44 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox43 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox42 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox41 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox40 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox39 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox38 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox37 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox36 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox35 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox34 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox33 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox32 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox31 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox30 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox29 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox28 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox27 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox26 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox25 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox24 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox23 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox22 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox21 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox20 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox19 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox18 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox17 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox16 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox15 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox14 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox13 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox12 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox11 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox10 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents SerialPort2 As System.IO.Ports.SerialPort
    Public Sub readandfill()
        Dim Portnames As String() = System.IO.Ports.SerialPort.GetPortNames
        Dim send As String
        Dim returnValues As String
        Dim f As String
        Dim G As String
        Dim S As String
        Dim u As Long
        Dim ST As String
        Dim ss As Integer


        Try

            SerialPort2 = New SerialPort
            If Portnames Is Nothing Then
                MsgBox("There are no Com Ports detected!")
                Me.Close()
            End If
            With SerialPort2   'DDDaemon change from (3) to (2)
                .PortName = Portnames(3)
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
            SerialPort2.Close()
            Threading.Thread.Sleep(200)

            Try
                SerialPort2.Open()
            Catch
                Return
            End Try


            If SerialPort2.IsOpen Then

                Try
                    send = "X01HS0100002884"
                    SerialPort2.WriteLine(send)
                    Threading.Thread.Sleep(200)

                Catch ex As Exception
                    MsgBox("Read " & ex.Message)
                End Try

            End If

            If SerialPort2.IsOpen Then

                Try
                    send = "N0170"
                    SerialPort2.WriteLine(send)
                    Threading.Thread.Sleep(200)

                Catch ex As Exception
                    MsgBox("Read " & ex.Message)
                End Try

            End If


            returnValues = SerialPort2.ReadLine()



            If SerialPort2.IsOpen Then

                Try
                    send = "N0170"
                    SerialPort2.WriteLine(send)
                    Threading.Thread.Sleep(200)

                Catch ex As Exception
                    MsgBox("Read " & ex.Message)
                End Try

            End If

            Threading.Thread.Sleep(5000)

            returnValues = SerialPort2.ReadLine()

            If returnValues.Length < 20 Then
                Return
            End If


            ss = 9

            For t = 14 To 166 Step +4
                u = t - 2
                S = ""

                f = Mid(returnValues, t, 2)
                G = Mid(returnValues, u, 2)
                S = f & G
                'S = "00F8"

                Try
                    ST = (Convert.ToInt32(S, 16) / 100).ToString("####0.000")
                Catch
                    ST = 0
                End Try
                ss = ss + 1


                Select Case t
                    Case 14

                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox10.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox10.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox10.BackColor = Color.White
                                                        End If
                                                    End Sub))

                    Case 18
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox11.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox11.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox11.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 22
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox12.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox12.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox12.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 26
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox13.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox13.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox13.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 30
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox14.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox14.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox14.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 34
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox15.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox15.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox15.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 38
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox16.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox16.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox16.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 42
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox17.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox17.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox17.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 46
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox18.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox18.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox18.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 50
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox19.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox19.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox19.BackColor = Color.White
                                                        End If
                                                    End Sub))
                    Case 54
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox20.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox20.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox20.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 58
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox21.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox21.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox21.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 62
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox22.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox22.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox22.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 66
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox23.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox23.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox23.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 70
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox24.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox24.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox24.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 74
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox25.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox25.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox25.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 78
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox26.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox26.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox26.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 82
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox27.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox27.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox27.BackColor = Color.White
                                                        End If
                                                    End Sub))
                    Case 86
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox28.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox28.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox28.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 90
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox29.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox29.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox29.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 94
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox30.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox30.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox30.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 98
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox31.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox31.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox31.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 102
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox32.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox32.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox32.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 106
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox33.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox33.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox33.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 110
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox34.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox34.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox34.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 114
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox35.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox35.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox35.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 118
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox36.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox36.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox36.BackColor = Color.White
                                                        End If
                                                    End Sub))
                    Case 122
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox37.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox37.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox37.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 126
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox38.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox38.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox38.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 130
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox39.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox39.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox39.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 134
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox40.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox40.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox40.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 138
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox41.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox41.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox41.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 142
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox42.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox42.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox42.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 146
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox43.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox43.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox43.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 150
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox44.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox44.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox44.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 154
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox45.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox45.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox45.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 158
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox46.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox46.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox46.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 162
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox47.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox47.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox47.BackColor = Color.White
                                                        End If

                                                    End Sub))
                    Case 166
                        Me.Invoke(New MethodInvoker(Sub()
                                                        Me.TextBox48.Text = ST
                                                        If CInt(ST) > 1 Then
                                                            Me.TextBox48.BackColor = Color.LimeGreen
                                                        Else
                                                            Me.TextBox48.BackColor = Color.White
                                                        End If

                                                    End Sub))
                End Select











                ST = ""
            Next t

            Application.DoEvents()







            SerialPort2.Close()

        Catch
        End Try
        SerialPort2.Close()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim Portnames As String() = System.IO.Ports.SerialPort.GetPortNames
        Dim send As String
        Dim returnValues As String
        Dim f As String
        Dim G As String
        Dim S As String
        Dim u As Long
        Dim ST As String
        Dim ss As Integer

        Return

        SerialPort2 = New SerialPort
        If Portnames Is Nothing Then
            MsgBox("There are no Com Ports detected!")
            Me.Close()
        End If
        With SerialPort2
            .PortName = Portnames(2)
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
        SerialPort2.Close()
        Threading.Thread.Sleep(200)

        Try
            SerialPort2.Open()
        Catch
            Return
        End Try




        If SerialPort2.IsOpen Then

            Try
                send = "X01HS0100002884"
                SerialPort2.WriteLine(send)
                Threading.Thread.Sleep(200)

            Catch ex As Exception
                MsgBox("Read " & ex.Message)
            End Try

        End If

        If SerialPort2.IsOpen Then

            Try
                send = "N0170"
                SerialPort2.WriteLine(send)
                Threading.Thread.Sleep(200)

            Catch ex As Exception
                MsgBox("Read " & ex.Message)
            End Try

        End If

        Threading.Thread.Sleep(2000)

        returnValues = SerialPort2.ReadLine()



        If SerialPort2.IsOpen Then

            Try
                send = "N0170"
                SerialPort2.WriteLine(send)
                Threading.Thread.Sleep(200)

            Catch ex As Exception
                MsgBox("Read " & ex.Message)
            End Try

        End If

        Threading.Thread.Sleep(3000)

        returnValues = SerialPort2.ReadLine()


        ss = 9

        For t = 14 To 166 Step +4
            u = t - 2
            f = Mid(returnValues, t, 2)
            G = Mid(returnValues, u, 2)
            S = f & G
            'S = "00F8"

            ST = (Convert.ToInt32(S, 16) / 100).ToString("####0.000")
            ss = ss + 1


            Select Case t
                Case 14
                    Application.DoEvents()
                    TextBox10.Text = ST
                Case 18
                    TextBox11.Text = ST
                Case 22
                    TextBox12.Text = ST
                Case 26
                    TextBox13.Text = ST
                Case 30
                    TextBox14.Text = ST
                Case 34
                    TextBox15.Text = ST
                Case 38
                    TextBox16.Text = ST
                Case 42
                    TextBox17.Text = ST
                Case 46
                    TextBox18.Text = ST
                Case 50
                    TextBox19.Text = ST
                Case 54
                    TextBox20.Text = ST
                Case 58
                    TextBox21.Text = ST
                Case 62
                    TextBox22.Text = ST
                Case 66
                    TextBox23.Text = ST
                Case 70
                    TextBox24.Text = ST
                Case 74
                    TextBox25.Text = ST
                Case 78
                    TextBox26.Text = ST
                Case 82
                    TextBox27.Text = ST
                Case 86
                    TextBox28.Text = ST
                Case 90
                    TextBox29.Text = ST
                Case 94
                    TextBox30.Text = ST
                Case 98
                    TextBox31.Text = ST
                Case 102
                    TextBox32.Text = ST
                Case 106
                    TextBox33.Text = ST
                Case 110
                    TextBox34.Text = ST
                Case 114
                    TextBox35.Text = ST
                Case 118
                    TextBox36.Text = ST
                Case 122
                    TextBox37.Text = ST
                Case 126
                    TextBox38.Text = ST
                Case 130
                    TextBox39.Text = ST
                Case 134
                    TextBox40.Text = ST
                Case 138
                    TextBox41.Text = ST
                Case 142
                    TextBox42.Text = ST
                Case 146
                    TextBox43.Text = ST
                Case 150
                    TextBox44.Text = ST
                Case 154
                    TextBox45.Text = ST
                Case 158
                    TextBox46.Text = ST
                Case 162
                    TextBox47.Text = ST
                Case 166
                    TextBox48.Text = ST
            End Select












        Next t









        SerialPort2.Close()
    End Sub
    Private Sub RLYOFF()
        'AxAdvDIO1.DeviceNumber = 2
        WriteDoChannel2(0, 22) 'rly2
        WriteDoChannel2(0, 21) 'rly3
        WriteDoChannel2(0, 20) 'rly4
        WriteDoChannel2(0, 19) 'rly5
        WriteDoChannel2(0, 18) 'rly6
        WriteDoChannel2(0, 17) 'rly7
        WriteDoChannel2(0, 16) 'rly8
        WriteDoChannel2(0, 15) 'rly9
        WriteDoChannel2(0, 14) 'rly10
        WriteDoChannel2(0, 13) 'rly11
        WriteDoChannel2(0, 12) 'rly12
        WriteDoChannel2(0, 11) 'rly13
        WriteDoChannel2(0, 10) 'rly14
        WriteDoChannel2(0, 9) 'rly15
        WriteDoChannel2(0, 8) 'rly16
        WriteDoChannel2(0, 7) 'rly17
        WriteDoChannel2(0, 6) 'rly18
        WriteDoChannel2(0, 5) 'rly19
        WriteDoChannel2(0, 4) 'rly20
        WriteDoChannel2(0, 3) 'rly21
        WriteDoChannel2(0, 2) 'rly22
        WriteDoChannel2(0, 1) 'rly21
        WriteDoChannel2(0, 0) 'rly24
        WriteDoChannel2(0, 47) 'rly25
        WriteDoChannel2(0, 46) 'rly26
        WriteDoChannel2(0, 45) 'rly27
        WriteDoChannel2(0, 44) 'rly28
        WriteDoChannel2(0, 43) 'rly29
        WriteDoChannel2(0, 42) 'rly30
        WriteDoChannel2(0, 41) 'rly31
        WriteDoChannel2(0, 40) 'rly32
        WriteDoChannel2(0, 39) 'rly33
        WriteDoChannel2(0, 38) 'rly34
        WriteDoChannel2(0, 37) 'rly35
        WriteDoChannel2(0, 36) 'rly36
        WriteDoChannel2(0, 35) 'rly37
        WriteDoChannel2(0, 34) 'rly38
        WriteDoChannel2(0, 33) 'rly39
        WriteDoChannel2(0, 32) 'rly40
        TextBox10.BackColor = Color.White
        TextBox10.Text = ""
        TextBox11.BackColor = Color.White
        TextBox11.Text = ""
        TextBox12.BackColor = Color.White
        TextBox12.Text = ""
        TextBox13.BackColor = Color.White
        TextBox13.Text = ""
        TextBox14.BackColor = Color.White
        TextBox14.Text = ""
        TextBox15.BackColor = Color.White
        TextBox15.Text = ""
        TextBox16.BackColor = Color.White
        TextBox16.Text = ""
        TextBox17.BackColor = Color.White
        TextBox17.Text = ""
        TextBox18.BackColor = Color.White
        TextBox18.Text = ""
        TextBox19.BackColor = Color.White
        TextBox19.Text = ""
        TextBox20.BackColor = Color.White
        TextBox20.Text = ""
        TextBox21.BackColor = Color.White
        TextBox21.Text = ""
        TextBox22.BackColor = Color.White
        TextBox22.Text = ""
        TextBox23.BackColor = Color.White
        TextBox23.Text = ""
        TextBox24.BackColor = Color.White
        TextBox24.Text = ""
        TextBox25.BackColor = Color.White
        TextBox25.Text = ""
        TextBox26.BackColor = Color.White
        TextBox26.Text = ""
        TextBox27.BackColor = Color.White
        TextBox27.Text = ""
        TextBox28.BackColor = Color.White
        TextBox28.Text = ""
        TextBox29.BackColor = Color.White
        TextBox29.Text = ""
        TextBox30.BackColor = Color.White
        TextBox30.Text = ""
        TextBox31.BackColor = Color.White
        TextBox31.Text = ""
        TextBox32.BackColor = Color.White
        TextBox32.Text = ""
        TextBox33.BackColor = Color.White
        TextBox33.Text = ""
        TextBox34.BackColor = Color.White
        TextBox34.Text = ""
        TextBox35.BackColor = Color.White
        TextBox35.Text = ""
        TextBox36.BackColor = Color.White
        TextBox36.Text = ""
        TextBox37.BackColor = Color.White
        TextBox37.Text = ""
        TextBox38.BackColor = Color.White
        TextBox38.Text = ""
        TextBox39.BackColor = Color.White
        TextBox39.Text = ""
        TextBox40.BackColor = Color.White
        TextBox40.Text = ""
        TextBox41.BackColor = Color.White
        TextBox41.Text = ""
        TextBox42.BackColor = Color.White
        TextBox42.Text = ""
        TextBox43.BackColor = Color.White
        TextBox43.Text = ""
        TextBox44.BackColor = Color.White
        TextBox44.Text = ""
        TextBox45.BackColor = Color.White
        TextBox45.Text = ""
        TextBox46.BackColor = Color.White
        TextBox46.Text = ""
        TextBox47.BackColor = Color.White
        TextBox47.Text = ""
        TextBox48.BackColor = Color.White
        TextBox48.Text = ""

    End Sub

    Private Sub RLYON()
        'AxAdvDIO1.DeviceNumber = 2
        WriteDoChannel2(1, 22) 'rly2
        WriteDoChannel2(1, 21) 'rly3
        WriteDoChannel2(1, 20) 'rly4
        WriteDoChannel2(1, 19) 'rly5
        WriteDoChannel2(1, 18) 'rly6
        WriteDoChannel2(1, 17) 'rly7
        WriteDoChannel2(1, 16) 'rly8
        WriteDoChannel2(1, 15) 'rly9
        WriteDoChannel2(1, 14) 'rly10
        WriteDoChannel2(1, 13) 'rly11
        WriteDoChannel2(1, 12) 'rly12
        WriteDoChannel2(1, 11) 'rly13
        WriteDoChannel2(1, 10) 'rly14
        WriteDoChannel2(1, 9) 'rly15
        WriteDoChannel2(1, 8) 'rly16
        WriteDoChannel2(1, 7) 'rly17
        WriteDoChannel2(1, 6) 'rly18
        WriteDoChannel2(1, 5) 'rly19
        WriteDoChannel2(1, 4) 'rly20
        WriteDoChannel2(1, 3) 'rly21
        WriteDoChannel2(1, 2) 'rly22
        WriteDoChannel2(1, 1) 'rly21
        WriteDoChannel2(1, 0) 'rly24
        WriteDoChannel2(1, 47) 'rly25
        WriteDoChannel2(1, 46) 'rly26
        WriteDoChannel2(1, 45) 'rly27
        WriteDoChannel2(1, 44) 'rly28
        WriteDoChannel2(1, 43) 'rly29
        WriteDoChannel2(1, 42) 'rly30
        WriteDoChannel2(1, 41) 'rly31
        WriteDoChannel2(1, 40) 'rly32
        WriteDoChannel2(1, 39) 'rly33
        WriteDoChannel2(1, 38) 'rly34
        WriteDoChannel2(1, 37) 'rly35
        WriteDoChannel2(1, 36) 'rly36
        WriteDoChannel2(1, 35) 'rly37
        WriteDoChannel2(1, 34) 'rly38
        WriteDoChannel2(1, 33) 'rly39
        WriteDoChannel2(1, 32) 'rly40



    End Sub
    Private Sub Body()
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)
        Dim IntCount(256) As Long
        Dim Voltage As String
        Dim FirstCharacter As Integer = TextBox8.Text.IndexOf("1    P")






        '''''sets control relays off

        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(0, 33)
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(0, 34)
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(0, 39)

        If Not ACCTESTFLAG = 1 Then
            TextBox8.Text = ""
            TextBox2.Text = ""
        End If

        'AxAdvDIO1.DeviceNumber = 0
        'AxAdvDIO1.WriteDoChannel(0, 22)
        'AxAdvDIO1.WriteDoChannel(0, 23)

        WriteDoChannel1(0, 33)
        WriteDoChannel1(0, 34)
        WriteDoChannel1(0, 39)
        WriteDoChannel0(0, 22)
        WriteDoChannel0(0, 23)



        TextBox3.BackColor = Color.White
        TextBox5.BackColor = Color.White
        TextBox6.BackColor = Color.White
        TextBox7.BackColor = Color.White
        TextBox2.BackColor = Color.White
        ''''' open relays



        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)
        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)
        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        'set FG420
        mbSession = CType(ResourceManager.GetLocalManager().Open("GPIB0::0::INSTR"), MessageBasedSession)

        '''''set Trigger to panel meter
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(1, 32)

        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(1, 34)
        ''MsgBox("set advantech board 1 ,34 power to emp")

        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(1, 39)
        'AxAdvDIO1.WriteDoChannel(0, 37)

        WriteDoChannel1(1, 32)

        WriteDoChannel1(1, 34)
        'MsgBox("set advantech board 1 ,34 power to emp")

        WriteDoChannel1(1, 39)
        WriteDoChannel1(0, 37)










        '''''================================================================if textbox 8 has a loaded PRN=================================================


        If Not ACCTESTFLAG = 1 Then
            If Mid(tboxprn, (FirstCharacter + 7), 3) = "120" Then
                GoTo 1231
            End If
        Else
            If TextBox2.Text = "120" Then
                GoTo 1231
            End If
        End If


        If Not ACCTESTFLAG = 1 Then

            If Mid(tboxprn, (FirstCharacter + 7), 3) = "208" Then
                GoTo 1232
            End If
        Else
            If TextBox2.Text = "208" Then
                GoTo 1232
            End If

        End If


        If Not ACCTESTFLAG = 1 Then
            If Mid(tboxprn, (FirstCharacter + 7), 3) = "240" Then
                GoTo 1233
            End If
        Else
            If TextBox2.Text = "240" Then
                GoTo 1233
            End If
        End If


        If Not ACCTESTFLAG = 1 Then
            If Mid(tboxprn, (FirstCharacter + 7), 3) = "277" Then
                GoTo 1234
            End If
        Else
            If TextBox2.Text = "277" Then
                GoTo 1234
            End If
        End If

        If Not ACCTESTFLAG = 1 Then
            If Mid(tboxprn, (FirstCharacter + 7), 3) = "347" Then
                GoTo 1235
            End If
        Else
            If TextBox2.Text = "347" Then
                GoTo 1235
            End If

        End If

        If Not ACCTESTFLAG = 1 Then
            If Mid(tboxprn, (FirstCharacter + 7), 3) = "416" Then
                GoTo 1236
            End If
        Else
            If TextBox2.Text = "416" Then
                GoTo 1236
            End If
        End If

        If Not ACCTESTFLAG = 1 Then
            If Mid(tboxprn, (FirstCharacter + 7), 3) = "480" Then
                GoTo 1237
            End If
        Else
            If TextBox2.Text = "480" Then
                GoTo 1237
            End If

        End If

        If Not ACCTESTFLAG = 1 Then
            If Mid(tboxprn, (FirstCharacter + 7), 3) = "600" Then
                GoTo 1238
            End If
        Else
            If TextBox2.Text = "600" Then
                GoTo 1238
            End If
        End If
        '*****       '----------------Get Voltage------------------------------------------------------------------------------------------------------
1231:
        tboxprn = ""
        TextBox8.Text = File.ReadAllText("C:\Bench\Console\mA cert files\mA120V100A.PRN")
        Voltage = Mid(TextBox8.Text, (FirstCharacter + 7), 3)
        TextBox2.Text = Voltage
        TextBox2.BackColor = Color.LimeGreen
        Thread.Sleep(200)

        ' Implemented November 2022 to handle the two different formula's required for the two processes
        ' If the Daily_Accuracy_Test is set to 1 that means the user called the Daily Accuracy Check from the File Menu
        ' If the Daily_Accuracy_Test is set to anything else that means it was called from the File Menu as Measurement Canada Certification Process
        If (Daily_Accuracy_Test = 1) Then
            'Call SetAfterPrn_Daily_Accuracy_Check()
            Call SetAfterPrn2()
        Else
            'Call SetAfterPrn_MC_Accuracy_Check()
            Call SetAfterPrn2()
            'Call SetAfterPrn()
        End If

        Thread.Sleep(200)
        If stopflag = 1 Then
            Exit Sub
        End If
        TextBox3.BackColor = Color.White
        TextBox5.BackColor = Color.White
        TextBox6.BackColor = Color.White
        TextBox7.BackColor = Color.White
        TextBox2.BackColor = Color.White

        ''################################## test 2#######################################################
        '''''change Prn
        If stopflag = 1 Then
            Exit Sub
        End If
        If ACCTESTFLAG = 1 Then
            Exit Sub
        End If
1232:

        tboxprn = ""
        TextBox8.Text = File.ReadAllText("C:\Bench\Console\mA cert files\mA208V100A.PRN")
        FirstCharacter = TextBox8.Text.IndexOf("1    P")
        Voltage = Mid(TextBox8.Text, (FirstCharacter + 7), 3)

        TextBox2.Text = Voltage
        TextBox2.BackColor = Color.LimeGreen
        Thread.Sleep(200)

        ' Implemented November 2022 to handle the two different formula's required for the two processes
        ' If the Daily_Accuracy_Test is set to 1 that means the user called the Daily Accuracy Check from the File Menu
        ' If the Daily_Accuracy_Test is set to anything else that means it was called from the File Menu as Measurement Canada Certification Process
        If (Daily_Accuracy_Test = 1) Then
            Call SetAfterPrn_Daily_Accuracy_Check()
        Else
            Call SetAfterPrn_MC_Accuracy_Check()
            'Call SetAfterPrn()
        End If

        Thread.Sleep(200)
        TextBox2.BackColor = Color.White

        '''''change Prn
        If stopflag = 1 Then
            Exit Sub
        End If
        If ACCTESTFLAG = 1 Then
            Exit Sub
        End If
1233:

        tboxprn = ""
        TextBox8.Text = File.ReadAllText("C:\Bench\Console\mA cert files\mA240V100A.PRN")
        FirstCharacter = TextBox8.Text.IndexOf("1    P")
        Voltage = Mid(TextBox8.Text, (FirstCharacter + 7), 3)

        TextBox2.Text = Voltage
        TextBox2.BackColor = Color.LimeGreen
        Thread.Sleep(200)

        ' Implemented November 2022 to handle the two different formula's required for the two processes
        ' If the Daily_Accuracy_Test is set to 1 that means the user called the Daily Accuracy Check from the File Menu
        ' If the Daily_Accuracy_Test is set to anything else that means it was called from the File Menu as Measurement Canada Certification Process
        If (Daily_Accuracy_Test = 1) Then
            Call SetAfterPrn_Daily_Accuracy_Check()
        Else
            Call SetAfterPrn_MC_Accuracy_Check()
            'Call SetAfterPrn()
        End If

        Thread.Sleep(200)
        TextBox2.BackColor = Color.White

        '''''change Prn
        If stopflag = 1 Then
            Exit Sub
        End If
        If ACCTESTFLAG = 1 Then
            Exit Sub
        End If
1234:

        tboxprn = ""
        TextBox8.Text = File.ReadAllText("C:\Bench\Console\mA cert files\mA277V100A.PRN")
        FirstCharacter = TextBox8.Text.IndexOf("1    P")

        Voltage = Mid(TextBox8.Text, (FirstCharacter + 7), 3)
        TextBox2.Text = Voltage
        TextBox2.BackColor = Color.LimeGreen
        Thread.Sleep(200)

        ' Implemented November 2022 to handle the two different formula's required for the two processes
        ' If the Daily_Accuracy_Test is set to 1 that means the user called the Daily Accuracy Check from the File Menu
        ' If the Daily_Accuracy_Test is set to anything else that means it was called from the File Menu as Measurement Canada Certification Process
        If (Daily_Accuracy_Test = 1) Then
            Call SetAfterPrn_Daily_Accuracy_Check()
        Else
            Call SetAfterPrn_MC_Accuracy_Check()
            'Call SetAfterPrn()
        End If

        Thread.Sleep(200)
        TextBox2.BackColor = Color.White

        '''''change Prn
        If stopflag = 1 Then
            Exit Sub
        End If
        If ACCTESTFLAG = 1 Then
            Exit Sub
        End If

1235:


        tboxprn = ""
        TextBox8.Text = File.ReadAllText("C:\Bench\Console\mA cert files\mA347V100A.PRN")
        FirstCharacter = TextBox8.Text.IndexOf("1    P")

        Voltage = Mid(TextBox8.Text, (FirstCharacter + 7), 3)
        TextBox2.Text = Voltage
        TextBox2.BackColor = Color.LimeGreen
        Thread.Sleep(200)

        ' Implemented November 2022 to handle the two different formula's required for the two processes
        ' If the Daily_Accuracy_Test is set to 1 that means the user called the Daily Accuracy Check from the File Menu
        ' If the Daily_Accuracy_Test is set to anything else that means it was called from the File Menu as Measurement Canada Certification Process
        If (Daily_Accuracy_Test = 1) Then
            Call SetAfterPrn_Daily_Accuracy_Check()
        Else
            Call SetAfterPrn_MC_Accuracy_Check()
            'Call SetAfterPrn()
        End If

        Thread.Sleep(200)

        '''''change Prn
        If stopflag = 1 Then
            Exit Sub
        End If
        If ACCTESTFLAG = 1 Then
            Exit Sub
        End If
1236:

        tboxprn = ""
        TextBox8.Text = File.ReadAllText("C:\Bench\Console\mA cert files\mA416V100A.PRN")
        FirstCharacter = TextBox8.Text.IndexOf("1    P")

        Voltage = Mid(TextBox8.Text, (FirstCharacter + 7), 3)
        TextBox2.Text = Voltage
        TextBox2.BackColor = Color.LimeGreen
        Thread.Sleep(200)

        ' Implemented November 2022 to handle the two different formula's required for the two processes
        ' If the Daily_Accuracy_Test is set to 1 that means the user called the Daily Accuracy Check from the File Menu
        ' If the Daily_Accuracy_Test is set to anything else that means it was called from the File Menu as Measurement Canada Certification Process
        If (Daily_Accuracy_Test = 1) Then
            Call SetAfterPrn_Daily_Accuracy_Check()
        Else
            Call SetAfterPrn_MC_Accuracy_Check()
            'Call SetAfterPrn()
        End If

        Thread.Sleep(200)
        TextBox2.BackColor = Color.White
        If stopflag = 1 Then
            Exit Sub
        End If
        '''''change Prn
        If stopflag = 1 Then
            Exit Sub
        End If

        If ACCTESTFLAG = 1 Then
            Exit Sub
        End If
1237:

        tboxprn = ""
        TextBox8.Text = File.ReadAllText("C:\Bench\Console\mA cert files\mA480V100A.PRN")
        FirstCharacter = TextBox8.Text.IndexOf("1    P")

        Voltage = Mid(TextBox8.Text, (FirstCharacter + 7), 3)
        TextBox2.Text = Voltage
        TextBox2.BackColor = Color.LimeGreen
        Thread.Sleep(200)

        ' Implemented November 2022 to handle the two different formula's required for the two processes
        ' If the Daily_Accuracy_Test is set to 1 that means the user called the Daily Accuracy Check from the File Menu
        ' If the Daily_Accuracy_Test is set to anything else that means it was called from the File Menu as Measurement Canada Certification Process
        If (Daily_Accuracy_Test = 1) Then
            Call SetAfterPrn_Daily_Accuracy_Check()
        Else
            Call SetAfterPrn_MC_Accuracy_Check()
            'Call SetAfterPrn()
        End If

        Thread.Sleep(200)
        TextBox2.BackColor = Color.White
        '''''change Prn
        If stopflag = 1 Then
            Exit Sub
        End If
        If ACCTESTFLAG = 1 Then
            Exit Sub
        End If

1238:

        tboxprn = ""
        TextBox8.Text = File.ReadAllText("C:\Bench\Console\mA cert files\mA600V100A.PRN")
        FirstCharacter = TextBox8.Text.IndexOf("1    P")

        Voltage = Mid(TextBox8.Text, (FirstCharacter + 7), 3)
        TextBox2.Text = Voltage
        TextBox2.BackColor = Color.LimeGreen
        Thread.Sleep(200)

        ' Implemented November 2022 to handle the two different formula's required for the two processes
        ' If the Daily_Accuracy_Test is set to 1 that means the user called the Daily Accuracy Check from the File Menu
        ' If the Daily_Accuracy_Test is set to anything else that means it was called from the File Menu as Measurement Canada Certification Process
        If (Daily_Accuracy_Test = 1) Then
            Call SetAfterPrn_Daily_Accuracy_Check()
        Else
            Call SetAfterPrn_MC_Accuracy_Check()
            'Call SetAfterPrn()
        End If

        Thread.Sleep(200)
        TextBox2.BackColor = Color.White




    End Sub
    Friend WithEvents Button3 As System.Windows.Forms.Button


    Public Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox53.Text = ""
        For Me.repx = 1 To 4
            RadioButton2.Checked = True
            TextBox53.Text = repx
            TextBox53.Refresh()
            Thread.Sleep(300)
            Call SETRELAYS_Click(0, e)
            If stopflag = 1 Then
                Exit Sub
            End If

        Next

        Me.repx = 0
        TextBox53.Text = ""



    End Sub

    Public Sub Button13_Click(sender As Object, e As EventArgs)

        '''''%0 amp overcurrent test Caution



        Dim test As String
        Dim style = MsgBoxStyle.YesNo
        Dim result As DialogResult = MsgBox("This will produce 50 amps to the output of the Vallhla. It was intended to be used to test the 10 times over current of the single point meter. Are you sure you want to continue?", MsgBoxStyle.YesNo, "CAUTION")
        If (result = DialogResult.Yes) Then





            'AxAdvDIO1.DeviceNumber = 1
            'AxAdvDIO1.WriteDoChannel(1, 39)
            WriteDoChannel1(1, 39)
            mbSession = CType(ResourceManager.GetLocalManager.Open("GPIB0::0::INSTR"), MessageBasedSession)
            mbSession.Write("IDEN?")
            test = mbSession.ReadString()

            If Len(test) > 1 Then
                'mbSession.Write("CHAN 2  FREQ 60.0000")
                mbSession.Write("CHAN 2  FREQ 60.4000")
                mbSession.Write("CHAN 1 SPH +01.000")
                mbSession.Write("CHAN 1  FREQ 60.0000")
                mbSession.Write("CHAN 1 Rang 1V AMR 0.5032")
            End If


            mbSession = CType(ResourceManager.GetLocalManager().Open("GPIB0::0::INSTR"), MessageBasedSession)
            mbSession.Write("CHAN 1 OUTP ON ")
            mbSession.Write("CHAN 2 OUTP ON ")
            'AxAdvDIO1.DeviceNumber = 1
            'AxAdvDIO1.WriteDoChannel(1, 33)
            WriteDoChannel1(1, 33)
            Threading.Thread.Sleep(100)
            'AxAdvDIO1.DeviceNumber = 1
            ' AxAdvDIO1.WriteDoChannel(0, 33)
            WriteDoChannel1(0, 33)
            mbSession.Write("CHAN 1 OUTP OFF ")
            mbSession.Write("CHAN 2 OUTP OFF ")
        Else
            Exit Sub
        End If




    End Sub


    Friend WithEvents AxAdvDIO1 As AxAdvDIOLib.AxAdvDIO
    Friend WithEvents AxAdvDIO2 As AxAdvDIOLib.AxAdvDIO

    Public Sub btnOpenFile_Click(sender As Object, e As EventArgs) Handles btnOpenFile.Click
        If repx = 0 Then
            On Error Resume Next
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If

        If repx = 1 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat1temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If

        If repx = 2 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat2temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If
        If repx = 3 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat3temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If
        If repx = 4 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat4temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If
        releaseObject(oExcel)
        releaseObject(oBook)
        releaseObject(oSheet)

        Call Form7.Button1_Click(0, e)



    End Sub
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork

    End Sub
    Public Sub radiobuttonChecked()
        Dim ListOFNames() As String = {"RadioButton1", "", "FiveLbRadioButton", _
                                            "TenLbRadioButton", "MoreTenLbRadioButton"}

        ' manifest all controls.

        For Each IsAControl As Control In Controls

            For Each Child As Control In IsAControl.Controls


:
                ' if this is a radio button, do work with it.

                If TypeOf Child Is RadioButton Then



                    ' Should probably find out if the radio button is CHECKED right here. If not no point in next part.



                    ' set index number to match (array) list of names.

                    For i As Integer = 0 To ListOFNames.Length - 1

                        ' if current child name matches list name, do work with it.

                        If String.Equals(Child.Name, ListOFNames(i), StringComparison.InvariantCultureIgnoreCase) Then





                        End If

                    Next

                End If

            Next

        Next




    End Sub
    Friend WithEvents TextBox52 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox51 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox50 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox49 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label



    Public Sub Find_check(sender As Object, e As EventArgs)

        Dim c As Object
        Dim a As Integer
        a = Me.GroupBox1.Controls.Count

        For Each c In Me.GroupBox1.Controls
            If TypeOf (c) Is RadioButton Then
                Dim rb As RadioButton = CType(c, RadioButton)
                If rb.Checked = True Then
                    'MsgBox(CStr(rb.Name) & " is selected")
                    CBNUM = CStr(rb.Name)


                    Exit For
                End If
            End If
        Next
    End Sub
    Public Sub Find_check2(sender As Object, e As EventArgs)
        Dim z As Integer
        Dim a As Integer
        a = Me.GroupBox3.Controls.Count
        Dim Counter As Integer = 0
        For x = 0 To checkboxarray.Count - 1
            If Len(checkboxarray(x)) < 1 Then
                Counter += 1
            End If
        Next

        If checkboxarray.Length > 1 Then

            For z = 2 To checkboxarray.Length

                Array.Resize(checkboxarray, checkboxarray.Length - 1)

            Next
        End If
        For Each checkbox In Me.GroupBox3.Controls.OfType(Of CheckBox)()
            If checkbox.Checked = True Then
                'MsgBox(checkbox.Name)
                Dim newitem As String = checkbox.Name
                checkboxarray = checkboxarray.Concat({newitem}).ToArray

            End If
        Next


    End Sub










    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton5 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton6 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton7 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton8 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton9 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton10 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton11 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton12 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton13 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton14 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton15 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton16 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton17 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton18 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton19 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton20 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton21 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton22 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton23 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton24 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton25 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton26 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton27 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton28 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton29 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton30 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton31 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton32 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton33 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton34 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton35 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton36 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton37 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton38 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton39 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton40 As System.Windows.Forms.RadioButton

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub BackgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker2.DoWork

        Application.DoEvents()
        Call readandfill()
        If (Me.BackgroundWorker1.CancellationPending = True) Then
            e.Cancel = True
        End If

    End Sub


    Public Sub closeSession()
        stopflag = 1

        mbSession = CType(ResourceManager.GetLocalManager().Open("GPIB0::0::INSTR"), MessageBasedSession)
        mbSession.Write("CHAN 1 GATE OFF")
        Thread.Sleep(100)
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")

        'Turn off EMP power
        WriteDoChannel1(0, 34)

        RLYOFF()

        BackgroundWorker2.CancelAsync()
        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy = True
            Application.DoEvents()
            Threading.Thread.Sleep(300)
        End While

        mbSession.Dispose()

        Try
            oExcel.DisplayAlerts = False
            Dim dt As New DateTime
            dt = DateTime.Now
            Dim mt As String
            mt = String.Format("C:\MCMKARESULTS_{0:MMM_dd_yyyy_HH-mm}.xlsx", dt)
            oBook.Worksheets(1).SaveAs(mt.ToString())
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        Catch ex As Exception
        End Try

        Dim processes() As Process = Process.GetProcesses
        For p As Integer = processes.Count - 1 To 0 Step -1
            If processes(p).ProcessName = "EXCEL" Then
                processes(p).Kill()
            End If
        Next
    End Sub


    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label


    Private Sub FTestApp_(sender As Object, e As EventArgs) Handles MyBase.DoubleClick

        mbSession = CType(ResourceManager.GetLocalManager.Open("GPIB0::0::INSTR"), MessageBasedSession)
        mbSession.Write("CHAN 2 Rang 10V AMR 0.0001")
        mbSession.Write("CHAN 2 OUTP OFF ")
        Call closeSession()
        'mbSession.Write("SOURce2:FREQuency:CW 00HZ")
        'mbSession.Write("OUTPut2:STATe OFF\n\r")
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(0, 33)
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(0, 34)
        'AxAdvDIO1.DeviceNumber = 1
        'AxAdvDIO1.WriteDoChannel(0, 39)
        'AxAdvDIO1.DeviceNumber = 0
        'AxAdvDIO1.WriteDoChannel(0, 22)
        'AxAdvDIO1.WriteDoChannel(0, 23)
        WriteDoChannel1(0, 33)
        WriteDoChannel1(0, 34)
        WriteDoChannel1(0, 39)
        WriteDoChannel0(0, 22)
        WriteDoChannel0(0, 23)


        TextBox8.Text = ""
        TextBox2.Text = ""
        TextBox2.BackColor = Color.White
        TextBox3.BackColor = Color.White
        TextBox5.BackColor = Color.White
        TextBox6.BackColor = Color.White
        TextBox7.BackColor = Color.White
        TextBox10.BackColor = Color.White
        TextBox10.Text = ""
        TextBox11.BackColor = Color.White
        TextBox11.Text = ""
        TextBox12.BackColor = Color.White
        TextBox12.Text = ""
        TextBox13.BackColor = Color.White
        TextBox13.Text = ""
        TextBox14.BackColor = Color.White
        TextBox14.Text = ""
        TextBox15.BackColor = Color.White
        TextBox15.Text = ""
        TextBox16.BackColor = Color.White
        TextBox16.Text = ""
        TextBox18.BackColor = Color.White
        TextBox18.Text = ""
        TextBox19.BackColor = Color.White
        TextBox19.Text = ""
        TextBox20.BackColor = Color.White
        TextBox20.Text = ""
        TextBox21.BackColor = Color.White
        TextBox21.Text = ""
        TextBox22.BackColor = Color.White
        TextBox22.Text = ""
        TextBox23.BackColor = Color.White
        TextBox23.Text = ""
        TextBox24.BackColor = Color.White
        TextBox24.Text = ""
        TextBox25.BackColor = Color.White
        TextBox25.Text = ""
        TextBox26.BackColor = Color.White
        TextBox26.Text = ""
        TextBox27.BackColor = Color.White
        TextBox27.Text = ""
        TextBox28.BackColor = Color.White
        TextBox28.Text = ""
        TextBox29.BackColor = Color.White
        TextBox29.Text = ""
        TextBox30.BackColor = Color.White
        TextBox30.Text = ""
        TextBox31.BackColor = Color.White
        TextBox31.Text = ""
        TextBox32.BackColor = Color.White
        TextBox32.Text = ""
        TextBox33.BackColor = Color.White
        TextBox33.Text = ""
        TextBox34.BackColor = Color.White
        TextBox34.Text = ""
        TextBox35.BackColor = Color.White
        TextBox35.Text = ""
        TextBox36.BackColor = Color.White
        TextBox36.Text = ""
        TextBox37.BackColor = Color.White
        TextBox37.Text = ""
        TextBox38.BackColor = Color.White
        TextBox38.Text = ""
        TextBox39.BackColor = Color.White
        TextBox39.Text = ""
        TextBox40.BackColor = Color.White
        TextBox40.Text = ""
        TextBox41.BackColor = Color.White
        TextBox41.Text = ""
        TextBox42.BackColor = Color.White
        TextBox42.Text = ""
        TextBox43.BackColor = Color.White
        TextBox43.Text = ""
        TextBox44.BackColor = Color.White
        TextBox44.Text = ""
        TextBox45.BackColor = Color.White
        TextBox45.Text = ""
        TextBox46.BackColor = Color.White
        TextBox46.Text = ""
        TextBox47.BackColor = Color.White
        TextBox47.Text = ""
        TextBox48.BackColor = Color.White
        TextBox48.Text = ""

        RadRDAccumStop(comDevice)
        RadRDAccumReset(comDevice, 0)


        Call RLYOFF()



        ''Save the Workbook and quit Excel.


        If repx = 0 Then
            On Error Resume Next
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If

        If repx = 1 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat1temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If

        If repx = 2 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat2temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If
        If repx = 3 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat3temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If
        If repx = 4 Then
            oExcel.DisplayAlerts = False
            oBook.Worksheets(1).SaveAs("C:\MCMKARepeat4temp.xlsx")
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

        End If



        repx = 0
    End Sub
    Private Sub WriteDoChannel1(nStatus, nRegister)
        Dim nport As Integer
        Dim nbit As Integer

        '//34 = 4, 2

        nport = Int(nRegister / 8)
        nbit = Int(nRegister - (nport * 8))
        AxInstantDoCtrl1.WriteBit(nport, nbit, nStatus)
    End Sub

    Private Sub WriteDoChannel2(nStatus, nRegister)
        Dim nport As Integer
        Dim nbit As Integer

        '//34 = 4, 2

        nport = Int(nRegister / 8)
        nbit = Int(nRegister - (nport * 8))

        'AxInstantDoCtrl2.CreateControl()
        AxInstantDoCtrl2.WriteBit(nport, nbit, nStatus)
    End Sub

    Private Sub WriteDoChannel0(nStatus, nRegister)
        Dim nport As Integer
        Dim nbit As Integer

        '//34 = 4, 2

        nport = Int(nRegister / 8)
        nbit = Int(nRegister - (nport * 8))
        AxInstantDoCtrl3.WriteBit(nport, nbit, nStatus)
    End Sub






    Public Sub accuracyMA()

        Application.DoEvents()

        ACCTESTFLAG = 1
        Daily_Accuracy_Test = 1
        'MsgBox("Ma test")
        Me.Show()
        Me.Label36.Text = "DAILY ACCURACY TEST "
        Me.Label36.Font = New Drawing.Font("Times New Roman", 20, FontStyle.Bold)
        Me.Label36.BackColor = Color.Red
        Me.GroupBox2.Hide()
        Me.TextBox54.BringToFront()
        Me.TextBox54.Text = ""
        Me.TextBox54.Multiline = True
        Me.TextBox54.ScrollBars = ScrollBars.Vertical
        Me.Button13.Visible = True




        strFileName = IO.File.ReadAllLines("C:\Bench2\Console\mmc_err.INI") '// add each line as String Array.
        Application.DoEvents()

        initials = InputBox("Enter your initials:", "1 InputBox Initials", "Type your initials here.")
        If Not initials = "" Or initials = "Type your initials here." Then
            'Me.TextBox54.Text = "Wh Chk 2.5%: 120 /" & Environment.NewLine & "No Results are Available for Selected Test"
            Call Me.Button4_Click(0, System.EventArgs.Empty)
        Else
            initials = InputBox("Enter your initials:", "InputBox Initials", "Type your initials here.")

        End If
        Application.DoEvents()

        'Form1.BackgroundWorker2.CancelAsync()
        'While Form1.BackgroundWorker2.IsBusy
        '    Application.DoEvents()
        '    Threading.Thread.Sleep(50)
        'End While
        'Form1.BackgroundWorker1.RunWorkerAsync(5)



    End Sub
    Public Sub CompleteaccuracyMA()
        Dim Mtap As String
        Dim MPRN As String
        Dim e As Object

        ACCTESTFLAG = 1
        'MsgBox("Ma test")
        Me.Show()
        Me.Label36.Text = "DAILY ACCURACY TEST "
        Me.Label36.Font = New Drawing.Font("Times New Roman", 20, FontStyle.Bold)
        Me.Label36.BackColor = Color.Red
        Me.GroupBox2.Hide()
        Me.TextBox54.BringToFront()
        Me.TextBox54.Text = ""
        Me.TextBox54.Multiline = True
        Me.TextBox54.ScrollBars = ScrollBars.Vertical
        Me.Button13.Visible = True
        EMPSerialPort.Close()



        strFileName = IO.File.ReadAllLines("C:\Bench2\Console\mmc_err.INI") '// add each line as String Array.


        initials = InputBox("Enter your initials:", "2 InputBox Initials", "Type your initials here.")
        If Not initials = "" Or initials = "Type your initials here." Then
            'Me.TextBox54.Text = "Wh Chk 2.5%: 120 /" & Environment.NewLine & "No Results are Available for Selected Test"
            Call Button13_Click_1(0, System.EventArgs.Empty)

            ' read text box list into array 

            For Each strLine As String In Form8.TextBox1.Text.Split(vbNewLine)
                Dim ChkBox As CheckBox = Nothing
                ' to unchecked all 
                For Each xObject As Object In Me.GroupBox3.Controls
                    If TypeOf xObject Is CheckBox Then
                        ChkBox = xObject
                        ChkBox.Checked = False
                    End If
                Next

                Mtap = Mid(strLine, 3, 2)
                MPRN = Mid(strLine, 8, 3)

                Select Case Mtap

                    Case "02"
                        CheckBox2.Checked = True
                    Case "03"
                        CheckBox3.Checked = True
                    Case "04"
                        CheckBox4.Checked = True
                    Case "05"
                        CheckBox5.Checked = True
                    Case "06"
                        CheckBox6.Checked = True
                    Case "07"
                        CheckBox7.Checked = True
                    Case "08"
                        CheckBox8.Checked = True
                    Case "09"
                        CheckBox9.Checked = True
                    Case "10"
                        CheckBox10.Checked = True
                    Case "11"
                        CheckBox11.Checked = True
                    Case "12"
                        CheckBox12.Checked = True
                    Case "13"
                        CheckBox13.Checked = True
                    Case "14"
                        CheckBox14.Checked = True
                    Case "15"
                        CheckBox15.Checked = True
                    Case "16"
                        CheckBox16.Checked = True
                    Case "17"
                        CheckBox17.Checked = True
                    Case "18"
                        CheckBox18.Checked = True
                    Case "19"
                        CheckBox19.Checked = True
                    Case "20"
                        CheckBox20.Checked = True
                    Case "21"
                        CheckBox21.Checked = True
                    Case "22"
                        CheckBox22.Checked = True
                    Case "23"
                        CheckBox23.Checked = True
                    Case "24"
                        CheckBox24.Checked = True
                    Case "25"
                        CheckBox25.Checked = True
                    Case "26"
                        CheckBox26.Checked = True
                    Case "27"
                        CheckBox27.Checked = True
                    Case "28"
                        CheckBox28.Checked = True
                    Case "29"
                        CheckBox29.Checked = True
                    Case "30"
                        CheckBox30.Checked = True
                    Case "31"
                        CheckBox31.Checked = True
                    Case "32"
                        CheckBox32.Checked = True
                    Case "33"
                        CheckBox33.Checked = True
                    Case "34"
                        CheckBox34.Checked = True
                    Case "35"
                        CheckBox35.Checked = True
                    Case "36"
                        CheckBox36.Checked = True
                    Case "37"
                        CheckBox37.Checked = True
                    Case "38"
                        CheckBox38.Checked = True
                    Case "39"
                        CheckBox39.Checked = True
                    Case "40"
                        CheckBox40.Checked = True

                End Select



                If Not Trim(MPRN) = "" Then
                    TextBox8.Text = File.ReadAllText("C:\Bench2\Console\mA cert files\mA" & MPRN & "V100A.PRN")
                    TextBox2.Text = MPRN
                Else

                End If

                If Not TextBox8.Text = "" Then
                    Call SETRELAYS_Click(0, e)
                End If





            Next








        Else
            initials = InputBox("Enter your initials:", "InputBox Initials", "Type your initials here.")

        End If

    End Sub
    Private Sub Form5_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim today As Date = Date.Today
        Dim dayIndex As Integer = today.DayOfWeek
        Dim FILE_NAME As String = ""
        Dim Mondaystr As String
        Dim myconnection As New ADODB.Connection
        Dim FileSize As UInt32
        Dim rawData() As Byte
        Dim fs As FileStream









        If dayIndex < DayOfWeek.Monday Then
            dayIndex += 7 'Monday is first day of week, no day of week should have a smaller index
        End If

        Dim dayDiff As Integer = dayIndex - DayOfWeek.Monday
        Dim monday As Date = today.AddDays(-dayDiff)
        Mondaystr = monday.ToString("yyyyMMdd")




        Dim path As String = "C:\Bench2\Console\MMCMA" & Mondaystr & "\"
        If Not Directory.Exists(path) Then
            Directory.CreateDirectory(path)

        End If

        FILE_NAME = path & "mmcmalog" & DateTime.Now.ToString("yyyyMMdd") & ".txt"

        If System.IO.File.Exists(FILE_NAME) = True Then

            Dim objWriter As System.IO.StreamWriter = File.AppendText(FILE_NAME)
            objWriter.Write(TextBox54.Text)
            objWriter.Close()

        Else
            If System.IO.File.Exists(FILE_NAME) = False Then
                System.IO.File.Create(FILE_NAME).Dispose()
                Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
                objWriter.Write(TextBox54.Text)
                objWriter.Close()
                ' MsgBox("Text written to file")
            End If
        End If

        myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
        myconnection.Execute("Update [MCTEST].[dbo].TestTable  set [MCTEST].[dbo].TestTable.[FileName] =" & "'" & path & "'" & "where [MCTEST].[dbo].TestTable.[filename] is null")
        myconnection.Execute("Update [MCTEST].[dbo].TestTable  set [MCTEST].[dbo].TestTable.[Extension] =" & "'" & FILE_NAME & "'" & "where [MCTEST].[dbo].TestTable.[Extension] is null")

        fs = New FileStream(FILE_NAME, FileMode.Open, FileAccess.Read)
        FileSize = fs.Length
        rawData = New Byte(FileSize) {}
        fs.Read(rawData, 0, FileSize)
        fs.Close()


        'command.Execute("Update [MCTEST].[dbo].TestTable  set [MCTEST].[dbo].TestTable.[Content] = @rawData  where [MCTEST].[dbo].TestTable.[Content] is null")
        ' command.Parameters.Cast("@FileContents", SqlDbType.VarBinary).Value = IO.File.ReadAllBytes("file path here")
        'command.ExecuteNonQuery()
        Call File2SqlBlob(FILE_NAME)










        myconnection.Close()























        ACCTESTFLAG = 0
        Uniq_ID_Flag = 1




    End Sub

    Friend WithEvents Label50 As System.Windows.Forms.Label
    Friend WithEvents TextBox53 As System.Windows.Forms.TextBox
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents TextBox54 As System.Windows.Forms.TextBox
    Friend WithEvents AxInstantDoCtrl2 As AxBDaqOcxLib.AxInstantDoCtrl
    Friend WithEvents AxInstantDoCtrl1 As AxBDaqOcxLib.AxInstantDoCtrl
    Private WithEvents axFramer1 As AxDSOFramer.AxFramerControl
    Private WithEvents axFramer2 As AxDSOFramer.AxFramerControl
    Private WithEvents axFramer3 As AxDSOFramer.AxFramerControl
    Private WithEvents axFramer4 As AxDSOFramer.AxFramerControl
    Friend WithEvents AxInstantDoCtrl3 As AxBDaqOcxLib.AxInstantDoCtrl








    Function find(s As String) As Boolean
        Return s.Contains(xstr)
    End Function
    Friend WithEvents Button13 As System.Windows.Forms.Button

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click

        If Button13.BackColor = Color.LightSteelBlue Then

            GroupBox1.Visible = False
            GroupBox3.Visible = True
            Button13.BackColor = Color.Green
            Button13.Text = "Automated Test Positions"
        ElseIf Button13.BackColor = Color.Green Then
            GroupBox1.Visible = True
            GroupBox3.Visible = False
            Button13.BackColor = Color.LightSteelBlue
            Button13.Text = "Manual Test Positions"
        End If



    End Sub
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox40 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox39 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox38 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox37 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox36 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox35 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox34 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox33 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox32 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox31 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox30 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox29 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox28 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox27 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox26 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox25 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox24 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox23 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox22 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox21 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox20 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox19 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox18 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox17 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox16 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox15 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox14 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox13 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox12 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox11 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox10 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox9 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox8 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox7 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox

    Private Sub File2SqlBlob(ByVal SourceFilePath As String)
        Dim cn As New SqlConnection("server=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
        Dim cmd As New SqlCommand("Update [MCTEST].[dbo].TestTable  set [MCTEST].[dbo].TestTable.[Content] = @rawData  where [MCTEST].[dbo].TestTable.[Content] is null", cn)
        Dim fs As New System.IO.FileStream(SourceFilePath, IO.FileMode.Open, IO.FileAccess.Read)
        Dim b(fs.Length() - 1) As Byte
        fs.Read(b, 0, b.Length)
        fs.Close()
        Dim P As New SqlParameter("@rawData", SqlDbType.Image, b.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, b)
        cmd.Parameters.Add(P)
        cn.Open()
        cmd.ExecuteNonQuery()
        cn.Close()
    End Sub
    Friend WithEvents AxInstantDoCtrl4 As AxBDaqOcxLib.AxInstantDoCtrl

    Public Sub SetAfterPrn2()
        Dim WHinternal As Single
        Dim WHexternal As Single
        Dim IntCount(256) As Long
        Dim cbox1 As String
        Dim voltage As String
        Dim commport1 As Byte = 4
        Dim commport As Byte = 1   'MWH:Fix this to use the menus...
        Dim Model As String = New String(" ", RAD_SIZE_MODEL)
        Dim Serial As String = New String(" ", RAD_SIZE_SERIAL)
        Dim Version As String = New String(" ", RAD_SIZE_VERSION)
        Dim DeviceName As String = New String(" ", RAD_SIZE_NAME)
        On Error Resume Next

        'Call SetTxtBoxtoEmpty()
        TextBox49.Text = ""
        TextBox49.BackColor = Color.White
        TextBox50.Text = ""
        TextBox50.BackColor = Color.White
        TextBox51.Text = ""
        TextBox51.BackColor = Color.White
        TextBox10.BackColor = Color.White
        TextBox10.Text = ""
        TextBox11.BackColor = Color.White
        TextBox11.Text = ""
        TextBox12.BackColor = Color.White
        TextBox12.Text = ""
        TextBox13.BackColor = Color.White
        TextBox13.Text = ""
        TextBox14.BackColor = Color.White
        TextBox14.Text = ""
        TextBox15.BackColor = Color.White
        TextBox15.Text = ""
        TextBox16.BackColor = Color.White
        TextBox16.Text = ""
        TextBox17.BackColor = Color.White
        TextBox17.Text = ""
        TextBox18.BackColor = Color.White
        TextBox18.Text = ""
        TextBox19.BackColor = Color.White
        TextBox19.Text = ""
        TextBox20.BackColor = Color.White
        TextBox20.Text = ""
        TextBox21.BackColor = Color.White
        TextBox21.Text = ""
        TextBox22.BackColor = Color.White
        TextBox22.Text = ""
        TextBox23.BackColor = Color.White
        TextBox23.Text = ""
        TextBox24.BackColor = Color.White
        TextBox24.Text = ""
        TextBox25.BackColor = Color.White
        TextBox25.Text = ""
        TextBox26.BackColor = Color.White
        TextBox26.Text = ""
        TextBox27.BackColor = Color.White
        TextBox27.Text = ""
        TextBox28.BackColor = Color.White
        TextBox28.Text = ""
        TextBox29.BackColor = Color.White
        TextBox29.Text = ""
        TextBox30.BackColor = Color.White
        TextBox30.Text = ""
        TextBox31.BackColor = Color.White
        TextBox31.Text = ""
        TextBox32.BackColor = Color.White
        TextBox32.Text = ""
        TextBox33.BackColor = Color.White
        TextBox33.Text = ""
        TextBox34.BackColor = Color.White
        TextBox34.Text = ""
        TextBox35.BackColor = Color.White
        TextBox35.Text = ""
        TextBox36.BackColor = Color.White
        TextBox36.Text = ""
        TextBox37.BackColor = Color.White
        TextBox37.Text = ""
        TextBox38.BackColor = Color.White
        TextBox38.Text = ""
        TextBox39.BackColor = Color.White
        TextBox39.Text = ""
        TextBox40.BackColor = Color.White
        TextBox40.Text = ""
        TextBox41.BackColor = Color.White
        TextBox41.Text = ""
        TextBox42.BackColor = Color.White
        TextBox42.Text = ""
        TextBox43.BackColor = Color.White
        TextBox43.Text = ""
        TextBox44.BackColor = Color.White
        TextBox44.Text = ""
        TextBox45.BackColor = Color.White
        TextBox45.Text = ""
        TextBox46.BackColor = Color.White
        TextBox46.Text = ""
        TextBox47.BackColor = Color.White
        TextBox47.Text = ""
        TextBox48.BackColor = Color.White
        TextBox48.Text = ""

        mbSession = CType(ResourceManager.GetLocalManager().Open("GPIB0::0::INSTR"), MessageBasedSession)
        mbSession.Write("CHAN 1 MODE GATE")
        mbSession.Write("CHAN 1 GATE OFF")
        voltage = TextBox2.Text
        'Call SetChan2(voltage)

        If voltage = "120" Then
            '*120 to PT1
            'AxAdvDIO1.DeviceNumber = 0
            'AxAdvDIO1.WriteDoChannel(1, 22)
            'AxAdvDIO1.WriteDoChannel(0, 23)
            ' mbSession.Write("CHAN 2 Rang 10V AMR 3.4944") 'bench3
            WriteDoChannel0(1, 22)
            WriteDoChannel0(0, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.1952") ''bench2
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.1992") ''bench2 changed on November 11, 2022 by CH as per details from KW
            mbSession.Write("CHAN 2 Rang 10V AMR 5.2072") ''bench2 changed on May 2, 2023 by CH as per details from KW
        End If

        If voltage = "208" Then
            '*208 to PT1
            'AxAdvDIO1.DeviceNumber = 0
            'AxAdvDIO1.WriteDoChannel(0, 22)
            'AxAdvDIO1.WriteDoChannel(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 2.0189") ''Bench3
            WriteDoChannel0(0, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.0015") ''Bench2
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.0038") ''Bench2 changed on November 11, 2022 by CH as per details from KW
            mbSession.Write("CHAN 2 Rang 10V AMR 3.0085") ''Bench2 changed on May 2, 2023 by CH as per details from KW
        End If


        If voltage = "240" Then
            '*240 to PT1
            'AxAdvDIO1.DeviceNumber = 0
            'AxAdvDIO1.WriteDoChannel(0, 22)
            'AxAdvDIO1.WriteDoChannel(1, 23)
            ' mbSession.Write("CHAN 2 Rang 10V AMR 2.3296") 'bench3
            WriteDoChannel0(0, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.4635") 'bench2
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.4661") 'bench2 changed on Novemberr 11, 2022 by CH as per details from KW
            mbSession.Write("CHAN 2 Rang 10V AMR 3.4715") 'bench2 changed on May 2, 2023 by CH as per details from KW
        End If

        If voltage = "277" Then
            '*277 to PT1
            'AxAdvDIO1.DeviceNumber = 0
            'AxAdvDIO1.WriteDoChannel(0, 22)
            'AxAdvDIO1.WriteDoChannel(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 2.6887") ''bench3
            WriteDoChannel0(0, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.9974") ''bench2
            'mbSession.Write("CHAN 2 Rang 10V AMR 4.0005") ''bench2 changed on November 11, 2022 by CH as per details from KW
            mbSession.Write("CHAN 2 Rang 10V AMR 4.0066") ''bench2 changed on May 2, 2023 by CH as per details from KW
        End If


        If voltage = "347" Then
            '*347 to PT1
            'AxAdvDIO1.DeviceNumber = 0
            'AxAdvDIO1.WriteDoChannel(0, 22)
            'AxAdvDIO1.WriteDoChannel(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.3682") ''bench3
            WriteDoChannel0(0, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.0076") ''bench2
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.0115") ''bench2 changed on November 11, 2022 by CH as per details from KW
            mbSession.Write("CHAN 2 Rang 10V AMR 5.0192") ''bench2 changed on May 2, 2023 by CH as per details from KW
        End If
        If voltage = "416" Then
            '*416 to PT1
            ' AxAdvDIO1.DeviceNumber = 0
            ' AxAdvDIO1.WriteDoChannel(1, 22)
            ' AxAdvDIO1.WriteDoChannel(1, 23)
            ' mbSession.Write("CHAN 2 Rang 10V AMR 2.4228") ''bench3
            WriteDoChannel0(1, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.6020") ''Bench2
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.6048") ''Bench2 changed on November 11, 2022 by CH as per details from KW
            mbSession.Write("CHAN 2 Rang 10V AMR 3.6103") ''Bench2 changed on May 2, 2023 by CH as per details from KW
        End If
        If voltage = "480" Then
            '*480 to PT1
            ' AxAdvDIO1.DeviceNumber = 0
            'AxAdvDIO1.WriteDoChannel(1, 22)
            ' AxAdvDIO1.WriteDoChannel(1, 23)
            ' mbSession.Write("CHAN 2 Rang 10V AMR 2.7955") ''Bench3
            WriteDoChannel0(1, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 4.1562 ") ''Bench2
            'mbSession.Write("CHAN 2 Rang 10V AMR 4.1594 ") ''Bench2 changed on November 11, 2022 by CH as per details from KW
            mbSession.Write("CHAN 2 Rang 10V AMR 4.1658") ''Bench2 changed on May 2, 2023 by CH as per details from KW

        End If
        If voltage = "600" Then
            '*600 to PT1
            'AxAdvDIO1.DeviceNumber = 0
            ' AxAdvDIO1.WriteDoChannel(1, 22)
            ' AxAdvDIO1.WriteDoChannel(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 3.4944") ''bench3
            WriteDoChannel0(1, 22)
            WriteDoChannel0(1, 23)
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.1952 ") ''Bench2
            'mbSession.Write("CHAN 2 Rang 10V AMR 5.1992") ''Bench2 changed on November 11, 2022 by CH as per details from KW
            mbSession.Write("CHAN 2 Rang 10V AMR 5.2072") ''Bench2 changed on May 2, 2023 by CH as per details from KW
        End If
        Call teststep1()

        '88888888888888888888888888888888888888888888888888888888888888888888 After Step 1 8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            'RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            'RadRDAccumStop(comDevice)
            RadRDAccumReset(comDevice, 0)
        End If


        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================
        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(150)
        End While

        Application.DoEvents()

        Me.BackgroundWorker2.RunWorkerAsync()
        Threading.Thread.Sleep(150)

        Form1.Form1_CallScans(5)
        Thread.Sleep(500)


        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If

        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If


        RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999
        'TEST STOPPED'
        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If
        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            Threading.Thread.Sleep(800)
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        'Threading.Thread.Sleep(250)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(60)
        End While
        Form1.Form1_CallScans(5)
        'TESTING RELAY STOP AND ADDING DELAY - JS 11-15-2023
        Threading.Thread.Sleep(5000)
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")
        'MsgBox("set Yokogawa Voltage , Current, Phase   OFF")
        'MsgBox("Read MC MKA to excel")

        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If



        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.



        Dim LastRow As Long
        'With oSheet
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row


        '    '''''''''''''''''''' Top Row Only'''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    If .cells(1, 1).Value = "" Then


        '        .cells(1, 1).Value = "TAP"
        '        .cells(1, 2).Value = "Console Radian"
        '        .cells(1, 3).Value = "MC Radian"
        '        .cells(1, 4).Value = "Current Step"
        '        .cells(1, 5).Value = "V Mult"
        '        .cells(1, 6).Value = "Rcons"
        '        .cells(1, 7).Value = "C Mult"
        '        .cells(1, 8).Value = "Rmc"
        '        .cells(1, 9).Value = "% Err"

        '    End If

        '''''''''''''''''''''''''''''''''CT and Votage '''''''''''''''''''''''''''''''''''''''''''''''''''
        ' '''''''add CT here
        'LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '.cells(LastRow + 1, 1).value = "CT"
        '.cells(LastRow + 1, 2).Value = voltage



        ' '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        'LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row

        'If Len(CBNUM) = 13 Then
        '    CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        'Else
        '    CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        'End If

        '.cells(LastRow + 1, 1).Value = CTNUM
        '.cells(LastRow + 1, 2).Value = WHinternal
        '.cells(LastRow + 1, 3).Value = WHexternal
        '.cells(LastRow + 1, 4).Value = "WH @ 2.5% Unity"

        'voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '.cells(LastRow + 1, 5).Value = voltage

        '.cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
        '.cells(LastRow + 1, 7).Value = 1000
        '.cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
        '.cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
        'TextBox49.Text = .cells(LastRow + 1, 9).Text
        'If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '    TextBox49.BackColor = Color.Red
        'Else
        '    TextBox49.BackColor = Color.Lime
        'End If

        'oExcel.DisplayAlerts = False
        'oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        'Application.DoEvents()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With
        '88888888888888888888888888888888888888888888888888888888888888 step 1 error calc 88888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888

        Dim tdate As DateTime
        Dim outp As String
        Dim opertr As String
        Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
        Dim tap As String = ""
        Dim reading As String = Math.Round(CDbl(WHinternal))
        Dim modl As String = Model
        Dim seral As String = Serial
        Dim xpercent As Double = 0
        Dim percent As String = ""
        Dim xeconsole As Double = 0
        Dim econsole As String = ""
        Dim i As Integer = 0
        Dim numtap As String = ""
        Dim Vtap As String = ""
        Dim Emeas As String = ""
        Dim Estd As String = ""
        Dim _Error As String = ""

        Dim mulitplier As String = ""
        Dim xEtrue As Double = 0
        Dim Emeter As Double = 0
        Dim Etrue As Double = 0
        Dim Change As Double = 0
        Dim Total As Double = 0
        Dim Radian As Double


        Dim Error_Calc_Meter As Double = 0
        Dim Error_Calc_Corrected_Meter As Double = 0
        Dim Scaled_WattHourMeter As Double = 0
        Dim Scaled_WattHourConsole As Double = 0
        Dim Accuracy_Test_Results As Double = 0
        Dim rnd_Accuracy_Test_Results As Double = 0


        tdate = DateTime.Now
        outp = TextBox2.Text
        opertr = initials
        modl = modl.Replace(ControlChars.NullChar, "")
        seral = seral.Replace(ControlChars.NullChar, "")

        If Button13.BackColor = Color.Green Then
            numtap = Mid(CBNUM, 12, 2)
            numtap = Trim(numtap)
            numtap = CInt(numtap) - 1
        Else
            numtap = Mid(rButton.Name, 12, 2)
            numtap = Trim(numtap)
            numtap = CInt(numtap) - 1
        End If
        numtap = numtap.ToString
        If Len(numtap) = 1 Then
            numtap = "0" & numtap
        End If
        For Each x As String In strFileName
            If x.Equals("[WH100_" & outp & "]") Then
                Dim index1 As Integer = Array.IndexOf(strFileName, x)
                Dim WH100array(117) As String
                For ii As Integer = 0 To WH100array.Count - 1
                    Dim iii As Integer = ii + (index1 + 4)
                    WH100array(ii) = strFileName(iii)
                    Dim fstring As String = "1.25_1.0_M" & numtap
                    Dim sndhalf As String
                    Dim spresult() As String
                    If WH100array(ii).Contains(fstring) Then
                        sndhalf = WH100array(ii)
                        spresult = sndhalf.Split("=")
                        Dim results() As String
                        results = spresult(1).Split(",")
                        Vtap = Trim(results(0))
                        Emeas = results(1)
                        Estd = results(2)
                        _Error = results(3)
                    End If

                Next

            End If
        Next

        mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


        _Error = CDbl(_Error)
        _Error = Math.Round(CDbl(_Error), 3)
        econsole = _Error
        Emeter = CDbl(Estd)
        Radian = Math.Round(CDbl(WHinternal), 3)
        Dim Percent2 As String
        If (Daily_Accuracy_Test = 1) Then

            Dim tuple_CalcResultsDAC1 As Tuple(Of Double, Double, Double, Double, Double, Double)
            tuple_CalcResultsDAC1 = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResultsDAC1.Item1
            Scaled_WattHourMeter = tuple_CalcResultsDAC1.Item2
            Scaled_WattHourConsole = tuple_CalcResultsDAC1.Item3
            Error_Calc_Meter = tuple_CalcResultsDAC1.Item4
            Error_Calc_Corrected_Meter = tuple_CalcResultsDAC1.Item5
            Accuracy_Test_Results = tuple_CalcResultsDAC1.Item6

            Call WriteDAC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                     , Error_Calc_Meter, Accuracy_Test_Results, voltage, "WH @ 2.5% Unity", oSheet, 1)

            tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)


            Call DAC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
                                     , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
                                     , rnd_Accuracy_Test_Results.ToString, 1)

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            external = WHexternal.ToString
            Load = "1.25"
            powerfactor = "1.0"
            Unit = "WH"


            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If

            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If

            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()

        Else
            Dim tuple_CalcResultsMC1 As Tuple(Of Double, Double, Double, Double)
            'tuple_CalcResults = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            tuple_CalcResultsMC1 = AccuracyCheckCalcFunctions.MC_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResultsMC1.Item1
            Scaled_WattHourMeter = tuple_CalcResultsMC1.Item2
            Scaled_WattHourConsole = tuple_CalcResultsMC1.Item3
            Error_Calc_Meter = tuple_CalcResultsMC1.Item4
            'Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            'Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteMC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(Estd), CDbl(mulitplier) _
                                 , Error_Calc_Meter, voltage, "WH @ 2.5% Unity", oSheet, 1)
        End If


        ' **** end of Test 1 **********


        Call teststep2()

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumReset(comDevice, 0)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumReset(comDevice, 0)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(200)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        '    Status = RadRDAssignDevice(CShort(commport1), comDevice)

        '    If Status = 0 Then
        '        'Successfully connected
        '        'Get unit information and populate status bar
        '        Status = RadRDModel(comDevice, Model)
        '        Status = RadRDSerial(comDevice, Serial)
        '        Status = RadRDVersion(comDevice, Version)
        '        Status = RadRDName(comDevice, DeviceName)
        '    End If
        'RadRDAccumStart(comDevice)
        '    If comDevice <> 0 Then
        '        RadRDReleaseDevice(comDevice)
        '        comDevice = 0
        '    End If




        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999







        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If
        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(100)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While

        BackgroundWorker2.RunWorkerAsync()

        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        Thread.Sleep(100)
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")
        'MsgBox("set Yokogawa Voltage , Current, Phase   OFF")
        'MsgBox("Read MC MKA to excel")
        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.
100:

        'With oSheet
        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "WH @ 25% Unity"

        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage



        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
        '    .cells(LastRow + 1, 7).Value = 1000
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
        '    TextBox50.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox50.BackColor = Color.Red
        '    Else
        '        TextBox50.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        Application.DoEvents()
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With
        'If ACCTESTFLAG = 1 Then
        'Dim tdate As DateTime
        'Dim outp As String
        outp = ""
        'Dim opertr As String
        'Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
        tap = ""
        reading = Math.Round(CDbl(WHinternal))
        modl = Model
        seral = Serial
        xpercent = 0
        percent = ""
        xeconsole = 0
        econsole = ""
        i = 0
        numtap = ""
        Vtap = ""
        Emeas = ""
        Estd = ""
        _Error = ""

        mulitplier = ""
        xEtrue = 0
        Emeter = 0
        Etrue = 0
        Change = 0
        Total = 0
        'Dim Radian As Double


        Error_Calc_Meter = 0
        Error_Calc_Corrected_Meter = 0
        Scaled_WattHourMeter = 0
        Scaled_WattHourConsole = 0
        Accuracy_Test_Results = 0
        rnd_Accuracy_Test_Results = 0


        tdate = DateTime.Now
        outp = TextBox2.Text
        opertr = initials
        modl = modl.Replace(ControlChars.NullChar, "")
        seral = seral.Replace(ControlChars.NullChar, "")

        If Button13.BackColor = Color.Green Then
            numtap = Mid(CBNUM, 12, 2)
            numtap = Trim(numtap)
            numtap = CInt(numtap) - 1
        Else
            numtap = Mid(rButton.Name, 12, 2)
            numtap = Trim(numtap)
            numtap = CInt(numtap) - 1
        End If
        numtap = numtap.ToString
        If Len(numtap) = 1 Then
            numtap = "0" & numtap
        End If
        For Each x As String In strFileName
            If x.Equals("[WH100_" & outp & "]") Then
                Dim index1 As Integer = Array.IndexOf(strFileName, x)
                Dim WH100array(116) As String
                For ii As Integer = 0 To WH100array.Count - 1
                    Dim iii As Integer = ii + (index1 + 4)
                    WH100array(ii) = strFileName(iii)
                    'Dim fstring As String = "1.25_1.0_M" & numtap
                    Dim fstring As String = "12.5_1.0_M" & numtap
                    Dim sndhalf As String
                    Dim spresult() As String
                    If WH100array(ii).Contains(fstring) Then
                        sndhalf = WH100array(ii)
                        spresult = sndhalf.Split("=")
                        Dim results() As String
                        results = spresult(1).Split(",")
                        Vtap = Trim(results(0))
                        Emeas = results(1)
                        Estd = results(2)
                        _Error = results(3)
                    End If

                Next

            End If
        Next

        mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


        _Error = CDbl(_Error)
        _Error = Math.Round(CDbl(_Error), 3)
        econsole = _Error
        Emeter = CDbl(Estd)
        Radian = Math.Round(CDbl(WHinternal), 3)
        Percent2 = ""
        If (Daily_Accuracy_Test = 1) Then

            Dim tuple_CalcResultsDAC2 As Tuple(Of Double, Double, Double, Double, Double, Double)
            tuple_CalcResultsDAC2 = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResultsDAC2.Item1
            Scaled_WattHourMeter = tuple_CalcResultsDAC2.Item2
            Scaled_WattHourConsole = tuple_CalcResultsDAC2.Item3
            Error_Calc_Meter = tuple_CalcResultsDAC2.Item4
            Error_Calc_Corrected_Meter = tuple_CalcResultsDAC2.Item5
            Accuracy_Test_Results = tuple_CalcResultsDAC2.Item6

            Call WriteDAC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                , Error_Calc_Meter, Accuracy_Test_Results, voltage, "WH @ 25% Unity", oSheet, 2)

            tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)



            Call DAC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
                                     , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
                                     , rnd_Accuracy_Test_Results.ToString, 2)
            ' Write results to DB

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            external = WHexternal.ToString
            Load = "12.5"
            powerfactor = "1.0"
            Unit = "WH"


            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If

            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If

            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()


        Else
            Dim tuple_CalcResultsMC2 As Tuple(Of Double, Double, Double, Double)
            tuple_CalcResultsMC2 = AccuracyCheckCalcFunctions.MC_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResultsMC2.Item1
            Scaled_WattHourMeter = tuple_CalcResultsMC2.Item2
            Scaled_WattHourConsole = tuple_CalcResultsMC2.Item3
            Error_Calc_Meter = tuple_CalcResultsMC2.Item4
            'Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            'Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteMC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                     , Error_Calc_Meter, voltage, "WH @ 25% Unity", oSheet, 2)



        End If

        ' **** end of Step 2 *******
DQDQ:

        Call teststep3()

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumReset(comDevice, 0)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumReset(comDevice, 0)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(110)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If




        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999


        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If
        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(100)
        ComboBox1.Text = ""


        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")

        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_WH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.


        'With oSheet

        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "WH @ 25% @ PF"


        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage



        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
        '    .cells(LastRow + 1, 7).Value = 1000
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
        '    TextBox51.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox51.BackColor = Color.Red
        '    Else
        '        TextBox51.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        Application.DoEvents()
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With
        'If ACCTESTFLAG = 1 Then
        outp = ""
        'Dim opertr As String
        'Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
        tap = ""
        reading = Math.Round(CDbl(WHinternal))
        modl = Model
        seral = Serial
        xpercent = 0
        percent = ""
        xeconsole = 0
        econsole = ""
        i = 0
        numtap = ""
        Vtap = ""
        Emeas = ""
        Estd = ""
        _Error = ""

        mulitplier = ""
        xEtrue = 0
        Emeter = 0
        Etrue = 0
        Change = 0
        Total = 0
        'Dim Radian As Double


        Error_Calc_Meter = 0
        Error_Calc_Corrected_Meter = 0
        Scaled_WattHourMeter = 0
        Scaled_WattHourConsole = 0
        Accuracy_Test_Results = 0
        rnd_Accuracy_Test_Results = 0

        tdate = DateTime.Now
        outp = TextBox2.Text
        opertr = initials
        modl = modl.Replace(ControlChars.NullChar, "")
        seral = seral.Replace(ControlChars.NullChar, "")

        If Button13.BackColor = Color.Green Then
            numtap = Mid(CBNUM, 12, 2)
            numtap = Trim(numtap)
            numtap = CInt(numtap) - 1
        Else
            numtap = Mid(rButton.Name, 12, 2)
            numtap = Trim(numtap)
            numtap = CInt(numtap) - 1
        End If
        numtap = numtap.ToString
        If Len(numtap) = 1 Then
            numtap = "0" & numtap
        End If
        For Each x As String In strFileName
            If x.Equals("[WH100_" & outp & "]") Then
                Dim index1 As Integer = Array.IndexOf(strFileName, x)
                Dim WH100array(116) As String
                For ii As Integer = 0 To WH100array.Count - 1
                    Dim iii As Integer = ii + (index1 + 4)
                    WH100array(ii) = strFileName(iii)
                    Dim fstring As String = "12.5_0.5_M" & numtap
                    Dim sndhalf As String
                    Dim spresult() As String
                    If WH100array(ii).Contains(fstring) Then
                        sndhalf = WH100array(ii)
                        spresult = sndhalf.Split("=")
                        Dim results() As String
                        results = spresult(1).Split(",")
                        Vtap = Trim(results(0))
                        Emeas = results(1)
                        Estd = results(2)
                        _Error = results(3)
                    End If

                Next

            End If
        Next

        mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


        _Error = CDbl(_Error)
        _Error = Math.Round(CDbl(_Error), 3)
        econsole = _Error
        Emeter = CDbl(Estd)
        Radian = Math.Round(CDbl(WHinternal), 3)
        Percent2 = ""
        If (Daily_Accuracy_Test = 1) Then

            Dim tuple_CalcResultsDAC3 As Tuple(Of Double, Double, Double, Double, Double, Double)
            tuple_CalcResultsDAC3 = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResultsDAC3.Item1
            Scaled_WattHourMeter = tuple_CalcResultsDAC3.Item2
            Scaled_WattHourConsole = tuple_CalcResultsDAC3.Item3
            Error_Calc_Meter = tuple_CalcResultsDAC3.Item4
            Error_Calc_Corrected_Meter = tuple_CalcResultsDAC3.Item5
            Accuracy_Test_Results = tuple_CalcResultsDAC3.Item6




            Call WriteDAC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                , Error_Calc_Meter, Accuracy_Test_Results, voltage, "WH @ 25% @ PF", oSheet, 3)


            tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)


            Call DAC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
                                     , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
                                     , rnd_Accuracy_Test_Results.ToString, 3)
            ' Write results to DB

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            external = WHexternal.ToString
            Load = "12.5"
            powerfactor = "0.5"
            Unit = "WH"


            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If

            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If

            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()

        Else
            Dim tuple_CalcResultsMC3 As Tuple(Of Double, Double, Double, Double)
            'tuple_CalcResults = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            tuple_CalcResultsMC3 = AccuracyCheckCalcFunctions.MC_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            rnd_Accuracy_Test_Results = tuple_CalcResultsMC3.Item1
            Scaled_WattHourMeter = tuple_CalcResultsMC3.Item2
            Scaled_WattHourConsole = tuple_CalcResultsMC3.Item3
            Error_Calc_Meter = tuple_CalcResultsMC3.Item4
            'Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            'Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteMC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                     , Error_Calc_Meter, voltage, "WH @ 25% @ PF", oSheet, 3)

        End If

        ' **** end of Step 3 *******

        Call teststep4()

        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If


        '==================================================stop accumulation clear both radians==========================================================================

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        '==================================================stop accumulation clear both radians==========================================================================

        Threading.Thread.Sleep(200)

        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(200)
        End While
        BackgroundWorker2.RunWorkerAsync()
        Form1.Form1_CallScans(5)
        Thread.Sleep(500)

        If ComboBox1.Text = "" Then
            ComboBox1.Text = 15
            cbox1 = ComboBox1.Text
        Else
            cbox1 = ComboBox1.Text
        End If
        '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888temp start radian
        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
        End If
        'RadRDAccumStart(comDevice)
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If
        '999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999

        If repx > 0 Then
            ComboBox1.Text = 5
            cbox1 = ComboBox1.Text
        End If

        Application.DoEvents()
        For z = 1 To CInt(cbox1) - 1
            ComboBox1.Text = CInt(cbox1) - z
            Threading.Thread.Sleep(250)
            Application.DoEvents()
            Threading.Thread.Sleep(250)
            If stopflag = 1 Then
                Exit Sub
            End If
        Next z
        Application.DoEvents()
        ComboBox1.Text = 0
        Threading.Thread.Sleep(250)
        ComboBox1.Text = ""
        Form1.BackgroundWorker1.CancelAsync()
        While Form1.BackgroundWorker1.IsBusy
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While
        Form1.Form1_CallScans(5)
        mbSession.Write("CHAN 1 GATE OFF")
        WriteDoChannel1(0, 33)
        'MsgBox("Deactivate Valhalla shorting relay")
        mbSession.Write("CHAN 1 OUTP OFF ")
        mbSession.Write("CHAN 2 OUTP Off ")

        ''''''''''''''''''''''''''''''''''''''''''''''''internal Radian''''''''''''''''''''''''''''''''''''''''''''''

        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_VAH, WHinternal)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''external radian''''''''''''''''''''''''''''''''''''''
        If comDevice <> 0 Then
            RadRDReleaseDevice(comDevice)
            comDevice = 0
        End If

        Status = RadRDAssignDevice(CShort(commport1), comDevice)

        If Status = 0 Then
            'Successfully connected
            'Get unit information and populate status bar
            Status = RadRDModel(comDevice, Model)
            Status = RadRDSerial(comDevice, Serial)
            Status = RadRDVersion(comDevice, Version)
            Status = RadRDName(comDevice, DeviceName)
            RadRDAccumMetric(comDevice, 0, RAD_ACCUM_VAH, WHexternal)

            'MsgBox("Read External Radian to excel")

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Add data to cells of the first worksheet in the new workbook.

        'With oSheet
        '    '''''''''''' enter Raw Reading ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    LastRow = .Cells(.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row
        '    If Len(CBNUM) = 13 Then
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 2)
        '    Else
        '        CTNUM = Microsoft.VisualBasic.Right(CBNUM, 1)
        '    End If

        '    .cells(LastRow + 1, 1).Value = CTNUM
        '    .cells(LastRow + 1, 2).Value = WHinternal
        '    .cells(LastRow + 1, 3).Value = WHexternal
        '    .cells(LastRow + 1, 4).Value = "VAH @ 25% @ PF"

        '    voltage = ReadAllTextFromINI(TextBox2.Text.ToString().TrimEnd(" ")).ToString()
        '    .cells(LastRow + 1, 5).Value = voltage

        '    .cells(LastRow + 1, 6).Value = .cells(LastRow + 1, 2).Value * .cells(LastRow + 1, 5).Value
        '    .cells(LastRow + 1, 7).Value = 1000
        '    .cells(LastRow + 1, 8).Value = .cells(LastRow + 1, 7).Value * .cells(LastRow + 1, 3).Value
        '    .cells(LastRow + 1, 9).Value = (.cells(LastRow + 1, 8).Value - .cells(LastRow + 1, 6).Value) / .cells(LastRow + 1, 8).Value * 100
        '    TextBox52.Text = .cells(LastRow + 1, 9).Text
        '    If .cells(LastRow + 1, 9).Value > 0.2 Or .cells(LastRow + 1, 9).Value < -0.2 Then
        '        TextBox52.BackColor = Color.Red
        '    Else
        '        TextBox52.BackColor = Color.Lime
        '    End If

        '    oExcel.DisplayAlerts = False
        '    oBook.Worksheets(1).SaveAs("C:\MCTemp.xlsx")

        Application.DoEvents()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'End With


        'If ACCTESTFLAG = 1 Then
        outp = ""
        'Dim opertr As String
        'Dim rButton As RadioButton = GroupBox1.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()
        tap = ""
        reading = Math.Round(CDbl(WHinternal))
        modl = Model
        seral = Serial
        xpercent = 0
        percent = ""
        xeconsole = 0
        econsole = ""
        i = 0
        numtap = ""
        Vtap = ""
        Emeas = ""
        Estd = ""
        _Error = ""

        mulitplier = ""
        xEtrue = 0
        Emeter = 0
        Etrue = 0
        Change = 0
        Total = 0
        'Dim Radian As Double




        tdate = DateTime.Now
        outp = TextBox2.Text
        opertr = initials
        modl = modl.Replace(ControlChars.NullChar, "")
        seral = seral.Replace(ControlChars.NullChar, "")

        If Button13.BackColor = Color.Green Then
            numtap = Mid(CBNUM, 12, 2)
            numtap = Trim(numtap)
            numtap = CInt(numtap) - 1
        Else
            numtap = Mid(rButton.Name, 12, 2)
            numtap = Trim(numtap)
            numtap = CInt(numtap) - 1
        End If
        numtap = numtap.ToString
        If Len(numtap) = 1 Then
            numtap = "0" & numtap
        End If
        For Each x As String In strFileName
            If x.Equals("[VAH100_" & outp & "]") Then
                Dim index1 As Integer = Array.IndexOf(strFileName, x)
                Dim WH100array(40) As String
                For ii As Integer = 0 To WH100array.Count - 1
                    Dim iii As Integer = ii + (index1 + 2)
                    WH100array(ii) = strFileName(iii)
                    Dim fstring As String = "12.5_0.5_M" & numtap
                    Dim sndhalf As String
                    Dim spresult() As String
                    If WH100array(ii).Contains(fstring) Then
                        sndhalf = WH100array(ii)
                        spresult = sndhalf.Split("=")
                        Dim results() As String
                        results = spresult(1).Split(",")
                        Vtap = Trim(results(0))
                        Emeas = results(1)
                        Estd = results(2)
                        _Error = results(3)
                    End If

                Next

            End If
        Next

        mulitplier = ReadAllTextFromINI(outp.TrimEnd(" ")).ToString()


        _Error = CDbl(_Error)
        _Error = Math.Round(CDbl(_Error), 3)
        econsole = _Error
        Emeter = CDbl(Estd)
        Radian = Math.Round(CDbl(WHinternal), 3)
        Percent2 = ""
        If (Daily_Accuracy_Test = 1) Then

            Dim tuple_CalcResultsDAC4 As Tuple(Of Double, Double, Double, Double, Double, Double)
            tuple_CalcResultsDAC4 = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(_Error), CDbl(Estd))
            rnd_Accuracy_Test_Results = tuple_CalcResultsDAC4.Item1
            Scaled_WattHourMeter = tuple_CalcResultsDAC4.Item2
            Scaled_WattHourConsole = tuple_CalcResultsDAC4.Item3
            Error_Calc_Meter = tuple_CalcResultsDAC4.Item4
            Error_Calc_Corrected_Meter = tuple_CalcResultsDAC4.Item5
            Accuracy_Test_Results = tuple_CalcResultsDAC4.Item6

            Call WriteDAC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
              , Error_Calc_Meter, Accuracy_Test_Results, voltage, "VAH @ 25% @ PF", oSheet, 4)

            tap = getTapNumber(numtap, Button13.BackColor, rButton.Name)


            ' **** Update Accuracy Check Window *************
            Call DAC_UpdateTextbox54(tdate, opertr, outp, tap, modl, seral, WHinternal, WHexternal, mulitplier _
                                     , Scaled_WattHourMeter, Estd, _Error, Scaled_WattHourConsole, Error_Calc_Meter, Error_Calc_Corrected_Meter _
                                     , rnd_Accuracy_Test_Results.ToString, 4)
            ' Write results to DB

            Dim myconnection As New ADODB.Connection
            Dim mycommand As New ADODB.Command
            Dim ra As Integer
            Dim Load As String
            Dim powerfactor As String
            Dim connt As ADODB.Connection
            Dim connectionString As String
            Dim external As String
            Dim Recset As New ADODB.Recordset
            Dim Recset1 As New ADODB.Recordset
            Dim Recset2 As New ADODB.Recordset
            Dim Mdate As String
            Dim mdDate As DateTime
            Dim Unit As String = ""
            external = WHexternal.ToString
            Load = "12.5"
            powerfactor = "0.5"
            Unit = "Vah"


            If Not Button13.BackColor = Color.Green Then
                Vtap = CDbl(numtap) + 1
            Else
                Vtap = CDbl(numtap)

            End If

            Vtap = Vtap.ToString
            If Len(Vtap) < 2 Then
                Vtap = "M0" & Vtap
            Else
                Vtap = "M" & Vtap
            End If

            myconnection.Open("Provider=SQLOLEDB;Data Source=10.10.10.251\UTILITYMANAGER;Initial Catalog=MCTEST; User Id=bench3;Password=carma;")
            myconnection.Execute("insert into[MCTEST].[dbo].TestResults([Units],[Voltage],[Load],[Powerfactor],[Vtap],[percent_error],[WHexternal],[WHinternal],[operator],[Date]) values  ( " & _
                       "'" & Unit & "', " & _
                       "'" & outp & "', " & _
                       "'" & Load & "', " & _
                       "'" & powerfactor & "', " & _
                       "'" & Vtap & "', " & _
                       "'" & _Error & "'," & _
                       "'" & WHexternal & "'," & _
                       "'" & WHinternal & "'," & _
                       "'" & opertr & "'," & _
                       "'" & tdate & "'" & _
                       ")")

            Recset.Open(("select Max(date)AS Mdate from [MCTEST].[dbo].[TestResults]"), myconnection)

            If Not Uniq_ID_Flag > 0 Then
                If Not Recset.EOF Then
                    Mdate = Recset.GetString
                    mdDate = DateTime.Parse(Mdate)
                    myconnection.Execute("insert into[MCTEST].[dbo].TestTable([date]) values (convert(datetime," & _
                                                        "'" & mdDate & "'" & _
                              "))")

                    Uniq_ID_Flag = 1
                End If
                Recset1.Open(("select Max(id)AS UniqID_text from [MCTEST].[dbo].[TestTable]"), myconnection)

                If Not Recset1.EOF Then
                    UniqID_text = Recset1.GetString
                    UniqID_text = CInt(UniqID_text)
                    myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[Date] = convert(datetime," & "'" & mdDate & "'" & ")")


                End If

            Else

                myconnection.Execute("Update [MCTEST].[dbo].TestResults set [MCTEST].[dbo].TestResults.[UniqID_ text]= " & "'" & UniqID_text & "'" & "where [MCTEST].[dbo].[TestResults].[UniqID_ text] is null")

            End If

            myconnection.Close()
        Else
            Dim tuple_CalcResultsMC4 As Tuple(Of Double, Double, Double, Double)
            'tuple_CalcResults = AccuracyCheckCalcFunctions.Daily_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            tuple_CalcResultsMC4 = AccuracyCheckCalcFunctions.MC_Accuracy_Calc(CDbl(WHinternal), CDbl(WHexternal), CDbl(mulitplier), CDbl(Estd), CDbl(_Error))
            rnd_Accuracy_Test_Results = tuple_CalcResultsMC4.Item1
            Scaled_WattHourMeter = tuple_CalcResultsMC4.Item2
            Scaled_WattHourConsole = tuple_CalcResultsMC4.Item3
            Error_Calc_Meter = tuple_CalcResultsMC4.Item4
            'Error_Calc_Corrected_Meter = tuple_CalcResults.Item5
            'Accuracy_Test_Results = tuple_CalcResults.Item6

            Call WriteMC_toExcel(CInt(numtap), CDbl(WHexternal), CDbl(WHinternal), Scaled_WattHourConsole, Scaled_WattHourMeter, CDbl(_Error), CDbl(mulitplier) _
                     , Error_Calc_Meter, voltage, "WH @ 25% @ PF", oSheet, 3)

        End If


        If ACCTESTFLAG = 1 And RadioButton40.Checked = True And TextBox2.Text = "600" Then
            Call Button2_Click(0, System.EventArgs.Empty)
            'Exit Sub

        End If





        Application.DoEvents()
    End Sub


End Class
