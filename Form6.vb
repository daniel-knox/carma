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

Public Class Form6
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


    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call FTestApp.Button13_Click(0, e)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
