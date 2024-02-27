Option Strict Off
Option Explicit On
Imports System.Runtime.InteropServices

Module RRKit_Declares_Constants

    ' Function declarations to calls to the RRKit DLL.

    Declare Function RadRDAssignDevice Lib "RadRRKit.dll" (ByVal comPort As Short, ByRef comID As Integer) As Integer
    Declare Function RadRDReleaseDevice Lib "RadRRKit.dll" (ByVal comID As Integer) As Integer
    Declare Function RadRDClosePort Lib "RadRRKit.dll" (ByVal comPort As Short) As Integer
    Declare Function RadRDPhaseCount Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef PhaseCount As Short) As Integer
    Declare Function RadRDModel Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Model As String) As Integer
    Declare Function RadRDSerial Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Serial As String) As Integer
    Declare Function RadRDName Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Name As String) As Integer
    Declare Function RadRDSetUnitName Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Name As String) As Integer
    Declare Function RadRDVersion Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Version As String) As Integer
    Declare Function RadRDStatus Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef Status As Integer) As Integer
    Declare Function RadRDMessage Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Message As String) As Integer
    Declare Function RadRDNOOP Lib "RadRRKit.dll" (ByVal comID As Integer) As Integer
    Declare Function RadRDInstMetric Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByVal Func As Short, ByRef value As Single) As Integer
    Declare Function RadRDInstMetricAll Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Func As Short, ByRef PhaseA As Single, ByRef PhaseB As Single, ByRef PhaseC As Single, ByRef phaseN As Single, ByRef phaseNet As Single) As Integer
    Declare Function RadRDInstMetricTable Lib "RadRRKit.dll" (ByVal comId As Integer, ByVal phase As Short, <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_R4)> ByRef metrics() As Single) As Integer
    Declare Function RadRDMaxMetric Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByVal Func As Short, ByRef value As Single) As Integer
    Declare Function RadRDMaxMetricAll Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Func As Short, ByRef PhaseA As Single, ByRef PhaseB As Single, ByRef PhaseC As Single, ByRef phaseN As Single, ByRef phaseNet As Single) As Integer
    Declare Function RadRDMaxMetricTable Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_R4)> ByRef MetricData() As Single) As Integer
    Declare Function RadRDMinMetric Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByVal Func As Short, ByRef value As Single) As Integer
    Declare Function RadRDMinMetricAll Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Func As Short, ByRef PhaseA As Single, ByRef PhaseB As Single, ByRef PhaseC As Single, ByRef phaseN As Single, ByRef phaseNet As Single) As Integer
    Declare Function RadRDMinMetricTable Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_R4)> ByRef MetricData() As Single) As Integer
    Declare Function RadRDAccumMetric Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByVal Func As Short, ByRef value As Single) As Integer
    Declare Function RadRDAccumMetricAll Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Func As Short, ByRef PhaseA As Single, ByRef PhaseB As Single, ByRef PhaseC As Single, ByRef phaseN As Single, ByRef phaseNet As Single) As Integer
    Declare Function RadRDAccumMetricTable Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_R4)> ByRef MetricData() As Single) As Integer
    Declare Function RadRDInstReset Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short) As Integer
    Declare Function RadRDAccumReset Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short) As Integer
    Declare Function RadRDAccumStart Lib "RadRRKit.dll" (ByVal comID As Integer) As Integer
    Declare Function RadRDAccumStop Lib "RadRRKit.dll" (ByVal comID As Integer) As Integer
    Declare Function RadRDAccumPulse Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef pulses As Short, ByVal mode As Short) As Integer
    Declare Function RadRDAccumTime Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef seconds As Single) As Integer
    Declare Function RadRDGetPulseOutput Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal port As Short, ByRef Func As Short, ByRef phase As Short) As Integer
    Declare Function RadRDSetPulseOutput Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal port As Short, ByVal Func As Short, ByVal phase As Short) As Integer
    Declare Function RadRD3PhaseSync Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef sync As Short) As Integer
    Declare Function RadRDGetPulseRate Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Func As Short, ByRef value As Single) As Integer
    Declare Function RadRDSetPulseRate Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Func As Short, ByVal value As Single) As Integer
    Declare Function RadRDInputControl Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef setup As Short) As Integer
    Declare Function RadRDBeepControl Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef beeb_onoff As Short) As Integer
    Declare Function RadRDLockTap Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByVal axis As Short, ByVal tap As Short) As Integer
    Declare Function RadRDUnlockTap Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByVal axis As Short) As Integer
    Declare Function RadRDGetTaps Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByRef vlock As Short, ByRef vtap As Short, ByRef ilock As Short, ByRef itap As Short) As Integer
    Declare Function RadRDResetDevice Lib "RadRRKit.dll" (ByVal comID As Integer) As Integer
    Declare Function RadRDPotentialGating Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal state As Short) As Integer
    Declare Function RadRDAutoCalStatus Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef acal_state As Short) As Integer
    Declare Function RadRDAutoCalSet Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal acal_state As Short) As Integer
    Declare Function RadRDModeSet Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal mode As Short) As Integer
    Declare Function RadRDModeStatus Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef mode As Short) As Integer
    Declare Function RadRDSystemStatus Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef Status As Integer) As Integer
    Declare Function RadRDIntegrationTime Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef value As Single) As Integer
    Declare Function RadRDGetTemperature Lib "RadRRKit.dll" (ByVal comID As Integer, ByRef temperature As Single) As Integer
    Declare Function RadRDHarmonicAnalysis Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByRef harmonic As HARMONIC_STRUC) As Integer
    Declare Function RadRDHarmonicData Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_R4)> ByRef HarmonicData() As Single) As Integer
    Declare Function RadRDMeterTest Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByRef meter_test As METER_TEST_STRUCT) As Integer
    Declare Function RadRDStandardTest Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByRef std_test As STD_TEST_STRUCT) As Integer
    Declare Function RadRDAnalogSenseTest Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal Func As Short, ByVal factor As Single, ByRef value As Single, ByRef ans As Single, ByRef diff As Single, ByRef Percent_Error As Single) As Integer
    Declare Function RadRS712SetOutput Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal value As Single) As Integer
    Declare Function RadRDWaveformCapture Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short) As Integer
    Declare Function RadRDWaveformData Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_R4)> ByRef WaveformData() As Single) As Integer
    Declare Function RadRDSendPacket Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal SendLength As Short, ByVal SendData As Byte(), ByVal RcvData As Byte(), ByRef ReturnLength As Short) As Integer
    Declare Function RadRDByteToFloat Lib "RadRRKit.dll" (ByVal Data As Byte(), ByRef value As Single) As Integer
    Declare Function RadRDFloatToByte Lib "RadRRKit.dll" (ByVal value As Single, ByVal Data As Byte()) As Integer
    Declare Function RadRDUserVoltageCalOffset Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByVal update As Short, ByRef value As Single) As Integer
    Declare Function RadRDUserCurrentCalOffset Lib "RadRRKit.dll" (ByVal comID As Integer, ByVal phase As Short, ByVal update As Short, ByRef value As Single) As Integer
    Declare Function RadRDRestoreFactoryDefaults Lib "RadRRKit.dll" (ByVal comID As Integer) As Integer

    ' Minimum size for data strings to use in functions that return data.
    Public Const RAD_SIZE_MODEL As Short = 10       'RadRDModel()
    Public Const RAD_SIZE_SERIAL As Short = 8       'RadRDSerial()
    Public Const RAD_SIZE_NAME As Short = 16        'RadRDName()
    Public Const RAD_SIZE_VERSION As Short = 8      'RadRDVersion()
    Public Const RAD_SIZE_MESSAGE As Short = 256    'RadRDMessage()
    Public Const RAD_INVALID_DATA As Double = 3.402823E+38 'Invalid Number (Used when reading metrics)

    ' RD-30 Phase Definitions
    Public Const RAD_PHASE_NONE As Short = 0 'Use this for single phase devices (RD-2x)
    Public Const RAD_PHASE_NET As Short = 0
    Public Const RAD_PHASE_A As Short = 1
    Public Const RAD_PHASE_B As Short = 2
    Public Const RAD_PHASE_C As Short = 3
    Public Const RAD_PHASE_N As Short = 4

    ' RD-2x/3x Auto-Calibration Definitions (RadRDAutocalSet and RadRDAutocalStatus)
    Public Const RAD_FULL_AUTOCAL As Short = 0
    Public Const RAD_PARTIAL_AUTOCAL As Short = 1

    ' RD-2x/3x Measurement Mode Definitions (RadRDModeSet)
    Public Const RAD_MODE_AC_RMS As Short = 0
    Public Const RAD_MODE_AC_AVG As Short = 1 'Only valid for RD-2X-4XX
    Public Const RAD_MODE_DC_RMS As Short = 2 'Only valid for RD-22-XXX
    Public Const RAD_MODE_DC_AVG As Short = 3 'Only valid for RD-22-4XX

    ' RD-2x/3x System Status Flag Definitions (RadRDSystemStatus)
    Public Const RAD_STATUS_AUTOCAL As Short = &H1S 'RD-2x Autocalibration Failure
    Public Const RAD_STATUS_AUTOCAL_A As Short = &H2S 'RD-3x Phase A Autocalibration Failure
    Public Const RAD_STATUS_AUTOCAL_B As Short = &H4S 'RD-3x Phase B Autocalibration Failure
    Public Const RAD_STATUS_AUTOCAL_C As Short = &H8S 'RD-3x Phase C Autocalibration Failure
    Public Const RAD_STATUS_MEMORY As Short = &H10S 'RD-2x Memory Error or RD-3x Master Memory Error
    Public Const RAD_STATUS_MEMORY_A As Short = &H20S 'RD-3x Phase A Memory Error
    Public Const RAD_STATUS_MEMORY_B As Short = &H40S 'RD-3x Phase B Memory Error
    Public Const RAD_STATUS_MEMORY_C As Short = &H80S 'RD-3x Phase C Memory Error
    Public Const RAD_STATUS_TDM_A As Short = &H200S 'RD-3x Phase A TDM Error
    Public Const RAD_STATUS_TDM_B As Short = &H400S 'RD-3x Phase B TDM Error
    Public Const RAD_STATUS_TDM_C As Short = &H800S 'RD-3x Phase C TDM Error
    Public Const RAD_STATUS_EBOARD As Short = &H1000S 'RD-2x E-Board Communication Error
    Public Const RAD_STATUS_EBOARD_A As Short = &H2000S 'RD-3x Phase A E-Board Communication Error
    Public Const RAD_STATUS_EBOARD_B As Short = &H4000S 'RD-3x Phase B E-Board Communication Error
    Public Const RAD_STATUS_EBOARD_C As Short = &H8000S 'RD-3x Phase C E-Board Communication Error
    Public Const RAD_STATUS_PULSEEXCEED As Integer = &H10000 'RD-2x Pulse Output Exceeded Max Limits
    Public Const RAD_STATUS_PULSEEXCEED_A As Integer = &H20000 'RD-3x Phase A Pulse Output Exceeded Max Limits
    Public Const RAD_STATUS_PULSEEXCEED_B As Integer = &H40000 'RD-3x Phase B Pulse Output Exceeded Max Limits
    Public Const RAD_STATUS_PULSEEXCEED_C As Integer = &H80000 'RD-3x Phase C Pulse Output Exceeded Max Limits

    ' Defines for the Instantaneous Metrics data (RadRDInstMetrics)
    Public Const RAD_INST_V As Short = 0 ' Volts
    Public Const RAD_INST_A As Short = 1 ' Amps
    Public Const RAD_INST_W As Short = 2 ' Watts
    Public Const RAD_INST_VA As Short = 3 ' VA
    Public Const RAD_INST_VAR As Short = 4 ' VAR
    Public Const RAD_INST_FREQ As Short = 5 ' Frequency
    Public Const RAD_INST_PHASE As Short = 6 ' Phase Angle (between Volts & Amps)
    Public Const RAD_INST_PF As Short = 7 ' Power Factor
    Public Const RAD_INST_ANS As Short = 8 ' Analog Sense
    Public Const RAD_INST_DPHASE As Short = 9 ' Delta Phase( Volts A & phase read )
    Public Const RAD_INST_V_DELTA As Short = 10 ' Volts Delta mode
    Public Const RAD_INST_W_DELTA As Short = 11 ' Watts Delta mode
    Public Const RAD_INST_VA_DELTA As Short = 12 ' VA Delta mode
    Public Const RAD_INST_VAR_DELTA As Short = 13 ' VAR Delta mode
    Public Const RAD_INST_VAR_WYE_XCONNECTED As Short = 14 ' VAR WYE Cross-connected
    Public Const RAD_INST_VAR_DELTA_XCONNECTED As Short = 15 ' VAR DELTA Cross-connected


    ' Defines for the Accumulating Metrics data (RadRDAccumMetrics and RadRDSetPulseRate )
    Public Const RAD_ACCUM_WH As Short = 0 ' Watt Hours
    Public Const RAD_ACCUM_VARH As Short = 1 ' VAR Hours
    Public Const RAD_ACCUM_QH As Short = 2 ' Q Hours
    Public Const RAD_ACCUM_VAH As Short = 3 ' VA Hours
    Public Const RAD_ACCUM_VH As Short = 4 ' Volt Hours
    Public Const RAD_ACCUM_AH As Short = 5 ' Amp Hours
    Public Const RAD_ACCUM_V2H As Short = 6 ' Volts Squared Hours
    Public Const RAD_ACCUM_A2H As Short = 7 ' Amps Squared Hours
    Public Const RAD_ACCUM_WHP As Short = 8 ' Positive Watt Hours
    Public Const RAD_ACCUM_WHN As Short = 9 ' Negative Watt Hours
    Public Const RAD_ACCUM_VARHP As Short = 10 ' Positive VAR Hours
    Public Const RAD_ACCUM_VARHN As Short = 11 ' Negative VAR Hours
    Public Const RAD_ACCUM_WH_DELTA As Short = 12 ' Watt hours Delta
    Public Const RAD_ACCUM_TIME As Short = 13 ' Accumulated Time
    Public Const RAD_ACCUM_VAH_DELTA As Short = 14 ' VA Hours Delta
    Public Const RAD_ACCUM_VARH_DELTA As Short = 15 ' VAR Hours Delta
    Public Const RAD_ACCUM_VARH_DELTA_XCONNECTED As Short = 16 ' VAR Hours Delta X
    Public Const RAD_ACCUM_VARH_WYE_XCONNECTED As Short = 17 ' VAR Hours WYE X
    Public Const RAD_ACCUM_WH_DELTA_POSITIVE As Short = 18 ' Watt Hours Delta +
    Public Const RAD_ACCUM_WH_DELTA_NEGATIVE As Short = 19 ' Watt Hours Delta -
    Public Const RAD_ACCUM_VARH_DELTA_POSITIVE As Short = 20 ' VAR Hours Delta +
    Public Const RAD_ACCUM_VARH_DELTA_NEGATIVE As Short = 21 ' VAR Hours Delta -
    Public Const RAD_ACCUM_VARH_DELTA_XCONNECTED_POSITIVE As Short = 22 ' VAR Hours Delta X +
    Public Const RAD_ACCUM_VARH_DELTA_XCONNECTED_NEGATIVE As Short = 23 ' VAR Hours Delta X -
    Public Const RAD_ACCUM_VARH_WYE_XCONNECTED_POSITIVE As Short = 24 ' Var Hours WYE X +
    Public Const RAD_ACCUM_VARH_WYE_XCONNECTED_NEGATIVE As Short = 25 ' Var Hours WYE X -


    ' Defines for the RadRDAccumPulse modes
    Public Const PULSE_MANUAL_MODE As Short = 0
    Public Const PULSE_SENSOR_MODE As Short = 1

    ' Valid values for the Pulse Output setting (RadRDGetPulseOutput and
    ' for RadRDSetPulseOutput)
    Public Const RAD_PO_RD3x_PORT1 As Short = 1 'Pulse Output Port 1 on RD-3x
    Public Const RAD_PO_RD3x_PORT2 As Short = 2 'Pulse Output Port 2 on RD-3x
    Public Const RAD_PO_RD3x_PORT3 As Short = 3 'Pulse Output Port 3 on RD-3x
    Public Const RAD_PO_RD2x As Short = 0 'Pulse Output on RD-2x (BNC 2)
    Public Const RAD_PO_WH As Short = 0 ' Watt Hours
    Public Const RAD_PO_VARH As Short = 1 ' VAR Hours
    Public Const RAD_PO_QH As Short = 2 ' Q Hours
    Public Const RAD_PO_VAH As Short = 3 ' VA Hours
    Public Const RAD_PO_VH As Short = 4 ' Volt Hours
    Public Const RAD_PO_AH As Short = 5 ' Amp Hours
    Public Const RAD_PO_V2H As Short = 6 ' Volts Squared Hours
    Public Const RAD_PO_A2H As Short = 7 ' Amps Squared Hours
    Public Const RAD_PO_WHP As Short = 8 ' Positive Watt Hours
    Public Const RAD_PO_WHN As Short = 9 ' Negative Watt Hours
    Public Const RAD_PO_VARHP As Short = 10 ' Positive VAR Hours
    Public Const RAD_PO_VARHN As Short = 11 ' Negative VAR Hours
    Public Const RAD_PO_WH_DELTA As Short = 12 ' Watt hours Delta
    Public Const RAD_PO_VAH_DELTA As Short = 14 ' VA Hours Delta
    Public Const RAD_PO_VARH_DELTA As Short = 15 ' VAR Hours Delta
    Public Const RAD_PO_VARH_DELTA_XCONNECTED As Short = 16 ' VAR Hours Delta X
    Public Const RAD_PO_VARH_WYE_XCONNECTED As Short = 17 ' VAR Hours WYE X
    Public Const RAD_PO_WH_DELTA_POSITIVE As Short = 18 ' Watt Hours Delta +
    Public Const RAD_PO_WH_DELTA_NEGATIVE As Short = 19 ' Watt Hours Delta -
    Public Const RAD_PO_VARH_DELTA_POSITIVE As Short = 20 ' VAR Hours Delta +
    Public Const RAD_PO_VARH_DELTA_NEGATIVE As Short = 21 ' VAR Hours Delta -
    Public Const RAD_PO_VARH_DELTA_XCONNECTED_POSITIVE As Short = 22 ' VAR Hours Delta X +
    Public Const RAD_PO_VARH_DELTA_XCONNECTED_NEGATIVE As Short = 23 ' VAR Hours Delta X -
    Public Const RAD_PO_VARH_WYE_XCONNECTED_POSITIVE As Short = 24 ' Var Hours WYE X +
    Public Const RAD_PO_VARH_WYE_XCONNECTED_NEGATIVE As Short = 25 ' Var Hours WYE X -

    Public Const RAD_PULSEOUT_POSITIVE As Short = 0
    Public Const RAD_PULSEOUT_NEGATIVE As Short = 1

    'Defines for setting BNC input mode (RadRDInputControl)
    Public Const RAD_INPUT_SSC_MANUAL As Short = 0 ' Start/Stop/Clear Gating (Manual Type)
    Public Const RAD_INPUT_CSS_MANUAL As Short = 1 ' Clear+Start/Stop Gating (Manual Type)
    Public Const RAD_INPUT_NEGATIVE As Short = 2 ' Negative Pulses Out     (RD-2x Only)
    Public Const RAD_INPUT_SSC_SENSOR As Short = 8 ' Start/Stop/Clear Gating (Sensor Type)
    Public Const RAD_INPUT_CSS_SENSOR As Short = 9 ' Clear+Start/Stop Gating (Sensor Type)
    Public Const RAD_INPUT_STATUS As Short = &HFFS ' Get status of pulse input mode

    ' Valid values for the Voltage and Current tap settings (from RadRDGetTaps and
    ' for RadRDSetTaps).  The Voltage axis uses taps 1 through 3.  The Current axis
    ' uses taps 1 through 13.
    Public Const RAD_TAP_UNLOCK As Short = 0
    Public Const RAD_TAP_LOCK As Short = 1
    Public Const RAD_TAP_UNLOCKED As Short = 0
    Public Const RAD_TAP_LOCKED As Short = 1
    Public Const RAD_TAP_OPEN As Short = 0
    Public Const RAD_TAP_1 As Short = 1
    Public Const RAD_TAP_2 As Short = 2
    Public Const RAD_TAP_3 As Short = 3
    Public Const RAD_TAP_4 As Short = 4
    Public Const RAD_TAP_5 As Short = 5
    Public Const RAD_TAP_6 As Short = 6
    Public Const RAD_TAP_7 As Short = 7
    Public Const RAD_TAP_8 As Short = 8
    Public Const RAD_TAP_9 As Short = 9
    Public Const RAD_TAP_10 As Short = 10
    Public Const RAD_TAP_11 As Short = 11
    Public Const RAD_TAP_12 As Short = 12
    Public Const RAD_TAP_13 As Short = 13

    'Harmonic Analysis Trigger Control (RadRDHarmonicTrigger)
    Public Const RAD_HARMONIC_STATUS As Short = 0 'Return status of harmonic analysis
    Public Const RAD_HARMONIC_RESTART As Short = 1  'Return status of harmonic analysis
    'Restart analysis if completed
    Public Const RAD_HARMONIC_VOLTAGE As Short = 2 'Trigger to start voltage harmonic analysis
    'on the requested harmonic number.
    'If the harmonic number is 0 then stop analysis.
    Public Const RAD_HARMONIC_CURRENT As Short = 3 'Trigger to start current harmonic analysis
    'on the requested harmonic number.
    'If the harmonic number is 0 then stop analysis.

    'Return status (RadRDHarmonicTrigger)
    Public Const RAD_HARMONIC_NOTRUNNING As Short = 0 'No harmonic analysis is being preformed
    Public Const RAD_HARMONIC_INPROGRESS As Short = 1 'Harmonic Analysis in progress
    Public Const RAD_HARMONIC_VREADY As Short = 2 'Voltage harmonic analysis ready
    Public Const RAD_HARMONIC_IREADY As Short = 3 'Current(I) harmonic analysis ready

    'HarmonicTrigger Structure
    Public Structure HARMONIC_STRUC
        Dim harmonic As Byte 'Harmonic to read
        Dim Status As Byte 'Control/Status register
        Dim zoom As Byte 'Subharmonic zoom factor
        Dim base As Byte 'Base Harmonic
        Dim magnitude As Single 'Magnitude measured
        Dim phase As Single 'Phase measured
        Dim distortion As Single 'Distortion measured
    End Structure

    'Meter Test Control Defines (RadRDMeterTest)
    Public Const RAD_METERTEST_STATUS As Short = 0 'Request status of running Meter Test
    Public Const RAD_METERTEST_STARTTEST As Short = 1 'Start Meter Test
    Public Const RAD_METERTEST_PULSE As Short = 0 'Pulse Based Sensor Test
    Public Const RAD_METERTEST_TIMED As Short = 2 'Timed Based Sensor Test
    Public Const RAD_METERTEST_MANUAL As Short = 4 'Pulse Based Manual Test
    Public Const RAD_METERTEST_DEMAND As Short = 6 'Demand Meter Test
    Public Const RAD_METERTEST_TOTALIZED As Short = 8 'Totalized test on RD-3x devices

    'Meter Test Element Defines (RadRDMeterTest)
    Public Const RAD_METERTEST_3 As Short = 0 '3 Element meter
    Public Const RAD_METERTEST_2_5 As Short = 1 '2.5 Element meter
    Public Const RAD_METERTEST_2 As Short = 2 '2 Element meter
    Public Const RAD_METERTEST_1 As Short = 3 '1 Element meter

    'Meter Test Structure
    Public Structure METER_TEST_STRUCT
        Dim func As Byte 'Use RAD_ACCUM_XXXX defines for function desired
        Dim Control As Byte 'Control and Status. Use RAD_METERTEST_XXX definitions
        Dim Duration As Single 'duration of test in seconds if Timed and Demand Test
        Dim factor As Single 'kH factor for pulse and times tests or
        'demand register for demand testing
        Dim elements As Byte 'Number of meter elements.
        'Use RAD_METERTEST_XXX definitions
        Dim c_elements As Byte 'Number of Current elements. 1,2 and 3 are supported.
        Dim Revs As Short 'The number of revolutions for the Meter during the test.
        'Only Sensor and Manual Tests.
        Dim pulses_rev As Short 'The number of pulses per revolution for the Meter
        'during the test.  Only Sensor and Manual Tests.
        Dim Percent_Error As Single 'Percent error when test is complete.
        Dim registration As Single 'Percent Registration when test is complete.
        Dim pulses As Integer 'Number of pulses counted by the RD
        Dim metric As Single 'Accumulated metric reading from the device
    End Structure

    'Standard Test Control Defines (RadRDStandardTest)
    Public Const RAD_STDTEST_STATUS As Short = 0 'Request status of running Meter Test
    Public Const RAD_STDTEST_STARTTEST As Short = 1 'Start a Standard Test
    Public Const RAD_STDTEST_NONRADIAN As Short = 0 'Non-Radian Standard
    Public Const RAD_STDTEST_RADIAN As Short = 2 'Radian Standard
    Public Const RAD_STDTEST_TIMED As Short = 0 'Time based standard test
    Public Const RAD_STDTEST_PULSE As Short = 4 'Pulse based standard test
    Public Const RAD_STDTEST_TOTALIZED As Short = 8 'Totalized test on RD-3x devices

    'Standard Test Structure
    Public Structure STD_TEST_STRUCT
        Dim func As Byte 'Use RAD_ACCUM_XXXX defines for function desired
        Dim Control As Byte 'Control and Status. Use RAD_STDTEST_XXX definitions
        Dim Duration As Single 'duration of test in seconds/pulses
        Dim constant As Single 'Pulse constant
        Dim Percent_Error As Single 'Percent error when test is complete.
        Dim registration As Single 'Percent Registration when test is complete.
        Dim pulses As Integer 'Number of pulses counted by the RD
        Dim metric As Single 'Accumulated metric reading from the device
    End Structure
End Module