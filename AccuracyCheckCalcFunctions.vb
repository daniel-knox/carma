Public Class AccuracyCheckCalcFunctions

    Public Shared Function Daily_Accuracy_Calc(ByVal WattHourConsole As Double, ByVal WattHourMeter As Double, ByVal Multiplier As Double _
                                    , ByVal CTError As Double, ByVal ConsoleError As Double) As Tuple(Of Double, Double, Double, Double, Double, Double)
        Dim DAC_percentError As Double
        'Dim Radian As Double
        Dim Scaled_WattHourMeter As Double
        Dim Scaled_WattHourConsole As Double
        Dim Error_meter As Double
        Dim Corrected_error_meter As Double
        Dim Accuracy_test_result As Double
        ' This function is the Accuracy check function that gets run daily


        'Radian = Math.Round(CDbl(WattHourInternal), 3)

        '*** Checked Device Error determined by Console  = (Check Device Reading X 1000 (convert from milliamp to amps) – Console Reading X the Voltage Multiplier(based on the voltage from the PRN) / Console Reading X the Voltage Multiplier(based on the voltage from the PRN)) X 100 
        Scaled_WattHourMeter = WattHourMeter * 1000 ' Scaling for mA
        Scaled_WattHourConsole = WattHourConsole * Multiplier ' Multiplier from MA_Settings.ini

        Error_meter = ((Scaled_WattHourMeter - Scaled_WattHourConsole) / Scaled_WattHourConsole) * 100

        '*** Corrected Error of the Checked device =  Checked Device Error determined by Console - %Econ(from MC Spreadsheet)
        Corrected_error_meter = Error_meter - ConsoleError

        '*** Accuracy Test Result = Corrected Error of the Checked device - P-E-01 6.1.1.2(1C) reference meter option results (to be determined by Steve)

        Accuracy_test_result = Corrected_error_meter - CTError

        DAC_percentError = Math.Abs(Math.Round(Accuracy_test_result, 3))




        Return New Tuple(Of Double, Double, Double, Double, Double, Double)(DAC_percentError, Scaled_WattHourMeter, Scaled_WattHourConsole, Error_meter, Corrected_error_meter, Accuracy_test_result)
    End Function


    Public Shared Function MC_Accuracy_Calc(ByVal WattHourConsole As Double, ByVal WattHourStandard As Double, ByVal Multiplier As Double _
                                , ByVal TapError As Double, ByVal ConsoleError As Double) As Tuple(Of Double, Double, Double, Double)
        Dim MCAC_percError As Double
        ' This function is the Accuracy Check that is performed once a year for the Measurement Canada Accuracy Check
        Dim Scaled_WattHourStandard As Double
        Dim Scaled_WattHourConsole As Double
        Dim Error_meter As Double
        'Dim Corrected_error_meter As Double
        'Dim Accuracy_test_result As Double
        ' This function is the Accuracy check function that gets run daily


        'Radian = Math.Round(CDbl(WattHourInternal), 3)

        '*** Checked Device Error determined by Console  = (Check Device Reading X 1000 (convert from milliamp to amps) – Console Reading X the Voltage Multiplier(based on the voltage from the PRN) / Console Reading X the Voltage Multiplier(based on the voltage from the PRN)) X 100 
        Scaled_WattHourStandard = WattHourStandard * 1000 ' Scaling for mA
        Scaled_WattHourConsole = WattHourConsole * Multiplier ' Multiplier from MA_Settings.ini

        Error_meter = ((Scaled_WattHourConsole - Scaled_WattHourStandard) / Scaled_WattHourStandard) * 100

        '*** Corrected Error of the Checked device =  Checked Device Error determined by Console - %Econ(from MC Spreadsheet)
        'Corrected_error_meter = Error_meter - ConsoleError

        '*** Accuracy Test Result = Corrected Error of the Checked device - P-E-01 6.1.1.2(1C) reference meter option results (to be determined by Steve)
        ' Not needed for the MC accuracy check as this is to find the errors in the Taps
        'Accuracy_test_result = Corrected_error_meter - TapError


        MCAC_percError = Math.Abs(Math.Round(Error_meter, 3)).ToString

        Return New Tuple(Of Double, Double, Double, Double)(MCAC_percError, Scaled_WattHourStandard, Scaled_WattHourConsole, Error_meter)
    End Function

End Class
