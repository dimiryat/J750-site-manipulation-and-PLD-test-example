Attribute VB_Name = "VBT_Module"
Public AFE_noisefloor_test_FailCount As New PinListData

Option Explicit

' This module should be used for VBT Tests.  All functions in this module
' will be available to be used from the Test Instance sheet.
' Additional modules may be added as needed (all starting with "VBT_").
'
' The required signature for a VBT Test is:
'
' Public Function FuncName(<arglist>) As Long
'   where <arglist> is any list of arguments supported by VBT Tests.
'
' See online help for supported argument types in VBT Tests.
'
'
' It is highly suggested to use error handlers in VBT Tests.  A sample
' VBT Test with a suggeseted error handler is shown below:
'
' Function FuncName() As Long
'     On Error GoTo errHandler
'
'     Exit Function
' errHandler:
'     If AbortTest Then Exit Function Else Resume Next
' End Function


Public Function VBT_AFE_Noise_Floor(patternName As String, mosi_Pin As String, miso_Pin As String, start_label As String, _
                            vec_offset As Long, start_Addr As Long, num_of_AFEs As Long, bytes_per_AFE As Long, num_of_Samples As Long, _
                            adc_min_ll As Double, adc_max_ul As Double, lower_limit As Double, upper_limit As Double, _
                            Debug_mode As Boolean, comments As String) As Long


                'New VBT Function created for F/W based AFE Noise Floor test  --- Wayne 02/10/2017
                'Data Format for each channel:
                '       - 0 - AFE# max val
                '       - 1 - AFE# min val
                '       - 2 - AFE# x0 val
                '       - 3,4 - AFE# values sum   (2s compliment)
                '       - 5,6 - AFE# values sum^2 (unsigned)
                '
                '       1. Loop through AFE's
                '       2. Loop through sites
                '       3. Process the data for one AFE channel
                '           (bytes_per_AFE = 7 bytes for Whitney)
                '

    'Set error handling
    'On Error GoTo ErrorHandler
                        
    Dim i As Long
    Dim j As Integer
    Dim k As Long
    Dim m As Integer
    Dim N As Integer

    Dim sitestatus As Long
    Dim thisSite As Long
    Dim loop_count As Integer
    Dim remaining_bytes As Integer
    Dim bytes_of_data As Integer
    Dim byte_counter As Integer

    Dim func_pass_flag As Boolean
    Dim Log_a_line As String

    Dim debug_str As String
    Dim this_instance As String
    Dim tName As String
    Dim chName As String
    Dim RetTestNums() As Long
    Dim RetTestCnt As Long
    Dim this_testnum As Long

    Dim a_16b_data As Long
    Dim one_afe_data(7) As Long
    Dim one_byte As Long
    Dim half_byte As Integer
    Dim str_byte As String
    Dim str_lsb As String
    Dim str_msb As String
    
    Dim min_value As Long
    Dim max_value As Long
    Dim max_noise As Double
    
    Dim adc_sum As Double
    Dim adc_sum2 As Double
    Dim noise_sigma As Double
    
    Dim adcMin_PLD As New PinListData
    Dim adcMax_PLD As New PinListData
    Dim nf_PLD As New PinListData
 
    this_instance = TheExec.DataManager.instanceName
    Call TheExec.DataManager.GetTestNumbers(this_instance, RetTestNums(), RetTestCnt)
    this_testnum = RetTestNums(0)

    If TheExec.Sites.ActiveCount = 0 Then GoTo NoSitesActive
    
    thehdw.digital.Patterns.pat(patternName).Load

    'Connect Pins
    thehdw.PinLevels.ConnectAllPins
    
    'Apply Levels/Timing
    thehdw.PinLevels.ApplyPower
    thehdw.digital.Timing.Load
    thehdw.Wait 0.001
    
    adcMin_PLD.AddPin ("AFE_x")
    adcMax_PLD.AddPin ("AFE_x")
    nf_PLD.AddPin ("AFE_x")
    adcMin_PLD.ResultType = tlResultTypeParametricValue
    adcMax_PLD.ResultType = tlResultTypeParametricValue
    nf_PLD.ResultType = tlResultTypeParametricValue
    
    AFE_noisefloor_test_FailCount.AddPin ("AFE_x")
    AFE_noisefloor_test_FailCount.ResultType = tlResultTypeParametricValue
    
    TheExec.datalog.WriteComment ""
    
    'Calculate loop parameters
    bytes_of_data = bytes_per_AFE
    
    For loop_count = 0 To num_of_AFEs - 1
        'Set up HRAM
        SetupHRAM captSTV, trigSTV, False, True, 0

        If loop_count = 0 Then
            Call update_pattern_fw_wr_16bit_data(patternName, mosi_Pin, start_label, vec_offset, start_Addr, 0, True)
        ElseIf loop_count = 1 Then
            Call update_pattern_fw_wr_16bit_data(patternName, mosi_Pin, start_label, vec_offset, 65535, 0, True) 'Put back addr = FFFF
        End If 'loop 1

        Call thehdw.digital.Patterns.pat(patternName).Start("")
        thehdw.digital.Patgen.HaltWait
        'siteIndex = TheExec.Sites.SelectFirst

        'For siteIndex = 0 To (TheExec.Sites.ExistingCount - 1)
        '    passed_func(siteIndex) = thehdw.digital.Patgen.PatternBurstPassed(siteIndex)
        'Next siteIndex
        
        'loop through active sites
        sitestatus = TheExec.Sites.SelectFirst

        Do While sitestatus <> loopDone
            thisSite = TheExec.Sites.SelectedSite
            
            func_pass_flag = thehdw.digital.Patgen.PatternBurstPassed(thisSite)
            
            'Log_a_line = "FW_DATA_CAPT " + CStr(this_testnum + loop_count) + " " + CStr(thisSite) + " "
            Log_a_line = CStr(this_testnum + 100 + loop_count) + " " + CStr(thisSite) + " "
            debug_str = "FW_DATA_DEBUG_MODE " + CStr(this_testnum + 200) + " " + CStr(thisSite) + " "
            
            
            If func_pass_flag = True Then
                Log_a_line = Log_a_line + "PASS "
                debug_str = debug_str + "PASS "
            Else
                Log_a_line = Log_a_line + "FAIL "
                debug_str = debug_str + "FAIL "
            End If
            
            If loop_count < 40 Then
                chName = "AFE" + CStr(loop_count)
            ElseIf loop_count < 42 Then
                chName = "REFL" + CStr(loop_count - 40)
            ElseIf loop_count < 82 Then
                chName = "AFE" + CStr(loop_count - 2)
            ElseIf loop_count < 84 Then
                chName = "REFR" + CStr(loop_count - 82)
            End If
            If num_of_AFEs = 2 Then
                'SB AFE's
                chName = "SB_RX" + CStr(loop_count)
            End If
            Log_a_line = Log_a_line + chName + " "
            
            one_byte = 0
            half_byte = 0
            byte_counter = 0
            str_byte = ""
            For i = 0 To 8 * bytes_of_data - 1
                If (TheExec.TesterMode = testModeOnline) Then
                    If (thehdw.digital.HRAM.pins(miso_Pin).PinData(i) = "L") Then
                        'debug_str = debug_str + "0"
                    ElseIf (thehdw.digital.HRAM.pins(miso_Pin).PinData(i) = "H") Then
                        'debug_str = debug_str + "1"
                        one_byte = one_byte + 1
                        half_byte = half_byte + 1
                    Else
                        TheExec.datalog.WriteComment "TEST_ERROR: Unexpected data while capturing data for site: " & CStr(thisSite)
                    End If
                Else
                    If (Rnd < 0.5) Then
                        debug_str = debug_str + "0"
                    Else
                        debug_str = debug_str + "1"
                        one_byte = one_byte + 1
                        half_byte = half_byte + 1
                    End If
                End If
                
                If (i Mod 4) = 3 Then
                    If half_byte < 10 Then
                        str_byte = str_byte + Chr(half_byte + 48)
                    Else
                        str_byte = str_byte + Chr(half_byte + 55)
                    End If
                    half_byte = 0
                End If

                If ((i + 1) Mod 8) = 0 Then
                    If (byte_counter Mod 2) = 0 Then
                        a_16b_data = one_byte
                        str_lsb = str_byte
                    Else
                        a_16b_data = a_16b_data + one_byte * 256
                        str_msb = str_byte
                    End If
                    one_byte = 0
                    str_byte = ""
                    byte_counter = byte_counter + 1
                End If
                
                If ((i + 1) Mod 16) = 0 Then
                    one_afe_data(Int(i / 16)) = a_16b_data
                    debug_str = debug_str + str_msb + str_lsb + " "
                    a_16b_data = 0
                End If

                half_byte = half_byte * 2
                one_byte = one_byte * 2
            Next i 'Loop of 1 bulk read
            
            For m = 0 To 2
                If one_afe_data(m) > 32768 Then
                    one_afe_data(m) = one_afe_data(m) - 65536
                End If
            Next m
            adc_sum = CDbl(one_afe_data(4)) * 65536 + one_afe_data(3)
            If one_afe_data(4) > 32768 Then
                adc_sum = adc_sum - CDbl(65536) * 65536
            End If
            adc_sum = adc_sum / num_of_Samples
            
            adc_sum2 = CDbl(one_afe_data(6)) * 65536 + one_afe_data(5)
            adc_sum2 = adc_sum2 / num_of_Samples
            
            If (adc_sum2 < adc_sum ^ 2) Then
                noise_sigma = -99
            Else
                noise_sigma = Sqr(adc_sum2 - adc_sum ^ 2)
            End If
            
            If loop_count = 0 Then
                adcMax_PLD.pins(0).Value(thisSite) = one_afe_data(0)
                adcMin_PLD.pins(0).Value(thisSite) = one_afe_data(1)
                nf_PLD.pins(0).Value(thisSite) = noise_sigma
            End If
                
            If adcMin_PLD.pins(0).Value(thisSite) > one_afe_data(1) Then
                adcMin_PLD.pins(0).Value(thisSite) = one_afe_data(1)
                'adcMin_PLD.AddPin (chName)
            End If
        
            If adcMax_PLD.pins(0).Value(thisSite) < one_afe_data(0) Then
                adcMax_PLD.pins(0).Value(thisSite) = one_afe_data(0)
                'adcMax_PLD.AddPin (chName)
            End If
        
            If nf_PLD.pins(0).Value(thisSite) < noise_sigma Then
                nf_PLD.pins(0).Value(thisSite) = noise_sigma
                'nf_PLD.Pins(0).Name(thisSite) = chName
                'nf_PLD.AddPin (chName)
            End If
                
            Log_a_line = Log_a_line + CStr(one_afe_data(0)) + " "
            Log_a_line = Log_a_line + CStr(one_afe_data(1)) + " "
            Log_a_line = Log_a_line + CStr(one_afe_data(2)) + " "
            Log_a_line = Log_a_line + Format(adc_sum, "0.00") + " "
            Log_a_line = Log_a_line + Format(adc_sum2, "0.0") + " "
            Log_a_line = Log_a_line + Format(noise_sigma, "0.000")
            
            If Debug_mode = True Then
                TheExec.datalog.WriteComment debug_str
            End If
            TheExec.datalog.WriteComment Log_a_line
            
            ' Go to next site.
            sitestatus = TheExec.Sites.SelectNext(loopTop)
        Loop 'all sites
        
    Next 'loop_count
    
    TheExec.datalog.WriteComment ""
    
    sitestatus = TheExec.Sites.SelectFirst
    Do While sitestatus <> loopDone
            thisSite = TheExec.Sites.SelectedSite
            this_testnum = TheExec.Sites.site(thisSite).TestNumber
            If this_testnum = 72162100 Then
                AFE_noisefloor_test_FailCount.pins(0).Value(thisSite) = 0
            End If
            If adcMin_PLD.pins(0).Value(thisSite) < adc_min_ll Or adcMax_PLD.pins(0).Value(thisSite) > adc_max_ul Or _
               nf_PLD.pins(0).Value(thisSite) < lower_limit Or nf_PLD.pins(0).Value(thisSite) > upper_limit Then
               AFE_noisefloor_test_FailCount.pins(0).Value(thisSite) = AFE_noisefloor_test_FailCount.pins(0).Value(thisSite) + 1
            End If
            sitestatus = TheExec.Sites.SelectNext(loopTop)
        Loop 'all sites
    
    tName = Replace(TheExec.DataManager.instanceName, "_data", "_Min")
    Call TheExec.Flow.TestLimit(resultVal:=adcMin_PLD, LowLimit:=adc_min_ll, testname:=tName, FormatString:="%5d")
    tName = Replace(TheExec.DataManager.instanceName, "_data", "_Max")
    Call TheExec.Flow.TestLimit(resultVal:=adcMax_PLD, highLimit:=adc_max_ul, testname:=tName, FormatString:="%5d")
    tName = Replace(TheExec.DataManager.instanceName, "_data", "_NF")
    Call TheExec.Flow.TestLimit(resultVal:=nf_PLD, LowLimit:=lower_limit, highLimit:=upper_limit, testname:=tName, FormatString:="%6.3f")
    
    Exit Function
 
NoSitesActive:
    
    On Error GoTo 0
    Exit Function
 
ErrorHandler:
    
End Function

Public Function AFE_noisefloor_test_result_VBT() As Long

    Dim tName As String
    tName = TheExec.DataManager.instanceName
    Call TheExec.Flow.TestLimit(resultVal:=AFE_noisefloor_test_FailCount, LowLimit:=0, lowcomparesign:=tlSignGreaterEqual, _
                                highLimit:=7, highcomparesign:=tlSignLessEqual, testname:=TheExec.DataManager.instanceName, _
                                FormatString:="%1d")

End Function

Public Function VBT_AFE_FW_Data_Capt(patternName As String, mosi_Pin As String, miso_Pin As String, start_label As String, _
                            vec_offset As Long, start_Addr As Long, num_of_Bytes As Long, _
                            comments As String) As Long

                'New VBT Function created for F/W based AFE tests  --- Wayne 02/08/2017
                'Data Format:
                '       TEST_ID site# ADC0 ADC1 ADC2 ... ADC15
                '       TEST_ID site# ADC0 ADC1 ADC2 ... ADC15
                '

    'Set error handling
    On Error GoTo ErrorHandler
        
    Dim i As Long
    Dim j As Integer
    Dim k As Long
    Dim m As Integer
    Dim N As Integer

    Dim sitestatus As Long
    Dim thisSite As Long
    Dim loop_count As Integer
    Dim remaining_bytes As Integer
    Dim bytes_of_data As Integer
    Dim byte_counter As Integer

    Dim func_pass_flag As Boolean
    Dim Log_a_line As String
    
    'Dim dump_hex As Boolean
    Dim debug_str As String

    Dim str_byte As String
    Dim str_lsb As String
    Dim str_msb As String
    
    
    Dim this_instance As String
    Dim RetTestNums() As Long
    Dim RetTestCnt As Long
    Dim this_testnum As Long

    Dim half_byte As Integer

    'Dim miso_Pin As String
    'miso_Pin = "SPI_SLV_MISO"

    this_instance = TheExec.DataManager.instanceName
    Call TheExec.DataManager.GetTestNumbers(this_instance, RetTestNums(), RetTestCnt)
    this_testnum = RetTestNums(0)

    If TheExec.Sites.ActiveCount = 0 Then GoTo NoSitesActive
    
    'ReDim passed_func(TheExec.Sites.ExistingCount - 1)
    
    thehdw.digital.Patterns.pat(patternName).Load

    'Connect Pins
    thehdw.PinLevels.ConnectAllPins
    
    'Apply Levels/Timing
    thehdw.PinLevels.ApplyPower
    thehdw.digital.Timing.Load
    thehdw.Wait 0.001

    TheExec.datalog.WriteComment ""
    
    remaining_bytes = num_of_Bytes
    
    'Calculate loop parameters
    loop_count = 0
    
    While remaining_bytes > 0
        'Set up HRAM
        SetupHRAM captSTV, trigSTV, False, True, 0

        If loop_count = 0 Then
            Call update_pattern_fw_wr_16bit_data(patternName, mosi_Pin, start_label, vec_offset, start_Addr, 0, True)
        ElseIf loop_count = 1 Then
            Call update_pattern_fw_wr_16bit_data(patternName, mosi_Pin, start_label, vec_offset, 65535, 0, True) 'Put back addr = FFFF
        End If 'loop 1

        Call thehdw.digital.Patterns.pat(patternName).Start("")
        thehdw.digital.Patgen.HaltWait
        'siteIndex = TheExec.Sites.SelectFirst

        'For siteIndex = 0 To (TheExec.Sites.ExistingCount - 1)
        '    passed_func(siteIndex) = thehdw.digital.Patgen.PatternBurstPassed(siteIndex)
        'Next siteIndex
        
        
        If remaining_bytes < 32 Then
            bytes_of_data = remaining_bytes
        Else
            bytes_of_data = 32
        End If
        
        'loop through active sites
        sitestatus = TheExec.Sites.SelectFirst

        Do While sitestatus <> loopDone
            thisSite = TheExec.Sites.SelectedSite
            
            func_pass_flag = thehdw.digital.Patgen.PatternBurstPassed(thisSite)
            
            'Log_a_line = "FW_DATA_CAPT " + CStr(this_testnum + loop_count) + " " + CStr(thisSite) + " "
            Log_a_line = CStr(this_testnum + loop_count) + " " + CStr(thisSite) + " "
            debug_str = "FW_DATA_Bit_Stream " + CStr(this_testnum) + " " + CStr(thisSite) + " "
            
            If func_pass_flag = True Then
                Log_a_line = Log_a_line + "PASS "
            Else
                Log_a_line = Log_a_line + "FAIL "
            End If
            
            half_byte = 0
            byte_counter = 0
            str_byte = ""
            For i = 0 To 8 * bytes_of_data - 1
                If (TheExec.TesterMode = testModeOnline) Then
                    If (thehdw.digital.HRAM.pins(miso_Pin).PinData(i) = "L") Then
                        'debug_str = debug_str + "0"
                    ElseIf (thehdw.digital.HRAM.pins(miso_Pin).PinData(i) = "H") Then
                        'debug_str = debug_str + "1"
                        half_byte = half_byte + 1
                    Else
                        TheExec.datalog.WriteComment "TEST_ERROR: Unexpected data while capturing data for site: " & CStr(thisSite)
                    End If
                Else
                    If (Rnd < 0.5) Then
                        debug_str = debug_str + "0"
                    Else
                        debug_str = debug_str + "1"
                        half_byte = half_byte + 1
                    End If
                End If
                If (i Mod 4) = 3 Then
                    If half_byte < 10 Then
                        str_byte = str_byte + Chr(half_byte + 48)
                    Else
                        str_byte = str_byte + Chr(half_byte + 55)
                    End If
                    half_byte = 0
                End If

                If ((i + 1) Mod 8) = 0 Then
                    If (byte_counter Mod 2) = 0 Then
                        str_lsb = str_byte
                    Else
                        str_msb = str_byte
                    End If
                    str_byte = ""
                    byte_counter = byte_counter + 1
                End If
                
                If ((i + 1) Mod 16) = 0 Then
                    Log_a_line = Log_a_line + str_msb + str_lsb + " "
                End If
                
                half_byte = half_byte * 2
            Next i 'Loop of 1 bulk read

            TheExec.datalog.WriteComment Log_a_line
            
            ' Go to next site.
            sitestatus = TheExec.Sites.SelectNext(loopTop)
        Loop

        remaining_bytes = remaining_bytes - 32
        loop_count = loop_count + 1
    Wend
    
    TheExec.datalog.WriteComment ""
    
    Exit Function

NoSitesActive:
    
    On Error GoTo 0
    Exit Function
    
ErrorHandler:

End Function


Public Function AFE_Itest_VBT(patternName As Pattern, mosi_Pin As pinList, analog_Pins As pinList, _
                            patVecLabel As String, numAdjustBits As Integer, regBitOffset As Integer, _
                            valueReg1 As Long, valueReg2 As Long, valueReg3 As Long, _
                            valueReg4 As Long, valueReg5 As Long, valueReg6 As Long, _
                            i_code0 As Double, i_step As Double, i_tolerance As Double, _
                            charzTest As Boolean, forceCond As Double, iRange As Double, _
                            RelayModePowered As Boolean, settlingTime As Double, numSamples As Long, _
                            InitLoPins As pinList, initHiPins As pinList, _
                            initHiZPins As pinList, initFloatPins As pinList, _
                            relayOnBits As pinList, relayOffBits As pinList, _
                            comments As String) As Long

                            
                            
    'Set error handling
    'On Error GoTo ErrorHandle
    
    Dim i As Integer, j As Integer
    Dim Max_code As Long
    Dim step_size As Long
    Dim def_value As Long
    Dim adj_code As Long
    
    
    
    Dim NUMSITES As Long
    Dim siteDone() As Boolean
    Dim bpmuIRange As bpmuIRange
    Dim bpmuVRange As bpmuVRange
    Dim vRange As Double
    
    vRange = 2 * v
    
    
    
    Dim measUnit As UnitType
    Dim forceUnit As UnitType
    Dim rtnMeas_Dbl() As Double
    Dim measScale As ScaleType

    Dim meas_LowLimit As Double
    Dim meas_HighLimit As Double

    Dim reg_label(6) As String
    'start_label lbl_1:              > tset_exec      1 1 0 0 X X X X X X X X X X X X 0 0 X X X X X X; // vector 0 / cycle 0 FP AFE CC, CM, DF adust
    'start_label lbl_2:              > tset_exec      1 1 0 0 X X X X X X X X X X X X 0 0 X X X X X X; // vector 129 / cycle 580 VCM_VREF_SEL (bit 11)
    'start_label lbl_3:              > tset_exec      1 1 0 0 X X X X X X X X X X X X 0 0 X X X X X X; // vector 258 / cycle 1160 WOF AFE(L) CC, CM, DF adust
    'start_label lbl_4:              > tset_exec      1 1 0 0 X X X X X X X X X X X X 0 0 X X X X X X; // vector 387 / cycle 1740 WOF AFE(R) CC, CM, DF adust
    'start_label lbl_5:              > tset_exec      1 1 0 0 X X X X X X X X X X X X 0 0 X X X X X X; // vector 516 / cycle 2320 AFE test mux select
    'start_label lbl_6:              > tset_exec      1 1 0 0 X X X X X X X X X X X X 0 0 X X X X X X; // vector 645 / cycle 2900 PS test mux select
    reg_label(1) = "lbl_1"
    reg_label(2) = "lbl_2"
    reg_label(3) = "lbl_3"
    reg_label(4) = "lbl_4"
    reg_label(5) = "lbl_5"
    reg_label(6) = "lbl_6"
    
    
    If (numAdjustBits + regBitOffset) > 16 Then
        'Something is wrong!
        Max_code = 0
        step_size = 0
    Else
        Max_code = 2 ^ numAdjustBits - 1
        step_size = 2 ^ regBitOffset
    End If
    
    
    NUMSITES = TheExec.Sites.ExistingCount
    ReDim siteDone(NUMSITES - 1)
    
    For j = 0 To NUMSITES - 1
        
        siteDone(j) = False
        
    Next j


    
    'Get the complete list of channels
    Call TheExec.DataManager.GetChanListForSelectedSitesByBoard(analog_Pins.Value, chIO, channels, _
        nchannels, nboards, ChPrBrd, nsites, Err, 1)
    
    'Power Down
    If RelayModePowered = False Then
        thehdw.PinLevels.PowerDown
    End If
    
    'Apply Levels, set Relays, init pins, Load Timing
    Call power_up(RelayModePowered, relayOnBits, relayOffBits, _
                initHiPins, InitLoPins, initHiZPins, initFloatPins)
    
    'SETUP BPMU
    bpmuIRange = getBPMUCurrentRange(iRange)
    bpmuVRange = getBPMUVoltageRange(forceCond)
    
    
    If patVecLabel = reg_label(1) Then
        def_value = valueReg1
    ElseIf patVecLabel = reg_label(1) Then
        def_value = valueReg1
    ElseIf patVecLabel = reg_label(2) Then
        def_value = valueReg2
    ElseIf patVecLabel = reg_label(3) Then
        def_value = valueReg3
    ElseIf patVecLabel = reg_label(4) Then
        def_value = valueReg4
    ElseIf patVecLabel = reg_label(5) Then
        def_value = valueReg5
    ElseIf patVecLabel = reg_label(6) Then
        def_value = valueReg6
    Else
        def_value = 0
    End If
    
    
    
    With thehdw.BPMU.pins(analog_Pins.Value)

        'set the BPMU in a ForceVoltageMeasureCurrent Mode
        Call .ModeFVMI(CLng(bpmuIRange), CLng(bpmuVRange))
        
        ' Tester HW does not support clamping current while in the 2uA range,
        '   so check the range setting prior to setting the clamp.
'        If Arg_Irange <> "0" Then
'            ' Clamp to the user specified value or to the limit of the range
'            If Arg_Clamp = TL_C_EMPTYSTR Then
'                RangeVal = thehdw.BPMU.CurrentRangeToValue(CLng(Arg_Irange))
'                .ClampCurrent(CLng(Arg_Irange)) = RangeVal
'            End If
'        End If
'        .ForceVoltage(CLng(Arg_Vrange)) = forceCond
    End With
    
    'Call BPMU_SETUP(meas_Pins.Value, measVoltage, forceVal, iRange, vRange)
    Call getForceMeas_Units(forceCond, False, measUnit, forceUnit, measScale)
    
    'Setup all 6 registers to default values
    Call update_pattern_fw_wr_16bit_data(patternName.Value, mosi_Pin.Value, reg_label(1), 50, valueReg1, 0, True)
    Call update_pattern_fw_wr_16bit_data(patternName.Value, mosi_Pin.Value, reg_label(2), 50, valueReg2, 0, True)
    Call update_pattern_fw_wr_16bit_data(patternName.Value, mosi_Pin.Value, reg_label(3), 50, valueReg3, 0, True)
    Call update_pattern_fw_wr_16bit_data(patternName.Value, mosi_Pin.Value, reg_label(4), 50, valueReg4, 0, True)
    Call update_pattern_fw_wr_16bit_data(patternName.Value, mosi_Pin.Value, reg_label(5), 50, valueReg5, 0, True)
    Call update_pattern_fw_wr_16bit_data(patternName.Value, mosi_Pin.Value, reg_label(6), 50, valueReg6, 0, True)
    
    'Run pattern
    thehdw.digital.Patterns.pat(patternName.Value).Start ""
    thehdw.digital.Patgen.HaltWait
    If TheExec.Sites.ActiveCount = 0 Then Exit Function
    
    
    meas_LowLimit = (i_code0 - i_tolerance) * 1000000
    meas_HighLimit = (i_code0 + i_tolerance) * 1000000
    
    'Do the first measurement
    Call BPMU_FVMI_Multisite_Afe(siteDone, analog_Pins.Value, forceCond, meas_LowLimit, meas_HighLimit, vRange, numSamples, settlingTime)
    
    'Is this single measurement? If numAdjustBits = 0 Then return
    If numAdjustBits = 0 Then GoTo datalog
    'End If
        
    'Is this production or characterization? charzTest True/False

    
    'Sweep through the bits or codes
    i = 1
    
    Do While i <= Max_code
        'APPLY LEVELS, TIMING, RELAYS  --- is this really needed???
        'Call power_up(RelayModePowered, relayOnBits, relayOffBits, _
        '   initHiPins, InitLoPins, initHiZPins, initFloatPins)
        
        adj_code = i * step_size
        
        'update the pattern for ADJ_CODE = i
        Call update_pattern_fw_wr_16bit_data(patternName.Value, mosi_Pin.Value, patVecLabel, 50, 50195 + adj_code, 0, True)
        
        If i_code0 > 0 Then
            meas_LowLimit = i_code0 + (i_step - i_tolerance) * i
            meas_HighLimit = i_code0 + (i_step + i_tolerance) * i
        Else
            meas_LowLimit = i_code0 - (i_step + i_tolerance) * i
            meas_HighLimit = i_code0 - (i_step - i_tolerance) * i
        End If
        
        meas_LowLimit = meas_LowLimit * 1000000
        meas_HighLimit = meas_HighLimit * 1000000

        thehdw.digital.Patterns.pat(patternName.Value).Start ""
        thehdw.digital.Patgen.HaltWait
        If TheExec.Sites.ActiveCount = 0 Then Exit Function
        
        'measure VCM or Itest
        Call BPMU_FVMI_Multisite_Afe(siteDone, analog_Pins.Value, forceCond, meas_LowLimit, meas_HighLimit, vRange, numSamples, settlingTime)

        'measure VCM or Itest
        'BPMU_FVMI_Multisite_Afe(False, analog_Pins.Value, forceCond, meas_LowLimit, meas_HighLimit, _
        '               vRange, numSamples, settlingTime, trimCode_Cur, rtnMeas_Dbl, measUnit, forceUnit)
        
        
        If charzTest Then
            i = i + 1
        Else
            i = i * 2
        End If
    Loop
    
    
    
datalog:
'-----------------------------------------------------------------------------------
'                               DATALOG
'-----------------------------------------------------------------------------------
    
    
        
    Exit Function

ErrorHandle:
  
  MsgBox "Error in BPMU_3Codes_Modify_Vector_VBT Routine!!", vbOKOnly
  Resume Next

End Function
    

Function BPMU_FVMI_Multisite_Afe(siteDone() As Boolean, pins As String, forceVal As Double, LowLimit As Double, highLimit As Double, _
    vRange As Double, samples As Long, settlingTime As Double) As Long
    'ByRef trimCodes_Cur() As Long, ByRef rtnMeas_Dbl() As Double, _
    'Optional measUnitcode As Long, Optional forceUnitcode As Long) As Long '   HiLoLimValid As String,
    
    Dim measured() As Double
    Dim thisChan As Long
    Dim thisChanGrp As Long
    Dim ChanCount As Long
    Dim lngX As Long
    Dim sampleLoop As Long
    Dim sampleAvg As Double
    Dim testStatus As Long      ' LogTestFlag
    Dim parmFlag As Long
    Dim loc As Long
    Dim ReturnStatus As Long

    Dim SortedMeasureVals() As Double
    Dim TmpChanArr() As Long
    ' Temporary variables to save IRange and VRange
    Dim tmpIrange() As Long, tmpVrange() As Long, tmpForceI As Double
    ' Don't know what "loc" is used for in Parametric Datalog
    loc = 0
    
    Dim siteList_BPMU() As Long
            
    Dim siteIndex As Long
    Dim checkforshare As Long
    Dim BpmuShared As Boolean
    Dim getIndex As Long
    Dim Index As Long
    Dim indextemp
    Dim thisSite As Long
    Dim pinName As String
    Dim measPin_bool As Boolean
    Dim bpmuVRange As bpmuVRange
    
    Dim vsim_PLD As New PinListData
    
    'On Error GoTo errHandler        ' trap driver errors
    
    
    'ReDim avgMeasVal(TheExec.Sites.ExistingCount - 1)
    'ReDim SortedMeasureVals(UBound(channelsArr) - LBound(channelsArr))
    
    bpmuVRange = getBPMUVoltageRange(vRange)
    
    vsim_PLD.AddPin (pins)
    vsim_PLD.ResultType = tlResultTypeParametricValue
    
    For thisChan = 0 To nchannelsShr - 1
        SortedMeasureVals(thisChan) = -1
    Next
    
    For thisChanGrp = 0 To ChPrBrd - 1
        'the number of groups to be tested will be based upon the number of channels per
        '   board.
        
        'Determine size of Array to create, the size is based upon how many boards are needed,
        ChanCount = nboards
        
        ReDim TmpChanArr(ChanCount - 1)
        ReDim siteList_BPMU(ChanCount - 1)
         measPin_bool = False
        lngX = 0
        For thisChan = 0 To ChanCount - 1
            If channels(thisChan * ChPrBrd + thisChanGrp) >= 0 Then
                Call thehdw.PinSiteFromChan(channels(thisChan * ChPrBrd + thisChanGrp), chIO, pinName, thisSite)
                If channels(thisChan * ChPrBrd + thisChanGrp) > -1 And siteDone(thisSite) = False Then
                    TmpChanArr(lngX) = channels(thisChan * ChPrBrd + thisChanGrp)
                    siteList_BPMU(lngX) = thisChan
                    lngX = lngX + 1
                    measPin_bool = True
                End If
            End If
        Next thisChan
        If measPin_bool = False Then
            GoTo nextChan
        End If
        If lngX = 0 Then
            lngX = 0
        End If
        
        ChanCount = lngX
        
        ReDim Preserve TmpChanArr(ChanCount - 1)
        ReDim Preserve siteList_BPMU(ChanCount - 1)
        
        ' Connect the bpmu, optionally powering down in order to do a cold connect
'        Call TheHdw.BPMU.Chans(TmpChanArr(0)).ReadRanges(tmpIrange, tmpVrange)
'        tmpForceI = TheHdw.BPMU.Chans(TmpChanArr(0)).ForceCurrent(tmpIrange(0))
'        'thehdw.BPMU.Chans(TmpChanArr).ForceCurrent(bpmu200uA) = 0#
        thehdw.BPMU.Chans(TmpChanArr).Connect

        thehdw.BPMU.Chans(TmpChanArr).ForceVoltage(bpmuVRange) = forceVal
        
        ' Set the Settling Timer for the user specified settling time
        Call thehdw.SetSettlingTimer(settlingTime)
        ' Wait for the settling timer
        Call thehdw.SettleWait(30#)
        If TheExec.Sites.ActiveCount = 0 Then GoTo NoSitesActive
        
        thehdw.BPMU.pins(pins).ClearLatchedOverload
    
        Call thehdw.BPMU.Chans(TmpChanArr).ParallelMeasurement(samples, measured)
        
        thehdw.BPMU.Chans(TmpChanArr).Disconnect
        
        ' Compare the value (either current or voltage)
        For thisChan = 0 To ChanCount - 1
 
            ' Average the results across all samples
            sampleAvg = 0#
            For sampleLoop = 0 To samples - 1
                sampleAvg = sampleAvg + measured(thisChan + sampleLoop * ChanCount)
            Next sampleLoop
            sampleAvg = sampleAvg / samples
            
            'TheExec.datalog.WriteComment "ITEST_Measurement = " + CStr(sampleAvg)
            
            'return function commented out by Wayne
            'will add pass/fail judges here
            
            
            Call thehdw.PinSiteFromChan(TmpChanArr(thisChan), chIO, pinName, thisSite)
            vsim_PLD.pins(0).Value(thisSite) = sampleAvg * 1000000      'in uA
            
            
            'rtnMeas_Dbl(thisSite, trimCodes_Cur(thisSite)) = sampleAvg
            

        Next thisChan
nextChan:
        thehdw.BPMU.Chans(TmpChanArr).Disconnect
        
    Next thisChanGrp
    
    'tName = Replace(TheExec.DataManager.instanceName, "_data", "_DCI")
    
    Call TheExec.Flow.TestLimit(resultVal:=vsim_PLD, unit:=unitCustom, LowLimit:=LowLimit, highLimit:=highLimit, _
                                testname:=TheExec.DataManager.instanceName, FormatString:="%6.2f", CustomUnit:="uA")
    
    
Exit Function

NoSitesActive:
    
    On Error GoTo 0
    Exit Function
    
errHandler:
    ReturnStatus = ReturnStatus + 1     ' count failures
    Resume Next
    ' Don't clear the On Error Goto.  This allows us to continue to
    ' count failures.
End Function

Public Function VBT_AFE_Noise_Floor_char(patternName As String, mosi_Pin As String, miso_Pin As String, start_label As String, _
                            vec_offset As Long, start_Addr As Long, num_of_AFEs As Long, bytes_per_AFE As Long, num_of_Samples As Long, _
                            adc_min_ll As Double, adc_max_ul As Double, afe_nf_limit As Double, ref_nf_limit As Double, _
                            Debug_mode As Boolean, comments As String) As Long


                'New VBT Function created for F/W based AFE Noise Floor test  --- Wayne 02/10/2017
                'Data Format for each channel:
                '       - 0 - AFE# max val
                '       - 1 - AFE# min val
                '       - 2 - AFE# x0 val
                '       - 3,4 - AFE# values sum   (2s compliment)
                '       - 5,6 - AFE# values sum^2 (unsigned)
                '
                '       1. Loop through AFE's
                '       2. Loop through sites
                '       3. Process the data for one AFE channel
                '           (bytes_per_AFE = 7 bytes for Whitney)
                '

    'Set error handling
    'On Error GoTo ErrorHandler
                        
    Dim i As Long
    Dim j As Integer
    Dim k As Long
    Dim m As Integer
    Dim N As Integer

    Dim sitestatus As Long
    Dim thisSite As Long
    Dim loop_count As Integer
    Dim remaining_bytes As Integer
    Dim bytes_of_data As Integer
    Dim byte_counter As Integer

    Dim func_pass_flag As Boolean
    Dim Log_a_line As String

    Dim debug_str As String
    Dim this_instance As String
    Dim tName As String
    Dim chName As String
    Dim RetTestNums() As Long
    Dim RetTestCnt As Long
    Dim this_testnum As Long

    Dim a_16b_data As Long
    Dim one_afe_data(7) As Long
    Dim one_byte As Long
    Dim half_byte As Integer
    Dim str_byte As String
    Dim str_lsb As String
    Dim str_msb As String
    
    Dim min_value As Long
    Dim max_value As Long
    Dim max_noise As Double
    
    Dim adc_sum As Double
    Dim adc_sum2 As Double
    Dim delata_mean_sq As Double
    Dim noise_sigma As Double
    
    Dim adcMin_PLD As New PinListData
    Dim adcMax_PLD As New PinListData
    Dim afe_nf_PLD As New PinListData
    Dim ref_nf_PLD As New PinListData
 
    this_instance = TheExec.DataManager.instanceName
    Call TheExec.DataManager.GetTestNumbers(this_instance, RetTestNums(), RetTestCnt)
    this_testnum = RetTestNums(0)

    If TheExec.Sites.ActiveCount = 0 Then GoTo NoSitesActive
    
    thehdw.digital.Patterns.pat(patternName).Load

    'Connect Pins
    thehdw.PinLevels.ConnectAllPins
    
    'Apply Levels/Timing
    thehdw.PinLevels.ApplyPower
    thehdw.digital.Timing.Load
    thehdw.Wait 0.001
    
    adcMin_PLD.AddPin ("AFE_x")
    adcMax_PLD.AddPin ("AFE_x")
    afe_nf_PLD.AddPin ("AFE_x")
    ref_nf_PLD.AddPin ("REF_x")
    adcMin_PLD.ResultType = tlResultTypeParametricValue
    adcMax_PLD.ResultType = tlResultTypeParametricValue
    afe_nf_PLD.ResultType = tlResultTypeParametricValue
    ref_nf_PLD.ResultType = tlResultTypeParametricValue
    
    TheExec.datalog.WriteComment ""
    
    'Calculate loop parameters
    bytes_of_data = bytes_per_AFE
    
    For loop_count = 0 To num_of_AFEs - 1
        'Set up HRAM
        SetupHRAM captSTV, trigSTV, False, True, 0

        If loop_count = 0 Then
            Call update_pattern_fw_wr_16bit_data(patternName, mosi_Pin, start_label, vec_offset, start_Addr, 0, True)
        ElseIf loop_count = 1 Then
            Call update_pattern_fw_wr_16bit_data(patternName, mosi_Pin, start_label, vec_offset, 65535, 0, True) 'Put back addr = FFFF
        End If 'loop 1

        Call thehdw.digital.Patterns.pat(patternName).Start("")
        thehdw.digital.Patgen.HaltWait
        'siteIndex = TheExec.Sites.SelectFirst

        'For siteIndex = 0 To (TheExec.Sites.ExistingCount - 1)
        '    passed_func(siteIndex) = thehdw.digital.Patgen.PatternBurstPassed(siteIndex)
        'Next siteIndex
        
        'loop through active sites
        sitestatus = TheExec.Sites.SelectFirst

        Do While sitestatus <> loopDone
            thisSite = TheExec.Sites.SelectedSite
            
            func_pass_flag = thehdw.digital.Patgen.PatternBurstPassed(thisSite)
            
            'Log_a_line = "FW_DATA_CAPT " + CStr(this_testnum + loop_count) + " " + CStr(thisSite) + " "
            Log_a_line = CStr(this_testnum + 100 + loop_count) + " " + CStr(thisSite) + " "
            debug_str = "FW_DATA_DEBUG_MODE " + CStr(this_testnum + 200) + " " + CStr(thisSite) + " "
            
            
            If func_pass_flag = True Then
                Log_a_line = Log_a_line + "PASS "
                debug_str = debug_str + "PASS "
            Else
                Log_a_line = Log_a_line + "FAIL "
                debug_str = debug_str + "FAIL "
            End If
            
            If loop_count < 40 Then
                chName = "AFE" + CStr(loop_count)
            ElseIf loop_count < 42 Then
                chName = "REFL" + CStr(loop_count - 40)
            ElseIf loop_count < 82 Then
                chName = "AFE" + CStr(loop_count - 2)
            ElseIf loop_count < 84 Then
                chName = "REFR" + CStr(loop_count - 82)
            End If
            If num_of_AFEs = 2 Then
                'SB AFE's
                chName = "SB_RX" + CStr(loop_count)
            End If
            Log_a_line = Log_a_line + chName + " "
            
            one_byte = 0
            half_byte = 0
            byte_counter = 0
            str_byte = ""
            For i = 0 To 8 * bytes_of_data - 1
                If (TheExec.TesterMode = testModeOnline) Then
                    If (thehdw.digital.HRAM.pins(miso_Pin).PinData(i) = "L") Then
                        'debug_str = debug_str + "0"
                    ElseIf (thehdw.digital.HRAM.pins(miso_Pin).PinData(i) = "H") Then
                        'debug_str = debug_str + "1"
                        one_byte = one_byte + 1
                        half_byte = half_byte + 1
                    Else
                        TheExec.datalog.WriteComment "TEST_ERROR: Unexpected data while capturing data for site: " & CStr(thisSite)
                    End If
                Else
                    If (Rnd < 0.5) Then
                        debug_str = debug_str + "0"
                    Else
                        debug_str = debug_str + "1"
                        one_byte = one_byte + 1
                        half_byte = half_byte + 1
                    End If
                End If
                
                If (i Mod 4) = 3 Then
                    If half_byte < 10 Then
                        str_byte = str_byte + Chr(half_byte + 48)
                    Else
                        str_byte = str_byte + Chr(half_byte + 55)
                    End If
                    half_byte = 0
                End If

                If ((i + 1) Mod 8) = 0 Then
                    If (byte_counter Mod 2) = 0 Then
                        a_16b_data = one_byte
                        str_lsb = str_byte
                    Else
                        a_16b_data = a_16b_data + one_byte * 256
                        str_msb = str_byte
                    End If
                    one_byte = 0
                    str_byte = ""
                    byte_counter = byte_counter + 1
                End If
                
                If ((i + 1) Mod 16) = 0 Then
                    one_afe_data(Int(i / 16)) = a_16b_data
                    debug_str = debug_str + str_msb + str_lsb + " "
                    a_16b_data = 0
                End If

                half_byte = half_byte * 2
                one_byte = one_byte * 2
            Next i 'Loop of 1 bulk read
            
            For m = 0 To 2
                If one_afe_data(m) > 32768 Then
                    one_afe_data(m) = one_afe_data(m) - 65536
                End If
            Next m
            adc_sum = CDbl(one_afe_data(4)) * 65536 + one_afe_data(3)
            If one_afe_data(4) > 32768 Then
                adc_sum = adc_sum - CDbl(65536) * 65536
            End If
            adc_sum = adc_sum / num_of_Samples
            delata_mean_sq = (adc_sum - one_afe_data(2)) ^ 2
                        
            adc_sum2 = CDbl(one_afe_data(6)) * 65536 + one_afe_data(5)
            adc_sum2 = adc_sum2 / num_of_Samples
            
            If (adc_sum2 < delata_mean_sq) Then
                noise_sigma = -99
            Else
                noise_sigma = Sqr(adc_sum2 - delata_mean_sq)
            End If
            
            If loop_count = 0 Then
                adcMax_PLD.pins(0).Value(thisSite) = one_afe_data(0)
                adcMin_PLD.pins(0).Value(thisSite) = one_afe_data(1)
                afe_nf_PLD.pins(0).Value(thisSite) = noise_sigma
            End If
                
            If loop_count = 40 Then
                ref_nf_PLD.pins(0).Value(thisSite) = noise_sigma
            End If
                
                
            If (loop_count < 40) Or ((loop_count > 41) And (loop_count < 82)) Then
                
                If adcMin_PLD.pins(0).Value(thisSite) > one_afe_data(1) Then
                    adcMin_PLD.pins(0).Value(thisSite) = one_afe_data(1)
                    'adcMin_PLD.AddPin (chName)
                End If
            
                If adcMax_PLD.pins(0).Value(thisSite) < one_afe_data(0) Then
                    adcMax_PLD.pins(0).Value(thisSite) = one_afe_data(0)
                    'adcMax_PLD.AddPin (chName)
                End If
            
                If afe_nf_PLD.pins(0).Value(thisSite) < noise_sigma Then
                    afe_nf_PLD.pins(0).Value(thisSite) = noise_sigma
                    'afe_nf_PLD.Pins(0).Name(thisSite) = chName
                    'afe_nf_PLD.AddPin (chName)
                End If
                
            Else
            
                If ref_nf_PLD.pins(0).Value(thisSite) < noise_sigma Then
                    ref_nf_PLD.pins(0).Value(thisSite) = noise_sigma
                End If
            
            End If
                
            Log_a_line = Log_a_line + CStr(one_afe_data(0)) + " "
            Log_a_line = Log_a_line + CStr(one_afe_data(1)) + " "
            Log_a_line = Log_a_line + CStr(one_afe_data(2)) + " "
            Log_a_line = Log_a_line + Format(adc_sum, "0.00") + " "
            Log_a_line = Log_a_line + Format(adc_sum2, "0.0") + " "
            Log_a_line = Log_a_line + Format(noise_sigma, "0.000")
            
            If Debug_mode = True Then
                TheExec.datalog.WriteComment debug_str
            End If
            TheExec.datalog.WriteComment Log_a_line
            
            ' Go to next site.
            sitestatus = TheExec.Sites.SelectNext(loopTop)
        Loop 'all sites
        
    Next 'loop_count
    
    TheExec.datalog.WriteComment ""
    
    tName = Replace(TheExec.DataManager.instanceName, "_data", "_Min")
    Call TheExec.Flow.TestLimit(resultVal:=adcMin_PLD, LowLimit:=adc_min_ll, testname:=tName, FormatString:="%5d")
    tName = Replace(TheExec.DataManager.instanceName, "_data", "_Max")
    Call TheExec.Flow.TestLimit(resultVal:=adcMax_PLD, highLimit:=adc_max_ul, testname:=tName, FormatString:="%5d")
    tName = Replace(TheExec.DataManager.instanceName, "_data", "_NF1")
    Call TheExec.Flow.TestLimit(resultVal:=afe_nf_PLD, LowLimit:=0.5, highLimit:=afe_nf_limit, testname:=tName, FormatString:="%6.3f")
    
    If num_of_AFEs <> 2 Then
        'If not SB AFE test
        tName = Replace(TheExec.DataManager.instanceName, "_data", "_NF2")
        Call TheExec.Flow.TestLimit(resultVal:=ref_nf_PLD, LowLimit:=0.5, highLimit:=ref_nf_limit, testname:=tName, FormatString:="%6.3f")
    End If
    
    
    Exit Function
 
NoSitesActive:
    
    On Error GoTo 0
    Exit Function
 
ErrorHandler:
    
End Function
