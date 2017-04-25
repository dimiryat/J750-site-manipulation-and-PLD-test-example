Attribute VB_Name = "VBT_Module"
Public AFE_noisefloor_test_FailCount As New PinListData

Option Explicit

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
