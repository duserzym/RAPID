VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Range"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarRangeType As Long 'local copy
Private mvarMinValue As Double 'local copy
Private mvarMaxValue As Double 'local copy
Public Enum MCCRangeConverter_Options
    VoltsToMCC_Counts = 0
    MCC_CountsToVolts = 1
    VoltsToMCCTrig = 2
    MCCTrigToVolts = 3
    MCC_CountsToFixedCounts = 4
    FixedCountsToMCC_Counts = 5
End Enum

Public Property Let MaxValue(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxValue = 5
    mvarMaxValue = vData
End Property

Public Property Get MaxValue() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxValue
    MaxValue = mvarMaxValue
End Property

Public Property Let MinValue(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinValue = 5
    mvarMinValue = vData
End Property

Public Property Get MinValue() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinValue
    MinValue = mvarMinValue
End Property

Public Property Let RangeType(ByVal vData As Long)
Attribute RangeType.VB_UserMemId = 0
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RangeType = 5
    mvarRangeType = vData
    
    'Find out what the min & max voltages are from the MCC Range value inputed
    Select Case mvarRangeType
    
        Case BIP60VOLTS
    
            mvarMinValue = -60
            mvarMaxValue = 60
    
        Case BIP20VOLTS
                       
            mvarMinValue = -20
            mvarMaxValue = 20
    
        Case BIP15VOLTS
    
            mvarMinValue = -15
            mvarMaxValue = 15
    
        Case BIP10VOLTS
    
            mvarMinValue = -10
            mvarMaxValue = 10
    
        Case BIP5VOLTS
        
            mvarMinValue = -5
            mvarMaxValue = 5
            
        Case BIP4VOLTS
        
            mvarMinValue = -4
            mvarMaxValue = 4
            
        Case BIP2PT5VOLTS
        
            mvarMinValue = -2.5
            mvarMaxValue = 2.5
            
        Case BIP2VOLTS
        
            mvarMinValue = -2
            mvarMaxValue = 2
            
        Case BIP1PT25VOLTS
        
            mvarMinValue = -1.25
            mvarMaxValue = 1.25
            
        Case BIP1VOLTS
        
            mvarMinValue = -1
            mvarMaxValue = 1
    
    
        Case BIPPT625VOLTS
            
            mvarMinValue = -0.625
            mvarMaxValue = 0.625
    
        Case BIPPT5VOLTS
        
            mvarMinValue = -0.5
            mvarMaxValue = 0.5
    
        Case BIPPT25VOLTS
        
            mvarMinValue = -0.25
            mvarMaxValue = 0.25
    
        Case BIPPT2VOLTS
        
            mvarMinValue = -0.2
            mvarMaxValue = 0.2
    
        Case BIPPT1VOLTS
        
            mvarMinValue = -0.1
            mvarMaxValue = 0.1
    
        Case BIPPT05VOLTS
            
            mvarMinValue = -0.05
            mvarMaxValue = 0.05
    
        Case BIPPT01VOLTS
        
            mvarMinValue = -0.01
            mvarMaxValue = 0.01
    
        Case BIPPT005VOLTS
        
            mvarMinValue = -0.005
            mvarMaxValue = 0.005
    
        Case BIP1PT67VOLTS
        
            mvarMinValue = -1.67
            mvarMaxValue = 1.67
            
        Case BIPPT312VOLTS
        
            mvarMinValue = -0.3125
            mvarMaxValue = 0.3125
            
        Case BIPPT156VOLTS
        
            mvarMinValue = -0.15625
            mvarMaxValue = 0.15625
            
        Case BIPPT125VOLTS
        
            mvarMinValue = -0.125
            mvarMaxValue = 0.125
            
        Case BIPPT078VOLTS
        
            mvarMinValue = -0.078125
            mvarMaxValue = 0.078125
    
        Case UNI10VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 10
            
        Case UNI5VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 5
            
        Case UNI4VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 4
            
        Case UNI2PT5VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 2.5
            
        Case UNI2VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 2
            
        Case UNI1PT67VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 1.67
        
        Case UNI1PT25VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 1.25
        
        Case UNI1VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 1
    
        Case UNIPT5VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 0.5
    
        Case UNIPT25VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 0.25
    
        Case UNIPT2VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 0.2
    
        Case UNIPT1VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 0.1
    
        Case UNIPT05VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 0.05
    
        Case UNIPT02VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 0.02
    
        Case UNIPT01VOLTS
        
            mvarMinValue = 0
            mvarMaxValue = 0.01
            
        Case Else
        
            'Bad value crept in here, be safe and set everything to zero
            mvarMinValue = 0
            mvarMaxValue = 0
            
    End Select
            
End Property

Public Property Get RangeType() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RangeType
    RangeType = mvarRangeType
End Property

Public Function ADWIN_RangeConverter(Optional ByVal voltage As Double = -100, Optional ByVal ADWIN_Val As Long = 100000) As Variant

    Dim LSB As Double
    Dim ADWIN_ZeroVal As Long
    
    'Set ADWIN_ZeroVal to 2^15
    ADWIN_ZeroVal = 2 ^ 15
    
    'Calculate Least Significant Bit (LSB)
    LSB = Abs(mvarMaxValue - mvarMinValue) / 2 ^ 16
    
    'Check to see if values have been passed into the function
    If voltage = -100 And ADWIN_Val = 100000 Then
    
        'no values passed in, exit function
        ADWIN_RangeConverter = -100
        
        Exit Function
        
    End If
    
    'If no voltage passed in, then need to convert ADWIN counts to Voltage
    If voltage = -100 Then
        
        If ADWIN_Val < -2147483468 + ADWIN_ZeroVal Then
            ADWIN_Val = -2147483468 + ADWIN_ZeroVal
        End If
        
        ADWIN_RangeConverter = CDbl((ADWIN_Val - ADWIN_ZeroVal) * LSB)
        
        
    Else
    'Must be a voltage passed in, need to convert it to an ADWIN counts value
        
        ADWIN_RangeConverter = CLng(CLng(voltage / LSB) + ADWIN_ZeroVal)
        
    End If
    
End Function

Private Function Calc_Fixed_Counts(ByVal FSR As Double, ByVal LSB As Double, ByVal PosVolt As Double, Optional ByVal NumbBits As Integer = 16) As Integer

    Dim temp As Long

    'Function that
    'Takes:
    'the absolute range of the voltage,
    'the least significant bit - or graininess of the voltage integer scale,
    'a positive voltage indicating the absolute voltage distance from the minimum voltage
    '
    'Returns:
    'the count corresponding to the RangePos voltage with:
    '    Min Value = -2^(# Bits - 1)
    '    Max Value = -1 + 2^(# Bits - 1)
    '    Median Value = 0
    
    'If RangePos is greater than the max voltage, this function will return -1 + 2^(# Bits - 1)
    If PosVolt > FSR Then PosVolt = FSR

    'Set calculated fixed counts to a long in case they exceed the bound of integers
    temp = CLng(-1 * 2 ^ (NumbBits - 1) + CLng(PosVolt / LSB))
    
    'To catch a rounding error when PosVolt = FSR
    If temp > -1 + 2 ^ (NumbBits - 1) Then temp = -1 + 2 ^ (NumbBits - 1)
        
    Calc_Fixed_Counts = CInt(temp)

End Function

Private Function Calc_MCC_Counts(ByVal FixedCounts As Integer, Optional ByVal NumbBits As Integer = 16) As Integer
        
    'Function that takes a voltage expressed as an integer value from
    '-2^(# Bits - 1) to -1 + 2^(# Bits - 1) without gaps to a gapped value consistent with MCC
    'desire to drive any sane programmer absolutely crazy
    'Fixed Counts:
    '   Min Value = -2^(# Bits - 1)
    '   Max Value = -1 + 2^(# Bits - 1)
    '   Median Value = 0
    '
    'MCC Counts:
    '   Max Value = -1
    '   Min Value = -1 + 2^(# Bits - 1)
    '   Zero Value = -2^(# Bits - 1)
    '
    'If NumbBits < 16, then MCC uses an unsigned integer counts range
    'MCC Counts:
    '   Median Value = 2^(# Bits - 1),  # Bits < 16
    '   Max Value = -1 + 2^(# Bits),    # Bits < 16
    '   Min Value = 0
    
    'Now Check if NumbBits < 16, if so, need to calculate unsigned MCC Counts range
    If NumbBits < 16 Then
    
        'Fixed counts and MCC counts are the same, just shifted by a half-range
        Calc_MCC_Counts = FixedCounts + 2 ^ (NumbBits - 1)
        
    'NumbBits = 16, if > 16, then this code will break
    Else
                
        'Fixed counts and MCC counts are essentially the same, except negative
        'voltages in MCC counts are positive, and positive voltages in MCC
        'counts are negative.  The person who thought this up, i hope, was fired.
        If FixedCounts < 0 Then
        
            Calc_MCC_Counts = CInt(FixedCounts + 2 ^ (NumbBits - 1))
        
        Else
        
            Calc_MCC_Counts = CInt(FixedCounts - 2 ^ (NumbBits - 1))
        
        End If

    End If
End Function

Private Function Calc_MCC_TrigThreshold(ByVal FixedCounts As Integer, Optional ByVal NumbBits As Integer = 16) As Integer

    'If Trigger is 16 bit, then just return normal fixed counts
    If NumbBits = 16 Then
    
        Calc_MCC_TrigThreshold = FixedCounts
        
    'If trigger is < 16 bit, then make fixed counts unsigned, but still continuous
    'and return unsigned values
    Else
        
        Calc_MCC_TrigThreshold = FixedCounts + 2 ^ (NumbBits - 1)

    End If

End Function

Public Function Get_RangeTypeStr() As String

    Dim TempStr As String
    
    Select Case mvarRangeType
    
        Case BIP60VOLTS

            TempStr = "BIP60VOLTS"
    
        Case BIP20VOLTS
                       
            TempStr = "BIP20VOLTS"
    
        Case BIP15VOLTS
    
            TempStr = "BIP15VOLTS"
            
        Case BIP10VOLTS
    
            TempStr = "BIP10VOLTS"
    
        Case BIP5VOLTS
        
            TempStr = "BIP5VOLTS"
            
        Case BIP4VOLTS
        
            TempStr = "BIP4VOLTS"
            
        Case BIP2PT5VOLTS
        
            TempStr = "BIP2PT5VOLTS"
            
        Case BIP2VOLTS
        
            TempStr = "BIP2VOLTS"
            
        Case BIP1PT25VOLTS
        
            TempStr = "BIP1PT25VOLTS"
            
        Case BIP1VOLTS
        
            TempStr = "BIP1VOLTS"
    
        Case BIPPT625VOLTS
            
            TempStr = "BIPPT625VOLTS"
    
        Case BIPPT5VOLTS
        
            TempStr = "BIPPT5VOLTS"
    
        Case BIPPT25VOLTS
        
            TempStr = "BIPPT25VOLTS"
    
        Case BIPPT2VOLTS
        
            TempStr = "BIPPT2VOLTS"
    
        Case BIPPT1VOLTS
        
            TempStr = "BIPPT1VOLTS"
    
        Case BIPPT05VOLTS
            
            TempStr = "BIPPT05VOLTS"
    
        Case BIPPT01VOLTS
        
            TempStr = "BIPPT01VOLTS"
    
        Case BIPPT005VOLTS
        
            TempStr = "BIPPT005VOLTS"
    
        Case BIP1PT67VOLTS
        
            TempStr = "BIPPT67VOLTS"
            
        Case BIPPT312VOLTS
        
            TempStr = "BIPPT312VOLTS"
            
        Case BIPPT156VOLTS
        
            TempStr = "BIPPT156VOLTS"
            
        Case BIPPT125VOLTS
        
            TempStr = "BIPPT125VOLTS"
            
        Case BIPPT078VOLTS
        
            TempStr = "BIPPT078VOLTS"
    
        Case UNI10VOLTS
        
            TempStr = "UNI10VOLTS"
            
        Case UNI5VOLTS
        
            TempStr = "UNI5VOLTS"
            
        Case UNI4VOLTS
        
            TempStr = "UNI4VOLTS"
            
        Case UNI2PT5VOLTS
        
            TempStr = "UNI2PT5VOLTS"
            
        Case UNI2VOLTS
        
            TempStr = "UNI2VOLTS"
            
        Case UNI1PT67VOLTS
        
            TempStr = "UNI1PT67VOLTS"
        
        Case UNI1PT25VOLTS
        
            TempStr = "UNI1PT25VOLTS"
        
        Case UNI1VOLTS
        
            TempStr = "UNI1VOLTS"
    
        Case UNIPT5VOLTS
        
            TempStr = "UNIPT5VOLTS"
    
        Case UNIPT25VOLTS
        
            TempStr = "UNIPT25VOLTS"
    
        Case UNIPT2VOLTS
        
            TempStr = "UNIPT2VOLTS"
    
        Case UNIPT1VOLTS
        
            TempStr = "UNIPT1VOLTS"
    
        Case UNIPT05VOLTS
        
            TempStr = "UNIPT05VOLTS"
    
        Case UNIPT02VOLTS
        
            TempStr = "UNIPT02VOLTS"
    
        Case UNIPT01VOLTS
        
            TempStr = "UNIPT01VOLTS"
            
        Case Else
        
            'Bad value crept in here, display "ERROR"
            TempStr = "ERROR"

    End Select
    
    Get_RangeTypeStr = TempStr

End Function

Public Function MCC_RangeConverter(ByRef voltage As Double, ByRef Counts As Integer, ByRef FixedCounts As Integer, ByVal UserOption As MCCRangeConverter_Options, Optional ByVal NumbBits As Integer = 16) As Long

    Dim LSB As Double   'Least significant bit
    Dim FSR As Double   'Full Scale Range
    Dim RangePos As Double  'Positive number indicating distance of Voltage from the minimum value of the range
    
    'Options controls what this function does
    '
    '0 = convert voltage into -2^(# bits -1) to -1 + 2^(# bits -1) signed integer MCC board counts
    '    where -2^(# bits -1) maps to 2^(# bits -1) and -1 maps to -2 + 2^(# bits) in unsigned integers
    '    and for +/- voltage ranges
    '    MaxVoltage = 2^(# bits -1), and MinVoltage = -1, Zero = 0
    '
    '1 = convert MCC signed integer board counts into voltage
    '
    '2 = convert voltage to MCC trigger threshold
    '
    '3 = convert MCC trigger threshold to a voltage
    '
    '4 = convert MCC signed integer board counts into fixed continuous board counts
    '
    '5 = continuous integer range from -2^(# bits -1) to -1 + 2^(# bits -1),
    '    where -2^(# bits -1) = MinVoltage in the range, and -1 + 2^(# bits -1)
    '    = maximum voltage - i.e. what the @#$%-heads at MCC should have used
    '    for their boards counts, but didn't into the MCC board counts
    
    'Switch on error checking in case this function is called before the min and max
    'values have been set
    On Error Resume Next
    
        FSR = Abs(mvarMaxValue - mvarMinValue)
        
        If Err.number <> 0 Then
        
            'max and min values haven't been set yet, send the error number along and exit the sub
            MCC_RangeConverter = Err.number
            
            Exit Function
            
        End If
    
    'Turn normal error flow back on
    On Error GoTo 0
        
    'Calculate the least significant bit from the function inputs
    LSB = FSR / 2 ^ NumbBits
    
    'Calculate the absolute difference between the minimum voltage
    'and the desired voltage
    RangePos = Abs(mvarMinValue - voltage)
    
    'Now switch to select case for the different output options
    Select Case UserOption
    
        'Convert voltage to counts
        Case 0
            
            'First Calculate the fixed counts where:
            'Min Value = -2^(# bits - 1)
            'Max Value = -1 + 2^(# bits - 1)
            'Median Value = 0
            
            FixedCounts = Calc_Fixed_Counts(FSR, LSB, RangePos, NumbBits)
    
            'Now need to convert Fixed Counts to MCC board counts
            '(Poo on MCC!)
            Counts = Calc_MCC_Counts(FixedCounts, NumbBits)
            
            
        'Convert counts to voltage
        Case 1
        
            'First convert the counts to fixed counts
            FixedCounts = MCCcounts_To_FixedCounts(Counts, NumbBits)
            
            'Now multiply the fixed counts by the LSB
            voltage = CDbl(FixedCounts * LSB)
            
        
        'Convert voltage to trigger threshold
        Case 2
        
            'First convert the voltage to fixed continuous counts
            FixedCounts = Calc_Fixed_Counts(FSR, LSB, RangePos, NumbBits)
        
            'Now need to convert fixed counts to threshold counts
            Counts = Calc_MCC_TrigThreshold(FixedCounts, NumbBits)
                    
                    
        'Convert trigger threshold to voltage
        Case 3
            
            'First convert the trigger threshold to fixed continuous counts
            FixedCounts = MCCTrig_To_FixedCounts(Counts, NumbBits)
            
            'Now multiply the fixed counts by the LSB
            voltage = CDbl(FixedCounts * LSB)
            
            
        'Convert MCC Counts to continuous fixed counts
        Case 4
        
            FixedCounts = MCCcounts_To_FixedCounts(Counts, NumbBits)
            
            
        'Convert continuous fixed counts to MCC Counts
        Case 5
        
            Counts = Calc_MCC_Counts(FixedCounts, NumbBits)
            
        
        Case Else
        
            'Bad user input, do nothing
            
    End Select
    
End Function

Private Function MCCcounts_To_FixedCounts(ByVal MCCCounts As Integer, Optional ByVal NumbBits As Integer = 16) As Integer

    Dim UnsignedCounts As Integer

    'Function takes MCC board counts: -2^(# Bits - 1) to -1 + 2^(# Bits - 1)
    'and converts those counts to continuous fixed counts
    
    'If MCC Board counts < 0, add 2^(# bits - 1) to the value
    If MCCCounts < 0 Then
    
        UnsignedCounts = CInt(MCCCounts + 2 ^ (NumbBits - 1))
    
    'Else, MCC ounts > 0, subtract 2^(# bits - 1)
    Else
            
        UnsignedCounts = CInt(MCCCounts - 2 ^ (NumbBits - 1))
        
    End If
    
    'If UnsignedCounts > -1 + 2^(# Bits - 1), then subtract 65536 from it
    'recast that as an integer and return the resulting value
    'Else, just recast UnsignedCounts as an integer and return that
    If UnsignedCounts > -1 + 2 ^ (NumbBits - 1) Then
    
        MCCcounts_To_FixedCounts = CInt(UnsignedCounts - 2 ^ NumbBits)
        
    Else
        
        MCCcounts_To_FixedCounts = CInt(UnsignedCounts)
        
    End If
    
End Function

Private Function MCCTrig_To_FixedCounts(ByVal MCCCounts As Integer, Optional ByVal NumbBits As Integer = 16) As Integer

    'If Trigger is 16 bit, then just return normal fixed counts
    If NumbBits = 16 Then
    
        MCCTrig_To_FixedCounts = MCCCounts
        
    'If trigger is < 16 bit, then make fixed counts unsigned, but still continuous
    'and return unsigned values
    Else
        
        MCCTrig_To_FixedCounts = MCCCounts - 2 ^ (NumbBits - 1)
        
    End If

End Function

