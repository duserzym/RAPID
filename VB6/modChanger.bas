Attribute VB_Name = "modChanger"
Option Explicit

' This module is for control of the physical sample changer.
' The changer form is for control of the sample placement.

Private curpos As Long           ' Current changer slot under holder
Private lastPos As Long          ' Previous slot under holder
                                 ' (used for moving to holes)

Public Function Changer_isHole(ByVal num As Integer) As Boolean
    ' This function returns true if 'num' identifies any hole in the
    ' sample changer.
        
    'Toni S. 2012 If XY table there is only one hole
    If UseXYTableAPS Then
        If num = HoleSlotNum Then
            Changer_isHole = True
        Else
            Changer_isHole = False
        End If
    Else
        If num Mod HoleSlotNum = 0 Then
            Changer_isHole = True
        Else
            Changer_isHole = False
        End If
    End If
End Function

'-----------------------------------------------------------------------------
'   Changer_Load
'
'   Description:        This procedure requests that the user load the sample
'                       changer and press OK.  It returns true if user says
'                       okay, false if not.
'   Revision History:
'      Albert Hsiao     2/19/99      added directions to dialog
'
Function Changer_Load(isUp As Boolean) As Boolean
    Dim QueryStr As String, dirstr As String
    Dim Response As VbMsgBoxResult
    
    If isUp Then dirstr = "up" Else dirstr = "down"
    
    QueryStr = "Make sure that the samples are loaded in the " & dirstr & _
        " direction and " & vbCr & "the glass rod sticks about 1 " & _
        "mm through the " & vbCr & "plexiglass plate." & vbCr & _
        vbCr & "Type anything to make the system test the " & _
        "limit switches."
    Response = MsgBox(QueryStr, vbOKCancel, "Notice!")
    Select Case Response
        Case vbOK:
            ' User selected okay, continue
            Changer_Load = True
        Case Else:
            ' Return to previous, user doesn't want to start
            Changer_Load = False
    End Select
End Function

Public Sub Changer_MoveTo(ByVal target As Long)
    ' This procedure moves the changer from the current position
    ' to the position specified.
    
    If DEBUG_MODE Then frmDebug.Msg "Move to hole" & target
    
    frmDCMotors.ChangerMotortoHole (val(target))
    curpos = frmDCMotors.ChangerHole           ' Current position is changed
End Sub

Public Sub Changer_NearestHole()
    ' This routine determines the location of the hole nearest to that
    ' of the present sample changer location, curPos, and moves the holder
    ' to it.  The last position moved from is stored in lastPos.

    Dim nearest_hole_pos As Integer
    
    nearest_hole_pos = Find_NearestChangerHole()
    
    If nearest_hole_pos = curpos Then Exit Sub
    
    Changer_MoveTo nearest_hole_pos
        
End Sub

'-----------------------------------------------------------------------------
'   Changer_ProcessSample
'
'   Description:       This is where we set the stuff for the sample changer
'                      loops.  We should already have made choices about Af
'                      demags, up or down, etc.  remember that we will
'                      increment the loop through the sample changer
'                      positions AFTER each measurement is done; control
'                      should reach here with the   variable having a
'                      valid sample or hole position (for a holder
'                      measurement) in place.
'   Revision History:
'       Albert Hsiao    2/19/99    added up and down directions
'
' updated to support multiple files
Public Sub Changer_ProcessSample(ByVal slot As Integer, Optional ReturnSample As Boolean = True)
    
    Dim doUp As Boolean
    Dim doBoth As Boolean
    Dim steps As Long
    Dim processingSample As Sample
        
    If Prog_halted Then Exit Sub
    
    
    If Changer_isHole(slot) Then
        Set processingSample = SampleHolder
        doUp = True
        doBoth = False
    Else
        Set processingSample = SampleIndexRegistry(MainChanger.ChangerFileName(slot)).sampleSet(MainChanger.ChangerSampleName(slot))
        doUp = processingSample.Parent.doUp
        doBoth = processingSample.Parent.doBoth
    End If
 
    
    ' First, move to the sample
    Changer_MoveTo (slot)
    '
    If Not Changer_isHole(slot) Then
        ' Read the specimen file
      
        ' if we're in rockmag mode and IRM is enabled, discharge the IRM coil
        ' before loading the sample
        
        If processingSample.Parent.RockmagMode And EnableAxialIRM Then
            frmIRMARM.optCoil(0).Value = True
            frmIRMARM.FireIRM 0
        End If
        
        processingSample.sampleHole = slot ' Set the hole number
        
        ' Now lower the sample holder GENTLY, turn on the vacuum, and go back
        ' to the home position
        'Motor_Home Motor_IdVert, 0      ' Stop at top of sample
        SampleHeight = MotorUPDN_home() - SampleBottom  ' SampleTop should now be the distance from
                                        ' the zero position to the top of the sample;
                                        ' use it to calculate height of sample .

'        Motor_WaitStop ("UPDOWN")      '   Wait for motor to stop
        
        frmVacuum.ValveConnect (True)            ' Grab the sample
        DelayTime (0.3)                    ' Pause to make sure connection is good
        'Motor_Home Motor_IdVert, 1      ' Lift back up to the sense switch
        
        
        'MotorUpDn_Move 0.05 * SampleTop, 1          ' move up to 0 position medium
        
        SampleHeight = SampleHeight - frmDCMotors.HomeToTop
        processingSample.SampleHeight = SampleHeight
        'Motor_MoveLoadLift              ' Try to lift sample 5mm higher
        
        Changer_NearestHole             ' Move to the nearest hole
        
        'Motor_MoveLoadToZero            ' Move from load position down to zero
        'ok - the new motors are scary.  First move down to the top of the
        'AF coils gently, then high-tail it to the zero position
        
        'MotorUpDn_Move Int((SCoilPos + AFPos) / 2), 1
        
        Set processingSample = SampleIndexRegistry(MainChanger.ChangerFileName(slot)).sampleSet(MainChanger.ChangerSampleName(slot))
        Measure_TreatAndRead processingSample, True ' Read the sample data
        'Motor_MoveZeroToLoad            ' Move back up to load position
        
        If ReturnSample Then
            MotorUpDn_Move 0, 2           '' move up to 0 position fast
            Changer_MoveTo (lastPos)        ' Move back to sample position
            'steps = MotorUPDN_home()         ' ignore steps here
            frmDCMotors.SampleDropOff
            frmVacuum.ValveConnect (False)  ' Drop sample into slot
            DelayTime DropoffVacuumDelay
            MotorUpDn_Move 0, 1           '' move up to 0 position medium
        End If
        '
      Else
        ' We want to make a holder measurement
        'Motor_MoveLoadToZero            ' Move from load position down to zero
        
        'MotorUpDn_Move ZeroPos, 2
        Measure_TreatAndRead processingSample, True ' Read the sample data
        
        If ReturnSample Then MotorUpDn_Move 0, 2           '' move up to 0 position fast
        '
    End If
End Sub

Sub Changer_TestLimits()
    ' This procedure homes the vertical motor in the top
    ' position, and makes sure the sample changer is in
    ' the nearest home position.
    '
    MotorUPDN_TopReset
    'MotorChanger_Move (0)
    '
    '
End Sub

Public Function Changer_ValidSlot(ByVal num As Double) As Boolean
    ' This function determines whether the slot at 'num' is a
    ' valid sample slot.
    
    Changer_ValidSlot = Changer_ValidStart(num) And Not Changer_isHole(num)
End Function

Public Function Changer_ValidStart(ByVal num As Double) As Boolean
    ' This function determines whether the slot at 'num' is a
    ' valid slot.  It must be between SLOTMIN and SLOTMAX.
    
    Changer_ValidStart = (num >= SlotMin And num <= SlotMax And num = Int(num))
End Function

Public Function Find_NearestChangerHole() As Integer

    Dim i, k As Integer
    Dim upper, lower As Integer
    
    ' If current position unknown, query user.
    frmChanger.GetCurrentChangerPos
    lastPos = frmDCMotors.ChangerHole
    curpos = lastPos
    
    If Not UseXYTableAPS Then
    'Using Chain Drive
    
        If DEBUG_MODE Then frmDebug.Msg "Search for hole near " & curpos
        
        'Check for case where current position = nearest hole
        If Changer_isHole(curpos) Then
            Find_NearestChangerHole = curpos
            Exit Function
        End If
        
        ' Find nearest hole by brute force.  I is the nearest hole above, K below
        i = curpos
        Do
            i = i + 1
            If i > SlotMax Then i = i - SlotMax
        Loop While Not Changer_isHole(i)
        
        k = curpos
        Do
            k = k - 1
            If k < SlotMin Then k = k + SlotMax
        Loop While Not Changer_isHole(k)
      
        ' Must worry about crossing the SLOTMAX barrier.
        upper = i - curpos: If upper < 0 Then upper = upper + SlotMax
        lower = curpos - k: If lower < 0 Then lower = lower + SlotMax
    
        If upper <= lower Then
            ' Closest hole is a higher number:  Move it to the nearest hole!
            Find_NearestChangerHole = i
        Else
            Find_NearestChangerHole = k
        End If
    Else
        'Using XY Table
        
        Find_NearestChangerHole = HoleSlotNum
    
    End If

End Function

