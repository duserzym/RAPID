'<ADbasic Header, Headerversion 001.001>
'<Header End>
'ADBASIC-SYSTEM-FILE--V1.00--0000--0000--INC--L16
'
'###############################################################
'#                        ADWL16.inc                           #
'#              INCLUDE-file for ADwin-Light 16                #
'###############################################################
'#                                                             #
'# Contents: Include functions for ADwin-Light 16              #
'#                                                             #
'#  13.06.2002  MF   creation                                  #
'#  14.08.2002  JK   cntno   ->  (cntno)                       #
'#  02.11.2005  MF   Seq Mode       (only for L16 Rev B)       #
'#  26.01.2006  MF   SSI_*           (only L16 REV B / DIO_02) #
'#  26.01.2006  MF   DIGIN/OUT_Long_E (only DIO_02 or DIO_03)  #
'#  24.02.2006  MF   L16_MODE         (only L16 REV B )        #
'#  29.03.2006  MF   SEQ_INIT         (only L16 REV B )        #
'#  23.08.2006  UR   LS-BUS                                    #
'#  13.09.2006  MF   Bugfix (LS-Bus)                           #
'#  20.09.2006  MF   LS-Bus optimized for smaler program-code  #
'#  11.05.2007  MF   designed for ADbasic 5                    #
'#  25.05.2007  MF   bugfix seq_init                           #
'#  14.06.2007  OL   New sub: Read_Msg_Con                     #
'#  18.02.2009  MF/OL read_msg_con                             #
'#  19.02.2009  MF  Transmit_Status                            #
'#  19.03.2009  MF   '''+                                      #
'#                                                             #
'#    Attention:                                               #
'#   - Do not copy the first source line (containing           #
'#     "'ADBASIC-SYSTEM-FILE...") to user files !              #
'#     Otherwise the Syntax color may be not correct.          #
'#                                                             #
'###############################################################

#define  enadrwr  204000dch ''<HIDE>     'wr 13            Address-table to access the hardware
#define  enadrrd  204001dch ''<HIDE>     'wr 29
#define  rdcan    204000fch ''<HIDE>     'rd 15
#define  wrcan    204000fch ''<HIDE>     'wr 15
#define  rescan   204001fch ''<HIDE>    'rd 14

import l16.li9

#if make_lib=off then  
dim can_msg[9] as Long         ' array with 9 elements
#endif


'---------------------------------------------------------------
' If a low-priority process accesses a LS Bus Modul it must not be
' interrupted by any other process !!!
'---------------------------------------------------------------
SUB SUB_BEGIN_CRITICAL_SECTION() ''<HIDE>
#IF PRIORITAET = LOW
  BEGIN_CRITICAL_SECTION
#ENDIF
ENDSUB

SUB SUB_END_CRITICAL_SECTION() ''<HIDE>
#IF PRIORITAET = LOW
  END_CRITICAL_SECTION
#ENDIF
ENDSUB


'***************************************************************
'***************** DIO on Light 16 extension *******************
'***************************************************************

'***************************************************************
''' Initialisation of digital I/O
''' value : each bit corresponds to 8 bit
'''+     0 = input
'''+     1 = output
'''*************************************************************
SUB Conf_DIO_E(value)
  Poke(2040046CH,value)
ENDSUB


'***************************************************************
''' Reads the digital inputs (channel 0 - 15)
'''*************************************************************
FUNCTION Digin_Word1_E() as Long
  Digin_Word1_E=peek(20400454H)
ENDFUNCTION


'***************************************************************
''' Reads the digital inputs (channel 16 - 31)
'''*************************************************************
FUNCTION Digin_Word2_E() as Long
  Digin_Word2_E=peek(20400464H)
ENDFUNCTION

'***************************************************************
''' Reads the digital inputs (channel 0 - 31)
'''*************************************************************
FUNCTION Digin_Long_E() as Long
  Digin_Long_E=peek(20400458H)
ENDFUNCTION


'***************************************************************
''' Sets the digital outputs (channel 0 - 15)
''' value: Bit pattern, corresponding to the TTL level desired at the digital outputs.
'''*************************************************************
SUB Digout_Word1_E(value)
  Poke(20400454H,value)
ENDSUB

'***************************************************************
''' Sets the digital outputs (channel 0 - 31)
''' value: Bit pattern, corresponding to the TTL level desired at the digital outputs.
'''*************************************************************
SUB Digout_Long_E(value)
  Poke(20400458H,value)
ENDSUB


'***************************************************************
''' Sets the digital outputs (channel 16 - 31)
''' value: Bit pattern, corresponding to the TTL level desired at the digital outputs.
'''*************************************************************
SUB Digout_Word2_E(value)
  Poke(20400464H,value)
ENDSUB


'***************************************************************
''' Sets selected digital outputs (channel 0-15)
''' set: Bitstring indicating the outputs which should be set
'''*************************************************************
SUB Digout_Set1_E(set)   'Set Register Bit 0-15
  Poke(20400474H,set)
ENDSUB


'***************************************************************
''' Sets selected digital outputs (channel 16-31)
''' set: Bitstring indicating the outputs which should be set
'''*************************************************************
SUB Digout_Set2_E(set)   'Set Register Bit 16-31
  Poke(20400484H,set)
ENDSUB


'***************************************************************
''' Resets selected digital outputs (channel 0-15)
''' reset: Bitstring indicating the outputs which should be reset
'''*************************************************************
SUB Digout_Reset1_E(clear) 'Reset Register Bit 0-15
  Poke(20400494H,clear)
ENDSUB


'***************************************************************
''' Resets selected digital outputs (channel 16-32)
''' reset: Bitstring indicating the outputs which should be reset
'''*************************************************************
SUB Digout_Reset2_E(clear) 'Reset Register Bit 16 - 31
  Poke(204004A4H,clear)
ENDSUB



'***************************************************************
'************************ Counter ******************************
'***************************************************************
''' sets the counter values of one or more counters on the specified module to the value 0 (zero).
''' pattern: Bit pattern.
'''+             Bit = 0: no influence.
'''+             Bit = 1: set counter to zero.
'''*************************************************************
SUB CNT_Clear(pattern)
  ' clear all counters selected with a bit 1
  poke(20400310h,pattern)
ENDSUB


' **************************************************************
''' enables or disables one or more counters on the specified module
'''pattern: Enable counters according to the bit pattern, for assignment of the counters see table.
'''+         Bit = 0: Disable counter / block.
'''+         Bit = 1: Enable counter / release
SUB CNT_Enable(pattern)
  ' enable all counters selected with a bit 1
  poke(20400300h,pattern)
ENDSUB


' **************************************************************
'''defines the operating mode of all counters by selecting which clock input they use according to the bit pattern in pattern.
'''pattern:
'''+              Bit = 0: external clock input (CLK/DIR or A/B).
'''+              Bit = 1: internal clock input (5 MHz or 20 MHz).
SUB CNT_Mode(pattern)
  ' Select Counter-mode  (0 = CLK external, 1 = CLK internal(PWM))
  poke(20400340h,pattern)
ENDSUB

' **************************************************************
''' defines the operating mode for all counters (depending on Cnt_Mode) according to the given bit pattern.
'''+ Bit value in pattern
''' pattern:
'''+  CNT_MODE=0
'''+             pattern 00 = 4-edge evaluation    
'''+             pattern 11 = Clock and direction input
'''+  CNT_MODE=1
'''+             pattern 00 = Reference clock 20 MHz
'''+             pattern 11 = Reference clock 5 MHz
''' **************************************************************
SUB CNT_Set(pattern)
  ' Set Counter   (0 = 4fach, 1 = CLK/DIR)
  ' Set PWM-Counter Frequency (0 = 20 MHz, 1 = 5 MHz)
  poke(20400350h,pattern)
ENDSUB


' **************************************************************
'''sets the function of the CLR/LATCH input of one or more counters.
'''pattern:
'''+        Bit = 0: Set CLR-mode.
'''+        Bit = 1: Set LATCH-mode.
SUB CNT_InputMode(pattern)
  ' Set Counterinputline   (0 = Clear counter, 1 = latch counter)
  poke(20400330h,pattern)
ENDSUB


' **************************************************************
'''transfers the current counter values of one or more counters on the specified module into the respective latch register(s) (= to latch).
'''pattern:
'''+        Bit = 0: no function.
'''+        Bit = 1: transfer counter values into Latch A .
SUB CNT_Latch(pattern)
  ' latch all counters selected with a Bit 1
  poke(20400320h,pattern)
ENDSUB


' **************************************************************
'''copies the current counter value into Latch A and returns it as return value.
'''cntno: Counter number: 1.2; L16-CO1: 1.
FUNCTION CNT_Read(cntno) AS Long
  ' performs a software-latch and reads the selected counter
  ' ATTENTION : this function may NOT be used, if the latch is
  '  done by an external signal (e.g. period-time or pulse
  '  low/high-time measurements) !

  dim cadr as Long
  dim cntsel as Long

  cntsel = shift_left(1,(cntno-1))
  poke(20400320h, cntsel)
  cadr = 204001F4h + (cntno) * 16 ' parentheses bec.  i + 1
  ' could be passed as cntno
  cnt_read = peek(cadr)
  ' (we do not need multiple reads, because for software
  '  latch the synchronisation is done by hardware.)
ENDFUNCTION


' **************************************************************
''' returns the value of a counter previously stored in Latch A.
''' cntno: Counter number: 1.2; L16-CO1: 1.
FUNCTION CNT_ReadLatch(cntno) AS Long
  ' reads the latch register triggerd by software-latch or
  '  signal-rise (of an external signal)

  dim cadr as Long
  dim ih as Long

  cadr = 204001F4h + (cntno) * 16 ' parentheses bec.  i + 1
  ' could be passed as cntno
  cnt_readlatch = peek(cadr)

  ' case a : the latch register is triggered by a software- latch :
  ' ---------------------------------------------------------------
  ' In this case, there is no need for multiple reads, because
  '  the synchronisation is done by hardware (see also the
  '  comments at CNT_READ()). The following instructions are
  '  not necessary, but they also do not harm (except wasting
  '  time). Therefore they are a potential for speed optimi-
  '  sations.

  ' case b : the latch register is triggered by an external signal :
  ' ----------------------------------------------------------------
  ' Check, if there was a (from an external signal triggered)
  '  LATCH just in the moment, when we read in the latch
  '  register. (This is neccessary because this case is not
  '  handled by the hardware.)
  '  If we have this case, the second reading of the latch
  '  register differs from the first reading and therefore
  '  one of the readings may be corrupt. Read a third time to
  '  be sure to have a valid reading.
  '        The measured time for these three readings is 25
  '        cycles = 625 ns   => the (period) time between two
  '        successive (external signal triggered) LATCHs may
  '        NOT be shorter than this limit.
  '        => valid measurements (using this method with LATCH
  '        done by external signal) cannot be guaranteed for
  '        smaller LATCH (period) times than 625 ns. (Longtime
  '        tests were done and o.k. with 664 ns). To have a
  '        safety margin for future changes in the ADbasic
  '        compiler, our statement to the customer is the
  '        limit of 1000 ns.

  ih = peek(cadr)
  if (ih <> cnt_readlatch) then
    cnt_readlatch = peek(cadr)
  endif
ENDFUNCTION


'***************************************************************
''' returns the value of a counter previously stored in Latch B.
''' cntno: Counter number: 1.2.
FUNCTION CNT_ReadFLatch(cntno) AS Long
  ' reads the latch register triggerd by signal-falling edge
  dim cadr as Long
  dim ih as Long

  cadr = 204001F8h + (cntno) * 16 ' parentheses bec.  i + 1
  ' could be passed as cntno
  cnt_readflatch = peek(cadr)
  ' (please see comments to READLATCH)
  ih = peek(cadr)
  if (ih <> cnt_readflatch) then
    cnt_readflatch = peek(cadr)
  endif
ENDFUNCTION


'***************************************************************
''' reads the status of a counter input
'''+   Bit 0 Reset-Enable (value which was set by Cnt_ClearEnable) Counter 1
'''+   Bit 1 Reset-Enable (value which was set by Cnt_ClearEnable) Counter 2
'''+   Bit 4 CLR-/LATCH-input (signal is not changing states) Counter 1
'''+   Bit 5 CLR-/LATCH-input (signal is not changing states) Counter 2
'''+   Bit 16 Signal A (signal is not changing states) Counter 1
'''+   Bit 17 Signal B (signal is not changing states) Counter 1
'''+   Bit 18 Signal A (signal is not changing states) Counter 2
'''+   Bit 19 Signal B (signal is not changing states) Counter 2
'''+   Bit 24 Correlation error* (signals A and B are identical, they are not phase-shifted by approx. 90∞) 
'''+          Counter 1
'''+   Bit 25 Line error (cable not connected or the line is broken)
'''+          Auto-Reset (is reset during reading out) Counter 1
'''+   Bit 26 Correlation error* (signals A and B are identical, they are not phase-shifted by approx. 90∞) 
'''+          Counter 2
'''+   Bit 27 Line error (cable not connected or the line is broken)
'''+          Auto-Reset (is reset during reading out) Counter 2
'''***************************************************************
FUNCTION CNT_GetStatus() AS Long
  CNT_GetStatus = PEEK(20400370h)
ENDFUNCTION


'***************************************************************
''' enables the hardware clear input of a counter
''' pattern: bitcombination ; each bit corresponds with counter
'''*************************************************************
SUB CNT_ClearEnable(pattern)
  dim value2 as Long
  value2=pattern AND 0FH
  Poke(20400370H,value2)
endsub


'***************************************************************
'***************************** CAN *****************************
'***************************************************************
'''Enables a message object for receiving messages.
'''+Disables all interrups of the message object.
'''msg_no: Number of the message object
'''Identifier : incomming frames are accepted by this message object, if the identifier matches the frame identifier
'''extended: Determines whether the identifier is 11 or 29 Bit
'''+      0 = 11 Bit Identifier
'''+      1 = 29 Bit Identifier
'''***************************************************************
sub En_Receive(msg_no,identifier,extended)
  set_can_reg(msg_no * 16, 149)
  if (extended = 0) then
    set_can_reg(msg_no * 16 + 6, 128)                         'read; DLC = 8 ; 11 Bit Identifier
    set_can_reg(msg_no * 16 + 2, (shift_right(identifier,3)and 255))  'Set Identifier
    set_can_reg(msg_no * 16 + 3, (shift_left(identifier,5)and 255))
  else
    set_can_reg(msg_no * 16 + 6, 132)                         'read; DLC = 8 ; 29 Bit Identifier
    set_can_reg(msg_no * 16 + 2, (shift_right(identifier,21)and 255)) 'Set Identifier
    set_can_reg(msg_no * 16 + 3, (shift_right(identifier,13)and 255))
    set_can_reg(msg_no * 16 + 4, (shift_right(identifier,5)and 255))
    set_can_reg(msg_no * 16 + 5, (shift_left(identifier,3)and 255))
  endif
endsub


'***************************************************************
'''Enables a message object for transmitting messages.
'''+Disables all interrups of the message object.
'''msg_no: Number of the message object
'''Identifier: identifier for the frames which will be transmitted
'''extended: Determines whether the identifier is 11 or 29 Bit
'''+      0 = 11 Bit Identifier
'''+      1 = 29 Bit Identifier
'''*************************************************************
sub En_Transmit(msg_no,identifier,extended)
  REM Message Objekt freigeben
  set_can_reg(msg_no * 16, 149)
  if (extended = 0) then
    set_can_reg(msg_no * 16 + 6, 136)                         'write; DLC = 8 ; 11 Bit Identifier
    set_can_reg(msg_no * 16 + 2, (shift_right(identifier,3)and 255))  'Set Identifier
    set_can_reg(msg_no * 16 + 3, (shift_left(identifier,5)and 255))
  else
    set_can_reg(msg_no * 16 + 6, 140)                         'write; DLC = 8 ; 29 Bit Identifier
    set_can_reg(msg_no * 16 + 2, (shift_right(identifier,21)and 255)) 'Set Identifier
    set_can_reg(msg_no * 16 + 3, (shift_right(identifier,13)and 255))
    set_can_reg(msg_no * 16 + 4, (shift_right(identifier,5)and 255))
    set_can_reg(msg_no * 16 + 5, (shift_left(identifier,3)and 255))
  endif
endsub


'***************************************************************
'''Initialisation of the CAN-controller.
'''+All message objects will be disabled.
'''+Bitrat: 1MBit/s
'''+Acceptance filter : Must match
'''*************************************************************
sub Init_CAN()
  dim count as Long

  poke(rescan,1)          ' Reset the CAN-Controllers (Hardwarereset)
  sleep(10000)            'Resetpin muﬂ mindestens 1 ms low sein
  poke(rescan,0)

  set_can_reg(0,65)       'Access to the registers of the controller only for CPU
  set_can_reg(2,65)     'CPU interface register
  set_can_reg(06h,0ffh) 'Setting all masks to "must match"
  set_can_reg(07h,0ffh)
  set_can_reg(08h,0ffh)
  set_can_reg(09h,0ffh)
  set_can_reg(0ah,0ffh)
  set_can_reg(0bh,0ffh)
  set_can_reg(0ch,0ffh)
  set_can_reg(0dh,0ffh)
  set_can_reg(0eh,0ffh)
  set_can_reg(0fh,0ffh)
  set_can_reg(01fh,0)     'Clock Out Register
  set_can_reg(02fh,0)     'Bus configuration register
  set_can_reg(03fh,0)     'Bit Timing register 0
  set_can_reg(04fh,32)    ' Bit Timing register 1

  for count = 1 to 15       'Disable all message objects
    set_can_reg(count * 16,85)
    set_can_reg(count * 16 + 1,85)
  next count

  set_can_reg(0,0)          'Access to the registers of the controller only for Application
endsub


'***************************************************************
'Function only for internal use.
'***************************************************************
function round_intern(value) as Long ''<HIDE>
  dim buffer as Float

  buffer = value + 0.5
  round_intern = buffer
endfunction


'***************************************************************
'Function only for internal use.
'***************************************************************
function int_check(fl,int) as Long ''<HIDE>
  dim fbuffer as Float

  fbuffer = fl-int
  if ((fbuffer >= 0.00001) or (fbuffer <= -0.00001)) then
    int_check = 1       'not equal
  else
    int_check = 0       'equal
  endif
endfunction


'***************************************************************
'''Function is setting the baudrate of the CAN-Bus.
'''+    return value 
'''+    0 = baudrate is set
'''+    1 = the selected baudrate is not set
'''baud: CAN-Bus baudrate
'''*************************************************************
function Set_CAN_Baudrate(baud) as Long
  dim scale as Long
  dim fbuffer as Float
  dim lbuffer as Long
  dim back as Long
  dim tseg1 as Long
  dim tseg2 as Long
  dim brp as Long
  dim btr0 as Long
  dim btr1 as Long

  set_can_baudrate = 0          'Initialisation
  fbuffer = 8000000 / baud
  scale = round_intern(fbuffer)       'total scale to achive the baudrate
  set_can_baudrate = int_check(fbuffer,scale) '1 if baudrate not possible
  if (set_can_baudrate = 0) then      'If baudrate possible then find settings
    brp = (scale - 1) / 25        'Prescaler

    do              'find the right brp
      fbuffer = (scale / (brp + 1) - 3)
      lbuffer = round_intern(fbuffer)
      back = int_check(fbuffer,lbuffer)   'inc prescaler until tseg1 + tseg 2 is an Long
      inc(brp)
    until((back = 0) or (fbuffer<5))
    dec(brp)
    if (back <> 0) then set_can_baudrate = 1
  endif

  if (set_can_baudrate = 0) then      'If there is a correct prescaler
    tseg2 = lbuffer * 0.4
    tseg1 = lbuffer - tseg2
    if (tseg2 > 7) then
      do
        dec(tseg2)
        inc(tseg1)
      until(tseg2 = 7)
    endif
    btr0 = brp            'Synchronization jump width = 1 (bits = 0)
    btr1 = tseg1 + shift_left(tseg2,4)    'one sampling each Bit
    set_can_reg(0,65)       'enable access to btr regisers
    set_can_reg(3fh,btr0)   'set btr0 regisers
    set_can_reg(4fh,btr1)   'set btr1 regisers
    set_can_reg(0,0)        'disable access to btr regisers
  endif
endfunction



'---------------------------------------------------------------
' Aid sub for Read_Msg_Con and Read_Msg.
'---------------------------------------------------------------
Sub read_controller_registers(msg_no,id0,id1,id2,id3,id_lenght)''<HIDE>
  Dim counter As Long
  Dim buffer As Long

  buffer = (Get_CAN_Reg(msg_no * 16 + 6))
  CAN_msg[9] = Shift_Right((buffer And 0F0h),4)             ' Number of databytes
  id_lenght = buffer And 4
  id0 = Get_CAN_Reg(msg_no * 16 + 2)         ' identifier registers
  id1 = Get_CAN_Reg(msg_no * 16 + 3)
  If (id_lenght = 4) Then
    id2 = Get_CAN_Reg(msg_no * 16 + 4)
    id3 = Get_CAN_Reg(msg_no * 16 + 5)
  EndIf
  For counter = 1 To CAN_msg[9]                             ' reading databytes
    CAN_msg[counter] = Get_CAN_Reg((msg_no * 16)  + counter + 6)
  Next counter
EndSub


'***************************************************************
'''The function checks, if there is a new frame in the selected
'''+message object. If there is a new frame all databytes are read
'''+and then written to the array "can_msg".
'''+The number of dadabytes is written to the element 9 of the array "can_msg".
'''+The return value of the function is the identifier of the frame.
'''+If there is no new frame in the message box the return value is -1.
'''msg_no: number of the message object
'''*************************************************************
function Read_Msg(msg_no) as Long
  dim new as Long                                           ' new message in the message object
  dim id0,id1,id2,id3 as Long                               ' Identifier Byte 0 und 1
  dim count as Long
  dim buffer as Long
  dim id_lenght as Long                                     ' buffer to distinguish 11 and 29 bit identifier

  new = get_can_reg(msg_no * 16 + 1) and 3                 'New Data Bit (check for new message)
  if (new =2) then                                             ' New frame is arrived
    read_controller_registers(msg_no,id0,id1,id2,id3,id_lenght)
    set_can_reg(msg_no * 16,0fdh)                     ' Reset interupt
    set_can_reg(msg_no * 16 +1,85)                    ' Reset the new data flag
    if (id_lenght = 0) then
      read_msg = (shift_right(id1,5)) or (shift_left(id0,3))        ' return value = 11 bit identifier
    else
      read_msg = (shift_right(id3,3)) or (shift_left(id2,5)) or (shift_left(id1,13))  or (shift_left(id0,21))       ' return value = 29 bit identifier
    endif
  else
    read_msg = -1                                                  ' return value = -1
  endif
endfunction


'***************************************************************
'''+The function checks, if there is a new frame in the selected
'''+message object. If there is a new frame all databytes are read
'''+and then written to the array "can_msg".
'''+The number of dadabytes is written to the element 9 of the array "can_msg".
'''+The return value of the function is the identifier of the frame.
'''+If there is no new frame in the message box the return value is -1.
'''+Checks the consistens of the data and if they ars not consistent
'''+the datas will be reread.
'''msg_no: number of the message object
'''***************************************************************
function Read_Msg_Con(msg_no) as Long
  dim new as Long                                           ' new message in the message object
  dim id0,id1,id2,id3 as Long                               ' Identifier Byte 0 und 1
  dim count as Long
  dim buffer as Long
  dim id_lenght as Long                                     ' buffer to distinguish 11 and 29 bit identifier

  new = get_can_reg(msg_no * 16 + 1) and 3                 'New Data Bit (check for new message)
  if (new =2) then                                             ' New frame is arrived
    If (msg_no = 15) then
      read_controller_registers( msg_no, id0, id1, id2, id3, id_lenght)
      Set_CAN_Reg(msg_no * 16, 0FDh)                                            ' reset interrupt
      Set_CAN_Reg(msg_no * 16 + 1, 85)                                          ' reset the new data flag
      new = Get_CAN_Reg(msg_no * 16 + 1) And 3                                  ' Check, if a new frame in arrived
      If (new = 2) Then                                                                             ' New frame is arrived
        read_controller_registers(msg_no, id0, id1, id2, id3, id_lenght)
        Set_CAN_Reg(msg_no * 16, 0FDh)                                          ' reset interrupt
        Set_CAN_Reg(msg_no * 16 + 1, 85)                                        ' reset the new data flag
      endif
    else
    
      set_can_reg(msg_no * 16 +1,85)                    ' Reset the new data flag
      read_controller_registers(msg_no,id0,id1,id2,id3,id_lenght)
      new = get_can_reg(msg_no * 16 + 1) and 3                 'New Data Bit (check for new message)
      if (new =2) then                                             ' New frame is arrived
        read_controller_registers(msg_no,id0,id1,id2,id3,id_lenght)
        set_can_reg(msg_no * 16,0fdh)                     ' Reset interupt
      endif
      set_can_reg(msg_no * 16 +1,85)                    ' Reset the new data flag
      
    endif
      
    if (id_lenght = 0) then
      read_msg_Con = (shift_right(id1,5)) or (shift_left(id0,3))        ' return value = 11 bit identifier
    else
      read_msg_Con = (shift_right(id3,3)) or (shift_left(id2,5)) or (shift_left(id1,13))  or (shift_left(id0,21))       ' return value = 29 bit identifier
    endif
  else
    read_msg_Con = -1                                                  ' return value = -1
  endif
endfunction


'***************************************************************
'''Transmitting a frame.
'''+The databytes and the number of databytes have to be written in the
'''+array "can_msg" before this subroutine is activated.
'''+After this subroutine the frame is transmitted as soon as the
'''+CAN-bus allows it.
'''msg_no: number of the message object
'''*************************************************************
SUB Transmit(msg_no)
  dim buffer as Long
  dim count as Long

  for count = 1 to can_msg[9]
    set_can_reg((msg_no * 10h)+count+6,can_msg[count]) 'copy datas
  next count
  REM transfer nomber of data bytes
  buffer = get_can_reg((msg_no * 10h)+6)         'read register
  buffer = buffer and 0fh                         ' change DLC Bits
  buffer = buffer or shift_left(can_msg[9],4)
  set_can_reg(msg_no *16 +6,buffer)           ' write register

  set_can_reg(msg_no * 16 +1, 102)' Control register (setting bit to transmit data)
ENDSUB

'******************************************************************************
''' Checking if a Messagebox is ready to send. (The last Message is
'''+ alrady send)
'''+ return: 
'''+      Bit 0 = ready to send
'''+      Bit 1 = not ready to send
''' msg_no: number of the message object
'''****************************************************************************
Function Transmit_Status(msg_no) as long
  Transmit_Status = shift_right(Get_CAN_Reg((msg_no * 10h) + 1), 5) and 1    ' Control 1register
EndFunction




'***************************************************************
'''Enable the interrupt of a message object.
'''msg_no: number of the message object
'''*************************************************************
SUB En_Interrupt(msg_no)
  set_can_reg(msg_no * 16, 0ebh)
  set_can_reg(0, 02h)
ENDSUB


'***************************************************************
'''Setting a register of the CAN-controller.
'''regno: register which should be set
'''value: value which should be written to the register.
'''*************************************************************
SUB Set_CAN_Reg(regno, value)
#IF PRIORITAET = LOW            'It is not allowed to interrupt the process while accessing an ADwin-Pro module. This part disables the interrupt.
  BEGIN_CRITICAL_SECTION
#ENDIF
  poke(enadrwr,regno)
  poke(rdcan,value)
#IF PRIORITAET = LOW
  END_CRITICAL_SECTION
#ENDIF
ENDSUB


'***************************************************************
'''Reading a register of the CAN-controller.
'''+return value: value of the register
'''regno: register to read
'''*************************************************************
Function Get_CAN_Reg(regno) as Long
#IF PRIORITAET = LOW            'It is not allowed to interrupt the process while accessing an ADwin-Pro module. This part disables the interrupt.
  BEGIN_CRITICAL_SECTION
#ENDIF
  poke(enadrrd, regno)
  sleep(6)
  get_can_reg = peek(rdcan)
#IF PRIORITAET = LOW
  END_CRITICAL_SECTION
#ENDIF
endfunction
'***************************************************************


'**************************************************************
'''read the input data in seq mode
'''+only L16 Rev B or higher
'''channel: Channel no. (1, 3, ., 15).
'''************************************************************
FUNCTION SEQ_Read(channel) as Long
  SEQ_Read=peek(20400800H+(channel)/2)
ENDFUNCTION


'**************************************************************
'''initialising the ain mode
'''+only L16 Rev B or higher
'''mode:
'''+    0 normal Mode
'''+    1 single shot mode
'''+    2 continuous mode Eventdelay
'''+    3 continuous mode max Speed
''' gain:
'''+      0 range -10 V -> +10 V
'''+      1 range -5V   -> + 5 V
'''+      2 range -2.5V -> + 2.5 V
'''+      3 range -1.25 -> + 1.25 V
'''channel:selected channels in binary
'''muxtime: 
'''+Number of time units, which sets the settling time of the sequential control:
'''+    0: Standard waiting time (200=> 5 µs).
'''+   >0: Waiting time in units of 25 ns.
'''************************************************************
SUB SEQ_Init(mode,gain,channel,muxtime)
  dim channeli as Long
  dim i as Long
  dim anz as Long
  dim delay as Long
  dim imux as Long
  dim startdelay as Long
  dim w as long
  dim t as long
	
  POKE(20400810H,1)
  channeli=channel AND 5555H
  anz=0
  for i=0 to 15
    if (shift_left(1,i) AND channeli>0) then inc anz
  next
  
  t=0
  do
    w=peek(20400020H) AND 1
    inc t
  until	((w=0) or (t>40000))
	
  POKE(20400810H,0)
  SET_MUX(shift_left(gain,6))

#IF PRIORITAET = LOW
  delay=globaldelay*4000
#else
  delay=globaldelay
#endif

  poke(20400814H,channel)
  if (muxtime<>0) then
    poke(20400818H,muxtime)
    imux=muxtime
  else
    poke(20400818H,200)
    imux=200
  endif

  startdelay=delay-((anz*imux)+200)
  if (startdelay<0) then startdelay=0

  if (mode=3) then
    POKE(20400810H,mode-1)
    poke(2040081CH,0)              'Start delay
    poke(20400820H,0)
  else
    POKE(20400810H,mode)
    poke(2040081CH,startdelay)     'Start delay
    poke(20400820H,delay)
  endif
ENDSUB


'**************************************************************
''' read SSI encoder value 32 Bit (new counter board)
''' channel: number of the counter
'''************************************************************
FUNCTION SSI_Read(channel) AS Long
  SSI_READ = PEEK(204005afh + channel)      'read 32 Bit-value
ENDFUNCTION


'**************************************************************
''' sets the number of databits for the SSI Interface
''' counter: number of the counter
''' value: number of bits
'''************************************************************
SUB SSI_Set_Bits(channel,value)
  poke((204005A7h + channel) , value)
ENDSUB


'**************************************************************
''' sets frequency of the SSI clock
''' channel: number of the counter
''' value: frequency = 10MHZ / value
'''************************************************************
SUB SSI_Set_Clock(channel,value)
  poke((204005A3h + channel), value)
ENDSUB


'**************************************************************
''' Starts the SSI interface
''' value: Bitcombination each bit corresponts to one counter
'''************************************************************
SUB SSI_Start(pattern)
  poke(204005ACh,pattern)
ENDSUB


'*************************************************************
''' Sets the SSI mode
''' value: Bitcombination each bit corresponts to one counter
'''+         0 = single start mode
'''+         1 = continuous start mode
'''***********************************************************
SUB SSI_Mode(value)
  poke(204005b4h,value)
ENDSUB


'*************************************************************
''' Reads the SSI status
'''+ return value : 
'''+             0 = SSI value compleated
'''+             1 = SSI transfer running
'''counter: number of the SSI counter
'''***********************************************************
FUNCTION SSI_Status(counter) AS Long
  dim value2 as Long

  value2 = peek(204005ACh)
  ssi_status=shift_right(value2,(counter-1)) AND 01H
ENDFUNCTION


'*************************************************************
'''select L16 mode
'''+only L16 Rev B or higher
''' mode:
'''+       0   : classic mode (wait_eoc after 5 µs) default
'''+       1   : 500 KHz mode (wait_eoc after 2 µs)
'''***********************************************************
SUB L16_Mode(mode)
  dim old as Long

  old=peek(20400008H) AND 0FFDH
  old=old or shift_left((mode AND 1),1)
  poke(20400008H,old)
ENDSUB

'-------------------LS-BUS


'***************************************************************
''' the function reset the watchdog timer, set digital outputs and
'''+ read digital inputs.
'''+ return  : state of digital outputs and inputs
'''+ ls_module: device nr lsbus (FIX 1)
''' value: each bit corresponds to a channel
'''*************************************************************
function LS_Dig_IO(value) as Long
  dim ld[4] as Long at dm_local
  dim cnt_i as Long

  for cnt_i=1 to 4               ' rx data read
    ld[cnt_i]=peek(20400880H)
  next cnt_i
  LS_DIG_IO = ld[1] or (shift_left(ld[2],8)) or (shift_left(ld[3],16)) or (shift_left(ld[4],24))

  ld[1] = 1                      ' WD trigger
  I_SEND_BC(44h, 1, ld)

  ld[1] = value and 0FFh        ' tx data transmit
  ld[2] = shift_right(value,8) and 0FFh
  ld[3] = shift_right(value,16) and 0FFh
  ld[4] = shift_right(value,24) and 0FFh
  I_SEND_NA(1, 7Ch, 4, ld)      ' channel 1..32

  LS_READ_REQ(1, 7Ch, 4)        'request new in data
endfunction


'***************************************************************
''' the function enable,disable and config the watchdog
'''+ return  : error code (0->ok)
''' ls_module: device nr lsbus (1..15)
''' enable:  
'''+       0->disable
'''+       1->enable
''' time: trigger time [ms]
'''*************************************************************
function LS_Watchdog_Init(ls_module, enable, time) as Long
  dim ld[4] as Long at dm_local
  dim ret, t as Long

  'Watchdog disable
  ld[1] = 0ADh
  ld[2] = 0
  ret = I_SEND(ls_module, 45h, 2, ld)

  'Watchdog time
  t = time / 25e-6
  ld[1] = t and 0FFh
  ld[2] = shift_right(t,8) and 0FFh
  ld[3] = shift_right(t,16) and 0FFh
  ld[4] = shift_right(t,24) and 0FFh
  ret = ret or (I_SEND(ls_module, 48h, 4, ld))

  ' Watchdog enable
  ld[1] = 0ADh
  ld[2] = enable and 1
  LS_WATCHDOG_INIT = ret or (I_SEND(ls_module, 45h, 2, ld))
endfunction


'***************************************************************
''' reset the watchdog timer
'''+ If the watchdog is activated, the user program has to call
'''+ this function at least every ....
'''*************************************************************
sub LS_Watchdog_Reset()
  dim ld[1] as Long at dm_local

  ld[1] = 1
  I_SEND_BC(44h, 1, ld)
endsub


'******************************************************************************
''' Initialisation of a digital I/O module (channel 0-31)
'''+ return: error code (0->ok)
''' ls_module: device nr LS-bus (1..15)
''' value: each bit corresponds to 8 channel
'''+           0 -> input
'''+           1 -> output
'''+           bit 0: channels 1..8
'''+               1: channels 9..16
'''+               2: channels 17..24
'''+               3: channels 25..32
'''****************************************************************************
function LS_DigProg(ls_module, value) as Long
  dim ld[2] as Long at dm_local
  dim ret as Long

  ld[1] = shift_left(value,3) and 8h
  ld[1] = ld[1] or (shift_left(value,7) and 80h)
  ld[1] = ld[1] or 11h                     ' as after reset
  ld[2] = shift_left(value,2) and 8h
  ld[2] = ld[2] or (shift_left(value,6) and 80h)
  ld[2] = ld[2] or 11h
  ret = I_SEND(ls_module, 16h, 2, ld)

  ld[1] = shift_left(value,1) and 8h
  ld[1] = ld[1] or (shift_left(value,5) and 80h)
  ld[1] = ld[1] or 11h
  ld[2] = value and 8h
  ld[2] = ld[2] or (shift_left(value,4) and 80h)
  ld[2] = ld[2] or 11h
  LS_DIGPROG = ret or (I_SEND(ls_module, 20h+16h, 2, ld))
endfunction


'***************************************************************
'''Sets the digital outputs (channel 0 - 31)
''' ls_module: device nr LS-bus (1..15)
''' value: each bit corresponds to a channel
'''*************************************************************
sub LS_Digout_Long(ls_module,value)
  dim ld[4] as Long at dm_local
  ld[1] = value and 0FFh
  ld[2] = shift_right(value,8) and 0FFh
  ld[3] = shift_right(value,16) and 0FFh
  ld[4] = shift_right(value,24) and 0FFh
  I_SEND_NA(ls_module, 7Ch, 4, ld)      ' channel 1..32
endsub


'***************************************************************
''' Reads the digital inputs (channel 0 - 31)
'''+ return: state of digital outputs and inputs
''' ls_module: device nr LS-bus (1..15)
'''*************************************************************
function LS_Digin_Long(ls_module) as Long
  dim ld[4] as Long at dm_local
  dim ret as Long

  ret = I_RECEIVE(ls_module, 7Ch, 4, ld)
  LS_DIGIN_Long = ld[1] or (shift_left(ld[2],8)) or (shift_left(ld[3],16)) or (shift_left(ld[4],24))
endfunction


'***************************************************************
''' Reads the digital output status (overcurrent) (channel 0 - 31)
'''+ Return: each bit corresponds a channel 0=OK 1=Error
''' ls_module: device nr lsbus
'''*************************************************************
function LS_Get_Output_Status(ls_module) as Long
  dim ld[4] as Long at dm_local
  dim ret,t as Long

  t = I_RECEIVE(ls_module, 50h, 4, ld)
  ret=ld[1] or shift_left(ld[2],8) or shift_left(ld[3],16) or shift_left(ld[4],24)
  LS_GET_OUTPUT_STATUS = ret
endfunction


'***************************************************************
''' Initialize the LS-Modul
'''+ return: error code (0->ok)
''' ls_module: device nr LS-bus (1..15)
'''*************************************************************
function LS_DIO_Init(ls_module) as Long
  dim ld[4] as Long at dm_local
  dim ret as Long

  poke(2040089CH, 1)                            ' autoreset rx ON
  ld[1] = 3
  ret = I_SEND(ls_module, 40h, 1, ld)             ' Reset JX 1,2
  ld[1] = 1
  ret = ret or (I_SEND(ls_module, 5Ch, 1, ld))    ' timeout reset

  ld[1] = 4
  ret = ret or (I_SEND(ls_module, 4Ch, 1, ld))    ' Set default direction time

  ld[1]=0
  ld[2]=0
  ld[3]=0
  ld[4]=0
  LS_DIO_Init = ret or (I_SEND(ls_module, 50h, 4, ld)) ' overcurrent reset
  ld[1]=30h
  LS_DIO_Init = ret or (I_SEND(ls_module, 43h, 1, ld)) ' overtemperature reset
endfunction

