Attribute VB_Name = "ADWIN"
Option Explicit
Rem collection of subroutines for the control of the ADwin measurement data acquisition
Rem systems from VISUAL-BASIC 4.0
Rem Version 1.4 M.H. 27.08.96
Rem Version 2.0 O.L. 26.04.99     Update for VB 5.0 and 6.0
Rem Version 2.1 O.L. 10.05.01     Compatibility to other drivers
Rem Version 2.11 O.L. 14.11.01    inverting bit for messages
Rem Version 2.12 O.L. 13.05.02    GetData_String changed and improved
Rem Version 2.13 O.L. 20.01.04    Changing Return Value of funcion get_globaldelay from Integer to long
Rem Version 2.14 J.K. 29.07.04    Get_Dev_ID()
Rem Version 2.15 J.K. 24.02.05    Free_Mem calls now AD_Memory_all_byte
Rem Version 2.16 O.L. 09.03.05    New Functions: Get_Processdelay; Set_Processdelay
Rem Version 2.17 O.L. 27.09.05    GetData_UPacked_Long added

Rem 1. Function of the DLL ADwin32(32-bit, respectively)
Rem =====================================================================
Rem Attention! the files "ADWIN32.DLL" respectively have to be included in a
Rem subdirectory that can be accessed by Windows ( at best in " \windows")
Rem in order to use these functions the program "ADwin9.btl" (or ADwin4.btl or ADwin5.btl or ADwin8.btl
Rem respectively) has to be loaded to
Rem the ADwin system!

Rem The DLL ADwin32.dll respectively include all functions necessary for
Rem the data transfer between PC and the ADwin system


Rem Import functions from adwin32.dll
Declare Function Get_ADBPar_All Lib "adwin32.dll" (ByVal Start As Integer, ByVal Count As Integer, arr As Any, ByVal DeviceNo As Integer) As Integer
Declare Function Get_ADBFPar_All Lib "adwin32.dll" (ByVal Start As Integer, ByVal Count As Integer, arr As Any, ByVal DeviceNo As Integer) As Integer
Declare Function ADGetErrorCode Lib "adwin32.dll" () As Long
Declare Function ADGetErrorText Lib "adwin32.dll" (ByVal ErrorCode As Long, ByVal text As String, ByVal lenght As Long) As Long

Declare Function ADSetLanguage Lib "adwin32.dll" (ByVal Language As Long) As Long
Declare Function ADboot Lib "adwin32.dll" (ByVal filename As String, ByVal DeviceNo As Integer, ByVal iboardsize As Long, ByVal msg As Integer) As Long
Declare Function ADBload96 Lib "adwin32.dll" Alias "ADBload" (ByVal filename As String, ByVal DeviceNo As Integer, ByVal msg As Integer) As Integer
Declare Function ADTest_Version Lib "adwin32.dll" (ByVal DeviceNo As Integer, ByVal msg As Integer) As Integer
Declare Function ADProzessorTyp Lib "adwin32.dll" (ByVal DeviceNo As Integer) As Integer
Declare Function AD_Auslastung Lib "adwin32.dll" (ByVal DeviceNo As Integer) As Integer
Declare Function AD_Workload Lib "adwin32.dll" (ByVal Priority As Integer, ByVal DeviceNo As Integer) As Integer
Declare Function AD_Memory Lib "adwin32.dll" (ByVal DeviceNo As Integer) As Long
Declare Function AD_Memory_all Lib "adwin32.dll" (ByVal typ As Integer, ByVal DeviceNo As Integer) As Long
Declare Function AD_Memory_all_byte Lib "adwin32.dll" (ByVal typ As Integer, ByVal DeviceNo As Integer) As Long
Declare Function Get_ADC Lib "adwin32.dll" (ByVal NADC As Integer, ByVal DeviceNo As Integer) As Long
Declare Function Set_DAC96 Lib "adwin32.dll" Alias "Set_DAC" (ByVal ndac As Integer, ByVal value As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Set_Digout_ Lib "adwin32.dll" Alias "Set_Digout" (ByVal value As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Set_DigoutX Lib "adwin32.dll" (ByVal iw As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Get_Digout_ Lib "adwin32.dll" Alias "Get_Digout" (ByVal DeviceNo As Integer) As Long
Declare Function Get_DigoutX Lib "adwin32.dll" (ByVal DeviceNo As Integer) As Long
Declare Function Get_Digin_ Lib "adwin32.dll" Alias "Get_Digin" (ByVal DeviceNo As Integer) As Long
Declare Function Get_DiginX Lib "adwin32.dll" (ByVal DeviceNo As Integer) As Long

Declare Function ADB_Start Lib "adwin32.dll" (ByVal np As Integer, ByVal DeviceNo As Integer) As Integer
Declare Function ADB_Stop Lib "adwin32.dll" (ByVal np As Integer, ByVal DeviceNo As Integer) As Integer

Declare Function Clear_process_ Lib "adwin32.dll" Alias "Clear_Process" (ByVal np As Integer, ByVal DeviceNo As Integer) As Integer

Declare Function Set_ADBPar Lib "adwin32.dll" (ByVal np As Integer, ByVal value As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Set_ADBFPar Lib "adwin32.dll" (ByVal np As Integer, ByVal value As Single, ByVal DeviceNo As Integer) As Integer
Declare Function Get_ADBPar Lib "adwin32.dll" (ByVal np As Integer, ByVal DeviceNo As Integer) As Long
Declare Function Get_ADBFPar Lib "adwin32.dll" (ByVal np As Integer, ByVal DeviceNo As Integer) As Single

Declare Function Get_List Lib "adwin32.dll" (arr As Any, ByVal typ As Integer, ByVal nr As Integer, ByVal proz As Integer, ByVal anzahl As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Set_List Lib "adwin32.dll" (arr As Any, ByVal typ As Integer, ByVal nr As Integer, ByVal Count As Long, ByVal DeviceNo As Integer) As Integer

Declare Function Get_Data_String Lib "adwin32.dll" (ByVal arr As String, ByVal Count As Long, ByVal nr As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Set_Data_String Lib "adwin32.dll" (ByVal arr As String, ByVal nr As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Get_Data96 Lib "adwin32.dll" Alias "Get_Data" (ByRef arr As Any, ByVal typ As Integer, ByVal nr As Integer, ByVal Start As Long, ByVal Count As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Set_Data96 Lib "adwin32.dll" Alias "Set_Data" (ByRef arr As Any, ByVal typ As Integer, ByVal nr As Integer, ByVal Start As Long, ByVal Count As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Get_Fifo96 Lib "adwin32.dll" Alias "Get_Fifo" (ByRef arr As Any, ByVal typ As Integer, ByVal nr As Integer, ByVal Count As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Set_Fifo96 Lib "adwin32.dll" Alias "Set_Fifo" (ByRef arr As Any, ByVal typ As Integer, ByVal nr As Integer, ByVal Count As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Get_Data_String_Length Lib "adwin32.dll" (ByVal nr As Integer, ByVal DeviceNo As Integer) As Long
Declare Function Get_Data_Length Lib "adwin32.dll" (ByVal nr As Integer, ByVal DeviceNo As Integer) As Long
Declare Function get_fifo_count96 Lib "adwin32.dll" Alias "Get_Fifo_Count" (ByVal np As Integer, ByVal DeviceNo As Integer) As Long
Declare Function get_fifo_empty96 Lib "adwin32.dll" Alias "Get_Fifo_Empty" (ByVal np As Integer, ByVal DeviceNo As Integer) As Long
Declare Function clear_fifo96 Lib "adwin32.dll" Alias "Clear_Fifo" (ByVal np As Integer, ByVal DeviceNo As Integer) As Integer

Declare Function SaveFast Lib "adwin32.dll" (ByVal Dateiname As String, ByVal datanr As Integer, ByVal startind As Long, ByVal Count As Long, ByVal Mode As Integer, ByVal DeviceNo As Integer) As Integer
Declare Function AD_Net_Connect Lib "adwin32.dll" (ByVal prot As String, ByVal addr As String, ByVal EPoint As String, ByVal Password As String, ByVal msg As Integer) As Integer
Declare Function AD_Net_Disconnect Lib "adwin32.dll" () As Integer

Declare Function Get_Data_packed Lib "adwin32.dll" (arr As Any, ByVal typ As Integer, ByVal nr As Integer, ByVal Start As Long, ByVal Count As Long, ByVal DeviceNo As Integer) As Integer
Declare Function Get_Data_Upacked Lib "adwin32.dll" (arr As Any, ByVal typ As Integer, ByVal nr As Integer, ByVal Start As Long, ByVal Count As Long, ByVal DeviceNo As Integer) As Integer

Declare Function Get_Btl_Directory Lib "adwin32.dll" (ByVal path As String, ByRef laenge As Long) As Integer
Declare Function Get_Dev_ID96 Lib "adwin32.dll" Alias "get_dev_id" (ByVal DeviceNo As Integer, ByRef arr As Any) As Long

Public DeviceNo As Integer
Public BinFolderPath As String
Public BootFileName As String
Public CurProcessFile As String
Global Const append As Integer = 1
Global Const overwrite As Integer = 0
Global message As Integer

Rem Checks if Process "PNo%" has executed the Activate_PC command
Function Activate(PNo%) As Long
Dim erg As Integer
Dim parno As Integer
  parno% = -(70 - PNo%)
  erg% = Get_ADBPar(parno%, DeviceNo%)
  Activate = erg%
  If (erg% = 1) Then Call Set_Par(parno%, 0)
End Function

Rem Returns state of ACTIVATE_PC-flag of process number "ProcessNo"
Function Activate_PC(ProcessNo As Integer) As Long
Dim parno As Integer
Dim erg As Long
  parno% = -(70 - ProcessNo%)
  erg& = Get_ADBPar(parno%, DeviceNo%)
  Activate_PC = erg&
  If (erg& = 1) Then Call Set_Par(parno%, 0)
End Function

Rem Loads an ADbasic-process (a bin file generated by ADbasic) to the ADwin board
Function ADBload(adbasicfile$) As Integer
  ADBload = ADBload96(adbasicfile$, DeviceNo%, message Xor 1)
End Function

Rem Returns measured value of anlog input channel
Function ADC(Channel As Integer) As Long
  ADC = Get_ADC(Channel%, DeviceNo%)
End Function

Public Function ADWIN_BootBoard(ByRef ADWINBoard As Board) As Boolean

    Dim ReturnVal As Long
    
    'Check for NOCOMM mode
    If NOCOMM_MODE = True Then
    
        ADWIN_BootBoard = False
        
        Exit Function
        
    End If
    
    If ADWINBoard.CommProtocol <> ADWIN_COM Then
    
        ADWIN_BootBoard = False
        
        Exit Function
        
    End If

    'First set the Device No
    ADWIN.SetDeviceNo ADWINBoard.BoardNum

    'Check to see if the ADWIN board has been booted
    On Error GoTo NoAdwinBoard:
    
        ReturnVal = ADWIN.Test_Version
        
    On Error GoTo 0
                               
    'If ReturnVal is Not zero, then need to boot the ADWIN board
    If ReturnVal <> 0 Then
    
        'Boot the ADWIN board
        ReturnVal = ADWIN.Boot(ADWIN.BinFolderPath & ADWIN.BootFileName, 0)
        
        'If Return Value is Not = 8000, then error occurred during the boot
        If ReturnVal <> 8000 Then
        
            MsgBox "Boot process failed for ADWIN-light-16 board." & _
                   vbNewLine & "ADWIN Device Number = " & Trim(str(ADWIN.GetDeviceNo)), _
                   vbCritical, _
                   "ADWIN Boot Error"
                      
             ADWIN_BootBoard = False
            
            'Leave the Function
            Exit Function
            
        End If
        
    End If

    ADWIN_BootBoard = True
    
    Exit Function
    
NoAdwinBoard:

    'Tell user that the ADWin board is missing
    MsgBox "Adwin Board has not been loaded and installed correctly." & vbNewLine & vbNewLine & _
           "Error Message:" & vbNewLine & _
           Err.Description, _
           vbCritical, _
           "Comm Setup Error!!"
           

    'Prompt User if they want to turn on No-comm mode
    modProg.Prompt_NOCOMM

End Function

Rem checks the momentary ADwin-System workload
Rem the return value is the workload in percent.'
Function auslast() As Integer
  auslast = AD_Auslastung(DeviceNo%)
End Function

Rem Down-loads the operating system (BTL-file) to the ADwin-System
Function Boot(filename As String, memsize As Long) As Long
  Boot = ADboot(filename$, DeviceNo%, memsize&, message Xor 1)
End Function

Rem clear the contents of FIFO
Sub clear_fifo(FNo%)
Dim erg As Integer
  erg% = clear_fifo96(FNo%, DeviceNo%)
End Sub

Rem Clear the Process
Function Clear_Process(ProcessNo As Long) As Long
  Clear_Process = Clear_process_(ProcessNo&, DeviceNo)
End Function

Function ClearAll_Processes() As Boolean

    Dim i As Integer
    Dim ReturnVal As Long
    Dim TempBool As Boolean
    
    TempBool = True
    
    For i = 1 To 10
    
        Do
        
            'Stop the process
            Stop_Process i
                   
            'Check process status
            ReturnVal = Process_Status(i)
            
        Loop Until ReturnVal <> 1
        
        'Process is stopped or doesn't exist
        'Clear the process
        ReturnVal = Clear_Process(CLng(i))
        
        If ReturnVal = 1 Then
        
            'Error occured
            TempBool = False
            
        End If
        
        TempBool = TempBool And True
        
    Next i

    ClearAll_Processes = TempBool

End Function

Rem Sets analog output "ndac" to value "iw" (in digits)
Sub DAC(Channel As Integer, value As Long)
  Call Set_DAC96(Channel%, value&, DeviceNo%)
End Sub

Rem Returns the lenght of a data
Function Data_Length(DataNo As Integer) As Long
  Data_Length = Get_Data_Length(DataNo%, DeviceNo%)
End Function

Rem Writes array-elements of ADwin-System immediately to harddisk
Function Data2File(filename As String, DataNo As Integer, StartIndex As Long, Count As Long, Mode As Integer) As Integer
  Data2File = SaveFast(filename$, DataNo%, StartIndex&, Count&, Mode%, DeviceNo%)
End Function

Rem digital input query
Rem the inputs are set in a 16-bit word'
Function Dig_In() As Long
  Dig_In = Get_Digin_(DeviceNo%)
End Function

Rem Returning the value of the digital outputs
Function Dig_Out() As Long
  Dig_Out = Get_Digout_(DeviceNo%)
End Function

Rem Error messages for Test, ADBload, Boot & Net_Connect    0/1 -> (off/on)
Sub Err_Message(OnOff As Integer)
  message = (OnOff And 1) Xor 1
End Sub

Rem Initiats read and write pointer of FIFO
Sub Fifo_Clear(FifoNo As Integer)
  Call clear_fifo96(FifoNo%, DeviceNo%)
End Sub

Rem Returns number of free elements in FIFO
Function Fifo_Empty(FifoNo As Integer) As Long
  Fifo_Empty = get_fifo_empty96(FifoNo%, DeviceNo%)
End Function

Rem Returns number of stored elements in FIFO
Function Fifo_Full(FifoNo As Integer) As Long
  Fifo_Full = get_fifo_count96(FifoNo%, DeviceNo%)
End Function

Rem Returns value of free memory from the ADwin-2, -4, -5 ,-8 and -9 System(in bytes)
Function Free_Mem(Mem_Spec As Long) As Long
    Select Case Mem_Spec
        Case 0
            Free_Mem = AD_Memory(DeviceNo) ' T2, T4, T5, T8
        Case Else
            ' 1 = PM   (for processors  T9, T10, T11 )
            ' 2 = EM   (for processor   T11 only)
            ' 3 = DM   (for processors  T9, T10, T11 )
            ' 4 = DX   (for processors  T9, T10, T11 )
            Free_Mem = AD_Memory_all_byte(Mem_Spec, DeviceNo)
    End Select
End Function

Rem checks the momentary free memory of the ADwin-System
Rem return value is the free memory in byte
Rem only for T4,T5 and T8
Function freemem() As Long
  freemem = AD_Memory(DeviceNo%)
End Function

Rem checks the momentary free memory of the ADwin-System
Rem return value is the free memory in byte
Rem only for T9
Rem typ : 1 = internal code
Rem       3 = internal data
Rem       4 = external data (DRAM)
Rem       5 = external code (SRAM)
Rem       6 = external data (SRAM)
Rem return value in longword
Function freemem_T9(typ%) As Long
  freemem_T9 = AD_Memory_all(typ%, DeviceNo%)
End Function

Rem Returns array-element(s) of WORD-type from ADwin-System-array (DATA_"DataNo")
Sub Get_Data(DataNo As Integer, StartIndex As Long, Count As Long, data() As Integer)
  Call Get_Data96(data%(1), 1, DataNo%, StartIndex&, Count&, DeviceNo%)
End Sub

Rem reads the float data set with the number "DNo%"
Rem "count&" elements (starting at the position "first&")
Sub get_data_float(DNo%, first&, Count&, Dest!())
Dim erg As Integer
  erg% = Get_Data96(Dest!(1), 5, DNo%, first&, Count&, DeviceNo%)
End Sub

Rem reads the long data set with the number "DNo%"
Rem "count&" elements (starting at the position "first&")
Sub get_data_long(DNo%, first&, Count&, Dest&())
Dim erg As Integer
  erg% = Get_Data96(Dest&(1), 2, DNo%, first&, Count&, DeviceNo%)
End Sub

Rem Undocumented function :
Rem Returns MAC-Address, IP-Address (or USB-Serial Number if USB device)
Function Get_Dev_ID(ID_Data() As Long) As Long
  ' ID_Data() : size must be 4 elements in minimum !

  Dim my_array(4) As Long
  Dim i As Long
  
  Get_Dev_ID = Get_Dev_ID96(DeviceNo%, my_array(1))
  ' return value :    0  = OK
  '                   1  = error
  
  For i = 1 To 4  ' use of a local array, to avoid "general protection  errors",
                  '   if the user passes an array, which is smaller than 4 longs.
     ID_Data(i) = my_array(i)
  Next i
  ' if ENET-device :
  '    ID_Data(1) = MAC-Adr HIGH-Word
  '    ID_Data(2) = MAC-Adr LOW-Word
  '    ID_Data(3) = IP-Address
  '    ID_Data(4) = reserve
  ' if USB-device :
  '    ID_Data(1) = USB-Serial Number
  '    ID_Data(2) = 0
  '    ID_Data(3) = 0
  '    ID_Data(4) = reserve
  
  ' Important :
  ' If an application program reads the MAC for identifying a special
  ' ADwin-system and will refuse operation , if the MAC does not match,
  ' then the following issue should be considered :
  ' What if an ADwin-system breaks and needs to be replaced ?
  ' Neither Jaeger Messtechnik nor Keithley Instruments will provide a
  ' spare part with an IDENTICAL  MAC address or USB- serialnumber !
  ' The application program has to be prepared somehow, that a replacement
  ' with an ADwin-system with a DIFFERENT MAC or USB-serialnumber might
  ' be necessary !
End Function

Rem Returns state of digital inputs (bits 0...15)
Function Get_Digin() As Long
  Get_Digin = Get_Digin_(DeviceNo%)
End Function

Rem Returns state (read-back) of digital outputs (bits 0...15)
Function Get_Digout() As Long
  Get_Digout = Get_Digout_(DeviceNo%)
End Function

Rem Sets FIFO-element(s) of ADwin-System from integer-array
Sub Get_Fifo(FifoNo As Integer, Count As Long, data() As Integer)
  Call Get_Fifo96(data%(1), 1, FifoNo%, Count&, DeviceNo%)
End Sub

Rem Getting the number of elements in the FIFO
Function get_fifo_count(FNo%) As Long
  get_fifo_count = get_fifo_count96(FNo%, DeviceNo%)
End Function

Rem Getting the number of empty positions in the FIFO
Function get_fifo_empty(FNo%) As Long
  get_fifo_empty = get_fifo_empty96(FNo%, DeviceNo%)
End Function

Rem Fetching the float-element(s) from a ADwin-System FIFO
Sub get_fifo_float(FNo%, Count&, Dest!())
Dim erg As Integer
  erg% = Get_Fifo96(Dest!(1), 5, FNo%, Count&, DeviceNo%)
End Sub

Rem Fetching the long-element(s) from a ADwin-System FIFO
Sub get_fifo_long(FNo%, Count&, Dest&())
Dim erg As Integer
  erg% = Get_Fifo96(Dest&(1), 2, FNo%, Count&, DeviceNo%)
End Sub

Rem Returns float value of parameter (FPAR_"Index") from the ADwin-System
Function Get_Fpar(Index As Integer) As Single
  Get_Fpar = Get_ADBFPar(Index%, DeviceNo%)
End Function

Rem Gets all 80 float parameters (FPar_1 - FPar_80) into a Single array
Function Get_Fpar_All(arr() As Single) As Integer
  Get_Fpar_All = Get_ADBFPar_All(1, 80, arr!(1), DeviceNo%)
End Function

Rem Gets all 80 float parameters (FPar_1 - FPar_80) into a double array
Function Get_Fpar_All_Double(arr() As Double) As Integer
Dim buffer(80) As Single
Dim lauf As Long
  Get_Fpar_All_Double = Get_ADBFPar_All(1, 80, buffer!(1), DeviceNo%)
    For lauf = 1 To 80
      arr(lauf) = buffer(lauf)
    Next lauf
End Function

Rem Gets a block of ADwin float parameters into a single array
Function Get_Fpar_Block(arr() As Single, StartIndex As Long, Count As Long) As Integer
  Get_Fpar_Block = Get_ADBFPar_All(StartIndex&, Count&, arr!(1), DeviceNo%)
End Function

Rem Gets a block of ADwin float parameters into a double array
Function Get_Fpar_Block_Double(arr() As Double, StartIndex As Long, Count As Long) As Integer
Dim buffer(80) As Single
Dim lauf As Long
  Get_Fpar_Block_Double = Get_ADBFPar_All(StartIndex&, Count&, buffer!(1), DeviceNo%)
  For lauf = 1 To Count
    arr(lauf) = buffer(lauf)
  Next lauf
End Function

Rem Gets the Globaldelay of Process "ProcessNo"
Function Get_Globaldelay(ProcessNo As Integer) As Long
  Get_Globaldelay = Get_Par(-90 + ProcessNo)
End Function

Rem Gets the error code of the last error occured
Function Get_Last_Error() As Long
  Get_Last_Error = ADGetErrorCode()
End Function

Rem Gets the error text of the "Last error"
Function Get_Last_Error_Text(Last_Error As Long) As String
Dim temp As String * 255
Dim merker As Long
Dim ret As Long
  merker = Len(temp)
  ret = ADGetErrorText(Last_Error, temp, merker)
  Get_Last_Error_Text = Mid(temp, 1, ret)
End Function

Rem Returns value of parameter (PAR_"Index") from the ADwin-System
Function Get_Par(Index As Integer) As Long
  Get_Par = Get_ADBPar(Index%, DeviceNo%)
End Function

Rem Gets all 80 integer parameters (Par_1 - Par_80) into a Long array
Function Get_Par_All(arr() As Long) As Integer
  Get_Par_All = Get_ADBPar_All(1, 80, arr&(1), DeviceNo%)
End Function

Rem Gets a block of ADwin integer parameters into a long array
Function Get_Par_Block(arr() As Long, StartIndex As Long, Count As Long) As Integer
  Get_Par_Block = Get_ADBPar_All(StartIndex&, Count&, arr&(1), DeviceNo%)
End Function

Rem Gets the Processdelay of Process "ProcessNo"
Function Get_Processdelay(ProcessNo As Integer) As Long
  Get_Processdelay = Get_Par(-90 + ProcessNo)
End Function

Rem Returns array-element(s) of Double-type from ADwin-System-array (DATA_"DataNo")
Sub GetData_Double(DataNo As Integer, StartIndex As Long, Count As Long, data() As Double)
  Call Get_Data96(data#(1), 6, DataNo%, StartIndex&, Count&, DeviceNo%)
End Sub

Rem Returns array-element(s) of Float-type from ADwin-System-array (DATA_"DataNo")
Sub GetData_Float(DataNo As Integer, StartIndex As Long, Count As Long, data() As Single)
  Call Get_Data96(data!(1), 5, DataNo%, StartIndex&, Count&, DeviceNo%)
End Sub

Rem Returns array-element(s) of LONGINT-type from ADwin-System-array (DATA_"DataNo")
Sub GetData_Long(DataNo As Integer, StartIndex As Long, Count As Long, data() As Long)
  Call Get_Data96(data&(1), 2, DataNo%, StartIndex&, Count&, DeviceNo%)
End Sub

Rem Undocumented function :
Rem Returns array-element(s) of DOUBLE-type from packed ADwin-System-array (DATA_"DataNo")
Sub GetData_Packed_Double(DataNo As Integer, StartIndex As Long, Count As Long, data() As Double)
  Call Get_Data_packed(data(1), 6, DataNo, StartIndex, Count, DeviceNo)
End Sub

Rem Undocumented function :
Rem Returns array-element(s) of FLOAT-type from packed ADwin-System-array (DATA_"DataNo")
Sub GetData_Packed_Float(DataNo As Integer, StartIndex As Long, Count As Long, data() As Single)
  Call Get_Data_packed(data(1), 5, DataNo, StartIndex, Count, DeviceNo)
End Sub

Rem Undocumented function :
Rem Returns array-element(s) of LONGINT-type from packed ADwin-System-array (DATA_"DataNo")
Sub GetData_Packed_Long(DataNo As Integer, StartIndex As Long, Count As Long, data() As Long)
  Call Get_Data_packed(data(1), 2, DataNo, StartIndex, Count, DeviceNo)
End Sub

Rem Undocumented function :
Rem Returns array-element(s) of WORD-type from packed ADwin-System-array (DATA_"DataNo")
Sub GetData_Packed_Short(DataNo As Integer, StartIndex As Long, Count As Long, data() As Integer)
  Call Get_Data_packed(data(1), 1, DataNo, StartIndex, Count, DeviceNo)
End Sub

Rem Returns array-element(s) of String-type from ADwin-System-array (DATA_"DataNo")
Function GetData_String(DataNo As Integer, MaxCount As Long, data As String) As Long
  data = String(MaxCount + 1, "0")
  GetData_String = Get_Data_String(data, MaxCount + 1, DataNo%, DeviceNo%)
  If GetData_String >= 0 Then
     data = Left(data, GetData_String)
  Else
     data = "" ' empty string, bec. an error happened
  End If
End Function

Rem Undocumented function :
Rem Returns array-element(s) of LONGINT-type from packed ADwin-System-array (DATA_"DataNo") / unsighned
Sub GetData_UPacked_Long(DataNo As Integer, StartIndex As Long, Count As Long, data() As Long)
  Call Get_Data_Upacked(data(1), 2, DataNo, StartIndex, Count, DeviceNo)
End Sub

Public Function GetDeviceNo() As Integer

    GetDeviceNo = DeviceNo
    
End Function

Rem Fetching the long-element(s) from a ADwin-System FIFO
Sub GetFifo(FNo As Integer, Count As Long, Dest() As Integer)
  Call Get_Fifo96(Dest%(1), 1, FNo%, Count&, DeviceNo%)
End Sub

Rem Sets FIFO-element(s) of ADwin-System from Double-array
Sub GetFifo_Double(FifoNo As Integer, Count As Long, data() As Double)
  Call Get_Fifo96(data#(1), 6, FifoNo%, Count&, DeviceNo%)
End Sub

Rem Sets FIFO-element(s) of ADwin-System from Single-array
Sub GetFifo_Float(FifoNo As Integer, Count As Long, data() As Single)
  Call Get_Fifo96(data!(1), 5, FifoNo%, Count&, DeviceNo%)
End Sub

Rem Sets FIFO-element(s) of ADwin-System from LONGINT-array
Sub GetFifo_Long(FifoNo As Integer, Count As Long, data() As Long)
  Call Get_Fifo96(data&(1), 2, FifoNo%, Count&, DeviceNo%)
End Sub

Rem Loads the driver to the ADwin board
Rem According to the ADwin board processor one of the following driver files is needed
Rem    T225 Processor -> btlfile$ = "adwin2.btl"
Rem    T400 Processor -> btlfile$ = "adwin4.btl"
Rem    T450 Processor -> btlfile$ = "adwin5.btl"
Rem    T805 Processor -> btlfile$ = "adwin8.btl"
Rem The memory size (size&) is indicated in hexadecimal code
Rem Examples: ADwin board with 1MB -> size& = 100000
Rem              "     "       4MB -> size& = 400000
Rem              "     "       8MB -> size& = 800000
Function iserver(btlfile$, size&) As Long
  iserver = ADboot(btlfile$, DeviceNo%, size&, message Xor 1)
End Function

Rem Down-loads an ADbasic-process (BIN-file) to the ADwin-System
Function Load_Process(filename As String) As Integer
  Load_Process = ADBload96(filename$, DeviceNo%, message Xor 1)
End Function

Rem  Gets access to the ADwin board in another server via network.
Rem  On this server the program ADserver has to run, which will be provided
Rem  together with ADbasic.

Function Net_Connect(Protocol As String, Address As String, Endpoint As String, Password As String) As Integer
  Net_Connect = AD_Net_Connect(Protocol$, Address$, Endpoint$, Password$, message Xor 1)
End Function

Rem  Disconnects network access to the ADwin-card
Rem  in another server

Function Net_Disconnect() As Integer
  Net_Disconnect = AD_Net_Disconnect()
End Function

Rem Returns the Processstatus
Function Process_Status(ProcessNo As Integer) As Integer
  Process_Status = Get_Par(-100 + ProcessNo)
End Function

Rem Returns the Processortype
Function Processor_Type() As Integer
Dim merker As Integer
  merker = ADProzessorTyp(DeviceNo)
  Processor_Type = merker
  If (merker = 146) Then Processor_Type = 5
  If (merker = 1000) Then Processor_Type = 9
End Function

Rem Saves the data from an ADbasic data set directly to hard disk in binary code
Rem under the specified file name filename$
Rem For mode% the following values are permitted:
Rem     0 -> a file with the same name is overwritten
Rem     1 -> data will be appended to a file with the same name
Function save_data(filename$, DataNo%, first&, Count&, Mode%) As Integer
  save_data = SaveFast(filename$, DataNo%, first&, Count&, Mode%, DeviceNo%)
End Function

Rem Setting of the DAC number "NDAC%" to the value "iw%"
Rem the measured value has to be in the lowest 12 bits of "iw%"
Sub set_dac(ndac%, iw&)
Dim erg As Integer
  erg% = Set_DAC96(ndac%, iw&, DeviceNo%)
End Sub

Rem Sets array-element(s) of ADwin-System from WORD-array
Sub Set_Data(DataNo As Integer, StartIndex As Long, Count As Long, data() As Integer)
  Call Set_Data96(data%(1), 1, DataNo%, StartIndex&, Count&, DeviceNo%)
End Sub

Rem "count&" elements of the Visual-Basic array "source()" are
Rem allocated to the ADbasic float data set
Rem with the number "DNo%" (starting at the position "first&")
Sub set_data_float(DNo%, first&, Count&, Source!())
Dim erg As Integer
  erg% = Set_Data96(Source!(1), 5, DNo%, first&, Count&, DeviceNo%)
End Sub

Rem "count&" elements of the Visual-Basic array "source()" are
Rem allocated to the ADbasic long data set
Rem with the number "DNo%" (starting at the position "first&")
Sub set_data_long(DNo%, first&, Count&, Source&())
Dim erg As Integer
  erg% = Set_Data96(Source&(1), 2, DNo%, first&, Count&, DeviceNo%)
End Sub

Rem outputs the value which is found in "iw%" on the digital outputs
Sub set_dig_out(iw&)
Dim erg As Integer
  erg% = Set_Digout_(iw&, DeviceNo%)
End Sub

Rem Sets digital outputs to "Value" (bits 0...15)
Sub Set_Digout(value As Long)
  Call Set_Digout_(value&, DeviceNo%)
End Sub

Public Sub Set_DigOutPort(ByVal PortNum As Long, ByVal value As Long)

    Dim Ou

End Sub

Rem allocates the array values to an ADbasic integer-FIFO
Sub set_fifo(FNo%, Count&, Source%())
Dim erg As Integer
  erg% = Set_Fifo96(Source%(1), 1, FNo%, Count&, DeviceNo%)
End Sub

Rem allocates the array values to an ADbasic float-FIFO
Sub set_fifo_float(FNo%, Count&, Source!())
Dim erg As Integer
  erg% = Set_Fifo96(Source!(1), 5, FNo%, Count&, DeviceNo%)
End Sub

Rem allocates the array values to an ADbasic long-FIFO
Sub set_fifo_long(FNo%, Count&, Source&())
Dim erg As Integer
  erg% = Set_Fifo96(Source&(1), 2, FNo%, Count&, DeviceNo%)
End Sub

Rem Sets a parameter (FPAR_"Index") on the ADwin-System to "Value"
Sub Set_Fpar(Index As Integer, value As Single)
  Call Set_ADBFPar(Index%, value!, DeviceNo%)
End Sub

Rem Set the Globaldelay of Process "ProcessNo" to "Globaldelay"
Sub Set_Globaldelay(ProcessNo As Integer, Globaldelay As Long)
  Call Set_Par((-90 + ProcessNo%), Globaldelay&)
End Sub

Rem Set the Language of the Errormessages
Function Set_Language(Language As Integer) As Integer
  Set_Language = ADSetLanguage(Language)
End Function

Rem Sets a parameter (PAR_"Index") on the ADwin-System to "Value"
Sub Set_Par(Index As Integer, value As Long)
  Call Set_ADBPar(Index%, value&, DeviceNo%)
End Sub

Rem Set the Processdelay of Process "ProcessNo" to "Processdelay"
Sub Set_Processdelay(ProcessNo As Integer, Processdelay As Long)
  Call Set_Par((-90 + ProcessNo%), Processdelay&)
End Sub

Rem Sets array-element(s) of ADwin-System from Double-array
Sub SetData_Double(DataNo As Integer, StartIndex As Long, Count As Long, data() As Double)
  Call Set_Data96(data#(1), 6, DataNo%, StartIndex&, Count&, DeviceNo%)
End Sub

Rem Sets array-element(s) of ADwin-System from Single-array
Sub SetData_Float(DataNo As Integer, StartIndex As Long, Count As Long, data() As Single)
  Call Set_Data96(data!(1), 5, DataNo%, StartIndex&, Count&, DeviceNo%)
End Sub

Rem Sets array-element(s) of ADwin-System from LONGINT-array
Sub SetData_Long(DataNo As Integer, StartIndex As Long, Count As Long, data() As Long)
  Call Set_Data96(data&(1), 2, DataNo%, StartIndex&, Count&, DeviceNo%)
End Sub

Sub SetData_String(DataNo As Integer, data As String)
  Call Set_Data_String(data, DataNo%, DeviceNo%)
End Sub

Public Sub SetDeviceNo(ByVal DevNumber As Integer)

    DeviceNo = DevNumber

End Sub

Rem Sets FIFO-element(s) of ADwin-System from Single-array
Sub SetFifo_Double(FifoNo As Integer, Count As Long, data() As Double)
  Call Set_Fifo96(data#(1), 6, FifoNo%, Count&, DeviceNo%)
End Sub

Rem Sets FIFO-element(s) of ADwin-System from Single-array
Sub SetFifo_Float(FifoNo As Integer, Count As Long, data() As Single)
  Call Set_Fifo96(data!(1), 5, FifoNo%, Count&, DeviceNo%)
End Sub

Rem Sets FIFO-element(s) of ADwin-System from LONGINT-array
Sub SetFifo_Long(FifoNo As Integer, Count As Long, data() As Long)
  Call Set_Fifo96(data&(1), 2, FifoNo%, Count&, DeviceNo%)
End Sub

Rem Error messages for Test, ADBload, Boot & Net_Connect    0/1 -> (off/on)
Sub Show_Errors(OnOff As Integer)
  message = (OnOff And 1) Xor 1
End Sub

Rem Starts process number "ProcessNo" on the ADwin-System
Sub Start_Process(ProcessNo As Integer)
  Call ADB_Start(ProcessNo%, DeviceNo%)
End Sub

Rem starts the ADbasic process with the number "PNo%"
Sub start_proz(PNo%)
Dim erg As Integer
  erg% = ADB_Start(PNo%, DeviceNo%)
End Sub

Rem Stops process number "ProcessNo" on the ADwin-System
Sub Stop_Process(ProcessNo As Integer)
  Call ADB_Stop(ProcessNo%, DeviceNo%)
End Sub

Rem stops the ADbasic process with the number "PNo%"
Sub Stop_Proz(PNo%)
Dim erg As Integer
  erg% = ADB_Stop(PNo%, DeviceNo%)
End Sub

Rem Returns the lenght of a string
Function String_Length(DataNo As Integer) As Long
  String_Length = Get_Data_String_Length(DataNo%, DeviceNo%)
End Function

Rem checks if the right driver has been
Rem loaded to the ADwin board
Function test() As Integer
  test = ADTest_Version(DeviceNo, message Xor 1)
End Function

Rem Checks if the CPU is accessible (previously booted and running)
Function Test_Version() As Integer
   Test_Version = ADTest_Version(DeviceNo, message Xor 1)
End Function

Rem Returns how busy the CPU is (processor usage in percent)
Function Workload(Priority As Long) As Integer
  Workload = AD_Workload(Priority&, DeviceNo%)
End Function

