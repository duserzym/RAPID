Attribute VB_Name = "modFileSave"
Option Explicit

Private Const BIF_STATUSTEXT = &H4&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As BrowseInfo) As Long
    
Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
     ByVal lpBuffer As String) As Long
     
Private Declare Function lstrcat Lib "kernel32" _
    Alias "lstrcatA" _
    (ByVal lpString1 As String, _
     ByVal lpString2 As String) As Long

Private Type BrowseInfo
 hWndOwner      As Long
 pIDLRoot       As Long
 pszDisplayName As Long
 lpszTitle      As Long
 ulFlags        As Long
 lpfnCallback   As Long
 lParam         As Long
 iImage         As Long
End Type

Private m_CurrentDirectory As String   'The current directory

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private m_dlgPosition As Integer
Private m_dlgHideOK As Integer
Private m_OKCaption As String
Private m_CancelCaption As String
Private m_LookInCaption As String
Private m_FileNameCaption As String
Private m_FileOfTypeCaption As String
          
Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_CREATEPROMPT As Long = &H2000
Private Const OFN_ENABLEHOOK As Long = &H20
Private Const OFN_ENABLETEMPLATE As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_EXTENSIONDIFFERENT As Long = &H400
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_LONGNAMES As Long = &H200000
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_NODEREFERENCELINKS As Long = &H100000
Private Const OFN_NOLONGNAMES As Long = &H40000
Private Const OFN_NONETWORKBUTTON As Long = &H20000
Private Const OFN_NOREADONLYRETURN As Long = &H8000& '*see comments
Private Const OFN_NOTESTFILECREATE As Long = &H10000
Private Const OFN_NOVALIDATE As Long = &H100
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_READONLY As Long = &H1
Private Const OFN_SHAREAWARE As Long = &H4000
Private Const OFN_SHAREFALLTHROUGH As Long = 2
Private Const OFN_SHAREWARN As Long = 0
Private Const OFN_SHARENOWARN As Long = 1
Private Const OFN_SHOWHELP As Long = &H10
Private Const OFN_ENABLESIZING As Long = &H800000
Private Const OFS_MAXPATHNAME As Long = 260

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or _
             OFN_LONGNAMES Or _
             OFN_CREATEPROMPT Or _
             OFN_NODEREFERENCELINKS

Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or _
             OFN_LONGNAMES Or _
             OFN_OVERWRITEPROMPT Or _
             OFN_HIDEREADONLY
             
'windows version constants
Private Const VER_PLATFORM_WIN32_NT As Long = 2
Private Const OSV_LENGTH As Long = 76
Private Const OSVEX_LENGTH As Long = 88
Private OSV_VERSION_LENGTH As Long

Private Const WM_INITDIALOG As Long = &H110
Private Const SW_SHOWNORMAL As Long = 1
Private Const SM_CYCAPTION As Long = 4

Private Const CDN_FIRST As Long = (-601)
Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_SETCONTROLTEXT As Long = CDM_FIRST + &H4
Private Const CDM_HIDECONTROL As Long = (CDM_FIRST + &H5)

Private Const IDOK As Long = 1
Private Const IDCANCEL As Long = 2
Private Const IDFILEOFTYPETEXT  As Long = &H441
Private Const IDFILENAMETEXT As Long = &H442
Private Const IDLOOKINTEXT As Long = &H443

Private FormObject As Form

Private Type OPENFILENAME
  nStructSize       As Long
  hWndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  sFile             As String
  nMaxFile          As Long
  sFileTitle        As String
  nMaxTitle         As Long
  sInitialDir       As String
  sDialogTitle      As String
  flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  sDefFileExt       As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
  pvReserved        As Long  'new in Windows 2000 and later
  dwReserved        As Long  'new in Windows 2000 and later
  FlagsEx           As Long  'new in Windows 2000 and later
End Type

Private OFN As OPENFILENAME

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function GetOpenFileName Lib "comdlg32" _
    Alias "GetOpenFileNameA" _
   (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetSaveFileName Lib "comdlg32" _
   Alias "GetSaveFileNameA" _
  (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" _
   (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" _
   (ByVal nIndex As Long) As Long
   
Private Declare Function GetParent Lib "user32" _
  (ByVal hWnd As Long) As Long

Private Declare Function SetWindowText Lib "user32" _
   Alias "SetWindowTextA" _
  (ByVal hWnd As Long, _
   ByVal lpString As String) As Long
   
Private Declare Function MoveWindow Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long
   
Private Declare Function GetWindowRect Lib "user32" _
  (ByVal hWnd As Long, _
   lpRect As RECT) As Long

'defined As Any to support either the
'OSVERSIONINFO or OSVERSIONINFOEX structure
Private Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
  (lpVersionInformation As Any) As Long

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

 
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
Dim lpIDList As Long
Dim ret As Long
Dim sBuffer As String
On Error Resume Next  'Sugested by MS to prevent an error from
                      'propagating back into the calling process.
 Select Case uMsg
  Case BFFM_INITIALIZED
   Call SendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
  Case BFFM_SELCHANGED
   sBuffer = space(MAX_PATH)
   ret = SHGetPathFromIDList(lp, sBuffer)
   If ret = 1 Then
    Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
   End If
 End Select
 BrowseCallbackProc = 0
End Function

'    If Len(folder) = 0 Then Exit Sub  'User Selected Cancel
'=====================================================================================
Public Function BrowseForFolder(owner As Form, _
                                Title As String, _
                                StartDir As String) As String
                                
'Opens a Treeview control that displays the directories in a computer
Dim lpIDList As Long
Dim szTitle As String
Dim sBuffer As String
Dim tBrowseInfo As BrowseInfo
 m_CurrentDirectory = StartDir & vbNullChar

 szTitle = Title
 With tBrowseInfo
  .hWndOwner = owner.hWnd
  .lpszTitle = lstrcat(szTitle, "")
  .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
  .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
 End With
  
 lpIDList = SHBrowseForFolder(tBrowseInfo)
 If (lpIDList) Then
  sBuffer = space(MAX_PATH)
  SHGetPathFromIDList lpIDList, sBuffer
  sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
  BrowseForFolder = sBuffer
 Else
  BrowseForFolder = ""
 End If
End Function

Private Property Let DialogHideOK(ByVal vNewValue As Boolean)

   m_dlgHideOK = vNewValue

End Property

Private Property Let DialogInitPosition(ByVal vNewValue As Integer)

   m_dlgPosition = vNewValue
   
End Property

Private Function FARPROC(ByVal pfn As Long) As Long
  
  'A dummy procedure that receives and returns
  'the return value of the AddressOf operator.
 
  'Obtain and set the address of the callback
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
 
  FARPROC = pfn

End Function

Public Function GenericFileOpenBrowser( _
            ByRef dialogFileBrowser As CommonDialog, _
            Optional ByVal DefaultPath As String = vbNullString, _
            Optional ByVal DialogTitle As String = "File Browser", _
            Optional ByVal DefaultFilter As String = vbNullString, _
            Optional ByVal DefaultFlags As Long = &H0&, _
            Optional ByVal DefaultFilename As String = vbNullString, _
            Optional ByVal DefaultExtension As String = vbNullString) As String
                                  

    Dim nfile As File
    
    Dim TempL As Long
    Dim TempL2 As Long
    
    Dim filename As String
    Dim FilePath As String
    Dim FileExtension As String
    Dim InitialPath As String
    Dim TempS As String

    Dim TempB As Boolean

    'Need to Browser for the new usage file to use
    'First check and see if the given Usage file exists
    If Not FileExists(DefaultPath) Or DefaultPath = vbNullString Then
    
        'Currently set usage file does not exist,
        'start the initial check in the folder one above
        'the VB project file folder
        FilePath = App.path
        
        'Get the position of the last "\" in the application filepath
        TempL = InStrRev(FilePath, "\")
        
        'Get the position of the second to last "\" in the application filepath
        TempL2 = InStrRev(FilePath, "\", TempL - 1)
        
        'Set the Initial Path, include the "\" at the end
        InitialPath = Mid(FilePath, 1, TempL2)
        
        'Set the filename to the default
        filename = DefaultFilename
        
        'Set the file extension to the default
        FileExtension = DefaultExtension
    
    Else
    
        'Get the parent folder of the usage file
        FilePath = DefaultPath
        
        'Get the position of the last "\" in filepath
        TempL = InStrRev(FilePath, "\")
    
        'Set Initial path to the parent directory
        InitialPath = Mid(FilePath, 1, TempL)
        
        'Set the Filename, if the default file name is empty
        If DefaultFilename = vbNullString Then
        
            filename = Mid(FilePath, TempL + 1)
                    
        Else
        
            filename = DefaultFilename
            
        End If
        
        'Set the Default extension
        If DefaultExtension = vbNullString Then
        
            'Get the extension from the end of the Filename
            TempL = InStrRev(filename, ".")
            
            If TempL = 0 Then
            
                'No matching "."
                FileExtension = vbNullString
                
            Else
            
                FileExtension = Mid(filename, TempL)
                
            End If
            
        End If
                    
    End If
    
    'Setup the file dialog
    With dialogFileBrowser

        'If Default filter = "", don't set a filter
        If DefaultFilter <> vbNullString Then
            
            .filter = DefaultFilter
            
        End If
        
        .flags = DefaultFlags
        .DefaultExt = FileExtension
        .filename = filename
        .DialogTitle = DialogTitle
        .InitDir = InitialPath
        
        'Pop-open the dialog box and get the user to pick a file
        .ShowOpen
        
        'Get the filepath
        GenericFileOpenBrowser = .filename
        
    End With

End Function

' This function allows you to assign a function pointer to a vaiable.
Private Function GetAddressofFunction(Add As Long) As Long
 GetAddressofFunction = Add
End Function

Private Function IsWin2000Plus() As Boolean

  'returns True if running Windows 2000 or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWin2000Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                      (osv.dwVerMajor = 5 And osv.dwVerMinor >= 0)
  
   End If

End Function

Private Function OFNHookProc(ByVal hWnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long
                                   
  'On initialization, set aspects of the
  'dialog that are not obtainable through
  'manipulating the OPENFILENAME structure members.
   Dim hwndParent As Long
   Dim rc As RECT
   
  'temporary vars for demo
   Dim newLeft As Long
   Dim newTop As Long
   Dim dlgWidth As Long
   Dim dlgHeight As Long
   Dim scrWidth As Long
   Dim scrHeight As Long
   Dim frmLeft As Long
   Dim frmTop As Long
   Dim frmWidth As Long
   Dim frmHeight As Long
            
   Select Case uMsg
      Case WM_INITDIALOG
      
        'obtain the handle to the parent dialog
         hwndParent = GetParent(hWnd)
         
         If hwndParent <> 0 Then
            
           'Get the current dialog size and position
            Call GetWindowRect(hwndParent, rc)
                        
           'Once again, to show the calculations involved
           'I'll use variables instead of creating a
           'one-line MoveWindow call.
            Select Case m_dlgPosition
               Case 0:  'normal position
                     
                  OFNHookProc = 0
               
               Case 1:  'centered on screen

                  dlgWidth = rc.Right - rc.Left
                  dlgHeight = rc.Bottom - rc.Top
                  scrWidth = Screen.Width \ Screen.TwipsPerPixelX
                  scrHeight = Screen.Height \ Screen.TwipsPerPixelY
                  newLeft = (scrWidth - dlgWidth) \ 2
                  newTop = (scrHeight - dlgHeight) \ 2
                  
                  Call MoveWindow(hwndParent, newLeft, newTop, dlgWidth, dlgHeight, True)
                  
                  OFNHookProc = 1

               Case 2:  'centered in parent
               
                  frmLeft = FormObject.Left \ Screen.TwipsPerPixelX
                  frmTop = FormObject.Top \ Screen.TwipsPerPixelY
                  frmWidth = FormObject.Width \ Screen.TwipsPerPixelX
                  frmHeight = FormObject.Height \ Screen.TwipsPerPixelX

                  dlgWidth = rc.Right - rc.Left
                  dlgHeight = rc.Bottom - rc.Top
                  
                  scrWidth = Screen.Width \ Screen.TwipsPerPixelX
                  scrHeight = Screen.Height \ Screen.TwipsPerPixelY
                  
                  newLeft = frmLeft + ((frmWidth - dlgWidth) \ 2)

                  If dlgHeight > frmHeight Then
                     newTop = frmTop + GetSystemMetrics(SM_CYCAPTION) + ((frmHeight - dlgHeight) \ 2)
                  Else
                     newTop = frmTop + GetSystemMetrics(SM_CYCAPTION)
                  End If
                  
                  Call MoveWindow(hwndParent, newLeft, newTop, dlgWidth, dlgHeight, True)
                  
                  OFNHookProc = 1
                           
            End Select
            
           'If the hide check is set, hide the OK button.
           'This simply shows how easy it is to hide
           'unwanted control elements.
            If m_dlgHideOK = True Then
               Call SendMessage(hwndParent, CDM_HIDECONTROL, _
                                IDOK, ByVal 0&)
            End If
            
           'If the length of the variables > 0, set
           'the new text to the respective control.
            If Len(m_OKCaption) > 0 Then
               Call SendMessage(hwndParent, CDM_SETCONTROLTEXT, _
                                IDOK, ByVal m_OKCaption)
            End If
            
            If Len(m_CancelCaption) > 0 Then
               Call SendMessage(hwndParent, CDM_SETCONTROLTEXT, _
                                IDCANCEL, ByVal m_CancelCaption)
            End If
            
            If Len(m_LookInCaption) > 0 Then
               Call SendMessage(hwndParent, CDM_SETCONTROLTEXT, _
                                IDLOOKINTEXT, ByVal m_LookInCaption)
            End If
  
            If Len(m_FileNameCaption) > 0 Then
               Call SendMessage(hwndParent, CDM_SETCONTROLTEXT, _
                                IDFILENAMETEXT, ByVal m_FileNameCaption)
            
            End If
            If Len(m_FileOfTypeCaption) > 0 Then
               
               Call SendMessage(hwndParent, CDM_SETCONTROLTEXT, _
                                IDFILEOFTYPETEXT, ByVal m_FileOfTypeCaption)
            End If
            
         End If
         
         Case Else
         
   End Select

End Function

'Function OpenDirectory
'Made March 2010, by Isaac Hilburn
'Adds a public function shell to the above code
'so that the user doesn't have to deal with all the nitty-gritty of the contents
'of the OPENFILENAME object

Public Function OpenDir _
    (ByVal StartDir As String, _
     ByVal BrowserTitle As String, _
     ByRef OwningFormObject As Form) As String

    Dim FolderPath As String
    
    'Set module private variable, FormObject to the reference of the owning form object
    'that was passed in by the user
    Set FormObject = OwningFormObject
    
    'The following are properties that need to be set using
    '.dll function aliases in the modFileSave module.
    'I'm explicitly referencing the functions using (function name)
    'to prevent any confusion/conflicts with other Paleomag code API / .dll file calls
        
    'if first time through set the appropriate OFN size
    If OSV_VERSION_LENGTH = 0 Then Call SetOSVersion
   
    'assign the new caption properties as needed
    'These options allow the common dialog API object to be changed into
    'a directory browser instead of a file browser
    SetOKCaption = "Sel Cur Dir..."
    SetCancelCaption = "Cancel"
    SetLookInCaption = StartDir
    SetFileNameCaption = ""
    SetFileOfTypeCaption = ""

    'Set the properties of the file browser object
    'that we're about to call from the file browse API in commdlg32.dll
     With OFN
     
        .nStructSize = OSV_VERSION_LENGTH
        .hWndOwner = FormObject.hWnd
        .sFilter = "*.dud"
        .nFilterIndex = 2
        .sFile = ".dud" & space$(1024) & vbNullChar & vbNullChar
        .nMaxFile = Len(.sFile)
        .sDefFileExt = "bas" & vbNullChar & vbNullChar
        .sFileTitle = vbNullChar & space$(512) & vbNullChar & vbNullChar
        .nMaxTitle = Len(OFN.sFileTitle)
        .sInitialDir = StartDir & vbNullChar & vbNullChar
        .sDialogTitle = BrowserTitle
        .flags = OFN_EXPLORER Or _
                 OFN_ENABLEHOOK Or _
                 OFN_NOVALIDATE
        .fnHook = FARPROC(AddressOf OFNHookProc)
    
     End With
     
    'call the API function from modFileSave
    Call GetOpenFileName(OFN)

    'Now extract the directory path from the OFN object file path
    FolderPath = Trim(OFN.sFile)
    
    'Use the Filename to remove that name from the FolderPath
    FolderPath = Mid(FolderPath, 1, InStrRev(FolderPath, "\"))
    
    'Return resulting folder path
    OpenDir = FolderPath
    
End Function
    


'Browse for a Folder using SHBrowseForFolder API function
'with a callback
' function BrowseCallbackProc.
' This Extends the functionality that was given in the
' MSDN Knowledge Base article Q179497 "HOWTO: Select a
'Directory
' Without the Common Dialog Control".
' After reading the MSDN knowledge base article Q179378
'"HOWTO: Browse for
' Folders from the Current Directory", I was able to
'figure out how to add
' a callback function that sets the starting directory
'and displays the
' currently selected path in the "Browse For Folder"
'dialog.
'==========================================================
' Usage:
'    Dim folder As String
'    folder = BrowseForFolder(Me, "Select A Directory", _
'     "C:\startdir\anywhere")

Public Sub PrintRichText(ByRef dlgObj As CommonDialog, _
                         ByRef richTextObj As RichTextBox, _
                         ByVal MsgStr As String, _
                         Optional ByVal PrintTitle As String = "Print ...")
                         
    'Dimension a font object
    Dim strLineSpacing As String
    Dim strRtf As String

    'Raise a print dialog to set the printer setting
    With dlgObj
        
        'use the default printer
        .PrinterDefault = True
        
        'Allow the cancel error
        .CancelError = True
        
        'Set settings to 1 copy, and landscape mode
        .Copies = 1
        .DialogTitle = PrintTitle
        .Orientation = cdlLandscape
        
        'Set the print dialog flags to
        'turn on the show page numbers and print all pages
        'and to return which printer was selected for printing
        .flags = cdlPDAllPages & _
                 cdlPDReturnDC & _
                 cdlPDPageNums
        
        'Show the printer dialog, checking for a cancel error
        On Error Resume Next
        
            .ShowPrinter
            
            If Err <> 0 Then
            
                'Pop-up message box
                MsgBox "Could not print text." & vbNewLine & vbNewLine & _
                       "Error: " & Err.Description, , _
                       "Print Error!"
            
                'Exit this subroutine
                Exit Sub
                
            End If
            
        'Resume normal error flow
        On Error GoTo 0
        
        'Set the Rich text box's text to that of the message
        richTextObj.text = MsgStr
        
        'select the contexts of the richtextbox
        richTextObj.SelStart = 1
        richTextObj.SelLength = Len(MsgStr)
        
        'Set the default font to arial, the size to 12, and the
        'spacing to double
        'Error check in case the arial font is missing
        'If so, the text will be printed with whatever font and font size
        'settings that the richText box already has
        On Error GoTo BadFont:
        
            richTextObj.SelFontName = "Arial"
            richTextObj.SelFontSize = 12
            
        On Error GoTo 0
        
BadFont:

        'Now change the rich-text box's line spacing
        'to double spaced
        SetRTFLineSpacing richTextObj, 24
        
        'Error check this
        On Error Resume Next
            
            'Now print the contents of the rich-text control object
            richTextObj.SelPrint .hDC
            
            If Err <> 0 Then
            
                'Pop-up message box
                MsgBox "Could not print text." & vbNewLine & vbNewLine & _
                       "Error: " & Err.Description, , _
                       "Print Error!"
                       
                Exit Sub
                
            End If
            
        On Error GoTo 0
        
    End With
        
End Sub

Private Property Let SetCancelCaption(ByVal vNewValue As String)

   m_CancelCaption = vNewValue

End Property

Private Property Let SetFileNameCaption(ByVal vNewValue As String)

   m_FileNameCaption = vNewValue

End Property

Private Property Let SetFileOfTypeCaption(ByVal vNewValue As String)

   m_FileOfTypeCaption = vNewValue

End Property

Private Property Let SetLookInCaption(ByVal vNewValue As String)

   m_LookInCaption = vNewValue

End Property

Private Property Let SetOKCaption(ByVal vNewValue As String)

   m_OKCaption = vNewValue

End Property

Private Sub SetOSVersion()
  
   Select Case IsWin2000Plus()
      Case True
         OSV_VERSION_LENGTH = OSVEX_LENGTH '5.0+ structure size
      
      Case Else
         OSV_VERSION_LENGTH = OSV_LENGTH   'pre-5.0 structure size
   End Select

End Sub

'Now rewrite the RTF formating in inputed RichTextBox control
'to make the lines double spaced
'Code taken from: http://www.bigresource.com/Tracker/Track-vb-h3LAFwMbGh/
Public Sub SetRTFLineSpacing(ByRef richTextObj As RichTextBox, _
                             ByVal HeightInTwips As Long)
                             
    Dim strRtf As String
    Dim strLineSpacing As String
                             
    strRtf = richTextObj.TextRTF
    strLineSpacing = "\sl" & Trim$(Str$(HeightInTwips))
    
    '[b]Add code to remove previous LineSpacing here[/b]
    
    strRtf = Replace(strRtf, "\pard", "\pard" & strLineSpacing, 1, -1, vbTextCompare)
    richTextObj.TextRTF = strRtf

End Sub

