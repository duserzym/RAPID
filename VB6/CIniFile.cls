VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' *************************************************************************
'  Copyright ©1994-2007 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Win32 API declarations
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpApplicationName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpApplicationName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

' Some default values
Private Const defBufferSize As Long = 8192  '(April 2010 - I Hilburn) Changed from 2048 to 8192
Private Const defStripComments As Boolean = True

' Member variables
Private m_FileName As String
Private m_Section As String
Private m_StripComments As Boolean

' Public events
Public Event EnumIniSection(ByVal SectionName As String, ByVal filename As String, Cancel As Boolean)
Public Event EnumIniSectionEntry(ByVal EntryName As String, ByVal SectionName As String, ByVal filename As String, Cancel As Boolean)

' *************************************************
'  Public Properties
' *************************************************
Public Property Get filename() As String
   filename = m_FileName
End Property

Public Property Let filename(ByVal IniFile As String)
   m_FileName = IniFile
End Property

Public Property Get Section() As String
   Section = m_Section
End Property

Public Property Let Section(ByVal IniSection As String)
   m_Section = IniSection
End Property

Public Property Get StripComments() As Boolean
   StripComments = m_StripComments
End Property

Public Property Let StripComments(ByVal NewSetting As Boolean)
   m_StripComments = NewSetting
End Property

Public Sub AddLineBreaks()

    Dim fso As FileSystemObject
    Dim FStream As TextStream
    Dim FileStr As String
    Dim LineStr As String

    'Need to crawl through the file and add a line break before every section heading
    'Need to use FileSystemObject and TextStreams to do this
    
    'Default Local string variables to vbnullstr
    FileStr = vbNullString
    LineStr = vbNullString
    
    'Allocate fso object
    Set fso = New FileSystemObject
    
    'Error check
    On Error GoTo BadFile:
        
        'Open the text stream
        Set FStream = fso.OpenTextFile(m_FileName, ForReading)
    
    On Error GoTo 0
    
    'Load the first linestring before the loop starts
    LineStr = FStream.ReadLine
    
    'Now loop through the INI file until hit EOF
    Do While FStream.AtEndOfStream = False
    
        If Left(Trim(LineStr), 1) = "[" Then
        
            FileStr = FileStr & vbCrLf & vbCrLf & LineStr
            
        Else
        
            FileStr = FileStr & vbCrLf & LineStr
            
        End If
        
        'Get the next line string
        LineStr = FStream.ReadLine
        
    Loop
    
    'Now have the whole file + line breaks stored to FileStr
    
    'Close the INI File Text Stream
    FStream.Close
    
    'Open a new file text stream to overwrite the INI file
    Set FStream = fso.OpenTextFile(m_FileName, ForWriting)
    
    'Write the FileStr to the INI file
    FStream.Write FileStr
    
    'Close the file text stream
    FStream.Close
    
    'Deallocate objects
    Set fso = Nothing
    Set FStream = Nothing

BadFile:

    'Do nothing

End Sub

' *************************************************
'  Initialization
' *************************************************
Private Sub Class_Initialize()
   m_StripComments = defStripComments
End Sub

Private Sub Class_Terminate()
   Me.Flush  'can't hurt
End Sub

' *************************************************
'  Public Methods
' *************************************************
Public Function EntryClear(ByVal EntryName As String, Optional ByVal Section As String, Optional ByVal filename As String) As Boolean
   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, filename)
   ' Write an empty string to entry.
   EntryClear = CBool(WritePrivateProfileString(m_Section, EntryName, "", m_FileName))
End Function

Public Function EntryDelete(ByVal EntryName As String, Optional ByVal Section As String, Optional ByVal filename As String) As Boolean
   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, filename)
   ' Totally zap this entry.
   EntryDelete = CBool(WritePrivateProfileString(m_Section, EntryName, vbNullString, m_FileName))
End Function

Public Function EntryRead(ByVal EntryName As String, Optional ByVal DefaultValue As String, Optional ByVal Section As String, Optional ByVal filename As String) As String
   Dim buffer As String
   Dim nRet As Long

   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, filename)

   ' Establish buffer for return info.
   buffer = space$(defBufferSize)
   
   ' May need several attempts if entry is huge.
   Do
      nRet = GetPrivateProfileString(m_Section, EntryName, DefaultValue, buffer, Len(buffer), m_FileName)
      If nRet = (Len(buffer) - 1) Then
         ' Need to expand buffer and try again.
         ' SDK Docs specify Len-1 for this case.
         buffer = space$(Len(buffer) + defBufferSize)
      ElseIf nRet > 0 Then
         ' Trim null and remaining buffer.
         buffer = Left$(buffer, nRet)
         Exit Do
      Else
         ' Function failed.
         buffer = ""
         Exit Do
      End If
   Loop
   
   ' Strip off everything from the semi-colon on?
   If m_StripComments Then
      nRet = InStr(buffer, ";")
      If nRet = 1 Then
         buffer = ""
      ElseIf nRet > 1 Then
         buffer = Left$(buffer, nRet - 1)
      End If
   End If
   
   ' Return results.
   EntryRead = Trim$(buffer)
End Function

Public Function EntryWrite(ByVal EntryName As String, ByVal EntryValue As String, Optional ByVal Section As String, Optional ByVal filename As String) As String
   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, filename)
   ' Write this value as requested.
   Call WritePrivateProfileString(m_Section, EntryName, EntryValue, m_FileName)
End Function

Public Sub EnumSectionEntries(Optional ByVal Section As String, Optional ByVal filename As String)
   Dim nRet As Long, i As Long
   Dim buffer As String
   Dim arrEntry() As String
   Dim BailOut As Boolean

   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, filename)

   ' Establish buffer for return info.
   buffer = space$(defBufferSize)
   
   ' May need several attempts if large
   ' number of entries.
   Do
      nRet = GetPrivateProfileString(m_Section, vbNullString, "", buffer, Len(buffer), m_FileName)
      If nRet = (Len(buffer) - 2) Then
         ' Need to expand buffer and try again.
         ' SDK Docs specify Len-2 for this case.
         buffer = space$(Len(buffer) + defBufferSize)
      Else
         Exit Do
      End If
   Loop
   
   ' Extract entry names.
   If ExtractStringsZ(buffer, arrEntry()) Then
      ' Raise an event for each element in the array.
      For i = LBound(arrEntry) To UBound(arrEntry)
         ' Offer user chance to stop enumeration by bailing
         ' from loop if the last parameter is set to True.
         RaiseEvent EnumIniSectionEntry(arrEntry(i), m_Section, m_FileName, BailOut)
         If BailOut Then Exit For
      Next i
   End If
End Sub

Public Sub EnumSections(Optional ByVal filename As String)
   Dim nRet As Long, i As Long
   Dim buffer As String
   Dim arrSection() As String
   Dim BailOut As Boolean
   
   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(, filename)
   
   ' Establish buffer for return info.
   buffer = space$(defBufferSize)
   
   ' May need several attempts if large
   ' number of sections.
   Do
      nRet = GetPrivateProfileSectionNames(buffer, Len(buffer), m_FileName)
      If nRet = (Len(buffer) - 2) Then
         ' Need to expand buffer and try again.
         ' SDK Docs specify Len-2 for this case.
         buffer = space$(Len(buffer) + defBufferSize)
      Else
         Exit Do
      End If
   Loop
   
   ' Extract section names.
   If ExtractStringsZ(buffer, arrSection()) Then
      ' Raise an event for each element in the array.
      For i = LBound(arrSection) To UBound(arrSection)
         ' Offer user chance to stop enumeration by bailing
         ' from loop if the last parameter is set to True.
         RaiseEvent EnumIniSection(arrSection(i), m_FileName, BailOut)
         If BailOut Then Exit For
      Next i
   End If
End Sub

Private Function ExtractStringsZ(buffer As String, OutArray() As String) As Long
   Dim StartPos As Long
   Dim NullPos As Long
   Dim BuffLen As Long
   Dim Elements As Long
   
   ' Extract null terminated strings from large
   ' double-null terminated buffer.
   StartPos = 1
   Elements = 0
   BuffLen = Len(buffer)
   
   ' Loop through buffer looking for nulls.
   Do While StartPos < BuffLen
      NullPos = InStr(StartPos, buffer, vbNullChar)
      If NullPos = StartPos Then
         ' We've hit the double-null terminator.
         Exit Do
      Else
         ' Expand array, store new substring, and
         ' increment counters.
         ReDim Preserve OutArray(0 To Elements) As String
         OutArray(Elements) = Mid$(buffer, StartPos, NullPos - StartPos)
         StartPos = NullPos + 1
         Elements = Elements + 1
      End If
   Loop
   
   ' Return number of substrings found.
   ExtractStringsZ = Elements
End Function

Public Sub Flush(Optional ByVal filename As String)
   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(, filename)
   ' To improve performance, Windows keeps a cached version of
   ' the most-recently accessed initialization file. If that
   ' filename is specified and the other three parameters are
   ' NULL, Windows flushes the cache.
   Call WritePrivateProfileString(vbNullString, vbNullString, vbNullString, m_FileName)
End Sub

   
'Public Function SectionClear
'
'   Created: July 1, 2010
'    Author: Isaac Hilburn
'
'   Summary:    Function is designed to clear out either all the entries and values in a section,
'               or just the values in a section.  This function uses SectionRead to get
'               all the entries in the section to modify, and Calls EntryClear or EntryDelete
'               to remove the entire entry or just the value in the entry, depending on the
'               boolean flag ClearAll
'
Public Function SectionClear(Optional ByVal Section As String, _
                             Optional ByVal ClearAll As Boolean = False, _
                             Optional ByVal filename As String) As Boolean

    
    Dim SectionEntries As Variant     'Allocate variant to store the returned
                                      'string array from SectionRead
                                        
    Dim i As Long
    Dim N As Long
    Dim TempBool As Boolean
    
    If Section = "" Then Section = m_Section
    If filename = "" Then filename = m_FileName
    
    SectionEntries = SectionRead(False, _
                                 True, _
                                 Section, _
                                 filename)
                                                              
    'Get the size of the section entries array
    N = UBound(SectionEntries)
    
    'Default TempBool = True
    TempBool = True
    
    'Iterate through the array, and perform the entry clear or delete action
    'for each entry
    For i = 0 To N - 1
    
        If ClearAll = True Then
        
            TempBool = TempBool And EntryDelete(SectionEntries(i), _
                                                Section, _
                                                filename)
                                                
        Else
        
            TempBool = TempBool And EntryClear(SectionEntries(i), _
                                               Section, _
                                               filename)
                                               
        End If
        
    Next i
    
    'Return TempBool as success / failure status
    SectionClear = TempBool
                        
End Function

Public Function SectionDelete(Optional ByVal Section As String, Optional ByVal filename As String) As Boolean
      ' Update cached section and filename, if provided.
      Call UpdateFileDetails(Section, filename)
      ' Remove specific section from INI file.
      SectionDelete = CBool(WritePrivateProfileString(m_Section, vbNullString, vbNullString, m_FileName))
   End Function

Public Function SectionExists(ByVal Section As String, _
                              Optional ByVal filename As String) As Boolean

    Dim buffer As String
    Dim nRet As Long
    Dim i As Long
    Dim SectionArray() As String
   
    ' Update cached section and filename, if provided.
    Call UpdateFileDetails(Section, filename)
    
    ' Establish buffer for return info.
    buffer = space$(defBufferSize)
    
    ' May need several attempts if large
    ' number of entries.
    Do
       nRet = GetPrivateProfileString(m_Section, vbNullString, "", buffer, Len(buffer), m_FileName)
       If nRet = (Len(buffer) - 2) Then
          ' Need to expand buffer and try again.
          ' SDK Docs specify Len-2 for this case.
          buffer = space$(Len(buffer) + defBufferSize)
       Else
          Exit Do
       End If
    Loop
    
    'Default SectionExists to false
    SectionExists = False
    
    If ExtractStringsZ(buffer, SectionArray()) > 0 Then SectionExists = True
    
End Function

Public Function SectionRead(Optional IncludeValues As Boolean = False, _
                            Optional ReturnAsArray As Boolean = False, _
                            Optional ByVal Section As String, _
                            Optional ByVal filename As String) As Variant
                            
   ' Modified:
   ' April 4, 2010 - Isaac Hilburn
   '
   ' If ReturnAsArray = False then:
   '
   ' This function will return a giant string
   ' containing either the entry names (similar to EnumSectionEntries)
   ' or the complete entries, including the values for each, separated by a
   ' Null Char
   '
   ' If ReturnAsArray = True then:
   '
   ' This function will return a string array filled with either
   ' the entry names (similar to EnumSectionEntries) or the
   ' complete entries (including the values for each).
   
   Dim nRet As Long, i As Long
   Dim buffer As String
   Dim arrEntry() As String

   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, filename)

   ' Establish buffer for return info.
   buffer = space$(defBufferSize)
   
   ' May need several attempts if large
   ' number of entries.
   Do
      
      ' Use different APIs depending on whether they
      ' want to include full entries, or just names.
      If IncludeValues Then
         
         nRet = GetPrivateProfileSection(m_Section, buffer, Len(buffer), m_FileName)
      
      Else
         
         nRet = GetPrivateProfileString(m_Section, vbNullString, "", buffer, Len(buffer), m_FileName)
      
      End If
         
      If nRet = (Len(buffer) - 2) Then
         
         ' Need to expand buffer and try again.
         ' SDK Docs specify Len-2 for this case.
         buffer = space$(Len(buffer) + defBufferSize)
         
      Else
      
         Exit Do
         
      End If
      
   Loop
  
'------------------------------------------------------'
'------------------------------------------------------'
'
'   Modified: April 4, 2010
'     Author: Isaac Hilburn
'
'    Summary: Want to change SectionRead so that it
'             gives the option for the user to specify
'             if the Section contents should be returned
'             as a string array or a giant string
'------------------------------------------------------'

    'If user has selected to return the results as a string array, do so
    If ReturnAsArray = True Then

        ' Extract entry names.
        If ExtractStringsZ(buffer, arrEntry()) Then
           
           SectionRead = arrEntry
           
        End If
        
    Else

        'Trim the empty, unfilled spaces from the returned
        'Buffer string, will overwrite Buffer with a trimmed
        'version of itself
        TrimReturnBuffer buffer
    
        'Return the trimmed Buffer string
        SectionRead = buffer

    End If
'------------------------------------------------------'
'------------------------------------------------------'
End Function

'New SectionWrite function
' Created:  April 4, 2010
'  Author:  Isaac Hilburn
'           RAPID Consortium
'
' Summary: Uses the WritePrivateProfileSection API function to
'          write a string buffer containing all the keys and associated values in an
'          entire section of a .ini file.
'
'  Inputs:
'
'   NewBuffer   -   Required string containing the new keys and values for the section.
'                   Format: <key as string> & "=" & <value as string> & vbNullString
'                   Each key + value combination must be terminated with a Null character
'
'   Section     -   Optional string containing the name of the section that you're trying
'                   to write to. If no value for sSection is input, it defaults to the
'                   value of the class private local variable m_Section that can be set
'                   through the Section property of CIniFile class
'
'   IniFileName -   Optional string containing the absolute path + filename of the .ini
'                   file that you're writing to.  If no value is input, it defaults to
'                   m_FileName, settable through the FileName property of the CIniFile class
'
Public Function SectionWrite(ByVal NewBuffer As String, Optional ByVal sSection As String = "", Optional ByVal IniFileName As String = "") As Boolean

    Dim StrByteSize As Long
    Dim lRet As Long
    
    'If no values for sSection and IniFileName were given, then need to
    'load the local variable values into one or both of them
    If sSection = "" Then sSection = m_Section
    If IniFileName = "" Then IniFileName = m_FileName
    
    'Get String length and Byte size
    StrByteSize = String_SizeInBytes(NewBuffer)
    
    'Check to see if string byte size is greater than the maximimum byte size the DLL function can take
    If StrByteSize > 65535 Then
    
        'This string is too large to written to the .ini file section!
        'Return a false and exit the function
        SectionWrite = False
        
        Exit Function
        
    End If
    
    'String size is okay, call the write profile
    lRet = WritePrivateProfileSection(sSection, NewBuffer, IniFileName)

    'Error check
    If lRet = 0 Then
    
        'Error occurred
        'Return a false and exit the function
        SectionWrite = False
        
        Exit Function
        
    End If
    
    'Write was successful!
    SectionWrite = True
    
End Function

Public Function ToBoolean(ByVal EntryValue As String) As Boolean
   ' Interpret entry as either true or false.
   Select Case Trim$(UCase$(EntryValue))
      Case "YES", "Y", "TRUE", "T", "ON", "1", "-1"
         ToBoolean = True
      Case "NO", "N", "FALSE", "F", "OFF", "0"
         ToBoolean = False
      Case Else
         ToBoolean = False
   End Select
End Function

Public Function ToDouble(ByVal EntryValue As String) As Double
   ' Return interpreted floating point value of entry.
   On Error Resume Next
   ToDouble = CDbl(val(EntryValue))
End Function

Public Function ToLong(ByVal EntryValue As String) As Long
   ' Return integral portion of entry.
   On Error Resume Next
   ToLong = CLng(val(EntryValue))
End Function

' *************************************************
'  Private Methods
' *************************************************

' Created: April 4, 2010
'  Author: Isaac Hilburn
'
' Summary: New function to trim the unfilled end of the return string
'          buffer from a successful section read.  Function uses InStr
'          to find the first double Null char in the Buffer string,
'          and then uses Mid() to remove the section of Buffer from the
'          the first character to the character just before the double Null.
'
'   Input:
'
'     Buffer    -   ByRef String to be trimmed of it's unfilled characters
'
'  Output:
'
'     Buffer    -   ByRef string - overwrites input string with filled
'                   portion of the Buffer string
'

Private Sub TrimReturnBuffer(ByRef buffer As String)

    Dim DoubleNullPos As Long
    Dim BuffLen As Long
    Dim TempStr As String
    
    DoubleNullPos = InStr(1, buffer, vbNullChar & vbNullChar)
    
    TempStr = Mid(buffer, 1, DoubleNullPos - 1)
    
    buffer = TempStr

End Sub

Private Sub UpdateFileDetails(Optional ByVal Section As String, Optional ByVal filename As String)
   ' Cache new values if different from old ones.
   If Len(Section) Then
      m_Section = Section
   End If
   If Len(filename) Then
      m_FileName = filename
   End If
End Sub

