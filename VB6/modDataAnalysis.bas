Attribute VB_Name = "modDataAnalysis"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32" _
    Alias "ShellExecuteA" _
   (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Option Explicit
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31

' This part by Caltech Paleomag lab
Public Sub DataAnalysis_SAMFile(filename As String, Optional filedir As String = 0&)
' (June 2008 L Carporzen) Visualisation of the data instead of looking at the text file
' This line has been deleted by MIT
    'RunShellExecute "Open", filename, vbNullString, filedir, SW_SHOWNORMAL
    frmPlots.RefreshSamples
    frmPlots.ZOrder
    frmPlots.Show
    frmPlots.SetFocus
    frmPlots.Actualize
End Sub

Public Sub DataAnalysis_SampleFile(filename As String, Optional filedir As String = 0&)
' (June 2008 L Carporzen) Visualisation of the data instead of looking at the text file
' This line has been deleted by MIT
    'RunShellExecute "Open", "notepad.exe", filename, filedir, SW_SHOWNORMAL
    If LenB(frmMagnetometerControl.cmbManSample.text) = 0 Then Exit Sub
    frmPlots.RefreshSamples
    frmPlots.cmbSamples.text = frmMagnetometerControl.cmbManSample.text
    frmPlots.ZOrder
    frmPlots.Show
    frmPlots.SetFocus
    frmPlots.Actualize
End Sub

Public Sub Open_SAMdirectory(filename As String, Optional filedir As String = 0&)
' (June 2008 L Carporzen) Open directory
    RunShellExecute "Open", vbNullString, filename, filedir, SW_SHOWNORMAL
End Sub

Public Sub Open_SAMFile(filename As String, Optional filedir As String = 0&)
' (June 2008 L Carporzen) Open text file
    RunShellExecute "Open", "notepad.exe", filename, filedir, SW_SHOWNORMAL
End Sub

Public Sub Open_SampleFile(filename As String, Optional filedir As String = 0&)
' (June 2008 L Carporzen) Open text file
    If LenB(frmPlots.cmbSamples.text) = 0 Then Exit Sub
    RunShellExecute "Open", "notepad.exe", filename, filedir, SW_SHOWNORMAL
End Sub

Private Sub RunShellExecute(sTopic As String, _
                           sFile As Variant, _
                           sParams As Variant, _
                           sDirectory As Variant, _
                           nShowCmd As Long)
   Dim hWndDesk As Long
   Dim success As Long
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)
  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  If success = SE_ERR_NOASSOC Then
     Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  End If
End Sub

