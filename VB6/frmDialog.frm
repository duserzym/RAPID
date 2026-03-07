VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7455
   ClientLeft      =   5685
   ClientTop       =   3060
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7455
   ScaleMode       =   0  'User
   ScaleWidth      =   10095
   Begin VB.TextBox txtMessageWithScroll 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmDialog.frx":0000
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox txtMessage 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmDialog.frx":0008
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdNegative 
      Caption         =   "Negative"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdPositive 
      Caption         =   "Positive"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UserResponse As Long
Public ActiveBox As String
Public NumButtons As Long

Private Sub cmdCancel_Click()

    UserResponse = vbCancel

End Sub

Private Sub cmdNegative_Click()

    UserResponse = vbNo

End Sub

Private Sub cmdPositive_Click()

    UserResponse = vbYes

End Sub

Public Function DialogBox( _
                    ByVal MsgStr As String, _
                    Optional ByVal TitleStr As String = "Dialog Window", _
                    Optional ByVal NumBtns As Long = 1, _
                    Optional ByVal ButtonName1 As String = "Positive", _
                    Optional ByVal ButtonName2 As String = "Negative", _
                    Optional ByVal ButtonName3 As String = "Cancel") As Long
                     
    'Have this form load
    Load Me
    
    txtMessage = MsgStr
    Caption = TitleStr
    
    'Coerce the number of buttons specified to the right number (1 - 3)
    If NumBtns < 1 Then NumBtns = 1
    If NumBtns > 3 Then NumBtns = 3
    
    NumButtons = NumBtns
    
    cmdPositive.Enabled = (NumButtons >= 1)
    cmdPositive.Caption = ButtonName1
    cmdPositive.Visible = (NumButtons >= 1)
           
    cmdNegative.Enabled = (NumButtons >= 2)
    cmdNegative.Caption = ButtonName2
    cmdNegative.Visible = (NumButtons >= 2)
            
    cmdCancel.Enabled = (NumButtons = 3)
    cmdCancel.Caption = ButtonName3
    cmdCancel.Visible = (NumButtons = 3)
    
    'Call the resize function to adjust the dialog window as needed
    ReSizeWindow
    
    'Now, wait until the user clicks a button
    Do
    
        DoEvents
        
        'Pause for 20 ms
        PauseTill timeGetTime() + 20
        
    Loop Until UserResponse <> -1
    
    DialogBox = UserResponse
    
    Unload Me
    
End Function

Private Sub Form_Load()

    If frmProgram.Visible = False Then frmProgram.Show
    Me.Show
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

    'Set normal dialog box defaults
    txtMessage.text = ""
    txtMessageWithScroll.text = ""
    txtMessage.Visible = True
    txtMessageWithScroll.Visible = False
    
    cmdPositive.Caption = "Positive"
    cmdNegative.Caption = "Negative"
    cmdCancel.Caption = "Cancel"
    
    cmdPositive.Enabled = True
    cmdNegative.Enabled = True
    cmdCancel.Enabled = True
    
    cmdPositive.Visible = True
    cmdNegative.Visible = True
    cmdCancel.Visible = True
    
    'Set UserResponse = -1 = Wait for response
    UserResponse = -1
    
    'Set everything to the default size and position
    With Me
    
        .Top = 4500
        .Left = 8640
        .Height = 3570
        .Width = 5010
        
    End With
    
    With txtMessage
    
        .Top = 120
        .Left = 120
        .Height = 2415
        .Width = 4695
        
    End With
    
    With cmdPositive
    
        .Top = 2640
        .Left = 120
        .Height = Me.TextHeight(.Caption) * 2
        .Width = Me.TextWidth(.Caption) * 2
        
    End With
    
    With cmdNegative
    
        .Top = 2640
        .Left = 1560
        .Height = Me.TextHeight(.Caption) * 2
        .Width = Me.TextWidth(.Caption) * 2
        
    End With
    
    With cmdCancel
    
        .Top = 2640
        .Left = 3600
        .Height = Me.TextHeight(.Caption) * 1.75
        .Width = Me.TextWidth(.Caption) * 1.5
        
    End With
    
End Sub

Private Sub mnuEditCopy_Click()

    If ActiveBox = "Scroll" Then
    
        Clipboard.Clear
        Clipboard.SetText Me.txtMessageWithScroll.text

    Else
    
        Clipboard.Clear
        Clipboard.SetText Me.txtMessage.text

    End If

End Sub

Public Sub ReSizeWindow()
                      
    Dim ButtonWidth As Single
    Dim LabelH As Single
    Dim LabelW As Single
    Dim TextH As Single
    Dim TextW As Single
    Dim MaxBtnW As Single
    Dim MaxBtnH As Single
    
    'Determine the window size based on the text width and heigth
    TextH = TextHeight(txtMessage.text)
    TextW = TextWidth(txtMessage.text)
    
    'Cap the width and height at a reasonable limit
    If TextW > Screen.Width - 2000 Then TextW = Screen.Width - 2000
    If TextH > Screen.Height - 2000 Then TextH = Screen.Height - 2000
    
    'Calculate the needed label dimensions
    LabelH = TextH + 1000
    LabelW = TextW + 1000
    
    'Set Those dimensions, and wrap the text
    txtMessage.Height = LabelH
    txtMessage.Width = LabelW
        
    'Set the form dimensions now
    Width = LabelW + 240
    Height = LabelH + 360 + cmdCancel.Height
    
    'Resize if bigger than screen
    If Width > Screen.Width - 1000 Then Width = Screen.Width - 2000
    If Height > Screen.Height - 1000 Then Height = Screen.Height - 2000
    
    'Center the form in the screen
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
        
    'Set the Button heights to the max height of the three
    MaxBtnH = 0
    If MaxBtnH < Me.TextHeight(cmdCancel.Caption) * 2 Then
    
        MaxBtnH = Me.TextHeight(cmdCancel.Caption) * 2
        
    End If
    
    If MaxBtnH < Me.TextHeight(cmdPositive.Caption) * 2 Then
    
        MaxBtnH = Me.TextHeight(cmdPositive.Caption) * 2
        
    End If
    
    If MaxBtnH < Me.TextHeight(cmdNegative.Caption) * 2 Then
    
        MaxBtnH = Me.TextHeight(cmdNegative.Caption) * 2
        
    End If
       
    cmdCancel.Height = MaxBtnH
    cmdPositive.Height = MaxBtnH
    cmdNegative.Height = MaxBtnH
        
    'Reset the label dimensions using the new ScaleHeight and ScaleWidth
    txtMessage.Height = ScaleHeight - 360 - MaxBtnH
    txtMessage.Width = ScaleWidth - 240
    
    'Determine if the vertical scroll bar needs to be turned on
    If TextHeight(txtMessage.text) > txtMessage.Height Then
    
        txtMessageWithScroll.Height = txtMessage.Height
        txtMessageWithScroll.Width = txtMessage.Width
        txtMessageWithScroll.text = txtMessage.text
        txtMessageWithScroll.Top = txtMessage.Top
        txtMessageWithScroll.Left = txtMessage.Left
        
        txtMessageWithScroll.Visible = True
        txtMessage.Visible = False
        
    End If
        
        
    'Set the Top positions of the buttons
    cmdCancel.Top = ScaleHeight - cmdCancel.Height - 120
    cmdPositive.Top = cmdCancel.Top
    cmdNegative.Top = cmdCancel.Top
    
    'Calculate the button width to use
    ButtonWidth = CSng(0.7 * (Width - 480) / 3)
    
    'Set the button widths to the max button width
    'Set the Button heights to the max height of the three
    MaxBtnW = ButtonWidth
    If MaxBtnW < Me.TextWidth(cmdCancel.Caption) * 2 Then
    
        MaxBtnW = Me.TextWidth(cmdCancel.Caption) * 2
        
    End If
    
    If MaxBtnW < Me.TextWidth(cmdPositive.Caption) * 2 Then
    
        MaxBtnW = Me.TextWidth(cmdPositive.Caption) * 2
        
    End If
    
    If MaxBtnW < Me.TextWidth(cmdNegative.Caption) * 2 Then
    
        MaxBtnW = Me.TextWidth(cmdNegative.Caption) * 2
        
    End If
       
    cmdCancel.Width = MaxBtnW
    cmdPositive.Width = MaxBtnW
    cmdNegative.Width = MaxBtnW
    
    'Set the button positions (All three buttons visible)
    If NumButtons = 3 Then
    
       cmdPositive.Visible = True
       cmdNegative.Visible = True
       cmdCancel.Visible = True
    
       cmdPositive.Left = 120
       cmdNegative.Left = CSng((ScaleWidth - cmdNegative.Width) / 2)
       cmdCancel.Left = ScaleWidth - cmdCancel.Width - 120
       
    'Only Two buttons visible (at either bottom corner)
    ElseIf NumButtons = 2 Then
           
        cmdPositive.Visible = True
        cmdNegative.Visible = True
        cmdCancel.Visible = False
       
        cmdPositive.Left = 120
        cmdNegative.Left = ScaleWidth - cmdNegative.Width - 120
        
    'Only One Button ( centered)
    ElseIf NumButtons = 1 Then
    
        cmdPositive.Visible = True
        cmdNegative.Visible = False
        cmdCancel.Visible = False
    
        cmdPositive.Left = CSng((ScaleWidth - cmdPositive.Width) / 2)
    
    End If
                   
End Sub

Private Sub txtMessage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
    
        ActiveBox = "NoScroll"
    
        'Popup the copy menu
        mnuEditCopy.Visible = True
        
        PopupMenu mnuEdit
        
    End If
        
End Sub

Private Sub txtMessageWithScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
    
        ActiveBox = "Scroll"
    
        'Popup the copy menu
        mnuEditCopy.Visible = True
        
        PopupMenu mnuEdit
        
    End If

End Sub

