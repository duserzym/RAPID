VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug messages"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3105
   ScaleWidth      =   4545
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   972
   End
   Begin VB.ListBox listDebug 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXLINEBUFFER = 200

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    
    listDebug.Clear
    Form_Resize
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
End Sub

Private Sub Form_Resize()
    cmdClose.Top = Me.ScaleHeight - 2 * cmdClose.Height - listDebug.Top
    listDebug.Height = cmdClose.Top - 2 * listDebug.Top
    listDebug.Width = Me.ScaleWidth - 2 * listDebug.Left
End Sub

Public Sub Msg(message As String)
    Dim curFracSec As Double
    If listDebug.ListCount > MAXLINEBUFFER Then listDebug.RemoveItem 0
    curFracSec = Timer - Int(Timer)
    listDebug.AddItem Format$(Now, "h:mm:ss") & Format$(curFracSec, ".00") & ": " & message
End Sub

