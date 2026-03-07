VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse for Directory..."
   ClientHeight    =   5748
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   5280
      Width           =   2172
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Select Current Directory"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   4800
      Width           =   2172
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
