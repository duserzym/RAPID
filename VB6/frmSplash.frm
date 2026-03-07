VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3870
   ClientLeft      =   255
   ClientTop       =   1425
   ClientWidth     =   7560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3675
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7305
      Begin MSComDlg.CommonDialog dialogGetINIFile 
         Left            =   6840
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin ComctlLib.ProgressBar ProgressBar 
         Height          =   132
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   7092
         _ExtentX        =   12515
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Image Image1 
         Height          =   1215
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Image imgLogo 
         Height          =   1815
         Left            =   120
         Picture         =   "frmSplash.frx":677C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Paleomagnetic Magnetometer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1950
         TabIndex        =   1
         Top             =   360
         Width           =   5100
      End
      Begin VB.Label lblProductName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Control System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   671
         Left            =   2827
         TabIndex        =   2
         Top             =   847
         Width           =   4224
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2013"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5520
         TabIndex        =   3
         Top             =   1560
         Width           =   1020
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6315
         TabIndex        =   4
         Top             =   2100
         Width           =   675
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright (C) 2010 by the California Institute of Technology"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   2640
         Width           =   4815
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         Caption         =   "Licensed under the GNU General Public License"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   2880
         Width           =   4815
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   7092
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblStatus.Caption = "Initializing..."
    progress 0
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Public Sub progress(ByVal fraction As Single)
    ProgressBar.Value = fraction * 100
End Sub

Public Sub SplashStatus(StatusText As String)
    lblStatus.Caption = StatusText
    lblStatus.refresh
End Sub

