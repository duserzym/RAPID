VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmINIConverter 
   Caption         =   "Paleomag INI File Upgrade Helper"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   11280
   Begin VB.TextBox txtStepNum 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   89
      Text            =   "Step __ of __"
      Top             =   360
      Width           =   1335
   End
   Begin ComctlLib.ProgressBar progStepsCompleted 
      Height          =   255
      Left            =   360
      TabIndex        =   88
      Top             =   600
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< Back"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdEndProgram 
      Caption         =   "Quit && Exit"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdNextFinishedSkip 
      Caption         =   "Next >>"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Frame Frame17 
      Height          =   6975
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   11055
      Begin VB.TextBox txtAFTransMaxVolt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   39
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox txtAFAxialMaxVolt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   38
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox txtAFTransResFreq 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   37
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtAFAxialResFreq 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label19 
         Caption         =   "Enter in the max pre-clipping voltage for the Transverse AF coil:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   35
         Top             =   4440
         Width           =   10575
      End
      Begin VB.Label Label17 
         Caption         =   "Enter in the resonance frequency for the Transverse AF coil:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   34
         Top             =   2280
         Width           =   10575
      End
      Begin VB.Label Label16 
         Caption         =   "Enter in the resonance frequency for the Axial AF coil:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   1320
         Width           =   10575
      End
      Begin VB.Label Label18 
         Caption         =   "Enter in the max pre-clipping voltage for the Axial AF coil: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   32
         Top             =   3480
         Width           =   10575
      End
   End
   Begin VB.Frame Frame16 
      Height          =   6975
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   11055
      Begin VB.Label Label15 
         Caption         =   $"frmINIConverter.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   30
         Top             =   3480
         Width           =   10575
      End
      Begin VB.Label Label14 
         Caption         =   $"frmINIConverter.frx":00B9
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   10575
      End
   End
   Begin VB.Frame Frame5 
      Height          =   6975
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   11055
      Begin MSComctlLib.ListView lvwEnabledModules 
         Height          =   4455
         Left            =   240
         TabIndex        =   60
         Top             =   1680
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7858
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label10 
         Caption         =   "Here's a list of all the rockmag modules.  Please check all the modules that you plan to use."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   10575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   6975
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   11055
      Begin VB.CheckBox chkADwinLight16 
         Caption         =   "ADwin - ADwin-light-16, USB external"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   18
         Top             =   3000
         Width           =   6735
      End
      Begin VB.CheckBox chkPCIDAS6030 
         Caption         =   "Measurement Computing - PCI-DAS6030"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   17
         Top             =   2040
         Width           =   7335
      End
      Begin VB.Label Label8 
         Caption         =   "The AF and IRM systems you have selected require the following DAQ Boards to be installed on or connected to your computer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   10455
      End
      Begin VB.Label Label9 
         Caption         =   "You will now be guided through the process of configuring the settings for the board(s) above."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   19
         Top             =   4080
         Width           =   10455
      End
   End
   Begin VB.Frame Frame2A 
      Height          =   6975
      Left            =   120
      TabIndex        =   51
      Top             =   0
      Width           =   11055
      Begin VB.ComboBox cmbIRMRelay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   59
         Top             =   5280
         Width           =   2655
      End
      Begin VB.ComboBox cmbAFTransRelay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   58
         Top             =   4200
         Width           =   2655
      End
      Begin VB.ComboBox cmbAFAxialRelay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   57
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label Label26 
         Caption         =   "IRM Relay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   56
         Top             =   4920
         Width           =   5535
      End
      Begin VB.Label Label25 
         Caption         =   "AF Tranverse Relay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   3840
         Width           =   5175
      End
      Begin VB.Label Label27 
         Caption         =   "AF Axial Relay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Top             =   2760
         Width           =   4935
      End
      Begin VB.Label Label5b 
         Caption         =   "Please Set the Digital Output Ports  for the three relays:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   53
         Top             =   2280
         Width           =   10335
      End
      Begin VB.Label Label5a 
         Caption         =   $"frmINIConverter.frx":0179
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   52
         Top             =   1080
         Width           =   10455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Would you like to continue with the INI  file upgrade process now?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   9
         Top             =   5280
         Width           =   10455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Your old Paleomag.INI file will not be deleted, modified, nor harmed (maybe taunted a few times....)."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   8
         Top             =   3120
         Width           =   10455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"frmINIConverter.frx":0233
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   10455
      End
   End
   Begin VB.Frame FrameWait 
      Height          =   6975
      Left            =   120
      TabIndex        =   86
      Top             =   0
      Width           =   11055
      Begin VB.Label LabelWait 
         Alignment       =   2  'Center
         Caption         =   "Please wait...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   87
         Top             =   3240
         Width           =   10575
      End
   End
   Begin VB.Frame Frame10 
      Height          =   6975
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   11055
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridDeletedSettings 
         Height          =   4455
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label13 
         Caption         =   $"frmINIConverter.frx":02E9
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   10575
      End
   End
   Begin VB.Frame Frame6 
      Height          =   6975
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   11055
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridBoardSettings 
         Height          =   4695
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   8281
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label11 
         Caption         =   "Board Name - DAQ Board settings list:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   10575
      End
   End
   Begin VB.Frame Frame8B 
      Height          =   6975
      Left            =   120
      TabIndex        =   68
      Top             =   0
      Width           =   11055
      Begin VB.ComboBox cmbIRMMonitorBoard 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   79
         Top             =   4560
         Width           =   3375
      End
      Begin VB.ComboBox cmbAltAFMonitorBoard 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   78
         Top             =   2280
         Width           =   3375
      End
      Begin VB.ComboBox cmbAltAFMonitorChan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   70
         Top             =   2760
         Width           =   2535
      End
      Begin VB.ComboBox cmbIRMMonitorChan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   69
         Top             =   5040
         Width           =   2535
      End
      Begin VB.Label Label37 
         Caption         =   "Channel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   77
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "Board:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   76
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "Channel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   75
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "Board:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   74
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "Please check the following AF / IRM Settings.  Fields are pre-loaded with default values:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   73
         Top             =   1080
         Width           =   10575
      End
      Begin VB.Label Label30 
         Caption         =   "Alternate AF Board && Channel for Ammeter (""Green Donut"") Input Voltage:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   72
         Top             =   1800
         Width           =   10455
      End
      Begin VB.Label Label29 
         Caption         =   "IRM Monitor Board && Channel for large Ammeter Input Voltage (may not be present on your system):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   71
         Top             =   3840
         Width           =   10575
      End
   End
   Begin VB.Frame Frame8A 
      Height          =   6975
      Left            =   120
      TabIndex        =   62
      Top             =   0
      Width           =   11055
      Begin VB.ComboBox cmbADWINMonitorChan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   67
         Top             =   4440
         Width           =   2895
      End
      Begin VB.ComboBox cmbADWINRampChan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   65
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label Label28 
         Caption         =   "AF Monitor Channel for Ammeter (""Green Donut"") Input Voltage:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   4080
         Width           =   10695
      End
      Begin VB.Label Label7 
         Caption         =   "AF Ramp Up / Ramp Down Channel for  ADWIN Output Volt. to Crest Audio Amplifier:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   64
         Top             =   2160
         Width           =   10695
      End
      Begin VB.Label Label5 
         Caption         =   "Please update the following ADWIN AF Settings.  Fields are pre-loaded with default values:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   63
         Top             =   1080
         Width           =   10695
      End
   End
   Begin VB.Frame Frame7 
      Height          =   6975
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   11055
      Begin VB.Label Label12 
         Caption         =   $"frmINIConverter.frx":0398
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   240
         TabIndex        =   24
         Top             =   2880
         Width           =   10575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   11055
      Begin VB.OptionButton optIRMCalDone 
         Caption         =   "Already calibrated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   84
         Top             =   4320
         Width           =   2895
      End
      Begin VB.OptionButton optIRMNeedsCalibration 
         Caption         =   "Needs to be calibrated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   83
         Top             =   3600
         Width           =   3255
      End
      Begin VB.ComboBox cmbIRMSystem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmINIConverter.frx":0497
         Left            =   240
         List            =   "frmINIConverter.frx":0499
         TabIndex        =   14
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label38 
         Caption         =   "Does this IRM System need to be calibrated?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   85
         Top             =   3240
         Width           =   10455
      End
      Begin VB.Label Label6 
         Caption         =   "Please select the IRM System that you are using or would like to start using:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   10455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11055
      Begin VB.OptionButton optAFCalDone 
         Caption         =   "Already calibrated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   81
         Top             =   4270
         Width           =   2895
      End
      Begin VB.OptionButton optAFNeedsCal 
         Caption         =   "Needs to be calibrated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   80
         Top             =   3600
         Width           =   3255
      End
      Begin VB.ComboBox cmbAFSystem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmINIConverter.frx":049B
         Left            =   240
         List            =   "frmINIConverter.frx":049D
         TabIndex        =   11
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label22 
         Caption         =   "Does this AF System need to be calibrated?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   82
         Top             =   3240
         Width           =   10455
      End
      Begin VB.Label Label4 
         Caption         =   "Please select the AF System that you are using or would like to start using:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   10455
      End
   End
   Begin VB.Frame Frame20 
      Height          =   6975
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Width           =   11055
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "Click 'Finish' to continue loading the Paleomag program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   50
         Top             =   4200
         Width           =   10575
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "The INI settings file upgrade is complete."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   49
         Top             =   2640
         Width           =   10575
      End
   End
   Begin VB.Frame Frame18 
      Height          =   6975
      Left            =   120
      TabIndex        =   40
      Top             =   0
      Width           =   11055
      Begin VB.CheckBox chkDoBackup 
         Caption         =   "Backup AF Data to remote folder?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   3600
         Width           =   9255
      End
      Begin VB.CommandButton cmdBrowseBackupFolder 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9960
         TabIndex        =   46
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtAFBackupDataFolder 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   240
         TabIndex        =   44
         Top             =   4440
         Width           =   9375
      End
      Begin VB.CommandButton cmdBrowseAFDataFolder 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9960
         TabIndex        =   43
         Top             =   1800
         Width           =   735
      End
      Begin MSComDlg.CommonDialog dlgBrowseAFDataFolder 
         Left            =   240
         Top             =   6120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtAFLocalDataFolder 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   240
         TabIndex        =   41
         Top             =   1800
         Width           =   9375
      End
      Begin VB.Label Label21 
         Caption         =   "Path to Backup AF Data folder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   45
         Top             =   4080
         Width           =   9375
      End
      Begin VB.Label Label20 
         Caption         =   "Path to AF Data folder (folder where AF monitor data will be saved for both 2G and ADWIN AF systems):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   42
         Top             =   1200
         Width           =   9375
      End
   End
   Begin VB.Frame Frame8 
      Height          =   6975
      Left            =   120
      TabIndex        =   90
      Top             =   0
      Width           =   11055
      Begin VB.OptionButton optTrimOnFalse 
         Caption         =   "False (Low-State)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   92
         Top             =   3000
         Width           =   2895
      End
      Begin VB.OptionButton optTrimOnTrue 
         Caption         =   "True (Hi-State)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   91
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label41 
         Caption         =   $"frmINIConverter.frx":049F
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   97
         Top             =   4920
         Width           =   9975
      End
      Begin VB.Label Label40 
         Caption         =   "This setting can be changed / corrected later in the Settings Window under the ""IRM"" tab."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   96
         Top             =   4320
         Width           =   9975
      End
      Begin VB.Label Label39 
         Caption         =   "** Default value for your IRM system."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   95
         Top             =   3600
         Width           =   7335
      End
      Begin VB.Label Label34 
         Caption         =   " If you get this setting wrong the IRM capacitor voltage will not be able to charge up properly."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   94
         Top             =   1800
         Width           =   9975
      End
      Begin VB.Label Label32 
         Caption         =   "What is the relay switch logical state at which the IRM Trim* turns on for your IRM hardware setup?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   93
         Top             =   1320
         Width           =   10455
      End
   End
   Begin VB.ComboBox cmbCellSelector 
      Height          =   315
      Left            =   3840
      TabIndex        =   61
      Top             =   6240
      Width           =   2655
   End
End
Attribute VB_Name = "frmINIConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ActiveFrame As String
Dim PriorFrame As String
Dim CurBoard As Board
Dim CurGrid As MSHFlexGrid
Dim ActiveCell(2) As Integer
Dim CurCmbIndex As Long
Dim Frame6Mode As String
Dim AlreadyHandled As Boolean

Public Enum TransferAxisType

    Transverse = 0
    Axial = 1

End Enum

Private Sub cmdBack_Click()

    Dim temp As Long

    'If the activeframe = "2" then set the prior frame = "0"
    If ActiveFrame = "2" Then PriorFrame = "0"
    
    'Set the active frame to the prior frame
    ActiveFrame = PriorFrame
    
    'Now call the cmdNextFinishedSkip button click event handling function
    cmdNextFinishedSkip_Click
    
End Sub

Private Sub cmdBrowseAFDataFolder_Click()

    Dim TempStr As String
    Dim TempL As Long
    
    'Parse App.Path to get the folder containing it
    TempStr = App.path
    
    'Clip off the last backslash if there is one
    If Right(TempStr, 1) = "\" Then TempStr = Mid(TempStr, 1, Len(TempStr) - 1)
    
    'Search for the next "\" at the end of the file name
    TempL = InStrRev(TempStr, "\")
    
    'We've found a matching backslash
    If TempL > 0 Then
    
        TempStr = Mid(TempStr, 1, TempL)
    
    Else
    
        'There isn't another backslash, just give the pathway to the folder that contains the VB6 project file
        TempStr = TempStr & "\"
        
    End If
    
    'Use Mod-filesave to select a directory for the AF Data folder
    Me.txtAFLocalDataFolder = modFileSave.OpenDir(TempStr, _
                                                  "Please select the folder to save the AF monitor data files.", _
                                                  Me)
                                
End Sub

Private Sub cmdBrowseBackupFolder_Click()

    Dim TempStr As String
    Dim TempL As Long
    
    'Parse App.Path to get the folder containing it
    TempStr = App.path
    
    'Clip off the last backslash if there is one
    If Right(TempStr, 1) = "\" Then TempStr = Mid(TempStr, 1, Len(TempStr) - 1)
    
    'Search for the next "\" at the end of the file name
    TempL = InStrRev(TempStr, "\")
    
    'We've found a matching backslash
    If TempL > 0 Then
    
        TempStr = Mid(TempStr, 1, TempL)
    
    Else
    
        'There isn't another backslash, just give the pathway to the folder that contains the VB6 project file
        TempStr = TempStr & "\"
        
    End If
    
    'Use Mod-filesave to select a directory for the AF Data folder
    Me.txtAFBackupDataFolder = modFileSave.OpenDir(TempStr, _
                                        "Please select the folder to backup the AF monitor data files.", _
                                        Me)
                                        
End Sub

Private Sub cmdEndProgram_Click()

    Dim fso As FileSystemObject

    'First, check to see if the new INI file has been created yet
    'If so, need to overwrite the application settings and restore the old .INI
    'file as the current one
    'Turn on error handling just in case
    On Error GoTo BadINIobj:
        
        If Not OldIniFile Is Nothing And _
           fso.FileExists(OldIniFile.filename) Then
        
            SaveSetting App.EXEName, "Settings", "INIFile", OldIniFile.filename
    
        End If

    On Error GoTo 0
    
BadINIobj:

    'Nothings been connected to yet.  Can just exit the program.
    End

End Sub

Private Sub cmdNextFinishedSkip_Click()

    Dim i, j As Integer
    Dim RowHeadAInChan As Long
    Dim RowHeadAOutChan As Long
    Dim RowHeadDInChan As Long
    Dim RowHeadDOutChan As Long
    Dim HeaderRows() As String
    
    Dim AFSection As String
    Dim CalDone As String
    
    Select Case ActiveFrame
    
        Case "0"
        
            'Reset buttons to how they are supposed to be on the first frame at the start of the
            'INI conversion process
            Me.cmdNextFinishedSkip.Caption = "< Start >"
            Me.cmdBack.Visible = False
    
            'Set the PriorFrame = "0"
            PriorFrame = "0"
            
            'Hide the progress bar
            Me.progStepsCompleted.Visible = False
    
            'Show the first frame
            ShowFrame "1"
    
        Case "1"
        
            'User has selected to proceed to frame 2
            '2 - Select AF system - new or old?
            
            'Show the progress bar
            Me.progStepsCompleted.Visible = True
                                    
            'Change the button names:
            Me.cmdNextFinishedSkip.Caption = "Next >>"
            
            'Show the back button
            Me.cmdBack.Visible = True
            
            'Set the PriorFrame = "0"
            PriorFrame = "0"
            
            'Show frame 2
            ShowFrame "2"
            
        Case "2"
        
            'User Has chosen to go to the frame after frame two
            
            'Set the prior frame = "1"
            PriorFrame = "1"
            
            'Check to see if we need to go to frame 2A before going to frame3
            
            'First verify that the necessary values on Frame 2 have been entered
            If Me.cmbAFSystem.text <> "ADWIN" And _
               Me.cmbAFSystem.text <> "2G" _
            Then
            
                'Pop-up message box to the user
                MsgBox "You must choose which AF system you are using or plan to use before proceding!", , _
                       "User Input Error!"
                       
                'Abort the button click event handler
                Exit Sub
            
            End If
            
            If (Me.optAFCalDone.Value = False And _
                Me.optAFNeedsCal.Value = False) _
            Then
        
                'Pop-up message box to the user
                MsgBox "You must choose whether or not your AF coil field calibrations have been done " & _
                       "before proceding!", , _
                       "User Input Error!"
                       
                'Abort the button click event handler
                Exit Sub
        
            End If
            
            
            'Is the AF system ADWIN?
            If Me.cmbAFSystem.text = "ADWIN" Then
            
                'Setup the combo-box channel selectors
                For i = 0 To 5
                
                    Me.cmbAFAxialRelay.AddItem "DIGOUT-" & Trim(Str(i)), i
                    Me.cmbAFTransRelay.AddItem "DIGOUT-" & Trim(Str(i)), i
                    Me.cmbIRMRelay.AddItem "DIGOUT-" & Trim(Str(i)), i
            
                Next i
            
                'Show frame 2A
                ShowFrame "2A"
                
            Else
            
                'Show frame 3
                ShowFrame "3"
                
            End If
            
        Case "2A"
        
            'User has selected to proceed to frame 3 to select which IRM system to use and whether or not
            'that IRM system is calibrated
            
            'Set the prior frame = "2"
            PriorFrame = "2"
            
            'First check to make sure that values have been entered in Frame 2A
            'All three relays need channel selections and the selections must each be a different channel
            If Me.cmbAFAxialRelay.text = "" Or _
               Me.cmbAFTransRelay.text = "" Or _
               Me.cmbIRMRelay.text = "" _
            Then
            
                'Pop-up a message for the user
                MsgBox "You must select Digital Ouput channel assignments for all three relays!", , _
                       "User Input Error!"
                       
                Exit Sub
           
            End If
            
            'Now check to make sure that all three are different
            If Me.cmbAFAxialRelay.text = Me.cmbAFTransRelay.text Or _
               Me.cmbAFAxialRelay.text = Me.cmbIRMRelay.text Or _
               Me.cmbIRMRelay.text = Me.cmbAFTransRelay.text _
            Then
            
                'two or more of the relays have the same digital channel assignments
                'Pop-up a message
                MsgBox "You must select different Digital Output channel assignments for each relay! " & _
                       vbNewLine & vbNewLine & "(No repeats allowed.  Sorry.... )", , _
                       "User Input Error!"
                       
                Exit Sub
                
            End If
           
            'Just need to show frame 3
            ShowFrame "3"
            
        Case "3"
        
            'Need to prepare frame "3A", which asks the user to set the logical state upon which the
            'IRM Trim is on or off.
            
            'Depending on the user chosen AF system, set the prior frame = "2" or = "2A"
            If Me.cmbAFSystem.text = "ADWIN" Then
            
                PriorFrame = "2A"
                
            Else
            
                PriorFrame = "2"
                
            End If
            
            'First check to make sure that the user made selections for the IRM system & the IRM calibration status
            If Me.cmbIRMSystem.text <> "Caltech Old" And _
               Me.cmbIRMSystem.text <> "ASC Scientific" _
            Then
            
                'Poop.  Pop-up a message for the user
                MsgBox "You must select which IRM system you are using or would like to use before proceding!", , _
                       "User Input Error!"
                       
                Exit Sub
                
            End If
            
            If Me.optIRMCalDone.Value = False And _
               Me.optIRMNeedsCalibration.Value = False _
            Then
            
                'Poop x 2.  Pop-up a message for the user
                MsgBox "You forgot to select the calibration status of your IRM system! " & _
                       "This info is needed before the INI Upgrade can procede any further.", , _
                       "User Input Error!"
                       
                Exit Sub
                
            End If
            
            
            
            'Change the labels on the Option buttons to indicate what the default value is for the users
            'selected IRM system
            'Default Values:
            '   Matsusada - Trim On False
            '   Old -   Trim On True
            If Me.cmbIRMSystem.text = "Caltech Old" Then
            
                Me.optTrimOnTrue.Caption = "True (Hi-state)**"
                Me.optTrimOnFalse.Caption = "False (Low-state)"
                
            Else
            
                Me.optTrimOnTrue.Caption = "True (Hi-state)"
                Me.optTrimOnFalse.Caption = "False (Low-state)**"
                
            End If
                      
            
                      
            'Show this frame
            ShowFrame "3A"
            
        Case "3A"
            
            'Set the prior frame = "3"
            PriorFrame = "3"
            
            'Check to see if the user selected a logic-state for the IRM Trim On setting
            If Me.optTrimOnTrue.Value = False And _
               Me.optTrimOnFalse.Value = False _
            Then
            
                'Ugh.  Pop-up a message box
                MsgBox "You forgot to select the logic state under which the IRM Trim is active. " & _
                       "The INI upgrade cannot continue without this information.", , _
                       "User Input Error"
                       
                Exit Sub
            
            End If
            
            'Need to prepare Frame 4 based on the AF system being used
            If Me.cmbAFSystem.text = "ADWIN" Then
            
                Me.chkADwinLight16.Value = Checked
                
            Else
            
                Me.chkADwinLight16.Value = Unchecked
                
            End If
            
            Me.chkPCIDAS6030.Value = Checked
            
            'Change the next button to a "< Continue >" button
            Me.cmdNextFinishedSkip.Caption = "< Continue >"
            
            'Show frame 4
            ShowFrame "4"
            
        Case "4"
        
            'Need to prepare the next frame quite a bit
            
            'Actually need to start looking at the settings in the old INI file
            'Need to populate the list view with all the possible
            'paleomag code modules, and check the modules that
            'are already enabled in the old .INI file
            
            'Set the prior frame = "3A"
            PriorFrame = "3A"
            
            'Clear the list view
            Me.lvwEnabledModules.ListItems.Clear
            Me.lvwEnabledModules.ColumnHeaders.Clear
            
            'Turn on the check-boxes in the list view
            Me.lvwEnabledModules.Checkboxes = True
            
            With Me.lvwEnabledModules
                
                'Set view style to report to show all of the columns
                .View = lvwReport
                
                'Set the columns in the list view
                .ColumnHeaders.Add 1, , "Module", .Width / 5
                .ColumnHeaders.Add 2, , "Description", .Width / 5 * 4
                                
                'Need to add Modules to the list view
                With .ListItems
                
                    .Add 1, "EnableAxialIRM", "Axial IRM"
                    .Item(1).SubItems(1) = "Enables IRM with the Axial Coil"
                                       
                    'Check to see if Axial IRM is enabled in the old INI file
                    If Trim(OldIniFile.EntryRead("EnableIRM", "ERROR", "Modules")) = "True" Then
                    
                        .Item(1).Checked = True
                        
                    Else
                    
                        .Item(1).Checked = False
                        
                    End If
                                       
                    .Add 2, "EnableTransIRM", "Transverse IRM"
                    .Item(2).SubItems(1) = "Enables IRM with the Transverse Coil"
                    
                    .Item(2).Checked = False
                    .Item(2).Ghosted = True
                    .Item(2).ToolTipText = "Under development. ADWIN AF system only."
                    
                    .Add 3, "EnableIRMBackfield", "IRM Backfield"
                    .Item(3).SubItems(1) = "Enables DC Demag experiments"
                    
                    'Check to see if IRM Backfield is enabled in the old INI file
                    If Trim(OldIniFile.EntryRead("EnableIRMBackfield", "ERROR", "Modules")) = "True" Then
                    
                        .Item(3).Checked = True
                        
                    Else
                    
                        .Item(3).Checked = False
                        
                    End If
                    
                    .Add 4, "EnableIRMMonitor", "IRM Monitor"
                    .Item(4).SubItems(1) = "Enables real-time monitoring of coil " & _
                                           "circuit voltage during IRM pulse"
                                          
                    .Item(4).Checked = False
                    .Item(4).Ghosted = True
                    .Item(4).ToolTipText = "Under development"
                                        
                    .Add 5, , "-----------"
                    .Item(5).SubItems(1) = "-----------"
                    .Item(5).Ghosted = True
                    
                    .Add 6, "EnableAF", "AF"
                    .Item(6).SubItems(1) = "Enabled Axial and Transverse AF demag"
                    
                    'Check to see if AF demag is enabled in the old INI file
                    If Trim(OldIniFile.EntryRead("EnableAF", "ERROR", "Modules")) = "True" Then
                    
                        .Item(6).Checked = True
                        
                    Else
                    
                        .Item(6).Checked = False
                        
                    End If
                    
                    .Add 7, "EnableAFAnalysis", "AF Analysis"
                    .Item(7).SubItems(1) = "Real-time display of AF monitor data. " & _
                                           "ADWIN AF system only."
                    
                    .Item(7).Checked = False
                                            
                    If Me.cmbAFSystem.text = "2G" Then
                    
                        .Item(7).Ghosted = True
                        .Item(7).ToolTipText = "ADWIN AF system only."
                        
                    Else
                        
                        .Item(7).Ghosted = False
                        .Item(7).ToolTipText = "ADWIN AF system only."
                        
                    End If
                    
                    .Add 8, "EnableAltAFMonitor", "Alt. AF Monitor"
                    .Item(8).SubItems(1) = "Alternate monitor channel for AF circuit. " & _
                                           "Works with both AF systems. " & _
                                           "Requires hardware modification."
                                           
                    .Item(8).Checked = False
                                           
                    .Add 9, , "-----------"
                    .Item(9).SubItems(1) = "-----------"
                    .Item(9).Ghosted = True
                    
                    .Add 10, "EnableARM", "ARM"
                    .Item(10).SubItems(1) = "Enables ARM acquisition & demag experiments " & _
                                           "(if the AF module is also activated)."
                                           
                    'Check to see if ARM is enabled in the old INI file
                    If Trim(OldIniFile.EntryRead("EnableARM", "ERROR", "Modules")) = "True" Then
                    
                        .Item(10).Checked = True
                        
                    Else
                    
                        .Item(10).Checked = False
                        
                    End If
                                           
                    .Add 11, , "-----------"
                    .Item(11).SubItems(1) = "-----------"
                    .Item(11).Ghosted = True
                    
                    .Add 12, "EnableSusceptibility", "Susceptibility"
                    .Item(12).SubItems(1) = "Enables the Bartington Susceptibility bridge"
                    
                    'Check to see if the Susceptibility bridge is enabled in the old INI file
                    If Trim(OldIniFile.EntryRead("EnableSusceptibility", "ERROR", "Modules")) = "True" Then
                    
                        .Item(12).Checked = True
                        
                    Else
                    
                        .Item(12).Checked = False
                        
                    End If
                    
                    .Add 13, , "-----------"
                    .Item(13).SubItems(1) = "-----------"
                    .Item(13).Ghosted = True
                    
                    .Add 14, "EnableT1", "AF Thermal Sensor #1"
                    .Item(14).SubItems(1) = "Enables software for monitoring of AF coil " & _
                                            "temperature.  Requires hardware modification."
                                            
                    If Trim(OldIniFile.EntryRead("EnableT1", "ERROR", "Modules")) = "True" Then
                    
                        .Item(14).Checked = True
                        
                    Else
                    
                        .Item(14).Checked = False
                        
                    End If
                    
                    .Add 15, "EnableT2", "AF Thermal Sensor #2"
                    .Item(15).SubItems(1) = "Enables software for monitoring of AF coil " & _
                                            "temperature.  Requires hardware modification."
                                            
                    If Trim(OldIniFile.EntryRead("EnableT2", "ERROR", "Modules")) = "True" Then
                    
                        .Item(15).Checked = True
                        
                    Else
                    
                        .Item(15).Checked = False
                        
                    End If
                                            
                End With
                                        
            End With
            
            'Change the Next button back to "Next >>"
            Me.cmdNextFinishedSkip.Caption = "Next >>"
                
            'Show Frame 5
            ShowFrame "5"
            
        Case "5"
        
            'First need to show frame 6 with the PCI Board settings that will be used
                        
            'Set the prior frame = "4"
            PriorFrame = "4"
            
            'Create the new Board settings in the Paleomag.INI file using the Defaults.INI file
            modConfig.Create_BoardsForINI
            
            'Load the the new INI boards settings into the System Boards collection
            modConfig.Get_BoardsFromIni
            
            'Set the Label
            Label11.Caption = "PCI-DAS6030 - List of DAQ Board Settings:"
            
            With Me.gridBoardSettings
            
                'Clear the grid
                .ClearStructure
                .Clear
                
                'Set the number of rows, cols, and fixed rows & cols
                '# Rows = 15 nonchannel settings rows,
                '         12 rows for channel section headers,
                '         rows for channel data
                .Rows = 15 + 13 + _
                        SystemBoards("PCI-DAS6030").AInChannels.Count + _
                        SystemBoards("PCI-DAS6030").AOutChannels.Count + _
                        SystemBoards("PCI-DAS6030").DInChannels.Count + _
                        SystemBoards("PCI-DAS6030").DOutChannels.Count
                        
                .Cols = 3
                .FixedRows = 0
                .FixedCols = 0
                
                'Set the Column Headers
                .TextMatrix(0, 0) = "Board Setting"
                .TextMatrix(0, 1) = "Setting Value"
                .TextMatrix(0, 2) = "Description"
                
            End With
             
            'Populate the Board Global Variables from the INI file
            modConfig.Get_BoardsFromIni
             
            With SystemBoards("PCI-DAS6030")
                 
                 'Non Channel Board Settings
                 Me.gridBoardSettings.TextMatrix(1, 0) = "BoardININum"
                 Me.gridBoardSettings.TextMatrix(1, 1) = Trim(Str(.BoardININum))
                 Me.gridBoardSettings.TextMatrix(1, 2) = "INI File ID # for DAQ Board"
                 Me.gridBoardSettings.TextMatrix(2, 0) = "BoardNum"
                 Me.gridBoardSettings.TextMatrix(2, 1) = Trim(Str(.BoardNum))
                 Me.gridBoardSettings.TextMatrix(2, 2) = "MCC or ADwin ID # for DAQ Board"
                 Me.gridBoardSettings.TextMatrix(3, 0) = "BoardName"
                 Me.gridBoardSettings.TextMatrix(3, 1) = .BoardName
                 Me.gridBoardSettings.TextMatrix(3, 2) = "Unique String ID in Paleomag code for DAQ Board"
                 Me.gridBoardSettings.TextMatrix(4, 0) = "BoardFunction"
                 Me.gridBoardSettings.TextMatrix(4, 1) = .BoardFunction
                 Me.gridBoardSettings.TextMatrix(4, 2) = "How the Board can be used in Paleomag code"
                 Me.gridBoardSettings.TextMatrix(5, 0) = "CommProtocol"
                 Me.gridBoardSettings.TextMatrix(5, 1) = Trim(Str(.CommProtocol))
                 Me.gridBoardSettings.TextMatrix(5, 2) = "MCC(1) vs ADwin(2) comm protocol"
                 Me.gridBoardSettings.TextMatrix(6, 0) = "BoardMode"
                 Me.gridBoardSettings.TextMatrix(6, 1) = Trim(Str(.BoardMode))
                 Me.gridBoardSettings.TextMatrix(6, 2) = "Single(1) wire vs Differential(0), 2 wire Analog Input mode"
                 Me.gridBoardSettings.TextMatrix(7, 0) = "MaxAInRate"
                 Me.gridBoardSettings.TextMatrix(7, 1) = Trim(Str(.MaxAInRate))
                 Me.gridBoardSettings.TextMatrix(7, 2) = "Max Analog Input A/D rate (Hz) for DAQ Board"
                 Me.gridBoardSettings.TextMatrix(8, 0) = "MaxAOutRate"
                 Me.gridBoardSettings.TextMatrix(8, 1) = Trim(Str(.MaxAOutRate))
                 Me.gridBoardSettings.TextMatrix(8, 2) = "Max Analog Output D/A rate (Hz) for DAQ Board"
                 Me.gridBoardSettings.TextMatrix(9, 0) = "RangeType"
                 Me.gridBoardSettings.TextMatrix(9, 1) = Trim(Str(.range.RangeType))
                 Me.gridBoardSettings.TextMatrix(9, 2) = "Board Range Constant, MCC Boards only"
                 Me.gridBoardSettings.TextMatrix(10, 0) = "AInChannelsCount"
                 Me.gridBoardSettings.TextMatrix(10, 1) = Trim(Str(.AInChannels.Count))
                 Me.gridBoardSettings.TextMatrix(10, 2) = "# of Analog Input Channels the Board has"
                 Me.gridBoardSettings.TextMatrix(11, 0) = "AOutChannelsCount"
                 Me.gridBoardSettings.TextMatrix(11, 1) = Trim(Str(.AOutChannels.Count))
                 Me.gridBoardSettings.TextMatrix(11, 2) = "# of Analog Output Channels the Board has"
                 Me.gridBoardSettings.TextMatrix(12, 0) = "DInChannelsCount"
                 Me.gridBoardSettings.TextMatrix(12, 1) = Trim(Str(.DInChannels.Count))
                 Me.gridBoardSettings.TextMatrix(12, 2) = "# of Digital Input Channels the Board has"
                 Me.gridBoardSettings.TextMatrix(13, 0) = "DOutChannelsCount"
                 Me.gridBoardSettings.TextMatrix(13, 1) = Trim(Str(.DOutChannels.Count))
                 Me.gridBoardSettings.TextMatrix(13, 2) = "# of Digital Output Channels the Board has"
                 Me.gridBoardSettings.TextMatrix(14, 0) = "DIOConfigured"
                 Me.gridBoardSettings.TextMatrix(14, 1) = Trim(Str(.DIOConfigured))
                 Me.gridBoardSettings.TextMatrix(14, 2) = "Are DIO channels pre-configured as Input or Output?"
                 Me.gridBoardSettings.TextMatrix(15, 0) = "DOutPortType"
                 Me.gridBoardSettings.TextMatrix(15, 1) = Trim(Str(.DoutPortType))
                 Me.gridBoardSettings.TextMatrix(15, 2) = "Bit(1) vs Port(>=10) configured DIO, MCC Boards only"
                 
                 'Store # of input and output channels to local variables
                 RowHeadAInChan = 17
                 RowHeadAOutChan = RowHeadAInChan + .AInChannels.Count + 3
                 RowHeadDInChan = RowHeadAOutChan + .AOutChannels.Count + 3
                 RowHeadDOutChan = RowHeadDInChan + .DInChannels.Count + 3
                    
            End With
                                
            With Me.gridBoardSettings
                    
                'Merge necessary rows for headers to each channel section
                .MergeCells = flexMergeRestrictRows
                .MergeRow(RowHeadAInChan - 1) = True
                .MergeRow(RowHeadAInChan) = True
                .MergeRow(RowHeadAOutChan - 1) = True
                .MergeRow(RowHeadAOutChan) = True
                .MergeRow(RowHeadDInChan - 1) = True
                .MergeRow(RowHeadDInChan) = True
                .MergeRow(RowHeadDOutChan - 1) = True
                .MergeRow(RowHeadDOutChan) = True
                
                'Make sure content in the cells in each of these rows are all the same
                For i = 0 To .Cols - 1
                
                    'Want a blank row before the header for each channel section
                    'with the same background cell color of the unpopulated section
                    'of the grid
                    .row = RowHeadAInChan - 1
                    .Col = i
                    .text = "     "
                    .CellBackColor = &HC0C0C0
                    
                    .row = RowHeadAOutChan - 1
                    .Col = i
                    .text = "     "
                    .CellBackColor = &HC0C0C0
                    
                    .row = RowHeadDInChan - 1
                    .Col = i
                    .text = "     "
                    .CellBackColor = &HC0C0C0
                    
                    .row = RowHeadDOutChan - 1
                    .Col = i
                    .text = "     "
                    .CellBackColor = &HC0C0C0
                    
                    'Now want channel section merged header with the name of the
                    'Channel Section + background color of fixed row (Row 0)
                    .row = RowHeadAInChan
                    .Col = i
                    .text = "Analog Input Channels"
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadAOutChan
                    .Col = i
                    .text = "Analog Output Channels"
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadDInChan
                    .Col = i
                    .text = "Digital Input Channels"
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadDOutChan
                    .Col = i
                    .text = "Digital Output Channles"
                    .CellBackColor = &H8000000F
                    
                    'Now set the cell background colors for the second header row
                    'to the fixed row background color
                    .row = 0
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadAInChan + 1
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadAOutChan + 1
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadDInChan + 1
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadDOutChan + 1
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                                                
                Next i
                
                'Now need to populate second header rows
                'By brute force coding
                Me.gridBoardSettings.TextMatrix(RowHeadAInChan + 1, 0) = "Chan INI ID"
                Me.gridBoardSettings.TextMatrix(RowHeadAInChan + 1, 1) = "Chan Name"
                Me.gridBoardSettings.TextMatrix(RowHeadAInChan + 1, 2) = "Chan Number"
                Me.gridBoardSettings.TextMatrix(RowHeadAOutChan + 1, 0) = "Chan INI ID"
                Me.gridBoardSettings.TextMatrix(RowHeadAOutChan + 1, 1) = "Chan Name"
                Me.gridBoardSettings.TextMatrix(RowHeadAOutChan + 1, 2) = "Chan Number"
                Me.gridBoardSettings.TextMatrix(RowHeadDInChan + 1, 0) = "Chan INI ID"
                Me.gridBoardSettings.TextMatrix(RowHeadDInChan + 1, 1) = "Chan Name"
                Me.gridBoardSettings.TextMatrix(RowHeadDInChan + 1, 2) = "Chan Number"
                Me.gridBoardSettings.TextMatrix(RowHeadDOutChan + 1, 0) = "Chan INI ID"
                Me.gridBoardSettings.TextMatrix(RowHeadDOutChan + 1, 1) = "Chan Name"
                Me.gridBoardSettings.TextMatrix(RowHeadDOutChan + 1, 2) = "Chan Number"
                
            End With
            
            With SystemBoards("PCI-DAS6030")
                
                'Now need to populate the data in the Analog In channel section
                For i = 1 To .AInChannels.Count
                
                    'Col 0 = the Channel's INI ID string
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAInChan + 1, 0) = _
                                modConfig.Create_INIChanStr(.AInChannels(i))
                                
                    'Col 1 = the Channel Name
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAInChan + 1, 1) = .AInChannels(i).ChanName
                    
                    'Col 2 = the Channel Number
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAInChan + 1, 2) = Trim(Str(.AInChannels(i).ChanNum))
                    
                Next i
                
                'Now need to populate the data in the Analog Out channel section
                For i = 1 To .AOutChannels.Count
                
                    'Col 0 = the Channel's INI ID string
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAOutChan + 1, 0) = _
                                modConfig.Create_INIChanStr(.AOutChannels(i))
                                
                    'Col 1 = the Channel Name
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAOutChan + 1, 1) = .AOutChannels(i).ChanName
                    
                    'Col 2 = the Channel Number
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAOutChan + 1, 2) = Trim(Str(.AOutChannels(i).ChanNum))
                    
                Next i
                
                'Now need to populate the data in the Dig. In channel section
                For i = 1 To .DInChannels.Count
                
                    'Col 0 = the Channel's INI ID string
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDInChan + 1, 0) = _
                                modConfig.Create_INIChanStr(.DInChannels(i))
                                
                    'Col 1 = the Channel Name
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDInChan + 1, 1) = .DInChannels(i).ChanName
                    
                    'Col 2 = the Channel Number
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDInChan + 1, 2) = Trim(Str(.DInChannels(i).ChanNum))
                    
                Next i
                                    
                'Now need to populate the data in the Dig. Out channel section
                For i = 1 To .DOutChannels.Count
                
                    'Col 0 = the Channel's INI ID string
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDOutChan + 1, 0) = _
                                modConfig.Create_INIChanStr(.DOutChannels(i))
                                
                    'Col 1 = the Channel Name
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDOutChan + 1, 1) = .DOutChannels(i).ChanName
                    
                    'Col 2 = the Channel Number
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDOutChan + 1, 2) = Trim(Str(.DOutChannels(i).ChanNum))
                    
                Next i
                                    
            End With
            
            'Change the column alignment for the 2nd & 3rd columns to Left justified
            Me.gridBoardSettings.ColAlignment(1) = flexAlignLeftCenter
            Me.gridBoardSettings.ColAlignment(2) = flexAlignLeftCenter
            
            'Set the top visible row
            Me.gridBoardSettings.TopRow = 0
            
            'Resize the board settings grid
            ResizeGrid Me.gridBoardSettings, _
                       Me, _
                       0, _
                       Me.gridBoardSettings.Rows - 1, , , , , _
                       False
            
            'Make the frame visible
            ShowFrame "6"
    
        Case "6"
        
            'Set the prior frame = "5"
            PriorFrame = "5"
        
            'Change the label on Frame 6B
            Me.Label11.Caption = "PCI-DAS6030 Board - The following pre-existing port assignments have " & _
                                "been identified in your Paleomag.ini file:"
                                
            'Need to read the comm-port settings that were in the old INI file
            Me.gridBoardSettings.Rows = 11
            Me.gridBoardSettings.FixedRows = 1
            Me.gridBoardSettings.Cols = 5
            
            With Me.gridBoardSettings
            
                'Set the column headers
                .TextMatrix(0, 0) = "Channel Description"
                .TextMatrix(0, 1) = "Old Chan. #"
                .TextMatrix(0, 2) = "Chan. Type"
                .TextMatrix(0, 3) = "Board INI #"
                .TextMatrix(0, 4) = "New Chan. ID"
                
                'Get the ARMVOutNum from the OldINIfile
                modConfig.ARMVOutNo = OldIniFile.EntryRead("ARMVoltageOut", _
                                                           "0", _
                                                           "IRM-ARM")
                
                .TextMatrix(1, 0) = "ARM Level Volts Out"
                .TextMatrix(1, 1) = Trim(Str(ARMVOutNo))
                .TextMatrix(1, 2) = "Analog Out"
                .TextMatrix(1, 3) = "0"
                .TextMatrix(1, 4) = "AO-0-CH" & Trim(Str(ARMVOutNo))
                
                'Add this channel to the INI file
                modConfig.Config_SaveSetting "Channels", _
                                             "ARMVoltageOut", _
                                             "AO-0-CH" & Trim(Str(ARMVOutNo))
                
                'Get the IRMVOutNum from the OldINIfile
                modConfig.IRMVOutNo = OldIniFile.EntryRead("IRMVoltageOut", _
                                                           "1", _
                                                           "IRM-ARM")
                
                .TextMatrix(2, 0) = "IRM Level Volts Out"
                .TextMatrix(2, 1) = Trim(Str(IRMVOutNo))
                .TextMatrix(2, 2) = "Analog Out"
                .TextMatrix(2, 3) = "0"
                .TextMatrix(2, 4) = "AO-0-CH" & Trim(Str(IRMVOutNo))
                
                'Add this channel to the INI file
                modConfig.Config_SaveSetting "Channels", _
                                             "IRMVoltageOut", _
                                             "AO-0-CH" & Trim(Str(IRMVOutNo))
                
                'Get the IRMCapVInNo from the OldINIfile
                modConfig.IRMCapVInNo = OldIniFile.EntryRead("IRMCapacitorVoltageIn", _
                                                             "0", _
                                                             "IRM-ARM")
                
                .TextMatrix(3, 0) = "IRM Capacitor Volts In"
                .TextMatrix(3, 1) = Trim(Str(IRMCapVInNo))
                .TextMatrix(3, 2) = "Analog In"
                .TextMatrix(3, 3) = "0"
                .TextMatrix(3, 4) = "AI-0-CH" & Trim(Str(IRMCapVInNo))
                
                'Add this channel to the INI file
                modConfig.Config_SaveSetting "Channels", _
                                             "IRMCapacitorVoltageIn", _
                                             "AI-0-CH" & Trim(Str(IRMCapVInNo))
                
                'Get the ARMSetNo from the OldINIfile
                modConfig.ARMSetNo = OldIniFile.EntryRead("ARMSet", _
                                                          "0", _
                                                          "IRM-ARM")
                
                .TextMatrix(4, 0) = "ARM Set TTL"
                .TextMatrix(4, 1) = Trim(Str(ARMSetNo))
                .TextMatrix(4, 2) = "Dig. Out"
                .TextMatrix(4, 3) = "0"
                .TextMatrix(4, 4) = "DO-0-CH" & Trim(Str(ARMSetNo))
                
                'Add this channel to the INI file
                modConfig.Config_SaveSetting "Channels", _
                                             "ARMSet", _
                                             "DO-0-CH" & Trim(Str(ARMSetNo))
                
                'Get the IRMFireNo from the OldINIfile
                modConfig.IRMFireNo = OldIniFile.EntryRead("IRMFire", _
                                                          "1", _
                                                          "IRM-ARM")
                                
                .TextMatrix(5, 0) = "IRM Fire TTL"
                .TextMatrix(5, 1) = Trim(Str(IRMFireNo))
                .TextMatrix(5, 2) = "Dig. Out"
                .TextMatrix(5, 3) = "0"
                .TextMatrix(5, 4) = "DO-0-CH" & Trim(Str(IRMFireNo))
                
                'Add this channel to the INI file
                modConfig.Config_SaveSetting "Channels", _
                                             "IRMFire", _
                                             "DO-0-CH" & Trim(Str(IRMFireNo))
                
                
                'Get the IRMTrimNo from the OldINIfile
                modConfig.IRMTrimNo = OldIniFile.EntryRead("IRMTrim", _
                                                           "3", _
                                                           "IRM-ARM")
                
                .TextMatrix(6, 0) = "IRM Trim TTL"
                .TextMatrix(6, 1) = Trim(Str(IRMTrimNo))
                .TextMatrix(6, 2) = "Dig. Out"
                .TextMatrix(6, 3) = "0"
                .TextMatrix(6, 4) = "DO-0-CH" & Trim(Str(IRMTrimNo))
         
                'Add this channel to the INI file
                modConfig.Config_SaveSetting "Channels", _
                                             "IRMTrim", _
                                             "DO-0-CH" & Trim(Str(IRMTrimNo))
                               
                
                SetRowColor Me.gridBoardSettings, 7, QBColor(8)
                .TextMatrix(7, 0) = "IRM Ready Indicator"
                .TextMatrix(7, 1) = Trim(Str(IRMReadyNo))
                .TextMatrix(7, 2) = "Dig. In"
                .TextMatrix(7, 3) = "NA"
                .TextMatrix(7, 4) = "Obsolete"
                
                'Not adding this channel.
                
                'Get the MotorToggleNo from the OldINIfile
                modConfig.MotorToggleNo = OldIniFile.EntryRead("MotorToggle", _
                                                               "5", _
                                                               "IRM-ARM")
                
                .TextMatrix(8, 0) = "Motor On/Off TTL"
                .TextMatrix(8, 1) = Trim(Str(MotorToggleNo))
                .TextMatrix(8, 2) = "Dig. Out"
                .TextMatrix(8, 3) = "0"
                .TextMatrix(8, 4) = "DO-0-CH" & Trim(Str(MotorToggleNo))
                
                
                'Add this channel to the INI file
                modConfig.Config_SaveSetting "Channels", _
                                             "MotorToggle", _
                                             "DO-0-CH" & Trim(Str(MotorToggleNo))
                
                'Get the VacTogANo from the OldINIfile
                modConfig.VacTogANo = OldIniFile.EntryRead("VacuumToggleA", _
                                                           "6", _
                                                           "IRM-ARM")
                
                .TextMatrix(9, 0) = "Vacuum Toggle A TTL"
                .TextMatrix(9, 1) = Trim(Str(VacTogANo))
                .TextMatrix(9, 2) = "Dig. Out"
                .TextMatrix(9, 3) = "0"
                .TextMatrix(9, 4) = "DO-0-CH" & Trim(Str(VacTogANo))
                
                'Add this channel to the INI file
                modConfig.Config_SaveSetting "Channels", _
                                             "VacuumToggleA", _
                                             "DO-0-CH" & Trim(Str(VacTogANo))
                
                
                'Get the VacTogBNo from the OldINIfile
                modConfig.VacTogBNo = OldIniFile.EntryRead("VacuumToggleB", _
                                                           "7", _
                                                           "IRM-ARM")
                       
                .TextMatrix(10, 0) = "Vacuum Toggle B TTL"
                .TextMatrix(10, 1) = Trim(Str(VacTogBNo))
                .TextMatrix(10, 2) = "Dig. Out"
                .TextMatrix(10, 3) = "0"
                .TextMatrix(10, 4) = "DO-0-CH" & Trim(Str(VacTogBNo))
                
                'Add this channel to the INI file
                modConfig.Config_SaveSetting "Channels", _
                                             "VacuumToggleB", _
                                             "DO-0-CH" & Trim(Str(VacTogBNo))
                                             
                'Get the DegausserToggle from the OldINIfile
                modConfig.DegCoolNo = OldIniFile.EntryRead("DegausserToggle", _
                                                           "2", _
                                                           "IRM-ARM")
                       
                .TextMatrix(11, 0) = "Degausser Cooler TTL"
                .TextMatrix(11, 1) = Trim(Str(DegCoolNo))
                .TextMatrix(11, 2) = "Dig. Out"
                .TextMatrix(11, 3) = "0"
                .TextMatrix(11, 4) = "DO-0-CH" & Trim(Str(DegCoolNo))
                
                'Add this channel to the INI file
                modConfig.Config_SaveSetting "Channels", _
                                             "DegausserToggle", _
                                             "DO-0-CH" & Trim(Str(DegCoolNo))
                
                
                ResizeGrid Me.gridBoardSettings, _
                           Me, _
                           0, _
                           Me.gridBoardSettings.Rows - 1, , , , , _
                           False
                
            End With
            
            'Now ready to show frame 6 again, but as frame 6B
            ShowFrame "6B"
            
        Case "6B"
            
            
            'Set the prior frame = "6"
            PriorFrame = "6"
            
            'Just need to show frame "7" - is just a message
            ShowFrame "7"
            
        Case "7"
    
            'Set the prior frame = "6B"
            PriorFrame = "6B"
        
            'Need to prepare Board settings grid to show the settings for the
            'ADWIN board
            
            'Set the Label
            Label11.Caption = "ADWin-light-16 - List of DAQ Board Settings:"
            
            With Me.gridBoardSettings
            
                'Clear the grid
                .ClearStructure
                .Clear
                
                'Set the number of rows, cols, and fixed rows & cols
                '# Rows = 15 nonchannel settings rows,
                '         12 rows for channel section headers,
                '         rows for channel data
                .Rows = 15 + 13 + _
                        SystemBoards("ADWIN-light-16").AInChannels.Count + _
                        SystemBoards("ADWIN-light-16").AOutChannels.Count + _
                        SystemBoards("ADWIN-light-16").DInChannels.Count + _
                        SystemBoards("ADWIN-light-16").DOutChannels.Count
                        
                .Cols = 3
                .FixedRows = 0
                .FixedCols = 0
                
                'Set the Column Headers
                .TextMatrix(0, 0) = "Board Setting"
                .TextMatrix(0, 1) = "Setting Value"
                .TextMatrix(0, 2) = "Description"
                
            End With
             
            'Populate the Board Global Variables from the INI file
            modConfig.Get_BoardsFromIni
             
            With SystemBoards("ADWIN-light-16")
                 
                 'Non Channel Board Settings
                 Me.gridBoardSettings.TextMatrix(1, 0) = "BoardININum"
                 Me.gridBoardSettings.TextMatrix(1, 1) = Trim(Str(.BoardININum))
                 Me.gridBoardSettings.TextMatrix(1, 2) = "INI File ID # for DAQ Board"
                 Me.gridBoardSettings.TextMatrix(2, 0) = "BoardNum"
                 Me.gridBoardSettings.TextMatrix(2, 1) = Trim(Str(.BoardNum))
                 Me.gridBoardSettings.TextMatrix(2, 2) = "MCC or ADwin ID # for DAQ Board"
                 Me.gridBoardSettings.TextMatrix(3, 0) = "BoardName"
                 Me.gridBoardSettings.TextMatrix(3, 1) = .BoardName
                 Me.gridBoardSettings.TextMatrix(3, 2) = "Unique String ID in Paleomag code for DAQ Board"
                 Me.gridBoardSettings.TextMatrix(4, 0) = "BoardFunction"
                 Me.gridBoardSettings.TextMatrix(4, 1) = .BoardFunction
                 Me.gridBoardSettings.TextMatrix(4, 2) = "How the Board can be used in Paleomag code"
                 Me.gridBoardSettings.TextMatrix(5, 0) = "CommProtocol"
                 Me.gridBoardSettings.TextMatrix(5, 1) = Trim(Str(.CommProtocol))
                 Me.gridBoardSettings.TextMatrix(5, 2) = "MCC(1) vs ADwin(2) comm protocol"
                 Me.gridBoardSettings.TextMatrix(6, 0) = "BoardMode"
                 Me.gridBoardSettings.TextMatrix(6, 1) = Trim(Str(.BoardMode))
                 Me.gridBoardSettings.TextMatrix(6, 2) = "Single(1) wire vs Differential(0), 2 wire Analog Input mode"
                 Me.gridBoardSettings.TextMatrix(7, 0) = "MaxAInRate"
                 Me.gridBoardSettings.TextMatrix(7, 1) = Trim(Str(.MaxAInRate))
                 Me.gridBoardSettings.TextMatrix(7, 2) = "Max Analog Input A/D rate (Hz) for DAQ Board"
                 Me.gridBoardSettings.TextMatrix(8, 0) = "MaxAOutRate"
                 Me.gridBoardSettings.TextMatrix(8, 1) = Trim(Str(.MaxAOutRate))
                 Me.gridBoardSettings.TextMatrix(8, 2) = "Max Analog Output D/A rate (Hz) for DAQ Board"
                 Me.gridBoardSettings.TextMatrix(9, 0) = "RangeMax"
                 Me.gridBoardSettings.TextMatrix(9, 1) = Trim(Str(.range.MaxValue))
                 Me.gridBoardSettings.TextMatrix(9, 2) = "Max Board Output Voltage"
                 Me.gridBoardSettings.TextMatrix(10, 0) = "RangeMin"
                 Me.gridBoardSettings.TextMatrix(10, 1) = Trim(Str(.range.MinValue))
                 Me.gridBoardSettings.TextMatrix(10, 2) = "Min Board Output Voltage"
                 Me.gridBoardSettings.TextMatrix(11, 0) = "AInChannelsCount"
                 Me.gridBoardSettings.TextMatrix(11, 1) = Trim(Str(.AInChannels.Count))
                 Me.gridBoardSettings.TextMatrix(11, 2) = "# of Analog Input Channels the Board has"
                 Me.gridBoardSettings.TextMatrix(12, 0) = "AOutChannelsCount"
                 Me.gridBoardSettings.TextMatrix(12, 1) = Trim(Str(.AOutChannels.Count))
                 Me.gridBoardSettings.TextMatrix(12, 2) = "# of Analog Output Channels the Board has"
                 Me.gridBoardSettings.TextMatrix(13, 0) = "DInChannelsCount"
                 Me.gridBoardSettings.TextMatrix(13, 1) = Trim(Str(.DInChannels.Count))
                 Me.gridBoardSettings.TextMatrix(13, 2) = "# of Digital Input Channels the Board has"
                 Me.gridBoardSettings.TextMatrix(14, 0) = "DOutChannelsCount"
                 Me.gridBoardSettings.TextMatrix(14, 1) = Trim(Str(.DOutChannels.Count))
                 Me.gridBoardSettings.TextMatrix(14, 2) = "# of Digital Output Channels the Board has"
                 Me.gridBoardSettings.TextMatrix(15, 0) = "DIOConfigured"
                 Me.gridBoardSettings.TextMatrix(15, 1) = Trim(Str(.DIOConfigured))
                 Me.gridBoardSettings.TextMatrix(15, 2) = "Are DIO channels pre-configured as Input or Output?"
                                  
                 'Store # of input and output channels to local variables
                 RowHeadAInChan = 17
                 RowHeadAOutChan = RowHeadAInChan + .AInChannels.Count + 3
                 RowHeadDInChan = RowHeadAOutChan + .AOutChannels.Count + 3
                 RowHeadDOutChan = RowHeadDInChan + .DInChannels.Count + 3
                    
            End With
                                
            With Me.gridBoardSettings
                    
                'Merge necessary rows for headers to each channel section
                .MergeCells = flexMergeRestrictRows
                .MergeRow(RowHeadAInChan - 1) = True
                .MergeRow(RowHeadAInChan) = True
                .MergeRow(RowHeadAOutChan - 1) = True
                .MergeRow(RowHeadAOutChan) = True
                .MergeRow(RowHeadDInChan - 1) = True
                .MergeRow(RowHeadDInChan) = True
                .MergeRow(RowHeadDOutChan - 1) = True
                .MergeRow(RowHeadDOutChan) = True
                
                'Make sure content in the cells in each of these rows are all the same
                For i = 0 To .Cols - 1
                
                    'Want a blank row before the header for each channel section
                    'with the same background cell color of the unpopulated section
                    'of the grid
                    .row = RowHeadAInChan - 1
                    .Col = i
                    .text = "     "
                    .CellBackColor = &HC0C0C0
                    
                    .row = RowHeadAOutChan - 1
                    .Col = i
                    .text = "     "
                    .CellBackColor = &HC0C0C0
                    
                    .row = RowHeadDInChan - 1
                    .Col = i
                    .text = "     "
                    .CellBackColor = &HC0C0C0
                    
                    .row = RowHeadDOutChan - 1
                    .Col = i
                    .text = "     "
                    .CellBackColor = &HC0C0C0
                    
                    'Now want channel section merged header with the name of the
                    'Channel Section + background color of fixed row (Row 0)
                    .row = RowHeadAInChan
                    .Col = i
                    .text = "Analog Input Channels"
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadAOutChan
                    .Col = i
                    .text = "Analog Output Channels"
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadDInChan
                    .Col = i
                    .text = "Digital Input Channels"
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadDOutChan
                    .Col = i
                    .text = "Digital Output Channles"
                    .CellBackColor = &H8000000F
                    
                    'Now set the cell background colors for the second header row
                    'to the fixed row background color
                    .row = 0
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadAInChan + 1
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadAOutChan + 1
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadDInChan + 1
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    .row = RowHeadDOutChan + 1
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                                                
                Next i
                
                'Now need to populate second header rows
                'By brute force coding
                .TextMatrix(RowHeadAInChan + 1, 0) = "Chan INI ID"
                .TextMatrix(RowHeadAInChan + 1, 1) = "Chan Name"
                .TextMatrix(RowHeadAInChan + 1, 2) = "Chan Number"
                .TextMatrix(RowHeadAOutChan + 1, 0) = "Chan INI ID"
                .TextMatrix(RowHeadAOutChan + 1, 1) = "Chan Name"
                .TextMatrix(RowHeadAOutChan + 1, 2) = "Chan Number"
                .TextMatrix(RowHeadDInChan + 1, 0) = "Chan INI ID"
                .TextMatrix(RowHeadDInChan + 1, 1) = "Chan Name"
                .TextMatrix(RowHeadDInChan + 1, 2) = "Chan Number"
                .TextMatrix(RowHeadDOutChan + 1, 0) = "Chan INI ID"
                .TextMatrix(RowHeadDOutChan + 1, 1) = "Chan Name"
                .TextMatrix(RowHeadDOutChan + 1, 2) = "Chan Number"
                
            End With
            
            With SystemBoards("ADWIN-light-16")
                
                'Now need to populate the data in the Analog In channel section
                For i = 1 To .AInChannels.Count
                
                    'Col 0 = the Channel's INI ID string
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAInChan + 1, 0) = _
                                modConfig.Create_INIChanStr(.AInChannels(i))
                                
                    'Col 1 = the Channel Name
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAInChan + 1, 1) = .AInChannels(i).ChanName
                    
                    'Col 2 = the Channel Number
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAInChan + 1, 2) = Trim(Str(.AInChannels(i).ChanNum))
                    
                Next i
                
                'Now need to populate the data in the Analog Out channel section
                For i = 1 To .AOutChannels.Count
                
                    'Col 0 = the Channel's INI ID string
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAOutChan + 1, 0) = _
                                modConfig.Create_INIChanStr(.AOutChannels(i))
                                
                    'Col 1 = the Channel Name
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAOutChan + 1, 1) = .AOutChannels(i).ChanName
                    
                    'Col 2 = the Channel Number
                    Me.gridBoardSettings.TextMatrix(i + RowHeadAOutChan + 1, 2) = Trim(Str(.AOutChannels(i).ChanNum))
                    
                Next i
                
                'Now need to populate the data in the Dig. In channel section
                For i = 1 To .DInChannels.Count
                
                    'Col 0 = the Channel's INI ID string
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDInChan + 1, 0) = _
                                modConfig.Create_INIChanStr(.DInChannels(i))
                                
                    'Col 1 = the Channel Name
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDInChan + 1, 1) = .DInChannels(i).ChanName
                    
                    'Col 2 = the Channel Number
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDInChan + 1, 2) = Trim(Str(.DInChannels(i).ChanNum))
                    
                Next i
                                    
                'Now need to populate the data in the Dig. Out channel section
                For i = 1 To .DOutChannels.Count
                
                    'Col 0 = the Channel's INI ID string
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDOutChan + 1, 0) = _
                                modConfig.Create_INIChanStr(.DOutChannels(i))
                                
                    'Col 1 = the Channel Name
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDOutChan + 1, 1) = .DOutChannels(i).ChanName
                    
                    'Col 2 = the Channel Number
                    Me.gridBoardSettings.TextMatrix(i + RowHeadDOutChan + 1, 2) = Trim(Str(.DOutChannels(i).ChanNum))
                    
                Next i
                                    
            End With
            
            'Change the column alignment for the 2nd & 3rd columns to Left justified
            Me.gridBoardSettings.ColAlignment(1) = flexAlignLeftCenter
            Me.gridBoardSettings.ColAlignment(2) = flexAlignLeftCenter
            
            'Set the top visible row
            Me.gridBoardSettings.TopRow = 0
            
            'Resize the board settings grid
            ResizeGrid Me.gridBoardSettings, _
                       Me, _
                       0, _
                       Me.gridBoardSettings.Rows - 1, , , , , _
                       False
            
            'Make the frame visible
            ShowFrame "8"
            
        Case "8"
        
            'Set the prior frame = "7"
            PriorFrame = "7"
        
            'Load in the Waveform objects from Defaults.INI into thenew Paleomag.INI file
            modConfig.Create_WavesForINI
            
            'Load this Waveform object info from Paleomag.INI into the System Globals
            modConfig.Get_WaveFormsFromIni
        
            'Need to check if user has selected to use the 2G or ADWIN AF system
            If Me.cmbAFSystem.text = "ADWIN" Then
            
                'Need to prepare frame 8A - for user to set channels for ADWIN Ramp output
                'and ADWIN Monitor input voltages
                
                'Need to populate the two combo-boxes
                LoadBoardAndChan WaveForms("AFMONITOR").BoardUsed.BoardName, _
                                 WaveForms("AFMONITOR").Chan.ChanName, _
                                 WaveForms("AFMONITOR").Chan.ChanType, _
                                 Nothing, _
                                 cmbADWINMonitorChan
                                 
                LoadBoardAndChan WaveForms("AFRAMPUP").BoardUsed.BoardName, _
                                 WaveForms("AFRAMPUP").Chan.ChanName, _
                                 WaveForms("AFRAMPUP").Chan.ChanType, _
                                 Nothing, _
                                 cmbADWINRampChan
                                 
                'Show frame 8A
                ShowFrame "8A"
                
            Else
            
                'Need to prepare Frame 8B - showing the Alternate AF Monitor Channel
                'and the IRM Monitor Channel settings
                
                'Need to populate the four combo-boxes
                LoadBoardAndChan WaveForms("ALTAFMONITOR").BoardUsed.BoardName, _
                                 WaveForms("ALTAFMONITOR").Chan.ChanName, _
                                 WaveForms("ALTAFMONITOR").Chan.ChanType, _
                                 Me.cmbAltAFMonitorBoard, _
                                 Me.cmbAltAFMonitorChan
                                 
                LoadBoardAndChan WaveForms("IRMMONITOR").BoardUsed.BoardName, _
                                 WaveForms("IRMMONITOR").Chan.ChanName, _
                                 WaveForms("IRMMONITOR").Chan.ChanType, _
                                 Me.cmbIRMMonitorBoard, _
                                 Me.cmbIRMMonitorChan

                'Show frame 8B
                ShowFrame "8B"
                
            End If
            
        Case "8A"
                
            'Set the prior frame = "8"
            PriorFrame = "8"
                
            'Need to prepare Frame 8B - showing the Alternate AF Monitor Channel
            'and the IRM Monitor Channel settings
            
            'Need to populate the four combo-boxes
            LoadBoardAndChan WaveForms("ALTAFMONITOR").BoardUsed.BoardName, _
                             WaveForms("ALTAFMONITOR").Chan.ChanName, _
                             WaveForms("ALTAFMONITOR").Chan.ChanType, _
                             Me.cmbAltAFMonitorBoard, _
                             Me.cmbAltAFMonitorChan
                             
            LoadBoardAndChan WaveForms("IRMMONITOR").BoardUsed.BoardName, _
                             WaveForms("IRMMONITOR").Chan.ChanName, _
                             WaveForms("IRMMONITOR").Chan.ChanType, _
                             Me.cmbIRMMonitorBoard, _
                             Me.cmbIRMMonitorChan

            'Show frame 8B
            ShowFrame "8B"
            
        Case "8B"
        
            'Need to prepare frame 8C with all of the waveform object settings
            'Use Frame 6 and the gridBoardSettings object
                        
            LoadWaveFormSettings Me.gridBoardSettings
            
            'Resize the grid
            ResizeGrid Me.gridBoardSettings, _
                       Me, _
                       0, _
                       Me.gridBoardSettings.Rows - 1, , , , , _
                       False
                                   
            'Change the caption in Frame 6
            Label11.Caption = "Wave Form Object Settings " & _
                              "(used for ADWIN AF, 2G AF monitor, and IRM monitor processes):"
                              
            'Need to set the prior frame, but that depends on the AF system that the user has inputed
            If Me.cmbAFSystem.text = "ADWIN" Then
            
                PriorFrame = "8A"
                
            Else
            
                PriorFrame = "8"
                
            End If
            
            'Show frame 6 as frame 8C
            ShowFrame "8C"
            
        Case "8C"
        
            'Set the prior frame = "8B"
            PriorFrame = "8B"
        
            'We now need to show the obsolete settings that won't be transfered
            'this will be done with gridDeletedSettings in Frame 10
            
            'This is a brute force job - very particular which settings need to be axed
            With Me.gridDeletedSettings
            
                'Set Grid dimensions
                .Rows = 61
                .Cols = 3
                
                'Set the alignment of the middle of the three columns (Col #1)
                'to left-horizontal, center-vertical justified
                .ColAlignment(1) = flexAlignLeftCenter
                                
                'Set fixed rows (no fixed columns)
                .FixedRows = 0
                .FixedCols = 0
                
                'Setup an array with the list of the first row that
                'is the header for a new block in the grid
                ReDim HeaderRows(9, 2)
                HeaderRows(0, 0) = "0"
                HeaderRows(0, 1) = "[IRMPulse]"
                HeaderRows(1, 0) = "16"
                HeaderRows(1, 1) = "[IRMPulseHF]"
                HeaderRows(2, 0) = "20"
                HeaderRows(2, 1) = "[Modules]"
                HeaderRows(3, 0) = "24"
                HeaderRows(3, 1) = "[AFDelay]"
                HeaderRows(4, 0) = "28"
                HeaderRows(4, 1) = "[AFRampRate]"
                HeaderRows(5, 0) = "32"
                HeaderRows(5, 1) = "[AFAxial]"
                HeaderRows(6, 0) = "39"
                HeaderRows(6, 1) = "[AFTrans]"
                HeaderRows(7, 0) = "46"
                HeaderRows(7, 1) = "[Vacuum]"
                HeaderRows(8, 0) = "52"
                HeaderRows(8, 1) = "[IRM-ARM]"
                
                'Setup Row Merges
                .MergeCells = flexMergeRestrictRows
                
                For i = 0 To UBound(HeaderRows, 1) - 1
                
                    If CLng(HeaderRows(i, 0)) > 0 Then
                            
                        'Setup blank rows
                        .MergeRow(CLng(HeaderRows(i, 0)) - 1) = True
                        .row = CLng(HeaderRows(i, 0)) - 1
                        
                        For j = 0 To .Cols - 1
                        
                            .Col = j
                            .text = "     "
                            .CellBackColor = &HC0C0C0
                            
                        Next j
                    
                    End If
                    
                    'Setup Named Rows
                    .MergeRow(CLng(HeaderRows(i, 0))) = True
                    .row = CLng(HeaderRows(i, 0))
                    
                    For j = 0 To .Cols - 1
                    
                        .Col = j
                        .text = HeaderRows(i, 1)
                        .CellBackColor = &H8000000F
                        
                    Next j
                    
                    'Setup Column Header rows
                    .row = CLng(HeaderRows(i, 0)) + 1
                    
                    For j = 0 To .Cols - 1
                    
                        .Col = j
                        .CellBackColor = &H8000000F
                        
                    Next j
                    
                    .TextMatrix(.row, 0) = "Setting Name"
                    .TextMatrix(.row, 1) = "Setting Value"
                    .TextMatrix(.row, 2) = "Reason for Deletion"
                    
                Next i
                
                .TextMatrix(2, 0) = "PulseY"
                .TextMatrix(2, 1) = Trim(OldIniFile.EntryRead("PulseY", "ERROR", "IRMPulse"))
                .TextMatrix(2, 2) = "Setting no longer used"
                                                    
                .TextMatrix(3, 0) = "PulseSlope"
                .TextMatrix(3, 1) = Trim(OldIniFile.EntryRead("PulseSlope", "ERROR", "IRMPulse"))
                .TextMatrix(3, 2) = "Setting no longer used"
                
                .TextMatrix(4, 0) = "IRMHFAxis"
                .TextMatrix(4, 1) = Trim(OldIniFile.EntryRead("IRMHFAxis", "ERROR", "IRMPulse"))
                .TextMatrix(4, 2) = "High Field IRM hardware does not exist"
                
                .TextMatrix(5, 0) = "IRMHFPlusLFAxis"
                .TextMatrix(5, 1) = Trim(OldIniFile.EntryRead("IRMHFPlusLFAxis", "ERROR", "IRMPulse"))
                .TextMatrix(5, 2) = "High Field IRM hardware does not exist"
                
                .TextMatrix(6, 0) = "PulseHFY"
                .TextMatrix(6, 1) = Trim(OldIniFile.EntryRead("PulseHFY", "ERROR", "IRMPulse"))
                .TextMatrix(6, 2) = "High Field IRM hardware does not exist"
                
                .TextMatrix(7, 0) = "PulseHFSlope"
                .TextMatrix(7, 1) = Trim(OldIniFile.EntryRead("PulseHFSlope", "ERROR", "IRMPulse"))
                .TextMatrix(7, 2) = "High Field IRM hardware does not exist"
                
                .TextMatrix(8, 0) = "PulseHFMax"
                .TextMatrix(8, 1) = Trim(OldIniFile.EntryRead("PulseHFMax", "ERROR", "IRMPulse"))
                .TextMatrix(8, 2) = "High Field IRM hardware does not exist"
                
                .TextMatrix(9, 0) = "PulseHFMin"
                .TextMatrix(9, 1) = Trim(OldIniFile.EntryRead("PulseHFMin", "ERROR", "IRMPulse"))
                .TextMatrix(9, 2) = "High Field IRM hardware does not exist"
                
                .TextMatrix(10, 0) = "PulseHFKeithleyVoltConverstion"
                .TextMatrix(10, 1) = Trim(OldIniFile.EntryRead("PulseHFKeithleyVoltConverstion", "ERROR", "IRMPulse"))
                .TextMatrix(10, 2) = "High Field IRM hardware does not exist"
                
                .TextMatrix(11, 0) = "PulseLFKeithleyVoltConverstion"
                .TextMatrix(11, 1) = Trim(OldIniFile.EntryRead("PulseLFKeithleyVoltConverstion", "ERROR", "IRMPulse"))
                .TextMatrix(11, 2) = "Setting no longer used"
                
                .TextMatrix(12, 0) = "PulseKeithleyVoltConverstion"
                .TextMatrix(12, 1) = Trim(OldIniFile.EntryRead("PulseKeithleyVoltConverstion", "ERROR", "IRMPulse"))
                .TextMatrix(12, 2) = "Setting no longer used"
                
                .TextMatrix(13, 0) = "KeithleyVoltConverstion"
                .TextMatrix(13, 1) = Trim(OldIniFile.EntryRead("KeithleyVoltConverstion", "ERROR", "IRMPulse"))
                .TextMatrix(13, 2) = "Setting no longer used"
                
                .TextMatrix(14, 0) = "PulseReturnKeithleyVoltConverstion"
                .TextMatrix(14, 1) = Trim(OldIniFile.EntryRead("PulseReturnKeithleyVoltConverstion", "ERROR", "IRMPulse"))
                .TextMatrix(14, 2) = "Setting no longer used"
                
                
                .TextMatrix(18, 0) = "Entire Section"
                .TextMatrix(18, 1) = "N/A"
                .TextMatrix(18, 2) = "High Field IRM hardware does not exist"
                
                
                .TextMatrix(22, 0) = "EnableIRMHi"
                .TextMatrix(22, 1) = Trim(OldIniFile.EntryRead("EnableIRMHi", "ERROR", "Modules"))
                .TextMatrix(22, 2) = "High Field IRM hardware does not exist"
                
                
                .TextMatrix(26, 0) = "AFDelay"
                .TextMatrix(26, 1) = Trim(OldIniFile.EntryRead("AFDelay", "ERROR", "AFDelay"))
                .TextMatrix(26, 2) = "Setting was moved to [AF] section"
                
                
                .TextMatrix(30, 0) = "AFRampRate"
                .TextMatrix(30, 1) = Trim(OldIniFile.EntryRead("AFRampRate", "ERROR", "AFRampRate"))
                .TextMatrix(30, 2) = "Setting was moved to [AF] section"
                
                
                .TextMatrix(34, 0) = "AFAxialLowSlope"
                .TextMatrix(34, 1) = OldIniFile.EntryRead("AFAxialLowSlope", _
                                                         "ERROR", _
                                                         "AFAxial")
                .TextMatrix(34, 2) = "Setting no longer used in Paleomag code"
                
                .TextMatrix(35, 0) = "AFAxialHighSlope"
                .TextMatrix(35, 1) = OldIniFile.EntryRead("AFAxialHighSlope", _
                                                         "ERROR", _
                                                         "AFAxial")
                .TextMatrix(35, 2) = "Setting no longer used in Paleomag code"
                
                .TextMatrix(36, 0) = "AFAxialXPoint"
                .TextMatrix(36, 1) = OldIniFile.EntryRead("AFAxialXPoint", _
                                                         "ERROR", _
                                                         "AFAxial")
                .TextMatrix(36, 2) = "Setting no longer used in Paleomag code"
                
                .TextMatrix(37, 0) = "AFAxialYPoint"
                .TextMatrix(37, 1) = OldIniFile.EntryRead("AFAxialYPoint", _
                                                         "ERROR", _
                                                         "AFAxial")
                .TextMatrix(37, 2) = "Setting no longer used in Paleomag code"
                
                
                .TextMatrix(41, 0) = "AFTransLowSlope"
                .TextMatrix(41, 1) = OldIniFile.EntryRead("AFTransLowSlope", _
                                                         "ERROR", _
                                                         "AFTrans")
                .TextMatrix(41, 2) = "Setting no longer used in Paleomag code"
                
                .TextMatrix(42, 0) = "AFTransHighSlope"
                .TextMatrix(42, 1) = OldIniFile.EntryRead("AFTransHighSlope", _
                                                         "ERROR", _
                                                         "AFTrans")
                .TextMatrix(42, 2) = "Setting no longer used in Paleomag code"
                
                .TextMatrix(43, 0) = "AFTransXPoint"
                .TextMatrix(43, 1) = OldIniFile.EntryRead("AFTransXPoint", _
                                                         "ERROR", _
                                                         "AFTrans")
                .TextMatrix(43, 2) = "Setting no longer used in Paleomag code"
                
                .TextMatrix(44, 0) = "AFTransYPoint"
                .TextMatrix(44, 1) = OldIniFile.EntryRead("AFTransYPoint", _
                                                         "ERROR", _
                                                         "AFTrans")
                .TextMatrix(44, 2) = "Setting no longer used in Paleomag code"
                
                
                .TextMatrix(48, 0) = "MotorToggle"
                .TextMatrix(48, 1) = Trim(OldIniFile.EntryRead("MotorToggle", "ERROR", "Vacuum"))
                .TextMatrix(48, 2) = "Setting moved to [Channels] section"
                                                    
                .TextMatrix(49, 0) = "VacuumToggleA"
                .TextMatrix(49, 1) = Trim(OldIniFile.EntryRead("VacuumToggleA", "ERROR", "Vacuum"))
                .TextMatrix(49, 2) = "Setting moved to [Channels] section"
                
                .TextMatrix(50, 0) = "VacuumToggleB"
                .TextMatrix(50, 1) = Trim(OldIniFile.EntryRead("VacuumToggleB", "ERROR", "Vacuum"))
                .TextMatrix(50, 2) = "Setting moved to [Channels] section"
                
                
                .TextMatrix(54, 0) = "ARMVoltageOut"
                .TextMatrix(54, 1) = Trim(OldIniFile.EntryRead("ARMVoltageOut", "ERROR", "IRM-ARM"))
                .TextMatrix(54, 2) = "Setting moved to [Channels] section"
                
                .TextMatrix(55, 0) = "IRMVoltageOut"
                .TextMatrix(55, 1) = Trim(OldIniFile.EntryRead("IRMVoltageOut", "ERROR", "IRM-ARM"))
                .TextMatrix(55, 2) = "Setting moved to [Channels] section"
                
                .TextMatrix(56, 0) = "IRMCapacitorVoltageIn"
                .TextMatrix(56, 1) = Trim(OldIniFile.EntryRead("IRMCapacitorVoltageIn", "ERROR", "IRM-ARM"))
                .TextMatrix(56, 2) = "Setting moved to [Channels] section"
                
                .TextMatrix(57, 0) = "ARMSet"
                .TextMatrix(57, 1) = Trim(OldIniFile.EntryRead("ARMSet", "ERROR", "IRM-ARM"))
                .TextMatrix(57, 2) = "Setting moved to [Channels] section"
                
                .TextMatrix(58, 0) = "IRMFire"
                .TextMatrix(58, 1) = Trim(OldIniFile.EntryRead("IRMFire", "ERROR", "IRM-ARM"))
                .TextMatrix(58, 2) = "Setting moved to [Channels] section"
                
                .TextMatrix(59, 0) = "IRMTrim"
                .TextMatrix(59, 1) = Trim(OldIniFile.EntryRead("IRMTrim", "ERROR", "IRM-ARM"))
                .TextMatrix(59, 2) = "Setting moved to [Channels] section"
                
                .TextMatrix(60, 0) = "IRMReady"
                .TextMatrix(60, 1) = Trim(OldIniFile.EntryRead("IRMReady", "ERROR", "IRM-ARM"))
                .TextMatrix(60, 2) = "Setting Obsolete. This hardware no longer supported."
                
                .TopRow = 0
                
            End With
                
            'Resize the grid
            ResizeGrid Me.gridDeletedSettings, _
                       Me, _
                       0, _
                       Me.gridDeletedSettings.Rows - 1, , , , , _
                       False
                                       
            'Show frame 9
            ShowFrame "9"
                                       
        Case "9"
                
            'Set the prior frame = "8C"
            PriorFrame = "8C"
                
            'Need to show transfered old AF Main settings and the New AF settings
            'Use gridBoardSettings & Frame 6
            With Me.gridBoardSettings
            
                'Clear out the grid
                .Clear
                .ClearStructure
            
                'Set grid dimensions
                .Rows = 20
                .Cols = 3
                
                'Set the fixed rows and cols
                .FixedCols = 0
                .FixedRows = 0
                
                'Set the row merge properties
                .MergeCells = flexMergeRestrictRows
                .MergeRow(0) = True
                .MergeRow(4) = True
                .MergeRow(5) = True
                
                'Set the column headers
                .TextMatrix(1, 0) = "Setting Name"
                .TextMatrix(1, 1) = "Setting Value"
                .TextMatrix(1, 2) = "Decription"
                
                'Setup blank row
                .row = 4
                
                For j = 0 To .Cols - 1
                
                    .Col = j
                    .text = "    "
                    .CellBackColor = &HC0C0C0
                    
                Next j
                
                
                'Setup Named Rows
                .row = 0
                    
                For j = 0 To .Cols - 1
                
                    .Col = j
                    .text = "Transfered Settings"
                    .CellBackColor = &H8000000F
                    
                Next j
                
                'Setup Column Header rows
                .row = 1
                
                For j = 0 To .Cols - 1
                
                    .Col = j
                    .CellBackColor = &H8000000F
                    
                Next j
            
                .row = 5
                    
                For j = 0 To .Cols - 1
                
                    .Col = j
                    .text = "New Settings"
                    .CellBackColor = &H8000000F
                    
                Next j
                
                'Setup Column Header rows
                .row = 6
                
                For j = 0 To .Cols - 1
                
                    .Col = j
                    .CellBackColor = &H8000000F
                    
                Next j
            
                .TextMatrix(6, 0) = "Setting Name"
                .TextMatrix(6, 1) = "Setting Value"
                .TextMatrix(6, 2) = "Decription"
                
                'Now transfered old settings data from the old INI file to the new INI file
                AFSection = OldIniFile.SectionRead(True, False, "AF")
                IniFile.SectionWrite AFSection, "AF"
                
                'Now add those settings to the Grid
                .TextMatrix(2, 0) = "AFDelay"
                .TextMatrix(2, 1) = IniFile.EntryRead("AFDelay", _
                                                      "ERROR", _
                                                      "AF")
                                                      
                .TextMatrix(2, 2) = "Sets the communication delay for the 2G AF box"
                
                .TextMatrix(3, 0) = "AFRampRate"
                .TextMatrix(3, 1) = IniFile.EntryRead("AFRampRate", _
                                                      "ERROR", _
                                                      "AF")
                                                      
                .TextMatrix(3, 2) = "Sets the Ramp Speed for the 2G AF box"
                
                'Now transfer the new AF settings from the defaults file to the new ini file
                With IniFile
                
                    .EntryWrite "AFUnits", _
                                DefaultINI.EntryRead( _
                                            "AFUnits", _
                                            "G", _
                                            "AF"), _
                                "AF"
                                
                    If Me.cmbAFSystem.text = "2G" Then
                                
                        modConfig.AFSystem = "2G"
                        
                    Else
                    
                        modConfig.AFSystem = "ADWIN"
                        
                    End If
                                
                    .EntryWrite "AFSystem", _
                                modConfig.AFSystem, _
                                "AF"
                                
                    .EntryWrite "AFWait", _
                                DefaultINI.EntryRead( _
                                            "AFWait", _
                                            "90", _
                                            "AF"), _
                                "AF"
                                
                    .EntryWrite "TSlope", _
                                DefaultINI.EntryRead( _
                                            "TSlope", _
                                            "58.86", _
                                            "AF"), _
                                "AF"
                                
                    .EntryWrite "Toffset", _
                                DefaultINI.EntryRead( _
                                            "Toffset", _
                                            "289.6", _
                                            "AF"), _
                                "AF"
                                
                    .EntryWrite "Thot", _
                                DefaultINI.EntryRead( _
                                            "Thot", _
                                            "40", _
                                            "AF"), _
                                "AF"
                                
                    .EntryWrite "Tmax", _
                                DefaultINI.EntryRead( _
                                            "Tmax", _
                                            "50", _
                                            "AF"), _
                                "AF"
                                
                    .EntryWrite "Tunits", _
                                DefaultINI.EntryRead( _
                                            "Tunits", _
                                            "C", _
                                            "AF"), _
                                "AF"
                                
                    .EntryWrite "MinRampUpTime_ms", _
                                DefaultINI.EntryRead( _
                                            "MinRampUpTime_ms", _
                                            "500", _
                                            "AF"), _
                                "AF"
                                
                    .EntryWrite "MaxRampUpTime_ms", _
                                DefaultINI.EntryRead( _
                                            "MaxRampUpTime_ms", _
                                            "1000", _
                                            "AF"), _
                                "AF"
                                                                
                    .EntryWrite "MinRampDown_NumPeriods", _
                                DefaultINI.EntryRead( _
                                            "MinRampDown_NumPeriods", _
                                            "250", _
                                            "AF"), _
                                "AF"
                                                                
                    .EntryWrite "MaxRampDown_NumPeriods", _
                                DefaultINI.EntryRead( _
                                            "MaxRampDown_NumPeriods", _
                                            "5000", _
                                            "AF"), _
                                "AF"
                                                                
                    .EntryWrite "RampDownNumPeriodsPerVolt", _
                                DefaultINI.EntryRead( _
                                            "RampDownNumPeriodsPerVolt", _
                                            "2000", _
                                            "AF"), _
                                "AF"
                                                                
                End With
                
                'Now add those settings to the grid
                .TextMatrix(7, 0) = "AFUnits"
                .TextMatrix(7, 1) = IniFile.EntryRead("AFUnits", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(7, 2) = "Magnetic Field Units for displaying the AF & IRM field values"
                
                .TextMatrix(8, 0) = "AFSystem"
                .TextMatrix(8, 1) = IniFile.EntryRead("AFSystem", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(8, 2) = "AF System to Use - 2G or ADWIN"
                
                .TextMatrix(9, 0) = "AFWait"
                .TextMatrix(9, 1) = IniFile.EntryRead("AFWait", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(9, 2) = "Wait time after AF Coil Temp. exceeds max. allowed Temp."
                
                .TextMatrix(10, 0) = "AFSystem"
                .TextMatrix(10, 1) = IniFile.EntryRead("AFSystem", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(10, 2) = "AF System to Use - 2G or ADWIN"
                
                .TextMatrix(10, 0) = "TSlope"
                .TextMatrix(10, 1) = IniFile.EntryRead("TSlope", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(10, 2) = "Slope calibration for AF Coil Temp. transducers"

                .TextMatrix(11, 0) = "Toffset"
                .TextMatrix(11, 1) = IniFile.EntryRead("Toffset", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(11, 2) = "Offset calibration for AF Coil Temp. transducers"


                .TextMatrix(12, 0) = "Thot"
                .TextMatrix(12, 1) = IniFile.EntryRead("Thot", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(12, 2) = "AF Temp. at which to start warning user."


                .TextMatrix(13, 0) = "Tmax"
                .TextMatrix(13, 1) = IniFile.EntryRead("Tmax", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(13, 2) = "AF Temp. at which to pause AF ramp cycle"

                .TextMatrix(14, 0) = "Tunits"
                .TextMatrix(14, 1) = IniFile.EntryRead("Tunits", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(14, 2) = "Units to use for AF Temp."


                .TextMatrix(15, 0) = "MinRampUpTime_ms"
                .TextMatrix(15, 1) = IniFile.EntryRead("MinRampUpTime_ms", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(15, 2) = "Minimum allowed ADWIN AF Ramp Up time"


                .TextMatrix(16, 0) = "MaxRampUpTime_ms"
                .TextMatrix(16, 1) = IniFile.EntryRead("MaxRampUpTime_ms", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(16, 2) = "Maxium allowed ADWIN AF Ramp Up time"


                .TextMatrix(17, 0) = "MinRampDown_NumPeriods"
                .TextMatrix(17, 1) = IniFile.EntryRead("MinRampDown_NumPeriods", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(17, 2) = "Min. # ADWIN AF Ramp down LC circuit periods"


                .TextMatrix(18, 0) = "MaxRampDown_NumPeriods"
                .TextMatrix(18, 1) = IniFile.EntryRead("MaxRampDown_NumPeriods", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(18, 2) = "Max. # ADWIN AF Ramp down LC circuit periods"


                .TextMatrix(19, 0) = "RampDownNumPeriodsPerVolt"
                .TextMatrix(19, 1) = IniFile.EntryRead("RampDownNumPeriodsPerVolt", _
                                                     "ERROR", _
                                                     "AF")
                .TextMatrix(19, 2) = "ADWIN AF Ramp Down rate"
                
            End With
            
            'Resize the grid
            ResizeGrid Me.gridBoardSettings, _
                       Me, _
                       0, _
                       Me.gridBoardSettings.Rows - 1, , , , , _
                       False
                       
            'Change the label caption
            Me.Label11 = "New AF System Main Settings"
            
            'Show the grid
            ShowFrame "10"
            
        Case "10"
        
            'Set the prior frame = "9"
            PriorFrame = "9"
            
            'Now need to show the AF Axial settings
            
            'Transfer AF Axial settings from default file to the New INI File
            TransferAFAxisSettings Axial
            
            'Add values to grid
            With Me.gridBoardSettings
            
                'Clear the grid
                .Clear
                .ClearStructure
            
                'Set dimensions & fixed cells
                .Rows = 10 + modConfig.AFAxialCount
                .Cols = 3
                .FixedRows = 0
                .FixedCols = 0
                
                'Setup Merged rows
                .MergeCells = flexMergeRestrictRows
                .MergeRow(0) = True
                .MergeRow(7) = True
                .MergeRow(8) = True
                
                'Setup 1st Header rows
                For i = 0 To .Cols - 1
                
                    'Set text
                    .TextMatrix(0, i) = "Non Calibration Settings"
                    .TextMatrix(7, i) = "    "
                    .TextMatrix(8, i) = "Field Calibration Settings"
                    
                    'Set back color, empty row
                    .row = 7
                    .Col = i
                    .CellBackColor = &HC0C0C0
                    
                    'Set back color 1st header row
                    .row = 0
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    'Set back color 2nd header row
                    .row = 1
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    'Set back color 1st cal. settings header row
                    .row = 8
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    'Set back color 2nd cal. settings header row
                    .row = 9
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                Next i
                
                'Setup 2nd Header Rows
                .TextMatrix(1, 0) = "Setting Name"
                .TextMatrix(1, 1) = "Setting Value"
                .TextMatrix(1, 2) = "Setting Description"
                
                .TextMatrix(9, 0) = "#"
                .TextMatrix(9, 1) = "X (2G OR Volts)"
                .TextMatrix(9, 2) = "Y (Field in G)"
                
                'Write in non-calibration settings to the grid
                .TextMatrix(2, 0) = "AFAxialCalDone"
                .TextMatrix(2, 1) = Trim(Str(modConfig.AFAxialCalDone))
                .TextMatrix(2, 2) = "Is the Axial Coil Field calibrated?"
                
                .TextMatrix(3, 0) = "AFAxialCoord"
                .TextMatrix(3, 1) = modConfig.AfAxialCoord
                .TextMatrix(3, 2) = "2G AF Relay control coordinate."
                
                .TextMatrix(4, 0) = "AFAxialMax"
                .TextMatrix(4, 1) = Trim(Str(modConfig.AfAxialMax))
                .TextMatrix(4, 2) = "Max. 2G Counts -or- peak volt. for AF Axial Ramp"
                
                .TextMatrix(5, 0) = "AFAxialMin"
                .TextMatrix(5, 1) = Trim(Str(modConfig.AfAxialMin))
                .TextMatrix(5, 2) = "Min. 2G Counts -or- peak volt. for AF Axial Ramp"
                
                .TextMatrix(6, 0) = "AFAxialCount"
                .TextMatrix(6, 1) = Trim(Str(modConfig.AFAxialCount))
                .TextMatrix(6, 2) = "# of AF Axial field cal. data points"
                                                    
                
                'Now need to load the values for the AF calibration data pairs
                For i = 10 To .Rows - 1
                
                    .TextMatrix(i, 0) = Trim(Str(i - 9))
                    .TextMatrix(i, 1) = Trim(Str(modConfig.AFAxial(i - 9, 0)))
                    .TextMatrix(i, 2) = Trim(Str(modConfig.AFAxial(i - 9, 1)))
                    
                Next i
                
            End With
            
            'Resize the grid
            ResizeGrid Me.gridBoardSettings, _
                       Me, _
                       0, _
                       Me.gridBoardSettings.Rows - 1, , , , , _
                       False
                       
            'Change the label caption
            Me.Label11.Caption = "AF Axial Settings"
            
            'Show the frame
            ShowFrame "11"
            
        Case "11"
            
            'Set the prior frame = "10"
            PriorFrame = "10"
            
            'Now need to show the AF Transverse settings
            
            'Transfer AF Transverse settings from the old INI file to the New INI File
            TransferAFAxisSettings Transverse
            
            'Add values to grid
            With Me.gridBoardSettings
            
                'Clear the grid
                .Clear
                .ClearStructure
            
                'Set dimensions & fixed cells
                .Rows = 10 + modConfig.AFTransCount
                .Cols = 3
                .FixedRows = 2
                .FixedCols = 0
                
                'Setup Merged rows
                .MergeCells = flexMergeRestrictRows
                .MergeRow(0) = True
                .MergeRow(7) = True
                .MergeRow(8) = True
                
                'Setup 1st Header rows
                For i = 0 To .Cols - 1
                
                    'Set text
                    .TextMatrix(0, i) = "Non Calibration Settings"
                    .TextMatrix(7, i) = "    "
                    .TextMatrix(8, i) = "Field Calibration Settings"
                    
                    'Set back color, empty row
                    .row = 7
                    .Col = i
                    .CellBackColor = &HC0C0C0
                    
                    'Set back color 1st cal. settings header row
                    .row = 8
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    'Set back color 2nd cal. settings header row
                    .row = 9
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                Next i
                
                'Setup 2nd Header Rows
                .TextMatrix(1, 0) = "Setting Name"
                .TextMatrix(1, 1) = "Setting Value"
                .TextMatrix(1, 2) = "Setting Description"
                
                .TextMatrix(9, 0) = "#"
                .TextMatrix(9, 1) = "X (2G counts / V)"
                .TextMatrix(9, 2) = "Y (Field in G)"
                
                'Write in non-calibration settings to the grid
                .TextMatrix(2, 0) = "AFTransCalDone"
                .TextMatrix(2, 1) = Trim(Str(modConfig.AFTransCalDone))
                .TextMatrix(2, 2) = "Is the Trans Coil Field Calibrated?"
                
                .TextMatrix(3, 0) = "AFTransCoord"
                .TextMatrix(3, 1) = modConfig.AfTransCoord
                .TextMatrix(3, 2) = "2G AF Relay control coordinate."
                
                .TextMatrix(4, 0) = "AFTransMax"
                .TextMatrix(4, 1) = Trim(Str(modConfig.AfTransMax))
                .TextMatrix(4, 2) = "Max. 2G Counts -or- peak volt. for AF Trans Ramp"
                
                .TextMatrix(5, 0) = "AFTransMin"
                .TextMatrix(5, 1) = Trim(Str(modConfig.AfTransMin))
                .TextMatrix(5, 2) = "Min. 2G Counts -or- peak volt. for AF Trans Ramp"
                
                .TextMatrix(6, 0) = "AFTransCount"
                .TextMatrix(6, 1) = Trim(Str(modConfig.AFTransCount))
                .TextMatrix(6, 2) = "# of AF Trans field cal. data points"
                                                    
                
                'Now need to load the values for the AF calibration data pairs
                For i = 10 To .Rows - 1
                
                    .TextMatrix(i, 0) = Trim(Str(i - 9))
                    .TextMatrix(i, 1) = Trim(Str(modConfig.AFTrans(i - 9, 0)))
                    .TextMatrix(i, 2) = Trim(Str(modConfig.AFTrans(i - 9, 1)))
                    
                Next i
                
            End With
            
            'Resize the grid
            ResizeGrid Me.gridBoardSettings, _
                       Me, _
                       0, _
                       Me.gridBoardSettings.Rows - 1, , , , , _
                       False
                       
            'Change the label caption
            Me.Label11.Caption = "AF Transverse Settings"
            
            'Show the frame
            ShowFrame "12"
            
        Case "12"
            
            'Set the prior frame = "11"
            PriorFrame = "11"
            
            'Now need to show the IRM Main settings
            
            'Transfer the IRM pulse settings that are being saved from the old INI file to the
            'new ini file
            'Now transfer the new AF settings from the defaults file to the new ini file
            'while also loading the global variables
            With IniFile
            
                If Me.cmbIRMSystem.text = "Caltech Old" Then
                            
                    modConfig.IRMSystem = "Old"
                    
                Else
                
                    modConfig.IRMSystem = "Matsusada"
                    
                End If
                            
                .EntryWrite "IRMSystem", _
                            modConfig.IRMSystem, _
                            "IRMPulse"
                            
                modConfig.PulseMCCVoltConversion = val( _
                                OldIniFile.EntryRead( _
                                            "PulseMCCVoltConversion", _
                                            ".02902", _
                                            "IRMPulse"))
                .EntryWrite "PulseMCCVoltConversion", _
                            Trim(Str(modConfig.PulseMCCVoltConversion)), _
                            "IRMPulse"
                            
                modConfig.PulseReturnMCCVoltConversion = val( _
                                OldIniFile.EntryRead( _
                                            "PulseReturnMCCVoltConversion", _
                                            ".02902", _
                                            "IRMPulse"))
                .EntryWrite "PulseReturnMCCVoltConversion", _
                            Trim(Str(modConfig.PulseReturnMCCVoltConversion)), _
                            "IRMPulse"
                            
                modConfig.PulseVoltMax = val( _
                                OldIniFile.EntryRead( _
                                            "PulseVoltMax", _
                                            "10", _
                                            "IRMPulse"))
                .EntryWrite "PulseVoltMax", _
                            Trim(Str(modConfig.PulseVoltMax)), _
                            "IRMPulse"
                            
                modConfig.IRMAxis = OldIniFile.EntryRead( _
                                            "IRMLFAxis", _
                                            "ERROR", _
                                            "IRMPulse")
                .EntryWrite "IRMAxis", _
                            modConfig.IRMAxis, _
                            "IRMPulse"
                          
                modConfig.IRMBackfieldAxis = OldIniFile.EntryRead( _
                                                "IRMLFBackfieldAxis", _
                                                "ERROR", _
                                                "IRMPulse")
                .EntryWrite "IRMBackfieldAxis", _
                            modConfig.IRMBackfieldAxis, _
                            "IRMPulse"
                            
                'Transfer the TrimOnTrue settings as entered by the User in Frame 3A
                .EntryWrite "TrimOnTrue", _
                            Trim(Str(Me.optTrimOnTrue.Value)), _
                            "IRMPulse"
                                                            
            End With
            
            'Now need to setup the grid
            With Me.gridBoardSettings
            
                'Clear the grid
                .Clear
                .ClearStructure
                
                'Setup the number of rows and cols
                .Rows = 8
                .Cols = 3
                
                'Set # of fixed cells
                .FixedCols = 0
                .FixedRows = 1
                
                'Setup the col headers
                .TextMatrix(0, 0) = "Setting Name"
                .TextMatrix(0, 1) = "Setting Value"
                .TextMatrix(0, 2) = "Setting Description"
                
                'Now load in the values
                .TextMatrix(1, 0) = "IRMSystem"
                .TextMatrix(1, 1) = modConfig.IRMSystem
                .TextMatrix(1, 2) = "Matsusada vs Old IRM system"
                
                .TextMatrix(2, 0) = "PulseMCCVoltConversion"
                .TextMatrix(2, 1) = Trim(Str(modConfig.PulseMCCVoltConversion))
                .TextMatrix(2, 2) = "MCC Output Volts -> IRM capacitor box volts"
                
                .TextMatrix(3, 0) = "PulseReturnMCCVoltConversion"
                .TextMatrix(3, 1) = Trim(Str(modConfig.PulseReturnMCCVoltConversion))
                .TextMatrix(3, 2) = "IRM capacitor box volts -> MCC Output Volts"
                
                .TextMatrix(4, 0) = "PulseVoltMax"
                .TextMatrix(4, 1) = Trim(Str(modConfig.PulseVoltMax))
                .TextMatrix(4, 2) = "Max MCC volt. output for IRM ( <= 10 )"
                
                .TextMatrix(5, 0) = "IRMAxis"
                .TextMatrix(5, 1) = modConfig.IRMAxis
                .TextMatrix(5, 2) = "2G Axis for setting the IRM Relay"
                
                .TextMatrix(6, 0) = "IRMBackfieldAxis"
                .TextMatrix(6, 1) = modConfig.IRMBackfieldAxis
                .TextMatrix(6, 2) = "2G Axis for setting the IRM Backfield Relay"
                
                .TextMatrix(7, 0) = "TrimOnTrue"
                .TextMatrix(7, 1) = Trim(Str(Me.optTrimOnTrue.Value))
                .TextMatrix(7, 2) = "IRM Trim logic setting, Trim On = True or False?"
            
            End With
            
            'Resize the grid
            ResizeGrid Me.gridBoardSettings, _
                       Me, _
                       0, _
                       Me.gridBoardSettings.Rows - 1, , , , , _
                       False
                       
            'Change the label
            Me.Label11.Caption = "IRM Main Settings"
                       
            'Show the frame
            ShowFrame "13"
            
        Case "13"
            
            'Set the prior frame = "12"
            PriorFrame = "12"
            
            'Need to show settings for the Axial IRM
            
            'Need to transfer old Axial IRM settings to new INI file & global variables
            TransferIRMAxisSettings Axial
           
            'Setup the grid
            With Me.gridBoardSettings
            
                'Clear the grid
                .Clear
                .ClearStructure
                
                'Set the grid dimensions
                .Rows = 10 + modConfig.PulseAxialCount
                .Cols = 3
                
                'Set the fixed cells
                .FixedRows = 2
                .FixedCols = 0
                
                'Setup the merged rows
                .MergeCells = flexMergeRestrictRows
                .MergeRow(0) = True
                .MergeRow(7) = True
                .MergeRow(8) = True
                
                'Setup the merged headers and colors
                For i = 0 To .Cols - 1
                
                    .TextMatrix(0, i) = "Axial Non-calibration Settings"
                    .TextMatrix(7, i) = "    "
                    .TextMatrix(8, i) = "IRM Calibration Settings"
                    
                    .row = 7
                    .Col = i
                    .CellBackColor = &HC0C0C0
                    
                    .row = 8
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                    .row = 9
                    .Col = i
                    .CellBackColor = &H8000000F
                    
                Next i
                
                'Setup non-merged headers
                .TextMatrix(1, 0) = "Setting Name"
                .TextMatrix(1, 1) = "Setting Value"
                .TextMatrix(1, 2) = "Setting Description"
                
                .TextMatrix(9, 0) = "#"
                .TextMatrix(9, 1) = "X (IRM Volt.)"
                .TextMatrix(9, 2) = "Y (Field in G)"
                
                'Add in non-calibration values now
                .TextMatrix(2, 0) = "IRMAxialCalDone"
                .TextMatrix(2, 1) = Trim(Str(modConfig.IRMAxialCalDone))
                .TextMatrix(2, 2) = "Is the Axial IRM calibrated?"
                
                .TextMatrix(3, 0) = "IRMAxialVoltMax"
                .TextMatrix(3, 1) = Trim(Str(modConfig.IRMAxialVoltMax))
                .TextMatrix(3, 2) = "Max allowed IRM Capacitor box voltage"
                
                .TextMatrix(4, 0) = "PulseAxialMax"
                .TextMatrix(4, 1) = Trim(Str(modConfig.PulseAxialMax))
                .TextMatrix(4, 2) = "Max allowed IRM Field in Gauss"
                
                .TextMatrix(5, 0) = "PulseAxialMin"
                .TextMatrix(5, 1) = Trim(Str(modConfig.PulseAxialMin))
                .TextMatrix(5, 2) = "Min allowed IRM Field in Gauss"
                
                .TextMatrix(6, 0) = "PulseAxialCount"
                .TextMatrix(6, 1) = Trim(Str(modConfig.PulseAxialCount))
                .TextMatrix(6, 2) = "# of IRM field calibration points"
                
                'Now add in the calibration data points
                For i = 10 To .Rows - 1
                
                    .TextMatrix(i, 0) = Trim(Str(i - 9))
                    .TextMatrix(i, 1) = Trim(Str(PulseAxial(i - 9, 0)))
                    .TextMatrix(i, 2) = Trim(Str(PulseAxial(i - 9, 1)))
                
                Next i
                
            End With
                
            'Resize the grid
            ResizeGrid Me.gridBoardSettings, _
                       Me, _
                       0, _
                       Me.gridBoardSettings.Rows - 1, , , , , _
                       False
                       
            'Change the Label caption
            Me.Label11.Caption = "IRM Axial Settings"
            
            'Show the frame
            ShowFrame "14"
            
        Case "14"
                            
            'Set the prior frame = "13"
            PriorFrame = "13"
            
            'Need to show settings for the Trans IRM
            
            'Transfer IRM Trans settings into new INI file + global variables
            TransferIRMAxisSettings Transverse
                        
            'Setup the grid
            With Me.gridBoardSettings
            
                'Clear the grid
                .Clear
                .ClearStructure
                
                'Set the grid dimensions
                If modConfig.PulseTransCount > 0 Then
                
                    .Rows = 10 + modConfig.PulseTransCount
                    
                Else
                
                    .Rows = 7
                    
                End If
                                    
                .Cols = 3
                
                'Set the fixed cells
                .FixedRows = 2
                .FixedCols = 0
                
                'Setup the merged rows
                .MergeCells = flexMergeRestrictRows
                .MergeRow(0) = True
                
                If modConfig.PulseTransCount > 0 Then
                    
                    .MergeRow(7) = True
                    .MergeRow(8) = True
                    
                End If
                
                'Setup the merged headers and colors
                For i = 0 To .Cols - 1
                
                    .TextMatrix(0, i) = "Trans Non-calibration Settings"
                    
                    If modConfig.PulseTransCount > 0 Then
                        
                        .TextMatrix(7, i) = "    "
                        .TextMatrix(8, i) = "IRM Calibration Settings"
                        
                        .row = 7
                        .Col = i
                        .CellBackColor = &HC0C0C0
                        
                        .row = 8
                        .Col = i
                        .CellBackColor = &H8000000F
                        
                        .row = 9
                        .Col = i
                        .CellBackColor = &H8000000F
                        
                    End If
                    
                Next i
                
                'Setup non-merged headers
                .TextMatrix(1, 0) = "Setting Name"
                .TextMatrix(1, 1) = "Setting Value"
                .TextMatrix(1, 2) = "Setting Description"
                
                If modConfig.PulseTransCount > 0 Then
                    
                    .TextMatrix(9, 0) = "#"
                    .TextMatrix(9, 1) = "X (IRM Volt.)"
                    .TextMatrix(9, 2) = "Y (Field in G)"
                
                End If
                
                'Add in non-calibration values now
                .TextMatrix(2, 0) = "IRMTransCalDone"
                .TextMatrix(2, 1) = Trim(Str(modConfig.IRMTransCalDone))
                .TextMatrix(2, 2) = "Is the Trans IRM calibrated?"
                
                .TextMatrix(3, 0) = "IRMTransVoltMax"
                .TextMatrix(3, 1) = Trim(Str(modConfig.IRMTransVoltMax))
                .TextMatrix(3, 2) = "Max allowed IRM Capacitor box voltage"
                
                .TextMatrix(4, 0) = "PulseTransMax"
                .TextMatrix(4, 1) = Trim(Str(modConfig.PulseTransMax))
                .TextMatrix(4, 2) = "Max allowed IRM Field in Gauss"
                
                .TextMatrix(5, 0) = "PulseTransMin"
                .TextMatrix(5, 1) = Trim(Str(modConfig.PulseTransMin))
                .TextMatrix(5, 2) = "Min allowed IRM Field in Gauss"
                
                .TextMatrix(6, 0) = "PulseTransCount"
                .TextMatrix(6, 1) = Trim(Str(modConfig.PulseTransCount))
                .TextMatrix(6, 2) = "# of IRM field calibration points"
                
                'Now add in the calibration data points
                If modConfig.PulseTransCount > 0 Then
                
                    For i = 10 To .Rows - 1
                    
                        .TextMatrix(i, 0) = Trim(Str(i - 9))
                        .TextMatrix(i, 1) = Trim(Str(PulseTrans(i - 9, 0)))
                        .TextMatrix(i, 2) = Trim(Str(PulseTrans(i - 9, 1)))
                    
                    Next i
                    
                End If
                    
            End With
                
            'Resize the grid
            ResizeGrid Me.gridBoardSettings, _
                       Me, _
                       0, _
                       Me.gridBoardSettings.Rows - 1, , , , , _
                       False
                       
            'Change the Label caption
            Me.Label11.Caption = "IRM Trans Settings"
            
            'Show the frame
            ShowFrame "15"
            
        Case "15"
            
            'Set the prior frame = "14"
            PriorFrame = "14"
            
            'Need to show additional settings message page
            
            'Just Showframe
            ShowFrame "16"
        
        Case "16"
            
            'Set the prior frame = "15"
            PriorFrame = "15"
            
            'Need to just show the additional AF settings page
            ShowFrame "17"
            
        Case "17"
            
            'Set the prior frame = "16"
            PriorFrame = "16"
            
            'Make sure button says "Next >>"
            Me.cmdNextFinishedSkip.Caption = "Next >>"
        
            'Need to just show the file settings page
            ShowFrame "18"
            
        Case "18"
            
            'Set the prior frame = "17"
            PriorFrame = "17"
            
            'Need to just show please wait frame
            ShowFrame "Wait"
            
            'Load all remaining settings into the new INI file
            LoadRemainingSettings
            
            'Change "Next >>" button to "- Finished - " button
            Me.cmdNextFinishedSkip.Caption = "- Finished -"
            
            'Show the finished process page
            ShowFrame "20"
            
        Case "20"
            
            'Need to unload this form
            Unload Me
            
    End Select
            
End Sub

                    
                    

Private Sub Form_Load()
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    'Load the AF System Combo-box
    Me.cmbAFSystem.AddItem "2G", 0
    Me.cmbAFSystem.AddItem "ADWIN", 1
    
    'Load the IRM System Combo-box
    Me.cmbIRMSystem.AddItem "Caltech Old", 0
    Me.cmbIRMSystem.AddItem "ASC Scientific", 1
    
    'Show the Quit & Exit button and the Next (Start) button
    Me.cmdEndProgram.Visible = True
    Me.cmdNextFinishedSkip.Visible = True
    Me.cmdNextFinishedSkip.Caption = "< Start >"
    
    'Hide the back button
    Me.cmdBack.Visible = False

    'Allocate the DefaultsINI global variable that contains the CINIFile object for the Default.INI file
    modConfig.Allocate_DefaultINI
    
    'Create the new INI file
    modConfig.Create_NewINIFile
    
    'Show the first frame
    ShowFrame "1"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Need to deallocate SystemBoards, SystemAssignedChannels, and WaveForms - these
    'will be reloaded when the code starts up again
    Set SystemBoards = Nothing
    Set SystemAssignedChannels = Nothing
    Set WaveForms = Nothing
    
    'Resume control to modProg.Main by setting INIConversionDone flag to true
    modProg.INIConversionDone = True
    
    'Hide this Form
    Me.Hide
    
End Sub

'This user interface - using the combo-box to adjust the values in column 4 of the
'grid showing the channel assignments is NOT working.  Grrrrrrrr!
'It's also not needed.  A more effective way to implement this would be
'to pop-up a small picture box with a more lables, two combo boxes (one for the board, one for the channel)
' and a text-box showing the resulting Channel INI file string that would be written into column 4 of the grid.
' This picture box would have "Cancel" & "Apply" buttons in it, negating the need for complex error handling
'
'Private Sub gridBoardSettings_Click()
'
'    Dim i, j As Long
'    Dim N As Long
'    Dim ChanINIStr As String
'    Dim BoardName As String
'
'    'Get currently active cell
'    With gridBoardSettings
'
'
'        ActiveCell(0) = .row
'        ActiveCell(1) = .Col
'
'        'Only pop-up the combo-box if the user clicks on column 4
'        If .Col <> 4 Or _
'           .row = 0 Or _
'           .TextMatrix(.row, .Col) = "obsolete" _
'        Then
'
'            'Check to see if the prior active cell was in Column 4
'            If ActiveCell(1) = 4 And _
'               ActiveCell(0) > 0 And _
'               .TextMatrix(ActiveCell(0), ActiveCell(1)) <> "obsolete" And _
'               Me.cmbCellSelector.ListCount > 3 _
'            Then
'
'                .TextMatrix(ActiveCell(0), ActiveCell(1)) = Me.cmbCellSelector.text
'
'                'Set already handled = true, don't need to transfer the value from the combo box to the cell
'                'using the LostFocus event
'                AlreadyHandled = True
'
'            End If
'
'            'Store -1's - this will short circuit the combo-box event handlers
'            ActiveCell(0) = -1
'            ActiveCell(1) = -1
'
'            'Hide the selector combo-box
'            Me.cmbCellSelector.Visible = False
'
'            Exit Sub
'
'        End If
'
'        'Check to see if the prior active cell was in Column 4
'        If ActiveCell(1) = 4 And _
'           ActiveCell(0) > 0 And _
'           .TextMatrix(ActiveCell(0), ActiveCell(1)) <> "obsolete" And _
'           Me.cmbCellSelector.ListCount > 3 _
'        Then
'
'            .TextMatrix(ActiveCell(0), ActiveCell(1)) = Me.cmbCellSelector.text
'
'            'Set already handled = true, don't need to transfer the value from the combo box to the cell
'            'using the LostFocus event
'            AlreadyHandled = True
'
'        Else
'
'            'Need to default already handled to false
'            'The user hasn't had the opportunity to click on the combo box yet
'            AlreadyHandled = False
'
'        End If
'
'        'Set the combo-box selector position
'        Me.cmbCellSelector.ZOrder (0)
'        Me.cmbCellSelector.Left = .CellLeft + .Left + 100
'        Me.cmbCellSelector.Top = .CellTop + .Top - 50
'
'        'Set the contents of the combo-box
'        'Crawl through col-4 and see what channels are available for this channel type
'        'and this board
'        Me.cmbCellSelector.Clear
'        Me.cmbCellSelector.AddItem .text, 0
'        Me.cmbCellSelector.AddItem "------", 1
'
'        j = 2
'
'        'Get the needed board object's name from the INI file
'        BoardName = Config_GetSetting("Boards", "BoardName" & Trim(.TextMatrix(.row, 3)), "-1")
'
'        'If no valid name could be retrieved, exit the subroutine
'        If BoardName = "-1" Then Exit Sub
'
'        'Using the newly retreived board name, get the current board object to use
'        Set CurBoard = SystemBoards(BoardName)
'
'        Select Case Trim(.TextMatrix(.row, 2))
'
'            Case "Analog In"
'
'                On Error GoTo wasError:
'
'                    N = CurBoard.AInChannels.Count
'
'                    If N > 0 Then
'
'                        For i = 1 To N
'
'                            'Get the channel's INI string
'                            ChanINIStr = modConfig.Create_INIChanStr( _
'                                            CurBoard.AInChannels(i))
'
'                            'Check to see if this channel is already listed
'                            'in the grid under the "New Chan. ID" column
''                            If ChannelFree(ChanINIStr) Then
'
'                                'Add the newfound channel to the cmbCellselector
'                                Me.cmbCellSelector.AddItem ChanINIStr, j
'
'                                'Increment the combo-box index counter
'                                j = j + 1
'
''                            End If
'
'                        Next i
'
'                    End If
'
'                On Error GoTo 0
'
'
'            Case "Analog Out"
'
'                On Error GoTo wasError:
'
'                    N = CurBoard.AOutChannels.Count
'
'                    If N > 0 Then
'
'                        For i = 1 To N
'
'                            'Get the channel's INI string
'                            ChanINIStr = modConfig.Create_INIChanStr( _
'                                            CurBoard.AOutChannels(i))
'
'                            'Check to see if this channel is already listed
'                            'in the grid under the "New Chan. ID" column
''                            If ChannelFree(ChanINIStr) Then
'
'                                'Add the newfound channel to the cmbCellselector
'                                Me.cmbCellSelector.AddItem ChanINIStr, j
'
'                                'Increment the combo-box index counter
'                                j = j + 1
'
''                            End If
'
'                        Next i
'
'                    End If
'
'                On Error GoTo 0
'
'
'            Case "Dig. In"
'
'                On Error GoTo wasError:
'
'                    N = CurBoard.DInChannels.Count
'
'                    If N > 0 Then
'
'                        For i = 1 To N
'
'                            'Get the channel's INI string
'                            ChanINIStr = modConfig.Create_INIChanStr( _
'                                            CurBoard.DInChannels(i))
'
'                            'Check to see if this channel is already listed
'                            'in the grid under the "New Chan. ID" column
''                            If ChannelFree(ChanINIStr) Then
'
'                                'Add the newfound channel to the cmbCellselector
'                                Me.cmbCellSelector.AddItem ChanINIStr, j
'
'                                'Increment the combo-box index counter
'                                j = j + 1
'
''                            End If
'
'                        Next i
'
'                    End If
'
'                On Error GoTo 0
'
'
'            Case "Dig. Out"
'
'                On Error GoTo wasError:
'
'                    N = CurBoard.DOutChannels.Count
'
'                    If N > 0 Then
'
'                        For i = 1 To N
'
'                            'Get the channel's INI string
'                            ChanINIStr = modConfig.Create_INIChanStr( _
'                                            CurBoard.DOutChannels(i))
'
'                            'Check to see if this channel is already listed
'                            'in the grid under the "New Chan. ID" column
''                            If ChannelFree(ChanINIStr) Then
'
'                                'Add the newfound channel to the cmbCellselector
'                                Me.cmbCellSelector.AddItem ChanINIStr, j
'
'                                'Increment the combo-box index counter
'                                j = j + 1
'
''                            End If
'
'                        Next i
'
'                    End If
'
'                On Error GoTo 0
'
'        End Select
'
'        'Resize the combo box to accomodate it's text contents
'        ResizeComboBox Me, Me.cmbCellSelector
'
'        'Set the current listindex for the combo box
'        Me.cmbCellSelector.ListIndex = 0
'        CurCmbIndex = 0
'
'        'Make the combo-box visible
'        Me.cmbCellSelector.Visible = True
'
'    End With
'
'wasError:
'
'End Sub

Private Sub gridBoardSettings_LostFocus()

    ActiveCell(0) = -1
    ActiveCell(1) = -1

    Me.cmbCellSelector.Visible = False

End Sub

Private Sub LoadBoardAndChan(ByVal BoardName As String, _
                             ByVal ChanName As String, _
                             ByVal ChanType As String, _
                             ByRef cmbBoard As ComboBox, _
                             ByRef cmbChan As ComboBox)
                             
    Dim i As Integer
    
    'Clauses to exit sub if system boards object is not loaded yet
    If SystemBoards Is Nothing Then Exit Sub
    If SystemBoards.Count < 1 Then Exit Sub
    
    
    'Clear the combo-boxes
    If Not cmbBoard Is Nothing Then
        
        cmbBoard.Clear
        
    End If
    
    If Not cmbChan Is Nothing Then
    
        cmbChan.Clear
        
    End If
        
        
    If Not cmbBoard Is Nothing Then
        
        'Populate the Board Combo-box
        For i = 1 To SystemBoards.Count
        
            cmbBoard.AddItem SystemBoards(i).BoardName, i - 1
            
        Next i
        
        'Select the indicated boardname
        For i = 0 To cmbBoard.ListCount - 1
        
            If cmbBoard.List(i) = BoardName Then
            
                cmbBoard.ListIndex = i
                
                Exit For
                
            End If
            
        Next i
        
    End If
         
    With SystemBoards(BoardName)
                    
        Select Case ChanType
        
            Case "AI"
                    
                'Populate Channel combobox
                For i = 1 To .AInChannels.Count
                
                    'Load in the channels names from the System Boards object
                    cmbChan.AddItem .AInChannels(i).ChanName, i - 1
                    
                Next i
                
            Case "AO"
                    
                'Populate Channel combobox
                For i = 1 To .AOutChannels.Count
                
                    'Load in the channels names from the System Boards object
                    cmbChan.AddItem .AOutChannels(i).ChanName, i - 1
                    
                Next i
                
            Case "DI"
                    
                'Populate Channel combobox
                For i = 1 To .DInChannels.Count
                
                    'Load in the channels names from the System Boards object
                    cmbChan.AddItem .DInChannels(i).ChanName, i - 1
                    
                Next i
                
            Case "DO"
                    
                'Populate Channel combobox
                For i = 1 To .DOutChannels.Count
                
                    'Load in the channels names from the System Boards object
                    cmbChan.AddItem .DOutChannels(i).ChanName, i - 1
                    
                Next i
                
        End Select
        
    End With
                
    'If the default (inputed) channel name = Channel name in the comobobox
    'then set that index as the listed one
    For i = 0 To cmbChan.ListCount - 1
    
        If ChanName = cmbChan.List(i) Then
        
            cmbChan.ListIndex = i
            
            Exit For
            
        End If
        
    Next i
                             
End Sub

Private Sub LoadRemainingSettings()

    Dim SectionStr As String
    
    'Copy out the [Program],[SampleChanger],[MotorPrograms],[MagnetometerCalibration],
    '[ARM],[COMPorts],[Email],[SusceptibilityCalibration],[Magnetometery], &
    '[RockmagRoutineDefaults] sections in their entirety to the new INI file.
    SectionStr = OldIniFile.SectionRead(True, False, "Program")
    IniFile.SectionWrite SectionStr, "Program"
    
    SectionStr = OldIniFile.SectionRead(True, False, "SampleChanger")
    IniFile.SectionWrite SectionStr, "SampleChanger"
    
    SectionStr = OldIniFile.SectionRead(True, False, "MotorPrograms")
    IniFile.SectionWrite SectionStr, "MotorPrograms"
    
    SectionStr = OldIniFile.SectionRead(True, False, "MagnetometerCalibration")
    IniFile.SectionWrite SectionStr, "MagnetometerCalibration"
    
    SectionStr = OldIniFile.SectionRead(True, False, "ARM")
    IniFile.SectionWrite SectionStr, "ARM"
    
    SectionStr = OldIniFile.SectionRead(True, False, "COMPorts")
    IniFile.SectionWrite SectionStr, "COMPorts"
    
    SectionStr = OldIniFile.SectionRead(True, False, "Email")
    IniFile.SectionWrite SectionStr, "Email"
    
    SectionStr = OldIniFile.SectionRead(True, False, "SusceptibilityCalibration")
    IniFile.SectionWrite SectionStr, "SusceptibilityCalibration"
    
    SectionStr = OldIniFile.SectionRead(True, False, "Magnetometery")
    IniFile.SectionWrite SectionStr, "Magnetometry"     'Note Settings Section spelling correction
                                                        'for the section name in the new INI file
    
    SectionStr = OldIniFile.SectionRead(True, False, "RockmagRoutineDefaults")
    IniFile.SectionWrite SectionStr, "RockmagRoutineDefaults"
    
    
    'Now - need to copy out the [SteppingMotor] section and then delete the IRMHiPos setting
    SectionStr = OldIniFile.SectionRead(True, False, "SteppingMotor")
    IniFile.SectionWrite SectionStr, "SteppingMotor"
    
    'Delete IRMHiPos setting
    IniFile.EntryDelete "IRMHiPos", "SteppingMotor"
    
    'Need to copy the DoVacuumReset setting in the [Vacuum] section
    IniFile.EntryWrite "DoVacuumReset", _
                       OldIniFile.EntryRead("DoVacuumReset", _
                                            "False", _
                                            "Vacuum"), _
                       "Vacuum"
                       
    'Need to copy the DoDegausserCooling setting in the [Vacuum] section
    IniFile.EntryWrite "DoDegausserCooling", _
                       OldIniFile.EntryRead("DoDegausserCooling", _
                                            "False", _
                                            "Vacuum"), _
                       "Vacuum"
                       
    'Now need to transfer the module settings from the list-view control to the new inifile
    With Me.lvwEnabledModules
    
        IniFile.EntryWrite "EnableAF", _
                           Trim(Str(.ListItems("EnableAF").Checked)), _
                           "Modules"
        
        IniFile.EntryWrite "EnableAFAnalysis", _
                           Trim(Str(.ListItems("EnableAFAnalysis").Checked)), _
                           "Modules"
        
        IniFile.EntryWrite "EnableAltAFMonitor", _
                           Trim(Str(.ListItems("EnableAltAFMonitor").Checked)), _
                           "Modules"
        
        IniFile.EntryWrite "EnableAxialIRM", _
                           Trim(Str(.ListItems("EnableAxialIRM").Checked)), _
                           "Modules"
        
        IniFile.EntryWrite "EnableTransIRM", _
                           Trim(Str(.ListItems("EnableTransIRM").Checked)), _
                           "Modules"
        
        IniFile.EntryWrite "EnableIRMBackfield", _
                           Trim(Str(.ListItems("EnableIRMBackfield").Checked)), _
                           "Modules"
        
        IniFile.EntryWrite "EnableIRMMonitor", _
                           Trim(Str(.ListItems("EnableIRMMonitor").Checked)), _
                           "Modules"
        
        IniFile.EntryWrite "EnableARM", _
                           Trim(Str(.ListItems("EnableARM").Checked)), _
                           "Modules"
        
        IniFile.EntryWrite "EnableSusceptibility", _
                           Trim(Str(.ListItems("EnableSusceptibility").Checked)), _
                           "Modules"
        
        IniFile.EntryWrite "EnableT1", _
                           Trim(Str(.ListItems("EnableT1").Checked)), _
                           "Modules"
        
        IniFile.EntryWrite "EnableT2", _
                           Trim(Str(.ListItems("EnableT2").Checked)), _
                           "Modules"
                           
    End With
    
    
    
    'Save the AF Data file settings
    IniFile.EntryWrite "ADWINDataFileSaveLocalDir", _
                       Me.txtAFLocalDataFolder.text, _
                       "AFFileSave"
    IniFile.EntryWrite "ADWINDataFileSaveBackupDir", _
                       Me.txtAFBackupDataFolder.text, _
                       "AFFileSave"
    IniFile.EntryWrite "2GDataFileSaveLocalDir", _
                       Me.txtAFLocalDataFolder.text, _
                       "AFFileSave"
    IniFile.EntryWrite "2GDataFileSaveBackupDir", _
                       Me.txtAFBackupDataFolder.text, _
                       "AFFileSave"
                       
    IniFile.EntryWrite "AFDataFileSaveDoBackup", _
                       Trim(Str(Me.chkDoBackup.Value = Checked)), _
                       "AFFileSave"
    
    'Now save the remaining 6 channels next yet saved to the INI File
    SaveRemaining6Channels
        
    'Now put line-breaks between sections in the INI file
    IniFile.AddLineBreaks
        
    'Done!  Yay!
    'The new inifile is ready to go!

End Sub

Private Sub LoadWaveFormSettings(ByRef gridobj As MSHFlexGrid)

    Dim i As Integer
    Dim j As Integer
    
    'Clear the grid object
    gridobj.Clear
    gridobj.ClearStructure
    
    If WaveForms Is Nothing Then Exit Sub
    If WaveForms.Count < 1 Then Exit Sub
    
    'For each waveform, there are 12 settings
    'for each setting, there are three columns
    'inbetween each waveform setting block are 3 rows for a new header
    'also need to remove one rows - because the first header doesn't need a blank row above it
    'Total # of rows = 1 + (# waveforms) * 9 + (# waveforms) * 3 - 2
    With gridobj
    
        'Set grid dimensions
        .Rows = 1 + 9 * WaveForms.Count + 3 * WaveForms.Count - 2
        .Cols = 3
        
        'Set the fixed rows and cols
        .FixedRows = 0
        .FixedCols = 0
        
        'Merge the necessary header rows
        .MergeCells = flexMergeRestrictRows
        .MergeRow(0) = True
        For i = 1 To WaveForms.Count - 1
            
            .MergeRow((i * 12) - 1) = True
            .MergeRow(i * 12) = True
            
        Next i
        
        'Create headers
        For j = 0 To WaveForms.Count - 1
            
            If j > 0 Then
                
                'Set the values for the blank rows
                .row = (j * 12) - 1
                
                For i = 0 To .Cols - 1
                        
                    .Col = i
                    .text = "     "
                    .CellBackColor = &HC0C0C0
            
                Next i
            
            End If
            
            'Set the values for the Waveform name header
            .row = j * 12
            
            For i = 0 To .Cols - 1
            
                .Col = i
                .text = WaveForms(j + 1).WaveName & ", " & WaveForms(j + 1).WaveDesc
                .CellBackColor = &H8000000F
                
            Next i
            
            'Set the values for the Column headers
            .TextMatrix(1 + (j * 12), 0) = "Setting Name"
            .TextMatrix(1 + (j * 12), 1) = "Setting Value"
            .TextMatrix(1 + (j * 12), 2) = "Description"
            .row = (j * 12) + 1
            
            For i = 0 To .Cols - 1
                
                .Col = i
                .CellBackColor = &H8000000F
                
            Next i
                
        Next j
        
    End With
    
    'All the header rows have been set, just need to load in the data values now
    For i = 0 To WaveForms.Count - 1
    
        With WaveForms(i + 1)
        
            'WaveININum
            gridobj.TextMatrix(i * 12 + 2, 0) = "WaveININum"
            gridobj.TextMatrix(i * 12 + 2, 1) = Trim(Str(.WaveININum))
            gridobj.TextMatrix(i * 12 + 2, 2) = "Wave object INI-file ID #"
            
            'WaveName
            gridobj.TextMatrix(i * 12 + 3, 0) = "WaveName"
            gridobj.TextMatrix(i * 12 + 3, 1) = .WaveName
            gridobj.TextMatrix(i * 12 + 3, 2) = "Wave's string ID used by Paleomag code"
            
            'BoardUsed
            gridobj.TextMatrix(i * 12 + 4, 0) = "BoardUsed"
            gridobj.TextMatrix(i * 12 + 4, 1) = .BoardUsed.BoardName
            gridobj.TextMatrix(i * 12 + 4, 2) = "String ID of DAQ Board used by Wave"
        
            'Channel
            gridobj.TextMatrix(i * 12 + 5, 0) = "Chan"
            gridobj.TextMatrix(i * 12 + 5, 1) = .Chan.ChanName
            gridobj.TextMatrix(i * 12 + 5, 2) = "String ID of Channel used by Wave"
            
            'IO direction
            gridobj.TextMatrix(i * 12 + 6, 0) = "IO"
            gridobj.TextMatrix(i * 12 + 6, 1) = .IO
            gridobj.TextMatrix(i * 12 + 6, 2) = "Input/Output direction of Wave"
            
            'IO Rate
            gridobj.TextMatrix(i * 12 + 7, 0) = "IORate"
            gridobj.TextMatrix(i * 12 + 7, 1) = Trim(Str(.IORate))
            gridobj.TextMatrix(i * 12 + 7, 2) = "A/D or D/A data conversion rate in Hz"
                            
            'Do Deallocate
            gridobj.TextMatrix(i * 12 + 8, 0) = "DoDeallocate"
            gridobj.TextMatrix(i * 12 + 8, 1) = Trim(Str(.DoDeallocate))
            gridobj.TextMatrix(i * 12 + 8, 2) = "Delete points from memory when Wave is done?"
            
            'Range Max
            gridobj.TextMatrix(i * 12 + 9, 0) = "RangeMax"
            gridobj.TextMatrix(i * 12 + 9, 1) = Trim(Str(.range.MaxValue))
            gridobj.TextMatrix(i * 12 + 9, 2) = "Max DAQ I/O voltage for Wave process"
            
            'Range Min
            gridobj.TextMatrix(i * 12 + 10, 0) = "RangeMin"
            gridobj.TextMatrix(i * 12 + 10, 1) = Trim(Str(.range.MinValue))
            gridobj.TextMatrix(i * 12 + 10, 2) = "Min DAQ I/O voltage for Wave process"
            
        End With
        
    Next i
            
End Sub

Private Sub lvwEnabledModules_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    'The purpose of this event handler is to prevent an item in the Modules list
    'view from being successfully checked if it's "Ghosted" property is set to true
    If Item.Ghosted = True Then Item.Checked = False

End Sub

Private Sub optAFCalDone_Click()

    If Me.optAFCalDone.Value = True And _
       Me.optAFNeedsCal.Value = True _
    Then
    
        Me.optAFNeedsCal.Value = False
        
    End If

End Sub

Private Sub optAFNeedsCal_Click()

    If Me.optAFCalDone.Value = True And _
       Me.optAFNeedsCal.Value = True _
    Then
    
        Me.optAFCalDone.Value = False
        
    End If

End Sub

Private Sub optIRMCalDone_Click()

    If Me.optIRMCalDone.Value = True And _
       Me.optIRMNeedsCalibration.Value = True _
    Then
    
        Me.optIRMNeedsCalibration.Value = False
        
    End If

End Sub

Private Sub optIRMNeedsCalibration_Click()

    If Me.optIRMCalDone.Value = True And _
       Me.optIRMNeedsCalibration.Value = True _
    Then
    
        Me.optIRMCalDone.Value = False
        
    End If

End Sub

Private Sub optTrimOnFalse_Click()

    If Me.optTrimOnFalse.Value = True And _
       Me.optTrimOnTrue.Value = True _
    Then
    
        Me.optTrimOnTrue.Value = False
        
    End If

End Sub

Private Sub optTrimOnTrue_Click()

    If Me.optTrimOnFalse.Value = True And _
       Me.optTrimOnTrue.Value = True _
    Then
    
        Me.optTrimOnFalse.Value = False
        
    End If

End Sub

Public Sub ResizeComboBox(ByRef FormObj As Form, _
                          ByRef cmbObj As ComboBox, _
                          Optional ByVal FirstItem As Integer = 1, _
                          Optional ByVal LastItem As Integer = -1, _
                          Optional ByVal ScalingFactor As Double = 1.5)
                          
    Dim i As Integer
    Dim LargestText As Integer
    
    'Coercing LastItem to an appropriate value
    If LastItem < 1 Then LastItem = cmbObj.ListCount
    
    'Error Checking
    If FirstItem > LastItem Then Exit Sub
    
    For i = FirstItem To LastItem
    
        If LargestText < FormObj.TextWidth(cmbObj.List(i)) * ScalingFactor Then
    
            LargestText = FormObj.TextWidth(cmbObj.List(i)) * ScalingFactor
            
        End If
        
    Next i
    
    cmbObj.Width = LargestText * 1.5
                          
End Sub

Private Sub SaveRemaining6Channels()

    Dim TempChannel As Channel
    Dim ChanINIStr As String

    'Allocate TempChannel
    Set TempChannel = New Channel

    'Need to snatch the default values for the Analog Thermal sensor channels from the Defaults.INI file
    'and write those values to the new INI file
    modConfig.Config_SaveSetting "Channels", _
                                 "AnalogT1", _
                                 DefaultINI.EntryRead("AnalogT1", _
                                                      "AI-0-CH3", _
                                                      "Channels")
                                                      
    modConfig.Config_SaveSetting "Channels", _
                                 "AnalogT2", _
                                 DefaultINI.EntryRead("AnalogT2", _
                                                      "AI-0-CH4", _
                                                      "Channels")
                                                      
    'Need to snatch the IRMPowerAmpVoltageIn channel infor from the Defaults.INI file
    modConfig.Config_SaveSetting "Channels", _
                                 "IRMPowerAmpVoltageIn", _
                                 DefaultINI.EntryRead("IRMPowerAmpVoltageIn", _
                                                      "AI-0-CH1", _
                                                      "Channels")
                                                      
                                                      
    'If this user has updated the three relay channel combo-boxes on Frame 2A, then
    'need to translate those channel names to Channel INI strings
    If Me.cmbAFAxialRelay.text <> "" Then
    
        'Attempt to convert the text in that combo-box into an INI channel string
        'First need to create a channel object with the correct channel name and the correct board name
        'and the correct channel type
        TempChannel.ChanName = Me.cmbAFAxialRelay.text
        TempChannel.ChanType = "DO"
        TempChannel.BoardName = WaveForms("AFRAMPUP").BoardUsed.BoardName
        
        'Now use the CreateINIChannelStr function
        ChanINIStr = modConfig.Create_INIChanStr(TempChannel)
        
        'Validate that chanINIstr
        If ChanINIStr <> "ERROR" Then
        
            modConfig.Config_SaveSetting "Channels", _
                                         "AFAxialRelay", _
                                         ChanINIStr
                                            
        Else
        
            'Load value from Defaults.INI file
            modConfig.Config_SaveSetting "Channels", _
                                         "AFAxialRelay", _
                                         DefaultINI.EntryRead("AFAxialRelay", _
                                                              "DO-1-CH2", _
                                                              "Channels")
                                                              
        End If
        
    Else
    
        'Load value from Defaults.INI file
            modConfig.Config_SaveSetting "Channels", _
                                         "AFAxialRelay", _
                                         DefaultINI.EntryRead("AFAxialRelay", _
                                                              "DO-1-CH2", _
                                                              "Channels")
                                                              
    End If
    
    If Me.cmbAFAxialRelay.text <> "" Then
    
        'Attempt to convert the text in that combo-box into an INI channel string
        'First need to create a channel object with the correct channel name and the correct board name
        'and the correct channel type
        TempChannel.ChanName = Me.cmbAFTransRelay.text
        TempChannel.ChanType = "DO"
        TempChannel.BoardName = WaveForms("AFRAMPUP").BoardUsed.BoardName
        
        'Now use the CreateINIChannelStr function
        ChanINIStr = modConfig.Create_INIChanStr(TempChannel)
        
        'Validate that chanINIstr
        If ChanINIStr <> "ERROR" Then
        
            modConfig.Config_SaveSetting "Channels", _
                                         "AFTransRelay", _
                                         ChanINIStr
                                            
        Else
        
            'Load value from Defaults.INI file
            modConfig.Config_SaveSetting "Channels", _
                                         "AFTransRelay", _
                                         DefaultINI.EntryRead("AFTransRelay", _
                                                              "DO-1-CH1", _
                                                              "Channels")
                                                              
        End If
        
    Else
    
        'Load value from Defaults.INI file
        modConfig.Config_SaveSetting "Channels", _
                                     "AFTransRelay", _
                                     DefaultINI.EntryRead("AFTransRelay", _
                                                          "DO-1-CH1", _
                                                          "Channels")
                                                          
    End If
    
    If Me.cmbIRMRelay.text <> "" Then
    
        'Attempt to convert the text in that combo-box into an INI channel string
        'First need to create a channel object with the correct channel name and the correct board name
        'and the correct channel type
        TempChannel.ChanName = Me.cmbIRMRelay.text
        TempChannel.ChanType = "DO"
        TempChannel.BoardName = WaveForms("AFRAMPUP").BoardUsed.BoardName
        
        'Now use the CreateINIChannelStr function
        ChanINIStr = modConfig.Create_INIChanStr(TempChannel)
        
        'Validate that chanINIstr
        If ChanINIStr <> "ERROR" Then
        
            modConfig.Config_SaveSetting "Channels", _
                                         "IRMRelay", _
                                         ChanINIStr
                                            
        Else
        
            'Load value from Defaults.INI file
            modConfig.Config_SaveSetting "Channels", _
                                         "IRMRelay", _
                                         DefaultINI.EntryRead("IRMRelay", _
                                                              "DO-1-CH0", _
                                                              "Channels")
                                                              
        End If
        
    Else
    
        'Load value from Defaults.INI file
        modConfig.Config_SaveSetting "Channels", _
                                     "IRMRelay", _
                                     DefaultINI.EntryRead("IRMRelay", _
                                                          "DO-1-CH0", _
                                                          "Channels")
                                                          
    End If
    

End Sub

Public Sub SetCellColor(ByRef gridobj As MSHFlexGrid, _
                        ByVal RowNum As Integer, _
                        ByVal ColNum As Integer, _
                        Optional ByVal ForeColor As Long = &H80000008, _
                        Optional ByVal BackColor As Long = &H80000005)
                        
    With gridobj
    
        On Error GoTo wasError:
    
            .row = RowNum
            .Col = ColNum
            
            .CellForeColor = ForeColor
            .CellBackColor = BackColor
            
        On Error GoTo 0
        
    End With

wasError:
                        
End Sub

Public Sub SetColColor(ByRef gridobj As MSHFlexGrid, _
                       ByVal ColNum As Integer, _
                       Optional ByVal ForeColor As Long = &H80000008, _
                       Optional ByVal BackColor As Long = &H80000005)
                        
    Dim i As Integer
    
    With gridobj
    
        On Error GoTo wasError:
    
            .Col = ColNum
            
            For i = 0 To .Rows - 1
            
                .row = i
                .CellForeColor = ForeColor
                .CellBackColor = BackColor
                
            Next i
            
        On Error GoTo 0
        
    End With
        
wasError:
                        
End Sub

Public Sub SetRowColor(ByRef gridobj As MSHFlexGrid, _
                       ByVal RowNum As Integer, _
                       Optional ByVal ForeColor As Long = &H80000008, _
                       Optional ByVal BackColor As Long = &H80000005)
                        
    Dim i As Integer
    
    With gridobj
    
        On Error GoTo wasError:
    
            .row = RowNum
            
            For i = 0 To .Cols - 1
            
                .Col = i
                .CellForeColor = ForeColor
                .CellBackColor = BackColor
                
            Next i
            
        On Error GoTo 0
        
    End With
        
wasError:
                        
End Sub

Private Sub ShowFrame(ByVal FrameIndex As String)

    Dim StepNum As Long

    'Hide all the frames
    Frame1.Visible = False
    Frame2.Visible = False
    Frame2A.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
    Frame8.Visible = False
    Frame8A.Visible = False
    Frame8B.Visible = False
    Frame10.Visible = False
    Frame16.Visible = False
    Frame17.Visible = False
    Frame18.Visible = False
    FrameWait.Visible = False
    Frame20.Visible = False

    'Show the asked for frame
    Select Case FrameIndex
    
        Case "1"
        
            'Welcome / start INI upgrade now?
            Frame1.Visible = True
            
            StepNum = 0
        
        Case "2"
        
            'Select AF System
            Frame2.Visible = True
            
            StepNum = 1
            
        Case "2A"
        
            'Set port assignments for Relays if ADWIN AF system
            Frame2A.Visible = True
            
            StepNum = 2

        Case "3"
            
            'Select IRM System
            Frame3.Visible = True
            
            StepNum = 3
            
        Case "3A"
        
            'Select the IRM Trim On True value
            Frame8.Visible = True
            StepNum = 4
            
        Case "4"
        
            'Boards required for install
            Frame4.Visible = True

            StepNum = 5
            
        Case "5"

            'Select modules to enable
            Frame5.Visible = True

            StepNum = 6
            
        Case "6"
        
            'PCI-DAS6030 board settings
            Frame6.Visible = True

            StepNum = 7
                        
        Case "6B"
        
            'Existing Port Assignments for
            'PCI-DAS6030 board
            Frame6.Visible = True

            StepNum = 8
                    
        Case "7"
            
            'Msg: Will setup ADWIN board, anyway
            '(so there!)
            Frame7.Visible = True

            StepNum = 9
                        
        Case "8"
        
            'Reusing Frame 6 to show the ADWIN Board Settings
            Frame6.Visible = True

            StepNum = 10
                        
        Case "8A"
        
            'Show ADWIN channel setting frame
            Frame8A.Visible = True

            StepNum = 11
                        
        Case "8B"
        
            'Show AF / IRM extra channels setting frame
            Frame8B.Visible = True

            StepNum = 12
                        
        Case "8C"
        
            'Reusing Frame 6 to show the Wave Form Object Settings
            Frame6.Visible = True

            StepNum = 13
                        
        Case "9"
        
            'INI File settings that are getting trashed
            Frame10.Visible = True

            StepNum = 14
                        
        Case "10"
        
            'AF Main settings being set
            Frame6.Visible = True

            StepNum = 15
                        
        Case "11"
        
            'AF Axial Settings being set
            Frame6.Visible = True

            StepNum = 16
                        
        Case "12"
        
            'AF Transverse settings being set
            Frame6.Visible = True

            StepNum = 17
                        
        Case "13"
        
            'IRM Main settings being set
            Frame6.Visible = True

            StepNum = 18
                        
        Case "14"
        
            'IRM Axial Settings being set
            Frame6.Visible = True

            StepNum = 19
                        
        Case "15"
        
            'IRM Transverse Settings being set
            Frame6.Visible = True

            StepNum = 20
            
        Case "16"
        
            'Msg: Additional information to set
            Frame16.Visible = True

            StepNum = 21
                                
        Case "17"
        
            'AF coil settings (res freq & clipping voltages)
            Frame17.Visible = True

            StepNum = 22
            
        Case "18"
        
            'AF Data File settings
            Frame18.Visible = True
 
            StepNum = 23
                       
        Case "19"
        
            'Final window displaying all settings being
            'transfered and created
            Frame19.Visible = True

            StepNum = 24
                        
        Case "20"
        
            'Upgrade done window
            Frame20.Visible = True
 
            StepNum = 25
                           
        Case "Wait"
        
            FrameWait.Visible = True
            
        Case Else
        
            'If bad active frame set,
            'jump back to start
            Frame1.Visible = True

            StepNum = 0
                        
    End Select
    
    'Store which frame is showing
    'Never set active frame to the Wait frame - that's just a temporary frame
    If FrameIndex <> "Wait" Then
    
        ActiveFrame = FrameIndex
        
        UpdateProgress StepNum
        
    End If

End Sub

'IMPORTANT NOTE: When calling this function, the only two valid inputs are:
'Axial & Transverse - created an enum type to ensure this, any other inputs
'will default to transverse and maybe mess stuff up.
Private Sub TransferAFAxisSettings(ByVal AxisObj As TransferAxisType)

    Dim CalDone As String
    Dim i As Long
    Dim SkipMod As Long
    Dim NumPoints As Long
    Dim Coord As String
    Dim TempArray(25, 1) As Double
    Dim TempD As Double
    Dim TempD2 As Double
    
    If AxisObj = Axial Then
    
        Coord = "Axial"
        
    Else
    
        Coord = "Trans"
        
    End If

    'Transfer the 5 non-calibration value settings
        
    'Figure out if user is changing to the ADWIN AF system
    'No new calibration is needed if the AF system is 2G
    CalDone = Trim(Str(Me.optAFCalDone.Value))
    
    If Coord = "Axial" Then
        
        'Write calibration status of AF system
        IniFile.EntryWrite "AF" & Coord & "CalDone", CalDone, "AF" & Coord
                
        modConfig.AFAxialCalDone = ("True" = CalDone)
                
        'Write Axis Coord
        modConfig.AfAxialCoord = OldIniFile.EntryRead("AF" & Coord & "Coord", _
                                                "ERROR", _
                                                "AF" & Coord)
        
        IniFile.EntryWrite "AF" & Coord & "Coord", _
                           modConfig.AfAxialCoord, _
                           "AF" & Coord
                           
        modConfig.AfAxialMax = val(OldIniFile.EntryRead("AF" & Coord & "Max", _
                                                "-1", _
                                                "AF" & Coord))

        'Write Axis Max
        IniFile.EntryWrite "AF" & Coord & "Max", _
                           Trim(Str(modConfig.AfAxialMax)), _
                           "AF" & Coord
    
    
        modConfig.AfAxialMin = val(OldIniFile.EntryRead("AF" & Coord & "Min", _
                                                "-1", _
                                                "AF" & Coord))

        'Write Axis Min
        IniFile.EntryWrite "AF" & Coord & "Min", _
                           Trim(Str(modConfig.AfAxialMin)), _
                           "AF" & Coord
        
        'First point is zero
        TempArray(0, 0) = 0
        TempArray(0, 1) = 0
        
        'Set SkipMod = 0
        SkipMod = 0     'No data values in the calibration data are being skipped
        
        'Now load the remain non-zero, < 3999 points
        For i = 1 To 25
        
            'Read the calibration pair from the INI file into two local Double-Type variables
            TempD = val(OldIniFile.EntryRead("AFAxialX" & Trim(Str(i)), _
                                                     "-1", _
                                                     "AFAxial"))
            
            TempD2 = val(OldIniFile.EntryRead("AFAxialY" & Trim(Str(i)), _
                                                     "-1", _
                                                     "AFAxial"))
            
            
            'If the 2G value or the magnetic field value is bad, then skip past this data pair
            If TempD > 3999 Or _
               TempD <= 0 Or _
               TempD2 <= 0 _
            Then
            
                'Skip by this data pair - don't add it to the TempArray
                
                'Need to update SkipMod so we don't leave a blank row
                'in TempArray
                SkipMod = SkipMod - 1
                
            Else
            
                'The data pair is good, add it
                TempArray(i + SkipMod, 0) = TempD
                TempArray(i + SkipMod, 1) = TempD2
                
            End If
                                                     
        Next i
        
        'The number of good data points = i + skipmod - 1
        NumPoints = i + SkipMod - 1
        
        'Now dimension the AF Axial array
        ReDim modConfig.AFAxial(NumPoints, 1)
        
        'Load the points from the Temp Array to the AFAxial Array
        ArrayCopy TempArray, modConfig.AFAxial, NumPoints + 1, 2
                    
        'Now save the count setting
        modConfig.AFAxialCount = NumPoints
        IniFile.EntryWrite "AFAxialCount", _
                           NumPoints, _
                           "AFAxial"
                           
        'Now save the X,Y calibration data pairs
        For i = 1 To NumPoints
        
            'Write in the X value
            IniFile.EntryWrite "AFAxialX" & Trim(Str(i)), _
                               Trim(Str(modConfig.AFAxial(i, 0))), _
                               "AFAxial"
                               
            'Write in the Y value
            IniFile.EntryWrite "AFAxialY" & Trim(Str(i)), _
                               Trim(Str(modConfig.AFAxial(i, 1))), _
                               "AFAxial"
                               
        Next i
                    
    Else
    
        'This is the Transverse Axis
        
        'Write calibration status of AF system
        IniFile.EntryWrite "AF" & Coord & "CalDone", CalDone, "AF" & Coord
                
        modConfig.AFTransCalDone = ("True" = CalDone)
                
        'Write Axis Coord
        modConfig.AfTransCoord = OldIniFile.EntryRead("AF" & Coord & "Coord", _
                                                "ERROR", _
                                                "AF" & Coord)
        
        IniFile.EntryWrite "AF" & Coord & "Coord", _
                           modConfig.AfTransCoord, _
                           "AF" & Coord
                           
        modConfig.AfTransMax = val(OldIniFile.EntryRead("AF" & Coord & "Max", _
                                                "-1", _
                                                "AF" & Coord))

        'Write Axis Max
        IniFile.EntryWrite "AF" & Coord & "Max", _
                           Trim(Str(modConfig.AfTransMax)), _
                           "AF" & Coord
    
    
        modConfig.AfTransMin = val(OldIniFile.EntryRead("AF" & Coord & "Min", _
                                                "-1", _
                                                "AF" & Coord))

        'Write Axis Min
        IniFile.EntryWrite "AF" & Coord & "Min", _
                           Trim(Str(modConfig.AfTransMin)), _
                           "AF" & Coord
        
        'First point is zero
        TempArray(0, 0) = 0
        TempArray(0, 1) = 0
        
        'Set SkipMod = 0
        SkipMod = 0     'No data values in the calibration data are being skipped
        
        'Now load the remain non-zero, < 3999 points
        For i = 1 To 25
        
            'Read the calibration pair from the INI file into two local Double-Type variables
            TempD = val(OldIniFile.EntryRead("AFTransX" & Trim(Str(i)), _
                                                     "-1", _
                                                     "AFTrans"))
            
            TempD2 = val(OldIniFile.EntryRead("AFTransY" & Trim(Str(i)), _
                                                     "-1", _
                                                     "AFTrans"))
            
            'If the 2G value or the magnetic field value is bad, then skip past this data pair
            If TempD > 3999 Or _
               TempD <= 0 Or _
               TempD2 <= 0 _
            Then
            
                'Skip by this data pair - don't add it to the TempArray
                
                'Need to update SkipMod so we don't leave a blank row
                'in TempArray
                SkipMod = SkipMod - 1
                
            Else
            
                'The data pair is good, add it
                TempArray(i + SkipMod, 0) = TempD
                TempArray(i + SkipMod, 1) = TempD2
                
            End If
                                                     
        Next i
        
        'The number of good data points = i + skipmod - 1
        NumPoints = i + SkipMod - 1
        
        'Now dimension the AF Trans array
        ReDim modConfig.AFTrans(NumPoints, 1)
        
        'Load the points from the Temp Array to the AFTrans Array
        ArrayCopy TempArray, modConfig.AFTrans, NumPoints + 1, 2
                    
        'Now save the count setting
        modConfig.AFTransCount = NumPoints
        IniFile.EntryWrite "AFTransCount", _
                           NumPoints, _
                           "AFTrans"
                           
        'Now save the X,Y calibration data pairs
        For i = 1 To NumPoints
        
            'Write in the X value
            IniFile.EntryWrite "AFTransX" & Trim(Str(i)), _
                               Trim(Str(modConfig.AFTrans(i, 0))), _
                               "AFTrans"
                               
            'Write in the Y value
            IniFile.EntryWrite "AFTransY" & Trim(Str(i)), _
                               Trim(Str(modConfig.AFTrans(i, 1))), _
                               "AFTrans"
                               
        Next i
        
    End If
    
End Sub

'IMPORTANT NOTE: When calling this function, the only two valid inputs are:
'Axial & Transverse - created an enum type to ensure this, any other inputs
'will default to transverse and maybe mess stuff up.
Private Sub TransferIRMAxisSettings(ByVal AxisObj As TransferAxisType)

    Dim CalDone As String
    Dim i As Long
    Dim NumPoints As Long
    Dim Coord As String
    Dim TempArray(25, 2) As Double
    
    If AxisObj = Axial Then
    
        Coord = "Axial"
        
    Else
    
        Coord = "Trans"
        
    End If
    
    'Transfer the 5 non-calibration value settings
        
    'Figure out if the IRM System needs to be calibrated
    If Me.optIRMCalDone.Value = True Then
    
        CalDone = "True"
        
    Else
    
        CalDone = "False"

    End If
    
    If Coord = "Axial" Then
        
        'Write calibration status of IRM system
        IniFile.EntryWrite "IRM" & Coord & "CalDone", CalDone, "IRM" & Coord
                
        modConfig.IRMAxialCalDone = ("True" = CalDone)
                
                
        'This setting doesn't seem to be in the old INI file. Load it from the default INI file
        modConfig.IRMAxialVoltMax = val(DefaultINI.EntryRead("IRMAxialVoltMax", _
                                                             "450", _
                                                             "IRMAxial"))
                                           
        'Write IRM Axial max voltage to the new INI file
        IniFile.EntryWrite "IRMAxialVoltMax", _
                           Trim(Str(modConfig.IRMAxialVoltMax)), _
                           "IRMAxial"
                                           
        modConfig.PulseAxialMax = val(OldIniFile.EntryRead("PulseMax", _
                                                "-1", _
                                                "IRMPulse"))

        'Write Axis Max
        IniFile.EntryWrite "Pulse" & Coord & "Max", _
                           Trim(Str(modConfig.PulseAxialMax)), _
                           "IRM" & Coord
    
    
        modConfig.PulseAxialMin = val(OldIniFile.EntryRead("PulseMin", _
                                                "-1", _
                                                "IRMPulse"))

        'Write Axis Min
        IniFile.EntryWrite "Pulse" & Coord & "Min", _
                           Trim(Str(modConfig.PulseAxialMin)), _
                           "IRM" & Coord
        
        'First point is zero
        TempArray(0, 0) = 0
        TempArray(0, 1) = 0
        
        'Set SkipMod = 0
        SkipMod = 0     'No data values in the calibration data are being skipped, yet
        
        'Now load the remain non-zero, =< IRM Max capacitor voltage points
        For i = 1 To 25
        
            'Read the calibration pair from the INI file into two local Double-Type variables
            TempD = val(OldIniFile.EntryRead("PulseLFX" & Trim(Str(i)), _
                                                     "-1", _
                                                     "IRMPulse"))
            
            TempD2 = val(OldIniFile.EntryRead("PulseLFY" & Trim(Str(i)), _
                                                     "-1", _
                                                     "IRMPulse"))
            
            'If the Capacitor voltage or the magnetic field value is bad, then skip past this data pair
            If TempD > modConfig.IRMAxialVoltMax Or _
               TempD <= 0 Or _
               TempD2 <= 0 _
            Then
            
                'Skip by this data pair - don't add it to the TempArray
                
                'Need to update SkipMod so we don't leave a blank row
                'in TempArray
                SkipMod = SkipMod - 1
                
            Else
            
                'The data pair is good, add it
                TempArray(i + SkipMod, 0) = TempD
                TempArray(i + SkipMod, 1) = TempD2
                
            End If
                                                     
        Next i
        
        'The number of good data points = i + skipmod - 1
        NumPoints = i + SkipMod - 1
        
        'Now dimension the IRM Axial array
        ReDim modConfig.PulseAxial(NumPoints, 1)
        
        'Load the points from the Temp Array to the AFTrans Array
        ArrayCopy TempArray, modConfig.PulseAxial, NumPoints + 1, 2
                    
        'Now save the count setting
        modConfig.PulseAxialCount = NumPoints
        IniFile.EntryWrite "Pulse" & Coord & "Count", _
                           NumPoints, _
                           "IRM" & Coord
                           
        'Now save the X,Y calibration data pairs
        For i = 1 To NumPoints
        
            'Write in the X value
            IniFile.EntryWrite "Pulse" & Coord & "X" & Trim(Str(i)), _
                               Trim(Str(modConfig.PulseAxial(i, 0))), _
                               "IRM" & Coord
                               
            'Write in the Y value
            IniFile.EntryWrite "Pulse" & Coord & "Y" & Trim(Str(i)), _
                               Trim(Str(modConfig.PulseAxial(i, 1))), _
                               "IRM" & Coord
                               
        Next i
                    
    Else
    
        'This is the Transverse Axis
        
        'Write calibration status of IRM system
        IniFile.EntryWrite "IRM" & Coord & "CalDone", CalDone, "IRM" & Coord
                
        modConfig.IRMTransCalDone = ("True" = CalDone)
                
                
        'This setting doesn't seem to be in the old INI file. Load it from the default INI file
        modConfig.IRMTransVoltMax = val(DefaultINI.EntryRead("IRMTransVoltMax", _
                                                             "400", _
                                                             "IRMTrans"))
                                           
        'Write IRM Trans max voltage to the new INI file
        IniFile.EntryWrite "IRMTransVoltMax", _
                           Trim(Str(modConfig.IRMTransVoltMax)), _
                           "IRMTrans"
                                           
        modConfig.PulseTransMax = 0

        'Write Axis Max
        IniFile.EntryWrite "Pulse" & Coord & "Max", _
                           Trim(Str(modConfig.PulseTransMax)), _
                           "IRM" & Coord
    
    
        modConfig.PulseTransMin = 0

        'Write Axis Min
        IniFile.EntryWrite "Pulse" & Coord & "Min", _
                           Trim(Str(modConfig.PulseTransMin)), _
                           "IRM" & Coord
        
        'There are no Transverse AF calibration points in the old INI file
        'so set count to Zero and only add the zero point to the PulseTrans array
        
        'Redim the modconfig PulseTrans array
        ReDim modConfig.PulseTrans(1, 2)
        
        'First point is zero
        PulseTrans(0, 0) = 0
        PulseTrans(0, 1) = 0
        
        'Now save the count setting
        modConfig.PulseTransCount = 0
        IniFile.EntryWrite "PulseTransCount", _
                           "0", _
                           "IRMTrans"
        
    End If
    
End Sub

Private Sub UpdateProgress(ByVal StepNum As Long)

    Me.progStepsCompleted.min = 0

    If StepNum = 0 Then
    
        Me.progStepsCompleted.Visible = False
        Me.txtStepNum.Visible = False
        
    End If

    If Me.cmbAFSystem.SelText = "2G" Then
    
        Me.progStepsCompleted.Max = 24
        
        If StepNum > 1 Then StepNum = StepNum - 1
        
    Else
    
        Me.progStepsCompleted.Max = 25
        
    End If
    
    Me.progStepsCompleted.Value = StepNum
    
    If StepNum < Me.progStepsCompleted.Max Then
    
        Me.txtStepNum.text = "Step " & Trim(Str(StepNum)) & " of " & _
                             Trim(Str(Me.progStepsCompleted.Max))
                                
    Else
    
        Me.txtStepNum.text = "Completed!"
                             
    End If

End Sub

