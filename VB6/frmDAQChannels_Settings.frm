VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDAQChannels_Settings 
   Caption         =   "DAQ Channel Settings"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   12285
   Begin VB.Frame frameOptions 
      BorderStyle     =   0  'None
      Caption         =   "SQUID Magnetometer Calibration factors"
      Height          =   9015
      Index           =   8
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin VB.Frame Frame5 
         Caption         =   "IRM Capacitor Voltage In"
         Height          =   1212
         Left            =   120
         TabIndex        =   62
         Top             =   3960
         Width           =   3495
         Begin VB.ComboBox Combo10 
            Height          =   315
            Left            =   1200
            TabIndex        =   64
            Text            =   "Combo1"
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox Combo9 
            Height          =   315
            Left            =   1200
            TabIndex        =   63
            Text            =   "Combo1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label13 
            Caption         =   "DAQ Board:"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Channel:"
            Height          =   255
            Left            =   360
            TabIndex        =   65
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "IRM Capacitor Voltage In"
         Height          =   1212
         Left            =   120
         TabIndex        =   57
         Top             =   2400
         Width           =   3495
         Begin VB.ComboBox Combo8 
            Height          =   315
            Left            =   1200
            TabIndex        =   59
            Text            =   "Combo1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   1200
            TabIndex        =   58
            Text            =   "Combo1"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "Channel:"
            Height          =   255
            Left            =   360
            TabIndex        =   61
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "DAQ Board:"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "IRM Capacitor Voltage In"
         Height          =   1212
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Top             =   840
         Width           =   3495
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   1200
            TabIndex        =   54
            Text            =   "Combo1"
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   1
            Left            =   1200
            TabIndex        =   53
            Text            =   "Combo1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "DAQ Board:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Channel:"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   55
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "IRM Capacitor Voltage In"
         Height          =   1212
         Left            =   3720
         TabIndex        =   47
         Top             =   4440
         Width           =   3495
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   1200
            TabIndex        =   49
            Text            =   "Combo1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   1200
            TabIndex        =   48
            Text            =   "Combo1"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Channel:"
            Height          =   255
            Left            =   360
            TabIndex        =   51
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "DAQ Board:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "IRM Capacitor Voltage In"
         Height          =   1212
         Left            =   3720
         TabIndex        =   42
         Top             =   3120
         Width           =   3495
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   1200
            TabIndex        =   44
            Text            =   "Combo1"
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1200
            TabIndex        =   43
            Text            =   "Combo1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "DAQ Board:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Channel:"
            Height          =   255
            Left            =   360
            TabIndex        =   45
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame frameIRMCapacitorVoltageIn 
         Caption         =   "IRM Capacitor Voltage In"
         Height          =   1212
         Left            =   3720
         TabIndex        =   36
         Top             =   480
         Width           =   3495
         Begin VB.ComboBox cmbIRMCapacitorVoltageInBoard 
            Height          =   315
            Left            =   1200
            TabIndex        =   38
            Text            =   "Combo1"
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox cmbIRMCapacitorVoltageInChan 
            Height          =   315
            Left            =   1200
            TabIndex        =   37
            Text            =   "Combo1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label78 
            Caption         =   "DAQ Board:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label77 
            Caption         =   "Channel:"
            Height          =   255
            Left            =   360
            TabIndex        =   39
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame frameARMSet 
         Caption         =   "ARM Set"
         Height          =   1212
         Left            =   8400
         TabIndex        =   31
         Top             =   7080
         Width           =   3972
         Begin VB.ComboBox cmbARMSetChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   33
            Text            =   "Combo1"
            Top             =   720
            Width           =   2052
         End
         Begin VB.ComboBox cmbARMSetBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   32
            Text            =   "Combo1"
            Top             =   360
            Width           =   2052
         End
         Begin VB.Label Label79 
            Caption         =   "Channel:"
            Height          =   252
            Left            =   480
            TabIndex        =   35
            Top             =   720
            Width           =   732
         End
         Begin VB.Label Label80 
            Caption         =   "DAQ Board:"
            Height          =   252
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.Frame frameIRMVoltageOut 
         Caption         =   "IRM Voltage Out"
         Height          =   1212
         Left            =   8280
         TabIndex        =   26
         Top             =   2760
         Width           =   3972
         Begin VB.ComboBox cmbIRMVoltageOutBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   28
            Text            =   "Combo1"
            Top             =   360
            Width           =   2052
         End
         Begin VB.ComboBox cmbIRMVoltageOutChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   27
            Text            =   "Combo1"
            Top             =   720
            Width           =   2052
         End
         Begin VB.Label Label81 
            Caption         =   "DAQ Board:"
            Height          =   252
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label Label82 
            Caption         =   "Channel:"
            Height          =   252
            Left            =   480
            TabIndex        =   29
            Top             =   720
            Width           =   732
         End
      End
      Begin VB.Frame frameIRMFire 
         Caption         =   "IRM Fire"
         Height          =   1212
         Left            =   8160
         TabIndex        =   21
         Top             =   0
         Width           =   3972
         Begin VB.ComboBox cmbIRMFireBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   23
            Text            =   "Combo1"
            Top             =   360
            Width           =   2052
         End
         Begin VB.ComboBox cmbIRMFireChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   22
            Text            =   "Combo1"
            Top             =   720
            Width           =   2052
         End
         Begin VB.Label Label9 
            Caption         =   "DAQ Board:"
            Height          =   252
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label Label61 
            Caption         =   "Channel:"
            Height          =   252
            Left            =   480
            TabIndex        =   24
            Top             =   720
            Width           =   732
         End
      End
      Begin VB.Frame frameIRMTrim 
         Caption         =   "IRM Trim"
         Height          =   1212
         Left            =   8520
         TabIndex        =   16
         Top             =   5640
         Width           =   3972
         Begin VB.ComboBox cmbIRMTrimBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   18
            Text            =   "Combo1"
            Top             =   360
            Width           =   2052
         End
         Begin VB.ComboBox cmbIRMTrimChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   17
            Text            =   "Combo1"
            Top             =   720
            Width           =   2052
         End
         Begin VB.Label Label62 
            Caption         =   "DAQ Board:"
            Height          =   252
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label Label83 
            Caption         =   "Channel:"
            Height          =   252
            Left            =   480
            TabIndex        =   19
            Top             =   720
            Width           =   732
         End
      End
      Begin VB.Frame frameIRMReady 
         Caption         =   "IRM Ready"
         Height          =   1212
         Left            =   8160
         TabIndex        =   11
         Top             =   4200
         Width           =   3972
         Begin VB.ComboBox cmbIRMReadyChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   13
            Text            =   "Combo1"
            Top             =   720
            Width           =   2052
         End
         Begin VB.ComboBox cmbIRMReadyBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   12
            Text            =   "Combo1"
            Top             =   360
            Width           =   2052
         End
         Begin VB.Label Label63 
            Caption         =   "Channel:"
            Height          =   252
            Left            =   480
            TabIndex        =   15
            Top             =   720
            Width           =   732
         End
         Begin VB.Label Label64 
            Caption         =   "DAQ Board:"
            Height          =   252
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.Frame frameARMVoltageOut 
         Caption         =   "ARM Voltage Out"
         Height          =   1212
         Left            =   8280
         TabIndex        =   6
         Top             =   1320
         Width           =   3972
         Begin VB.ComboBox cmbARMVoltageOutChan 
            Height          =   315
            Left            =   1560
            TabIndex        =   8
            Text            =   "Combo1"
            Top             =   720
            Width           =   2052
         End
         Begin VB.ComboBox cmbARMVoltageOutBoard 
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Text            =   "Combo1"
            Top             =   360
            Width           =   2052
         End
         Begin VB.Label Label8 
            Caption         =   "Channel:"
            Height          =   252
            Left            =   480
            TabIndex        =   10
            Top             =   720
            Width           =   732
         End
         Begin VB.Label Label7 
            Caption         =   "DAQ Board:"
            Height          =   252
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.Frame frameIRMMonitor 
         Caption         =   "IRM Monitor"
         Height          =   1212
         Left            =   3720
         TabIndex        =   1
         Top             =   1800
         Width           =   3495
         Begin VB.ComboBox cmbIRMMonitorChan 
            Height          =   315
            Left            =   1200
            TabIndex        =   3
            Text            =   "Combo1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.ComboBox cmbIRMMonitorBoard 
            Height          =   315
            Left            =   1200
            TabIndex        =   2
            Text            =   "Combo1"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label96 
            Caption         =   "Channel:"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label97 
            Caption         =   "DAQ Board:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1095
         End
      End
      Begin ComctlLib.TabStrip tbsARMIRMChannels 
         Height          =   5895
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   10398
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   4
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Analog Out"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Analog In"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Digital In"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Digital Out"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDAQChannels_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
