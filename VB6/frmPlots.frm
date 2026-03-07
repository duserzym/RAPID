VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPlots 
   Caption         =   "Plots"
   ClientHeight    =   8685
   ClientLeft      =   1305
   ClientTop       =   2145
   ClientWidth     =   15420
   Icon            =   "frmPlots.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   15420
   Begin VB.CommandButton cmdOpenSAMdirectory 
      Caption         =   "Open SAM directory"
      Height          =   495
      Left            =   12360
      TabIndex        =   18
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpenSampleFile 
      Caption         =   "Open Sample File"
      Height          =   495
      Left            =   12360
      TabIndex        =   17
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpenSAMFile 
      Caption         =   "Open SAM file"
      Height          =   495
      Left            =   12360
      TabIndex        =   16
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CheckBox ChkCSD 
      Caption         =   "CSD"
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   120
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox ChkLabels 
      Caption         =   "Labels"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   120
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.PictureBox MomentX 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   3660
      Left            =   5400
      ScaleHeight     =   1.14
      ScaleLeft       =   -0.16
      ScaleMode       =   0  'User
      ScaleTop        =   -0.07
      ScaleWidth      =   1.2
      TabIndex        =   1
      Top             =   5000
      Width           =   5000
   End
   Begin VB.CheckBox ChkX 
      Caption         =   "Susceptibility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10440
      TabIndex        =   2
      Top             =   6120
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox ChkM 
      Caption         =   "Moment magnitude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10440
      TabIndex        =   3
      Top             =   5880
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.OptionButton optBedding 
      Caption         =   "Bedding coordinates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   5
      Top             =   6000
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.OptionButton optGeographic 
      Caption         =   "Geographic coordinates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   6
      Top             =   5760
      Width           =   2415
   End
   Begin VB.OptionButton optCore 
      Caption         =   "Core coordinates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   7
      Top             =   5520
      Width           =   1935
   End
   Begin VB.PictureBox EqualArea 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   5000
      Left            =   5400
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   1
      TabIndex        =   9
      Top             =   0
      Width           =   5000
   End
   Begin VB.TextBox txtZijLines 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11205
      TabIndex        =   10
      Text            =   "32"
      Top             =   5565
      Width           =   615
   End
   Begin VB.PictureBox Zijderveld 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   5000
      Left            =   10400
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   1
      TabIndex        =   11
      Top             =   0
      Width           =   5000
   End
   Begin VB.ComboBox cmbSamples 
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Text            =   "cmbSamples"
      Top             =   80
      Width           =   1815
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRMGfile 
      Height          =   8055
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   14208
      _Version        =   393216
      Rows            =   15
      Cols            =   8
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   10440
      TabIndex        =   21
      Top             =   5040
      Width           =   4935
      Begin VB.OptionButton optNS 
         Caption         =   "N - S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2080
         TabIndex        =   19
         Top             =   0
         Value           =   -1  'True
         Width           =   760
      End
      Begin VB.OptionButton optEW 
         Caption         =   "E - W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2080
         TabIndex        =   20
         Top             =   230
         Width           =   805
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Zijderveld [1967] plot (             orthographic projection)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   40
         TabIndex        =   12
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.Label Label27 
      Caption         =   "Sample:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label28 
      Caption         =   "previous steps"
      Height          =   255
      Left            =   11850
      TabIndex        =   8
      Top             =   5610
      Width           =   1095
   End
End
Attribute VB_Name = "frmPlots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' arc cosine
' error if NUMBER is outside the range [-1,1]
Function acos(ByVal number As Double) As Double
    If Abs(number) <> 1 Then
        acos = 1.5707963267949 - Atn(number / Sqr(1 - number * number))
    ElseIf number = -1 Then
        acos = 3.14159265358979
    End If
    'elseif number=1 --> Acos=0 (implicit)
End Function

' arc cotangent
' error if NUMBER is zero
Function ACot(Value As Double) As Double
    ACot = Atn(1 / Value)
End Function

' arc cosecant
' error if value is inside the range [-1,1]
Function ACsc(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ACsc = ASin(1 / value)
    If Abs(Value) <> 1 Then
        ACsc = Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ACsc = 1.5707963267949 * Sgn(Value)
    End If
End Function

Public Sub Actualize()
    If Me.WindowState = vbNormal Then
        Me.Width = 15520
        Me.Height = 9180
    End If
    If optCore.Value = True Then
        InitEqualArea
        EqualArea.CurrentX = 0
        EqualArea.CurrentY = 0.92
        EqualArea.FontBold = True
        EqualArea.Print "Core" & vbCrLf & "coordinates"
        EqualArea.FontBold = False
        ImportZijRoutine cmbSamples.text
    ElseIf optGeographic.Value = True Then
        InitEqualArea
        EqualArea.CurrentX = 0
        EqualArea.CurrentY = 0.92
        EqualArea.FontBold = True
        EqualArea.Print "Geographic" & vbCrLf & "coordinates"
        EqualArea.FontBold = False
        ImportZijRoutine cmbSamples.text
    ElseIf optBedding.Value = True Then
        InitEqualArea
        EqualArea.CurrentX = 0
        EqualArea.CurrentY = 0.92
        EqualArea.FontBold = True
        EqualArea.Print "Bedding" & vbCrLf & "coordinates"
        EqualArea.FontBold = False
        ImportZijRoutine cmbSamples.text
    End If
    EqualArea.Circle (0.8, 0.03), 0.01, RGB(255, 0, 0)
    EqualArea.Circle (0.89, 0.03), 0.01, RGB(0, 0, 255)
End Sub

' arc secant
' error if value is inside the range [-1,1]
Function ASec(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ASec = ACos(1 / value)
    If Abs(Value) <> 1 Then
        ASec = 1.5707963267949 - Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ASec = 3.14159265358979 * Sgn(Value)
    End If
End Function

' arc sine
' error if value is outside the range [-1,1]
Function ASin(Value As Double) As Double
    If Abs(Value) <> 1 Then
        ASin = Atn(Value / Sqr(1 - Value * Value))
    Else
        ASin = 1.5707963267949 * Sgn(Value)
    End If
End Function

Private Sub AveragePlotEqualArea(ByVal dec As Double, ByVal inc As Double, ByVal CSD As Double)
    ' (June 2008 L Carporzen) Plot the CSD ellipsoid
    Dim L0 As Double
    Dim L As Double
    Dim ax As Double
    Dim bx As Double
    Dim ay As Double
    Dim by As Double
    Dim X1 As Double
    Dim X2 As Double
    Dim Y1 As Double
    Dim Y2 As Double
    Dim i As Integer
    If CSD > 180 Then CSD = 0
    L0 = 1 / Sqr(Cos(inc * Pi / 180) * Cos(dec * Pi / 180) * Cos(inc * Pi / 180) * Cos(dec * Pi / 180) + Cos(inc * Pi / 180) * Sin(dec * Pi / 180) * Cos(inc * Pi / 180) * Sin(dec * Pi / 180))
    If inc >= 0 Then ' Down direction
        L = L0 * Sqr(1 - Sin(inc * Pi / 180))
        EqualArea.Circle ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5), 0.005, RGB(0, 0, 255)
        If CSD > 5 Then ' No calcul for small CSD
        If inc + CSD >= 90 Then
        ' The center of the equal area is include in the a95 which will be draw as a circle with the CSD as radius
        ax = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5
        ay = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + Sqr(1 - Sin((90 - CSD) * Pi / 180)) / 2
        bx = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + Sqr(1 - Sin((90 - CSD) * Pi / 180)) / 2
        by = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5
        Else
        ' Calcul of the coordinates of the axis of the ellipsoid
        ax = (Sin((dec + ASin(Sin((CSD) * Pi / 180) / Cos(inc * Pi / 180)) * 180 / Pi) * Pi / 180) * Sqr(1 - Sin((ASin(1 - (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180))) * (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180)))))) * 180 / Pi) * Pi / 180))) / 2 + 0.5
        ay = Abs(-(Cos((dec + ASin(Sin((CSD) * Pi / 180) / Cos(inc * Pi / 180)) * 180 / Pi) * Pi / 180) * Sqr(1 - Sin((ASin(1 - (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180))) * (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180)))))) * 180 / Pi) * Pi / 180))) / 2 + 0.5)
        bx = (Sin(dec * Pi / 180) * Sqr(1 - Sin((inc + CSD) * Pi / 180))) / 2 + 0.5
        by = Abs(-(Cos(dec * Pi / 180) * Sqr(1 - Sin((inc + CSD) * Pi / 180))) / 2 + 0.5)
        If ay > 1 Then ay = 1 - (ay - 1)
        If by > 1 Then by = 1 - (by - 1)
        End If
        ' Plot of the ellipsoid/circle by small segments (5 degrees)
        For i = 0 To 30
            ' The up ellipsoid/circle is a dash line
            If i = 1 Or i = 3 Or i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Or i = 17 Or i = 19 Or i = 21 Or i = 23 Or i = 25 Or i = 27 Or i = 29 Then
            Y1 = (Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by))) * Sin(3 * i * Pi / 180)
            X1 = Sqr((Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay))) ^ 2 * (1 - Sin(3 * i * Pi / 180) * Sin(3 * i * Pi / 180)))
            Y2 = (Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by))) * Sin(3 * (i + 1) * Pi / 180)
            X2 = Sqr((Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay))) ^ 2 * (1 - Sin(3 * (i + 1) * Pi / 180) * Sin(3 * (i + 1) * Pi / 180)))
            ' Test to don't plot the parts of the ellipsoid/circle which are outside of the plane inc = 0
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Cos((-dec) * Pi / 180) + Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Sin((-dec) * Pi / 180) + Y2 * Cos((-dec) * Pi / 180)), RGB(0, 0, 255)
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Cos((-dec) * Pi / 180) - Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Sin((-dec) * Pi / 180) - Y2 * Cos((-dec) * Pi / 180)), RGB(0, 0, 255)
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Cos((-dec) * Pi / 180) + Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Sin((-dec) * Pi / 180) + Y2 * Cos((-dec) * Pi / 180)), RGB(0, 0, 255)
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Cos((-dec) * Pi / 180) - Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Sin((-dec) * Pi / 180) - Y2 * Cos((-dec) * Pi / 180)), RGB(0, 0, 255)
            End If
        Next i
        End If
    Else ' Up direction
        L = L0 * Sqr(1 + Sin(inc * Pi / 180))
        EqualArea.Circle ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5), 0.005, RGB(255, 0, 0)
        If CSD > 5 Then ' No calcul for small CSD
        If Abs(inc) + CSD >= 90 Then
        ' The center of the equal area is include in the a95 which will be draw as a circle with the CSD as radius
        ax = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5
        ay = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + Sqr(1 - Sin((90 - CSD) * Pi / 180)) / 2
        bx = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + Sqr(1 - Sin((90 - CSD) * Pi / 180)) / 2
        by = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5
        Else
        ' Calcul of the coordinates of the axis of the ellipsoid
        ax = (Sin((dec + ASin(Sin((CSD) * Pi / 180) / Cos(Abs(inc) * Pi / 180)) * 180 / Pi) * Pi / 180) * Sqr(1 - Sin((ASin(1 - (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180))) * (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180)))))) * 180 / Pi) * Pi / 180))) / 2 + 0.5
        ay = Abs(-(Cos((dec + ASin(Sin((CSD) * Pi / 180) / Cos(Abs(inc) * Pi / 180)) * 180 / Pi) * Pi / 180) * Sqr(1 - Sin((ASin(1 - (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180))) * (2 * ((-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5) - 0.5) / (-(Cos(dec * Pi / 180)))))) * 180 / Pi) * Pi / 180))) / 2 + 0.5)
        bx = (Sin(dec * Pi / 180) * Sqr(1 - Sin((Abs(inc) + CSD) * Pi / 180))) / 2 + 0.5
        by = Abs(-(Cos(dec * Pi / 180) * Sqr(1 - Sin((Abs(inc) + CSD) * Pi / 180))) / 2 + 0.5)
        If ay > 1 Then ay = 1 - (ay - 1)
        If by > 1 Then by = 1 - (by - 1)
        End If
        ' Plot of the ellipsoid/circle by small segments (5 degrees)
        For i = 0 To 30
            ' The up ellipsoid/circle is a dash line
            If i = 1 Or i = 3 Or i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Or i = 17 Or i = 19 Or i = 21 Or i = 23 Or i = 25 Or i = 27 Or i = 29 Then
            Y1 = (Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by))) * Sin(3 * i * Pi / 180)
            X1 = Sqr((Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay))) ^ 2 * (1 - Sin(3 * i * Pi / 180) * Sin(3 * i * Pi / 180)))
            Y2 = (Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - bx) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - by))) * Sin(3 * (i + 1) * Pi / 180)
            X2 = Sqr((Sqr(((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) * ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - ax) + (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay) * (-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - ay))) ^ 2 * (1 - Sin(3 * (i + 1) * Pi / 180) * Sin(3 * (i + 1) * Pi / 180)))
            ' Test to don't plot the parts of the ellipsoid/circle which are outside of the plane inc = 0
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Cos((-dec) * Pi / 180) + Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Sin((-dec) * Pi / 180) + Y2 * Cos((-dec) * Pi / 180)), RGB(255, 0, 0)
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Cos((-dec) * Pi / 180) - Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Sin((-dec) * Pi / 180) - Y2 * Cos((-dec) * Pi / 180)), RGB(255, 0, 0)
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) + Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) + Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Cos((-dec) * Pi / 180) + Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Sin((-dec) * Pi / 180) + Y2 * Cos((-dec) * Pi / 180)), RGB(255, 0, 0)
            If (Abs(-(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180) - 0.5)) ^ 2 < Abs(0.5 ^ 2 - ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180) - 0.5) ^ 2) Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X1 * Cos((-dec) * Pi / 180) - Y1 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X1 * Sin((-dec) * Pi / 180) - Y1 * Cos((-dec) * Pi / 180))-((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5 - X2 * Cos((-dec) * Pi / 180) - Y2 * Sin((-dec) * Pi / 180), -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5 + X2 * Sin((-dec) * Pi / 180) - Y2 * Cos((-dec) * Pi / 180)), RGB(255, 0, 0)
            End If
        Next i
        End If
    End If
End Sub

Private Sub ChkCSD_Click()
    Actualize
End Sub

Private Sub ChkLabels_Click()
    Actualize
End Sub

Private Sub chkM_Click()
    Actualize
End Sub

Private Sub chkX_Click()
    Actualize
End Sub

Private Sub cmbSamples_Click()
    If LenB(cmbSamples.text) = 0 Then Exit Sub
    Actualize
End Sub

Private Sub cmdOpenSAMdirectory_Click()
    Dim r As Integer
    If LenB(cmbSamples.text) = 0 Then Exit Sub
    For r = 0 To frmSampleIndexRegistry.cmbSampCode.ListCount - 1
    If SampleIndexRegistry.IsValidSample(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1), cmbSamples.text) Then
        If frmSampleIndexRegistry.txtDir & frmSampleIndexRegistry.cmbSampCode.List(r) = SampleIndexRegistry(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1)).filedir & SampleIndexRegistry(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1)).SampleCode Then
            Open_SAMdirectory SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1), SampleIndexRegistry(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1)).filedir & SampleIndexRegistry(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1)).SampleCode
        End If
    Else
        Exit Sub
    End If
    Next r
    'Open_SAMdirectory frmSampleIndexRegistry.txtDir & frmSampleIndexRegistry.cmbSampCode & ".sam", frmSampleIndexRegistry.txtDir & frmSampleIndexRegistry.cmbSampCode
End Sub

Private Sub cmdOpenSAMFile_Click()
    Dim r As Integer
    If LenB(cmbSamples.text) = 0 Then Exit Sub
    For r = 0 To frmSampleIndexRegistry.cmbSampCode.ListCount - 1
    If SampleIndexRegistry.IsValidSample(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1), cmbSamples.text) Then
        If frmSampleIndexRegistry.txtDir & frmSampleIndexRegistry.cmbSampCode.List(r) & "\" & frmSampleIndexRegistry.cmbSampCode.List(r) & ".sam" = SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1) Then
        Open_SAMFile SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1)
        End If
    Else
        Exit Sub
    End If
    Next r
    'Open_SAMFile frmSampleIndexRegistry.txtDir & frmSampleIndexRegistry.cmbSampCode & "\" & frmSampleIndexRegistry.cmbSampCode & ".sam"
End Sub

Private Sub cmdOpenSampleFile_Click()
    Dim r As Integer
    If LenB(cmbSamples.text) = 0 Then Exit Sub
    For r = 0 To frmSampleIndexRegistry.cmbSampCode.ListCount - 1
    If SampleIndexRegistry.IsValidSample(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1), cmbSamples.text) Then
        If frmSampleIndexRegistry.txtDir & frmSampleIndexRegistry.cmbSampCode.List(r) = SampleIndexRegistry(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1)).filedir & SampleIndexRegistry(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1)).SampleCode Then
            Open_SampleFile SampleIndexRegistry(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1)).filedir & SampleIndexRegistry(SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1)).SampleCode & "\" & cmbSamples.text
        End If
    Else
        Exit Sub
    End If
    Next r
    'Open_SampleFile frmSampleIndexRegistry.txtDir & frmSampleIndexRegistry.cmbSampCode & "\" & cmbSamples.Text
End Sub

Private Sub Form_Load()

    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me

End Sub

Private Sub ImportZijRoutine(ByVal FilePath As String)
    ' (June 2008 L Carporzen) Visualisation of the data in a Zijderveld diagram near the equal area plot
    Dim filenum As Integer
    Dim whole_file As String ' To read the sample file for the Zijderveld diagram
    Dim lines As Variant
    Dim num_rows As Long
    Dim ZijLines As Long
    Dim r As Long
    Dim dec As Double
    Dim inc As Double
    Dim readMoment As Double
    Dim readcrdec As Double
    Dim readcrinc As Double
    Dim readcrdec2 As Double
    Dim readcrinc2 As Double
    Dim readcsd As Double
    Dim MaxZijX As Double
    Dim MaxZijY As Double
    Dim MinZijX As Double
    Dim MinZijY As Double
    Dim ZijScale As Double
    Dim ZijHoriOrig As Double
    Dim ZijVertOrig As Double
    Dim ZijX As Variant
    Dim ZijY As Variant
    Dim ZijZ As Variant
    Dim RMGlines As Variant ' To read the RMG file for the susceptibility versus demagnetization
    Dim RMGarray As Variant
    Dim numRMGrows As Long
    Dim SusceLines As Long
    Dim p As Long
    Dim Q As Long
    Dim MaxMoment As Double
    Dim MaxDemag As Double
    Dim MaxSusceptibility As Double
    Dim MinSusceptibility As Double
    Dim SusceScale As Double
    Dim SusceOrig As Double
    Dim DemagStep As Variant
    Dim Susceptibility As Variant
    Dim AF As Boolean
    Dim Thermal As Boolean
    Dim L0 As Double
    Dim L As Double
    Dim NewRMG As Variant
    Dim specParent As String
    Dim specimen As Sample
    If LenB(cmbSamples.text) = 0 Then Exit Sub
    specParent = SampleIndexRegistry.SampleFileByIndex(cmbSamples.ListIndex + 1)
    If SampleIndexRegistry.IsValidSample(specParent, FilePath) Then
        Set specimen = SampleIndexRegistry(specParent).sampleSet(FilePath)
    Else
        Exit Sub
    End If
    If Not LenB(dir$(specimen.SpecFilePath)) > 0 Then
        Exit Sub
    Else
        If FilePath = "" Then Exit Sub
    End If
    If Not LenB(dir$(specimen.SpecFilePath & ".rmg")) > 0 Then
        ChkX.Visible = False
    Else
        ChkX.Visible = True ' Allow reading the RMG file for susceptibility versus demagnetization
    End If
    If txtZijLines = "" Then txtZijLines = 0
    If txtZijLines < 0 Then txtZijLines = 0
    ZijLines = txtZijLines ' Nb of previous steps plot for the comparison
    Zijderveld.Cls ' Clean the plot
    ChkM.Visible = True
    MomentX.Cls ' Clean the plot
    filenum = FreeFile ' Read the sample file
    Open specimen.SpecFilePath For Input As #filenum
    whole_file = Input$(LOF(filenum), #filenum)
    Close #filenum
    lines = Split(whole_file, vbCrLf) ' Cut the file in lines
    whole_file = ""
    num_rows = UBound(lines)
    If num_rows < ZijLines + 2 Then ZijLines = num_rows - 2
    If ZijLines < 1 Then Exit Sub
    ReDim ZijX(ZijLines)
    ReDim ZijY(ZijLines)
    ReDim ZijZ(ZijLines)
    MaxZijX = 0.0000000001
    MaxZijY = 0.0000000001
    MinZijX = -0.0000000001
    MinZijY = -0.0000000001
    p = 1
    MaxMoment = 0.0000000001
    MaxDemag = 0
    MaxSusceptibility = 0.00001
    MinSusceptibility = 0
    If ChkX.Visible = True Then ' (June 2008 L Carporzen) Visualisation of the data as Susceptibility versus demagnetization below the equal area plot
        filenum = FreeFile ' Read the RMG file
        Open specimen.SpecFilePath & ".rmg" For Input As #filenum
        whole_file = Input$(LOF(filenum), #filenum)
        Close #filenum
        RMGlines = Split(whole_file, vbCrLf)
        whole_file = ""
        numRMGrows = UBound(RMGlines)
        ReDim RMGarray(3 * ZijLines)
        If numRMGrows > 3 * ZijLines Then
            SusceLines = 3 * ZijLines
        Else
            SusceLines = numRMGrows
        End If
        For r = 1 To SusceLines
            If RMGlines(numRMGrows - r) = "" Then Exit Sub
            RMGarray(r) = Split(RMGlines(numRMGrows - r), ",")
        Next r
        ReDim Susceptibility(ZijLines)
    End If
    ReDim DemagStep(ZijLines)
    grdRMGfile.ColWidth(0) = 660
    grdRMGfile.ColWidth(1) = 600
    grdRMGfile.ColWidth(3) = 900
    grdRMGfile.ColWidth(4) = 500
    grdRMGfile.ColWidth(5) = 500
    grdRMGfile.ColWidth(6) = 500
    grdRMGfile.ColWidth(7) = 500
    grdRMGfile.TextMatrix(0, 0) = "Step"
    grdRMGfile.TextMatrix(0, 1) = "Level"
    grdRMGfile.TextMatrix(0, 2) = "X (emu/Oe)"
    grdRMGfile.TextMatrix(0, 3) = "M (emu)"
    grdRMGfile.TextMatrix(0, 4) = "Dec"
    grdRMGfile.TextMatrix(0, 5) = "Inc"
    grdRMGfile.TextMatrix(0, 6) = "CSD"
    grdRMGfile.TextMatrix(0, 7) = "cm"
    grdRMGfile.Rows = ZijLines + 1
    For r = 1 To ZijLines
        readMoment = val(Mid$(lines(num_rows - r), 32, 8))
        readcsd = val(Mid$(lines(num_rows - r), 41, 5))
        grdRMGfile.TextMatrix(ZijLines - r + 1, 3) = Format$(val(Mid$(lines(num_rows - r), 32, 8)), "0.00E+")
        grdRMGfile.TextMatrix(ZijLines - r + 1, 6) = val(Mid$(lines(num_rows - r), 41, 5))
        If optCore.Value = True Then ' Core coordinates
            readcrdec = val(Mid$(lines(num_rows - r), 47, 5))
            readcrinc = val(Mid$(lines(num_rows - r), 53, 5))
            grdRMGfile.TextMatrix(ZijLines - r + 1, 4) = val(Mid$(lines(num_rows - r), 47, 5))
            grdRMGfile.TextMatrix(ZijLines - r + 1, 5) = val(Mid$(lines(num_rows - r), 53, 5))
            If r > 1 Then
            readcrdec2 = val(Mid$(lines(num_rows - r + 1), 47, 5))
            readcrinc2 = val(Mid$(lines(num_rows - r + 1), 53, 5))
            Else
            readcrdec2 = 0
            readcrinc2 = 0
            End If
        End If
        If optGeographic.Value = True Then ' Geographic coordinates
            readcrdec = val(Mid$(lines(num_rows - r), 8, 5))
            readcrinc = val(Mid$(lines(num_rows - r), 14, 5))
            grdRMGfile.TextMatrix(ZijLines - r + 1, 4) = val(Mid$(lines(num_rows - r), 8, 5))
            grdRMGfile.TextMatrix(ZijLines - r + 1, 5) = val(Mid$(lines(num_rows - r), 14, 5))
            If r > 1 Then
            readcrdec2 = val(Mid$(lines(num_rows - r + 1), 8, 5))
            readcrinc2 = val(Mid$(lines(num_rows - r + 1), 14, 5))
            Else
            readcrdec2 = 0
            readcrinc2 = 0
            End If
        End If
        If optBedding.Value = True Then ' Bedding coordinates
            readcrdec = val(Mid$(lines(num_rows - r), 20, 5))
            readcrinc = val(Mid$(lines(num_rows - r), 26, 5))
            grdRMGfile.TextMatrix(ZijLines - r + 1, 4) = val(Mid$(lines(num_rows - r), 20, 5))
            grdRMGfile.TextMatrix(ZijLines - r + 1, 5) = val(Mid$(lines(num_rows - r), 26, 5))
            If r > 1 Then
            readcrdec2 = val(Mid$(lines(num_rows - r + 1), 20, 5))
            readcrinc2 = val(Mid$(lines(num_rows - r + 1), 26, 5))
            Else
            readcrdec2 = 0
            readcrinc2 = 0
            End If
        End If
        PlotHistory readcrdec, readcrinc, readcrdec2, readcrinc2, readcsd
        ZijX(r) = readMoment * Cos(readcrinc * Pi / 180) * Cos(readcrdec * Pi / 180)
        ZijY(r) = readMoment * Cos(readcrinc * Pi / 180) * Sin(readcrdec * Pi / 180)
        ZijZ(r) = readMoment * Sin(readcrinc * Pi / 180)
      If optNS.Value = True Then
        If -ZijX(r) > MaxZijX Then MaxZijX = -ZijX(r)
        If ZijY(r) > MaxZijY Then MaxZijY = ZijY(r)
        If ZijZ(r) > MaxZijX Then MaxZijX = ZijZ(r)
        If -ZijX(r) < MinZijX Then MinZijX = -ZijX(r)
        If ZijY(r) < MinZijY Then MinZijY = ZijY(r)
        If ZijZ(r) < MinZijX Then MinZijX = ZijZ(r)
      Else
        If ZijX(r) > MaxZijX Then MaxZijX = ZijX(r)
        If ZijY(r) > MaxZijY Then MaxZijY = ZijY(r)
        If ZijZ(r) > MaxZijY Then MaxZijY = ZijZ(r)
        If ZijX(r) < MinZijX Then MinZijX = ZijX(r)
        If ZijY(r) < MinZijY Then MinZijY = ZijY(r)
        If ZijZ(r) < MinZijY Then MinZijY = ZijZ(r)
      End If
        DemagStep(r) = Abs(val(Mid$(lines(num_rows - r), 4, 3)))
        grdRMGfile.TextMatrix(ZijLines - r + 1, 0) = Left$(lines(num_rows - r), 3)
        grdRMGfile.TextMatrix(ZijLines - r + 1, 1) = Abs(val(Mid$(lines(num_rows - r), 4, 3)))
        If ChkX.Visible = True Then
            For Q = p To SusceLines
                If Left$(lines(num_rows - r), 2) = RMGarray(Q)(0) Then ' Are the 2 letters of the step labels are equals (AF or TT)
                    If RMGarray(Q)(0) = "TT" Then
                        Thermal = True
                    Else
                        Thermal = False
                    End If
                    If RMGarray(Q)(0) = "AF" Then
                        AF = True
                    Else
                        AF = False
                    End If
                    If DemagStep(r) = Abs(val(Right$(RMGarray(Q)(1), 3))) Then  ' Are the 3 last digits of the step numbers are equals
                        DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                        Susceptibility(r) = val(RMGarray(Q)(8))
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 0) = RMGarray(Q)(0)
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 1) = val(RMGarray(Q)(1))
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 2) = Format$(val(RMGarray(Q)(8)), "0.0000E+")
                        If IsInArray(Q, 18, RMGarray) Then grdRMGfile.TextMatrix(ZijLines - r + 1, 7) = Format$(RMGarray(Q)(18), "0.00")
                        p = Q + 1
                        Exit For
                    ElseIf DemagStep(r) = Abs(val(Right$(RMGarray(Q)(1), 4))) Then ' Are the 4 last digits of the step numbers are equals
                        DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                        Susceptibility(r) = val(RMGarray(Q)(8))
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 0) = RMGarray(Q)(0)
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 1) = val(RMGarray(Q)(1))
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 2) = Format$(val(RMGarray(Q)(8)), "0.0000E+")
                        If IsInArray(Q, 18, RMGarray) Then grdRMGfile.TextMatrix(ZijLines - r + 1, 7) = Format$(RMGarray(Q)(18), "0.00")
                        p = Q + 1
                        Exit For
                    End If
                ElseIf Left$(lines(num_rows - r), 3) = RMGarray(Q)(0) Then ' Are the 3 letters of the step labels are equals
                    If RMGarray(Q)(0) = "IRM" Or RMGarray(Q)(0) = "ARM" Or RMGarray(Q)(0) = "AFz" Then
                        AF = True
                    Else
                        AF = False
                    End If
                    If DemagStep(r) = Abs(val(Right$(RMGarray(Q)(1), 3))) Then ' Are the 3 last digits of the step numbers are equals
                        DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                        Susceptibility(r) = val(RMGarray(Q)(8))
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 0) = RMGarray(Q)(0)
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 1) = val(RMGarray(Q)(1))
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 2) = Format$(val(RMGarray(Q)(8)), "0.0000E+")
                        If IsInArray(Q, 18, RMGarray) Then grdRMGfile.TextMatrix(ZijLines - r + 1, 7) = Format$(RMGarray(Q)(18), "0.00")
                        p = Q + 1
                        Exit For
                    ElseIf DemagStep(r) = Abs(val(Right$(RMGarray(Q)(1), 4))) Then ' Are the 4 last digits of the step numbers are equals
                        DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                        Susceptibility(r) = val(RMGarray(Q)(8))
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 0) = RMGarray(Q)(0)
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 1) = val(RMGarray(Q)(1))
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 2) = Format$(val(RMGarray(Q)(8)), "0.0000E+")
                        If IsInArray(Q, 18, RMGarray) Then grdRMGfile.TextMatrix(ZijLines - r + 1, 7) = Format$(RMGarray(Q)(18), "0.00")
                        p = Q + 1
                        Exit For
                    ElseIf Left$(lines(num_rows - r), 3) = "ARM" Then ' Is their only a real number only in the RMG (ARM)
                        DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                        Susceptibility(r) = val(RMGarray(Q)(8))
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 0) = RMGarray(Q)(0)
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 1) = val(RMGarray(Q)(1))
                        grdRMGfile.TextMatrix(ZijLines - r + 1, 2) = Format$(val(RMGarray(Q)(8)), "0.0000E+")
                        If IsInArray(Q, 18, RMGarray) Then grdRMGfile.TextMatrix(ZijLines - r + 1, 7) = Format$(RMGarray(Q)(18), "0.00")
                        p = Q + 1
                        Exit For
                    End If
                ElseIf Left$(lines(num_rows - r), 5) = RMGarray(Q)(0) Then ' Are the step labels are equals to AFmax
                    If RMGarray(Q)(0) = "AFmax" Then
                        AF = True
                    Else
                        AF = False
                    End If
                    DemagStep(r) = Abs(val(RMGarray(Q)(1)))
                    Susceptibility(r) = val(RMGarray(Q)(8))
                    grdRMGfile.TextMatrix(ZijLines - r + 1, 0) = RMGarray(Q)(0)
                    grdRMGfile.TextMatrix(ZijLines - r + 1, 1) = val(RMGarray(Q)(1))
                    grdRMGfile.TextMatrix(ZijLines - r + 1, 2) = Format$(val(RMGarray(Q)(8)), "0.0000E+")
                    If IsInArray(Q, 18, RMGarray) Then grdRMGfile.TextMatrix(ZijLines - r + 1, 7) = Format$(RMGarray(Q)(18), "0.00")
                    p = Q + 1
                    Exit For
                End If
            Next Q
            If Susceptibility(r) = "" Then Susceptibility(r) = 0
          End If
          If readMoment > MaxMoment Then MaxMoment = readMoment
          If DemagStep(r) > MaxDemag Then MaxDemag = DemagStep(r)
          If ChkX.Visible = True And ChkX.Value = Checked Then
            If Susceptibility(r) > MaxSusceptibility Then MaxSusceptibility = Susceptibility(r)
            If Susceptibility(r) < MinSusceptibility Then MinSusceptibility = Susceptibility(r)
        End If
    Next r
    If ChkM.Value = Checked Or ChkX.Value = Checked Then
        MomentX.Line (0, 1)-(1, 1) ' horizontal axis
        MomentX.Line (1, 1 - 0.02)-(1, 1 + 0.02)
        MomentX.CurrentX = 0
        MomentX.CurrentY = 1.02
        MomentX.Print "0"
        MomentX.Line (0, 1 - 0.02)-(0, 1 + 0.02)
        MomentX.CurrentX = 1 - 0.06
        MomentX.CurrentY = 1.02
        MomentX.Print MaxDemag
        If Thermal = True Then
          MomentX.CurrentX = 0.5
          MomentX.CurrentY = 1.01
          MomentX.Print "°C"
        ElseIf AF = True Then
          MomentX.CurrentX = 0.5
          MomentX.CurrentY = 1.01
          MomentX.Print "Oe"
        Else
          MomentX.CurrentX = 0.4
          MomentX.CurrentY = 1.01
          MomentX.Print "Oe & °C"
        End If
    End If
    If ChkM.Value = Checked Then
        MomentX.Line (0, 0)-(0, 1), RGB(255, 0, 0) ' vertical axis
        MomentX.Line (-0.02, 1)-(0.02, 1), RGB(255, 0, 0)
        MomentX.Line (-0.02, 0)-(0.02, 0), RGB(255, 0, 0)
        MomentX.ForeColor = RGB(255, 0, 0)
        MomentX.CurrentX = -0.15
        MomentX.CurrentY = 0.5
        MomentX.Print "emu"
        MomentX.CurrentX = -0.15
        MomentX.CurrentY = -0.07
        MomentX.Print Format$(MaxMoment, "0.00E+")
        MomentX.CurrentX = -0.05
        MomentX.CurrentY = 1 - 0.05
        MomentX.Print "0"
        MomentX.ForeColor = RGB(0, 0, 0)
    End If
    If ChkX.Visible = True And ChkX.Value = Checked Then
        SusceScale = Abs(MaxSusceptibility - MinSusceptibility)
        SusceOrig = Abs(MinSusceptibility / SusceScale)
        If ChkM.Value = Checked Then
            MomentX.Line (0, 0)-(0, 1) ' vertical axis
            MomentX.Line (-0.02, 1)-(0.02, 1)
            MomentX.Line (-0.02, 0)-(0.02, 0)
        Else
            MomentX.Line (0, 0)-(0, 1), RGB(0, 0, 255) ' vertical axis
            MomentX.Line (-0.02, 1)-(0.02, 1), RGB(0, 0, 255)
            MomentX.Line (-0.02, 0)-(0.02, 0), RGB(0, 0, 255)
        End If
        MomentX.Line (0, 1 - SusceOrig)-(1, 1 - SusceOrig), RGB(0, 0, 255) ' horizontal axis
        MomentX.Line (1, 1 - SusceOrig - 0.02)-(1, 1 - SusceOrig + 0.02), RGB(0, 0, 255)
        MomentX.ForeColor = RGB(0, 0, 255)
        MomentX.CurrentX = -0.05
        MomentX.CurrentY = 1 - SusceOrig - 0.05
        MomentX.Print "0"
        MomentX.CurrentX = -0.15
        If 1 - SusceOrig > 0.5 Then
            MomentX.CurrentY = 1 - SusceOrig - 0.05 - 0.05
        Else
            MomentX.CurrentY = 1 - SusceOrig + 0.05 - 0.05
        End If
        MomentX.Print "emu/Oe"
        MomentX.CurrentX = -0.15
        MomentX.CurrentY = 0.01
        MomentX.Print Format$(MaxSusceptibility, "0.00E+")
        If Not MinSusceptibility = 0 Then
            MomentX.CurrentX = -0.15
            MomentX.CurrentY = 1.01
            MomentX.Print Format$(MinSusceptibility, "0.00E+")
        End If
        MomentX.ForeColor = RGB(0, 0, 0)
    End If
    If MaxZijX = MinZijX Or MaxZijY = MinZijY Then
        ' Do nothing, avoid bugs???
    Else ' We can plot
        ZijScale = Abs(MaxZijX - MinZijX)
        MaxZijX = MaxZijX + 0.05 * ZijScale
        MinZijX = MinZijX - 0.05 * ZijScale
        ZijScale = Abs(MaxZijY - MinZijY)
        MaxZijY = MaxZijY + 0.05 * ZijScale
        MinZijY = MinZijY - 0.05 * ZijScale
      If optNS.Value = True Then
        If Abs(MaxZijX - MinZijX) > Abs(MaxZijY - MinZijY) Then
            ZijScale = Abs(MaxZijX - MinZijX) ' Same scale for both axis
            ZijHoriOrig = Abs(MinZijY / ZijScale) + (1 - Abs(MaxZijY - MinZijY) / ZijScale) / 2 'Center the plot in the page
            ZijVertOrig = Abs(MinZijX / ZijScale) 'The lowest and highest values are on the borders of the plot
        Else
            ZijScale = Abs(MaxZijY - MinZijY) ' Same scale for both axis
            ZijHoriOrig = Abs(MinZijY / ZijScale) 'The lowest and highest values are on the borders of the plot
            ZijVertOrig = Abs(MinZijX / ZijScale) + (1 - Abs(MaxZijX - MinZijX) / ZijScale) / 2 'Center the plot in the page
        End If
        Zijderveld.Line (ZijHoriOrig, 0)-(ZijHoriOrig, 1) ' vertical axis
        Zijderveld.Line (0, ZijVertOrig)-(1, ZijVertOrig) ' horizontal axis
        Zijderveld.CurrentX = 0
        Zijderveld.CurrentY = ZijVertOrig
        Zijderveld.Print "W"
        Zijderveld.CurrentX = 1 - 0.025
        Zijderveld.CurrentY = ZijVertOrig
        Zijderveld.Print "E"
        Zijderveld.CurrentX = ZijHoriOrig - 0.03
        Zijderveld.CurrentY = 0
        Zijderveld.Print "N  Up"
        Zijderveld.CurrentX = ZijHoriOrig - 0.03
        Zijderveld.CurrentY = 1 - 0.04
        Zijderveld.Print "S  Down"
        Zijderveld.Circle (ZijHoriOrig - 0.06, 0.02), 0.02, RGB(0, 0, 255)
        Zijderveld.Circle (ZijHoriOrig + 0.08, 0.02), 0.02, RGB(255, 0, 0)
        If ZijHoriOrig >= 0.5 Then
            Zijderveld.CurrentX = 0.04
        Else
            Zijderveld.CurrentX = 0.7
        End If
        Zijderveld.CurrentY = 0
        Zijderveld.Print "View: " & Format$(ZijScale, "0.00E+") & " emu"   ' scale
        For r = 1 To ZijLines 'Circles for each step
            lines(num_rows - r) = ""
            If ChkM.Value = Checked And Not Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) = 0 And Not MaxDemag = 0 Then MomentX.Circle (DemagStep(r) / MaxDemag, 1 - Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) / MaxMoment), 0.005, RGB(255, 0, 0)
            If ChkX.Visible = True And ChkX.Value = Checked Then
                If Not Susceptibility(r) = 0 And Not MaxDemag = 0 Then MomentX.Circle (DemagStep(r) / MaxDemag, 1 - (Susceptibility(r) / SusceScale + SusceOrig)), 0.005, RGB(0, 0, 255)
            End If
            If ChkLabels.Value = Checked Then
                Zijderveld.CurrentX = ZijY(r) / ZijScale + ZijHoriOrig
                Zijderveld.CurrentY = -ZijX(r) / ZijScale + ZijVertOrig
                Zijderveld.Print DemagStep(r)
                Zijderveld.CurrentX = ZijY(r) / ZijScale + ZijHoriOrig
                Zijderveld.CurrentY = ZijZ(r) / ZijScale + ZijVertOrig
                Zijderveld.Print DemagStep(r)
                dec = val(grdRMGfile.TextMatrix(ZijLines - r + 1, 4))
                inc = val(grdRMGfile.TextMatrix(ZijLines - r + 1, 5))
                L0 = 1 / Sqr(Cos(inc * Pi / 180) * Cos(dec * Pi / 180) * Cos(inc * Pi / 180) * Cos(dec * Pi / 180) + Cos(inc * Pi / 180) * Sin(dec * Pi / 180) * Cos(inc * Pi / 180) * Sin(dec * Pi / 180))
                If inc >= 0 Then ' Down direction
                    L = L0 * Sqr(1 - Sin(inc * Pi / 180))
                    EqualArea.CurrentX = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5
                    EqualArea.CurrentY = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5
                Else ' Up direction
                    L = L0 * Sqr(1 + Sin(inc * Pi / 180))
                    EqualArea.CurrentX = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5
                    EqualArea.CurrentY = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5
                End If
                EqualArea.Print DemagStep(r)
            End If
            Zijderveld.Circle (ZijY(r) / ZijScale + ZijHoriOrig, -ZijX(r) / ZijScale + ZijVertOrig), 0.005, RGB(0, 0, 255)
            Zijderveld.Circle (ZijY(r) / ZijScale + ZijHoriOrig, ZijZ(r) / ZijScale + ZijVertOrig), 0.005, RGB(255, 0, 0)
        Next r
        For r = 1 To ZijLines - 1 'Link each step by a line
            If ChkM.Value = Checked And Not (Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) = 0 Or Sqr(ZijX(r + 1) ^ 2 + ZijY(r + 1) ^ 2 + ZijZ(r + 1) ^ 2) = 0) And Not MaxDemag = 0 Then MomentX.Line (DemagStep(r) / MaxDemag, 1 - Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) / MaxMoment)-(DemagStep(r + 1) / MaxDemag, 1 - Sqr(ZijX(r + 1) ^ 2 + ZijY(r + 1) ^ 2 + ZijZ(r + 1) ^ 2) / MaxMoment), RGB(255, 0, 0)
            If ChkX.Visible = True And ChkX.Value = Checked Then
                If Not (Susceptibility(r) = 0 Or Susceptibility(r + 1) = 0) And Not MaxDemag = 0 Then MomentX.Line (DemagStep(r) / MaxDemag, 1 - (Susceptibility(r) / SusceScale + SusceOrig))-(DemagStep(r + 1) / MaxDemag, 1 - (Susceptibility(r + 1) / SusceScale + SusceOrig)), RGB(0, 0, 255)
            End If
            Zijderveld.Line (ZijY(r) / ZijScale + ZijHoriOrig, -ZijX(r) / ZijScale + ZijVertOrig)-(ZijY(r + 1) / ZijScale + ZijHoriOrig, -ZijX(r + 1) / ZijScale + ZijVertOrig), RGB(0, 0, 255)
            Zijderveld.Line (ZijY(r) / ZijScale + ZijHoriOrig, ZijZ(r) / ZijScale + ZijVertOrig)-(ZijY(r + 1) / ZijScale + ZijHoriOrig, ZijZ(r + 1) / ZijScale + ZijVertOrig), RGB(255, 0, 0)
        Next r
      Else
        If Abs(MaxZijX - MinZijX) > Abs(MaxZijY - MinZijY) Then
            ZijScale = Abs(MaxZijX - MinZijX) ' Same scale for both axis
            ZijHoriOrig = Abs(MinZijX / ZijScale) 'Center the plot in the page
            ZijVertOrig = Abs(MinZijY / ZijScale) + (1 - Abs(MaxZijY - MinZijY) / ZijScale) / 2 'The lowest and highest values are on the borders of the plot
        Else
            ZijScale = Abs(MaxZijY - MinZijY) ' Same scale for both axis
            ZijHoriOrig = Abs(MinZijX / ZijScale) + (1 - Abs(MaxZijX - MinZijX) / ZijScale) / 2 'The lowest and highest values are on the borders of the plot
            ZijVertOrig = Abs(MinZijY / ZijScale) 'Center the plot in the page
        End If
        Zijderveld.Line (ZijHoriOrig, 0)-(ZijHoriOrig, 1) ' vertical axis
        Zijderveld.Line (0, ZijVertOrig)-(1, ZijVertOrig) ' horizontal axis
        Zijderveld.CurrentX = 0
        Zijderveld.CurrentY = ZijVertOrig
        Zijderveld.Print "S"
        Zijderveld.CurrentX = 1 - 0.025
        Zijderveld.CurrentY = ZijVertOrig
        Zijderveld.Print "N"
        Zijderveld.CurrentX = ZijHoriOrig - 0.037
        Zijderveld.CurrentY = 0
        Zijderveld.Print "W  Up"
        Zijderveld.CurrentX = ZijHoriOrig - 0.03
        Zijderveld.CurrentY = 1 - 0.04
        Zijderveld.Print "E  Down"
        Zijderveld.Circle (ZijHoriOrig - 0.06, 0.02), 0.02, RGB(0, 0, 255)
        Zijderveld.Circle (ZijHoriOrig + 0.08, 0.02), 0.02, RGB(255, 0, 0)
        If ZijHoriOrig >= 0.5 Then
            Zijderveld.CurrentX = 0.04
        Else
            Zijderveld.CurrentX = 0.7
        End If
        Zijderveld.CurrentY = 0
        Zijderveld.Print "View: " & Format$(ZijScale, "0.00E+") & " emu"   ' scale
        For r = 1 To ZijLines 'Circles for each step
            lines(num_rows - r) = ""
            If ChkM.Value = Checked And Not Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) = 0 And Not MaxDemag = 0 Then MomentX.Circle (DemagStep(r) / MaxDemag, 1 - Sqr(ZijX(r) ^ 2 + ZijY(r) ^ 2 + ZijZ(r) ^ 2) / MaxMoment), 0.005, RGB(255, 0, 0)
            If ChkX.Visible = True And ChkX.Value = Checked Then
                If Not Susceptibility(r) = 0 And Not MaxDemag = 0 Then MomentX.Circle (DemagStep(r) / MaxDemag, 1 - (Susceptibility(r) / SusceScale + SusceOrig)), 0.005, RGB(0, 0, 255)
            End If
            If ChkLabels.Value = Checked Then
                Zijderveld.CurrentX = ZijX(r) / ZijScale + ZijHoriOrig
                Zijderveld.CurrentY = ZijY(r) / ZijScale + ZijVertOrig
                Zijderveld.Print DemagStep(r)
                Zijderveld.CurrentX = ZijX(r) / ZijScale + ZijHoriOrig
                Zijderveld.CurrentY = ZijZ(r) / ZijScale + ZijVertOrig
                Zijderveld.Print DemagStep(r)
                dec = val(grdRMGfile.TextMatrix(ZijLines - r + 1, 4))
                inc = val(grdRMGfile.TextMatrix(ZijLines - r + 1, 5))
                L0 = 1 / Sqr(Cos(inc * Pi / 180) * Cos(dec * Pi / 180) * Cos(inc * Pi / 180) * Cos(dec * Pi / 180) + Cos(inc * Pi / 180) * Sin(dec * Pi / 180) * Cos(inc * Pi / 180) * Sin(dec * Pi / 180))
                If inc >= 0 Then ' Down direction
                    L = L0 * Sqr(1 - Sin(inc * Pi / 180))
                    EqualArea.CurrentX = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5
                    EqualArea.CurrentY = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5
                Else ' Up direction
                    L = L0 * Sqr(1 + Sin(inc * Pi / 180))
                    EqualArea.CurrentX = (Sin(dec * Pi / 180) * L / L0) / 2 + 0.5
                    EqualArea.CurrentY = -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5
                End If
                EqualArea.Print DemagStep(r)
            End If
            Zijderveld.Circle (ZijX(r) / ZijScale + ZijHoriOrig, ZijY(r) / ZijScale + ZijVertOrig), 0.005, RGB(0, 0, 255)
            Zijderveld.Circle (ZijX(r) / ZijScale + ZijHoriOrig, ZijZ(r) / ZijScale + ZijVertOrig), 0.005, RGB(255, 0, 0)
        Next r
        For r = 1 To ZijLines - 1 'Link each step by a line
            If ChkM.Value = Checked And Not (Sqr(ZijY(r) ^ 2 + ZijX(r) ^ 2 + ZijZ(r) ^ 2) = 0 Or Sqr(ZijY(r + 1) ^ 2 + ZijX(r + 1) ^ 2 + ZijZ(r + 1) ^ 2) = 0) And Not MaxDemag = 0 Then MomentX.Line (DemagStep(r) / MaxDemag, 1 - Sqr(ZijY(r) ^ 2 + ZijX(r) ^ 2 + ZijZ(r) ^ 2) / MaxMoment)-(DemagStep(r + 1) / MaxDemag, 1 - Sqr(ZijY(r + 1) ^ 2 + ZijX(r + 1) ^ 2 + ZijZ(r + 1) ^ 2) / MaxMoment), RGB(255, 0, 0)
            If ChkX.Visible = True And ChkX.Value = Checked Then
                If Not (Susceptibility(r) = 0 Or Susceptibility(r + 1) = 0) And Not MaxDemag = 0 Then MomentX.Line (DemagStep(r) / MaxDemag, 1 - (Susceptibility(r) / SusceScale + SusceOrig))-(DemagStep(r + 1) / MaxDemag, 1 - (Susceptibility(r + 1) / SusceScale + SusceOrig)), RGB(0, 0, 255)
            End If
            Zijderveld.Line (ZijX(r) / ZijScale + ZijHoriOrig, ZijY(r) / ZijScale + ZijVertOrig)-(ZijX(r + 1) / ZijScale + ZijHoriOrig, ZijY(r + 1) / ZijScale + ZijVertOrig), RGB(0, 0, 255)
            Zijderveld.Line (ZijX(r) / ZijScale + ZijHoriOrig, ZijZ(r) / ZijScale + ZijVertOrig)-(ZijX(r + 1) / ZijScale + ZijHoriOrig, ZijZ(r + 1) / ZijScale + ZijVertOrig), RGB(255, 0, 0)
        Next r
      End If
    End If
End Sub

Private Sub InitEqualArea()
    ' (June 2008 L Carporzen) Visualisation of the data in Equal area plot
    Dim i As Integer
    EqualArea.Cls ' Clean the plot
    EqualArea.CurrentX = 0
    EqualArea.CurrentY = 0
    EqualArea.FontBold = True
    EqualArea.Print "Equal area" & vbCrLf & "stereoplot"
    EqualArea.FontBold = False
    EqualArea.Circle (0.5, 0.5), 0.5 ' external circle
    EqualArea.Line (0.5, 0)-(0.5, 1) ' vertical axis
    EqualArea.Line (0, 0.5)-(1, 0.5) ' horizontal axis
    For i = 1 To 8
        EqualArea.Line (0.5 + Sqr(1 - Sin(10 * i * Pi / 180)) / 2, 0.49)-(0.5 + Sqr(1 - Sin(10 * i * Pi / 180)) / 2, 0.52) ' vertical ticks
        EqualArea.Line (0.5 - Sqr(1 - Sin(10 * i * Pi / 180)) / 2, 0.49)-(0.5 - Sqr(1 - Sin(10 * i * Pi / 180)) / 2, 0.52) ' vertical ticks
        EqualArea.Line (0.49, 0.5 + Sqr(1 - Sin(10 * i * Pi / 180)) / 2)-(0.52, 0.5 + Sqr(1 - Sin(10 * i * Pi / 180)) / 2) ' horizontal ticks
        EqualArea.Line (0.49, 0.5 - Sqr(1 - Sin(10 * i * Pi / 180)) / 2)-(0.52, 0.5 - Sqr(1 - Sin(10 * i * Pi / 180)) / 2) ' horizontal ticks
    Next i
    EqualArea.CurrentX = 0.5 + 0.01
    EqualArea.CurrentY = 0.01
    EqualArea.Print "N" ' Label
    EqualArea.CurrentX = 0.5 + 0.01
    EqualArea.CurrentY = 1 - 0.05
    EqualArea.Print "S" ' Label
    EqualArea.CurrentX = 1 - 0.03
    EqualArea.CurrentY = 0.5 + 0.01
    EqualArea.Print "E" ' Label
    EqualArea.CurrentX = 0.01
    EqualArea.CurrentY = 0.5 + 0.01
    EqualArea.Print "W" ' Label
    EqualArea.CurrentX = 0.82
    EqualArea.CurrentY = 0.01
    EqualArea.Print "Up"
    EqualArea.CurrentX = 0.91
    EqualArea.CurrentY = 0.01
    EqualArea.Print "Down"
End Sub

Private Function IsInArray(p As Long, Column As _
    Long, arrSearch As Variant) As Boolean
    On Error GoTo LocalError
    If Not IsArray(arrSearch) Then Exit Function
    IsInArray = arrSearch(p)(Column)
Exit Function
LocalError:
    'Justin (just in case)
End Function

Private Sub optBedding_Click()
    optCore.Value = False
    optGeographic.Value = False
    optBedding.Value = True
    Actualize
End Sub

Private Sub optCore_Click()
    optCore.Value = True
    optGeographic.Value = False
    optBedding.Value = False
    Actualize
End Sub

Private Sub optEW_Click()
    optEW.Value = True
    optNS.Value = False
    Actualize
End Sub

Private Sub optGeographic_Click()
    optCore.Value = False
    optGeographic.Value = True
    optBedding.Value = False
    Actualize
End Sub

Private Sub optNS_Click()
    optNS.Value = True
    optEW.Value = False
    Actualize
End Sub

Private Sub PlotHistory(ByVal dec As Double, ByVal inc As Double, ByVal dec2 As Double, ByVal inc2 As Double, ByVal CSD As Double)
    ' (June 2008 L Carporzen) Visualisation of the previous directions
    Dim L0 As Double
    Dim L As Double
    Dim L02 As Double
    Dim L2 As Double
    L0 = 1 / Sqr(Cos(inc * Pi / 180) * Cos(dec * Pi / 180) * Cos(inc * Pi / 180) * Cos(dec * Pi / 180) + Cos(inc * Pi / 180) * Sin(dec * Pi / 180) * Cos(inc * Pi / 180) * Sin(dec * Pi / 180))
    L02 = 1 / Sqr(Cos(inc2 * Pi / 180) * Cos(dec2 * Pi / 180) * Cos(inc2 * Pi / 180) * Cos(dec2 * Pi / 180) + Cos(inc2 * Pi / 180) * Sin(dec2 * Pi / 180) * Cos(inc2 * Pi / 180) * Sin(dec2 * Pi / 180))
    If inc >= 0 Then ' Down direction
        L = L0 * Sqr(1 - Sin(inc * Pi / 180))
        L2 = L02 * Sqr(1 - Sin(inc2 * Pi / 180))
        If ChkCSD.Value = Checked Then
            AveragePlotEqualArea dec, inc, CSD
        Else
            EqualArea.Circle ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5), 0.005, RGB(0, 0, 255)
        End If
        If Not inc2 = 0 And Not L02 = 0 And Not inc = inc2 Then
            If inc / inc2 > 0 Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5)-((Sin(dec2 * Pi / 180) * L2 / L02) / 2 + 0.5, -(Cos(dec2 * Pi / 180) * L2 / L02) / 2 + 0.5), RGB(0, 0, 255)
        End If
    Else ' Up direction
        L = L0 * Sqr(1 + Sin(inc * Pi / 180))
        L2 = L02 * Sqr(1 + Sin(inc2 * Pi / 180))
        If ChkCSD.Value = Checked Then
            AveragePlotEqualArea dec, inc, CSD
        Else
            EqualArea.Circle ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5), 0.005, RGB(255, 0, 0)
        End If
        If Not inc2 = 0 And Not L02 = 0 And Not inc = inc2 Then
            If inc / inc2 > 0 Then EqualArea.Line ((Sin(dec * Pi / 180) * L / L0) / 2 + 0.5, -(Cos(dec * Pi / 180) * L / L0) / 2 + 0.5)-((Sin(dec2 * Pi / 180) * L2 / L02) / 2 + 0.5, -(Cos(dec2 * Pi / 180) * L2 / L02) / 2 + 0.5), RGB(255, 0, 0)
        End If
    End If
End Sub

Public Sub RefreshSamples()
    ' Adds fields to the combobox
    ' cmbSamples, so the user can manually select samples to view.
    Dim i As Integer, j As Integer
    cmbSamples.Clear
On Error GoTo fin
    If SampleIndexRegistry.Count = 0 Then Exit Sub
    For i = 1 To SampleIndexRegistry.Count
        With SampleIndexRegistry(i).sampleSet
        If .Count > 0 Then
            For j = 1 To .Count
                cmbSamples.AddItem .Item(j).Samplename
            Next j
        End If
        End With
    Next i
    On Error GoTo 0
fin:
End Sub

' A random number in the range (low, high)
Function Rnd2(low As Single, high As Single) As Single
    Rnd2 = Rnd * (high - low) + low
End Function

