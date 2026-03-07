VERSION 5.00
Begin VB.Form frmPlotAF 
   Caption         =   "AF Plot"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   9465
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7200
      Width           =   1215
   End
End
Attribute VB_Name = "frmPlotAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHide_Click()

    Me.Hide

End Sub
