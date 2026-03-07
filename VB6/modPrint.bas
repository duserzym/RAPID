Attribute VB_Name = "modPrint"
' This module controls printing on the default printer.
' It is just a series of functions that determines when
' page breaks should occur.

Option Explicit

Private numlines     As Integer ' number of lines printed on page
Const pagelines = 80            ' number of lines on page

Public Sub Print_Line(Optional line As String = vbNullString)
    If numlines < pagelines Then
        numlines = numlines + 1
    Else
        Print_PageBreak
    End If
    Printer.Print line
End Sub

Public Function Print_LinesLeft() As Integer
    Print_LinesLeft = pagelines - numlines
End Function

Public Sub Print_PageBreak()
    numlines = 0
    Printer.NewPage
End Sub

