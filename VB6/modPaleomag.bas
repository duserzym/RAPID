Attribute VB_Name = "modPaleomag"
Option Explicit
Global Const Geographic As Integer = 0
Global Const Stratigraphic As Integer = 1
Global Const Core As Integer = 2
Global Const Action_Measurement As String = "Measure"
Global Const Action_AFDemag As String = "AF"
Global Const Action_ThermalDemag As String = "Thermal"
Global Const Action_ChemicalDemag As String = "Chemical"
Global Const Action_IRMDemag As String = "IRM"
Global Const Action_Measurement_Up As String = "U"
Global Const Action_Measurement_Down As String = "D"
Global Const Action_Measurement_Both As String = "B"
Global Const DemagType_NRM As String = "NRM"
Global Const DemagType_AF As String = "AF"
Global Const DemagType_Thermal As String = "TH"
Global Const DemagType_Chemical As String = "CH"
Global Const DemagType_IRM As String = "IRM"
Dim OneUp As String
Dim AFDemag As String, ThermalDemag As String, ChemicalDemag As String, IRMDemag As String
Dim t As String, C As String

Function AFSequence(Min As Double, Delta As Double, Max As Double)
    Dim SeqStr As String
    Dim R As Double
    SeqStr = OneUp
    R = Min
    While R <= Max
        SeqStr = SeqStr + C + AFDemag + Format(R) + C + OneUp
        R = R + Delta
    Wend
    AFSequence = SeqStr
End Function

Sub PerformAction(ActionString As String, Samplename As String)
End Sub

Sub SetDefaultItems(NameList As ListBox, SeqList As ListBox)
    t = Chr(9): C = Chr(13)
    Dim SeqStr As String
    OneUp = Action_Measurement + t + "1" + t + Action_Measurement_Up
    AFDemag = Action_AFDemag + t
    ThermalDemag = Action_ThermalDemag + t
    ChemicalDemag = Action_ChemicalDemag + t
    IRMDemag = Action_IRMDemag + t
    NameList.AddItem "One Up Measurement": SeqList.AddItem OneUp
    NameList.AddItem "AF Demagnetization": SeqList.AddItem AFDemag
    NameList.AddItem "Thermal Demagnetization": SeqList.AddItem ThermalDemag
    NameList.AddItem "Chemical Demagnetization": SeqList.AddItem ChemicalDemag
    NameList.AddItem "IRM Demagnetization": SeqList.AddItem IRMDemag
    NameList.AddItem "Standard 25 G steps up to 200": SeqList.AddItem AFSequence(25, 25, 200)
    NameList.AddItem "5 Gauss steps up to 200": SeqList.AddItem AFSequence(5, 5, 200)
    NameList.AddItem "12.5 Gauss steps up to 200": SeqList.AddItem AFSequence(12.5, 12.5, 200)
    NameList.AddItem "25 Gauss steps up to 200": SeqList.AddItem AFSequence(25, 25, 200)
    NameList.AddItem "50 Gauss steps up to 200": SeqList.AddItem AFSequence(50, 50, 200)
    NameList.AddItem "100 Gauss steps up to 200": SeqList.AddItem AFSequence(100, 100, 200)
    NameList.AddItem "Hawaii Standard - 25,50,100,200,...,800"
    SeqStr = OneUp + C + AFDemag + "25" + C + OneUp + C + AFDemag + "50" + C
    SeqStr = SeqStr + OneUp + C + AFDemag + "100" + C + OneUp + C + AFDemag + "200" + C
    SeqStr = SeqStr + OneUp + C + AFDemag + "400" + C + OneUp + C + AFDemag + "800" + C
    SeqList.AddItem SeqStr
'      Print "S -  Standard 25 G steps up to 200"
'      Print "    or: "
'      Print "A -  5    Gauss steps "
'      Print "B -  12.5 Gauss steps "
'      Print "C -  25   Gauss steps "
'      Print "D -  50   Gauss steps "
'      Print "E -  100  Gauss steps "
'      Print "F -  A Double step at 800 Gauss, with no measurement"
'      Print "H -  Hawaiian Standard - 25, 50, 100, 200, .. 800"
'      Print "R -  Repeat last selection, or"
'      Print "I -  Enter one at a time (Individually)"
End Sub
