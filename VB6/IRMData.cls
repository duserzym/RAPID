VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRMData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mCol As Collection

Private mvaraverage_change_entire_charging_cycle As Double
Private mvaraverage_change_over_window As Double
Private mvarwindow_size As Integer
Private num_total_points As Long
Private array_size As Integer
Private collection_size As Integer

Public Enum IrmDataKeyTypeEnum
    time = 0
    Index = 1
End Enum

Public Property Let average_change_entire_charging_cycle(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.read_voltage = 5
    mvaraverage_change_entire_charging_cycle = vData
End Property

Public Property Get average_change_entire_charging_cycle() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.read_voltage
    average_change_entire_charging_cycle = mvaraverage_change_entire_charging_cycle
End Property

Public Property Let average_change_over_window(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.read_voltage = 5
    mvaraverage_change_over_window = vData
End Property

Public Property Get average_change_over_window() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.read_voltage
    average_change_over_window = mvaraverage_change_over_window
End Property

Public Property Get Item(key As Variant) As IrmDataPoint
    
    On Error GoTo DataPointGetError
    
    Set Item = mCol.Item(CInt(key))
        
    On Error GoTo 0
    Exit Property
                
DataPointGetError:

    Set Item = Nothing
            
End Property

Public Property Get NewIndex() As Long
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    
    If mCol Is Nothing Then Set mCol = New Collection
        
    NewIndex = mCol.Count
        
End Property

Public Property Let window_size(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.read_voltage = 5
    mvarwindow_size = vData
End Property

Public Property Get window_size() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.read_voltage
    window_size = mvarwindow_size
    collection_size = window_size * 10
End Property

Public Function Add(ByVal time As Long, ByVal irm_read_voltage As Double) As IrmDataPoint
    
    If mCol Is Nothing Then
        
        Set mCol = New Collection
        
    End If
    
    If mCol.Count >= collection_size Then
    
        mCol.Remove (1)
        
    End If
    
    Dim irm_data_point As IrmDataPoint
        
    Set irm_data_point = New IrmDataPoint
    irm_data_point.time_stamp = time
    irm_data_point.read_voltage = irm_read_voltage
        
    Dim window_start As Integer
    window_start = mCol.Count - Me.window_size - 1
    
    If mCol.Count > 1 Then
        
        Dim last_data_point As IrmDataPoint
        Set last_data_point = mCol.Item(mCol.Count)
       
        Dim delta_t As Double
        Dim delta_v As Double
        
        delta_t = irm_data_point.time_stamp - last_data_point.time_stamp
        delta_v = irm_data_point.read_voltage - last_data_point.read_voltage
       
        If delta_t <> 0 Then
            
            irm_data_point.delta_rate = delta_v / delta_t
            
        Else
        
            irm_data_point.delta_rate = 0
        
        End If
        
    Else
    
        irm_data_point.delta_rate = 0
            
    End If
                                                  
                                                  
    Me.average_change_entire_charging_cycle = Me.average_change_entire_charging_cycle * num_total_points + _
                                              irm_data_point.delta_rate
    
    Me.average_change_entire_charging_cycle = Me.average_change_entire_charging_cycle / (num_total_points + 1)
        
    
    If window_start > 0 Then
           
        Dim window_start_point As IrmDataPoint
        Set window_start_point = mCol.Item(window_start)
                
        Me.average_change_over_window = Me.average_change_over_window * Me.window_size - _
                                        window_start_point.delta_rate + _
                                        irm_data_point.delta_rate
                                                                                
        Me.average_change_over_window = Me.average_change_over_window / Me.window_size
    
    Else
    
        Me.average_change_over_window = Me.average_change_entire_charging_cycle
            
    End If
    
    
    mCol.Add irm_data_point
    num_total_points = num_total_points + 1
        
    'return the object created
    Set Add = irm_data_point

End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    
    If Me.window_size <= 0 Then Me.window_size = 20
    If collection_size <= 0 Then collection_size = 10 * Me.window_size
    If num_total_points > 0 Then num_total_points = 0
    Me.average_change_entire_charging_cycle = 0
    Me.average_change_over_window = 0
    
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Me.Clear
    Set mCol = Nothing
End Sub

Public Sub Clear()

    On Error GoTo Clear_Error

    If mCol Is Nothing Then Set mCol = New Collection: Exit Sub
    
    Dim i As Integer
    
    For i = 0 To mCol.Count - 1
    
        mCol.Remove (i)
    
    Next i
    
    On Error GoTo 0
    
Clear_Error:

    Set mCol = Nothing
    Set mCol = New Collection

End Sub

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
    
End Property

Private Function GetIrmDataPoint_ByTimeValue(ByVal time As Date) As IrmDataPoint
    
    On Error GoTo GetIrmDataPoint_ByTimeValue_Error
    
        Dim Index As Integer
        Index = GetIrmDataPointIndex_ByTimeValue(time)
    
        Set GetIrmDataPoint_ByTimeValue = mCol.Item(Index)
    
    On Error GoTo 0
    Exit Function
    
GetIrmDataPoint_ByTimeValue_Error:
    
    Set GetIrmDataPoint_ByTimeValue = Nothing
    
End Function

Private Function GetIrmDataPointIndex_ByTimeValue(ByVal time As Date) As Integer

    On Error GoTo GetIrmDataPointIndex_ByTimeValue_Error

    If mCol Is Nothing Then
    
        GetIrmDataPointIndex_ByTimeValue = -1
        Exit Function
    
    End If

    'Search two end-points
    If time > mCol.Item(mCol.Count - 1) Then
    
        GetIrmDataPointIndex_ByTimeValue = -1
        Exit Function
        
    End If
    
    If time < mCol.Item(0) Then
    
        GetIrmDataPointIndex_ByTimeValue = -1
        Exit Function
    
    End If
    
    GetIrmDataPointIndex_ByTimeValue = RecursiveSearch_ByTime(time, 0, mCol.Count - 1)
    
    On Error GoTo 0
    Exit Function
    
GetIrmDataPointIndex_ByTimeValue_Error:

    GetIrmDataPointIndex_ByTimeValue = -1
    

End Function

Private Function RecursiveSearch_ByTime(ByVal time_stamp As Date, _
                                        ByVal start_index As Integer, _
                                        ByVal end_index As Integer) As Integer



    Dim start_point As IrmDataPoint
    Dim end_point As IrmDataPoint
    
    Set start_point = mCol.Item(start_index)
    Set end_point = mCol.Item(end_index)
    
    If time_stamp = start_point.time_stamp Then
    
        RecursiveSearch_ByTime = start_index
        
    ElseIf time_stamp = end_point.time_stamp Then
    
        RecursiveSearch_ByTime = end_index
        
    ElseIf time_stamp > start_point.time_stamp And _
           time_stamp < end_point.time_stamp Then
       
        'Time is bounded in this region of the collection
        If end_index - start_index <= 1 Then
       
            RecursiveSearch_ByTime = start_index
            Exit Function
            
        End If
        
        If end_index - start_index = 2 Then
        
            RecursiveSearch_ByTime = start_index + 1
            Exit Function
            
        End If
        
        'Otherwise, split the section in two
        Dim upper_result As Integer
        Dim lower_result As Integer
        
        upper_result = RecursiveSearch_ByTime(time_stamp, start_index, CInt((end_index - start_index) / 2) + start_index)
        
        If upper_result > 0 Then
            RecursiveSearch_ByTime = upper_result
            Exit Function
        End If
        
        lower_result = RecursiveSearch_ByTime(time_stamp, CInt((end_index - start_index) / 2) + start_index + 1, end_index)
        
        RecursiveSearch_ByTime = lower_result
        
    Else
    
        RecursiveSearch_ByTime = -1
        
    End If

End Function

Public Sub Remove(key As Variant, key_type As IrmDataKeyTypeEnum)
    On Error GoTo DeleteDataPointError
    
    Dim Index As Integer
    
    Select Case key_type
    
        Case IrmDataKeyTypeEnum.time
        
            Index = GetIrmDataPointIndex_ByTimeValue(CDate(key))
        
        Case Else
        
            Index = CInt(key)
    
    End Select
        
    mCol.Remove Index
        
    On Error GoTo 0
    Exit Sub
                   
DeleteDataPointError:
            
End Sub

