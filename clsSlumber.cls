VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSlumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************
' Name:        clsSlumber
'
' Description: This class encapsulate the QueryPerfomanceXXX and Sleep
'              API functions to put the thread of your application to
'              Sleep for a period of time. The Sleep API is preferable to
'              a loop since it does not consume Processor time. The problem
'              is that during a Sleep, no events can be received. Slumber
'              provides an encapsulated alternative to Sleep. Slumber will
'              Sleep for small intervals at a time, after each return from
'              Sleep it will call DoEvents so that your application can
'              receive events. It also raises its own Slumber event so
'              that if you want your application to check on something each
'              Slumber. If you want to Exit out of the Slumber period during
'              one of these events, you can use the WakeUp property to
'              notify that it is time to wake up. Otherwise, clsSlumber will
'              put the thread asleep again until either a WakeUp is indicated
'              or the Slumber period has finished. The SlumberInterval property
'              allows you to set the interval time for the Sleep cycles, by default
'              this is set for 0.5 seconds or 500 milliseconds.
'              By using QueryPerfomanceXXX, this class mesures time to a precision
'              of 0.0001 seconds, 10 times better then timeGetTime() or GetTickCount()
'              APIs which have a precision of 0.1 seconds.
'
' Application: This class is very versatile. It can be used as a simple way to put
'              your app on idle while it is waiting for something, or it can be used
'              to keep track of complicated asynchronous processes.
'
' Example 1:   'This is an example for idling your application
'              Private WithEvents mobjSlumber As clsSlumber
'              Private Sub RunProcess()
'                Set mobjSlumber = New clsSlumber
'                Do
'                 If mbWorkToDo Then
'                   Call ProcessWork()
'                 Else
'                   mobjSlumber.Slumber(5000) 'Slumber for 5 seconds
'                 End If
'                 If mbStop Then Exit Do
'                Loop
'                Set mobjSlumber = nothing
'              End Sub
'
'              Private Sub mobjSlumber_Slumber()
'                'Now every 0.5 seconds, it will check to WakeUp
'                If mbStop or mbWorkToDo Then mobjSlumber.WakeUp
'              End Sub
'
' Example 2:   'This is an example for asyncronous processing
'              Private WithEvents mobjSlumber As clsSlumber
'              Private Const GETDATA_TIMEOUT = 60000 'Timeout after 1 minute
'
'              Private Sub GetData()
'                Set mobjSlumber = New clsSlumber
'                Call mobjAsynchGetData.Start()
'
'                mobjSlumber.Slumber(GETDATA_TIMEOUT) 'Slumber for Timeout Period
'
'                If Not mobjAsynchGetData.DataArrived Then
'                  Err.Raise vbObjectError, "GetData", "A Timeout occurred while trying to get data!"
'                End if
'
'                Set mobjSlumber = nothing
'              End Sub
'
'              Private Sub mobjSlumber_Slumber()
'                'Every 0.5 seconds, it will check for data arrival
'                If mobjAsynchGetData.DataArrived Then
'                  mobjSlumber.WakeUp
'                Elseif
'                  'You can call back into clsSlumber to figure out the percentage of time elapsed
'                  Debug.Print Format$((mobjSlumber.ElapsedMilliseconds / GETDATA_TIMEOUT) * 100), "0.00") & "% to Timeout"
'                End If
'              End Sub
'
' Revision History:
' 02/01/2000 - DDRAKE - Created class using timGetTime() API
' 03/03/2000 - DDRAKE - Added Slumber Event
' 12/13/2000 - DDRAKE - Changed class to use QueryPerformanceCounter() API for improved accuracy
'                     - Added ElapsedMilliseconds Property

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Large) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Large) As Long

Private Type Large
    Low As Long
    High As Long
End Type

Private Const CONST_2_31 = 2147483648#  '2^31
Private Const CONST_2_32 = 4294967296#  '2^32

Private miInterval As Long
Private mbWakeUp As Boolean
Private mdFrequency As Double
Private mdElapsedSeconds As Double

Public Event Slumber()

Private Sub Class_Initialize()
    Dim mlPerfFrequency As Large
    
    QueryPerformanceFrequency mlPerfFrequency
    mdFrequency = CDouble(mlPerfFrequency)
    
    If mdFrequency = 0 Then Err.Raise vbObjectError, "[clsSlumber.Class_Initialize]", "[clsSlumber.Class_Initialize]Error while initializing clsSlumber!"
    
    miInterval = 500 '0.5 Seconds
End Sub

Public Property Get SlumberInterval() As Long
    On Error GoTo GetSlumberInterval_Err
    SlumberInterval = miInterval
    Exit Property
    
GetSlumberInterval_Err:
    Err.Raise Err.Number, Err.Source, "[clsSlumber.Get.SlumberInterval]" & Err.Description
End Property

Public Property Let SlumberInterval(MilliSeconds As Long)
    On Error GoTo LetSlumberInterval_Err
    If MilliSeconds <= 0 Then Err.Raise vbObjectError, "clsSlumber.Let.SlumberInterval", "SlumberInterval must be greater than 0!"
    miInterval = MilliSeconds
    Exit Property
    
LetSlumberInterval_Err:
    Err.Raise Err.Number, Err.Source, "[clsSlumber.Let.SlumberInterval]" & Err.Description
End Property

Public Property Get ElapsedMilliseconds() As Double
    On Error GoTo ElapsedMilliseconds_Err
    'Timing is accurate to 0.1 Milliseconds
    ElapsedMilliseconds = Int(CStr(mdElapsedSeconds * 10000)) / 10
    Exit Property
    
ElapsedMilliseconds_Err:
    Err.Raise Err.Number, Err.Source, "[clsSlumber.Get.ElapsedMilliseconds]" & Err.Description
End Property

Private Function CDouble(LargeNum As Large) As Double
    On Error GoTo ErrHandler
    With LargeNum
        If .Low > 0& Then
            CDouble = .Low + (.High * CONST_2_32)
        Else
            CDouble = CONST_2_31 + CDbl(.Low And &H7FFFFFFF) + (.High * CONST_2_32)
        End If
    End With
    Exit Function

ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsSlumber.CDouble]" & Err.Description
End Function

Private Function DiffDbl(Big As Large, Small As Large) As Double
    On Error GoTo ErrHandler
    With Small
        If .Low > 0& Then
            If Big.Low > 0& Then
                DiffDbl = Big.Low - .Low + (Big.High - .High) * CONST_2_32
            Else
                DiffDbl = CONST_2_31 + CDbl(Big.Low And &H7FFFFFFF) - .Low + (Big.High - .High) * CONST_2_32
            End If
        Else
            If Big.Low > 0& Then
                DiffDbl = Big.Low - CONST_2_31 - CDbl(.Low And &H7FFFFFFF) + (Big.High - .High) * CONST_2_32
            Else
                DiffDbl = CDbl(Big.Low And &H7FFFFFFF) - CDbl(.Low And &H7FFFFFFF) + (Big.High - .High) * CONST_2_32
            End If
        End If
    End With
    Exit Function

ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsSlumber.DiffDbl]" & Err.Description
End Function

Public Sub Slumber(MilliSeconds As Long)
    On Error GoTo Slumber_Err
    Dim dSecs As Double
    Dim lCounterStart As Large
    Dim lCounterEnd As Large
    
    If MilliSeconds <= 0 Then Err.Raise vbObjectError, "[clsSlumber.Slumber]", "Parameter MilliSeconds must be greater than 0!"
    
    dSecs = MilliSeconds / 1000
    
    mbWakeUp = False
    mdElapsedSeconds = 0
    QueryPerformanceCounter lCounterStart
    Do
        Sleep miInterval
        
        On Error GoTo SlumberEvent_Err
        RaiseEvent Slumber
        DoEvents
        
        On Error GoTo Slumber_Err
        QueryPerformanceCounter lCounterEnd
        
        mdElapsedSeconds = DiffDbl(lCounterEnd, lCounterStart) / mdFrequency
        
        If mdElapsedSeconds >= dSecs Or mbWakeUp Then Exit Do
    Loop
    mdElapsedSeconds = dSecs
    Exit Sub
    
Slumber_Err:
    Err.Raise Err.Number, Err.Source, "[clsSlumber.Slumber]" & Err.Description
    
SlumberEvent_Err:
    Err.Raise Err.Number, Err.Source, "[clsSlumber.Slumber][Slumber.Event]" & Err.Description
End Sub

Public Sub WakeUp()
    On Error GoTo WakeUp_Err
    mbWakeUp = True
    Exit Sub
    
WakeUp_Err:
    Err.Raise Err.Number, Err.Source, "[clsSlumber.WakeUp]" & Err.Description
End Sub


