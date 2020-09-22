VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPieBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written By David Drake

Option Explicit
Private Const PI = 3.14159265
Private Const PIx2 = 6.2831853
Private Const PIx133 = 4.712388975
Private Const PIv5 = 1.570796325

Private mobjPie As PictureBox
Private mlProgress As Long
Private mdStatusValue As Double
Private mdCenterX As Double
Private mdCenterY As Double
Private mdRadius As Double

Property Set PictureBox(Obj As PictureBox)
    On Error GoTo ErrHandler
    
    If Obj Is Nothing Then Err.Raise vbObjectError, "[clsPieBar.PictureBox.Set]", "Invalid PictureBox object!"
    If Not TypeOf Obj Is PictureBox Then Err.Raise vbObjectError, "[clsPieBar.PictureBox.Set]", "Object is not of Type PictureBox!"
    
    Set mobjPie = Nothing
    Set mobjPie = Obj
    
    With mobjPie
        .ScaleMode = vbPixels
        .FillStyle = vbFSSolid
        .ScaleMode = 3   ' Set scale to pixels.
        .Cls

        mdCenterX = .ScaleWidth / 2
        mdCenterY = .ScaleHeight / 2
    End With
    
    If mdCenterX > mdCenterY Then mdRadius = mdCenterY Else mdRadius = mdCenterX
    mdStatusValue = 0
    Exit Property
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsPieBar.PictureBox.Set]" & Err.Description
End Property

Public Property Let Value(StatusValue As Double)
    On Error GoTo ErrHandler
    Dim dRadians As Double
    
    If mobjPie Is Nothing Then Err.Raise vbObjectError, "[clsPieBar.Value.Let]", "PictureBox property has not been initialized!"
    
    mdStatusValue = StatusValue

    If StatusValue = 0 Then
        mlProgress = 0
    Else
         dRadians = Int(CStr(StatusValue))
        
         If dRadians > mlProgress Then
             mlProgress = dRadians
         Else
             Exit Property
         End If
    End If
   
    dRadians = PIx2 * (StatusValue / -100)
    
    If dRadians = 0 Then
        mobjPie.Cls
        Exit Property
    End If
    
    If dRadians = -PIx2 Then
         mobjPie.Circle (mdCenterX, mdCenterY), mdRadius
    ElseIf dRadians <= -PIx133 Then
        mobjPie.Circle (mdCenterX, mdCenterY), mdRadius, , PIv5 * -1, dRadians + PIx133
    Else
        mobjPie.Circle (mdCenterX, mdCenterY), mdRadius, , PIv5 * -1, dRadians - PIv5
    End If
    Exit Property
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsPieBar.Value.Let]" & Err.Description
End Property

Public Property Get Value() As Double
    On Error GoTo ErrHandler
    Value = mdStatusValue
    Exit Property
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsPieBar.Value.Get]" & Err.Description
End Property

Private Sub Class_Terminate()
    Set mobjPie = Nothing
End Sub
