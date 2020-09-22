VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjBar As PictureBox

Private mlCurrentValue As Long
Private mlMin As Long
Private mlMax As Long
Private mlForeColor As Long
Private mlBackColor As Long
Private miCellWidth As Integer
Private miCellHeight As Integer
Private miXMargin As Integer
Private miYMargin As Integer
Private miCells As Integer
Private mbRecursing As Boolean

Private Sub Class_Initialize()
    Set mobjBar = Nothing
    mlForeColor = vbHighlight
    mlBackColor = vb3DFace
    mlMax = 100
End Sub

Public Property Set PictureBox(Obj As PictureBox)
    On Error GoTo ErrHandler
    
    If Not Obj Is Nothing Then
        Set mobjBar = Obj
        mobjBar.ForeColor = mlForeColor
        mobjBar.BackColor = mlBackColor
        mobjBar.ScaleMode = vbPixels
        Call ResizePictureBox
    Else
        Err.Raise vbObjectError, "[clsProgressBar.PictureBox.Set]", "Invalid PictureBox object."
    End If
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.PictureBox.Set]" & Err.Description
End Property

Public Property Get PictureBox() As Object
    On Error GoTo ErrHandler
    
    Set PictureBox = mobjBar
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.PictureBox.Get]" & Err.Description
End Property

Public Property Let Value(NewVal As Long)
    On Error GoTo ErrHandler
    
    If NewVal <> mlCurrentValue Then
        Select Case NewVal
            Case Is > mlMax
                mlCurrentValue = mlMax
            Case Is < mlMin
                mlCurrentValue = mlMin
            Case Else
                mlCurrentValue = NewVal
        End Select
        Call Refresh
    End If
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.Value.Let]" & Err.Description
End Property

Public Property Get Value() As Long
    On Error GoTo ErrHandler
   
    Value = mlCurrentValue
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.Value.Get]" & Err.Description
End Property

Public Property Let Min(NewVal As Long)
    On Error GoTo ErrHandler
   
    If NewVal < mlMax Then
        mlMin = NewVal
        Call Refresh
    Else
        Err.Raise vbObjectError, "[clsProgressBar.Min.Let]", "Min property must be less than Max."
    End If
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.Min.Let]" & Err.Description
End Property

Public Property Get Min() As Long
    On Error GoTo ErrHandler
   
    Min = mlMin
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.Min.Get]" & Err.Description
End Property

Public Property Let Max(NewVal As Long)
    On Error GoTo ErrHandler
   
    If NewVal > mlMin Then
        mlMax = NewVal
        Call Refresh
    Else
        Err.Raise vbObjectError, "[clsProgressBar.Max.Let]", "Max property must be greater than Min."
    End If
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.Max.Let]" & Err.Description
End Property

Public Property Get Max() As Long
    On Error GoTo ErrHandler
   
    Max = mlMax
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.Max.Get]" & Err.Description
End Property

Public Property Let ForeColor(NewVal As Long)
    On Error GoTo ErrHandler
   
    mlForeColor = NewVal
    If Not mobjBar Is Nothing Then
        mobjBar.ForeColor = mlForeColor
        Call Refresh
    End If
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.ForeColor.Let]" & Err.Description
End Property

Public Property Get ForeColor() As Long
    On Error GoTo ErrHandler
   
    ForeColor = mlForeColor
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.ForeColor.Get]" & Err.Description
End Property

Public Property Let BackColor(NewVal As Long)
    On Error GoTo ErrHandler
   
    mlBackColor = NewVal
    If Not mobjBar Is Nothing Then
        mobjBar.BackColor = mlBackColor
        Call Refresh
    End If
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.ForeColor.Get]" & Err.Description
End Property

Public Property Get BackColor() As Long
    On Error GoTo ErrHandler
   
    BackColor = mlBackColor
    Exit Property
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.BackColor.Get]" & Err.Description
End Property

Public Sub Refresh(Optional ClearFirst As Boolean)
    On Error GoTo ErrHandler
   
    If Not mobjBar Is Nothing Then
        If ClearFirst Then mobjBar.Cls
        Call RedrawMe
    End If
    Exit Sub
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.Refresh]" & Err.Description
End Sub

Public Sub Resize()
    On Error GoTo ErrHandler
   
    Call CalcCellSize
    Call Refresh(True)
    Exit Sub
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.Resize]" & Err.Description
End Sub

Public Sub ResizePictureBox()
    On Error GoTo ErrHandler
   
    Dim iActual As Integer
    Dim iPicture As Integer
    Dim iLarge As Integer
    Dim iSmall As Integer
    Dim iExtra As Integer

    If mbRecursing Then Exit Sub
    If mobjBar Is Nothing Then Exit Sub
    
    Call CalcCellSize
   
    iActual = miXMargin + miCells * (miCellWidth + miXMargin)
    iPicture = mobjBar.ScaleWidth
    iLarge = iActual - iPicture
    iSmall = iPicture - (iActual - miCellWidth - miXMargin)
   
    If iSmall <= iLarge Then
        miCells = miCells - 1
        iActual = miXMargin + miCells * (miCellWidth + miXMargin)
    End If
   
    iExtra = mobjBar.Width - (iPicture * Screen.TwipsPerPixelX)
    mbRecursing = True
    mobjBar.Width = iActual * Screen.TwipsPerPixelX + iExtra
    mbRecursing = False
    Call Refresh(True)
    Exit Sub
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.ResizePictureBox]" & Err.Description
End Sub

Private Sub RedrawMe()
    On Error GoTo ErrHandler
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim n As Long
   
    ' Calc number of live cells to draw.
    n = (mlCurrentValue / (mlMax - mlMin)) * miCells
   
    ' Draw live cells.
    mobjBar.ForeColor = mlForeColor
    y = miYMargin + miCellHeight - 1
    x = miYMargin
    For i = 1 To n
        mobjBar.Line (x, miYMargin)-(x + miCellWidth, y), , BF
        x = x + miXMargin + miCellWidth
    Next i

    ' Draw dead cells.
    If n < miCells Then
        mobjBar.ForeColor = mlBackColor
        For i = n + 1 To miCells
            mobjBar.Line (x, miYMargin)-(x + miCellWidth, y), , BF
            x = x + miXMargin + miCellWidth
        Next i
    End If
    Exit Sub
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.RedrawMe]" & Err.Description
End Sub

Private Sub CalcCellSize()
    On Error GoTo ErrHandler
    
    miYMargin = 2
    miXMargin = 3
    miCellHeight = mobjBar.ScaleHeight - miYMargin * 2
    miCellWidth = miCellHeight * (2 / 3)
    miCells = mobjBar.ScaleWidth \ (miCellWidth + miXMargin) + 1
    Exit Sub
   
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsProgressBar.CalcCellSize]" & Err.Description
End Sub

Private Sub Class_Terminate()
    Set mobjBar = Nothing
End Sub
