VERSION 5.00
Begin VB.Form frmCountdown 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Countdown Timer"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbTextBuffer 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      ScaleHeight     =   735
      ScaleWidth      =   3555
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5355
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Default         =   -1  'True
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtSecs 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3420
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "0"
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtMins 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "1"
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtHours 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   720
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2580
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1260
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      Height          =   1035
      Left            =   3900
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   1395
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   60
      ScaleHeight     =   255
      ScaleWidth      =   5295
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1860
      Width           =   5355
   End
End
Attribute VB_Name = "frmCountdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written By David Drake

Option Explicit
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
   
Private WithEvents mobjSlumber As clsSlumber
Attribute mobjSlumber.VB_VarHelpID = -1
Private mobjProgBar As clsProgressBar
Private mobjPieBar As clsPieBar

Private mdTotalTime As Double
Private mdCurrentTime As Double
Private mdCurrentHours As Long
Private mdCurrentMins As Long
Private mdCurrentSecs As Double
Private mbInterrupted As Boolean

Private Sub cmdStart_Click()
    On Error GoTo ErrHandler
    Dim sMsg As String
    
    If cmdStart.Caption = "Stop" Then
        mbInterrupted = True
        mobjSlumber.WakeUp
        Exit Sub
    End If
        
    'Ensure Data is Numeric
    txtHours.Text = Val(txtHours.Text)
    txtMins.Text = Val(txtMins.Text)
    txtSecs.Text = Val(txtSecs.Text)
    
    'Ensure Values are whole numbers
    txtHours.Text = Int(txtHours.Text)
    txtMins.Text = Int(txtMins.Text)
    txtSecs.Text = Int(txtSecs.Text)
       
    'Validate Input Data
    Select Case Val(txtHours.Text)
        Case Is < 0
            sMsg = "Hours cannot be a negative number!"
            GoTo DisplayMsg
    End Select
    
    Select Case Val(txtMins.Text)
        Case Is < 0
            sMsg = "Minutes cannot be a negative number!"
            GoTo DisplayMsg
        Case Is >= 60
            sMsg = "Minutes cannot be greater than or equal to 60!"
            GoTo DisplayMsg
    End Select
    
    Select Case Val(txtSecs.Text)
        Case Is < 0
            sMsg = "Seconds cannot be a negative number!"
            GoTo DisplayMsg
        Case Is >= 60
            sMsg = "Seconds cannot be greater than or equal to 60!"
            GoTo DisplayMsg
    End Select

    If Val(txtHours.Text) = 0 And Val(txtMins.Text) = 0 And Val(txtSecs.Text) = 0 Then
        sMsg = "Please input countdown time!"
        GoTo DisplayMsg
    End If
    
DisplayMsg:
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If
    
    'Lock Input Boxes
    txtHours.Locked = True
    txtMins.Locked = True
    txtSecs.Locked = True
    mbInterrupted = False
    cmdStart.Caption = "Stop"
    
    'Set to Slumber every 10 milliseconds
    '(This means that clsSlumber will put the thread to sleep for
    ' 10 milliseconds at a time until the total time has elapsed)
    mobjSlumber.SlumberInterval = 10
    
    'Calculate total number of milliseconds
    mdTotalTime = (((Val(txtHours.Text) * 60) + Val(txtMins.Text)) * 60) + Val(txtSecs.Text)
    
    'Start Slumbering
    mobjSlumber.Slumber CLng(mdTotalTime * 1000)
    
    'Reset form state
    cmdStart.Caption = "Start"
    If Not mbInterrupted Then
        Call Display("000.00.00.0000")
        mobjProgBar.Value = 100
        mobjPieBar.Value = 100
    End If

    'UnLock Input Boxes
    txtHours.Locked = False
    txtMins.Locked = False
    txtSecs.Locked = False
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    pbTextBuffer.BackColor = vbBlack
    pbTextBuffer.Cls
    pbTextBuffer.FontBold = True
    pbTextBuffer.FontSize = 24
    pbTextBuffer.ForeColor = vbGreen
    pbTextBuffer.ScaleMode = vbPixels
    pbTextBuffer.AutoRedraw = True
    pbTextBuffer.Visible = False
            
    Set mobjSlumber = New clsSlumber
    Set mobjProgBar = New clsProgressBar
    Set mobjProgBar.PictureBox = Picture1
    mobjProgBar.BackColor = Picture1.BackColor
    Set mobjPieBar = New clsPieBar
    Set mobjPieBar.PictureBox = Picture2
    Me.ScaleMode = vbPixels
    Me.Show
    mobjProgBar.Value = 100
    mobjPieBar.Value = 100
    DoEvents
    Call ConvertTime
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjSlumber = Nothing
    Set mobjProgBar = Nothing
    Set mobjPieBar = Nothing
    End
End Sub

Private Sub mobjSlumber_Slumber()
    On Error GoTo ErrHandler
    
    If mobjSlumber.ElapsedMilliseconds > 0 Then
        mdCurrentTime = mdTotalTime - CDbl(mobjSlumber.ElapsedMilliseconds / 1000)
        mobjProgBar.Value = (100 * (mobjSlumber.ElapsedMilliseconds / (mdTotalTime * 1000)))
        mobjPieBar.Value = mobjProgBar.Value
    Else
        mobjProgBar.Value = 0
        mobjPieBar.Value = 0
        mdCurrentTime = mdTotalTime
    End If
    mdCurrentHours = Int(CStr(mdCurrentTime)) \ 3600
    mdCurrentSecs = mdCurrentTime - (mdCurrentHours * 3600)
    mdCurrentMins = Int(CStr(mdCurrentSecs)) \ 60
    mdCurrentSecs = mdCurrentSecs - (mdCurrentMins * 60)
        
    Call Display(Format$(mdCurrentHours, "000") & "." & Format$(mdCurrentMins, "00") & "." & Format$(mdCurrentSecs, "00.0000"))
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[frmCountdown.mobjSlumber_Slumber]" & Err.Description
End Sub

Private Sub Display(Text As String)
    With pbTextBuffer
        .Cls
        .CurrentX = 0
        .CurrentY = 10
        pbTextBuffer.Print Text
        BitBlt Me.hDC, .Left, .Top, .ScaleWidth, .ScaleHeight, .hDC, 0, 0, SRCCOPY
    End With
End Sub

Private Sub ConvertTime()
    Call Display(Format$(Val(txtHours.Text), "000") & "." & Format$(Val(txtMins.Text), "00") & "." & Format$(Val(txtSecs.Text), "00") & ".0000")
End Sub

Private Sub txtHours_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ConvertTime
End Sub

Private Sub txtMins_Change()
    Call ConvertTime
End Sub

Private Sub txtSecs_Change()
    Call ConvertTime
End Sub
