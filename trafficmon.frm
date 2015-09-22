VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_trafficmon 
   Caption         =   "TrafficMon"
   ClientHeight    =   1725
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox P_in 
      Height          =   375
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox P_out 
      Height          =   375
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5520
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBarOut 
      Height          =   222
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBarIn 
      Height          =   222
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   1320
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "out:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "in:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Menu m_set 
      Caption         =   "&Settings"
      Begin VB.Menu m_int 
         Caption         =   "&Interval"
         Shortcut        =   ^I
      End
      Begin VB.Menu m_max 
         Caption         =   "&max Values"
         Shortcut        =   ^M
      End
      Begin VB.Menu m_automax 
         Caption         =   "&auto max"
         Checked         =   -1  'True
      End
      Begin VB.Menu m_filter 
         Caption         =   "&filter"
      End
   End
   Begin VB.Menu m_help 
      Caption         =   "&?"
      Begin VB.Menu m_about 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "F_trafficmon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc _
    As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
     

Dim m_cin As Long
Dim m_cout As Long
Dim m_din As Long
Dim m_dout As Long
Dim m_maxIn As Double
Dim m_maxOut As Double
Const max_meter = 100
Dim oFilterIn As New C_filter
Dim oFilterOut As New C_filter
Dim m_bgColor As ColorConstants

Private Sub Form_Load()
    Caption = "TrafficMon Rev." + Mid("$Rev:: 955  $", 7, 4)
    m_maxIn = max_meter
    m_maxOut = max_meter
    Timer1.Interval = 750
    ProgressBarIn.max = max_meter
    ProgressBarOut.max = max_meter
    
    oFilterIn.filterLen = 2
    oFilterOut.filterLen = 2
    'm_bgColor = ColorConstants.vbYellow
    
    P_in.ForeColor = ColorConstants.vbYellow
    P_out.ForeColor = ColorConstants.vbRed
    P_in.BackColor = ColorConstants.vbBlack
    P_out.BackColor = ColorConstants.vbBlack
End Sub

Private Sub m_about_Click()
    MsgBox Me.Caption + vbCrLf + vbCrLf + "This is freeware." + vbCrLf + "2015 by lifeSim.de", vbOKOnly, "About " + Me.Caption
End Sub

Private Sub m_automax_Click()
    m_automax.Checked = Not m_automax.Checked
End Sub

Private Sub m_filter_Click()
    Dim f As Integer
    f = oFilterIn.filterLen
    f = Val(InputBox("New Filterlength: ", "Filter", Str(f)))
    If f > 0 And f < 1000 Then
        oFilterIn.filterLen = f
        oFilterOut.filterLen = f
    End If
End Sub

Private Sub m_int_Click()
    Dim s As String
    Dim i As Integer
    
    s = InputBox("New Inverval in ms:", "Set Interval", Str(Timer1.Interval))
    i = Val(s)
    If i < 200 Then i = 75
    Timer1.Interval = i
        
End Sub

Private Sub Timer1_Timer()
    Dim s As String
    Dim lines() As String
    Static oi As Long
    Static oo As Long
    Dim r As String

    s = GetCommandOutput("netstat -e")
    lines = helper.TextToLines(s)
    r = lines(4)
    While InStr(r, "  ")
        r = Replace(r, "  ", " ")
    Wend
    lines = Split(r, " ")
    m_cin = Val(lines(1))
    m_cout = Val(lines(2))
    
    If oi = 0 Then oi = m_cin   '1st time?
    m_din = m_cin - oi
    m_din = oFilterIn.filter(m_din)
    oi = m_cin
    
    If oo = 0 Then oo = m_cout  '1st time?
    m_dout = m_cout - oo
    m_dout = oFilterOut.filter(m_dout)
    oo = m_cout
    
    If m_automax.Checked Then
        Call scaleMax(m_maxIn, m_din)
        scaleMax m_maxOut, m_dout
    End If
    
    updateGfx
    
End Sub
Sub updateGfx()
    Dim d As Double
    Dim s As String
    
    d = m_din
    s = CStr(Round(d / 1024, 1)) + " KiB/s. |" + vbCrLf + "100%=" + CStr(Round(m_maxIn / 1024, 1)) + " KiB/s."
    ProgressBarIn.ToolTipText = s  'CStr(Round(d / 1024, 1)) + " KiB/s."
    P_in.ToolTipText = s
    d = scaled(m_din, m_maxIn) * ProgressBarIn.max
    If d > ProgressBarIn.max Then d = ProgressBarIn.max
    ProgressBarIn.value = d
    
    d = m_dout
    s = CStr(Round(d / 1024, 1)) + " KiB/s. |" + vbCrLf + "100%=" + CStr(Round(m_maxOut / 1024, 1)) + " KiB/s."
    ProgressBarOut.ToolTipText = s      'CStr(Round(d / 1024, 1)) + " KiB/s."
    P_out.ToolTipText = s
    d = scaled(m_dout, m_maxOut) * ProgressBarOut.max
    If d > ProgressBarOut.max Then d = ProgressBarOut.max
    ProgressBarOut.value = d
    
    Text2 = m_cin
    Text3 = m_cout
    
    Text1 = m_maxIn
    Text4 = m_maxOut
    
    drawChart P_in, m_din, m_maxIn
    drawChart P_out, m_dout, m_maxOut
    

End Sub

Sub drawChart(pic As PictureBox, ByVal value As Double, ByVal max As Long)
    Dim w As Integer
    
    Call BitBlt(pic.hDC, 0, 0, pic.Width, pic.Height, pic.hDC, 1, 0, vbSrcCopy)
    value = value * pic.Height
    If max = 0 Then max = 1
    value = value / max
    value = pic.Height - value
    
    w = pic.Width - 100
    pic.Line (w, 0)-(w, value), pic.BackColor
    pic.Line (w, value)-(w, pic.Height), pic.ForeColor
End Sub

Function scaled(ByVal diff As Long, ByVal max As Double) As Double
    'results in max=100%
    Dim d As Double
    On Error GoTo hell
    
    d = diff * 100  '100%
    If max = 0 Then max = 0.01
    d = d / max
    d = d / Timer1.Interval
    If d < 0 Then d = 0
    scaled = d
    Exit Function
hell:
    scaled = max
End Function
Function scaleMax(ByRef max As Double, ByVal value As Long) As Integer
    If value > max Then
        max = max + (value - max) / 2   'go smooth to upper border
        max = value
    End If
    'adjust if value < 10%
    If value < max / 10 Then
        max = max - (max - value) / 13   'go very smooth to lower border
    End If
'    max = max
End Function

  
