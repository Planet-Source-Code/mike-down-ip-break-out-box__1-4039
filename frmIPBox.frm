VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIPBox 
   Caption         =   "IP BreakOut Box By Mike Down"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CLS"
      Height          =   375
      Left            =   9390
      TabIndex        =   12
      Top             =   30
      Width           =   435
   End
   Begin MSWinsockLib.Winsock Inet2 
      Left            =   3180
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   375
      Left            =   8790
      TabIndex        =   11
      Top             =   30
      Width           =   585
   End
   Begin VB.TextBox txtLocalPort 
      Height          =   285
      Left            =   3420
      TabIndex        =   8
      Text            =   "23"
      Top             =   60
      Width           =   525
   End
   Begin VB.TextBox txtLocalIP 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   510
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   60
      Width           =   2505
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   285
      Left            =   7500
      TabIndex        =   4
      Text            =   "23"
      Top             =   60
      Width           =   525
   End
   Begin VB.OptionButton opUDP 
      Caption         =   "UDP"
      Enabled         =   0   'False
      Height          =   225
      Left            =   8070
      TabIndex        =   3
      Top             =   0
      Width           =   645
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   285
      Left            =   4590
      TabIndex        =   1
      Top             =   60
      Width           =   2505
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4005
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   450
      Width           =   9435
   End
   Begin MSWinsockLib.Winsock Inet 
      Left            =   1110
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   53
      LocalPort       =   53
   End
   Begin VB.OptionButton opTCP 
      Caption         =   "TCP"
      Height          =   225
      Left            =   8070
      TabIndex        =   2
      Top             =   240
      Value           =   -1  'True
      Width           =   645
   End
   Begin VB.Label Label4 
      Caption         =   "Local:"
      Height          =   225
      Left            =   30
      TabIndex        =   10
      Top             =   90
      Width           =   465
   End
   Begin VB.Label Label3 
      Caption         =   "Port:"
      Height          =   225
      Left            =   3060
      TabIndex        =   9
      Top             =   90
      Width           =   345
   End
   Begin VB.Label Label2 
      Caption         =   "Remote:"
      Height          =   225
      Left            =   3960
      TabIndex        =   6
      Top             =   90
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Port:"
      Height          =   225
      Left            =   7140
      TabIndex        =   5
      Top             =   90
      Width           =   345
   End
End
Attribute VB_Name = "frmIPBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdListen_Click()

Select Case cmdListen.Caption


Case Is = "Listen"
  If opTCP.Value Then
    Inet.Protocol = sckTCPProtocol
    Inet2.Protocol = sckTCPProtocol
    Inet.LocalPort = CInt(txtLocalPort.Text)
    Inet.RemoteHost = txtRemoteIP.Text
    Inet.RemotePort = CInt(txtRemotePort.Text)
    txtLocalPort.Enabled = False
    txtRemoteIP.Enabled = False
    txtRemotePort.Enabled = False
    cmdListen.Caption = "Reset"
    Inet.Close
    Inet.Listen
    log "I>Using TCP to Listin on " & Inet.LocalIP & ":" & Inet.LocalPort
  Else
    Inet.Close
    Inet2.Close
    Inet.Protocol = sckUDPProtocol
    Inet2.Protocol = sckUDPProtocol
    Inet.LocalPort = CInt(txtLocalPort.Text)
    Inet2.RemoteHost = txtRemoteIP.Text
    Inet2.RemotePort = CInt(txtRemotePort.Text)
    txtLocalPort.Enabled = False
    txtRemoteIP.Enabled = False
    txtRemotePort.Enabled = False
    cmdListen.Caption = "Reset"
    Inet.Bind CInt(txtLocalPort.Text)
'    Inet2.Bind CInt(txtRemotePort.Text)
    log "I>Using UDP to Listin on " & Inet.LocalIP & ":" & Inet.LocalPort
  
  End If
  
Case Is = "Reset"
  Inet.Close
  txtLocalPort.Enabled = True
  txtRemoteIP.Enabled = True
  txtRemotePort.Enabled = True
  cmdListen.Caption = "Listen"
  
End Select
  
End Sub

Private Sub Command1_Click()
txtLog.Text = ""
End Sub

Private Sub Form_Load()
txtLocalIP.Text = Inet.LocalIP

End Sub

Private Sub Form_Resize()
If Not Me.WindowState = vbMinimized Then

  txtLog.Width = Me.ScaleWidth
  txtLog.Height = Me.Height - 850

End If
End Sub


Private Sub Inet_Close()
log "I>INET EVENT: CLOSED CONNECTION"
Inet2.Close
cmdListen_Click
cmdListen_Click
End Sub

Private Sub Inet_Connect()
log "I>INET EVENT: CONNECT"
End Sub

Private Sub Inet_ConnectionRequest(ByVal requestID As Long)
log "I>INET EVENT: CONNECTION REQUEST [ " & requestID & " ]"
If Inet.State <> sckClosed Then Inet.Close
log "I>CONNECTING 0 TO " & txtRemoteIP.Text & ":" & CInt(txtRemotePort.Text)
Inet2.Close
Inet2.Connect txtRemoteIP.Text, CInt(txtRemotePort.Text)

Do Until Inet2.State = sckConnected
  DoEvents
Loop

Inet.Accept requestID




End Sub


Private Sub Inet_DataArrival(ByVal bytesTotal As Long)
'On Error GoTo erred
Dim sData As String
Dim bData() As Byte

If opTCP.Value Then
  Inet.PeekData sData, vbString
  Inet.GetData bData(), vbArray + vbByte
  Inet2.SendData bData()

Else
  Inet.GetData sData
  Inet2.SendData sData

End If



log "I>" & sData
Exit Sub

erred:
Inet.Close
Inet2.Close
cmdListen_Click
cmdListen_Click
End Sub

Private Sub Inet_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
log "I>INET ERROR: " & Number & " = " & Description
End Sub


Public Sub log(Text As String)
On Error GoTo erred
txtLog.Text = txtLog.Text & Text & vbCrLf
txtLog.SelStart = Len(txtLog.Text)
Exit Sub
erred:
txtLog.Text = ""
txtLog.Text = txtLog.Text & Text & vbCrLf
txtLog.SelStart = Len(txtLog.Text)
End Sub


Private Sub Inet2_Close()
log "0>INET EVENT: CLOSED CONNECTION"
Inet.Close
cmdListen_Click
cmdListen_Click
End Sub

Private Sub Inet2_DataArrival(ByVal bytesTotal As Long)
On Error GoTo erred
Dim sData As String
Dim bData2() As Byte

If opTCP.Value Then
  Inet2.PeekData sData, vbString
  Inet2.GetData bData2(), vbArray + vbByte
  Inet.SendData bData2()

Else
  Inet2.GetData sData
  Inet.SendData sData

End If

log "O>" & sData
Exit Sub
erred:
Inet.Close
Inet2.Close
cmdListen_Click
cmdListen_Click


End Sub

Private Sub Inet2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
log "O>INET ERROR: " & Number & " = " & Description

End Sub


Private Sub txtLocalPort_Change()
txtRemotePort.Text = txtLocalPort.Text
End Sub


