VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Clientfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mike's Chat - [Client]"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer SendClientInfo 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   2520
   End
   Begin VB.TextBox ChatSendBox 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox ChatBox 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"Client.frx":0000
   End
   Begin MSWinsockLib.Winsock ClientWinsock 
      Left            =   3600
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton SendText 
      Caption         =   "Send Text"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2890
      Width           =   1095
   End
End
Attribute VB_Name = "Clientfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ClientWinsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
Dim Data2 As String
ClientWinsock.GetData Data, vbString
Data2 = Left(Data, 1)
Data = Mid(Data, 2)
If Data2 = "C" Then
Select Case (Data)
Case "Connection Requested"
ServerORClient.StatusLabel1 = "Connection Granted..."
ClientWinsock.SendData "C" & "Connection Accepted"
Me.Show
Case "Disconnected"
MsgBox "You were disconnected from the server.", vbSystemModal + vbExclamation, "Notice"
ClientWinsock.Close
Shell App.Path & "\" & App.EXEName, vbNormalFocus
End
Case "Ejection"
MsgBox "You were Booted from the server for bad conduct.", vbSystemModal + vbCritical, "Server Admin:"
ClientWinsock.Close
Shell App.Path & "\" & App.EXEName, vbNormalFocus
Unload Me
End Select
ElseIf Data2 = "T" Then
On Error Resume Next
ChatBox.SelStart = Len(ChatBox.Text)
ChatBox.SelColor = vbRed
GetNick$ = Mid(Data, InStr(Data, ":"), Len(Data))
GetNick2$ = Mid(Data, InStr(Data, ":") + 1, Len(Data))
GetNick3$ = Replace(Data, GetNick$, "")
ChatBox.SelBold = True
ChatBox.SelText = GetNick3$ & ":"
ChatBox.SelBold = False
ChatBox.SelText = GetNick2$ & vbCrLf
Playwav App.Path & "\Sounds\Chatrcv.wav"
End If

End Sub


Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ServerORClient.Visible = False And Me.Visible = True And Serverfrm.Visible = False Then
ClientWinsock.SendData "D" & NickName & " - " & ClientWinsock.LocalIP
ClientWinsock.Close: End: End If
End Sub

Private Sub SendClientInfo_Timer()
SendClientInfo.Enabled = False
ClientWinsock.SendData "N" & NickName & " - " & ClientWinsock.LocalIP
ChatSendBox.SetFocus
SendClientInfo.Enabled = False
End Sub

Private Sub SendText_Click()
If ChatSendBox = "" Then Exit Sub
ClientWinsock.SendData "T" & NickName & ":  " & ChatSendBox
ChatBox.SelStart = Len(ChatBox.Text)
ChatBox.SelColor = vbBlue
ChatBox.SelBold = True
ChatBox.SelText = NickName & ":   "
ChatBox.SelBold = False
ChatBox.SelText = ChatSendBox & vbCrLf
ChatSendBox = ""
Playwav App.Path & "\Sounds\Chatsnd.wav"
End Sub
