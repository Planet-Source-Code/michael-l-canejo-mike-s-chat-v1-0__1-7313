VERSION 5.00
Begin VB.Form ServerORClient 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mike's Chatroom - [Menu]"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ClientIp 
      Height          =   285
      Left            =   120
      MaxLength       =   15
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton ConnectMaker 
      Caption         =   "Connect as?"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.OptionButton ClientOpt 
      Alignment       =   1  'Right Justify
      Caption         =   "         Client"
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton ServerOpt 
      Caption         =   "Server"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox UserName 
      Height          =   285
      Left            =   120
      MaxLength       =   12
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label StatusLabel1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "ServerORClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const vbDataErrDisplay = 1
Private Sub ClientIp_Change()
Dim a As Integer, b As Integer
a = Len(UserName)
b = Len(ClientIp)
If ClientIp = "" Then ConnectMaker.Enabled = False
If a > 0 And b > 0 Then ConnectMaker.Enabled = True
End Sub

Private Sub ClientOpt_Click()
Clientfrm.Left = 50000
Clientfrm.Show
Clientfrm.Hide
CenterForm Clientfrm
Unload Serverfrm
ConnectMaker.SetFocus
End Sub

Private Sub ConnectMaker_Click()
On Error Resume Next
If ConnectMaker.Tag = 3 Then GoTo SkipStartUp
If ConnectMaker.Tag = 4 Then GoTo SkipStartUp
If ServerOpt.Value = False And ClientOpt.Value = False Then MsgBox "Please choose whether to connect as the Server Admin or Client user.", vbSystemModal + vbExclamation, "Choose One!": Exit Sub
If ConnectMaker.Tag = 1 Then If ConnectMaker.Caption = "Connect as?" Then ConnectMaker.Enabled = False: If ServerOpt.Value = True Then ConnectMaker.Caption = "Connect as Server": UserName = "Server" Else If ClientOpt.Value = True Then ConnectMaker.Caption = "Connect as Client": ConnectMaker.Tag = 2
If ClientOpt.Value = True Then Me.Caption = "Mike's Chatroom - [Client]"
If ServerOpt.Value = True Then Me.Caption = "Mike's Chatroom - [Server]"
If ClientOpt.Value = True Then: ServerOpt.Visible = False: ClientOpt.Visible = False: StatusLabel1.Top = 0: StatusLabel1.Left = 120: UserName.Visible = True: UserName.Top = 240: UserName = "Guest": UserName.Left = 120: UserName.SetFocus
If ServerOpt.Value = True Then: ServerOpt.Visible = False: ClientOpt.Visible = False: StatusLabel1.Top = 0: StatusLabel1.Left = 120: UserName.Visible = True: UserName.Top = 240: UserName = "Server": UserName.Left = 120: ClientIp.Top = 600: ClientIp.Left = 120: ClientIp = "127.0.0.1": ClientIp.Visible = True: UserName.SetFocus
ConnectMaker.Tag = 2
If ConnectMaker.Tag = 2 Then ConnectMaker.Tag = 3: Exit Sub
MsgBox ConnectMaker.Tag
SkipStartUp:
If ConnectMaker.Caption = "Connect as Server" Then
If ConnectMaker.Tag = 4 Then GoTo SkipServerConnect
With Serverfrm.ServerWinsock
.Protocol = sckUDPProtocol
  .RemotePort = 11110
    .Bind
End With
SkipServerConnect:
ConnectMaker.Tag = 4
StatusLabel1.Caption = "Allowing All Connections..."
NickName = UserName
Serverfrm.ServerWinsock.RemoteHost = ClientIp
Serverfrm.ServerWinsock.SendData "C" & "Connection Requested"
Exit Sub
End If
If ConnectMaker.Caption = "Connect as Client" Then
With Clientfrm.ClientWinsock
.Protocol = sckUDPProtocol
  .LocalPort = 11110
    .Bind
End With
NickName = UserName
StatusLabel1.Caption = "Connecting to Server..."
End If
End Sub

Private Sub Form_Load()
StayOnTop Me
Me.Visible = False
Me.Height = 1740
Me.Width = 2250
ConnectMaker.Tag = 1
CenterForm Me
Me.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub ServerOpt_Click()
ServerPassword$ = InputBox("You need to optain a server password for this feature.", "Server Admin", "")
If Not ServerPassword$ Like "serveradmin" Then ServerOpt.Value = False: ClientOpt.Value = False
Serverfrm.Left = 50000
Serverfrm.Show
Serverfrm.Hide
CenterForm Serverfrm
Unload Clientfrm
ConnectMaker.SetFocus
End Sub

Private Sub StatusLabel1_Change()
If StatusLabel1 = "Connection Granted..." Then Clientfrm.Visible = True: ServerORClient.Hide: Clientfrm.SendClientInfo.Enabled = True
End Sub

Private Sub UserName_Change()
Dim a As Integer, b As Integer
UserName = Replace(UserName, " ", ""): UserName = Replace(UserName, "`", ""): UserName = Replace(UserName, "~", ""): UserName = Replace(UserName, "!", ""): UserName = Replace(UserName, "@", ""): UserName = Replace(UserName, "#", ""): UserName = Replace(UserName, "$", ""): UserName = Replace(UserName, "%", ""): UserName = Replace(UserName, "^", ""): UserName = Replace(UserName, "&", ""): UserName = Replace(UserName, "*", ""): UserName = Replace(UserName, "(", ""): UserName = Replace(UserName, ")", ""): UserName = Replace(UserName, "_", ""): UserName = Replace(UserName, "-", ""): UserName = Replace(UserName, "+", ""): UserName = Replace(UserName, "=", ""): UserName = Replace(UserName, "    ", ""): UserName = Replace(UserName, "[", ""): UserName = Replace(UserName, "]", ""): UserName = Replace(UserName, "{", ""): UserName = Replace(UserName, "}", ""): UserName = Replace(UserName, "|", ""): UserName = Replace(UserName, "\", ""): UserName = Replace(UserName, ":", ""): UserName = Replace(UserName, ";", "") _
: UserName = Replace(UserName, "'", ""): UserName = Replace(UserName, """", ""): UserName = Replace(UserName, ",", ""): UserName = Replace(UserName, "<", ""): UserName = Replace(UserName, ">", ""): UserName = Replace(UserName, ".", ""): UserName = Replace(UserName, "?", ""): UserName = Replace(UserName, "/", ""): UserName = Replace(UserName, "§", ""): UserName = Replace(UserName, "æ", "")
a = Len(UserName)
b = Len(ClientIp)
UserName.SelStart = a
If UserName = "" Then ConnectMaker.Enabled = False
If ClientIp.Visible = True Then If a > 0 And b > 0 Then ConnectMaker.Enabled = True
If ClientIp.Visible = False Then If a > 0 Then ConnectMaker.Enabled = True

End Sub

