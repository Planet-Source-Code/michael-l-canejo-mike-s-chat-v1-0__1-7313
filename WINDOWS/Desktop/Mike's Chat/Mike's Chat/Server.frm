VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Serverfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mike's Chat - [Server]"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton IPBooter 
      Caption         =   "Boot Ip"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ListBox RoomList 
      Height          =   1815
      ItemData        =   "Server.frx":0000
      Left            =   4080
      List            =   "Server.frx":0002
      TabIndex        =   3
      Top             =   480
      Width           =   2175
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
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   2890
      Width           =   4935
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
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"Server.frx":0004
   End
   Begin MSWinsockLib.Winsock ServerWinsock 
      Left            =   4080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton SendText 
      Caption         =   "Send Text"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2890
      Width           =   1095
   End
   Begin VB.Label CountUsers 
      Alignment       =   2  'Center
      Caption         =   "In Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Liner1 
         Caption         =   "-"
      End
      Begin VB.Menu BootIP 
         Caption         =   "Boot"
      End
      Begin VB.Menu Liner2 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "Serverfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BootIP_Click()
ServerWinsock.SendData "C" & "Ejection"
End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ServerORClient.Visible = False And Me.Visible = True And Clientfrm.Visible = False Then ServerWinsock.SendData "C" & "Disconnected": ServerWinsock.Close: End
End Sub


Private Sub IPBooter_Click()
On Error GoTo here:
If RoomList.Text = "" Then MsgBox "Please select a User to boot then click Boot Ip.", vbSystemModal + vbCritical, "Error:": Exit Sub
GetUserIP$ = Mid(RoomList.Text, InStr(RoomList.Text, "-") + 2)
GetUser$ = Replace(RoomList.Text, GetUserIP$, "")
GetUser$ = Replace(GetUser$, "-", ""): GetUser$ = Replace(GetUser$, " ", "")
BootIP.Caption = "Boot [" & GetUser$ & "]" & " - [" & GetUserIP$ & "]"
here:
Me.PopupMenu Menu1, 1
End Sub

Private Sub RoomList_DblClick()
On Error GoTo here:
GetUserIP$ = Mid(RoomList.Text, InStr(RoomList.Text, "-") + 2)
GetUser$ = Replace(RoomList.Text, GetUserIP$, "")
GetUser$ = Replace(GetUser$, "-", ""): GetUser$ = Replace(GetUser$, " ", "")
BootIP.Caption = "Boot [" & GetUser$ & "]" & " - [" & GetUserIP$ & "]"
here:
Me.PopupMenu Menu1, 1
End Sub

Private Sub SendText_Click()
If ChatSendBox = "" Then Exit Sub
ServerWinsock.SendData "T" & NickName & ":  " & ChatSendBox
ChatBox.SelStart = Len(ChatBox.Text)
ChatBox.SelColor = vbBlue
ChatBox.SelBold = True
ChatBox.SelText = NickName & ":   "
ChatBox.SelBold = False
ChatBox.SelText = ChatSendBox & vbCrLf
ChatSendBox = ""
Playwav App.Path & "\Sounds\Chatsnd.wav"
End Sub

Private Sub ServerWinsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
Dim Data2 As String
ServerWinsock.GetData Data, vbString
Data2 = Left(Data, 1)
Data = Mid(Data, 2)
If Data2 = "N" Then
RoomList.AddItem Data
CountUsers.Caption = RoomList.ListCount & " In Room"
End If
If Data2 = "C" Then
Select Case (Data)
Case "Connection Accepted"
Serverfrm.Show
ServerORClient.Hide
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
ElseIf Data2 = "D" Then
Dim x As Integer, Counter
RoomList.ListIndex = 0
Counter = 0
For x = 1 To RoomList.ListCount
RoomList.ListIndex = Counter
If RoomList.Text = Data Then RoomList.RemoveItem Counter: CountUsers.Caption = RoomList.ListCount & " In Room": Exit Sub
Counter = Counter + 1
Next x
End If
End Sub

