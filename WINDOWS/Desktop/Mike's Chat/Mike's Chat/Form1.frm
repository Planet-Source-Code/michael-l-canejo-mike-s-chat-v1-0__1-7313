VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   2175
      Left            =   2760
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1455
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox RoomList 
      Height          =   6300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x As Integer, z As Integer
'RoomList.Clear
RoomList.AddItem "<START>"
For x = 1 To 100
RoomList.AddItem "Number" & x
Next x
RoomList.AddItem "*************************"
For x = 1 To 100
RoomList.AddItem "Number" & x
Next x
RoomList.AddItem "<END>"

End Sub

Private Sub Command2_Click()
RemoveDoubles RoomList
End Sub

Private Sub Command3_Click()
Dim x As Integer
For x = 1 To 10
Command1_Click
Command2_Click
Next x
End Sub

Private Sub Form_Load()
xer = 0
End Sub
