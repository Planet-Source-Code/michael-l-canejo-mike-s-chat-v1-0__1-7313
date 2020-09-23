Attribute VB_Name = "DataStore"
'Mike's Chat v1.0
'
'Email me at: mike_3d@hotmail.com
'
'or AIM me at: mike3dd
'
'
'
'
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SND_NODEFAULT = &H2
Public Const SND_ASYNC = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const HWND_TOPMOST = -1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Global NickName As String
Global ServerName As String
Public Function TimeOUT(HesitateTime)
Hesitator = Timer
Do While Timer - Hesitator < Val(HesitateTime)
DoEvents
Loop
End Function
Public Function CenterForm(TENProg As Form)
TENProg.Move (Screen.Width) / 2 - (TENProg.Width) / 2, (Screen.Height) / 2 - (TENProg.Height) / 2
End Function
Public Function StayOnTop(TheForm As Form)
Dim SetWinOnTop As Long
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Function
Public Function Playwav(file)
Dim SoundName As String, wFlags As Long, x As Long
On Error Resume Next
SoundName$ = file
wFlags& = SND_ASYNC Or SND_NODEFAULT
x& = sndPlaySound(SoundName$, wFlags&)
End Function
