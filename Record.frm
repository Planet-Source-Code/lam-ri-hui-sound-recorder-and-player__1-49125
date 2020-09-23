VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sound Recorder and Player"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   Icon            =   "Record.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Play Recorded Sound"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************
'Code (C) Lam Ri Hui 2003
'************************
Private Declare Function mciSendString Lib "winmm.dll" _
                                   Alias "mciSendStringA" _
                                   (ByVal lpstrCommand As String, _
                                   ByVal lpstrReturnString As String, _
                                   ByVal uReturnLength As Long, _
                                   ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" _
                                   Alias "mciGetErrorStringA" _
                                   (ByVal dwError As Long, _
                                   ByVal lpstrBuffer As String, _
                                   ByVal uLength As Long) As Long

Sub CloseSound()
    Dim Result&
    Dim errormsg%
    Dim ReturnString As String * 1024
    Dim ErrorString As String * 1024

    Result& = mciSendString("close mysound", ReturnString, 1024, 0)

End Sub
Sub RecordSound()
'Record a sound up to 5 minutes
    Dim Result&
    Dim errormsg%
    Dim ReturnString As String * 1024
    Dim ErrorString As String * 1024

    CloseSound

    Result& = mciSendString("open new type waveaudio alias mysound", ReturnString, 1024, 0)
    If Not Result& = 0 Then
        errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
        MsgBox ErrorString, 0, "Error"
        Exit Sub
    End If

    Result& = mciSendString("set mysound time format ms bitspersample 8 samplespersec 11025", ReturnString, 1024, 0)
    If Not Result& = 0 Then
        errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
        MsgBox ErrorString, 0, "Error"
        Exit Sub
    End If

'Record for 300 seconds (5 minutes)
    Result& = mciSendString("record mysound to 300000", ReturnString, 1024, 0)
    If Not Result& = 0 Then
        errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
        MsgBox ErrorString, 0, "Error"
        Exit Sub
    End If
End Sub

Sub PlayRecSound()
'Play the recorded sound
    Dim Result&
    Dim errormsg%
    Dim ReturnString As String * 1024
    Dim ErrorString As String * 1024
    DoEvents
    Result& = mciSendString("stop mysound", ReturnString, 1024, 0)
    If Not Result& = 0 Then
        errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
        MsgBox ErrorString, 0, "Error"
    End If
    
    Result& = mciSendString("play mysound from 1 wait", ReturnString, 1024, 0)
    If Not Result& = 0 Then
        errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
        MsgBox ErrorString, 0, "Error"
    End If
    
End Sub

Private Sub Command1_Click()
RecordSound
End Sub

Private Sub Command2_Click()
Call PlayRecSound
End Sub

Private Sub Form_Unload(Cancel As Integer)
CloseSound
End Sub



