VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmvavaudioplayer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "VAV Audio Player"
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   Picture         =   "frmvavaudioplayer.frx":0000
   ScaleHeight     =   1725
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox imgColon 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":05DC
      ScaleHeight     =   195
      ScaleWidth      =   90
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   9
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":061B
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   8
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":0676
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   7
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":06C2
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   6
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":0719
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   5
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":0775
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   4
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":07D0
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   3
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":082B
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   2
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":0886
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   1
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":08E1
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   2880
      Picture         =   "frmvavaudioplayer.frx":0937
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSongPlayed 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   480
      ScaleHeight     =   255
      ScaleWidth      =   3735
      TabIndex        =   6
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Timer tmrScrollSongPlayed 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   0
   End
   Begin VB.Timer tmrPlayState 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   0
   End
   Begin VB.Timer tmrposition 
      Interval        =   100
      Left            =   3480
      Top             =   0
   End
   Begin VB.PictureBox picTimeEllapsedMinuteTens 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   1
      Top             =   600
      Width           =   210
   End
   Begin VB.PictureBox picTimeEllapsedMinuteOnes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   330
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   2
      Top             =   600
      Width           =   210
   End
   Begin VB.PictureBox picTimeEllapsedColon 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   540
      ScaleHeight     =   255
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   600
      Width           =   150
   End
   Begin VB.PictureBox picTimeEllapsedSecondTens 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   750
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   4
      Top             =   600
      Width           =   210
   End
   Begin VB.PictureBox picTimeEllapsedSecondOnes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1050
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   5
      Top             =   600
      Width           =   210
   End
   Begin VB.Image imgTitleAudioPlayer 
      Height          =   300
      Left            =   1410
      Picture         =   "frmvavaudioplayer.frx":0993
      Top             =   120
      Width           =   1500
   End
   Begin VB.Image cmdMuteMouseOver 
      Height          =   225
      Left            =   3480
      Picture         =   "frmvavaudioplayer.frx":0FB4
      ToolTipText     =   "Mute"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblVolume 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Volume "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   720
   End
   Begin VB.Image cmdVolumeMinusMouseOver 
      Height          =   225
      Left            =   3000
      Picture         =   "frmvavaudioplayer.frx":11B6
      ToolTipText     =   "Decrease"
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdVolumeMinus 
      Height          =   225
      Left            =   3000
      Picture         =   "frmvavaudioplayer.frx":13C0
      ToolTipText     =   "Decrease"
      Top             =   240
      Width           =   225
   End
   Begin VB.Image cmdVolumePlusMouseOver 
      Height          =   225
      Left            =   3960
      Picture         =   "frmvavaudioplayer.frx":1574
      ToolTipText     =   "Increase"
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdVolumePlus 
      Height          =   225
      Left            =   3960
      Picture         =   "frmvavaudioplayer.frx":178D
      ToolTipText     =   "Increase"
      Top             =   240
      Width           =   225
   End
   Begin VB.Image cmdMute 
      Height          =   225
      Left            =   3480
      Picture         =   "frmvavaudioplayer.frx":196C
      ToolTipText     =   "Mute"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdMoveAudioPlaylistLeftMouseOver 
      Height          =   300
      Left            =   120
      Picture         =   "frmvavaudioplayer.frx":1B82
      ToolTipText     =   "Show Audio Playlist"
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdMoveAudioPlaylistLeft 
      Height          =   300
      Left            =   120
      Picture         =   "frmvavaudioplayer.frx":1DF9
      ToolTipText     =   "Show Audio Playlist"
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdMoveAudioPlaylistRightMouseOver 
      Height          =   300
      Left            =   120
      Picture         =   "frmvavaudioplayer.frx":2058
      ToolTipText     =   "Hide Audio Playlist"
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdMoveAudioPlaylistRight 
      Height          =   300
      Left            =   120
      Picture         =   "frmvavaudioplayer.frx":22CC
      ToolTipText     =   "Hide Audio Playlist"
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin MediaPlayerCtl.MediaPlayer MainMediaVAVAudio 
      Height          =   495
      Left            =   -1000
      TabIndex        =   0
      Top             =   -1000
      Visible         =   0   'False
      Width           =   495
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Image cmdPlayActive 
      Height          =   225
      Left            =   1688
      Picture         =   "frmvavaudioplayer.frx":2525
      ToolTipText     =   "Play"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPlayMouseOver 
      Height          =   225
      Left            =   1688
      Picture         =   "frmvavaudioplayer.frx":26D9
      ToolTipText     =   "Play"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPauseActive 
      Height          =   225
      Left            =   2048
      Picture         =   "frmvavaudioplayer.frx":28EF
      ToolTipText     =   "Pause"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPauseMouseOver 
      Height          =   225
      Left            =   2048
      Picture         =   "frmvavaudioplayer.frx":2AA4
      ToolTipText     =   "Pause"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdStopActive 
      Height          =   225
      Left            =   2408
      Picture         =   "frmvavaudioplayer.frx":2CA8
      ToolTipText     =   "Stop"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdStopMouseOver 
      Height          =   225
      Left            =   2408
      Picture         =   "frmvavaudioplayer.frx":2E63
      ToolTipText     =   "Stop"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdForwardActive 
      Height          =   225
      Left            =   2768
      Picture         =   "frmvavaudioplayer.frx":3083
      ToolTipText     =   "Forward"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdForwardMouseOver 
      Height          =   225
      Left            =   2768
      Picture         =   "frmvavaudioplayer.frx":3245
      ToolTipText     =   "Forward"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdRewindActive 
      Height          =   225
      Left            =   1328
      Picture         =   "frmvavaudioplayer.frx":3476
      ToolTipText     =   "Rewind"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdRewindMouseOver 
      Height          =   225
      Left            =   1328
      Picture         =   "frmvavaudioplayer.frx":3639
      ToolTipText     =   "Rewind"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdNextActive 
      Height          =   225
      Left            =   3128
      Picture         =   "frmvavaudioplayer.frx":386B
      ToolTipText     =   "Next"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdNextMouseOver 
      Height          =   225
      Left            =   3128
      Picture         =   "frmvavaudioplayer.frx":3A32
      ToolTipText     =   "Next"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPreviousActive 
      Height          =   225
      Left            =   968
      Picture         =   "frmvavaudioplayer.frx":3C67
      ToolTipText     =   "Previous"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPreviousMouseOver 
      Height          =   225
      Left            =   968
      Picture         =   "frmvavaudioplayer.frx":3E2F
      ToolTipText     =   "Previous"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdCloseMouseDown 
      Height          =   225
      Left            =   120
      Picture         =   "frmvavaudioplayer.frx":4061
      ToolTipText     =   "Close Audio Player"
      Top             =   120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdclose 
      Height          =   225
      Left            =   120
      Picture         =   "frmvavaudioplayer.frx":4289
      ToolTipText     =   "Close Audio Player"
      Top             =   120
      Width           =   225
   End
   Begin VB.Image cmdCloseMouseOver 
      Height          =   225
      Left            =   120
      Picture         =   "frmvavaudioplayer.frx":44C4
      ToolTipText     =   "Close Audio Player"
      Top             =   120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPlay 
      Height          =   225
      Left            =   1688
      Picture         =   "frmvavaudioplayer.frx":46EC
      ToolTipText     =   "Play"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdRewind 
      Height          =   225
      Left            =   1328
      Picture         =   "frmvavaudioplayer.frx":4903
      ToolTipText     =   "Rewind"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdPrevious 
      Height          =   225
      Left            =   968
      Picture         =   "frmvavaudioplayer.frx":4B41
      ToolTipText     =   "Previous"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdPause 
      Height          =   225
      Left            =   2048
      Picture         =   "frmvavaudioplayer.frx":4D7D
      ToolTipText     =   "Pause"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdStop 
      Height          =   225
      Left            =   2408
      Picture         =   "frmvavaudioplayer.frx":4F89
      ToolTipText     =   "Stop"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdForward 
      Height          =   225
      Left            =   2768
      Picture         =   "frmvavaudioplayer.frx":51AC
      ToolTipText     =   "Forward"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdNext 
      Height          =   225
      Left            =   3128
      Picture         =   "frmvavaudioplayer.frx":53DE
      ToolTipText     =   "Next"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdMuteActive 
      Height          =   225
      Left            =   3480
      Picture         =   "frmvavaudioplayer.frx":5617
      ToolTipText     =   "Mute"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmvavaudioplayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
Dim cursorposition As POINTAPI
Dim retvalue As Long
Public playing As Boolean, paused As Boolean, stopped As Boolean
Dim MinuteTime As Byte, SecondTime As Byte
Dim tempcal As Byte
Dim strcap As String
Dim strcapshow As String
Dim StrCapStart As Integer, StrCapLeftOver As Integer

Private Sub cmdclose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdclose.Visible = False
    cmdCloseMouseDown.Visible = False
    cmdCloseMouseOver.Visible = True
End Sub

Private Sub cmdclose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        ExitAction
    End If
End Sub

Private Sub cmdCloseMouseOver_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdCloseMouseOver.Visible = False
    cmdclose.Visible = False
    cmdCloseMouseDown.Visible = True
End Sub

Private Sub cmdCloseMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        ExitAction
    End If
End Sub

Private Sub cmdCloseMouseDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        ExitAction
    End If
End Sub

Private Sub ExitAction()
    If VideoEnable = True Then
        frmvavvideoplayer.Top = frmMainControls.Top + frmMainControls.Height
        frmVAVPlayerVideoPlaylist.Top = frmMainControls.Top + frmMainControls.Height
        frmVAVPlayerVideoPlaylist.Left = frmvavvideoplayer.Left + frmvavvideoplayer.Width
    End If
    AudioEnable = False
    Unload frmVAVPlayerAudioPlaylist
    Unload Me
End Sub

Private Sub cmdMoveAudioPlaylistLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdMoveAudioPlaylistLeftMouseOver.Visible = True
End Sub

Private Sub cmdMoveAudioPlaylistLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        AudioPlaylistLeft = True
        frmVAVPlayerAudioPlaylist.Show
        cmdMoveAudioPlaylistRight.Visible = True
        cmdMoveAudioPlaylistLeft.Visible = False
        cmdMoveAudioPlaylistLeftMouseOver.Visible = False
    End If
End Sub

Private Sub cmdMoveAudioPlaylistLeftMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        AudioPlaylistLeft = True
        frmVAVPlayerAudioPlaylist.Show
        cmdMoveAudioPlaylistRight.Visible = True
        cmdMoveAudioPlaylistLeft.Visible = False
        cmdMoveAudioPlaylistLeftMouseOver.Visible = False
    End If
End Sub

Private Sub cmdMoveAudioPlaylistRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdMoveAudioPlaylistRightMouseOver.Visible = True
End Sub

Private Sub cmdMoveAudioPlaylistRight_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        AudioPlaylistLeft = False
        frmVAVPlayerAudioPlaylist.Hide
        cmdMoveAudioPlaylistRight.Visible = False
        cmdMoveAudioPlaylistRightMouseOver.Visible = False
        cmdMoveAudioPlaylistLeft.Visible = True
    End If
End Sub

Private Sub cmdMoveAudioPlaylistRightMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        AudioPlaylistLeft = False
        frmVAVPlayerAudioPlaylist.Hide
        cmdMoveAudioPlaylistRight.Visible = False
        cmdMoveAudioPlaylistRightMouseOver.Visible = False
        cmdMoveAudioPlaylistLeft.Visible = True
    End If
End Sub

Private Sub cmdMute_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdMuteMouseOver.Visible = True
End Sub

Private Sub cmdMute_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        MainMediaVAVAudio.Mute = True
        cmdMute.Visible = False
        cmdMuteMouseOver.Visible = False
        cmdMuteActive.Visible = True
    End If
End Sub

Private Sub cmdMuteActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        MainMediaVAVAudio.Mute = False
        cmdMute.Visible = True
        cmdMuteMouseOver.Visible = False
        cmdMuteActive.Visible = False
    End If
End Sub

Private Sub cmdMuteMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        MainMediaVAVAudio.Mute = True
        cmdMute.Visible = False
        cmdMuteMouseOver.Visible = False
        cmdMuteActive.Visible = True
    End If
End Sub

Private Sub cmdVolumeMinus_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdVolumeMinusMouseOver.Visible = True
End Sub

Private Sub cmdVolumeMinus_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If MainMediaVAVAudio.Volume > -3000 Then
            MainMediaVAVAudio.Volume = MainMediaVAVAudio.Volume - 100
        Else
            MainMediaVAVAudio.Volume = -3000
        End If
    End If
End Sub

Private Sub cmdVolumeMinusMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If MainMediaVAVAudio.Volume > -3000 Then
            MainMediaVAVAudio.Volume = MainMediaVAVAudio.Volume - 100
        Else
            MainMediaVAVAudio.Volume = -3000
        End If
    End If
End Sub

Private Sub cmdVolumePlus_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdVolumePlusMouseOver.Visible = True
End Sub

Private Sub cmdVolumePlus_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If MainMediaVAVAudio.Volume < -100 Then
            MainMediaVAVAudio.Volume = MainMediaVAVAudio.Volume + 100
        Else
            MainMediaVAVAudio.Volume = 0
        End If
    End If
End Sub

Private Sub cmdVolumePlusMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If MainMediaVAVAudio.Volume < -90 Then
            MainMediaVAVAudio.Volume = MainMediaVAVAudio.Volume + 100
        Else
            MainMediaVAVAudio.Volume = 0
        End If
    End If
End Sub

Private Sub Form_Click()
    frmMainControls.SetFocus
    If AudioPlaylistLeft = True And frmMainControls.minimized = False Then
        frmVAVPlayerAudioPlaylist.SetFocus
    End If
    If VideoEnable = True And frmMainControls.minimized = False Then
        If VideoPlayListRight = True And frmMainControls.minimized = False Then
            frmVAVPlayerVideoPlaylist.SetFocus
        End If
        frmvavvideoplayer.SetFocus
    End If
End Sub

Private Sub Form_GotFocus()
    Form_Click
End Sub

Private Sub Form_Load()
    picTimeEllapsedMinuteOnes.Left = picTimeEllapsedMinuteTens.Left + picTimeEllapsedMinuteTens.Width
    picTimeEllapsedColon.Left = picTimeEllapsedMinuteOnes.Left + picTimeEllapsedMinuteOnes.Width
    picTimeEllapsedSecondTens.Left = picTimeEllapsedColon.Left + picTimeEllapsedColon.Width
    picTimeEllapsedSecondOnes.Left = picTimeEllapsedSecondTens.Left + picTimeEllapsedSecondTens.Width
    StopAction
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub imgTitleAudioPlayer_Click()
    FormMouseMove
End Sub

Private Sub lblVolume_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub lblVolume_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton And (Shift And vbAltMask) > 0 Then
        Load frmAdjustVolumeAudio
        frmAdjustVolumeAudio.SliderMainMediaAudioPlayerVolume.Value = MainMediaVAVAudio.Volume
        frmAdjustVolumeAudio.Show
    End If
End Sub

Private Sub picSongPlayed_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub MainMediaVAVAudio_EndOfStream(ByVal Result As Long)
    NextAction
End Sub

Private Sub MainMediaVAVAudio_PositionChange(ByVal oldPosition As Double, ByVal newPosition As Double)
    TimeEllapsed
End Sub

Private Sub picTimeEllapsedColon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub picTimeEllapsedMinuteOnes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub picTimeEllapsedMinuteTens_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub picTimeEllapsedSecondOnes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub picTimeEllapsedSecondTens_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub tmrPlayState_Timer()
    TimeEllapsed
End Sub

Private Sub tmrposition_Timer()
    retvalue = GetCursorPos(cursorposition)
    If cursorposition.x < ScaleX(Me.Left, vbTwips, vbPixels) Or cursorposition.x > ScaleX(Me.Left + Me.Width, vbTwips, vbPixels) Or cursorposition.y < ScaleY(Me.Top, vbTwips, vbPixels) Or cursorposition.y > ScaleY(Me.Top + Me.Height, vbTwips, vbPixels) Then
        FormMouseMove
    End If
End Sub

Private Sub FormMouseMove()
    cmdclose.Visible = True
    cmdCloseMouseOver.Visible = False
    cmdCloseMouseDown.Visible = False
    cmdPlayMouseOver.Visible = False
    cmdPauseMouseOver.Visible = False
    cmdStopMouseOver.Visible = False
    cmdRewindMouseOver.Visible = False
    cmdForwardMouseOver.Visible = False
    cmdNextMouseOver.Visible = False
    cmdPreviousMouseOver.Visible = False
    cmdVolumePlusMouseOver.Visible = False
    cmdVolumeMinusMouseOver.Visible = False
    If AudioPlaylistLeft = True Then
        cmdMoveAudioPlaylistLeft.Visible = False
        cmdMoveAudioPlaylistLeftMouseOver.Visible = False
        cmdMoveAudioPlaylistRight.Visible = True
        cmdMoveAudioPlaylistRightMouseOver.Visible = False
    Else
        cmdMoveAudioPlaylistLeft.Visible = True
        cmdMoveAudioPlaylistLeftMouseOver.Visible = False
        cmdMoveAudioPlaylistRight.Visible = False
        cmdMoveAudioPlaylistRightMouseOver.Visible = False
    End If
    If MainMediaVAVAudio.Mute = True Then
        cmdMute.Visible = False
        cmdMuteMouseOver.Visible = False
    Else
        cmdMuteMouseOver.Visible = False
        cmdMuteActive.Visible = True
    End If
End Sub

Private Sub cmdforward_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdRewindMouseOver.Visible = False
    cmdForwardMouseOver.Visible = True
    cmdPlayMouseOver.Visible = False
    cmdStopMouseOver.Visible = False
    cmdPauseMouseOver.Visible = False
    cmdNextMouseOver.Visible = False
    cmdPreviousMouseOver.Visible = False
End Sub

Private Sub cmdforward_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        ForwardAction
    End If
End Sub

Private Sub cmdForwardMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        ForwardAction
    End If
End Sub

Private Sub cmdnext_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdPreviousMouseOver.Visible = False
    cmdNextMouseOver.Visible = True
    cmdPlayMouseOver.Visible = False
    cmdForwardMouseOver.Visible = False
    cmdPauseMouseOver.Visible = False
    cmdStopMouseOver.Visible = False
    cmdRewindMouseOver.Visible = False
End Sub

Private Sub cmdnext_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        NextAction
    End If
End Sub

Private Sub cmdNextMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        NextAction
    End If
End Sub

Private Sub cmdpause_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdRewindMouseOver.Visible = False
    cmdForwardMouseOver.Visible = False
    cmdPauseMouseOver.Visible = True
    cmdPlayMouseOver.Visible = False
    cmdStopMouseOver.Visible = False
    cmdNextMouseOver.Visible = False
    cmdPreviousMouseOver.Visible = False
End Sub

Private Sub cmdpause_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If paused = False Then
            PauseAction
            paused = True
        End If
    End If
End Sub

Private Sub cmdPauseMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If paused = False Then
            PauseAction
            paused = True
        End If
    End If
End Sub

Private Sub cmdplay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdRewindMouseOver.Visible = False
    cmdForwardMouseOver.Visible = False
    cmdPlayMouseOver.Visible = True
    cmdPauseMouseOver.Visible = False
    cmdStopMouseOver.Visible = False
    cmdNextMouseOver.Visible = False
    cmdPreviousMouseOver.Visible = False
End Sub

Private Sub cmdplay_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If playing = False Then
            playing = True
            PlayAction
        End If
    End If
End Sub

Private Sub cmdPlayMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If playing = False Then
            playing = True
            PlayAction
        End If
    End If
End Sub

Private Sub cmdprevious_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdPreviousMouseOver.Visible = True
    cmdNextMouseOver.Visible = False
    cmdPlayMouseOver.Visible = False
    cmdForwardMouseOver.Visible = False
    cmdPauseMouseOver.Visible = False
    cmdStopMouseOver.Visible = False
    cmdRewindMouseOver.Visible = False
End Sub


Private Sub cmdprevious_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        PreviousAction
    End If
End Sub

Private Sub cmdPreviousMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        PreviousAction
    End If
End Sub

Private Sub cmdrewind_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        RewindAction
    End If
End Sub

Private Sub cmdrewind_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdRewindMouseOver.Visible = True
    cmdForwardMouseOver.Visible = False
    cmdPlayMouseOver.Visible = False
    cmdStopMouseOver.Visible = False
    cmdPauseMouseOver.Visible = False
    cmdNextMouseOver.Visible = False
    cmdPreviousMouseOver.Visible = False
End Sub

Private Sub cmdRewindMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        RewindAction
    End If
End Sub

Private Sub cmdstop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdRewindMouseOver.Visible = False
    cmdForwardMouseOver.Visible = False
    cmdStopMouseOver.Visible = True
    cmdPauseMouseOver.Visible = False
    cmdPlayMouseOver.Visible = False
    cmdNextMouseOver.Visible = False
    cmdPreviousMouseOver.Visible = False
End Sub

Private Sub cmdstop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If stopped = False Then
            StopAction
            stopped = True
        End If
    End If
End Sub

Private Sub cmdStopMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If stopped = False Then
            StopAction
            stopped = True
        End If
    End If
End Sub

Public Sub PlayAction()
On Error GoTo playerror
If Len(frmVAVPlayerAudioPlaylist.lstAudioPlayer.Text) > 0 Then
    cmdPlay.Visible = False
    cmdPlayActive.Visible = True
    cmdPlayMouseOver.Visible = False
    cmdPause.Visible = True
    cmdPauseActive.Visible = False
    cmdPauseMouseOver.Visible = False
    cmdStop.Visible = True
    cmdStopActive.Visible = False
    cmdStopActive.Visible = False
    cmdRewind.Visible = True
    cmdRewindActive.Visible = False
    cmdRewindMouseOver.Visible = False
    cmdForward.Visible = True
    cmdForwardActive.Visible = False
    cmdForwardMouseOver.Visible = False
    cmdNext.Visible = True
    cmdNextActive.Visible = False
    cmdNextMouseOver.Visible = False
    cmdPrevious.Visible = True
    cmdPreviousActive.Visible = False
    cmdPreviousMouseOver.Visible = False
    If paused = False Then
        MainMediaVAVAudio.FileName = frmVAVPlayerAudioPlaylist.lstAudioPlayer.Text
    End If
    MinuteTime = Int(MainMediaVAVAudio.Duration / 60)
    SecondTime = Int(MainMediaVAVAudio.Duration Mod 60)
    StrCapStart = 0
    StrCapLeftOver = 0
    strcap = ""
    strcap = "            *****  " & frmVAVPlayerAudioPlaylist.lstAudioPlayer.Text & "[ " & Str$(Int(MinuteTime / 10)) & Str$(Int(MinuteTime Mod 10)) & ":" & Str$(Int(SecondTime / 10)) & Str$(Int(SecondTime Mod 10)) & " ]      "
    strcap = UCase$(strcap)
    tmrScrollSongPlayed.Enabled = True
    MainMediaVAVAudio.Play
    TimeEllapsed
    tmrPlayState.Enabled = True
    stopped = False
    paused = False
Else
    playing = False
End If
Exit Sub
playerror:
StopAction
retvalue = MsgBox("Error Occured During trying to Play file " & frmVAVPlayerAudioPlaylist.lstAudioPlayer.Text & vbCrLf, vbCritical + vbOKOnly, "Invalid Format.")
End Sub

Public Sub StopAction()
    cmdPlay.Visible = False
    cmdPlayActive.Visible = True
    cmdPlayMouseOver.Visible = False
    cmdPause.Visible = False
    cmdPauseActive.Visible = True
    cmdPauseMouseOver.Visible = False
    cmdStop.Visible = False
    cmdStopActive.Visible = True
    cmdStopMouseOver.Visible = False
    cmdForward.Visible = False
    cmdForwardActive.Visible = True
    cmdForwardMouseOver.Visible = False
    cmdRewind.Visible = False
    cmdRewindActive.Visible = True
    cmdRewindMouseOver.Visible = False
    cmdNext.Visible = False
    cmdNextActive.Visible = True
    cmdNextMouseOver.Visible = False
    cmdPrevious.Visible = False
    cmdPreviousActive.Visible = True
    cmdPreviousMouseOver.Visible = False
    MainMediaVAVAudio.Stop
    picTimeEllapsedColon.Cls
    picTimeEllapsedMinuteTens.Cls
    picTimeEllapsedMinuteOnes.Cls
    picTimeEllapsedSecondTens.Cls
    picTimeEllapsedSecondOnes.Cls
    tmrPlayState.Enabled = False
    tmrScrollSongPlayed.Enabled = False
    picSongPlayed.Cls
    strcap = ""
    MainMediaVAVAudio.FileName = "c:\msp.dat"
    playing = False
    stopped = False
    paused = False
End Sub

Private Sub PauseAction()
    cmdPlay.Visible = True
    cmdPlayActive.Visible = False
    cmdPlayMouseOver.Visible = False
    cmdPause.Visible = False
    cmdPauseActive.Visible = True
    cmdPauseMouseOver.Visible = False
    cmdStop.Visible = True
    cmdStopActive.Visible = False
    cmdStopMouseOver.Visible = False
    cmdNext.Visible = True
    cmdNextActive.Visible = False
    cmdNextMouseOver.Visible = False
    cmdPrevious.Visible = True
    cmdPreviousActive.Visible = False
    cmdPreviousMouseOver.Visible = False
    MainMediaVAVAudio.Pause
    playing = False
    stopped = False
    paused = True
End Sub

Private Sub ForwardAction()
On Error GoTo forwarderror
If MainMediaVAVAudio.CurrentPosition > MainMediaVAVAudio.Duration - 5 Then
    MainMediaVAVAudio.CurrentPosition = Int(MainMediaVAVAudio.Duration)
Else
    MainMediaVAVAudio.CurrentPosition = MainMediaVAVAudio.CurrentPosition + 5
End If
Exit Sub
forwarderror:
If Err.Number = 380 Then
    PauseAction
    MainMediaVAVAudio.CurrentPosition = Int(MainMediaVAVAudio.Duration)
End If
End Sub

Private Sub RewindAction()
On Error GoTo rewinderror
If MainMediaVAVAudio.CurrentPosition < 5 Then
    MainMediaVAVAudio.CurrentPosition = -1
Else
    MainMediaVAVAudio.CurrentPosition = MainMediaVAVAudio.CurrentPosition - 5
End If
Exit Sub
rewinderror:
If Err.Number = 380 Then
    PauseAction
    MainMediaVAVAudio.CurrentPosition = -1
End If
End Sub

Private Sub PreviousAction()
    If frmVAVPlayerAudioPlaylist.lstAudioPlayer.ListCount > 0 Then
        If frmVAVPlayerAudioPlaylist.lstAudioPlayer.ListIndex > 0 Then
            frmVAVPlayerAudioPlaylist.lstAudioPlayer.ListIndex = frmVAVPlayerAudioPlaylist.lstAudioPlayer.ListIndex - 1
        Else
            frmVAVPlayerAudioPlaylist.lstAudioPlayer.Selected(frmVAVPlayerAudioPlaylist.lstAudioPlayer.ListCount - 1) = True
        End If
        MainMediaVAVAudio.FileName = frmVAVPlayerAudioPlaylist.lstAudioPlayer.Text
        PlayAction
    Else
        StopAction
    End If
End Sub

Private Sub NextAction()
    If frmVAVPlayerAudioPlaylist.lstAudioPlayer.ListCount > 0 Then
        If frmVAVPlayerAudioPlaylist.lstAudioPlayer.ListIndex < frmVAVPlayerAudioPlaylist.lstAudioPlayer.ListCount - 1 Then
            frmVAVPlayerAudioPlaylist.lstAudioPlayer.ListIndex = frmVAVPlayerAudioPlaylist.lstAudioPlayer.ListIndex + 1
        Else
            frmVAVPlayerAudioPlaylist.lstAudioPlayer.Selected(0) = True
        End If
        MainMediaVAVAudio.FileName = frmVAVPlayerAudioPlaylist.lstAudioPlayer.Text
        PlayAction
    Else
        StopAction
    End If
End Sub

Private Sub TimeEllapsed()
    picTimeEllapsedColon.PaintPicture imgColon.Picture, 0, 0
    MinuteTime = Int(MainMediaVAVAudio.CurrentPosition / 60)
    SecondTime = Int(MainMediaVAVAudio.CurrentPosition Mod 60)
    tempcal = Int(MinuteTime / 10)
    picTimeEllapsedMinuteTens.PaintPicture imgDigit(tempcal).Picture, 0, 0
    tempcal = Int(MinuteTime Mod 10)
    picTimeEllapsedMinuteOnes.PaintPicture imgDigit(tempcal).Picture, 0, 0
    tempcal = Int(SecondTime / 10)
    picTimeEllapsedSecondTens.PaintPicture imgDigit(tempcal).Picture, 0, 0
    tempcal = Int(SecondTime Mod 10)
    picTimeEllapsedSecondOnes.PaintPicture imgDigit(tempcal).Picture, 0, 0
End Sub

Private Sub tmrScrollSongPlayed_Timer()
    strcapshow = ""
    StrCapStart = StrCapStart + 1
    If Len(strcap) - StrCapStart >= 40 Then
        strcapshow = Mid(strcap, StrCapStart, 40)
    Else
        StrCapLeftOver = 40 - (Len(strcap) - StrCapStart)
        strcapshow = Mid(strcap, StrCapStart, Len(strcap) - StrCapStart)
        strcapshow = strcapshow & Mid(strcap, 1, StrCapLeftOver)
        If StrCapStart = Len(strcap) Then
            StrCapStart = 0
        End If
    End If
    picSongPlayed.Cls
    picSongPlayed.Print strcapshow
End Sub
