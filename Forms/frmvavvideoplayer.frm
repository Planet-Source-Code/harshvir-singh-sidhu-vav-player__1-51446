VERSION 5.00
Begin VB.Form frmvavvideoplayer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "VAV Video Player"
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmvavvideoplayer.frx":0000
   ScaleHeight     =   931.338
   ScaleMode       =   0  'User
   ScaleWidth      =   4425.688
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTimeEllapsedSecondOnes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   900
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   17
      Top             =   600
      Width           =   210
   End
   Begin VB.PictureBox picTimeEllapsedSecondTens 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   690
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   600
      Width           =   146
   End
   Begin VB.PictureBox picTimeEllapsedMinuteOnes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   330
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   14
      Top             =   600
      Width           =   210
   End
   Begin VB.PictureBox picTimeEllapsedMinuteTens 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   13
      Top             =   600
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   2880
      Picture         =   "frmvavvideoplayer.frx":05DC
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
      Index           =   1
      Left            =   2880
      Picture         =   "frmvavvideoplayer.frx":0638
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
      Picture         =   "frmvavvideoplayer.frx":068E
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
      Index           =   3
      Left            =   2880
      Picture         =   "frmvavvideoplayer.frx":06E9
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
      Index           =   4
      Left            =   2880
      Picture         =   "frmvavvideoplayer.frx":0744
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   5
      Left            =   2880
      Picture         =   "frmvavvideoplayer.frx":079F
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   6
      Left            =   2880
      Picture         =   "frmvavvideoplayer.frx":07FA
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   7
      Left            =   2880
      Picture         =   "frmvavvideoplayer.frx":0856
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   8
      Left            =   2880
      Picture         =   "frmvavvideoplayer.frx":08AD
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   9
      Left            =   2880
      Picture         =   "frmvavvideoplayer.frx":08F9
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgColon 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2880
      Picture         =   "frmvavvideoplayer.frx":0954
      ScaleHeight     =   195
      ScaleWidth      =   90
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer tmrScrollSongPlayed 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   600
   End
   Begin VB.PictureBox picSongPlayed 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
   Begin VB.Timer tmrPlayState 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   600
   End
   Begin VB.Timer tmrposition 
      Interval        =   100
      Left            =   360
      Top             =   720
   End
   Begin VB.Image imgTitleVideoPlayer 
      Height          =   300
      Left            =   1410
      Picture         =   "frmvavvideoplayer.frx":0993
      Top             =   120
      Width           =   1500
   End
   Begin VB.Image cmdMoveVideoPlaylistLeftMouseOver 
      Height          =   300
      Left            =   3960
      Picture         =   "frmvavvideoplayer.frx":0FA8
      ToolTipText     =   "Hide Video Playlist"
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdPreviousMouseOver 
      Height          =   225
      Left            =   615
      Picture         =   "frmvavvideoplayer.frx":121F
      ToolTipText     =   "Previous"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPreviousActive 
      Height          =   225
      Left            =   615
      Picture         =   "frmvavvideoplayer.frx":1451
      ToolTipText     =   "Previous"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdNextMouseOver 
      Height          =   225
      Left            =   2775
      Picture         =   "frmvavvideoplayer.frx":1619
      ToolTipText     =   "Next"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdNextActive 
      Height          =   225
      Left            =   2775
      Picture         =   "frmvavvideoplayer.frx":184E
      ToolTipText     =   "Next"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdRewindMouseOver 
      Height          =   225
      Left            =   975
      Picture         =   "frmvavvideoplayer.frx":1A15
      ToolTipText     =   "Rewind"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdRewindActive 
      Height          =   225
      Left            =   975
      Picture         =   "frmvavvideoplayer.frx":1C47
      ToolTipText     =   "Rewind"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdForwardMouseOver 
      Height          =   225
      Left            =   2415
      Picture         =   "frmvavvideoplayer.frx":1E0A
      ToolTipText     =   "Forward"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdForwardActive 
      Height          =   225
      Left            =   2415
      Picture         =   "frmvavvideoplayer.frx":203B
      ToolTipText     =   "Forward"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdStopMouseOver 
      Height          =   225
      Left            =   2055
      Picture         =   "frmvavvideoplayer.frx":21FD
      ToolTipText     =   "Stop"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdStopActive 
      Height          =   225
      Left            =   2055
      Picture         =   "frmvavvideoplayer.frx":241D
      ToolTipText     =   "Stop"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPauseMouseOver 
      Height          =   225
      Left            =   1695
      Picture         =   "frmvavvideoplayer.frx":25D8
      ToolTipText     =   "Pause"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPauseActive 
      Height          =   225
      Left            =   1695
      Picture         =   "frmvavvideoplayer.frx":27DC
      ToolTipText     =   "Pause"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPlayMouseOver 
      Height          =   225
      Left            =   1335
      Picture         =   "frmvavvideoplayer.frx":2991
      ToolTipText     =   "Play"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdPlayActive 
      Height          =   225
      Left            =   1335
      Picture         =   "frmvavvideoplayer.frx":2BA7
      ToolTipText     =   "Play"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMoveVideoPlaylistLeft 
      Height          =   300
      Left            =   3960
      Picture         =   "frmvavvideoplayer.frx":2D5B
      ToolTipText     =   "Hide Video Playlist"
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdMuteMouseOver 
      Height          =   225
      Left            =   3120
      Picture         =   "frmvavvideoplayer.frx":2FBA
      ToolTipText     =   "Mute"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdCloseMouseOver 
      Height          =   225
      Left            =   120
      Picture         =   "frmvavvideoplayer.frx":31BC
      ToolTipText     =   "Close Video Player"
      Top             =   120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdclose 
      Height          =   225
      Left            =   120
      Picture         =   "frmvavvideoplayer.frx":33E4
      ToolTipText     =   "Close Video Player"
      Top             =   120
      Width           =   225
   End
   Begin VB.Image cmdCloseMouseDown 
      Height          =   225
      Left            =   120
      Picture         =   "frmvavvideoplayer.frx":361F
      ToolTipText     =   "Close Video Player"
      Top             =   120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdVolumePlusMouseOver 
      Height          =   225
      Left            =   3960
      Picture         =   "frmvavvideoplayer.frx":3847
      ToolTipText     =   "Increase"
      Top             =   120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdVolumeMinusMouseOver 
      Height          =   225
      Left            =   3000
      Picture         =   "frmvavvideoplayer.frx":3A60
      ToolTipText     =   "Decrease"
      Top             =   120
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
      TabIndex        =   1
      Top             =   120
      Width           =   720
   End
   Begin VB.Image cmdFullScreenMouseOver 
      Height          =   225
      Left            =   3480
      Picture         =   "frmvavvideoplayer.frx":3C6A
      ToolTipText     =   "Full Screen"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdFullScreen 
      Height          =   225
      Left            =   3480
      Picture         =   "frmvavvideoplayer.frx":3E8E
      ToolTipText     =   "Full Screen"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdVolumeMinus 
      Height          =   225
      Left            =   3000
      Picture         =   "frmvavvideoplayer.frx":40C8
      ToolTipText     =   "Decrease"
      Top             =   120
      Width           =   225
   End
   Begin VB.Image cmdVolumePlus 
      Height          =   225
      Left            =   3960
      Picture         =   "frmvavvideoplayer.frx":427C
      ToolTipText     =   "Increase"
      Top             =   120
      Width           =   225
   End
   Begin VB.Image cmdMoveVideoPlaylistRightMouseOver 
      Height          =   300
      Left            =   3960
      Picture         =   "frmvavvideoplayer.frx":445B
      ToolTipText     =   "Show Video Playlist"
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdMoveVideoPlaylistRight 
      Height          =   300
      Left            =   3960
      Picture         =   "frmvavvideoplayer.frx":46CF
      ToolTipText     =   "Show Video Playlist"
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdRewind 
      Height          =   225
      Left            =   975
      Picture         =   "frmvavvideoplayer.frx":4928
      ToolTipText     =   "Rewind"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdPlay 
      Height          =   225
      Left            =   1335
      Picture         =   "frmvavvideoplayer.frx":4B66
      ToolTipText     =   "Play"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdPrevious 
      Height          =   225
      Left            =   615
      Picture         =   "frmvavvideoplayer.frx":4D7D
      ToolTipText     =   "Previous"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdPause 
      Height          =   225
      Left            =   1695
      Picture         =   "frmvavvideoplayer.frx":4FB9
      ToolTipText     =   "Pause"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdStop 
      Height          =   225
      Left            =   2055
      Picture         =   "frmvavvideoplayer.frx":51C5
      ToolTipText     =   "Stop"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdForward 
      Height          =   225
      Left            =   2415
      Picture         =   "frmvavvideoplayer.frx":53E8
      ToolTipText     =   "Forward"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdNext 
      Height          =   225
      Left            =   2775
      Picture         =   "frmvavvideoplayer.frx":561A
      ToolTipText     =   "Next"
      Top             =   1320
      Width           =   225
   End
   Begin VB.Image cmdMuteActive 
      Height          =   225
      Left            =   3120
      Picture         =   "frmvavvideoplayer.frx":5853
      ToolTipText     =   "Mute"
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMute 
      Height          =   225
      Left            =   3120
      Picture         =   "frmvavvideoplayer.frx":59DC
      ToolTipText     =   "Mute"
      Top             =   1320
      Width           =   225
   End
End
Attribute VB_Name = "frmvavvideoplayer"
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
    VideoEnable = False
    Unload frmVAVPlayerVideoPlaylist
    Unload frmVAVVideoPlayerScreen
    Unload Me
End Sub

Private Sub cmdFullScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdFullScreenMouseOver.Visible = True
End Sub

Private Sub cmdFullScreen_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        frmVAVVideoPlayerScreen.MainMediaVAVVideo.DisplaySize = mpFullScreen
    End If
End Sub

Private Sub cmdFullScreenMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        frmVAVVideoPlayerScreen.MainMediaVAVVideo.DisplaySize = mpFullScreen
    End If
End Sub

Private Sub cmdMoveVideoPlaylistLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdMoveVideoPlaylistLeftMouseOver.Visible = True
End Sub

Private Sub cmdMoveVideoPlaylistLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        VideoPlayListRight = False
        frmVAVPlayerVideoPlaylist.Hide
        cmdMoveVideoPlaylistRight.Visible = True
        cmdMoveVideoPlaylistLeft.Visible = False
        cmdMoveVideoPlaylistLeftMouseOver.Visible = False
    End If
End Sub

Private Sub cmdMoveVideoPlaylistLeftMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        VideoPlayListRight = False
        frmVAVPlayerVideoPlaylist.Hide
        cmdMoveVideoPlaylistRight.Visible = True
        cmdMoveVideoPlaylistLeft.Visible = False
        cmdMoveVideoPlaylistLeftMouseOver.Visible = False
    End If
End Sub

Private Sub cmdMoveVideoPlaylistRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdMoveVideoPlaylistRightMouseOver.Visible = True
End Sub

Private Sub cmdMoveVideoPlaylistRight_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        VideoPlayListRight = True
        frmVAVPlayerVideoPlaylist.Show
        cmdMoveVideoPlaylistLeft.Visible = True
        cmdMoveVideoPlaylistRight.Visible = False
        cmdMoveVideoPlaylistRightMouseOver.Visible = False
    End If
End Sub

Private Sub cmdMoveVideoPlaylistRightMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        VideoPlayListRight = True
        frmVAVPlayerVideoPlaylist.Show
        cmdMoveVideoPlaylistLeft.Visible = True
        cmdMoveVideoPlaylistRight.Visible = False
        cmdMoveVideoPlaylistRightMouseOver.Visible = False
    End If
End Sub

Private Sub cmdMute_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdMuteMouseOver.Visible = True
End Sub

Private Sub cmdMute_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Mute = True
            cmdMute.Visible = False
            cmdMuteMouseOver.Visible = False
            cmdMuteActive.Visible = True
    End If
End Sub

Private Sub cmdMuteActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Mute = False
            cmdMute.Visible = True
            cmdMuteMouseOver.Visible = False
            cmdMuteActive.Visible = False
    End If
End Sub

Private Sub cmdMuteMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Mute = True
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
        If frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume > -3000 Then
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume = frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume - 100
        Else
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume = -3000
        End If
    End If
End Sub

Private Sub cmdVolumeMinusMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume > -3000 Then
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume = frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume - 100
        Else
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume = -3000
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
        If frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume < -100 Then
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume = frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume + 100
        Else
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume = 0
        End If
    End If
End Sub

Private Sub cmdVolumePlusMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume < -90 Then
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume = frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume + 100
        Else
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume = 0
        End If
    End If
End Sub

Private Sub Form_Click()
    frmMainControls.SetFocus
    If VideoPlayListRight = True Then
        frmVAVPlayerVideoPlaylist.SetFocus
    End If
    If AudioEnable = True Then
        If AudioPlaylistLeft = True Then
            frmVAVPlayerAudioPlaylist.SetFocus
        End If
        frmvavaudioplayer.SetFocus
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

Private Sub imgTitleVideoPlayer_Click()
    FormMouseMove
End Sub

Private Sub imgTitleVideoPlayer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton And (Shift And vbAltMask) And (frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex >= 0) And (frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex < 100) Then
        Load frmRecordSkip
        frmRecordSkip.Show
    End If
End Sub

Private Sub lblVolume_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub lblVolume_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton And (Shift And vbAltMask) > 0 Then
        Load frmAdjustVolumeVideo
        frmAdjustVolumeVideo.SliderMainMediaVideoPlayerVolume.Value = frmVAVVideoPlayerScreen.MainMediaVAVVideo.Volume
        frmAdjustVolumeVideo.Show
    End If
End Sub

Private Sub picSongPlayed_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
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
    If SkipNo <= NoOfSkips Then
        If Int(frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition) = RecordSkip(frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex, SkipNo, 1) Then
            frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition = RecordSkip(frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex, SkipNo, 2)
            SkipNo = SkipNo + 1
        End If
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
    cmdFullScreenMouseOver.Visible = False
    cmdVolumePlusMouseOver.Visible = False
    cmdVolumeMinusMouseOver.Visible = False
    If VideoPlayListRight = True Then
        cmdMoveVideoPlaylistLeft.Visible = True
        cmdMoveVideoPlaylistLeftMouseOver.Visible = False
        cmdMoveVideoPlaylistRight.Visible = False
        cmdMoveVideoPlaylistRightMouseOver.Visible = False
    Else
        cmdMoveVideoPlaylistLeft.Visible = False
        cmdMoveVideoPlaylistLeftMouseOver.Visible = False
        cmdMoveVideoPlaylistRight.Visible = True
        cmdMoveVideoPlaylistRightMouseOver.Visible = False
    End If
    If frmVAVVideoPlayerScreen.MainMediaVAVVideo.Mute = True Then
        cmdMute.Visible = False
        cmdMuteMouseOver.Visible = False
        cmdMuteActive.Visible = True
    Else
        cmdMuteMouseOver.Visible = False
        cmdMuteActive.Visible = False
        cmdMute.Visible = True
    End If
End Sub

Private Sub cmdforward_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdForwardMouseOver.Visible = True
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
    cmdNextMouseOver.Visible = True
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
    cmdPauseMouseOver.Visible = True
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
    cmdPlayMouseOver.Visible = True
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
End Sub

Private Sub cmdRewindMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        RewindAction
    End If
End Sub

Private Sub cmdstop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdStopMouseOver.Visible = True
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
If Len(frmVAVPlayerVideoPlaylist.lstVideoPlayer.Text) > 0 Then
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
    Load frmVAVVideoPlayerScreen
    If paused = False Then
        frmVAVVideoPlayerScreen.MainMediaVAVVideo.FileName = frmVAVPlayerVideoPlaylist.lstVideoPlayer.Text
    End If
    MinuteTime = Int(frmVAVVideoPlayerScreen.MainMediaVAVVideo.Duration / 60)
    SecondTime = Int(frmVAVVideoPlayerScreen.MainMediaVAVVideo.Duration Mod 60)
    StrCapStart = 0
    StrCapLeftOver = 0
    strcap = ""
    strcap = "            *****  " & frmVAVPlayerVideoPlaylist.lstVideoPlayer.Text & "[ " & Str$(Int(MinuteTime / 10)) & Str$(Int(MinuteTime Mod 10)) & ":" & Str$(Int(SecondTime / 10)) & Str$(Int(SecondTime Mod 10)) & " ]      "
    strcap = UCase$(strcap)
    tmrScrollSongPlayed.Enabled = True
    frmVAVVideoPlayerScreen.Show
    frmVAVVideoPlayerScreen.MainMediaVAVVideo.Play
    tmrPlayState.Enabled = True
    stopped = False
    paused = False
Else
    playing = False
End If
Exit Sub
playerror:
StopAction
retvalue = MsgBox("Error Occured During trying to Play file " & frmVAVPlayerVideoPlaylist.lstVideoPlayer.Text & vbCrLf, vbCritical + vbOKOnly, "Invalid Format.")
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
    frmVAVVideoPlayerScreen.MainMediaVAVVideo.Stop
    tmrPlayState.Enabled = False
    tmrScrollSongPlayed.Enabled = False
    picSongPlayed.Cls
    picTimeEllapsedColon.Cls
    picTimeEllapsedMinuteTens.Cls
    picTimeEllapsedMinuteOnes.Cls
    picTimeEllapsedSecondTens.Cls
    picTimeEllapsedSecondOnes.Cls
    frmVAVVideoPlayerScreen.MainMediaVAVVideo.FileName = "c:\msp.dat"
    Unload frmVAVVideoPlayerScreen
    playing = False
    stopped = False
    paused = False
End Sub

Public Sub PauseAction()
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
    frmVAVVideoPlayerScreen.MainMediaVAVVideo.Pause
    playing = False
    stopped = False
    paused = True
End Sub

Public Sub ForwardAction()
On Error GoTo forwarderror
If frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition > frmVAVVideoPlayerScreen.MainMediaVAVVideo.Duration - 5 Then
    frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition = Int(frmVAVVideoPlayerScreen.MainMediaVAVVideo.Duration)
Else
    frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition = frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition + 5
End If
Exit Sub
forwarderror:
If Err.Number = 380 Then
    PauseAction
    frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition = Int(frmVAVVideoPlayerScreen.MainMediaVAVVideo.Duration)
End If
End Sub

Public Sub RewindAction()
On Error GoTo rewinderror
If frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition < 5 Then
    frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition = -1
Else
    frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition = frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition - 5
End If
Exit Sub
rewinderror:
If Err.Number = 380 Then
    PauseAction
    frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition = -1
End If
End Sub

Private Sub PreviousAction()
    If frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListCount > 0 Then
        If frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex > 0 Then
            frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex = frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex - 1
        Else
            frmVAVPlayerVideoPlaylist.lstVideoPlayer.Selected(frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListCount - 1) = True
        End If
        frmVAVVideoPlayerScreen.MainMediaVAVVideo.FileName = frmVAVPlayerVideoPlaylist.lstVideoPlayer.Text
        PlayAction
    Else
        StopAction
    End If
End Sub

Public Sub NextAction()
    If frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListCount > 0 Then
        If frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex < frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListCount - 1 Then
            frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex = frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex + 1
        Else
            frmVAVPlayerVideoPlaylist.lstVideoPlayer.Selected(0) = True
        End If
        frmVAVVideoPlayerScreen.MainMediaVAVVideo.FileName = frmVAVPlayerVideoPlaylist.lstVideoPlayer.Text
        PlayAction
    Else
        StopAction
    End If
End Sub

Public Sub TimeEllapsed()
    picTimeEllapsedColon.PaintPicture imgColon.Picture, 0, 0
    MinuteTime = Int(frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition / 60)
    SecondTime = Int(frmVAVVideoPlayerScreen.MainMediaVAVVideo.CurrentPosition Mod 60)
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
