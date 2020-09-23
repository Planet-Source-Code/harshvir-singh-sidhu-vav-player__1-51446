VERSION 5.00
Begin VB.Form frmMainControls 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "VAV Player"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   Icon            =   "frmMainControls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainControls.frx":030A
   ScaleHeight     =   1695
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrLogoAnimation 
      Interval        =   200
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Timer tmrclock 
      Interval        =   500
      Left            =   3000
      Top             =   1200
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   600
      Picture         =   "frmMainControls.frx":0944
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   1
      Left            =   600
      Picture         =   "frmMainControls.frx":09A0
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   2
      Left            =   600
      Picture         =   "frmMainControls.frx":09F6
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   3
      Left            =   600
      Picture         =   "frmMainControls.frx":0A51
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   4
      Left            =   600
      Picture         =   "frmMainControls.frx":0AAC
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   5
      Left            =   600
      Picture         =   "frmMainControls.frx":0B07
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   6
      Left            =   600
      Picture         =   "frmMainControls.frx":0B62
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   7
      Left            =   600
      Picture         =   "frmMainControls.frx":0BBE
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   8
      Left            =   600
      Picture         =   "frmMainControls.frx":0C15
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgDigit 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   9
      Left            =   600
      Picture         =   "frmMainControls.frx":0C61
      ScaleHeight     =   195
      ScaleWidth      =   150
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox imgColon 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   600
      Picture         =   "frmMainControls.frx":0CBC
      ScaleHeight     =   195
      ScaleWidth      =   90
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox imgBlank 
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   600
      Picture         =   "frmMainControls.frx":0CFB
      ScaleHeight     =   375
      ScaleWidth      =   210
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox imgAM 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   600
      Picture         =   "frmMainControls.frx":104A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox imgPM 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   600
      Picture         =   "frmMainControls.frx":11BE
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox imgHourTens 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   5
      Top             =   1320
      Width           =   210
   End
   Begin VB.PictureBox imgHourOnes 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   330
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   4
      Top             =   1320
      Width           =   210
   End
   Begin VB.PictureBox imgHourColon 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   540
      ScaleHeight     =   255
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   1320
      Width           =   150
   End
   Begin VB.PictureBox imgMinuteTens 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   690
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   2
      Top             =   1320
      Width           =   210
   End
   Begin VB.PictureBox imgMinuteOnes 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   900
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   1
      Top             =   1320
      Width           =   210
   End
   Begin VB.PictureBox imgHourAMPM 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   840
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   0
      Top             =   240
      Width           =   210
   End
   Begin VB.Timer tmrposition 
      Interval        =   100
      Left            =   3480
      Top             =   1200
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   8
      Left            =   3240
      Picture         =   "frmMainControls.frx":1332
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   7
      Left            =   3240
      Picture         =   "frmMainControls.frx":1BCB
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   6
      Left            =   3240
      Picture         =   "frmMainControls.frx":2482
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   5
      Left            =   3240
      Picture         =   "frmMainControls.frx":2BC5
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   4
      Left            =   3240
      Picture         =   "frmMainControls.frx":316D
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   3
      Left            =   3240
      Picture         =   "frmMainControls.frx":35B1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   2
      Left            =   3240
      Picture         =   "frmMainControls.frx":3B5B
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   1
      Left            =   3240
      Picture         =   "frmMainControls.frx":4298
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   0
      Left            =   3240
      Picture         =   "frmMainControls.frx":4B43
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image cmdAudioPlayerOFFMouseOver 
      Height          =   480
      Left            =   1920
      Picture         =   "frmMainControls.frx":53E4
      ToolTipText     =   "Audio Player Disabled"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdAudioPlayerOFF 
      Height          =   480
      Left            =   1920
      Picture         =   "frmMainControls.frx":5733
      ToolTipText     =   "Audio Player Disabled"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdAudioPlayerONMouseOver 
      Height          =   480
      Left            =   1920
      Picture         =   "frmMainControls.frx":5A8F
      ToolTipText     =   "Audio Player Enabled"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdAudioPlayerON 
      Height          =   480
      Left            =   1920
      Picture         =   "frmMainControls.frx":5E8C
      ToolTipText     =   "Audio Player Enabled"
      Top             =   240
      Width           =   480
   End
   Begin VB.Image cmdVideoPlayerOFFMouseOver 
      Height          =   480
      Left            =   1200
      Picture         =   "frmMainControls.frx":62A3
      ToolTipText     =   "Video Player Disabled"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdVideoPlayerOFF 
      Height          =   480
      Left            =   1200
      Picture         =   "frmMainControls.frx":6628
      ToolTipText     =   "Video Player Disabled"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdVideoPlayerONMouseOver 
      Height          =   480
      Left            =   1200
      Picture         =   "frmMainControls.frx":69B6
      ToolTipText     =   "Video Player Enabled"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdVideoPlayerON 
      Height          =   480
      Left            =   1200
      Picture         =   "frmMainControls.frx":6E22
      ToolTipText     =   "Video Player Enabled"
      Top             =   240
      Width           =   480
   End
   Begin VB.Image cmdMinimizeMouseDown 
      Height          =   225
      Left            =   120
      Picture         =   "frmMainControls.frx":729A
      ToolTipText     =   "Minimizes VAV Player"
      Top             =   480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMinimizeMouseOver 
      Height          =   225
      Left            =   120
      Picture         =   "frmMainControls.frx":74B6
      ToolTipText     =   "Minimizes VAV Player"
      Top             =   480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdminimize 
      Height          =   225
      Left            =   120
      Picture         =   "frmMainControls.frx":76CF
      ToolTipText     =   "Minimizes VAV Player"
      Top             =   480
      Width           =   225
   End
   Begin VB.Image cmdCloseMouseDown 
      Height          =   225
      Left            =   120
      Picture         =   "frmMainControls.frx":78FC
      ToolTipText     =   "Closes VAV Player"
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdCloseMouseOver 
      Height          =   225
      Left            =   120
      Picture         =   "frmMainControls.frx":7B24
      ToolTipText     =   "Closes VAV Player"
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdclose 
      Height          =   225
      Left            =   120
      Picture         =   "frmMainControls.frx":7D4C
      ToolTipText     =   "Closes VAV Player"
      Top             =   240
      Width           =   225
   End
End
Attribute VB_Name = "frmMainControls"
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
Public minimized As Boolean
Dim HourTime As Integer, MinuteTime As Integer, Colon As Boolean, tempcal As Integer
Dim varLogoAnimation As Byte, controlAnimation As Byte, aniLoop As Byte

Private Sub cmdAudioPlayerOFF_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        AudioEnable = True
        Load frmvavaudioplayer
        Load frmVAVPlayerAudioPlaylist
        frmvavaudioplayer.Left = Me.Left
        frmVAVPlayerAudioPlaylist.Left = Me.Left - frmVAVPlayerAudioPlaylist.Width
        frmvavaudioplayer.Top = Me.Top + Me.Height
        frmVAVPlayerAudioPlaylist.Top = Me.Top + Me.Height
        frmvavaudioplayer.Show
        If AudioPlaylistLeft = True Then
            frmVAVPlayerAudioPlaylist.Show
        End If
        cmdAudioPlayerOFF.Visible = False
        cmdAudioPlayerOFFMouseOver.Visible = False
        cmdAudioPlayerONMouseOver.Visible = False
        cmdAudioPlayerON.Visible = True
        If VideoEnable = True Then
            frmvavvideoplayer.Top = Me.Top + frmvavaudioplayer.Top + frmvavaudioplayer.Height + Me.Height
            frmVAVPlayerVideoPlaylist.Top = Me.Top + frmvavaudioplayer.Top + frmvavaudioplayer.Height
            frmVAVPlayerVideoPlaylist.Left = Me.Left + Me.Width
        End If
    End If
End Sub

Private Sub cmdAudioPlayerOFF_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdAudioPlayerOFF.Visible = False
    cmdAudioPlayerOFFMouseOver.Visible = True
End Sub

Private Sub cmdAudioPlayerOFFMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        AudioEnable = True
        Load frmvavaudioplayer
        Load frmVAVPlayerAudioPlaylist
        frmvavaudioplayer.Left = Me.Left
        frmVAVPlayerAudioPlaylist.Left = Me.Left - frmVAVPlayerAudioPlaylist.Width
        frmvavaudioplayer.Top = Me.Top + Me.Height
        frmVAVPlayerAudioPlaylist.Top = Me.Top + Me.Height
        frmvavaudioplayer.Show
        If AudioPlaylistLeft = True Then
            frmVAVPlayerAudioPlaylist.Show
        End If
        cmdAudioPlayerOFF.Visible = False
        cmdAudioPlayerOFFMouseOver.Visible = False
        cmdAudioPlayerONMouseOver.Visible = False
        cmdAudioPlayerON.Visible = True
        If VideoEnable = True Then
            frmvavvideoplayer.Top = frmvavaudioplayer.Height + Me.Top + Me.Height
            frmVAVPlayerVideoPlaylist.Top = Me.Top + frmvavaudioplayer.Top + frmvavaudioplayer.Height
            frmVAVPlayerVideoPlaylist.Left = Me.Left + Me.Width
        End If
    End If
End Sub

Private Sub cmdAudioPlayerON_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        AudioEnable = False
        frmvavaudioplayer.Hide
        frmVAVPlayerAudioPlaylist.Hide
        Unload frmvavaudioplayer
        Unload frmVAVPlayerAudioPlaylist
        cmdAudioPlayerON.Visible = False
        cmdAudioPlayerONMouseOver.Visible = False
        cmdAudioPlayerOFF.Visible = True
        frmvavvideoplayer.Top = Me.Height
        frmVAVPlayerVideoPlaylist.Top = Me.Top + Me.Height
        frmVAVPlayerVideoPlaylist.Left = Me.Left + Me.Width
    End If
End Sub

Private Sub cmdAudioPlayerON_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdAudioPlayerON.Visible = False
    cmdAudioPlayerONMouseOver.Visible = True
End Sub

Private Sub cmdAudioPlayerONMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        AudioEnable = False
        frmvavaudioplayer.Hide
        frmVAVPlayerAudioPlaylist.Hide
        Unload frmvavaudioplayer
        Unload frmVAVPlayerAudioPlaylist
        cmdAudioPlayerON.Visible = False
        cmdAudioPlayerONMouseOver.Visible = False
        cmdAudioPlayerOFF.Visible = True
        frmvavvideoplayer.Top = Me.Height
        frmVAVPlayerVideoPlaylist.Top = Me.Top + Me.Height
        frmVAVPlayerVideoPlaylist.Left = Me.Left + Me.Width
    End If
End Sub

Private Sub cmdclose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdclose.Visible = False
    cmdCloseMouseOver.Visible = True
    cmdCloseMouseDown.Visible = False
End Sub

Private Sub cmdclose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        ExitAction
    End If
End Sub

Private Sub cmdCloseMouseOver_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        FormMouseMove
        cmdclose.Visible = False
        cmdCloseMouseOver.Visible = False
        cmdCloseMouseDown.Visible = True
    End If
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
    SaveVAVSettings
    Unload frmvavaudioplayer
    Unload frmVAVPlayerAudioPlaylist
    Unload frmvavvideoplayer
    Unload frmVAVPlayerVideoPlaylist
    Unload frmVAVVideoPlayerScreen
    Unload Me
    End
End Sub

Private Sub cmdminimize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdminimize.Visible = False
    cmdMinimizeMouseOver.Visible = True
    cmdMinimizeMouseDown.Visible = False
End Sub

Private Sub cmdminimize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        Me.WindowState = vbMinimized
        frmvavaudioplayer.Hide
        frmVAVPlayerAudioPlaylist.Hide
        frmvavvideoplayer.Hide
        frmVAVPlayerVideoPlaylist.Hide
        minimized = True
    End If
End Sub

Private Sub cmdMinimizeMouseDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        Me.WindowState = vbMinimized
        frmvavaudioplayer.Hide
        frmVAVPlayerAudioPlaylist.Hide
        frmvavvideoplayer.Hide
        frmVAVPlayerVideoPlaylist.Hide
        minimized = True
    End If
End Sub

Private Sub cmdMinimizeMouseOver_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdMinimizeMouseDown.Visible = True
        cmdminimize.Visible = False
        cmdMinimizeMouseOver.Visible = False
    End If
End Sub

Private Sub cmdMinimizeMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        Me.WindowState = vbMinimized
        frmvavaudioplayer.Hide
        frmVAVPlayerAudioPlaylist.Hide
        frmvavvideoplayer.Hide
        frmVAVPlayerVideoPlaylist.Hide
        minimized = True
    End If
End Sub

Private Sub cmdVideoPlayerOFF_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdVideoPlayerOFF.Visible = False
    cmdVideoPlayerOFFMouseOver.Visible = True
End Sub

Private Sub cmdVideoPlayerOFF_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        VideoEnable = True
        Load frmvavvideoplayer
        Load frmVAVPlayerVideoPlaylist
        frmvavvideoplayer.Left = Me.Left
        frmVAVPlayerVideoPlaylist.Left = Me.Left + Me.Width
        If AudioEnable = True Then
            frmvavvideoplayer.Top = frmvavaudioplayer.Top + frmvavaudioplayer.Height
            frmVAVPlayerVideoPlaylist.Top = frmvavaudioplayer.Top + frmvavaudioplayer.Height
        Else
            frmvavvideoplayer.Top = Me.Top + Me.Height
            frmVAVPlayerVideoPlaylist.Top = Me.Top + Me.Height
        End If
        frmvavvideoplayer.Show
        If VideoPlayListRight = True Then
            frmVAVPlayerVideoPlaylist.Show
        End If
        cmdVideoPlayerOFF.Visible = False
        cmdVideoPlayerOFFMouseOver.Visible = False
        cmdVideoPlayerONMouseOver.Visible = False
        cmdVideoPlayerON.Visible = True
    End If
End Sub

Private Sub cmdVideoPlayerOFFMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        VideoEnable = True
        Load frmvavvideoplayer
        Load frmVAVPlayerVideoPlaylist
        frmvavvideoplayer.Left = Me.Left
        frmVAVPlayerVideoPlaylist.Left = Me.Left + Me.Width
        frmVAVPlayerVideoPlaylist.Left = Me.Left + Me.Width
        If AudioEnable = True Then
            frmvavvideoplayer.Top = frmvavaudioplayer.Top + frmvavaudioplayer.Height
            frmVAVPlayerVideoPlaylist.Top = frmvavaudioplayer.Top + frmvavaudioplayer.Height
        Else
            frmvavvideoplayer.Top = Me.Top + Me.Height
            frmVAVPlayerVideoPlaylist.Top = Me.Top + Me.Height
        End If
        frmvavvideoplayer.Show
        If VideoPlayListRight = True Then
            frmVAVPlayerVideoPlaylist.Show
        End If
        cmdVideoPlayerOFF.Visible = False
        cmdVideoPlayerOFFMouseOver.Visible = False
        cmdVideoPlayerONMouseOver.Visible = False
        cmdVideoPlayerON.Visible = True
    End If
End Sub

Private Sub cmdVideoPlayerON_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdVideoPlayerON.Visible = False
    cmdVideoPlayerONMouseOver.Visible = True
End Sub

Private Sub cmdVideoPlayerON_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        VideoEnable = False
        frmvavvideoplayer.Hide
        frmVAVPlayerVideoPlaylist.Hide
        frmVAVVideoPlayerScreen.Hide
        Unload frmvavvideoplayer
        Unload frmVAVPlayerVideoPlaylist
        Unload frmVAVVideoPlayerScreen
        cmdVideoPlayerON.Visible = False
        cmdVideoPlayerONMouseOver.Visible = False
        cmdVideoPlayerOFF.Visible = True
    End If
End Sub

Private Sub cmdVideoPlayerONMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        FormMouseMove
        VideoEnable = False
        frmvavvideoplayer.Hide
        frmVAVPlayerVideoPlaylist.Hide
        frmVAVVideoPlayerScreen.Hide
        Unload frmvavvideoplayer
        Unload frmVAVPlayerVideoPlaylist
        Unload frmVAVVideoPlayerScreen
        cmdVideoPlayerON.Visible = False
        cmdVideoPlayerONMouseOver.Visible = False
        cmdVideoPlayerOFF.Visible = True
    End If
End Sub

Private Sub Form_Click()
    If AudioEnable = True Then
        If AudioPlaylistLeft = True Then
            frmVAVPlayerAudioPlaylist.SetFocus
        End If
        frmvavaudioplayer.SetFocus
    End If
    If VideoEnable = True Then
        frmvavvideoplayer.SetFocus
    End If
End Sub

Private Sub Form_GotFocus()
    Form_Click
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        frmAppPrevInstance.Show 1
        ExitAction
    End If
    Load frmvavaudioplayer
    Load frmvavvideoplayer
    Load frmVAVPlayerAudioPlaylist
    Load frmVAVPlayerVideoPlaylist
    SkipNo = 1
    Me.Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
    LoadVAVSettings
    minimized = False
    If AudioEnable = True Then
        frmvavaudioplayer.Show
        frmvavaudioplayer.Left = frmMainControls.Left
        frmvavaudioplayer.Top = frmMainControls.Top + frmMainControls.Height
        Load frmVAVPlayerAudioPlaylist
        frmVAVPlayerAudioPlaylist.Left = frmvavaudioplayer.Left - frmVAVPlayerAudioPlaylist.Width
        frmVAVPlayerAudioPlaylist.Top = frmvavaudioplayer.Top
        If AudioPlaylistLeft = True Then
            frmVAVPlayerAudioPlaylist.Show
        Else
            frmVAVPlayerAudioPlaylist.Hide
        End If
    End If
    If VideoEnable = True Then
        frmvavvideoplayer.Show
        frmvavvideoplayer.Left = frmMainControls.Left
        frmvavvideoplayer.Top = frmvavaudioplayer.Top + frmvavaudioplayer.Height
        Load frmVAVPlayerVideoPlaylist
        frmVAVPlayerVideoPlaylist.Left = frmvavvideoplayer.Left + frmvavvideoplayer.Width
        frmVAVPlayerVideoPlaylist.Top = frmvavvideoplayer.Top
        If VideoPlayListRight = True Then
            frmVAVPlayerVideoPlaylist.Show
        Else
            frmVAVPlayerVideoPlaylist.Hide
        End If
        If AudioEnable = False Then
            frmvavvideoplayer.Top = Me.Top + Me.Height
            frmVAVPlayerVideoPlaylist.Top = frmvavvideoplayer.Top
            frmVAVPlayerVideoPlaylist.Left = frmvavvideoplayer.Left + frmvavvideoplayer.Width
        End If
        Load frmVAVVideoPlayerScreen
        frmVAVVideoPlayerScreen.Hide
    End If
    Colon = True
    imgHourOnes.Top = imgHourTens.Top
    imgHourColon.Top = imgHourTens.Top
    imgMinuteTens.Top = imgHourTens.Top
    imgMinuteOnes.Top = imgHourTens.Top
    imgHourAMPM.Top = imgHourTens.Top
    imgHourOnes.Left = imgHourTens.Left + imgHourTens.Width
    imgHourColon.Left = imgHourOnes.Left + imgHourOnes.Width
    imgMinuteTens.Left = imgHourColon.Left + imgHourColon.Width
    imgMinuteOnes.Left = imgMinuteTens.Left + imgMinuteTens.Width
    imgHourAMPM.Left = imgMinuteOnes.Left + imgMinuteOnes.Width
    FormMouseMove
    varLogoAnimation = 0
    controlAnimation = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub imgHourAMPM_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub imgHourColon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub imgHourOnes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub imgHourTens_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub imgMinuteOnes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub imgMinuteTens_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub tmrLogoAnimation_Timer()
    If controlAnimation = 0 Then
        varLogoAnimation = varLogoAnimation + 1
    ElseIf controlAnimation = 1 Then
        varLogoAnimation = varLogoAnimation - 1
    End If
    For aniLoop = 0 To 8
        imgLogoAnimation(aniLoop).Visible = False
    Next
    imgLogoAnimation(varLogoAnimation).Visible = True
    If varLogoAnimation = 8 Then
        controlAnimation = 1
    ElseIf varLogoAnimation = 0 Then
        controlAnimation = 0
    End If
End Sub

Private Sub tmrposition_Timer()
    retvalue = GetCursorPos(cursorposition)
    If cursorposition.x < ScaleX(Me.Left, vbTwips, vbPixels) Or cursorposition.x > ScaleX(Me.Left + Me.Width, vbTwips, vbPixels) Or cursorposition.y < ScaleY(Me.Top, vbTwips, vbPixels) Or cursorposition.y > ScaleY(Me.Top + Me.Height, vbTwips, vbPixels) Then
        FormMouseMove
    End If
    If Me.WindowState = vbNormal And minimized = True Then
        If AudioEnable = True Then
            frmvavaudioplayer.Show
            If AudioPlaylistLeft = True Then
                frmVAVPlayerAudioPlaylist.Show
            End If
        End If
        If VideoEnable = True Then
            frmvavvideoplayer.Show
            If VideoPlayListRight = True Then
                frmVAVPlayerVideoPlaylist.Show
            End If
        End If
        minimized = False
    End If
End Sub

Private Sub FormMouseMove()
    cmdclose.Visible = True
    cmdCloseMouseOver.Visible = False
    cmdCloseMouseDown.Visible = False
    cmdminimize.Visible = True
    cmdMinimizeMouseOver.Visible = False
    cmdMinimizeMouseDown.Visible = False
    If VideoEnable = True Then
        cmdVideoPlayerON.Visible = True
        cmdVideoPlayerONMouseOver.Visible = False
        cmdVideoPlayerOFF.Visible = False
        cmdVideoPlayerOFFMouseOver.Visible = False
    Else
        cmdVideoPlayerON.Visible = False
        cmdVideoPlayerONMouseOver.Visible = False
        cmdVideoPlayerOFF.Visible = True
        cmdVideoPlayerOFFMouseOver.Visible = False
    End If
    If AudioEnable = True Then
        cmdAudioPlayerON.Visible = True
        cmdAudioPlayerONMouseOver.Visible = False
        cmdAudioPlayerOFF.Visible = False
        cmdAudioPlayerOFFMouseOver.Visible = False
    Else
        cmdAudioPlayerON.Visible = False
        cmdAudioPlayerONMouseOver.Visible = False
        cmdAudioPlayerOFF.Visible = True
        cmdAudioPlayerOFFMouseOver.Visible = False
    End If
End Sub

Private Sub tmrclock_Timer()
    HourTime = Hour(Now)
    MinuteTime = Minute(Now)
    If HourTime >= 12 Then
        imgHourAMPM.PaintPicture imgPM.Picture, 0, 0
        HourTime = HourTime - 12
    Else
        imgHourAMPM.PaintPicture imgAM.Picture, 0, 0
    End If
    tempcal = Int(HourTime / 10)
    imgHourTens.PaintPicture imgDigit(tempcal).Picture, 0, 0
    tempcal = Int(HourTime Mod 10)
    imgHourOnes.PaintPicture imgDigit(tempcal).Picture, 0, 0
    tempcal = Int(MinuteTime / 10)
    imgMinuteTens.PaintPicture imgDigit(tempcal).Picture, 0, 0
    tempcal = Int(MinuteTime Mod 10)
    imgMinuteOnes.PaintPicture imgDigit(tempcal).Picture, 0, 0
    If Colon = True Then
        imgHourColon.PaintPicture imgColon.Picture, 0, 0
        Colon = False
    ElseIf Colon = False Then
        imgHourColon.PaintPicture imgBlank.Picture, 0, 0
        Colon = True
    End If
End Sub

Private Sub SaveVAVSettings()
    Dim fso As Object
    Dim fs As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(App.Path) > 3 Then
        Set fs = fso.createtextfile(App.Path & "\VAVSettings.ini")
    Else
        Set fs = fso.createtextfile(App.Path & "VAVSettings.ini")
    End If
    fs.writeline ("[VAV Player Settings]")
    If AudioPlaylistLeft = True Then
        fs.writeline ("AudioPlayListLeft=True")
    Else
        fs.writeline ("AudioPlayListLeft=False")
    End If
    If VideoPlayListRight = True Then
        fs.writeline ("VideoPlayListRight=True")
    Else
        fs.writeline ("VideoPlayListRight=False")
    End If
    If VideoEnable = True Then
        fs.writeline ("VideoEnable=True")
    Else
        fs.writeline ("VideoEnable=False")
    End If
    If AudioEnable = True Then
        fs.writeline ("AudioEnable=True")
    Else
        fs.writeline ("AudioEnable=False")
    End If
End Sub

Private Sub LoadVAVSettings()
On Error GoTo ErrorHandlerLoadList
    Dim fso As Object
    Dim fs As Object
    Dim tempstr As String, retstr1 As String, retstr2 As String
    Dim valid As Boolean
    Dim pos As Integer
    valid = True
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(App.Path) > 3 Then
        Set fs = fso.opentextfile(App.Path & "\VAVSettings.ini", 1, False)
    Else
        Set fs = fso.opentextfile(App.Path & "VAVSettings.ini", 1, False)
    End If
    tempstr = fs.readline()
    pos = 1
    Do While fs.AtEndOfStream <> True
    tempstr = fs.readline()
    pos = InStr(1, tempstr, "=")
    pos = pos - 1
    retstr1 = Mid(tempstr, 1, pos)
    retstr2 = Mid(tempstr, pos + 2, Len(tempstr) - pos)
    If LCase(retstr1) = "audioplaylistleft" Then
        If LCase(retstr2) = "false" Then
            AudioPlaylistLeft = False
        Else
            AudioPlaylistLeft = True
        End If
    ElseIf LCase(retstr1) = "videoplaylistright" Then
        If LCase(retstr2) = "false" Then
            VideoPlayListRight = False
        Else
            VideoPlayListRight = True
        End If
    ElseIf LCase(retstr1) = "videoenable" Then
        If LCase(retstr2) = "false" Then
            VideoEnable = False
        Else
            VideoEnable = True
        End If
    ElseIf LCase(retstr1) = "audioenable" Then
        If LCase(retstr2) = "false" Then
            AudioEnable = False
        Else
            AudioEnable = True
        End If
    End If
    Loop
    fs.Close
    Set fso = Nothing
    Exit Sub
ErrorHandlerLoadList:
    AudioEnable = True
    VideoEnable = True
    AudioPlaylistLeft = True
    VideoPlayListRight = True
End Sub
