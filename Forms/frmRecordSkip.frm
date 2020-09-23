VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRecordSkip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MSP's VAV Video Player Skipping Recorder."
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   Icon            =   "frmRecordSkip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "R&eset"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Thats IT!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4170
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdRecordSkip 
      Caption         =   "&Record Jump Start Time"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1170
      TabIndex        =   2
      Top             =   4440
      Width           =   2655
   End
   Begin ComctlLib.Slider sldRecordControl 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   327682
      Max             =   100
   End
   Begin MediaPlayerCtl.MediaPlayer mpRecordSkip 
      CausesValidation=   0   'False
      Height          =   3375
      Left            =   1800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4335
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   0   'False
      AllowChangeDisplaySize=   0   'False
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
      Enabled         =   0   'False
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
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
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmRecordSkip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdRecordSkip_Click()
    If cmdRecordSkip.Caption = "&Record Jump Start Time" Then
        NoOfSkips = NoOfSkips + 1
        cmdRecordSkip.Caption = "&Record Jump Stop Time"
        RecordSkip(frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex, NoOfSkips, 1) = Int(mpRecordSkip.CurrentPosition)
        cmdOK.Enabled = False
    ElseIf cmdRecordSkip.Caption = "&Record Jump Stop Time" Then
        cmdRecordSkip.Caption = "&Record Jump Start Time"
        RecordSkip(frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex, NoOfSkips, 2) = Int(mpRecordSkip.CurrentPosition)
        cmdOK.Enabled = True
    End If
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer
    For i = 1 To NoOfSkips
        RecordSkip(frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex, i, 1) = 0
        RecordSkip(frmVAVPlayerVideoPlaylist.lstVideoPlayer.ListIndex, i, 2) = 0
    Next
    cmdRecordSkip.Caption = "&Record Jump Start Time"
    cmdOK.Enabled = True
    SkipNo = 1
    NoOfSkips = 0
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    If Len(frmVAVPlayerVideoPlaylist.lstVideoPlayer.Text) > 1 Then
        mpRecordSkip.FileName = frmVAVPlayerVideoPlaylist.lstVideoPlayer.Text
    End If
    sldRecordControl.Max = mpRecordSkip.Duration
    mpRecordSkip.Play
End Sub

Private Sub sldRecordControl_Change()
    mpRecordSkip.CurrentPosition = sldRecordControl.Value
End Sub

Private Sub sldRecordControl_Scroll()
    mpRecordSkip.CurrentPosition = sldRecordControl.Value
End Sub
