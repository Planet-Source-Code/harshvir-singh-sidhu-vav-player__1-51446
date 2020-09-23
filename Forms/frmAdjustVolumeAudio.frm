VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAdjustVolumeAudio 
   BorderStyle     =   0  'None
   Caption         =   "Adjust Volume For Audio Player"
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Slider SliderMainMediaAudioPlayerVolume 
      Height          =   510
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Adjust Audio Player's Volume"
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   900
      _Version        =   327682
      LargeChange     =   50
      Min             =   -3000
      Max             =   0
      TickStyle       =   2
   End
End
Attribute VB_Name = "frmAdjustVolumeAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SliderMainMediaAudioPlayerVolume_Click()
    frmvavaudioplayer.MainMediaVAVAudio.Volume = SliderMainMediaAudioPlayerVolume.Value
End Sub

Private Sub SliderMainMediaAudioPlayerVolume_LostFocus()
    Unload Me
End Sub

Private Sub SliderMainMediaAudioPlayerVolume_Scroll()
    frmvavaudioplayer.MainMediaVAVAudio.Volume = SliderMainMediaAudioPlayerVolume.Value
End Sub
