VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVAVPlayerAudioPlaylist 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   Picture         =   "frmVAVPlayerAudioPlaylist.frx":0000
   ScaleHeight     =   4605
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDVAVPlaylist 
      Left            =   3600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrpos 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox lstAudioPlayer 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   3765
      ItemData        =   "frmVAVPlayerAudioPlaylist.frx":0A1A
      Left            =   0
      List            =   "frmVAVPlayerAudioPlaylist.frx":0A1C
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Image cmdSavePlayListMouseOver 
      Height          =   300
      Left            =   360
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":0A1E
      ToolTipText     =   "Save Playlist."
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdSavePlayList 
      Height          =   300
      Left            =   360
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":0D0E
      ToolTipText     =   "Save Playlist."
      Top             =   4200
      Width           =   300
   End
   Begin VB.Image cmdLoadPlayListMouseOver 
      Height          =   300
      Left            =   960
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":0FFA
      ToolTipText     =   "Load Playlist."
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdLoadPlayList 
      Height          =   300
      Left            =   960
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":12ED
      ToolTipText     =   "Load Playlist."
      Top             =   4200
      Width           =   300
   End
   Begin VB.Image cmdDeleteMouseOver 
      Height          =   300
      Left            =   2784
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":15D6
      ToolTipText     =   "Delete Selected Item From Playlist."
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdDeleteAllMouseOver 
      Height          =   300
      Left            =   3390
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":18D7
      ToolTipText     =   "Delete All Files From Playlist."
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdOpenMouseOver 
      Height          =   300
      Left            =   2178
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":1BFF
      ToolTipText     =   "Add File to Playlist."
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdOpenMultipleMouseOver 
      Height          =   300
      Left            =   1572
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":1E94
      ToolTipText     =   "Add Multiple Files to Playlist."
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AudioPlayer PlayList"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   990
      TabIndex        =   1
      Top             =   0
      Width           =   2115
   End
   Begin VB.Image cmdDelete 
      Height          =   300
      Left            =   2784
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":21B5
      ToolTipText     =   "Delete Selected Item From Playlist."
      Top             =   4200
      Width           =   300
   End
   Begin VB.Image cmdDeleteAll 
      Height          =   300
      Left            =   3390
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":248D
      ToolTipText     =   "Delete All Files From Playlist."
      Top             =   4200
      Width           =   300
   End
   Begin VB.Image cmdOpen 
      Height          =   300
      Left            =   2178
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":27B0
      ToolTipText     =   "Add File to Playlist."
      Top             =   4200
      Width           =   300
   End
   Begin VB.Image cmdOpenMultiple 
      Height          =   300
      Left            =   1572
      Picture         =   "frmVAVPlayerAudioPlaylist.frx":2A1A
      ToolTipText     =   "Add Multiple Files to Playlist."
      Top             =   4200
      Width           =   300
   End
End
Attribute VB_Name = "frmVAVPlayerAudioPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public AudioFile As Byte
Dim retval As Integer, i As Integer
Private Type POINTAPI
        x As Long
        y As Long
End Type
Dim cursorposition As POINTAPI

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdDeleteMouseOver.Visible = True
    cmdDelete.Visible = False
End Sub

Private Sub cmdDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        DeleteAction
    End If
End Sub

Private Sub cmdDeleteMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        DeleteAction
    End If
End Sub

Private Sub cmdOpen_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        AudioFile = 2
        OpenAction
    End If
End Sub

Private Sub cmdOpen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdOpenMouseOver.Visible = True
    cmdOpen.Visible = False
End Sub

Private Sub cmdOpenmouseover_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        AudioFile = 2
        OpenAction
    End If
End Sub

Private Sub cmdOpenMultiple_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdOpenMultipleMouseOver.Visible = True
    cmdOpenMultiple.Visible = False
End Sub

Private Sub cmdOpenMultiple_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        AudioFile = 1
        OpenAction
    End If
End Sub

Private Sub cmdOpenMultiplemouseover_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        AudioFile = 1
        OpenAction
    End If
End Sub

Private Sub cmdDeleteAll_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdDeleteAllMouseOver.Visible = True
    cmdDeleteAll.Visible = False
End Sub

Private Sub cmdDeleteAll_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        If lstAudioPlayer.ListCount > 0 Then
            retval = MsgBox("Remove All The Items From Playlist", vbQuestion + vbYesNo, "MSP's Media Player List Clear")
                If retval = vbYes Then
                    lstAudioPlayer.Clear
                End If
        Else
            retval = MsgBox("List Already Empty.", vbInformation + vbOKOnly, "MSP's Media Player List Clear")
        End If
    End If
End Sub

Private Sub cmdDeleteAllMouseover_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
    If lstAudioPlayer.ListCount > 0 Then
        retval = MsgBox("Remove All The Items From Playlist", vbQuestion + vbYesNo, "MSP's Media Player List Clear")
            If retval = vbYes Then
                lstAudioPlayer.Clear
            End If
    Else
        retval = MsgBox("List Already Empty.", vbInformation + vbOKOnly, "MSP's Media Player List Clear")
    End If
    End If
End Sub

Private Sub OpenAction()
    FormMouseMove
    frmFileSelectAudioPlayer.Show 1
End Sub

Private Sub DeleteAction()
    Dim Index As Integer
    If lstAudioPlayer.ListCount > 0 And lstAudioPlayer.ListIndex >= 0 Then
        Index = lstAudioPlayer.ListIndex
        lstAudioPlayer.RemoveItem (lstAudioPlayer.ListIndex)
        If Index < lstAudioPlayer.ListCount - 2 Then
            lstAudioPlayer.Selected(Index) = True
        Else
            If Index = 0 And lstAudioPlayer.ListCount > 0 Then
                lstAudioPlayer.Selected(0) = True
            ElseIf Index > 0 Then
                lstAudioPlayer.Selected(Index - 1) = True
            End If
        End If
    End If
End Sub

Private Sub FormMouseMove()
    cmdOpen.Visible = True
    cmdOpenMouseOver.Visible = False
    cmdDelete.Visible = True
    cmdDeleteMouseOver.Visible = False
    cmdOpenMultiple.Visible = True
    cmdOpenMultipleMouseOver.Visible = False
    cmdDeleteAll.Visible = True
    cmdDeleteAllMouseOver.Visible = False
    cmdSavePlayList.Visible = True
    cmdSavePlayListMouseOver.Visible = False
    cmdLoadPlayList.Visible = True
    cmdLoadPlayListMouseOver.Visible = False
End Sub

Private Sub cmdSavePlayList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdSavePlayListMouseOver.Visible = True
End Sub

Private Sub cmdSavePlayList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        SavePlayList
    End If
End Sub

Private Sub cmdSavePlayListMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        SavePlayList
    End If
End Sub

Private Sub cmdLoadPlayList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    cmdLoadPlayListMouseOver.Visible = True
End Sub

Private Sub cmdLoadPlayList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        LoadPlayList
    End If
End Sub

Private Sub cmdLoadPlayListMouseOver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Click
    If Button = vbLeftButton Then
        LoadPlayList
    End If
End Sub

Private Sub Form_Click()
    frmMainControls.SetFocus
    frmvavaudioplayer.SetFocus
    If VideoEnable = True Then
        frmvavvideoplayer.SetFocus
        If VideoPlayListRight = True Then
            frmVAVPlayerVideoPlaylist.SetFocus
        End If
    End If
End Sub

Private Sub Form_GotFocus()
    Form_Click
End Sub

Private Sub Form_Load()
On Error GoTo errorhandler:
    Dim strname As String
    Dim fso As Object
    Dim fs As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.opentextfile(App.Path & "\fileplay.va")
    strname = fs.readline()
    Do While fs.AtEndOfStream <> True
    strname = fs.readline()
    lstAudioPlayer.AddItem (strname)
    Loop
    fs.Close
    Set fso = Nothing
    If lstAudioPlayer.ListCount > 0 Then
        lstAudioPlayer.Selected(0) = True
    End If
    Exit Sub
errorhandler:
    If Err.Number = 51 Then Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strname As String
    Dim fso As Object
    Dim fs As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.createtextfile(App.Path & "\fileplay.va")
    fs.writeline (" ")
    For i = 0 To lstAudioPlayer.ListCount - 1
        strname = lstAudioPlayer.List(i)
        fs.writeline (strname)
    Next
    fs.Close
    Set fso = Nothing
End Sub

Private Sub lstAudioPlayer_Click()
    If frmvavaudioplayer.playing = False Then
        frmvavaudioplayer.cmdPlay.Visible = True
        frmvavaudioplayer.cmdPlayActive.Visible = False
    End If
End Sub

Private Sub lstAudioPlayer_DblClick()
    frmvavaudioplayer.MainMediaVAVAudio.FileName = lstAudioPlayer.List(lstAudioPlayer.ListIndex)
    frmvavaudioplayer.PlayAction
End Sub

Private Sub lstAudioPlayer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        DeleteAction
    ElseIf KeyCode = 13 Then
        frmvavaudioplayer.PlayAction
    End If
    lstAudioPlayer.SetFocus
End Sub

Private Sub lstAudioPlayer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormMouseMove
    If lstAudioPlayer.ListCount > 0 And lstAudioPlayer.ListIndex >= 0 Then
        lstAudioPlayer.ToolTipText = lstAudioPlayer.Text
    Else
        lstAudioPlayer.ToolTipText = "Audio Player Playlist"
    End If
End Sub

Private Sub tmrpos_Timer()
    retval = GetCursorPos(cursorposition)
    If cursorposition.x < ScaleX(Me.Left, vbTwips, vbPixels) Or cursorposition.x > ScaleX(Me.Left + Me.Width, vbTwips, vbPixels) Or cursorposition.y < ScaleY(Me.Top, vbTwips, vbPixels) Or cursorposition.y > ScaleY(Me.Top + Me.Height, vbTwips, vbPixels) Then
        FormMouseMove
    End If
End Sub

Private Sub SavePlayList()
Dim PlayListFilename As String
Dim strname As String
Dim fso As Object
Dim fs As Object
On Error GoTo SavePlayListError
    CDVAVPlaylist.CancelError = True
    CDVAVPlaylist.FileName = ""
    CDVAVPlaylist.Flags = cdlOFNOverwritePrompt
    CDVAVPlaylist.DialogTitle = "Save Audio PlayList"
    CDVAVPlaylist.DefaultExt = "*.vapl"
    CDVAVPlaylist.Filter = "VAV PlayList(Audio)|*.vapl"
    CDVAVPlaylist.ShowSave
    PlayListFilename = CDVAVPlaylist.FileName
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.createtextfile(PlayListFilename)
    fs.writeline (" ")
    For i = 0 To lstAudioPlayer.ListCount - 1
        strname = lstAudioPlayer.List(i)
        fs.writeline (strname)
    Next
    fs.Close
    Set fso = Nothing
    Exit Sub
SavePlayListError:
    If Err.Number = cdlCancel Then Exit Sub
    retval = MsgBox("UnKnown Error Occured While Saving PlayList.", vbCritical + vbOKOnly, "PlayList Save Error")
End Sub

Private Sub LoadPlayList()
Dim PlayListFilename As String
Dim strname As String
Dim fso As Object
Dim fs As Object
On Error GoTo SavePlayListError
    CDVAVPlaylist.CancelError = True
    CDVAVPlaylist.FileName = ""
    CDVAVPlaylist.Flags = cdlOFNFileMustExist
    CDVAVPlaylist.DialogTitle = "Load Audio PlayList"
    CDVAVPlaylist.DefaultExt = "*.vapl"
    CDVAVPlaylist.Filter = "VAV PlayList(Audio)|*.vapl"
    CDVAVPlaylist.ShowOpen
    PlayListFilename = CDVAVPlaylist.FileName
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso.opentextfile(PlayListFilename)
    strname = fs.readline()
    Do While fs.AtEndOfStream <> True
    strname = fs.readline()
    lstAudioPlayer.AddItem (strname)
    Loop
    fs.Close
    Set fso = Nothing
    Exit Sub
SavePlayListError:
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number = 51 Then Exit Sub
    retval = MsgBox("UnKnown Error Occured While Loading PlayList.", vbCritical + vbOKOnly, "PlayList Save Error")
End Sub
