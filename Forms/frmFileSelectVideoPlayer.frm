VERSION 5.00
Begin VB.Form frmFileSelectVideoPlayer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Video File"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   Icon            =   "frmFileSelectVideoPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      ToolTipText     =   "Adds the list of Selected Songs in Your Video Playlist."
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      ToolTipText     =   "Closes the Video Selection Window and Discards the list of selected Videos."
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox cboFilePattern 
      Height          =   315
      ItemData        =   "frmFileSelectVideoPlayer.frx":000C
      Left            =   2640
      List            =   "frmFileSelectVideoPlayer.frx":0019
      TabIndex        =   16
      Text            =   "*.dat"
      Top             =   3000
      Width           =   2600
   End
   Begin VB.DriveListBox drvOpenVideo 
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Top             =   3000
      Width           =   2600
   End
   Begin VB.DirListBox dirOpenVideo 
      Height          =   3015
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   2600
   End
   Begin VB.FileListBox filOpenVideoMultiple 
      Height          =   3015
      Left            =   2640
      MultiSelect     =   1  'Simple
      Pattern         =   "*.dat"
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.Frame frasearch 
      Caption         =   "Search"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   3360
      Width           =   4455
      Begin VB.CommandButton cmdSearchStart 
         Caption         =   "&Start"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         ToolTipText     =   "Start Search For Selected Pattern in Selected Drive."
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cboDrives 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "s"
         Top             =   360
         Width           =   2600
      End
      Begin VB.ComboBox cboSearchPattern 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmFileSelectVideoPlayer.frx":0033
         Left            =   1680
         List            =   "frmFileSelectVideoPlayer.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   2600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Drive"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pattern"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   60
      End
   End
   Begin VB.FileListBox filOpenVideo 
      Height          =   3015
      Left            =   2640
      Pattern         =   "*.dat"
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.ListBox lstFileSelected 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      ItemData        =   "frmFileSelectVideoPlayer.frx":005A
      Left            =   5880
      List            =   "frmFileSelectVideoPlayer.frx":005C
      MultiSelect     =   1  'Simple
      TabIndex        =   4
      Top             =   0
      Width           =   4815
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5250
      TabIndex        =   3
      ToolTipText     =   "Add Selected Video To Selection List"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdRemoveFile 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5250
      TabIndex        =   2
      ToolTipText     =   "Removes Selected Video From Selection List"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdMoveAll 
      Caption         =   ">>>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5250
      TabIndex        =   1
      ToolTipText     =   "Add All Videos To Selection List"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdRemoveFileAll 
      Caption         =   "<<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5250
      TabIndex        =   0
      ToolTipText     =   "Removes All Video From Selection List"
      Top             =   2520
      Width           =   615
   End
End
Attribute VB_Name = "frmFileSelectVideoPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private StringDrives() As String
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const MAX_PATH = 260
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type
Dim i As Integer
Dim retval As Integer

Private Sub DriveChange()
On Error GoTo errorhandelopen
    dirOpenVideo.Path = drvOpenVideo.Drive
Exit Sub
errorhandelopen:
    If Err.Number = 68 Then
        retval = MsgBox("Drive :  " & UCase$(drvOpenVideo.Drive) & " Not Ready.", vbCritical + vbRetryCancel, "VAV Player File Select Error")
        If retval = vbRetry Then
            DriveChange
        ElseIf retval = vbCancel Then
            retval = 0
            drvOpenVideo.Drive = "C:\"
            dirOpenVideo.Path = drvOpenVideo.Drive
            filOpenVideo.Path = dirOpenVideo.Path
            filOpenVideoMultiple.Path = dirOpenVideo.Path
        End If
    Else
        retval = MsgBox("UnKnown Error Occured.", vbCritical + vbRetryCancel, "MSP's Media Player File Select Error")
    End If
End Sub

Private Sub cmdAddFile_Click()
On Error GoTo ErrorAddFile
    If filOpenVideo.Visible = True Then
        If filOpenVideo.ListIndex >= 0 Then
            If Len(dirOpenVideo.Path) > 3 Then
                If filOpenVideo.Selected(filOpenVideo.ListIndex) = True Then
                    lstFileSelected.AddItem (dirOpenVideo.Path) & "\" & filOpenVideo.List(filOpenVideo.ListIndex)
                End If
            Else
                If filOpenVideo.Selected(filOpenVideo.ListIndex) = True Then
                    lstFileSelected.AddItem (dirOpenVideo.Path) & filOpenVideo.List(filOpenVideo.ListIndex)
                End If
            End If
                cmdOpen.Enabled = True
        Else
            retval = MsgBox("No File Selected To Add.", vbCritical + vbOKOnly, "VAV Player Add File Error.")
        End If
    ElseIf filOpenVideoMultiple.Visible = True Then
        For i = 0 To filOpenVideoMultiple.ListCount - 1
            If filOpenVideoMultiple.Selected(i) = True Then
                If Len(dirOpenVideo.Path) > 3 Then
                    lstFileSelected.AddItem (dirOpenVideo.Path & "\" & filOpenVideoMultiple.List(i))
                Else
                    lstFileSelected.AddItem (dirOpenVideo.Path & filOpenVideoMultiple.List(i))
                End If
            End If
        Next
        cmdOpen.Enabled = True
    End If
    Exit Sub
ErrorAddFile:
    If Err.Number = 381 Then
        retval = MsgBox("No File Selected To Add.", vbCritical + vbOKOnly, "VAV Player Add File Error.")
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdMoveAll_Click()
    If Len(dirOpenVideo.Path) > 3 Then
        For i = 0 To filOpenVideoMultiple.ListCount - 1
            lstFileSelected.AddItem (dirOpenVideo.Path & "\" & filOpenVideoMultiple.List(i))
        Next
    Else
        For i = 0 To filOpenVideoMultiple.ListCount - 1
            lstFileSelected.AddItem (dirOpenVideo.Path & filOpenVideoMultiple.List(i))
        Next
    End If
    If lstFileSelected.ListCount > 0 Then
        cmdOpen.Enabled = True
    End If
End Sub

Private Sub cmdopen_Click()
    For i = 0 To lstFileSelected.ListCount - 1 Step 1
        frmVAVPlayerVideoPlaylist.lstVideoPlayer.AddItem (lstFileSelected.List(i))
    Next
    Unload Me
End Sub

Private Sub cmdRemoveFile_Click()
    On Error GoTo errorRemove
    For i = lstFileSelected.ListCount - 1 To 0 Step -1
        If lstFileSelected.Selected(i) = True Then
                lstFileSelected.RemoveItem (i)
        End If
    Next
    If lstFileSelected.ListCount <= 0 Then
        cmdOpen.Enabled = False
    End If
    Exit Sub
errorRemove:
    If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboFilePattern_Change()
    filOpenVideo.Pattern = cboFilePattern.List(cboFilePattern.ListIndex)
    filOpenVideoMultiple.Pattern = cboFilePattern.List(cboFilePattern.ListIndex)
End Sub

Private Sub cboFilePattern_Click()
    filOpenVideo.Pattern = cboFilePattern.List(cboFilePattern.ListIndex)
    filOpenVideoMultiple.Pattern = cboFilePattern.List(cboFilePattern.ListIndex)
End Sub

Private Sub cmdRemoveFileAll_Click()
    lstFileSelected.Clear
    cmdOpen.Enabled = False
End Sub

Private Sub cmdSearchStart_Click()
Dim loopvar As Integer
    lblSearch.Caption = "Searching..."
    If InStr(cboDrives.Text, "All Local Fixed Drives") > 0 Then
        For loopvar = 0 To UBound(StringDrives)
            GetFiles StringDrives(loopvar), True, cboSearchPattern.Text
        Next
    Else
        GetFiles cboDrives.Text, True, cboSearchPattern.Text
    End If
    If lstFileSelected.ListCount > 0 Then
        cmdOpen.Enabled = True
    End If
    lblSearch.Caption = ""
End Sub

Private Sub dirOpenVideo_Change()
    filOpenVideo.Path = dirOpenVideo.Path
    filOpenVideoMultiple.Path = dirOpenVideo.Path
End Sub

Private Sub drvOpenVideo_Change()
    DriveChange
End Sub

Private Sub Form_Load()
Dim loopvar As Integer
Dim DriveCount As Integer
Dim StringAllDrive As String
Dim StringDrive As String
ReDim StringDrives(0) As String
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    If frmVAVPlayerVideoPlaylist.VideoFile = 1 Then
        filOpenVideoMultiple.Visible = True
        cmdSearchStart.Enabled = True
        cboDrives.Enabled = True
        cboSearchPattern.Enabled = True
        Label1.Enabled = True
        Label2.Enabled = True
        frasearch.Enabled = True
        cmdMoveAll.Enabled = True
        cmdRemoveFileAll.Enabled = True
    ElseIf frmVAVPlayerVideoPlaylist.VideoFile = 2 Then
        filOpenVideo.Visible = True
    End If
    For loopvar = 66 To 90
        StringDrive = Chr(loopvar) & ":\"
        If DriveType(StringDrive) = "Fixed Drive" Then
            If StringDrives(0) = "" Then
                StringDrives(0) = StringDrive
            Else
                ReDim Preserve StringDrives(UBound(StringDrives) + 1) As String
                StringDrives(UBound(StringDrives)) = StringDrive
            End If
            cboDrives.AddItem StringDrive
            If StringAllDrive <> "" Then StringAllDrive = StringAllDrive & ", "
            StringAllDrive = StringAllDrive & StringDrive
            DriveCount = DriveCount + 1
        End If
    Next
    If DriveCount > 1 Then
        StringAllDrive = "All Local Fixed Drives (" & StringAllDrive & ")"
        cboDrives.AddItem StringAllDrive
    End If
    cboDrives.ListIndex = 0
    cboSearchPattern.ListIndex = 0
End Sub

Private Function DriveType(Drive As String) As String
Dim sAns As String, lAns As Long
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" _
        Then Drive = Drive & "\"
    lAns = GetDriveType(Drive)
    Select Case lAns
     Case 2
       sAns = "Removable Drive"
     Case 3
       sAns = "Fixed Drive"
     Case 4
        sAns = "Remote Drive"
     Case 5
        sAns = "CD-ROM"
     Case 6
        sAns = "RAM Disk"
     Case Else
        sAns = "Drive Doesn't Exist"
    End Select
    DriveType = sAns
End Function

Public Sub GetFiles(Path As String, SubFolder As Boolean, Optional Pattern As String = "*.*")
    Screen.MousePointer = vbHourglass
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, fPath As String, fName As String
    fPath = AddBackslash(Path)
    Dim sPattern As String
    sPattern = Pattern
    fName = fPath & sPattern
    hFile = FindFirstFile(fName, WFD)
    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
        lstFileSelected.AddItem (fPath & StripNulls(WFD.cFileName))
    End If
    If hFile > 0 Then
    While FindNextFile(hFile, WFD)
        If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
        lstFileSelected.AddItem (fPath & StripNulls(WFD.cFileName))
        End If
    Wend
    End If
    If SubFolder Then
        hFile = FindFirstFile(fPath & "*.*", WFD)
        If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
        StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
           GetFiles fPath & StripNulls(WFD.cFileName), True, sPattern
        End If
        While FindNextFile(hFile, WFD)
            If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
            StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
                GetFiles fPath & StripNulls(WFD.cFileName), True, sPattern
            End If
        Wend
    End If
    FindClose hFile
    Screen.MousePointer = vbDefault
End Sub

Private Function StripNulls(f As String) As String
    StripNulls = Left$(f, InStr(1, f, Chr$(0)) - 1)
End Function

Private Function AddBackslash(s As String) As String
    If Len(s) Then
       If Right$(s, 1) <> "\" Then
          AddBackslash = s & "\"
       Else
          AddBackslash = s
       End If
    Else
       AddBackslash = "\"
    End If
End Function

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

