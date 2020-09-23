VERSION 5.00
Begin VB.Form frmFileSelectAudioPlayer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Audio File"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   ControlBox      =   0   'False
   Icon            =   "frmFileSelectAudioPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   18
      ToolTipText     =   "Add Selected Song To Selection List"
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
      TabIndex        =   17
      ToolTipText     =   "Delete Selected Song From Selection List"
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
      TabIndex        =   16
      ToolTipText     =   "Add All Songs To Selection List"
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
      TabIndex        =   15
      ToolTipText     =   "Delete All Songs From Selection List"
      Top             =   2520
      Width           =   615
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
      ItemData        =   "frmFileSelectAudioPlayer.frx":000C
      Left            =   5880
      List            =   "frmFileSelectAudioPlayer.frx":000E
      MultiSelect     =   1  'Simple
      TabIndex        =   13
      Top             =   0
      Width           =   4815
   End
   Begin VB.FileListBox filOpenAudio 
      Height          =   3015
      Left            =   2640
      Pattern         =   "*.mp3"
      TabIndex        =   12
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
      Begin VB.ComboBox cboSearchPattern 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmFileSelectAudioPlayer.frx":0010
         Left            =   1680
         List            =   "frmFileSelectAudioPlayer.frx":0023
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   2600
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
         TabIndex        =   7
         ToolTipText     =   "Start Search For Selected Pattern in Selected Drive."
         Top             =   1200
         Width           =   1215
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
         TabIndex        =   14
         Top             =   1320
         Width           =   60
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
         TabIndex        =   10
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.FileListBox filOpenAudioMultiple 
      Height          =   3015
      Left            =   2640
      MultiSelect     =   1  'Simple
      Pattern         =   "*.mp3"
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.DirListBox dirOpenAudio 
      Height          =   3015
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2600
   End
   Begin VB.DriveListBox drvOpenAudio 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   2600
   End
   Begin VB.ComboBox cboFilePattern 
      Height          =   315
      ItemData        =   "frmFileSelectAudioPlayer.frx":004A
      Left            =   2640
      List            =   "frmFileSelectAudioPlayer.frx":005D
      TabIndex        =   2
      Text            =   "*.mp3"
      Top             =   3000
      Width           =   2600
   End
   Begin VB.CommandButton cmdcancel 
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
      TabIndex        =   1
      ToolTipText     =   "Closes the Song Selection Window and Discards the list of selected Songs."
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdopen 
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
      TabIndex        =   0
      ToolTipText     =   "Adds the list of Selected Songs in Your Audio Playlist."
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmFileSelectAudioPlayer"
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
    dirOpenAudio.Path = drvOpenAudio.Drive
Exit Sub
errorhandelopen:
    If Err.Number = 68 Then
        retval = MsgBox("Drive :  " & UCase$(drvOpenAudio.Drive) & " Not Ready.", vbCritical + vbRetryCancel, "VAV Player File Select Error")
        If retval = vbRetry Then
            DriveChange
        ElseIf retval = vbCancel Then
            retval = 0
            drvOpenAudio.Drive = "C:\"
            dirOpenAudio.Path = drvOpenAudio.Drive
            filOpenAudio.Path = dirOpenAudio.Path
            filOpenAudioMultiple.Path = dirOpenAudio.Path
        End If
    Else
        retval = MsgBox("UnKnown Error Occured.", vbCritical + vbRetryCancel, "MSP's Media Player File Select Error")
    End If
End Sub

Private Sub cmdAddFile_Click()
On Error GoTo ErrorAddFile
    If filOpenAudio.Visible = True Then
        If filOpenAudio.ListIndex >= 0 Then
            If Len(dirOpenAudio.Path) > 3 Then
                If filOpenAudio.Selected(filOpenAudio.ListIndex) = True Then
                    lstFileSelected.AddItem (dirOpenAudio.Path) & "\" & filOpenAudio.List(filOpenAudio.ListIndex)
                End If
            Else
                If filOpenAudio.Selected(filOpenAudio.ListIndex) = True Then
                    lstFileSelected.AddItem (dirOpenAudio.Path) & filOpenAudio.List(filOpenAudio.ListIndex)
                End If
            End If
                cmdOpen.Enabled = True
        Else
            retval = MsgBox("No File Selected To Add.", vbCritical + vbOKOnly, "VAV Player Add File Error.")
        End If
    ElseIf filOpenAudioMultiple.Visible = True Then
        For i = 0 To filOpenAudioMultiple.ListCount - 1
            If filOpenAudioMultiple.Selected(i) = True Then
                If Len(dirOpenAudio.Path) > 3 Then
                    lstFileSelected.AddItem (dirOpenAudio.Path & "\" & filOpenAudioMultiple.List(i))
                Else
                    lstFileSelected.AddItem (dirOpenAudio.Path & filOpenAudioMultiple.List(i))
                End If
            End If
        Next
        If lstFileSelected.ListCount > 0 Then
            cmdOpen.Enabled = True
        End If
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
    If Len(dirOpenAudio.Path) > 3 Then
        For i = 0 To filOpenAudioMultiple.ListCount - 1
            lstFileSelected.AddItem (dirOpenAudio.Path & "\" & filOpenAudioMultiple.List(i))
        Next
    Else
        For i = 0 To filOpenAudioMultiple.ListCount - 1
            lstFileSelected.AddItem (dirOpenAudio.Path & filOpenAudioMultiple.List(i))
        Next
    End If
    If lstFileSelected.ListCount > 0 Then
        cmdOpen.Enabled = True
    End If
End Sub

Private Sub cmdopen_Click()
    For i = 0 To lstFileSelected.ListCount - 1 Step 1
        frmVAVPlayerAudioPlaylist.lstAudioPlayer.AddItem (lstFileSelected.List(i))
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
    filOpenAudio.Pattern = cboFilePattern.List(cboFilePattern.ListIndex)
    filOpenAudioMultiple.Pattern = cboFilePattern.List(cboFilePattern.ListIndex)
End Sub

Private Sub cboFilePattern_Click()
    filOpenAudio.Pattern = cboFilePattern.List(cboFilePattern.ListIndex)
    filOpenAudioMultiple.Pattern = cboFilePattern.List(cboFilePattern.ListIndex)
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

Private Sub dirOpenAudio_Change()
    filOpenAudio.Path = dirOpenAudio.Path
    filOpenAudioMultiple.Path = dirOpenAudio.Path
End Sub

Private Sub drvOpenAudio_Change()
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
    If frmVAVPlayerAudioPlaylist.AudioFile = 1 Then
        filOpenAudioMultiple.Visible = True
        cmdSearchStart.Enabled = True
        cboDrives.Enabled = True
        cboSearchPattern.Enabled = True
        Label1.Enabled = True
        Label2.Enabled = True
        frasearch.Enabled = True
        cmdMoveAll.Enabled = True
        cmdRemoveFileAll.Enabled = True
    ElseIf frmVAVPlayerAudioPlaylist.AudioFile = 2 Then
        filOpenAudio.Visible = True
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
