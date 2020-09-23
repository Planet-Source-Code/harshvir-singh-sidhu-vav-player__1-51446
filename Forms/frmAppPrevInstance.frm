VERSION 5.00
Begin VB.Form frmAppPrevInstance 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "VAV Player Already Started"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrposition 
      Interval        =   100
      Left            =   4080
      Top             =   2040
   End
   Begin VB.Timer tmrControlAnimation 
      Interval        =   200
      Left            =   4560
      Top             =   2040
   End
   Begin VB.Image cmdPrevOkMouseDown 
      Height          =   450
      Left            =   1920
      Picture         =   "frmAppPrevInstance.frx":0000
      Top             =   2040
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblPrevInstance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "An Instance Of  VAV Player is Already Running."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4605
   End
   Begin VB.Label lblPrevInstance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAV Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1710
   End
   Begin VB.Image cmdPrevOKMouseOver 
      Height          =   450
      Left            =   1920
      Picture         =   "frmAppPrevInstance.frx":04C3
      Top             =   2040
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image cmdPrevOk 
      Height          =   450
      Left            =   1920
      Picture         =   "frmAppPrevInstance.frx":09DA
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   0
      Left            =   3960
      Picture         =   "frmAppPrevInstance.frx":0E94
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   1
      Left            =   3960
      Picture         =   "frmAppPrevInstance.frx":1735
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   2
      Left            =   3960
      Picture         =   "frmAppPrevInstance.frx":1FE0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   3
      Left            =   3960
      Picture         =   "frmAppPrevInstance.frx":271D
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   4
      Left            =   3960
      Picture         =   "frmAppPrevInstance.frx":2CC7
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   5
      Left            =   3960
      Picture         =   "frmAppPrevInstance.frx":310B
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   6
      Left            =   3960
      Picture         =   "frmAppPrevInstance.frx":36B3
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   7
      Left            =   3960
      Picture         =   "frmAppPrevInstance.frx":3DF6
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLogoAnimation 
      Height          =   750
      Index           =   8
      Left            =   3960
      Picture         =   "frmAppPrevInstance.frx":46AD
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmAppPrevInstance"
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
Dim varLogoAnimation As Byte, controlAnimation As Byte, aniLoop As Byte

Private Sub cmdPrevOk_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdPrevOk.Visible = False
    cmdPrevOkMouseDown.Visible = False
    cmdPrevOKMouseOver.Visible = True
End Sub

Private Sub cmdPrevOk_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Unload Me
    End If
End Sub

Private Sub cmdPrevOkMouseDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdPrevOKMouseOver.Visible = True
        cmdPrevOkMouseDown.Visible = False
        cmdPrevOk.Visible = False
        Unload Me
    End If
End Sub

Private Sub cmdPrevOkMouseDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdPrevOKMouseOver.Visible = True
        cmdPrevOkMouseDown.Visible = False
        cmdPrevOk.Visible = False
        Unload Me
    End If
End Sub

Private Sub cmdPrevOKMouseOver_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdPrevOkMouseDown.Visible = True
        cmdPrevOKMouseOver.Visible = False
        cmdPrevOk.Visible = False
    End If
End Sub

Private Sub Form_Load()
    controlAnimation = 0
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdPrevOk.Visible = True
    cmdPrevOKMouseOver.Visible = False
    cmdPrevOkMouseDown.Visible = False
End Sub

Private Sub tmrControlAnimation_Timer()
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
        cmdPrevOk.Visible = True
        cmdPrevOKMouseOver.Visible = False
        cmdPrevOkMouseDown.Visible = False
    End If
End Sub
