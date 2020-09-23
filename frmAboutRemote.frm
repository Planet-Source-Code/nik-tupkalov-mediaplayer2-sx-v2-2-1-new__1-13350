VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   3465
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   5250
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmAboutRemote.frx":0000
   ScaleHeight     =   2391.605
   ScaleMode       =   0  'User
   ScaleWidth      =   4930.022
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4095
      Top             =   1125
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   405
      Picture         =   "frmAboutRemote.frx":237E
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      Top             =   675
      Width           =   540
   End
   Begin VB.Label lTime 
      BackStyle       =   0  'Transparent
      Caption         =   "21:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   8
      Top             =   315
      Width           =   960
   End
   Begin VB.Label lCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "MediaPlayer About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   7
      Top             =   315
      Width           =   2085
   End
   Begin VB.Label cmdOK 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3915
      TabIndex        =   6
      Top             =   2565
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "tuniks@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1440
      MouseIcon       =   "frmAboutRemote.frx":3048
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2520
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "   If you have found errors in this program, and also your sentences and wishes can be sent me to the address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   630
      TabIndex        =   4
      Top             =   1710
      Width           =   3930
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(20 November 2000)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1620
      TabIndex        =   1
      Top             =   1485
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "CyrillicGoth"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   630
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   2160
      TabIndex        =   3
      Top             =   1215
      Width           =   825
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Font.Size = 10
Label2.ForeColor = vbBlue
End Sub

Private Sub Form_Load()
On Error Resume Next
    lCaption.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = Format(Now, "dd-mmmm-yyyy")
Set MouseIcon = Remote.ImageList1.ListImages(15).Picture
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
On Error Resume Next
Set MouseIcon = Remote.ImageList1.ListImages(16).Picture
ReleaseCapture
SendMessage hwnd, WM_SYSCOMMAND, SC_MOVE, 0
Set MouseIcon = Remote.ImageList1.ListImages(15).Picture
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Font.Size = 10
Label2.ForeColor = vbBlue
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Font.Size = 10
Label2.ForeColor = vbBlue
End Sub

Private Sub Label2_Click()
frmSendMail.Show
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Font.Size = 12
Label2.ForeColor = vbRed
End Sub

Private Sub Timer1_Timer()
lTime.Caption = Time
End Sub
