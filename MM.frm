VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form MM 
   Caption         =   "MediaPlayer 2.0 SX"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin MediaPlayerCtl.MediaPlayer Control1 
      Height          =   5040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6645
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   30
      BaseURL         =   ""
      BufferingTime   =   10
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
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
      EnableFullScreenControls=   -1  'True
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   -1  'True
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
      SendMouseClickEvents=   -1  'True
      SendMouseMoveEvents=   -1  'True
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuFullScreen 
         Caption         =   "Show Full Screen"
      End
      Begin VB.Menu mnuStandart 
         Caption         =   "Standart Resize"
      End
      Begin VB.Menu mnuMute 
         Caption         =   "Mute Volume"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties File"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Start/Stop Play"
      End
   End
End
Attribute VB_Name = "MM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Control1_MouseDown(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
If Button <> 2 Then Exit Sub
'PopupMenu mnuMain
End Sub

Private Sub Form_Load()
With Control1
        .AutoStart = False
        .DisplaySize = mpFitToSize
        .EnableContextMenu = False
        .SendMouseClickEvents = True
        .SendMouseMoveEvents = True
        .ShowPositionControls = False
        .EnableFullScreenControls = mnuFullScreen.Checked
End With
End Sub

Private Sub Form_Resize()
Control1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuFullScreen_Click()
If Control1.DisplaySize = mpDefaultSize Then Control1.DisplaySize = mpFullScreen
End Sub
