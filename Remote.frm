VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EDF439C0-99E5-11CF-AFF3-004005100200}#6.0#0"; "PVMARQ.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Remote 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   9015
   ControlBox      =   0   'False
   Icon            =   "Remote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Remote.frx":0CCA
   MousePointer    =   99  'Custom
   OLEDropMode     =   1  'Manual
   Picture         =   "Remote.frx":0E1C
   ScaleHeight     =   4845
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MediaPlayer2.PBarY PBarY3 
      Height          =   60
      Left            =   150
      TabIndex        =   36
      ToolTipText     =   "The Balance control."
      Top             =   3285
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   106
      Value           =   50
      Min             =   1
      Max             =   99
      BackColor       =   128
      FillColor       =   255
      BorderColor     =   16777215
      picForeColor    =   49152
   End
   Begin MediaPlayer2.PBarY PBarY2 
      Height          =   195
      Left            =   150
      TabIndex        =   35
      ToolTipText     =   "The Volume control."
      Top             =   2970
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   344
      Value           =   650
      Max             =   1000
      BackColor       =   0
      BorderColor     =   0
      MouseIcon       =   "Remote.frx":92A8
      MousePointer    =   99
      picFillColor    =   65280
      picStep         =   40
      Style           =   1
   End
   Begin MediaPlayer2.LButton cmdPlay 
      Height          =   465
      Left            =   270
      TabIndex        =   31
      ToolTipText     =   "Play"
      Top             =   1800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
      Caption         =   "Play"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LikeColor       =   7
      CapColor        =   32768
   End
   Begin PVMarqueeLib.PVMarquee Label7 
      Height          =   330
      Left            =   270
      TabIndex        =   28
      ToolTipText     =   "Îòîáðàæåíèå èìåíè òåêóùåãî ìóëüòèìåäèéíîãî ôàéëà."
      Top             =   315
      Width           =   2625
      _Version        =   393216
      _ExtentX        =   4630
      _ExtentY        =   582
      _StockProps     =   29
      Text            =   "SX MediaPlayer 2.0  (WAV, MID, MP3, MP2, AVI, CDA, MOV, WMA)"
      ForeColor       =   65280
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "EuroStyleDiai"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Frame           =   0
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Text            =   "SX MediaPlayer 2.0  (WAV, MID, MP3, MP2, AVI, CDA, MOV, WMA)"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4470
      Left            =   3240
      TabIndex        =   25
      Top             =   180
      Width           =   5595
      Begin VB.DirListBox Dir1 
         Height          =   540
         Left            =   90
         TabIndex        =   4
         ToolTipText     =   "Âûáîð ïàïêè äëÿ ïîèñêà."
         Top             =   3870
         Width           =   5370
      End
      Begin MediaPlayer2.LButton cmdMnu 
         Height          =   345
         Index           =   0
         Left            =   2565
         TabIndex        =   6
         ToolTipText     =   "New List."
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Caption         =   "LButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         Style           =   1
      End
      Begin MediaPlayer2.LButton cmdAddAll 
         Height          =   375
         Left            =   2250
         TabIndex        =   14
         ToolTipText     =   "Add all files in List."
         Top             =   1260
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "Add All"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         CapColor        =   12582912
      End
      Begin MediaPlayer2.LButton cmdAdd 
         Height          =   375
         Left            =   2250
         TabIndex        =   13
         ToolTipText     =   "Add file in List."
         Top             =   810
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "Add"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         CapColor        =   12582912
      End
      Begin MSComDlg.CommonDialog CommD 
         Left            =   405
         Top             =   2925
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   2250
         TabIndex        =   3
         ToolTipText     =   "Âûáîð óñòðîéñòâà."
         Top             =   3375
         Width           =   960
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   1080
         Top             =   1890
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   15000
         Left            =   1035
         Top             =   2745
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3825
         Top             =   2340
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   16
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":940A
               Key             =   "ExpandUp"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":9566
               Key             =   "ExpandDown"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":96C2
               Key             =   "Play"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":97CE
               Key             =   "Stop"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":98DA
               Key             =   "Pause"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":99E6
               Key             =   "Rewind"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":9AF2
               Key             =   "New"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":A042
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":A592
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":AAE2
               Key             =   "Del"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":ABF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":AD02
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":AE0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":AF1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":B026
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Remote.frx":B18A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.FileListBox File1 
         DragIcon        =   "Remote.frx":B2EE
         Height          =   3015
         Left            =   90
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   765
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.CommandButton cmdDn 
         Enabled         =   0   'False
         Height          =   280
         Left            =   2925
         Picture         =   "Remote.frx":B730
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ïåðåìåñòèòü âûáðàííûé ýëåìåíò âíèç."
         Top             =   2970
         Width           =   285
      End
      Begin VB.CommandButton cmdUp 
         Enabled         =   0   'False
         Height          =   280
         Left            =   2925
         Picture         =   "Remote.frx":BC72
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ïåðåìåñòèòü âûáðàííûé ýëåìåíò ââåðõ."
         Top             =   2610
         Width           =   285
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   90
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Text            =   "Combo1"
         ToolTipText     =   "Èìÿ èñïîëíèòåëÿ èëè äðóãàÿ èíôîðìàöèÿ."
         Top             =   270
         Width           =   2085
      End
      Begin VB.ListBox List1 
         Height          =   2985
         ItemData        =   "Remote.frx":BD74
         Left            =   3285
         List            =   "Remote.frx":BD76
         MultiSelect     =   2  'Extended
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         ToolTipText     =   "Ôàéëû çàíåñ¸ííûå â ñïèñîê äëÿ âîñïðîèçâåäåíèÿ."
         Top             =   765
         Width           =   2175
      End
      Begin VB.ListBox List3 
         DragIcon        =   "Remote.frx":BD78
         Height          =   2985
         Left            =   3285
         TabIndex        =   27
         Top             =   765
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ListBox List2 
         Height          =   2985
         ItemData        =   "Remote.frx":C1BA
         Left            =   90
         List            =   "Remote.frx":C1BC
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   26
         ToolTipText     =   "Íàéäåííûå ìóëüòèìåäèéíûå ôàéëû."
         Top             =   765
         Width           =   2085
      End
      Begin VB.CheckBox chkMsg 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4410
         TabIndex        =   29
         Top             =   2925
         Visible         =   0   'False
         Width           =   255
      End
      Begin MediaPlayer2.LButton cmdMnu 
         Height          =   345
         Index           =   1
         Left            =   2925
         TabIndex        =   7
         ToolTipText     =   "Open List."
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Caption         =   "LButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         Style           =   1
      End
      Begin MediaPlayer2.LButton cmdMnu 
         Height          =   345
         Index           =   2
         Left            =   3285
         TabIndex        =   8
         ToolTipText     =   "Save List as..."
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Caption         =   "LButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         Style           =   1
      End
      Begin MediaPlayer2.LButton cmdMnu 
         Height          =   345
         Index           =   3
         Left            =   3645
         TabIndex        =   9
         ToolTipText     =   "Delete List from windows."
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Caption         =   "LButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         Style           =   1
      End
      Begin MediaPlayer2.LButton cmdMnu 
         Height          =   345
         Index           =   4
         Left            =   4365
         TabIndex        =   10
         ToolTipText     =   "Play selected file."
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Caption         =   "LButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         Style           =   1
         Enabled         =   0   'False
      End
      Begin MediaPlayer2.LButton cmdMnu 
         Height          =   345
         Index           =   5
         Left            =   4725
         TabIndex        =   11
         ToolTipText     =   "Stop playing."
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Caption         =   "LButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         Style           =   1
         Enabled         =   0   'False
      End
      Begin MediaPlayer2.LButton cmdMnu 
         Height          =   345
         Index           =   6
         Left            =   5085
         TabIndex        =   12
         ToolTipText     =   "Play forward file."
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Caption         =   "LButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         Style           =   1
         Enabled         =   0   'False
      End
      Begin MediaPlayer2.PBarY PBarY1 
         Height          =   105
         Left            =   0
         TabIndex        =   30
         ToolTipText     =   "Èíäèêàòîð ïðîãðåññà è ðó÷íàÿ ïðîêðóòêà."
         Top             =   0
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   185
         BackColor       =   0
         MouseIcon       =   "Remote.frx":C1BE
         MousePointer    =   99
         picForeColor    =   16384
         picFillColor    =   65280
         picStep         =   17
         Style           =   1
      End
      Begin MediaPlayer2.LButton cmdRemove 
         Height          =   375
         Left            =   2250
         TabIndex        =   15
         ToolTipText     =   "Remove selected file."
         Top             =   1710
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "Remove"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         CapColor        =   255
      End
      Begin MediaPlayer2.LButton cmdRemoveAll 
         Height          =   375
         Left            =   2250
         TabIndex        =   16
         ToolTipText     =   "Remove all files."
         Top             =   2160
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "Remove All"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LikeColor       =   4
         CapColor        =   255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4050
      Top             =   2745
   End
   Begin MediaPlayer2.LButton cmdStop 
      Height          =   465
      Left            =   1710
      TabIndex        =   32
      ToolTipText     =   "Stop"
      Top             =   1800
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   820
      Caption         =   "Stop"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LikeColor       =   4
      CapColor        =   32768
   End
   Begin MediaPlayer2.LButton cmdRewLeft 
      Height          =   465
      Left            =   270
      TabIndex        =   33
      ToolTipText     =   "Step back."
      Top             =   2340
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
      Caption         =   "<<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LikeColor       =   4
      CapColor        =   32768
   End
   Begin MediaPlayer2.LButton cmdRewRight 
      Height          =   465
      Left            =   1710
      TabIndex        =   34
      ToolTipText     =   "Step forward."
      Top             =   2340
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   820
      Caption         =   ">>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LikeColor       =   4
      CapColor        =   32768
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FFFFFF&
      X1              =   1560
      X2              =   1560
      Y1              =   3330
      Y2              =   3420
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   480
      Left            =   1710
      Shape           =   4  'Rounded Rectangle
      Top             =   2340
      Width           =   1170
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   480
      Left            =   270
      Shape           =   4  'Rounded Rectangle
      Top             =   2340
      Width           =   1170
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   480
      Left            =   270
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   1170
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Height          =   480
      Left            =   1710
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   1170
   End
   Begin VB.Label lHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2970
      TabIndex        =   37
      Top             =   0
      Width           =   195
   End
   Begin VB.Line Line17 
      BorderColor     =   &H0000FF00&
      X1              =   2925
      X2              =   2925
      Y1              =   225
      Y2              =   1680
   End
   Begin VB.Line Line16 
      BorderColor     =   &H0000FF00&
      X1              =   225
      X2              =   225
      Y1              =   240
      Y2              =   1665
   End
   Begin VB.Line Line15 
      BorderColor     =   &H0000FF00&
      X1              =   225
      X2              =   2925
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   225
      X2              =   3015
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   270
      X2              =   2925
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   930
      Left            =   225
      Top             =   3555
      Width           =   2760
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   3015
      X2              =   3015
      Y1              =   3555
      Y2              =   4500
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   2925
      X2              =   2925
      Y1              =   3555
      Y2              =   4410
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   2925
      X2              =   270
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   270
      X2              =   270
      Y1              =   3600
      Y2              =   4410
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   180
      X2              =   180
      Y1              =   3510
      Y2              =   4455
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   315
      OLEDropMode     =   1  'Manual
      Picture         =   "Remote.frx":C320
      Stretch         =   -1  'True
      ToolTipText     =   "The context  Menu and OLE the receiver. DblClick - open List."
      Top             =   3645
      Width           =   2595
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   1665
      TabIndex        =   0
      ToolTipText     =   "Click - change View."
      Top             =   810
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1665
      TabIndex        =   24
      Top             =   1440
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1620
      TabIndex        =   23
      Top             =   1215
      Width           =   585
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50:50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   2235
      TabIndex        =   22
      Top             =   1215
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "650"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   2250
      TabIndex        =   21
      Top             =   1395
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1575
      X2              =   1575
      Y1              =   2835
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   150
      X2              =   3000
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   150
      X2              =   3000
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      X1              =   1575
      X2              =   1575
      Y1              =   1650
      Y2              =   750
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FF00&
      X1              =   2925
      X2              =   225
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  0  of  0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   360
      TabIndex        =   20
      ToolTipText     =   "Êàêîé ôàéë ïðîèãðûâàåòñÿ è ñêîëüêî âñåãî â ñïèñêå."
      Top             =   1305
      Width           =   1035
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0000FF00&
      X1              =   225
      X2              =   2925
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Notify"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   180
      TabIndex        =   19
      ToolTipText     =   "Òèï ìóëüòèìåäèéíîãî ôàéëà."
      Top             =   855
      Width           =   1365
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0000FF00&
      X1              =   225
      X2              =   2925
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu L1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuSysTray 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuList 
         Caption         =   "&List Show"
      End
      Begin VB.Menu L2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "At start to Hide"
      End
      Begin VB.Menu mnuAllFiles 
         Caption         =   "Show all files"
      End
      Begin VB.Menu mnuFullScreen 
         Caption         =   "FullScreen"
      End
      Begin VB.Menu L3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Unload"
      End
   End
End
Attribute VB_Name = "Remote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Number_MM As Integer
Dim ModeTime As Integer     ' Ìîäåëü îòîáðàæåíèÿ ÷àñîâ
Dim V As iVol               ' Ïåðåìåííàÿ äëÿ Volume
Dim Vol As tVol
Dim ArgArray() As String
Dim CommandLineEmpty As Boolean, FileType As String

Const Caption_Const = "SX MediaPlayer 2.0(WAV, SND, AU, AIF, AIFC, AIFF," & _
" MID, RMI, MP3, M3U, M1V, MP2, MPA, MPE, MPEG, ASF, ASX, MOV, QT, RA," & _
"RM, RAM, RMM, AVI, DAT)"
Const m_FileType = "*.wav;*.snd;*.au;*.aif;*.aifc;*.aiff;*.mid;*.rmi;*.mp3;" & _
"*.m3u;*.m1v;*.mp2;*.mpa;*.mpe;*.mpeg;*.asf;*.asx;*.mov;*.qt;*.ra;*.rm;*.ram;" & _
"*.rmm;*.avi;*.dat"

Private Type iVol
   RightV As Integer
   LeftV As Integer
End Type
Private Type tVol
    tempVol As Long
End Type

Const VolumeStep = 65       ' 65535/1000 (&HFFFF/&H3E8=&H41)
Const Wait_Timer = 15000    '15 Sec
Const Max_Width = 9045 '9180
Const Min_Width = 3260 '3300
Const const_Height = 4835 '5175

Private Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long

Private Sub chkMsg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim wParam As Long
    
    'X contains the wParam from the Windows messaging system
    'when a mouse event occurs over the System Tray Icon.
    'Since the checkbox coordinates are in twips, VB has
    'already converted the message by multiplying the X by
    'Screen.TwipsPerPixelX, so convert it back.
    wParam = X / Screen.TwipsPerPixelX
    
    'This is only using Double-Click and Right-Click,
    'but all of the following events are returned.
    Select Case wParam
        'Case WM_MOUSEMOVE
        'Case WM_LBUTTONDOWN
        'Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
            mnuShow_Click
        Case WM_RBUTTONDOWN
            PopupMenu mnuFile, , , , mnuShow
        'Case WM_RBUTTONUP
        'Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub cmdPlay_Click()
If MM.Control1.PlayState = mpPlaying Then Exit Sub
If List1.ListCount = List1.ListIndex + 1 Then MM.Control1.Stop: Exit Sub
On Error Resume Next
    List1.ListIndex = List1.ListIndex + 1
    List3.ListIndex = List1.ListIndex
    List1.SetFocus
Static Name As String
With MM.Control1
     Name = List3.Text & List1.Text
     If Name = "" Then Exit Sub
    .Stop
If Number_MM <> 0 Then Number_MM = List1.ListIndex + 1
    .FileName = Name
            PBarY1.Max = .Duration
     Label7.Text = Name
SetSysTrayIcon ModifyIcon, chkMsg.hwnd, Me.Icon, Name
     Label10.Caption = UCase(Right(Name, 3))
    Timer1.Enabled = False
        If Label10.Caption = "DAT" Or Label10.Caption = "AVI" _
        Or Label10.Caption = "MOV" Or Label10.Caption = "MPE" Then
            MM.Show
        '.AutoStart = False
        '.EnableContextMenu = False
        '.SendMouseClickEvents = True
        '.SendMouseMoveEvents = True
        '.ShowPositionControls = False
        .EnableFullScreenControls = mnuFullScreen.Checked
        Else
           MM.Hide
        End If
    .Open (Name)
    Do Until .PlayState = mpPlaying
    .Play
    DoEvents
    Loop
    Timer1.Enabled = True
End With
VolumeUpDown_Change
End Sub

Private Sub cmdRemove_Click()
Static L As Integer, i As Integer
If List1.List(List1.ListIndex) = "" Then MM.Control1.Stop: Number_MM = 0: _
List1.ListIndex = -1: Label7.Text = Caption_Const: Exit Sub
On Error GoTo CountEndRemove
        For i = 0 To List1.ListCount
            If List1.Selected(i) Then
    L = InStrRev(List1.List(i), "\")
        L = Len(List1.List(i)) - L
                List2.AddItem Right(List1.List(i), L)
                    List3.RemoveItem (i)
                        List1.RemoveItem (i)
                i = i - 1
            End If
        Next i
CountEndRemove:
End Sub

Private Sub cmdRemoveAll_Click()
ButtonEnabled (False)
Set cmdMnu(4).Picture = ImageList1.ListImages(3).Picture
List3.Clear
List1.Clear
File1_PathChange
MM.Control1.Stop
Number_MM = 0
List1.ListIndex = -1
Label7.Text = Caption_Const
End Sub

Private Sub cmdMnu_Click(Index As Integer)
    Static i As Integer, NextLine As String
Select Case Index
    Case 0
Dir1.Path = "C:\"
cmdRemoveAll_Click
Combo1.Text = ""
    Case 1
            If CommandLineEmpty Then
With CommD
cmdRemoveAll_Click
    .Filter = "File catalog|*.mpc"
    .FilterIndex = 1
    .DialogTitle = "Open file Catalog"
    .CancelError = False
    .ShowOpen
End With
            End If
If CommD.FileName = Empty Then Exit Sub
    Open CommD.FileName For Input As #1
        ButtonEnabled (True)
            Line Input #1, NextLine
                Combo1.Text = NextLine
    Do While Not EOF(1)
        If CommandLineEmpty Then DoEvents
    Line Input #1, NextLine
        If NextLine = "" Then GoTo M1
            i = InStrRev(NextLine, "\")
                List1.AddItem Right(NextLine, Len(NextLine) - i)
        List3.AddItem Left(NextLine, i)
    Loop
M1: Close #1
List1.ListIndex = -1
On Error Resume Next
Dir1.Path = List3.List(0)
        If Err.Number = 68 Then MsgBox Err.Description, vbQuestion + vbCritical, "Attention, attention!"
    Case 2
With CommD
    .Filter = "File catalog|*.mpc"
    .FilterIndex = 1
    .DialogTitle = "Save file Catalog"
    .CancelError = False
    .ShowSave
If .FileName = Empty Then Exit Sub
Open .FileName For Output As #1
End With
Print #1, Combo1.Text
For i = 0 To List1.ListCount
Print #1, List3.List(i) & List1.List(i)
Next i
Close #1
    Case 3
cmdRemoveAll_Click
    Case 4
GoSub MM_Play
Do While MM.Control1.PlayState = mpPlaying
DoEvents
Loop
MM.Control1.Stop
    Case 5
    MM.Control1.Stop
    MM.Control1.FileName = ""
    Case 6
If List1.ListCount = List1.ListIndex + 1 Then MM.Control1.Stop: Exit Sub
    List1.ListIndex = List1.ListIndex + 1
    List3.ListIndex = List1.ListIndex
    List1.SetFocus
    GoSub MM_Play
Do While MM.Control1.PlayState = mpPlaying
DoEvents
Loop
MM.Control1.Stop
    End Select
    Exit Sub
MM_Play:
Static Name As String
With MM.Control1
     Name = List3.Text & List1.Text
     If Name = "" Then Exit Sub
    .Stop
     Number_MM = List1.ListIndex + 1
    .FileName = Name
     Label7.Text = Name
     SetSysTrayIcon ModifyIcon, chkMsg.hwnd, Me.Icon, Name
     Label10.Caption = UCase(Mid(Name, InStr(1, Name, ".") + 1))
    Timer1.Enabled = False
    PBarY1.Max = .Duration
        If Label10.Caption = "DAT" Or Label10.Caption = "AVI" _
        Or Label10.Caption = "MOV" Or Label10.Caption = "MPE" Then
            MM.Show
        '.AutoStart = False
        '.EnableContextMenu = False
        '.SendMouseClickEvents = True
        '.SendMouseMoveEvents = True
        '.ShowPositionControls = False
        .EnableFullScreenControls = mnuFullScreen.Checked
        Else
        MM.Hide
        End If
    .Open (Name)
On Error Resume Next
    Do Until .PlayState = mpPlaying
    .Play
    DoEvents
    Loop
    Timer1.Enabled = True
End With
Return
End Sub

Private Sub cmdUp_Click()
Static str As String, i As Integer

i = List1.ListIndex
If i = 0 Then Exit Sub
str = List1.List(i - 1)
List1.List(i - 1) = List1.List(i)
List1.List(i) = str
str = List3.List(i - 1)
List3.List(i - 1) = List3.List(i)
List3.List(i) = str
List1.ListIndex = i - 1
End Sub

Private Sub cmdDn_Click()
Static str As String, i As Integer

i = List1.ListIndex
If i = List1.ListCount - 1 Then Exit Sub
str = List1.List(i + 1)
List1.List(i + 1) = List1.List(i)
List1.List(i) = str
str = List3.List(i + 1)
List3.List(i + 1) = List3.List(i)
List3.List(i) = str
List1.ListIndex = i + 1
End Sub

Private Sub cmdStop_Click()
MM.Control1.Stop
MM.Control1.FileName = ""
End Sub

Private Sub cmdRewLeft_Click()
On Error Resume Next
If List1.ListIndex <> 0 Then List1.ListIndex = List1.ListIndex - 2: cmdRewRight_Click
End Sub

Private Sub cmdRewRight_Click()
cmdMnu_Click (6)
End Sub

Private Sub Combo1_Click()
MousePointer = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MousePointer = 14 Then Exit Sub
MousePointer = 99
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Set Me.MouseIcon = ImageList1.ListImages(16).Picture
ReleaseCapture
SendMessage hwnd, WM_SYSCOMMAND, SC_MOVE, 0
Set Me.MouseIcon = ImageList1.ListImages(15).Picture
End Sub

Private Sub Image2_DblClick()
mnuList_Click
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 12
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 12
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 12
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 12
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 12
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 12
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 12
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 2
End Sub

Private Sub lHelp_Click()
Unload Me
End Sub

Public Sub mnuSysTray_Click()
Me.Move Left, Top, Min_Width, Height
'Set cmdList.Picture = ImageList1.ListImages(1).Picture
mnuList.Visible = False
mnuSysTray.Visible = False
Timer2.Enabled = False
SetSysTrayIcon AddIcon, chkMsg.hwnd, Me.Icon, Label7.Text
Me.Hide
'ImPruv.Visible = False
End Sub

Private Sub mnuList_Click()
    Dir1.Move Dir1.Left, 3870, Dir1.Width, 540
If Me.Width <> Max_Width Then
Me.Move Left, Top, Max_Width, Height
'Set cmdList.Picture = ImageList1.ListImages(2).Picture
mnuList.Caption = "List UnShow"
Timer_Refresh
'ImPruv.Visible = True
Else
Me.Move Left, Top, Min_Width, Height
'Set cmdList.Picture = ImageList1.ListImages(1).Picture
mnuList.Caption = "List Show"
Timer2.Enabled = False
Timer3.Enabled = False
'ImPruv.Visible = False
End If
End Sub

Private Sub cmdAdd_Click()
    Dir1.Move Dir1.Left, 3870, Dir1.Width, 540
If List2.ListIndex = -1 Then Exit Sub
ButtonEnabled (True)
Static i As Integer
On Error GoTo CountEnd
Err.Number = 0
    For i = 0 To List2.ListCount - 1
        If List2.Selected(i) Then
 If Right(File1.Path, 1) <> "\" Then
        List1.AddItem List2.List(i)
    List3.AddItem File1.Path & "\"
        List2.RemoveItem (i)
        i = i - 1
 Else
    List1.AddItem List2.List(i)
        List3.AddItem File1.Path
    List2.RemoveItem (i)
    i = i - 1
 End If
        End If
    Next i
List1.SetFocus
Exit Sub

CountEnd:
    If Err.Number = 381 Then
        List1.SetFocus
    End If
End Sub

Private Sub cmdAddAll_Click()
    Dir1.Move Dir1.Left, 3870, Dir1.Width, 540
ButtonEnabled (True)
Dim i As Integer
For i = 0 To List2.ListCount - 1
If Right(File1.Path, 1) <> "\" Then
List1.AddItem List2.List(0)
List3.AddItem File1.Path & "\"
List2.RemoveItem (0)
Else
List1.AddItem List2.List(0)
List3.AddItem File1.Path
List2.RemoveItem (0)
End If
Next i
List1.SetFocus
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Refresh
End Sub

Private Sub Dir1_Click()
    Dir1.Move Dir1.Left, 1710, Dir1.Width, 540 * 5
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
Timer_Refresh
If Timer3.Enabled Then Me.Move Left, Top, Max_Width, Height: Timer3.Enabled = False

End Sub

Private Sub Drive1_Change()
On Error GoTo No_Ready_device
Dir1.Path = Drive1.Drive
Dir1.Refresh
File1.Path = Dir1.Path
File1.Refresh
Exit Sub
No_Ready_device:
    MsgBox Err.Description, vbQuestion + vbCritical, "Attention, attention!"
Drive1.Drive = Dir1.Path
End Sub

Private Sub Drive1_GotFocus()

Drive1.Move Drive1.Left - 400, Drive1.Top, Drive1.Width + 800
End Sub

Private Sub Drive1_LostFocus()
Drive1.Move Drive1.Left + 400, Drive1.Top, Drive1.Width - 800
End Sub

Private Sub File1_Click()
Timer_Refresh
End Sub

Private Sub File1_PathChange()
Static i As Integer

List2.Clear
    For i = 0 To File1.ListCount - 1
    List2.AddItem File1.List(i)
    Next i
End Sub

Private Sub Form_Load()
FileType = m_FileType
Label7.Text = Caption_Const
On Error Resume Next
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = Min_Width
    Me.Height = const_Height
    Dir1.Path = GetSetting(App.Title, "Settings", "Path", "C:\")

atStart = GetSetting(App.Title, "Settings", "Hide", "0")
mnuHide.Checked = atStart

mnuAllFiles.Checked = GetSetting(App.Title, "Settings", "Filter", "0")
mnuFullScreen.Checked = GetSetting(App.Title, "Settings", "Screen", "1")

PBarY2.Value = GetSetting(App.Title, "Settings", "Volume", 650)
PBarY3.Value = GetSetting(App.Title, "Settings", "Balance", 50)

Combo1.Text = "New Artist"
File1.Pattern = FileType
Dir1_Change
File1_PathChange

Set cmdMnu(0).Picture = ImageList1.ListImages(7).Picture
Set cmdMnu(1).Picture = ImageList1.ListImages(8).Picture
Set cmdMnu(2).Picture = ImageList1.ListImages(9).Picture
Set cmdMnu(3).Picture = ImageList1.ListImages(10).Picture
Set cmdMnu(4).Picture = ImageList1.ListImages(3).Picture
Set cmdMnu(5).Picture = ImageList1.ListImages(4).Picture
Set cmdMnu(6).Picture = ImageList1.ListImages(6).Picture
Me.Show
cmdPlay.SetFocus
List1.ListIndex = -1

CommandLineEmpty = True
GetCommandLine 1
If ArgArray(1) = Empty Then Exit Sub
'------------ Íà÷èíàåòñÿ îáðàáîòêà êîììàíäíîé ñòðîêè-------------
Complete_Command_Line ArgArray(1)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Set Me.MouseIcon = ImageList1.ListImages(16).Picture
ReleaseCapture
SendMessage hwnd, WM_SYSCOMMAND, SC_MOVE, 0
Set Me.MouseIcon = ImageList1.ListImages(15).Picture
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Data.GetFormat(vbCFFiles) Then Exit Sub
On Error Resume Next
Complete_Command_Line Data.Files(1)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "Path", Dir1.Path
        SaveSetting App.Title, "Settings", "Volume", PBarY2.Value
        SaveSetting App.Title, "Settings", "Hide", atStart
        SaveSetting App.Title, "Settings", "Filter", mnuAllFiles.Checked
        SaveSetting App.Title, "Settings", "Screen", mnuFullScreen.Checked
        SaveSetting App.Title, "Settings", "Balance", PBarY3.Value
    End If
     If MsgBox("You are sure?", vbQuestion + vbYesNo, "Menu Quit") = vbYes Then
            Cancel = False
 SetSysTrayIcon DeleteIcon, chkMsg.hwnd, Me.Icon, Caption_Const
            End
        Else
            Cancel = True
    End If
End Sub

Private Sub Form_Resize()
'Me.Caption = "SX MediaPlayer 2.0"
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 99
    Dir1.Move Dir1.Left, 3870, Dir1.Width, 540
        Timer_Refresh
    If Timer3.Enabled Then Me.Move Left, Top, Max_Width, Height: Timer3.Enabled = False
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then Exit Sub
mnuShow.Visible = False
L1.Visible = False
PopupMenu mnuFile, , , , mnuList
L1.Visible = True
mnuShow.Visible = True
End Sub

Private Sub Image2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Label9_Click()
Select Case ModeTime
Case 0              ' Òåêóùåå âðåìÿ
ModeTime = 1
Case 1              ' Îñòàòîê âðåìåíè ïðîèãðûâàíèÿ â Ñåê
ModeTime = 2
Case 2              ' Îñòàòîê âðåìåíè ïðîèãðûâàíèÿ â ôðàãìåíòàõ
ModeTime = 0
End Select
cmdStop.SetFocus
End Sub

Private Sub List1_Click()
    Dir1.Move Dir1.Left, 3870, Dir1.Width, 540
Set cmdMnu(4).Picture = ImageList1.ListImages(12).Picture
List3.ListIndex = List1.ListIndex
End Sub

Private Sub List1_DblClick()
    Dir1.Move Dir1.Left, 3870, Dir1.Width, 540
cmdRemove_Click
End Sub

Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
cmdAdd_Click
End Sub

Private Sub List1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
If State = 0 Then
List1.DragIcon = File1.DragIcon
Else
List2.DragIcon = List3.DragIcon
End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
    Dir1.Move Dir1.Left, 3870, Dir1.Width, 540
    Timer_Refresh
        If Timer3.Enabled Then Me.Move Left, Top, Max_Width, Height: Timer3.Enabled = False
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Data.GetFormat(vbCFFiles) Then Exit Sub
On Error Resume Next
Static i As Integer, L As Integer, Ext As String
i = 1
Do
  Err.Clear
    Ext = UCase(Right(Data.Files(i), 3))
         If Err.Source = App.Title Then Exit Sub
Select Case Ext
    Case "MPC"
MM.Control1.Stop
CommandLineEmpty = False
CommD.FileName = Data.Files(i)
cmdMnu_Click (1)
CommandLineEmpty = True
cmdPlay_Click
    Case "WAV", "SND", "AU", "AIF", "MID", "RMI", "MP3", "M3U", "M1V", _
    "MP2", "MPA", "MPE", "ASF", "ASX", "MOV", "RAM", "RMM", "AVI", "DAT"
    L = InStrRev(Data.Files(i), "\")
        List3.AddItem Left(Data.Files(i), L)
            L = Len(Data.Files(i)) - L
                List1.AddItem Right(Data.Files(i), L)
    Label7.Text = Data.Files(i)
End Select
i = i + 1
Loop
End Sub

Private Sub List2_Click()
    Dir1.Move Dir1.Left, 3870, Dir1.Width, 540
End Sub

Private Sub List2_DblClick()
cmdAdd_Click
    Dir1.Move Dir1.Left, 3870, Dir1.Width, 540
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static i As Integer
For i = 0 To List2.ListCount - 1
If List2.Selected(i) Then GoTo GragList
Next i
Exit Sub

GragList: List2.DragIcon = File1.DragIcon
    List2.Drag
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
    Dir1.Move Dir1.Left, 3870, Dir1.Width, 540
        Timer_Refresh
    If Timer3.Enabled Then Me.Move Left, Top, Max_Width, Height: Timer3.Enabled = False
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuAllFiles_Click()
mnuAllFiles.Checked = Not mnuAllFiles.Checked
If mnuAllFiles.Checked Then
FileType = "*.*"
Else
FileType = m_FileType
End If
File1.Pattern = FileType
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFullScreen_Click()
mnuFullScreen.Checked = Not mnuFullScreen.Checked
If Remote.mnuFullScreen Then
MM.Control1.DisplaySize = mpFullScreen
Else
MM.Control1.DisplaySize = mpDefaultSize
End If
MM.Control1.AllowChangeDisplaySize = True
MM.Control1.AutoStart = False
MM.Control1.ClickToPlay = True
End Sub

Private Sub mnuHide_Click()
mnuHide.Checked = Not mnuHide.Checked
atStart = mnuHide.Checked
End Sub

Private Sub mnuShow_Click()
Me.Show
SetSysTrayIcon DeleteIcon, chkMsg.hwnd, Me.Icon, Caption_Const
mnuList.Visible = True
mnuSysTray.Visible = True
End Sub

Private Sub PBarY1_ChangeValue(NewValue As Long, OldValue As Long)
If MM.Control1.PlayState = mpPlaying Then
If NewValue = OldValue Then Exit Sub
MM.Control1.CurrentPosition = NewValue
End If
End Sub

Private Sub PBarY2_ChangeValue(NewValue As Long, OldValue As Long)
Label1.Caption = NewValue
VolumeUpDown_Change
End Sub

Private Sub PBarY3_ChangeValue(NewValue As Long, OldValue As Long)
    Label4.Caption = str(100 - NewValue) & ":" & Mid(str(NewValue), 2)
VolumeUpDown_Change
End Sub

Private Sub Timer1_Timer()
GetVolume_Balance
If ModeTime = 0 Then
Label9.Caption = Format(Now, "hh:mm:ss")
ElseIf ModeTime = 1 Then
MM.Control1.DisplayMode = mpTime
Label9.Caption = Format((MM.Control1.Duration - MM.Control1.CurrentPosition), "0.000")
Else
MM.Control1.DisplayMode = mpFrames
Label9.Caption = Format(CDbl(MM.Control1.CurrentPosition), "0.000")
End If
    Label8.Caption = Format(List1.ListIndex + 1, "##0") & "  of  " & _
                 Format(List1.ListCount, "##0")
If List1.ListIndex = -1 Then Set cmdMnu(4).Picture = ImageList1.ListImages(3).Picture
If MM.Control1.PlayState = mpStopped Then cmdPlay_Click
If MM.Control1.PlayState = mpPlaying Then PBarY1.Value = MM.Control1.CurrentPosition
End Sub

Private Sub Timer_Refresh()
        Timer2.Enabled = False
    Timer2.Interval = Wait_Timer
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
        Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Me.Move Left, Top, Width - (Max_Width - Min_Width) / 60, Height
    If Me.Width <= Min_Width Then
        mnuList.Caption = "List Show"
        Me.Width = Min_Width
    Timer3.Enabled = False
  End If
End Sub

Private Sub ButtonEnabled(ByVal Priznak As Boolean)
cmdMnu(4).Enabled = Priznak
cmdMnu(5).Enabled = Priznak
cmdMnu(6).Enabled = Priznak
cmdUp.Enabled = Priznak
cmdDn.Enabled = Priznak
End Sub

Private Sub VolumeUpDown_Change()
Static SLong As Long, StepBalance As Integer

        SLong = PBarY2.Value * VolumeStep
        StepBalance = CInt(SLong / 50)
    
    If PBarY3.Value < 50 Then SLong = SLong - (50 - PBarY3.Value) * StepBalance
If SLong > 32768 Then V.LeftV = CInt(SLong + &HFFFF0000) Else V.LeftV = CInt(SLong)
        
        SLong = PBarY2.Value * VolumeStep
    If PBarY3.Value > 50 Then SLong = SLong - (PBarY3.Value - 50) * StepBalance
If SLong > 32768 Then V.RightV = CInt(SLong + &HFFFF0000) Else V.RightV = CInt(SLong)

LSet Vol = V
waveOutSetVolume 0, Vol.tempVol
midiOutSetVolume 0, Vol.tempVol
End Sub

Function GetCommandLine(Optional MaxArgs)
    'Declare variables.
    Static C As String, CmdLine As String, CmdLnLen As Integer
    Static InArg As Boolean, i As Integer, NumArgs As Integer
    'See if MaxArgs was provided.
    If IsMissing(MaxArgs) Then MaxArgs = 10
    'Make array of the correct size.
    ReDim ArgArray(MaxArgs)
    NumArgs = 0: InArg = False
    'Get command line arguments.
    CmdLine = Command()
    CmdLnLen = Len(CmdLine)
    'Go thru command line one character
    'at a time.
    For i = 1 To CmdLnLen
        C = Mid(CmdLine, i, 1)
        'Test for space or tab.
        If (C <> "/" And C <> vbTab) Then
            'Neither space nor tab.
            'Test if already in argument.
            If Not InArg Then
            'New argument begins.
            'Test for too many arguments.
                If NumArgs = MaxArgs Then Exit For
                NumArgs = NumArgs + 1
                InArg = True
            End If
            'Concatenate character to current argument.
            ArgArray(NumArgs) = ArgArray(NumArgs) & C
        Else
            'Found a space or tab.
            'Set InArg flag to False.
            InArg = False
        End If
    Next i
    'Resize array just enough to hold arguments.
    ReDim Preserve ArgArray(NumArgs)
    'Return Array in Function name.
    GetCommandLine = ArgArray()
End Function


Private Function Complete_Command_Line(ByVal DataFileName As String)
Static Ext As String

Ext = UCase(Right(DataFileName, 3))
If Ext = "MPC" Then
CommandLineEmpty = False
CommD.FileName = DataFileName
cmdRemoveAll_Click
cmdMnu_Click (1)
CommandLineEmpty = True
cmdPlay_Click
Else
  Select Case Ext
    Case "WAV", "SND", "AU", "AIF", "MID", "RMI", "MP3", "M3U", "M1V", _
    "MP2", "MPA", "MPE", "ASF", "ASX", "MOV", "RAM", "RMM", "AVI", "DAT"
    With MM.Control1
            Timer1.Enabled = False
    '    .Open DataFileName
    'Do Until .OpenState = mpClosed
    'DoEvents
    'Loop
    '        PBarY1.Value = .Duration
               If Ext = "AVI" Or Ext = "MOV" Or Ext = "DAT" Then
    MM.Show
        '.AutoStart = False
        '.EnableContextMenu = False
        '.SendMouseClickEvents = True
        '.SendMouseMoveEvents = True
        '.ShowPositionControls = False
        .EnableFullScreenControls = mnuFullScreen.Checked
                Else
    MM.Hide
                End If
    .Open DataFileName
    Do Until .OpenState = mpClosed
    DoEvents
    Loop
PBarY1.Max = .Duration
    Timer1.Enabled = True
 Label7.Text = DataFileName
     .Play
   End With
Timer1.Enabled = True
  End Select
End If
End Function

Private Sub GetVolume_Balance()
Static LeftV As Long, RightV As Long, RightBig As Boolean

On Error Resume Next

waveOutGetVolume 0, Vol.tempVol

LSet V = Vol

LeftV = CLng(V.LeftV): RightV = CLng(V.RightV)
    If LeftV < 0 Then LeftV = LeftV - &HFFFF0000
        If RightV < 0 Then RightV = RightV - &HFFFF0000
            Vol.tempVol = RightV: RightBig = True
                If RightV < LeftV Then Vol.tempVol = LeftV: RightBig = False
PBarY2.Value = Vol.tempVol / VolumeStep

    If RightBig Then
        PBarY3.Value = LeftV / RightV * 50
    Else
        PBarY3.Value = 100 - RightV / LeftV * 50
    End If
End Sub
