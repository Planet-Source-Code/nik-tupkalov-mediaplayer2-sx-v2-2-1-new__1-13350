VERSION 5.00
Begin VB.Form frmSendMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Mail"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox strNote 
      Height          =   2760
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   4290
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send &Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   2925
      Width           =   4290
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oMapi As Object
Dim S As Object
Dim R As Object
Dim errcod As Long, errtxt As String

Private Sub cmdSend_Click()
On Error GoTo errid
    If MsgBox("Send Mail?", vbYesNo) = vbYes Then
        oMapi.DownloadMail = False
            If Not oMapi.LogOn(errcod, errtxt) Then GoTo errid
                S.SendTo = "tuniks@hotmail.com"
    S.Subject = "MediaPlayer2 SX"
        S.Message = strNote.Text
            If Not S.SendMail(errcod, errtxt) Then GoTo errid
                If Not oMapi.LogOff(errcod, errtxt) Then GoTo errid
    End If
        Unload Me
            Exit Sub
errid:
        MsgBox errtxt & errcod, vbCritical
End Sub

Private Sub Form_Load()
Me.Show
    
    Set oMapi = CreateObject("MAPIMail.cMAPI")
    Set S = oMapi.Sender
    Set R = oMapi.Receiver

End Sub
