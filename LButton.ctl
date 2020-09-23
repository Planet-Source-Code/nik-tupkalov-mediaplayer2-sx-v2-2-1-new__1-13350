VERSION 5.00
Begin VB.UserControl LButton 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   Picture         =   "LButton.ctx":0000
   PropertyPages   =   "LButton.ctx":13CA
   ScaleHeight     =   1050
   ScaleWidth      =   2490
   ToolboxBitmap   =   "LButton.ctx":13ED
   Begin VB.Label LCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LButtom"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   945
      TabIndex        =   0
      Top             =   405
      Width           =   615
   End
   Begin VB.Image Img1 
      Height          =   1050
      Left            =   0
      Picture         =   "LButton.ctx":16FF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2490
   End
   Begin VB.Image Img2 
      Height          =   1050
      Index           =   7
      Left            =   0
      Picture         =   "LButton.ctx":2AC9
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Image Img2 
      Height          =   1050
      Index           =   6
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Image Img2 
      Height          =   1050
      Index           =   5
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Image Img2 
      Height          =   1050
      Index           =   4
      Left            =   0
      Picture         =   "LButton.ctx":44C7
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Image Img2 
      Height          =   1050
      Index           =   3
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Image Img2 
      Height          =   1050
      Index           =   2
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Image Img2 
      Height          =   1050
      Index           =   1
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Image LPicture 
      Height          =   465
      Left            =   900
      Top             =   315
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image Img2 
      Height          =   1050
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Image Img3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1110
      Left            =   0
      Picture         =   "LButton.ctx":5EC5
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Image Img4 
      Height          =   375
      Left            =   810
      Picture         =   "LButton.ctx":728F
      Top             =   315
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "LButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Dim lLeft As Integer, lTop As Integer

Public Enum LStyle
    Normal
    Grafics
       End Enum

Public Enum eLColor
    Grey
    Blue
    Magenta
    Ultra
    Wave
    Yellow
    Red
    Green
       End Enum
       
'Default Property Values:
Const m_def_Enabled = True
Const m_def_Style = 0
Const m_def_LikeColor = 0

'Property Variables:
Dim m_Enabled As Boolean
Dim m_Picture As Object
Dim m_Style As LStyle
Dim m_LikeColor As eLColor
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

Private Sub Img1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not m_Enabled Then Exit Sub
If Button <> 0 Then Exit Sub
Static StepY As Integer, StepX
StepX = ScaleWidth / 10
StepY = ScaleHeight / 5
If (X > StepX And X <= ScaleWidth - StepX) And (Y > StepY And Y < ScaleHeight - StepY) Then
Img1.Picture = Img2(m_LikeColor).Picture
Else
Img1.Picture = Img4.Picture
End If
End Sub

Private Sub Img1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not m_Enabled Then Exit Sub
If Button <> 1 Then Exit Sub
'Img1.Picture = Img3.Picture
Img1.BorderStyle = 1
If m_Style = Normal Then
LCaption.Move lLeft + 15, lTop + 15
Else
LPicture.Move lLeft + 15, lTop + 15
End If
End Sub

Private Sub Img1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not m_Enabled Then Exit Sub
If Button <> 1 Then Exit Sub
Img1.Picture = Img4.Picture
If m_Style = Normal Then
LCaption.Move lLeft, lTop
Else
LPicture.Move lLeft, lTop
End If
Img1.BorderStyle = 0
   RaiseEvent Click
End Sub

Private Sub LCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not m_Enabled Then Exit Sub
If Button <> 1 Then Exit Sub
Img1.Picture = Img3.Picture
Img1.BorderStyle = 1
LCaption.Move lLeft + 15, lTop + 15
End Sub

Private Sub LCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not m_Enabled Then Exit Sub
Img1.Picture = Img2(m_LikeColor).Picture
End Sub

Private Sub LCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not m_Enabled Then Exit Sub
If Button <> 1 Then Exit Sub
Img1.Picture = Img4.Picture
LCaption.Move lLeft, lTop
Img1.BorderStyle = 0
   RaiseEvent Click
End Sub

Private Sub LPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not m_Enabled Then Exit Sub
If Button <> 1 Then Exit Sub
Img1.Picture = Img3.Picture
Img1.BorderStyle = 1
LPicture.Move lLeft + 15, lTop + 15
End Sub

Private Sub LPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not m_Enabled Then Exit Sub
Img1.Picture = Img2(m_LikeColor).Picture
End Sub

Private Sub LPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not m_Enabled Then Exit Sub
If Button <> 1 Then Exit Sub
Img1.Picture = Img4.Picture
Img1.BorderStyle = 0
   RaiseEvent Click
LPicture.Move lLeft, lTop
End Sub

Private Sub UserControl_ExitFocus()
Img1.Picture = Img4.Picture
End Sub

Private Sub UserControl_Resize()
Img1.Width = ScaleWidth
Img1.Height = ScaleHeight
If m_Style = Normal Then
lLeft = (ScaleWidth - LCaption.Width) / 2
lTop = (ScaleHeight - LCaption.Height) / 2
LCaption.Move lLeft, lTop
Else
lLeft = (ScaleWidth - LPicture.Width) / 2
lTop = (ScaleHeight - LPicture.Height) / 2
LPicture.Move lLeft, lTop
End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCaption,LCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = LCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    LCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCaption,LCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = LCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set LCaption.Font = New_Font
    PropertyChanged "Font"
lLeft = (ScaleWidth - LCaption.Width) / 2
lTop = (ScaleHeight - LCaption.Height) / 2
LCaption.Left = lLeft
LCaption.Top = lTop
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Img1,Img1,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = Img1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Img1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbWhite
Public Property Get LikeColor() As eLColor
    LikeColor = m_LikeColor
End Property

Public Property Let LikeColor(ByVal New_LikeColor As eLColor)
    m_LikeColor = New_LikeColor
    PropertyChanged "LikeColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_LikeColor = m_def_LikeColor
    LCaption.Caption = Extender.Name
    m_Style = m_def_Style
    Set m_Picture = LoadPicture("")
    m_Enabled = m_def_Enabled
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    LCaption.Caption = PropBag.ReadProperty("Caption", "LButtom")
    Set LCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Img1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_LikeColor = PropBag.ReadProperty("LikeColor", m_def_LikeColor)
    LCaption.ForeColor = PropBag.ReadProperty("CapColor", &H80000008)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
End Sub

Private Sub UserControl_Show()
If m_Style = Grafics Then
LPicture.Visible = True
LCaption.Visible = False
LPicture.Move (ScaleWidth - LPicture.Width) / 2, (ScaleHeight - LPicture.Height) / 2
LPicture.ZOrder
Else
LPicture.Visible = False
LCaption.Visible = True
End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", LCaption.Caption, "LButtom")
    Call PropBag.WriteProperty("Font", LCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("ToolTipText", Img1.ToolTipText, "")
    Call PropBag.WriteProperty("LikeColor", m_LikeColor, m_def_LikeColor)
    Call PropBag.WriteProperty("CapColor", LCaption.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCaption,LCaption,-1,ForeColor
Public Property Get CapColor() As OLE_COLOR
Attribute CapColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    CapColor = LCaption.ForeColor
End Property

Public Property Let CapColor(ByVal New_CapColor As OLE_COLOR)
    LCaption.ForeColor() = New_CapColor
    PropertyChanged "CapColor"
lLeft = (ScaleWidth - LCaption.Width) / 2
lTop = (ScaleHeight - LCaption.Height) / 2
LCaption.Left = lLeft
LCaption.Top = lTop
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Style() As LStyle
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As LStyle)
    m_Style = New_Style
    PropertyChanged "Style"
If m_Style = Grafics Then
LPicture.Visible = True
LCaption.Visible = False
LPicture.Move (ScaleWidth - LPicture.Width) / 2, (ScaleHeight - LPicture.Height) / 2
LPicture.ZOrder
Else
LPicture.Visible = False
LCaption.Visible = True
End If
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LPicture,LPicture,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = LPicture.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set LPicture.Picture = New_Picture
    PropertyChanged "Picture"
Static Wi As Integer, He As Integer
Wi = ScaleWidth * 4 / 5
He = ScaleHeight * 3 / 5
If LPicture.Width > ScaleWidth Then LPicture.Width = Wi
If LPicture.Height > ScaleHeight Then LPicture.Height = He
LPicture.Left = (ScaleWidth - LPicture.Width) / 2
LPicture.Top = (ScaleHeight - LPicture.Height) / 2
'LPicture.ZOrder
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

