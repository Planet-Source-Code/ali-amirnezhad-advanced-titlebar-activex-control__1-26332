VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TitleBar 
   BackColor       =   &H00FF0000&
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   51
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   309
   ToolboxBitmap   =   "TitleBar.ctx":0000
   Begin VB.Timer timTitleBar 
      Interval        =   100
      Left            =   900
      Top             =   60
   End
   Begin VB.CommandButton cmdMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   75
      Width           =   240
   End
   Begin VB.CommandButton cmdMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   240
   End
   Begin VB.CommandButton cmdClose 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   240
   End
   Begin MSComctlLib.ImageList imlPictures 
      Left            =   1380
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TitleBar.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TitleBar.ctx":0606
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TitleBar.ctx":08FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TitleBar.ctx":0BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TitleBar.ctx":0EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TitleBar.ctx":11D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TitleBar.ctx":14CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image picIcon 
      Height          =   240
      Left            =   90
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "TitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Click()
Public Event MouseOver()
Public Event MouseOut()
Public Event DoMinimize()
Public Event DoMaximize()
Public Event DoClose()
Public Enum Aligns
  tb_Left = 0
  tb_Right = 1
End Enum
Public Enum ColorStyles
  tb_Solid = 0
  tb_Gradient = 1
End Enum
Public Enum DrawStyles
  tb_full = 0
  tb_minimal = 1
End Enum
Private stlBackground As ColorStyles
Private picMain As Picture
Private alnTitleBar As Aligns
Private alnPicture As Aligns
Private blnShowImage As Boolean
Private colMain As OLE_COLOR
Private colLeft As OLE_COLOR
Private colRight As OLE_COLOR
Private strCaption As String
Private colCaption As OLE_COLOR

Private blnMouseOnTitleBar As Boolean
Private blnUserResize As Boolean
Private blnUserDraw As Boolean
Private lngUsedWidth As Long

Public Sub Gradient(StartColor As Long, EndColor As Long)
  Dim rgbStart As rgbColor
  Dim rgbEnd As rgbColor
  Dim rgbTemp As rgbColor
  Dim lngCounter As Long
  Dim colTemp As Long
  
  rgbStart = GetRGBColor(StartColor)
  rgbEnd = GetRGBColor(EndColor)
  UserControl.ScaleMode = 3
  For lngCounter = 0 To UserControl.ScaleWidth
    rgbTemp.lngRed = ((rgbEnd.lngRed - rgbStart.lngRed) / UserControl.ScaleWidth * lngCounter) + rgbStart.lngRed
    rgbTemp.lngGreen = ((rgbEnd.lngGreen - rgbStart.lngGreen) / UserControl.ScaleWidth * lngCounter) + rgbStart.lngGreen
    rgbTemp.lngBlue = ((rgbEnd.lngBlue - rgbStart.lngBlue) / UserControl.ScaleWidth * lngCounter) + rgbStart.lngBlue
    colTemp = RGB(rgbTemp.lngRed, rgbTemp.lngGreen, rgbTemp.lngBlue)
    UserControl.Line (lngCounter, 0)-(lngCounter, UserControl.ScaleHeight), colTemp, B
  Next lngCounter
End Sub

Private Sub Find_Object_Place()
  Dim lngPictureHeight, lngFontHeight, lngButtonsHeight As Long
  Dim lngWidth, lngHeight As Long
  
  lngWidth = UserControl.Parent.Width - 60
  lngUsedWidth = 25
  lngHeight = 22
  If (blnShowImage) Then
    lngPictureHeight = Int((Int(UserControl.ScaleY(picMain.Height, vbTwips, vbPixels)) + 1) / 2)
    If (lngHeight < lngPictureHeight) Then lngHeight = lngPictureHeight
  End If
  lngButtonsHeight = Int((Int(UserControl.ScaleY(cmdClose.Picture.Height, vbTwips, vbPixels)) + 1) / 2)
  If (lngHeight < lngButtonsHeight) Then lngHeight = lngButtonsHeight
  lngFontHeight = Int(UserControl.TextHeight(strCaption) / 2)
  If (lngHeight < lngFontHeight) Then lngHeight = lngFontHeight
  
  lngHeight = UserControl.ScaleY(lngHeight, vbPixels, vbTwips)
  blnUserResize = False
  Call UserControl.Size(lngWidth, lngHeight)
  blnUserResize = True
  
  lngWidth = Int((Int(UserControl.ScaleX(cmdClose.Picture.Width, vbTwips, vbPixels)) + 1) / 2)
  lngHeight = Int((Int(UserControl.ScaleY(cmdClose.Picture.Height, vbTwips, vbPixels)) + 1) / 2)
  lngUsedWidth = lngUsedWidth + (3 * lngWidth) + 15
  Call cmdMax.Move(0, 0, lngWidth, lngHeight)
  Call cmdMin.Move(0, 0, lngWidth, lngHeight)
  Select Case alnTitleBar
    Case tb_Right:
      Call picIcon.Move(5, Int((UserControl.ScaleHeight - 16) / 2))
      Call cmdClose.Move(UserControl.ScaleWidth - lngWidth - 5, Int((UserControl.ScaleHeight - lngHeight) / 2), lngWidth, lngHeight)
      Call cmdMax.Move(UserControl.ScaleWidth - 2 * lngWidth - 10, Int((UserControl.ScaleHeight - lngHeight) / 2), lngWidth, lngHeight)
      Call cmdMin.Move(UserControl.ScaleWidth - 3 * lngWidth - 11, Int((UserControl.ScaleHeight - lngHeight) / 2), lngWidth, lngHeight)
    Case tb_Left:
      Call picIcon.Move(UserControl.ScaleWidth - 21, Int((UserControl.ScaleHeight - 16) / 2))
      Call cmdClose.Move(5, Int((UserControl.ScaleHeight - lngHeight) / 2), lngWidth, lngHeight)
      Call cmdMax.Move(10 + lngWidth, Int((UserControl.ScaleHeight - lngHeight) / 2), lngWidth, lngHeight)
      Call cmdMin.Move(11 + 2 * lngWidth, Int((UserControl.ScaleHeight - lngHeight) / 2), lngWidth, lngHeight)
  End Select
End Sub

Private Sub Draw_Objects()
  Dim lngPictureHeight, lngPictureWidth As Long
  
  blnUserResize = False
  UserControl.Cls
  Select Case stlBackground
    Case tb_Solid: Call Gradient(colMain, colMain)
    Case tb_Gradient: Call Gradient(colLeft, colRight)
  End Select
  If (blnShowImage) Then
    lngPictureHeight = Int(UserControl.ScaleY(picMain.Height, vbTwips, vbPixels) / 2)
    lngPictureWidth = Int(UserControl.ScaleX(picMain.Width, vbTwips, vbPixels) / 2)
    Select Case alnPicture
      Case tb_Left: Call UserControl.PaintPicture(picMain, 0, Int((UserControl.ScaleHeight - lngPictureHeight) / 2))
      Case tb_Right:  Call UserControl.PaintPicture(picMain, UserControl.ScaleWidth - lngPictureWidth - 1, Int((UserControl.ScaleHeight - lngPictureHeight) / 2))
    End Select
  End If
  UserControl.ForeColor = colCaption
  UserControl.CurrentY = Int((UserControl.ScaleHeight - UserControl.TextHeight(strCaption)) / 2)
  Select Case alnTitleBar
    Case tb_Left: UserControl.CurrentX = (UserControl.ScaleWidth - 26) - UserControl.TextWidth(strCaption)
    Case tb_Right: UserControl.CurrentX = 26
  End Select
  UserControl.Print strCaption
  blnUserResize = True
End Sub

Public Sub Draw_TitleBar(ByVal dwsTemp As DrawStyles)
  blnUserResize = False
  If (dwsTemp = tb_full) Then
    Call Find_Object_Place
  End If
  Call Draw_Objects
  blnUserResize = True
End Sub

Public Property Get Align() As Aligns
  Align = alnTitleBar
End Property

Public Property Let Align(ByVal alnNewAlign As Aligns)
  alnTitleBar = alnNewAlign
  PropertyChanged "Align"
  Call Draw_TitleBar(tb_full)
End Property

Public Property Get BackgroundStyle() As ColorStyles
  BackgroundStyle = stlBackground
End Property

Public Property Let BackgroundStyle(ByVal stlNewBackgroundStyle As ColorStyles)
  stlBackground = stlNewBackgroundStyle
  PropertyChanged "BackgroundStyle"
  Call Draw_TitleBar(tb_minimal)
End Property

Public Property Get Image() As Picture
  Set Image = picMain
End Property

Public Property Set Image(ByVal picNewImage As Picture)
  Set picMain = picNewImage
  PropertyChanged "Image"
  Call Draw_TitleBar(tb_full)
End Property

Public Property Get ImageAlign() As Aligns
  ImageAlign = alnPicture
End Property

Public Property Let ImageAlign(ByVal alnNewImageAlign As Aligns)
  alnPicture = alnNewImageAlign
  PropertyChanged "ImageAlign"
  Call Draw_TitleBar(tb_minimal)
End Property

Public Property Get ShowImage() As Boolean
  ShowImage = blnShowImage
End Property

Public Property Let ShowImage(ByVal blnNewShowImage As Boolean)
  blnShowImage = blnNewShowImage
  PropertyChanged "ShowImage"
  Call Draw_TitleBar(tb_full)
End Property

Public Property Get Icon() As Picture
  Set Icon = picIcon.Picture
End Property

Public Property Set Icon(ByVal picNewIcon As Picture)
  Dim blnShowIcon As Boolean
  
  blnShowIcon = picIcon.Visible
  picIcon.Visible = False
  Set picIcon.Picture = picNewIcon
  PropertyChanged "Icon"
  picIcon.Visible = blnShowIcon
End Property

Public Property Get ShowIcon() As Boolean
  ShowIcon = picIcon.Visible
End Property

Public Property Let ShowIcon(ByVal blnNewShowIcon As Boolean)
  picIcon.Visible = blnNewShowIcon
  PropertyChanged "ShowIcon"
End Property

Public Property Get ColorMain() As OLE_COLOR
  ColorMain = colMain
End Property

Public Property Let ColorMain(ByVal colNewColorMain As OLE_COLOR)
  colMain = colNewColorMain
  PropertyChanged "ColorMain"
  Call Draw_TitleBar(tb_minimal)
End Property

Public Property Get ColorLeft() As OLE_COLOR
  ColorLeft = colLeft
End Property

Public Property Let ColorLeft(ByVal colNewColorLeft As OLE_COLOR)
  colLeft = colNewColorLeft
  PropertyChanged "ColorLeft"
  Call Draw_TitleBar(tb_minimal)
End Property

Public Property Get ColorRight() As OLE_COLOR
  ColorRight = colRight
End Property

Public Property Let ColorRight(ByVal colNewColorRight As OLE_COLOR)
  colRight = colNewColorRight
  PropertyChanged "ColorRight"
  Call Draw_TitleBar(tb_minimal)
End Property

Public Property Get Caption() As String
  Caption = strCaption
End Property

Public Property Let Caption(ByVal strNewCaption As String)
  strCaption = strNewCaption
  PropertyChanged "Caption"
  Call Draw_TitleBar(tb_minimal)
End Property

Public Property Get CaptionColor() As OLE_COLOR
  CaptionColor = colCaption
End Property

Public Property Let CaptionColor(ByVal colNewCaptionColor As OLE_COLOR)
  colCaption = colNewCaptionColor
  PropertyChanged "CaptionColor"
  Call Draw_TitleBar(tb_minimal)
End Property

Public Property Get CaptionFont() As Font
  Set CaptionFont = UserControl.Font
End Property

Public Property Set CaptionFont(ByVal fntNewCaptionFont As Font)
  Set UserControl.Font = fntNewCaptionFont
  PropertyChanged "CaptionFont"
  Call Draw_TitleBar(tb_full)
End Property

Public Property Get ShowMaxButton() As Boolean
  ShowMaxButton = cmdMax.Visible
End Property

Public Property Let ShowMaxButton(ByVal blnNewShowMaxButton As Boolean)
  cmdMax.Visible = blnNewShowMaxButton
  PropertyChanged "ShowMaxButton"
End Property

Public Property Get EnabledMaxButton() As Boolean
  EnabledMaxButton = cmdMax.Enabled
End Property

Public Property Let EnabledMaxButton(ByVal blnNewEnabledMaxButton As Boolean)
  cmdMax.Enabled = blnNewEnabledMaxButton
  PropertyChanged "EnabledMaxButton"
End Property

Public Property Get ShowMinButton() As Boolean
  ShowMinButton = cmdMin.Visible
End Property

Public Property Let ShowMinButton(ByVal blnNewShowMinButton As Boolean)
  cmdMin.Visible = blnNewShowMinButton
  PropertyChanged "ShowMinButton"
End Property

Public Property Get EnabledMinButton() As Boolean
  EnabledMinButton = cmdMin.Enabled
End Property

Public Property Let EnabledMinButton(ByVal blnNewEnabledMinButton As Boolean)
  cmdMin.Enabled = blnNewEnabledMinButton
  PropertyChanged "EnabledMinButton"
End Property

Public Property Get ShowCloseButton() As Boolean
  ShowCloseButton = cmdClose.Visible
End Property

Public Property Let ShowCloseButton(ByVal blnNewShowCloseButton As Boolean)
  cmdClose.Visible = blnNewShowCloseButton
  PropertyChanged "ShowCloseButton"
End Property

Public Property Get EnabledCloseButton() As Boolean
  EnabledCloseButton = cmdClose.Enabled
End Property

Public Property Let EnabledCloseButton(ByVal blnNewEnabledCloseButton As Boolean)
  cmdClose.Enabled = blnNewEnabledCloseButton
  PropertyChanged "EnabledCloseButton"
End Property

Public Property Get PictureMaxButton() As Picture
  Set PictureMaxButton = cmdMax.Picture
End Property

Public Property Set PictureMaxButton(ByVal picNewPictureMaxButton As Picture)
  Set cmdMax.Picture = picNewPictureMaxButton
  PropertyChanged "PictureMaxButton"
End Property

Public Property Get PictureMaxButtonDisable() As Picture
  Set PictureMaxButtonDisable = cmdMax.DisabledPicture
End Property

Public Property Set PictureMaxButtonDisable(ByVal picNewPictureMaxButtonDisable As Picture)
  Set cmdMax.DisabledPicture = picNewPictureMaxButtonDisable
  PropertyChanged "PictureMaxButtonDisable"
End Property

Public Property Get PictureMinButton() As Picture
  Set PictureMinButton = cmdMin.Picture
End Property

Public Property Set PictureMinButton(ByVal picNewPictureMinButton As Picture)
  Set cmdMin.Picture = picNewPictureMinButton
  PropertyChanged "PictureMinButton"
End Property

Public Property Get PictureMinButtonDisable() As Picture
  Set PictureMinButtonDisable = cmdMin.DisabledPicture
End Property

Public Property Set PictureMinButtonDisable(ByVal picNewPictureMinButtonDisable As Picture)
  Set cmdMin.DisabledPicture = picNewPictureMinButtonDisable
  PropertyChanged "PictureMinButtonDisable"
End Property

Public Property Get PictureCloseButton() As Picture
  Set PictureCloseButton = cmdClose.Picture
End Property

Public Property Set PictureCloseButton(ByVal picNewPictureCloseButton As Picture)
  Set cmdClose.Picture = picNewPictureCloseButton
  PropertyChanged "PictureCloseButton"
  Call Draw_TitleBar(tb_full)
End Property

Public Property Get PictureCloseButtonDisable() As Picture
  Set PictureCloseButtonDisable = cmdClose.DisabledPicture
End Property

Public Property Set PictureCloseButtonDisable(ByVal picNewPictureCloseButtonDisable As Picture)
  Set cmdClose.DisabledPicture = picNewPictureCloseButtonDisable
  PropertyChanged "PictureCloseButtonDisable"
End Property

Private Sub cmdClose_Click()
  RaiseEvent DoClose
End Sub

Private Sub cmdMax_Click()
  RaiseEvent DoMaximize
End Sub

Private Sub cmdMin_Click()
  RaiseEvent DoMinimize
End Sub

Private Sub timTitleBar_Timer()
  If (IsMouseInto(UserControl.hWnd)) Then
    If (blnMouseOnTitleBar = False) Then
      blnMouseOnTitleBar = True
      RaiseEvent MouseOver
    End If
  Else
    If (blnMouseOnTitleBar) Then
      blnMouseOnTitleBar = False
      RaiseEvent MouseOut
    End If
  End If
End Sub

Private Sub UserControl_Initialize()
  blnUserResize = True
  blnUserDraw = True
  
  blnMouseOnTitleBar = False
  stlBackground = tb_Solid
  Set picMain = imlPictures.ListImages(7).Picture
  alnTitleBar = tb_Right
  alnPicture = tb_Right
  blnShowImage = False
  picIcon.Visible = True
  colMain = vbBlue
  colLeft = vbRed
  colRight = vbYellow
  strCaption = "Amirnezhad's TitleBar"
  colCaption = vbBlack
  cmdMax.Visible = True
  cmdMax.Enabled = False
  cmdMin.Visible = True
  cmdMin.Enabled = True
  cmdClose.Visible = True
  cmdClose.Enabled = True
  Set cmdMax.Picture = imlPictures.ListImages(3).Picture
  Set cmdMin.Picture = imlPictures.ListImages(5).Picture
  Set cmdClose.Picture = imlPictures.ListImages(1).Picture
  Set cmdClose.DisabledPicture = imlPictures.ListImages(2).Picture
  Set cmdMax.DisabledPicture = imlPictures.ListImages(4).Picture
  Set cmdMin.DisabledPicture = imlPictures.ListImages(6).Picture
  Set picIcon.Picture = imlPictures.ListImages(7).Picture
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  DragParentForm (UserControl.Parent.hWnd)
End Sub

Private Sub UserControl_Paint()
  If (blnUserDraw) Then
    Draw_TitleBar (tb_full)
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  stlBackground = PropBag.ReadProperty("BackgroundStyle", tb_Solid)
  Set picMain = PropBag.ReadProperty("Image", imlPictures.ListImages(7).Picture)
  Set picIcon.Picture = PropBag.ReadProperty("Icon", imlPictures.ListImages(7).Picture)
  alnTitleBar = PropBag.ReadProperty("Align", tb_Right)
  alnPicture = PropBag.ReadProperty("PictureAlign", tb_Right)
  blnShowImage = PropBag.ReadProperty("ShowImage", False)
  picIcon.Visible = PropBag.ReadProperty("ShowIcon", True)
  colMain = PropBag.ReadProperty("MainColor", vbBlue)
  colLeft = PropBag.ReadProperty("LeftColor", vbRed)
  colRight = PropBag.ReadProperty("RightColor", vbYellow)
  strCaption = PropBag.ReadProperty("Caption", "Amirnezhad's TitleBar")
  colCaption = PropBag.ReadProperty("CaptionColor", vbBlack)
  cmdMax.Visible = PropBag.ReadProperty("ShowMaxButton", True)
  cmdMax.Enabled = PropBag.ReadProperty("EnabledMaxButton", False)
  cmdMin.Visible = PropBag.ReadProperty("ShowMinButton", True)
  cmdMin.Enabled = PropBag.ReadProperty("EnabledMinButton", True)
  cmdClose.Visible = PropBag.ReadProperty("ShowCloseButton", True)
  cmdClose.Enabled = PropBag.ReadProperty("EnabledCloseButton", True)
  Set cmdClose.Picture = PropBag.ReadProperty("ButtonClosePicture", imlPictures.ListImages(1).Picture)
  Set cmdClose.DisabledPicture = PropBag.ReadProperty("ButtonCloseDisabledPicture", imlPictures.ListImages(2).Picture)
  Set cmdMax.Picture = PropBag.ReadProperty("ButtonMaxPicture", imlPictures.ListImages(3).Picture)
  Set cmdMax.DisabledPicture = PropBag.ReadProperty("ButtonMaxDisabledPicture", imlPictures.ListImages(4).Picture)
  Set cmdMin.Picture = PropBag.ReadProperty("ButtonMinPicture", imlPictures.ListImages(5).Picture)
  Set cmdMin.DisabledPicture = PropBag.ReadProperty("ButtonMinDisabledPicture", imlPictures.ListImages(6).Picture)
  Set UserControl.Font = PropBag.ReadProperty("CaptionFont", Ambient.Font)
End Sub

Private Sub UserControl_Resize()
  If (blnUserResize) Then
    Call Draw_TitleBar(tb_full)
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackgroundStyle", stlBackground, tb_Solid)
  Call PropBag.WriteProperty("Image", picMain, imlPictures.ListImages(7).Picture)
  Call PropBag.WriteProperty("Icon", picIcon.Picture, imlPictures.ListImages(7).Picture)
  Call PropBag.WriteProperty("Align", alnTitleBar, tb_Right)
  Call PropBag.WriteProperty("PictureAlign", alnPicture, tb_Right)
  Call PropBag.WriteProperty("ShowImage", blnShowImage, False)
  Call PropBag.WriteProperty("ShowIcon", picIcon.Visible, True)
  Call PropBag.WriteProperty("MainColor", colMain, vbBlue)
  Call PropBag.WriteProperty("LeftColor", colLeft, vbRed)
  Call PropBag.WriteProperty("RightColor", colRight, vbYellow)
  Call PropBag.WriteProperty("Caption", strCaption, "Amirnezhad's TitleBar")
  Call PropBag.WriteProperty("CaptionColor", colCaption, vbBlack)
  Call PropBag.WriteProperty("ShowMaxButton", cmdMax.Visible, True)
  Call PropBag.WriteProperty("EnabledMaxButton", cmdMax.Enabled, False)
  Call PropBag.WriteProperty("ShowMinButton", cmdMin.Visible, True)
  Call PropBag.WriteProperty("EnabledMinButton", cmdMin.Enabled, True)
  Call PropBag.WriteProperty("ShowCloseButton", cmdClose.Visible, True)
  Call PropBag.WriteProperty("EnabledCloseButton", cmdClose.Enabled, True)
  Call PropBag.WriteProperty("ButtonClosePicture", cmdClose.Picture, imlPictures.ListImages(1).Picture)
  Call PropBag.WriteProperty("ButtonCloseDisabledPicture", cmdClose.DisabledPicture, imlPictures.ListImages(2).Picture)
  Call PropBag.WriteProperty("ButtonMaxPicture", cmdMax.Picture, imlPictures.ListImages(3).Picture)
  Call PropBag.WriteProperty("ButtonMaxDisabledPicture", cmdMax.DisabledPicture, imlPictures.ListImages(4).Picture)
  Call PropBag.WriteProperty("ButtonMinPicture", cmdMin.Picture, imlPictures.ListImages(5).Picture)
  Call PropBag.WriteProperty("ButtonMinDisabledPicture", cmdMin.DisabledPicture, imlPictures.ListImages(6).Picture)
  Call PropBag.WriteProperty("CaptionFont", UserControl.Font, Ambient.Font)
End Sub
