VERSION 5.00
Begin VB.UserControl OfficeXPToolbar 
   BackColor       =   &H00D1D8DB&
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   ScaleHeight     =   1005
   ScaleWidth      =   5835
   ToolboxBitmap   =   "Main.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00D1D8DB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   600
      Picture         =   "Main.ctx":0312
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00B59285&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   120
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00A6A6A6&
      Index           =   0
      Visible         =   0   'False
      X1              =   375
      X2              =   375
      Y1              =   330
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   45
      Picture         =   "Main.ctx":069C
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000374&
      FillColor       =   &H008D90FF&
      FillStyle       =   0  'Solid
      Height          =   330
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
End
Attribute VB_Name = "OfficeXPToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim LastX As Integer
Dim TotalButtons As Integer
Dim CurrentOn As Integer
Dim MovedAlready As Boolean
Dim DefinedBorderColor As OLE_COLOR
Dim DefinedBackColor As OLE_COLOR
Dim DefinedImageDisplacement As Integer
Dim DefinedPressColor As OLE_COLOR
Dim TempColor As OLE_COLOR

Dim ButtonPressed As Integer
Private poiCursorPos As POINTAPI

Public Event Click(Index As Integer, Key As String)
Public Event MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, X As Single)
Public Event MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, X As Single)

Public Property Get HighlightBackColor() As OLE_COLOR
Attribute HighlightBackColor.VB_Description = "Returns/sets the background color of a highlighted toolbar item."
HighlightBackColor = DefinedBackColor
End Property
Public Property Let HighlightBackColor(ByVal NewValue As OLE_COLOR)
DefinedBackColor = NewValue
UserControl.PropertyChanged "HighlightBackColor"
End Property

Public Property Get PressColor() As OLE_COLOR
PressColor = DefinedPressColor
End Property
Public Property Let PressColor(ByVal NewValue As OLE_COLOR)
DefinedPressColor = NewValue
UserControl.PropertyChanged "PressColor"
End Property
Public Property Get HighlightBorderColor() As OLE_COLOR
Attribute HighlightBorderColor.VB_Description = "Returns/sets the bordercolor of a highlighted toolbar item."
HighlightBorderColor = DefinedBorderColor
End Property
Public Property Let HighlightBorderColor(ByVal NewValue As OLE_COLOR)
DefinedBorderColor = NewValue
UserControl.PropertyChanged "HighlightBorderColor"
End Property
Public Property Get ImageDisplacement() As Integer
Attribute ImageDisplacement.VB_Description = "Returns/sets the amount of pixels a toolbar item will be displaced by when the mouse moves over it."
ImageDisplacement = DefinedImageDisplacement
End Property
Public Property Let ImageDisplacement(ByVal NewValue As Integer)
DefinedImageDisplacement = NewValue
UserControl.PropertyChanged "ImageDisplacement"
End Property
Private Sub Image1_Click(Index As Integer)
RaiseEvent Click(Index, Image1(Index).Tag)
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Index, Button, Shift, X, Y)
ButtonPressed = Index
TempColor = Shape1(Index).FillColor
Shape1(Index).FillColor = Shape2(Index).FillColor
Image1(Index).Left = Image1(Index).Left + DefinedImageDisplacement
Image1(Index).Top = Image1(Index).Top + DefinedImageDisplacement
MovedAlready = False
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not CurrentOn = Index Then
If Shape1(CurrentOn).Visible = True Then Shape1(CurrentOn).Visible = False
If MovedAlready = True Then
MovedAlready = False
Image1(CurrentOn).Left = Image1(CurrentOn).Left + DefinedImageDisplacement
Image1(CurrentOn).Top = Image1(CurrentOn).Top + DefinedImageDisplacement
End If
End If

CurrentOn = Index
If MovedAlready = True Then Exit Sub
MovedAlready = True
Shape1(Index).Left = Image1(Index).Left - 45
Shape1(Index).Top = 0
Shape1(Index).ZOrder 1
Shape1(Index).Visible = True
Image1(Index).Top = Image1(Index).Top - DefinedImageDisplacement
Image1(Index).Left = Image1(Index).Left - DefinedImageDisplacement


End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Index, Button, Shift, X, Y)
ButtonPressed = -1
If Not Shape1(Index).FillColor = TempColor Then Shape1(Index).FillColor = TempColor
End Sub

Private Sub Timer1_Timer()
'Check if the mouse is on the toolbar.  If not, make all highlights invisible
On Error Resume Next
Dim lonCStat As Long
Dim lonCurrhWnd As Long

Timer1.Enabled = False


lonCStat = GetCursorPos&(poiCursorPos)
lonCurrhWnd = WindowFromPoint(poiCursorPos.X, poiCursorPos.Y)

If lonCurrhWnd <> UserControl.hWnd Then
z = 0
Do While Not z = TotalButtons + 1
If ButtonPressed = z Then
Else
Shape1(z).Visible = False
If MovedAlready = True Then
MovedAlready = False
Image1(CurrentOn).Left = Image1(CurrentOn).Left + DefinedImageDisplacement
Image1(CurrentOn).Top = Image1(CurrentOn).Top + DefinedImageDisplacement
End If
End If
z = z + 1

Loop
Else
End If

Timer1.Enabled = True
End Sub

Private Sub UserControl_Initialize()
MovedAlready = False
TotalButtons = -1
LastX = -280
HighlightBackColor = &HD6BEB5
HighlightBorderColor = &H6B2408
PressColor = &HB59285
DefinedImageDisplacement = 10
ButtonPressed = -1 'no buttons pressed
End Sub
Public Sub AddButton(ButtonType As Integer, Icon, Optional ButtonKey As String, Optional ToolTipText As String, Optional LineColor As OLE_COLOR, Optional FillColor As OLE_COLOR, Optional PressColor As OLE_COLOR)
TotalButtons = TotalButtons + 1 'increment

If LineColor Then
Else
LineColor = DefinedBorderColor
End If

If FillColor Then
Else
FillColor = DefinedBackColor
End If

If PressColor Then
Else
PressColor = DefinedPressColor
End If

If Not ButtonKey = "" Then
Else
ButtonKey = Str(TotalButtons)
End If

If ButtonType = 1 Then 'if type = button
LastX = LastX + 350
If Not TotalButtons = 0 Then
Load Image1(TotalButtons)
Load Shape1(TotalButtons)
Load Shape2(TotalButtons) 'load the image and place it.
Load Picture1(TotalButtons)
End If

Image1(TotalButtons).Tag = ButtonKey
Image1(TotalButtons).Picture = Icon
Image1(TotalButtons).Top = 45
Image1(TotalButtons).Left = LastX
Image1(TotalButtons).ToolTipText = ToolTipText
Image1(TotalButtons).Visible = True
Shape1(TotalButtons).BorderColor = LineColor
Shape1(TotalButtons).FillColor = FillColor
Shape2(TotalButtons).FillColor = PressColor
End If

If ButtonType = 0 Then 'if type = separator
If Not TotalButtons = 0 Then
Load Line1(TotalButtons)
End If
Line1(TotalButtons).X1 = LastX + 330
Line1(TotalButtons).X2 = LastX + 330
Line1(TotalButtons).Y1 = 330
Line1(TotalButtons).Y2 = 0
Line1(TotalButtons).Visible = True
LastX = LastX + 90
End If

End Sub
Public Sub SetStatus(Index As Integer, Status As Integer)
If Status = 0 Then 'disabled
    Const PICSIZE = 16
    Picture1(Index).ScaleMode = vbPixels
    Picture1(Index).Picture = Image1(Index).Picture
    Picture1(Index).Left = Image1(Index).Left
    Picture1(Index).Top = Image1(Index).Top
    DisableHDC Picture1(Index).hdc, PICSIZE, PICSIZE
    Image1(Index).Visible = False
    Picture1(Index).Visible = True
    Picture1(Index).Refresh
Else 'enabled

    Image1(Index).Visible = True
    Picture1(Index).Visible = False
End If
End Sub
Public Function GetStatus(Index As Integer)
If Image1(Index).Visible = True Then
GetStatus = 1
Else
GetStatus = 0
End If
End Function

Private Sub UserControl_InitProperties()
Timer1.Enabled = Ambient.UserMode
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not ButtonPressed = CurrentOn Then
If Shape1(CurrentOn).Visible = True Then Shape1(CurrentOn).Visible = False
If MovedAlready = True Then
MovedAlready = False
Image1(CurrentOn).Left = Image1(CurrentOn).Left + DefinedImageDisplacement
Image1(CurrentOn).Top = Image1(CurrentOn).Top + DefinedImageDisplacement
End If
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Timer1.Enabled = Ambient.UserMode
HighlightBorderColor = PropBag.ReadProperty("HighlightBorderColor", &H6B2408)
HighlightBackColor = PropBag.ReadProperty("HighlightBackColor", &HD6BEB5)
PressColor = PropBag.ReadProperty("PressColor", &HB59285)
DefinedImageDisplacement = PropBag.ReadProperty("DefinedImageDisplacement", 10)
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 330

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "HighlightBorderColor", HighlightBorderColor, &H6B2408
PropBag.WriteProperty "HighlightBackColor", HighlightBackColor, &HD6BEB5
PropBag.WriteProperty "PressColor", PressColor, &HB59285
PropBag.WriteProperty "DefinedImageDisplacement", DefinedImageDisplacement, 10
End Sub
