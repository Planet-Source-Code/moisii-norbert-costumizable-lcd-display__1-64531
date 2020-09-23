VERSION 5.00
Begin VB.UserControl LCD 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   148
   Begin VB.PictureBox picLCD 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   480
      Picture         =   "cLCD.ctx":0000
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "LCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=============================='
'         LCD DISPLAY          '
'         only numbers         '
'       by Moisii Norbert      '
'------------------------------'
'  i created this for my       '
'  skinnable mp3 player.       '
'  You can create your own     '
'  display style very fast     '
'  with a size as you want.    '
'  All you need is to set      '
'  the number width and height.'
'  The code is very small      '
'  and it's easy to understand '
'  please vote if you find     '
'  this usefull.               '
'sorry for my english.. :)     '
'=============================='
Option Explicit
'=================================================================================
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Const SRCCOPY = &HCC0020

Dim TempDC As Long
Dim TempBMP As Long

Dim nrW As Integer 'the numbers Width in the picture
Dim nrH As Integer ' - Height

Dim i As Integer, cap As String
'=================================================================================
Public Event Click()
Private Sub UserControl_Click()
RaiseEvent Click
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub LoadDC()
TempDC = CreateCompatibleDC(GetDC(0))
TempBMP = picLCD.Picture
SelectObject TempDC, TempBMP
End Sub
Private Sub UserControl_Initialize()
LoadDC
End Sub
Private Sub UserControl_InitProperties()
nrW = 10
nrH = 15
cap = "00:00"
LCDWrite cap
End Sub
Private Sub UserControl_Resize()
LCDWrite cap
End Sub
Public Property Let Caption(sCaption As String)
cap = sCaption
PropertyChanged "Caption"
LCDWrite sCaption
End Property
Public Property Get Caption() As String
Caption = cap
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
cap = PropBag.ReadProperty("Caption", "00:00")
nrW = PropBag.ReadProperty("NumberWidth", "10")
nrH = PropBag.ReadProperty("NumberHeight", "15")
Caption = cap
If Caption = "" Then
Refresh
Else
LCDWrite Caption
End If
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", Caption, "00:00"
PropBag.WriteProperty "NumberWidth", nrW, "10"
PropBag.WriteProperty "NumberHeight", nrH, "15"
End Sub
Private Sub LCDWrite(str As String)
  
  UserControl.Width = (nrW * (Len(str))) * Screen.TwipsPerPixelX
  UserControl.Height = nrH * Screen.TwipsPerPixelY

 If str = "" Then
 Refresh
 Else
  UserControl.Cls
  Dim xStr As String
  For i = 1 To Len(str)
  xStr = Mid(str, i, 1)
  Display xStr, i
  Debug.Print xStr
  Next i
 End If
End Sub
Private Sub Display(bNumber As Variant, pos As Integer)
  On Error Resume Next
  If bNumber = 0 Then
  BitBlt UserControl.hdc, (nrW * pos - nrW), 0, nrW, nrH, TempDC, nrW * 9, 0, SRCCOPY
  ElseIf bNumber = ":" Then
  BitBlt UserControl.hdc, (nrW * pos - nrW), 0, nrW, nrH, TempDC, nrW * 10, 0, SRCCOPY
  Else
  BitBlt UserControl.hdc, (nrW * pos - nrW), 0, nrW, nrH, TempDC, nrW * bNumber - nrW, 0, SRCCOPY
  End If
End Sub
' \\\
Public Sub LoadLCD(ffilename As String, nrWidth, nrHeight As Integer)
nrW = nrWidth: nrH = nrHeight
picLCD.Picture = LoadPicture(ffilename)
UserControl.Cls
LoadDC
End Sub
Public Function ScaleWidthPx()
ScaleWidthPx = UserControl.ScaleWidth
End Function
Public Function ScaleHeightPx()
ScaleHeightPx = UserControl.ScaleHeight
End Function
Public Sub Refresh()
LCDWrite "00:00"
End Sub

