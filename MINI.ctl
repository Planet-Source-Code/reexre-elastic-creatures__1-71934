VERSION 5.00
Begin VB.UserControl MINI 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   90
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.HScrollBar SH2 
      Height          =   270
      LargeChange     =   10
      Left            =   45
      Max             =   99
      TabIndex        =   0
      Top             =   975
      Visible         =   0   'False
      Width           =   4590
   End
   Begin VB.VScrollBar SH 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "MINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
 X As Long
 Y As Long
End Type
Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
Private Declare Sub ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long)
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, pRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long 'OK
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long  'Gets the hdc of Desktop
Private Declare Function GetDesktopWindow Lib "user32" () As Long   'Gets the hwnd of Desktop
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long     'Clear up the screen
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long     'Draw Animated rectangles( Using as the last event of animation )
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtVisible Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const DT_BOTTOM As Long = &H8
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4
Private Const DT_LEFT As Long = &H0
Private Const DT_CENTER As Long = &H1
Private Const DT_RIGHT As Long = &H2
Private Const FormatDes = DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX Or DT_CENTER
Private Const SRCAND = &H8800C6          ' (DWORD) dest = origen AND dest
Private Const SRCCOPY = &HCC0020        ' (DWORD) dest = origen
Private Const SRCERASE = &H440328        ' (DWORD) dest = origen AND (NOT dest)
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = origen OR dest
Public Enum AnimeMode
 aHide = 0
 aShow = 1
 aNothing = 2
End Enum
Public Enum AnimeSpeed
 aFast = 1
 aMedium = 10
 aSlow = 50
End Enum
Private DrawCol  As Long 'Color de la animacion
Private OldEvent As AnimeMode

Dim CurPos As POINTAPI, Cr As RECT, Ni As Integer, m_blnIsIn As Boolean
Const m_def_BorderColor = &HAE8962
Const m_def_BackColor = &HE1D4CB
Const m_def_BorderSize = 1
Const m_def_Size = 64
Const m_def_Spaces = 5
Const m_def_SelectColor = &HFF8080
Const m_def_AnimColor = &HAE8962
Dim m_BorderColor As OLE_COLOR
Dim m_BorderSize As Long
Dim m_Size As Long
Dim m_Spaces As Long
Dim m_SelectColor As OLE_COLOR
Dim m_AnimColor As OLE_COLOR

Private Type Cuadrado
 X As Long
 Y As Long
 X2 As Long
 Y2 As Long
End Type
Private Type tPicture
 hRgn As Long
 Caption As String
 POS As Cuadrado
 Imagen As Picture
End Type
Private PicName() As String
Private Pictures() As tPicture, NPictures As Long
Private NSelected As Long, OldSelected As Long
Private XPos As Single, YPos As Single
'Event Click(ThePicture As Picture, Nome As String)

Event Click(N As Integer)



Public Property Get BorderColor() As OLE_COLOR
 BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
 If m_BorderColor <> New_BorderColor Then
  m_BorderColor = New_BorderColor
  PropertyChanged "BorderColor"
  Call UserControl_Resize
 End If
End Property
Public Property Get BackColor() As OLE_COLOR
 BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
 If UserControl.BackColor() <> New_BackColor Then
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
  UserControl_Resize
 End If
End Property
Public Property Get BorderSize() As Long
 BorderSize = m_BorderSize
End Property
Public Property Let BorderSize(ByVal New_BorderSize As Long)
 If m_BorderSize <> New_BorderSize Then
  m_BorderSize = New_BorderSize
  PropertyChanged "BorderSize"
  UserControl_Resize
 End If
End Property
Public Property Get Size() As Long
 Size = m_Size
End Property
Public Property Let Size(ByVal New_Size As Long)
 If m_Size <> New_Size Then
  m_Size = New_Size
  PropertyChanged "Size"
  UserControl_Resize
 End If
End Property
Public Property Get Spaces() As Long
 Spaces = m_Spaces
End Property
Public Property Let Spaces(ByVal New_Spaces As Long)
 If m_Spaces <> New_Spaces Then
  m_Spaces = New_Spaces
  PropertyChanged "Spaces"
  UserControl_Resize
 End If
End Property
Public Property Get SelectColor() As OLE_COLOR
 SelectColor = m_SelectColor
End Property
Public Property Let SelectColor(ByVal New_SelectColor As OLE_COLOR)
 m_SelectColor = New_SelectColor
 PropertyChanged "SelectColor"
End Property
Public Property Get AnimColor() As OLE_COLOR
 AnimColor = m_AnimColor
End Property
Public Property Let AnimColor(ByVal New_AnimColor As OLE_COLOR)
 m_AnimColor = New_AnimColor
 PropertyChanged "AnimColor"
End Property

Private Sub UserControl_Initialize()



 m_blnIsIn = False
 OldEvent = aNothing
End Sub
Private Sub UserControl_InitProperties()
 On Error Resume Next
 Extender.Name = "MINI"
 While Err <> 0
  Err.Clear
  Extender.Name = "MINI" + CStr(Ni)
  Ni = Ni + 1
 Wend
 On Error GoTo 0
 m_BorderColor = m_def_BorderColor
 m_BorderSize = m_def_BorderSize
 m_Size = m_def_Size
 m_Spaces = m_def_Spaces
 XPos = m_Spaces
 YPos = m_Spaces
 UserControl.BackColor = m_def_BackColor
 m_SelectColor = m_def_SelectColor
 m_AnimColor = m_def_AnimColor
 UserControl_Resize
 
 

End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Mouse As POINTAPI, N As Long
 
 Call GetCursorPos(Mouse)
 'Call GetWindowRect(UserControl.hwnd, R)
 N = TestRegions(X, Y) '(Mouse.X, Mouse.Y)
 If N <> 0 Then
'       Stop
       
'RaiseEvent Click(Pictures(n).Imagen, PicName(n))
RaiseEvent Click(CInt(N))
  
 End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Mouse As POINTAPI, R As RECT
 Call GetCursorPos(Mouse)
 Call GetWindowRect(UserControl.hwnd, R)
 If 0 = PtInRect(R, Mouse.X, Mouse.Y) Then
  ReleaseCapture
  m_blnIsIn = False
  Exit Sub
 Else
  If WindowFromPoint(Mouse.X, Mouse.Y) <> UserControl.hwnd Then
   ReleaseCapture
   m_blnIsIn = False
   Exit Sub
  Else
   If m_blnIsIn = True Then
    '
   Else
    m_blnIsIn = True
   End If
   NSelected = TestRegions(X, Y) '(Mouse.X, Mouse.Y)
   If NSelected <> 0 Then UserControl_Resize
  End If
 End If
 Call SetCapture(UserControl.hwnd)
End Sub
Private Sub UserControl_Paint()
 UserControl_Resize
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
  UserControl.BackColor = .ReadProperty("BackColor", m_def_BackColor)
  m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
  m_BorderSize = .ReadProperty("BorderSize", m_def_BorderSize)
  m_Size = .ReadProperty("Size", m_def_Size)
  m_Spaces = .ReadProperty("Spaces", m_def_Spaces)
  m_SelectColor = .ReadProperty("SelectColor", m_def_SelectColor)
  m_AnimColor = .ReadProperty("AnimColor", m_def_AnimColor)
 End With
 UserControl_Resize
 'On Local Error Resume Next
 If Ambient.UserMode Then
  'If Err Then Err.Clear
  NPictures = 0
  ReDim Pictures(1)
  XPos = m_Spaces
  YPos = m_Spaces
 End If
 'On Local Error GoTo 0
End Sub
Private Sub UserControl_Terminate()
 On Local Error Resume Next
 If Ambient.UserMode Then
  Erase Pictures
 End If
 On Local Error GoTo 0
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  Call .WriteProperty("BackColor", UserControl.BackColor, m_def_BackColor)
  Call .WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
  Call .WriteProperty("BorderSize", m_BorderSize, m_def_BorderSize)
  Call .WriteProperty("Size", m_Size, m_def_Size)
  Call .WriteProperty("Spaces", m_Spaces, m_def_Spaces)
  Call .WriteProperty("SelectColor", m_SelectColor, m_def_SelectColor)
  Call .WriteProperty("AnimColor", m_AnimColor, m_def_AnimColor)
 End With
End Sub
Private Sub UserControl_Resize()
 Dim HScaleX As Integer, VScaleY As Integer
 Dim Cuadro As RECT, OldColor As OLE_COLOR, Th As Long, Lw As Long
 Dim Bs As Single
 HScaleX = ScaleWidth
 VScaleY = ScaleHeight
' Stop
 
 SH2.Move m_BorderSize + 1, VScaleY - 16, HScaleX - (m_BorderSize + 2), 16 - (m_BorderSize + 1)
 'SH2.Move HScaleX - 20, m_BorderSize + 1, 20 - (m_BorderSize + 1), VScaleY - (m_BorderSize + 2)

 
 UserControl.Cls
 UserControl.DrawWidth = 1 'm_BorderSize
 Bs = m_BorderSize * 0.5
 Call GetWindowRect(UserControl.hwnd, Cr)
 Call GetCursorPos(CurPos)
 Line (Bs, Bs)-(HScaleX - Bs - 1, VScaleY - Bs - 1), m_BorderColor, B
 If NPictures <> 0 Then
  If NSelected <> 0 Then
   UserControl.DrawWidth = 3
   With Pictures(NSelected)
    Line (.POS.X, .POS.Y)-(.POS.X2, .POS.Y2), m_SelectColor, B
   End With
  End If
 End If
End Sub
Public Sub GeneraMiniaturas(Path As String)
Dim i2 As Integer

Dim lPATH As String

i2 = InStrRev(Path, "\")

lPATH = Left$(Path, i2)

 On Error Resume Next
 Dim MiArchivo As String, MiRuta As String, MiNombre As String
 Dim Imagen As Picture, Ext As String, Fichero As String, i As Integer
 Clear
 MiRuta = Path
 If Right$(Path, 1) <> "\" Then MiRuta = MiRuta + "\"
 'MiArchivo = MiRuta + "*.*"
 MiArchivo = Left$(MiRuta, Len(MiRuta) - 1)
 MiNombre = Dir(MiArchivo, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
 i = 0
 SH2.Enabled = False
 SH2.Value = 0
 SH2.Visible = False
 Do While MiNombre <> ""
  'If MiNombre <> "." And MiNombre <> ".." Then
  
  'Fichero = MiRuta + MiNombre
  Fichero = lPATH + MiNombre
  
  If (GetAttr(Fichero) And vbDirectory) = vbDirectory Then GoTo Siguiente
   Ext = UCase$(Mid$(Fichero, InStr(1, Fichero, ".", vbTextCompare) + 1, 3))
   Select Case Ext
    Case "BMP", "JPG", "JPE", "ICO", "GIF", "WMF", "CUR" ', "ANI"
     Set Imagen = Nothing
'     Stop
     
'     Set Imagen = LoadPicture(MiRuta + MiNombre)
     Set Imagen = LoadPicture(Fichero)
'     Stop
     
     If Not (Imagen Is Nothing) Then Call AddPicture(Imagen, MiNombre)
     TestResize
   End Select
Siguiente:
  MiNombre = Dir
  i = i + 1
 Loop
' Stop
 
 SH2.Enabled = True
 Set Imagen = Nothing
 UserControl_Resize
 
' SH2.Value = SH2.Max
 
End Sub
Public Sub AddPicture(NImagen As Picture, Optional Titulo As String = "")
 NPictures = NPictures + 1
 ReDim Preserve Pictures(1 To NPictures)
 ReDim Preserve PicName(1 To NPictures)
 If Titulo = "" Then Titulo = CStr(NPictures)
' Stop
 
 PicName(NPictures) = Titulo
 With Pictures(NPictures)
  Set .Imagen = NImagen
  .Caption = Titulo
 ' .POS.X = YPos
 ' .POS.Y = XPos
 ' .POS.X2 = YPos + m_Size
 ' .POS.Y2 = XPos + m_Size
  .POS.X = XPos
  .POS.Y = YPos
  .POS.X2 = XPos + m_Size
  .POS.Y2 = YPos + m_Size
  
  
  
  .hRgn = CreateRectRgn(.POS.X, .POS.Y, .POS.X2, .POS.Y2)
 End With
 XPos = XPos + m_Size + m_Spaces
 Call DrawPicture(NPictures)
 TestResize
End Sub
Private Sub DrawPicture(Numero As Long)
 On Local Error Resume Next
 With Pictures(Numero)
  Call PaintPicture(.Imagen, .POS.X, .POS.Y, m_Size, m_Size)
 End With
 Set UserControl.Picture = UserControl.Image
 UserControl.Refresh
 On Local Error GoTo 0
End Sub
Private Sub DrawPictures()
 If NPictures = 0 Then Exit Sub
 Set UserControl.Picture = LoadPicture("")
 UserControl.Cls
 UserControl.Refresh
 Dim J As Long
 For J = 1 To NPictures
  With Pictures(J)
   Call PaintPicture(.Imagen, .POS.X, .POS.Y, m_Size, m_Size)
  End With
 Next J
 Set UserControl.Picture = UserControl.Image
 UserControl.Refresh
End Sub
Private Sub Clear()
 Dim K As Long
 If NPictures <> 0 Then
  For K = 1 To NPictures
   Call DeleteObject(Pictures(K).hRgn)
   Set Pictures(K).Imagen = Nothing
  Next
 End If
 NPictures = 0
 NSelected = 0
 OldSelected = 0
 Erase Pictures
 Set UserControl.Picture = LoadPicture("")
 UserControl.Cls
 UserControl.Refresh
 XPos = m_Spaces
 YPos = m_Spaces
End Sub
Private Function TestRegions(ByVal X As Long, ByVal Y As Long) As Long
 Dim Nr As Long, i As Long
 If NPictures = 0 Then
  Nr = 0
 Else
  For i = 1 To NPictures
   If PtInRegion(Pictures(i).hRgn, X, Y) <> 0 Then
    Nr = i
    Exit For
   End If
  Next
 End If
 TestRegions = Nr
End Function
Private Sub TestResize()
 Dim NVis As Long, NPd As Long
 If NPictures = 0 Then Exit Sub
 NPd = NPictures
' Stop
 
 With Pictures(NPictures)
  If .POS.X2 > ScaleWidth Then

 '   If .POS.Y2 > ScaleHeight Then
'  Stop
  
   SH2.Visible = True
   SH2.Max = NPd - 2 ' NPd - 3
   SH2.Min = 2
   NVis = Round(ScaleWidth / (m_Spaces + m_Size + m_Spaces))
   SH2.LargeChange = NVis
   SH2.SmallChange = 1
   '''''''' messo io
   SH2.Value = SH2.Max
   ''''''''
  Else
   'SH2.Visible = False
  End If
 End With
End Sub
Public Sub MMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub
Public Sub MDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub sh2_Change()
 MovePictures
End Sub
Private Sub sh2_Scroll()
 MovePictures
End Sub
Private Sub MovePictures()


 Dim Valor As Single, K As Long
 Dim HScaleX As Integer, VScaleY As Integer
 Dim Cuadro As RECT, OldColor As OLE_COLOR, Th As Long, Lw As Long
 Dim Bs As Single
 Valor = SH2.Value + 1
 If NPictures = 0 Then Exit Sub
 Set UserControl.Picture = LoadPicture("")
 UserControl.Cls
 UserControl.Refresh
' Stop
 
 'XPos = (-Valor + 1) * (m_Spaces + m_Size) + m_Spaces
 XPos = ScaleWidth - (Valor + 1) * (m_Spaces + m_Size) + m_Spaces
' Stop
 
 YPos = m_Spaces
 'XPos = m_Spaces
 
 For K = 1 To NPictures
  With Pictures(K)
   Call DeleteObject(.hRgn)
   .POS.X = XPos
   .POS.Y = YPos
   .POS.X2 = XPos + m_Size
   .POS.Y2 = YPos + m_Size
'   .POS.X = YPos
'   .POS.Y = XPos
'   .POS.X2 = YPos + m_Size
'   .POS.Y2 = XPos + m_Size
   
   .hRgn = CreateRectRgn(.POS.X, .POS.Y, .POS.X2, .POS.Y2)
   Call DrawPicture(K)
  End With
  XPos = XPos + m_Size + m_Spaces
 Next K
 'Dibuja el borde del control
 HScaleX = ScaleWidth
 VScaleY = ScaleHeight
 UserControl.DrawWidth = m_BorderSize
 Bs = m_BorderSize * 0.5
 Call GetWindowRect(UserControl.hwnd, Cr)
 Line (Bs, Bs)-(HScaleX - Bs - 1, VScaleY - Bs - 1), m_BorderColor, B
End Sub
'Funciones para la animacion del control
Private Function DeskDc()
 DeskDc = GetWindowDC(GetDesktopWindow)
End Function
Private Function DeskHwnd()
 DeskHwnd = GetDesktopWindow
End Function
Private Sub ClearScreen()
 InvalidateRect 0&, 0&, True
End Sub
Public Sub Anima(Optional aSpeed As AnimeSpeed = 10, Optional SleepTime As Integer = 1)
 Dim ScrX As Long, ScrY As Long, Rct1 As RECT, Rct2 As RECT
 Dim aEvent As AnimeMode
 Static CurPos As POINTAPI
 ScrX = Screen.TwipsPerPixelX
 ScrY = Screen.TwipsPerPixelY
 DrawCol = m_AnimColor
 If OldEvent = aNothing Then
  If Extender.Visible = True Then
   aEvent = aHide
  Else
   aEvent = aShow
  End If
 Else
  If OldEvent = aHide Then
   aEvent = aShow
  Else
   aEvent = aHide
  End If
 End If
 OldEvent = aEvent
 'If aEvent = aShow Then GetCursorPos CurPos
 GetCursorPos CurPos
 Call GetWindowRect(UserControl.hwnd, Rct1)
 With Rct2
  .Left = CurPos.X
  .Right = CurPos.X
  .Top = CurPos.Y
  .Bottom = CurPos.Y
 End With
 If aEvent = aShow Then
  PrivateAnime Rct2, Rct1, aSpeed, 10
  DrawAnimatedRects UserControl.hwnd, 3, Rct2, Rct1
  Extender.Visible = True
 Else
  Extender.Visible = False
  PrivateAnime Rct1, Rct2, aSpeed, 10
  DrawAnimatedRects UserControl.hwnd, 3, Rct1, Rct2
 End If
 ClearScreen
End Sub
Private Function PrivateAnime(sRct As RECT, eRct As RECT, ByVal aSpeed As AnimeSpeed, Optional ByVal RctCount = 25)
 Dim X As Integer, XIncr As Double, YIncr As Double
 Dim HIncr As Double, WIncr As Double, TempRect As RECT
 XIncr = (eRct.Left - sRct.Left) / RctCount
 YIncr = (eRct.Top - sRct.Top) / RctCount
 HIncr = ((eRct.Bottom - eRct.Top) - (sRct.Bottom - sRct.Top)) / RctCount
 WIncr = ((eRct.Right - eRct.Left) - (sRct.Right - sRct.Left)) / RctCount
 TempRect = sRct
 For X = 1 To RctCount
  Sleep aSpeed
  TempRect.Left = TempRect.Left + XIncr
  TempRect.Right = TempRect.Right + XIncr + WIncr
  TempRect.Top = TempRect.Top + YIncr
  TempRect.Bottom = TempRect.Bottom + YIncr + HIncr
  TransRectangle DeskDc, TempRect
 Next X
End Function
Private Sub TransRectangle(Dhdc As Long, VRct As RECT, Optional ByVal DrawWidth As Long = 1)
 Dim X As Integer, hBrush As Long, TempRect(1 To 4) As RECT
 For X = 1 To 4
  TempRect(X) = VRct
  If X = 1 Then TempRect(X).Bottom = TempRect(X).Top + DrawWidth
  If X = 2 Then TempRect(X).Left = TempRect(X).Right - DrawWidth
  If X = 3 Then TempRect(X).Top = TempRect(X).Bottom - DrawWidth
  If X = 4 Then TempRect(X).Right = TempRect(X).Left + DrawWidth
  hBrush = CreateSolidBrush(DrawCol)
  FillRect DeskDc, TempRect(X), hBrush
  DeleteObject hBrush
 Next X
End Sub





