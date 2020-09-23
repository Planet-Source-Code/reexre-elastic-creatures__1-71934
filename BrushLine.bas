Attribute VB_Name = "BrushLine"
Public Type POINTAPI
    X As Long
    y As Long
End Type

Public Poi As POINTAPI
Public PI As Double
Public PI2 As Double
Public PIm As Single




Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, _
                                          ByVal xInizioRettangolo As Long, _
                                          ByVal yInizioRettangolo As Long, _
                                          ByVal xFineRettangolo As Long, _
                                          ByVal yFineRettangolo As Long, _
                                          ByVal xInizioArco As Long, _
                                          ByVal yInizioArco As Long, _
                                          ByVal xFineArco As Long, _
                                          ByVal yFineArco As Long) As Long


'Declare Function Arc Lib "gdi32.dll" (ByVal HDC As Long, ByVal X1 As Long, _
ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long




Public Sub SetBrush(ByVal hdc As Long, ByVal PenWidth As Long, ByVal PenColor As Long)


DeleteObject (SelectObject(hdc, CreatePen(vbSolid, PenWidth, PenColor)))
'kOBJ = SelectObject(hDC, CreatePen(vbSolid, PenWidth, PenColor))
'SetBrush = kOBJ


End Sub



Public Sub FastLine(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long, ByVal W As Long, ByVal color As Long)
Attribute FastLine.VB_Description = "disegna line veloce"

SetBrush hdc, W, color

MoveToEx hdc, X1, Y1, Poi
LineTo hdc, X2, Y2

End Sub

Sub MyCircle(ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal R As Long, W, color)
SetBrush hdc, W, color
Arc hdc, X - R, y - R, X + R, y + R, X + R, y, X + R, y


End Sub


Public Function GetAngle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single

Dim dx As Single
Dim dy As Single

dx = X2 - X1
dy = Y2 - Y1

GetAngle = Atn(dy / (dx + 0.00000001)) * 180 / PI
If dx < 0 Then: GetAngle = (90 + GetAngle) + 90
If dy < 0 And dx >= 0 Then: GetAngle = 360 + GetAngle


GetAngle = GetAngle / 360 * PI2
End Function

'Public Function Distance(X1, Y1, X2, Y2, Toll) As Single
'Dim ddx As Single
'Dim ddy As Single
'
'ddx = X2 - X1
'ddy = Y2 - Y1
'
'Distance = Sqr(ddx * ddx + ddy * ddy)
''If Distance = 0 Then Stop
'
'If Distance < Toll Then Distance = 0'
''
'
'End Function
