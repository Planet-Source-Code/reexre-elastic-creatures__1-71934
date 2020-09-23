Attribute VB_Name = "modPointLink"
'   Author : Roberto Mior
'   reexre@gmail.COM
'   If you use or modify source code or part of it for
'   developping a private application please report the author

Type Vector2
    X As Double
    Y As Double
End Type


Public Type tPOINT
    
    X As Double
    Y As Double
    vX As Double
    vY As Double
    NewVX As Double
    NewVY As Double
    isFix As Boolean
    LINKED As Boolean
    HowManyLinks As Integer
    ORGx As Double
    ORGy As Double
    
    TMPdelaunay As Boolean
    
    DynamicSpeed As Double
    DynamicFase As Double
    DynamicFaseLEN As Double
    
End Type

Public Type tLINK
    
    P1 As Integer
    P2 As Integer
    MainLenght As Double
    
    isNOTcross As Boolean
    BreakDist As Double
    isFix As Boolean
    SpringStrenght As Double
    
    DynamicLenght As Double
    DynamicSpeed As Double
    DynamicFase As Double
    
End Type

Public ICre As Integer
Public NCre As Integer

Public GlobalStiffness As Double
Public Gravity As Double
Public GravityX As Double
Public GravityY As Double
Public MaxH As Double
Public MaxW As Double
Public GlobalBreak As Double
Public GlobalSPRING As Double
'public GlobalTime As Double
'Public TimeSTEP As Double

Public PI As Double
Public FileName As String

Public AABB1 As tPOINT
Public AABB2 As tPOINT

Public B() As New clsBODY

Public Function Distance(P1 As tPOINT, P2 As tPOINT) As Double
Dim dx As Double
Dim dy As Double

dx = P1.X - P2.X
dy = P1.Y - P2.Y


Distance = Sqr(dx * dx + dy * dy)

End Function


Public Function vNormalize(Valu As Vector2) As Vector2
Dim factor As Double
Dim ZeroV As Vector2
'ZeroV.X = 0
'ZeroV.Y = 0

factor = 1 / (vDistance(Valu))

vNormalize.X = Valu.X * factor
vNormalize.Y = Valu.Y * factor


End Function
Public Function vDistance(Value1 As Vector2) As Double
Dim dx As Double
Dim dy As Double
dx = Value1.X '- Value2.X
dy = Value1.Y '- Value2.Y

vDistance = Sqr(dx * dx + dy * dy)
End Function
Public Function GetAngle(POS As Vector2, CenterPos As Vector2) As Double 'I borrowed this function from someone.
'Returns the angle between two points in
'     degrees
Dim intA As Double 'Integer
Dim intB As Double 'Integer
Dim intC As Double 'Integer
Dim PI As Double

PI = Atn(1) * 4
intB = Abs(CenterPos.X - POS.X) 'distance is always positive-->abs()
intC = Abs(CenterPos.Y - POS.Y)

If intB <> 0 Then 'don't divide by zero ...
    GetAngle = Atn(intC / intB) * 180 / PI
End If

If POS.X < CenterPos.X Then
    'the point is at the left of CenterPos
    If POS.Y = CenterPos.Y Then GetAngle = 180
    
    If POS.Y < CenterPos.Y Then
        GetAngle = 180 - GetAngle
    End If
    
    If POS.Y > CenterPos.Y Then
        GetAngle = 180 + GetAngle
    End If
End If

If POS.X > CenterPos.X Then
    'the point is at the right of CenterPos
    If POS.Y > CenterPos.Y Then
        GetAngle = 360 - GetAngle
    End If
End If

If POS.X = CenterPos.X Then
    
    If POS.Y < CenterPos.Y Then
        GetAngle = 90
    End If
    
    If POS.Y > CenterPos.Y Then
        GetAngle = 270
    End If
End If
'be sure the GetAngle is between [0,360]
GetAngle = Abs(GetAngle Mod 360)

GetAngle = (GetAngle / 180) * (PI)





End Function



Public Sub ReactToCreature()
Dim L1 As Integer
Dim L2 As Integer

Dim P1 As Vector2
Dim P2 As Vector2
Dim P3 As Vector2
Dim P4 As Vector2
Dim tmpP1 As tPOINT
Dim tmpP2 As tPOINT
Dim tL1 As tLINK
Dim tL2 As tLINK
'Stop
Dim Touch As Boolean

For c1 = 1 To NCre - 1
    For c2 = c1 + 1 To NCre
        Touch = False
        
        For L1 = 1 To B(c1).NLink
            tL1 = B(c1).GetLink(L1)
            If tL1.DynamicLenght = 0 Then
                
                tmpP1 = B(c1).GetPoint(tL1.P1)
                tmpP2 = B(c1).GetPoint(tL1.P2)
                P1.X = tmpP1.X
                P1.Y = tmpP1.Y
                P2.X = tmpP2.X
                P2.Y = tmpP2.Y
                
                
                For L2 = 1 To B(c2).NLink
                    
                    
                    tL2 = B(c2).GetLink(L2)
                    If tL2.DynamicLenght = 0 Then
                        
                        tmpP1 = B(c2).GetPoint(tL2.P1)
                        tmpP2 = B(c2).GetPoint(tL2.P2)
                        P3.X = tmpP1.X
                        P3.Y = tmpP1.Y
                        P4.X = tmpP2.X
                        P4.Y = tmpP2.Y
                        
                        If SegmentsIntersect(P1.X, P1.Y, _
                                P2.X, P2.Y, _
                                P3.X, P3.Y, _
                                P4.X, P4.Y, _
                                rx, ry) Then
                        '           Stop
                        
                        B(c2).TimeStep = -B(c2).TimeStep
                        B(c1).TimeStep = -B(c1).TimeStep
                        Touch = True
                        'Exit For
                        
                    End If
                    
                End If
            Next L2
        End If
        '   If Touch Then Exit For
    Next L1
    '''''''''''''If Touch Then Exit For
Next c2
'''''''''''If Touch Then Exit For
Next c1

End Sub
