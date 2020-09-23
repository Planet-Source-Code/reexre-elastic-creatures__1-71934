VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBODY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Author : Roberto Mior
'   reexre@gmail.COM
'   If you use or modify source code or part of it for
'   developping a private application please report the author

Private Point() As tPOINT
Public Npoint As Long
Private Link() As tLINK
Public NLink As Long

Public Dissipation As Double

Private COM As tPOINT

Public FirstDELALink As Integer

Public pAABB1x As Double
Public pAABB1y As Double
Public pAABB2x As Double
Public pAABB2y As Double

'Private pAABB2 As tPOINT


Public InternalTimer As Double
Public TimeStep As Double

Public INTERSECTwithC As Boolean


Public Sub SetNpoint(numPoints)
Npoint = numPoints
ReDim Preserve Point(Npoint)

End Sub
Public Sub SetNlink(numlinks)
NLink = numlinks
ReDim Preserve Link(NLink)

End Sub



Friend Property Get GetPoint(NP As Integer) As tPOINT
GetPoint = Point(NP)
End Property
Friend Property Let SetPoint(NP As Integer, P As tPOINT)
Point(NP) = P
End Property

Friend Property Get GetLink(NL As Integer) As tLINK
GetLink = Link(NL)
End Property
Friend Property Let SetLink(NL As Integer, L As tLINK)
Link(NL) = L
End Property


Public Sub ADDPoint(X As Double, Y As Double, Optional IsStatic = False, _
        Optional Delaunay As Boolean = False, _
        Optional dynSPEED As Double = 0, _
        Optional dynFASE As Double = 0, _
        Optional FASElen As Double = 0)

'''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''
Npoint = Npoint + 1
ReDim Preserve Point(Npoint)

With Point(Npoint)
    
    .X = X
    .Y = Y
    .isFix = IsStatic
    .HowManyLinks = 0
    
    
    .TMPdelaunay = Delaunay
    
    .DynamicSpeed = dynSPEED
    .DynamicFase = dynFASE
    .DynamicFaseLEN = FASElen
    
    .ORGx = X
    .ORGy = Y
End With

End Sub

Public Function ADDLink(P1 As Integer, P2 As Integer, MultipleBreakLenght As Double, _
        Optional isFix As Boolean = False, Optional SpringStr As Double = 15, _
        Optional DynLEN = 0, Optional dynSPEED = 0, Optional dynFASE = 0)


NLink = NLink + 1
ReDim Preserve Link(NLink)
With Link(NLink)
    .P1 = P1
    .P2 = P2
    '        .MainLenght = Distance(GetPoint(P1), GetPoint(p2))
    .MainLenght = Distance(Point(P1), Point(P2))
    '    Stop
    .DynamicLenght = .MainLenght * DynLEN
    .DynamicSpeed = dynSPEED
    .DynamicFase = dynFASE
    
    .BreakDist = .MainLenght * MultipleBreakLenght
    Point(.P1).HowManyLinks = Point(.P1).HowManyLinks + 1
    Point(.P2).HowManyLinks = Point(.P2).HowManyLinks + 1
    .isFix = isFix
    .SpringStrenght = SpringStr
    
    
    
    
End With


End Function


Public Sub DRAW(dPIC As PictureBox, Optional Scala As Single = 1)
Dim PX1 As Double
Dim PY1 As Double
Dim PX2 As Double
Dim PY2 As Double


For i = 1 To Npoint
    
    PX1 = Point(i).X * Scala
    PY1 = Point(i).Y * Scala
    
    If Point(i).isFix Then
        
        
        
        dPIC.Line (PX1 - 3, PY1 - 4)- _
                (PX1 + 5, PY1 + 4), vbRed
        dPIC.Line (PX1 + 4, PY1 - 4)- _
                (PX1 - 4, PY1 + 4), vbRed
        
    Else
        
        dPIC.Circle (PX1, PY1), 2, vbRed
        
    End If
Next

For L = 1 To NLink
    i = Link(L).P1
    J = Link(L).P2
    
    PX1 = Point(i).X * Scala
    PY1 = Point(i).Y * Scala
    PX2 = Point(J).X * Scala
    PY2 = Point(J).Y * Scala
    
    dPIC.Line (PX1, PY1)- _
            (PX2, PY2), IIf(Link(L).DynamicLenght <> 0, vbGreen, vbRed)
    
    
    
Next




End Sub


Public Sub UpDateForces()
Dim DIst As Double
Dim DistDiff
Dim vDIFF As Vector2
Dim vNorm As Vector2
Dim pMED As Vector2
Dim Vec1 As Vector2
Dim Vec2 As Vector2
Dim TmpV1 As Vector2
Dim TmpV2 As Vector2
Dim LengV1 As Double
Dim LengV2 As Double
Dim ParV1 As Vector2
Dim ParV2 As Vector2
Dim PerpV1 As Vector2
Dim PerpV2 As Vector2
Dim P1 As Integer
Dim P2 As Integer

'Stop

InternalTimer = InternalTimer + TimeStep

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BOUNDAIRS COLLISION
Dim KK
KK = 0.997
For i = 1 To Npoint
    With Point(i)
        If .Y < 3 Or .Y > MaxH Then
            .Y = .Y - .NewVY
            .vY = -.vY * KK
            .NewVY = -.NewVY * KK
            .vX = .vX * 0.1 'Friction
            '            Stop
            If GravityX = 0 And GravityY = 0 Then TimeStep = -TimeStep
        End If
        If .X < 3 Or .X > MaxW Then
            .X = .X - .NewVX
            .vX = -.vX * KK
            .NewVX = -.NewVX * KK
            '.vY = .vY * 0.1 'Friction
            TimeStep = -TimeStep
        End If
    End With
Next

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For i = 1 To Npoint
    Point(i).NewVX = Point(i).vX + GravityX '
    Point(i).NewVY = Point(i).vY + GravityY ' Gravity
Next
'ìììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììì


For i = 1 To Npoint
    Dim T1 As Single
    Dim T2 As Single
    Dim T3 As Single
    
    With Point(i)
        T1 = Sin((InternalTimer * .DynamicSpeed) * 20 + (.DynamicFase - .DynamicFaseLEN) * PI * 2)
        T2 = Sin((InternalTimer * .DynamicSpeed) * 20 + .DynamicFase * PI * 2)
        T3 = Sin((InternalTimer * .DynamicSpeed) * 20 + (.DynamicFase + .DynamicFaseLEN) * PI * 2)
        '    Stop
        '    MsgBox Abs(T2) - .DynamicFaseLEN
        If (T2 + 1) - (1 - .DynamicFaseLEN) * 2 > 0 Then
            '    Stop
            
            .isFix = True
        Else
            .isFix = False
        End If
    End With ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Next
'ììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììììì

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Shape DEFORM
For L = 1 To NLink
    i = Link(L).P1
    J = Link(L).P2
    'Stop
    
    vDIFF.X = Point(i).X - Point(J).X
    vDIFF.Y = Point(i).Y - Point(J).Y
    DIst = Distance(Point(i), Point(J))
    'dIst = vDistance(VDIFF)
    vNorm = vNormalize(vDIFF)
    
    
    
    With Link(L)
        If .DynamicLenght = 0 Then
            DistDiff = DIst - .MainLenght
        Else
            '       Stop
            
            DistDiff = DIst - (.MainLenght + _
                    .DynamicLenght * Sin((InternalTimer * .DynamicSpeed) * 20 + .DynamicFase * PI * 2))
        End If
        
    End With
    
    
    
    ''
    '''
    'second Way adding this line:   (so reaction force proportional to NormDist) CORRECT!
    DistDiff = Link(L).SpringStrenght * DistDiff / Link(L).MainLenght '15
    ''''
    ''
    
    If DIst < Link(L).BreakDist Then
        '        If Link(L).isFix Then
        '
        ''GoTo secondoModo
        '
        '            Point(I).NewVX = Point(I).NewVX - DistDiff * vNorm.x * 0.5 * 1.01
        '            Point(I).NewVY = Point(I).NewVY - DistDiff * vNorm.y * 0.5 * 1.01
        '            Point(j).NewVX = Point(j).NewVX + DistDiff * vNorm.x * 0.5 * 1.01
        '            Point(j).NewVY = Point(j).NewVY + DistDiff * vNorm.y * 0.5 * 1.01
        '
        '
        '            GoTo Fatto
        '
        'secondoModo:
        '            '''non va
        '            pMED.x = (Point(I).x + Point(j).x) / 2
        '            pMED.y = (Point(I).x + Point(j).y) / 2
        '
        '            TmpV1.x = Point(I).vX
        '            TmpV1.y = Point(I).vY
        '            TmpV2.x = Point(j).vX
        '            TmpV2.y = Point(j).vY
        '
        '            LengV1 = vDistance(TmpV1)
        '            LengV2 = vDistance(TmpV2)
        '
        '            ParV1.x = TmpV1.x * Abs(vNorm.x)
        '            ParV1.y = TmpV1.y * Abs(vNorm.y)
        '            ParV2.x = TmpV2.x * Abs(vNorm.x)
        '            ParV2.y = TmpV2.y * Abs(vNorm.y)
        '
        '            PerpV1.x = TmpV1.x * Abs(vNorm.y)
        '            PerpV1.y = TmpV1.y * Abs(vNorm.x)
        '            PerpV2.x = TmpV2.x * Abs(vNorm.y)
        '            PerpV2.y = TmpV2.y * Abs(vNorm.x)
        '            sx = -Sgn(ParV1.x - ParV2.x)
        '            sy = -Sgn(ParV1.x - ParV2.x)
        '
        '            Point(I).NewVX = (ParV1.x + ParV2.x) / 2 + PerpV1.x
        '            Point(I).NewVY = (ParV1.y + ParV2.y) / 2 + PerpV1.y
        '            Point(j).NewVX = (ParV1.x + ParV2.x) / 2 + PerpV2.x
        '            Point(j).NewVX = (ParV1.y + ParV2.y) / 2 + PerpV2.y
        '            '
        '            'Point(I).X = pMED.X + Link(L).MainLenght * vNorm.X / 2
        '            'Point(I).Y = pMED.Y + Link(L).MainLenght * vNorm.Y / 2
        '            'Point(j).X = pMED.X - Link(L).MainLenght * vNorm.X / 2
        '            'Point(j).Y = pMED.Y - Link(L).MainLenght * vNorm.Y / 2
        '
        'Fatto:
        '        Else
        'Debug.Print DistDiff * VNorm.X & " " & DistDiff * VNorm.Y
        '        Stop
        
        Point(i).NewVX = Point(i).NewVX - DistDiff * vNorm.X * 0.5 * GlobalStiffness
        Point(i).NewVY = Point(i).NewVY - DistDiff * vNorm.Y * 0.5 * GlobalStiffness
        Point(J).NewVX = Point(J).NewVX + DistDiff * vNorm.X * 0.5 * GlobalStiffness
        Point(J).NewVY = Point(J).NewVY + DistDiff * vNorm.Y * 0.5 * GlobalStiffness
        '        End If
        
    Else
        BREAKLink (L)
    End If
    
    'End If
Next L
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'UPdate and Air Restsence
For i = 1 To Npoint
    With Point(i)
        If .isFix Then
            .vX = 0
            .vY = 0
        Else
            
            
            .vX = .NewVX * 0.997 '0.997
            .vY = .NewVY * 0.997 '0.997
            
            .X = .X + .vX
            .Y = .Y + .vY
        End If
        
    End With
    
Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FIXED LINK
'For L = 1 To NLink
'If Link(L).isFix Then
'    I = Link(L).P1
'    j = Link(L).p2
'    vDIFF.X = Point(I).X - Point(j).X
'    vDIFF.Y = Point(I).Y - Point(j).Y
'    DIst = Distance(Point(I), Point(j))
'    vNorm = vNormalize(vDIFF)
''    Stop
' '
'    DistDiff = DIst - Link(L).MainLenght
'            pMED.X = (Point(I).X + Point(j).X) / 2
'            pMED.Y = (Point(I).X + Point(j).Y) / 2
'            Point(I).X = pMED.X - Link(L).MainLenght * vNorm.X * 0.5 * Sgn(DistDiff)
'            Point(I).Y = pMED.Y - Link(L).MainLenght * vNorm.Y * 0.5 * Sgn(DistDiff)
'            Point(j).X = pMED.X + Link(L).MainLenght * vNorm.X * 0.5 * Sgn(DistDiff)
'            Point(j).Y = pMED.Y + Link(L).MainLenght * vNorm.Y * 0.5 * Sgn(DistDiff)
'End If
'Next L





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FindCROSSLinks_andUPDATE
'GLUE


End Sub
Sub BREAKLink(Li As Integer)

Dim NL
Dim i
Dim HowManyLinks As Integer

''''''''''''''''''''''''''''''''''''''''
NL = NLink
For i = Li To NL - 1
    Link(i) = Link(i + 1)
Next
NLink = NL - 1
'''''''''''''''''''''''''''''''''''''''''''''''''''

For i = 1 To Npoint
    Point(i).HowManyLinks = 0
    For L = 1 To NLink
        If (Link(L).P1 = i) Or (Link(L).P2 = i) Then
            Point(i).HowManyLinks = Point(i).HowManyLinks + 1
        End If
    Next
    'If Point(I).HowManyLinks Then Stop
    
    
Next

End Sub
Sub REMOVEPoint(P As Integer)

Dim NP
Dim i
Dim remove As Boolean

remove = True
For i = 1 To NLink
    If Link(i).P1 = P Then remove = False
    If Link(i).P2 = P Then remove = False
Next


If remove Then
    
    NP = Npoint
    For i = P To NP - 1
        Point(i) = Point(i + 1)
    Next
    Npoint = NP - 1
    ''''followin useless
    For i = 1 To NLink
        If Link(i).P1 >= P Then Link(i).P1 = Link(i).P1 - 1
        If Link(i).P2 >= P Then Link(i).P2 = Link(i).P2 - 1
    Next
End If




End Sub


Public Sub SaveMe(FileName As String)
Open FileName For Output As 1

'Print #1, "Stifness" 'obsolete
'Print #1, Replace(GlobalStiffness, ",", ".") 'obsolete
Print #1, "N-Points"
Print #1, Npoint
For i = 1 To Npoint
    PrintPoint Point(i)
    
Next
Print #1, "N-Links"
Print #1, NLink
For i = 1 To NLink
    PrintLink Link(i)
Next
Close 1



End Sub

Friend Sub PrintPoint(P As tPOINT)
'Stop

'Print #1, Replace(P.ORGx, ",", ".")
'Print #1, Replace(P.ORGy, ",", ".")
Print #1, Replace(P.X, ",", ".")
Print #1, Replace(P.Y, ",", ".")

Print #1, Replace(P.vX, ",", ".")
Print #1, Replace(P.vY, ",", ".")
Print #1, IIf(P.isFix, 1, 0)
Print #1, P.HowManyLinks

Print #1, Replace(P.DynamicSpeed, ",", ".")
Print #1, Replace(P.DynamicFase, ",", ".")
Print #1, Replace(P.DynamicFaseLEN, ",", ".")
End Sub
Friend Sub PrintLink(L As tLINK)
Print #1, L.P1
Print #1, L.P2
Print #1, Replace(L.MainLenght, ",", ".")
Print #1, Replace(L.BreakDist, ",", ".")
Print #1, Replace(L.SpringStrenght, ",", ".")
Print #1, IIf(L.isFix, 1, 0)
Print #1, IIf(L.isNOTcross, 1, 0)
Print #1, Replace(L.DynamicLenght, ",", ".")
Print #1, Replace(L.DynamicSpeed, ",", ".")
Print #1, Replace(L.DynamicFase, ",", ".")
End Sub
Friend Function InputPoint() As tPOINT
Input #1, InputPoint.X
Input #1, InputPoint.Y

Input #1, InputPoint.vX
Input #1, InputPoint.vY
Input #1, s
InputPoint.isFix = IIf(s = "1", True, False)
Input #1, InputPoint.HowManyLinks

Input #1, InputPoint.DynamicSpeed
Input #1, InputPoint.DynamicFase
Input #1, InputPoint.DynamicFaseLEN

InputPoint.ORGx = InputPoint.X
InputPoint.ORGy = InputPoint.Y

End Function
Friend Function InputLink() As tLINK
Input #1, InputLink.P1
Input #1, InputLink.P2
Input #1, InputLink.MainLenght
Input #1, InputLink.BreakDist
Input #1, InputLink.SpringStrenght

Input #1, s
InputLink.isFix = IIf(s = "1", True, False)
Input #1, s
InputLink.isNOTcross = IIf(s = "1", True, False)

Input #1, InputLink.DynamicLenght
Input #1, InputLink.DynamicSpeed
Input #1, InputLink.DynamicFase
End Function

Public Sub LoadMe(FileName As String)
Dim s As String

Clear

Open FileName For Input As 1

'Input #1, S
'Input #1, S ' GlobalStiffness

Input #1, s
Input #1, Npoint
ReDim Point(Npoint)
For i = 1 To Npoint
    Point(i) = InputPoint
Next
Input #1, s
Input #1, NLink
ReDim Link(NLink)
For i = 1 To NLink
    Link(i) = InputLink
Next

Close 1


End Sub


Public Sub Clear()
Npoint = 0
NLink = 0
End Sub

Public Sub FindCROSSLinks_InitialState()
Dim isB As Boolean
Dim XX As Double
Dim yy As Double

'Stop


For L1 = 1 To NLink
    isB = True
    For L2 = 1 To NLink
        If L1 <> L2 Then
            
            If (Link(L1).P1 <> Link(L2).P1) And _
                    (Link(L1).P1 <> Link(L2).P2) And _
                    (Link(L1).P2 <> Link(L2).P1) And _
                    (Link(L1).P2 <> Link(L2).P2) Then
            
            If (SegmentsIntersect(Point(Link(L1).P1).X, _
                    Point(Link(L1).P1).Y, _
                    Point(Link(L1).P2).X, _
                    Point(Link(L1).P2).Y, _
                    Point(Link(L2).P1).X, _
                    Point(Link(L2).P1).Y, _
                    Point(Link(L2).P2).X, _
                    Point(Link(L2).P2).Y, _
                    XX, yy)) = True Then isB = False: Exit For
            
        Else
            '    Stop
            
        End If
        
    End If
Next
'If isB Then Stop

Link(L1).isNOTcross = isB

Next


End Sub


Sub FindCROSSLinks_andUPDATE() '''non funziona''not working
Dim isB As Boolean
Dim XX As Double
Dim yy As Double
Dim L1 As Integer
Dim L2 As Integer
Dim LC As Integer


For L1 = 1 To NLink
    
    
    isB = True
    For L2 = 1 To NLink
        If L1 <> L2 Then
            
            If (Link(L1).P1 <> Link(L2).P1) And _
                    (Link(L1).P1 <> Link(L2).P2) And _
                    (Link(L1).P2 <> Link(L2).P1) And _
                    (Link(L1).P2 <> Link(L2).P2) Then
            
            If (SegmentsIntersect(Point(Link(L1).P1).X, _
                    Point(Link(L1).P1).Y, _
                    Point(Link(L1).P2).X, _
                    Point(Link(L1).P2).Y, _
                    Point(Link(L2).P1).X, _
                    Point(Link(L2).P1).Y, _
                    Point(Link(L2).P2).X, _
                    Point(Link(L2).P2).Y, _
                    XX, yy)) = True Then isB = False: Exit For
            
        Else
            '    Stop
            
        End If
        
    End If
Next

If (Link(L1).isNOTcross) Then
    If Link(L1).isNOTcross <> isB Then
        NearP = findNearestPoint(L1)
        '    Stop
        
        Point(Link(L1).P1).X = Point(Link(L1).P1).X - Point(Link(L1).P1).vX
        Point(Link(L1).P1).Y = Point(Link(L1).P1).Y - Point(Link(L1).P1).vY
        Point(Link(L1).P2).X = Point(Link(L1).P2).X - Point(Link(L1).P2).vX
        Point(Link(L1).P2).Y = Point(Link(L1).P2).Y - Point(Link(L1).P2).vY
        
        
        Point(Link(L1).P1).vX = -Point(Link(L1).P1).vX '+ Point(nearp).vX / 2
        Point(Link(L1).P1).vY = -Point(Link(L1).P1).vY '+ Point(nearp).vY / 2
        Point(Link(L1).P2).vX = -Point(Link(L1).P2).vX '+ Point(nearp).vX / 2
        Point(Link(L1).P2).vY = -Point(Link(L1).P2).vY '+ Point(nearp).vY / 2
        
        
        Point(NearP).X = Point(NearP).X - Point(NearP).vX
        Point(NearP).Y = Point(NearP).Y - Point(NearP).vY
        Point(NearP).vX = -Point(NearP).vX * 2 '1.2 + Point(Link(L1).P1).vX + Point(Link(L1).P2).vX
        Point(NearP).vY = -Point(NearP).vY * 2 '1.2 + Point(Link(L1).P1).vY + Point(Link(L1).P2).vY
        
        
    End If
    
End If

Next

End Sub
Public Function findNearestPoint(L As Integer) As Integer
Dim D As Double
Dim Dmin As Double
Dim P1 As Integer
Dim P2 As Integer

P1 = Link(L).P1
P2 = Link(L).P2

Dim pMED As tPOINT

pMED.X = (Point(Link(L).P1).X + Point(Link(L).P2).X) / 2
pMED.Y = (Point(Link(L).P1).Y + Point(Link(L).P2).Y) / 2


Dmin = 99999999
For i = 1 To Npoint
    If i <> P1 And i <> P2 Then
        If Point(i).isFix = False And Point(i).LINKED = True Then
            
            D = Distance(Point(i), pMED)
            If D < Dmin Then Dmin = D: findNearestPoint = i
        End If
    End If
    
Next
'Stop

End Function

Sub GLUE()
'Stop

Dim i As Integer
Dim J As Integer

For i = 1 To Npoint
    
    If Point(i).HowManyLinks < 2 Then
        For J = 1 To Npoint
            If i <> J Then
                If Distance(Point(i), Point(J)) < 40 Then
                    ADDLink i, J, Val(Replace(frmMAIN.cmbBREAK, ",", ".")), _
                            IIf(frmMAIN.chLFix.Value = Checked, True, False), _
                            Val(Replace(frmMAIN.cmbSPRING, ",", "."))
                End If
                
            End If
        Next J
    End If
    
Next

End Sub



Public Sub FindEXTERNALlinks() ''''' to develope
Dim V() As Vector2
Dim Vtmp() As Vector2
Dim PAng() As Double

ReDim V(Npoint)
ReDim Vtmp(Npoint)
ReDim PAng(Npoint)


For i = 1 To Npoint
    V(i).X = Point(i).X
    V(i).Y = Point(i).Y
Next


StartP = 1

For i = 1 To Npoint
    If i <> StartP Then PAng(i) = GetAngle(V(i), V(StartP))
    Debug.Print PAng(i)
Next

End Sub
Sub UnCheckDELApoints()
For i = 1 To Npoint
    Point(i).TMPdelaunay = False
Next
End Sub

Sub Delaunay()
Dim i As Integer

Dim Dcount As Integer
Dim StartP As Integer

clearDELAresult

For i = FirstDELALink To NLink
    BREAKLink i
Next


Dcount = 0
For i = 1 To Npoint
    If Point(i).TMPdelaunay Then
        If Dcount = 0 Then StartP = i - 1
        Dcount = Dcount + 1
        '  ReDim Preserve Vertex(Dcount)
        Vertex(Dcount).X = Point(i).X
        Vertex(Dcount).Y = Point(i).Y
        Vertex(Dcount).z = 0
    End If
    
Next

nTria = Triangulate(Dcount)


If nTria > 0 Then 'createDELALinks Ntria, StartP
    '''''''''''''''''''''''''''''''''''''
    ' ADD Triangles LINKS
    frmMAIN.Caption = "Adding Links.. "
    DoEvents
    
    Dim BR As Double
    Dim IsF As Boolean
    Dim SPR As Double
    
    BR = Val(Replace(frmMAIN.cmbBREAK, ",", "."))
    IsF = IIf(frmMAIN.chLFix.Value = Checked, True, False)
    SPR = Val(Replace(frmMAIN.cmbSPRING, ",", "."))
    
    
    For t = 1 To nTria
        ADDLink Triangle(t).vv0 + StartP, Triangle(t).vv1 + StartP, BR, IsF, SPR
        ADDLink Triangle(t).vv1 + StartP, Triangle(t).vv2 + StartP, BR, IsF, SPR
        ADDLink Triangle(t).vv0 + StartP, Triangle(t).vv2 + StartP, BR, IsF, SPR
    Next
    
    frmMAIN.PIC.Cls
    DRAW frmMAIN.PIC
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If nTria > 1 Then
        RemoveOverLapLinks
    End If
    
End If


End Sub







Private Sub RemoveOverLapLinks()

frmMAIN.Caption = "Remove Overlappig Links.. "
DoEvents


Dim L1 As Integer
Dim L2 As Integer
Dim P1L1 As Integer
Dim P2L1 As Integer
Dim P1L2 As Integer
Dim P2L2 As Integer

For L1 = 1 To NLink
    For L2 = L1 + 1 To NLink - 1
        P1L1 = Link(L1).P1
        P2L1 = Link(L1).P2
        P1L2 = Link(L2).P1
        P2L2 = Link(L2).P2
        
        If (P1L1 = P1L2) And (P2L1 = P2L2) Or _
                (P1L1 = P2L2) And (P2L1 = P1L2) Then
        BREAKLink L2
    End If
    
    
Next
Next

End Sub

'Public Sub LinkInvertDynSpeed(L As Integer)
'Link(L).DynamicSpeed = -Link(L).DynamicSpeed
'End Sub
'useless


Public Sub FindAABB()

pAABB1x = 999999
pAABB1y = 999999
pAABB2x = -999999
pAABB2y = -999999

For i = 1 To Npoint
    With Point(i)
        
        If .X > pAABB2x Then pAABB2x = .X
        If .X < pAABB1x Then pAABB1x = .X
        If .Y > pAABB2y Then pAABB2y = .Y
        If .Y < pAABB1y Then pAABB1y = .Y
        
        
        
    End With
Next
End Sub

Private Sub Class_Initialize()
TimeStep = 0.005

End Sub


Public Sub CreatePreview(picP As PictureBox, sFileName As String)
Dim NewScale As Single
Dim s As String
Dim II As Integer
'Stop

II = InStrRev(sFileName, ".")
s = Left$(sFileName, II - 1) & ".bmp"




'picP.ScaleMode = 3
'picP.ScaleWidth = 1000
'picP.ScaleHeight = 1000
'picP.BackColor = 0

'Stop

FindAABB
'Stop
pAABB1x = pAABB1x - 5
pAABB1y = pAABB1y - 5
pAABB2x = pAABB2x + 5
pAABB2y = pAABB2y + 5

If (pAABB2x - pAABB1x) > (pAABB2y - pAABB1y) Then
    
    NewScale = picP.ScaleWidth / (pAABB2x - pAABB1x)
    
Else
    NewScale = picP.ScaleHeight / (pAABB2y - pAABB1y)
    
    
End If

'Stop



For i = 1 To Npoint
    Point(i).X = (Point(i).X - pAABB1x) '* NewScale
    Point(i).Y = (Point(i).Y - pAABB1y) '* NewScale
Next
'Stop
picP.Line (0, 0)-(500, 500), 0, BF


DRAW picP, NewScale
picP.Refresh

SavePicture picP.Image, s

For i = 1 To Npoint
    Point(i).X = Point(i).X + pAABB1x
    Point(i).Y = Point(i).Y + pAABB1y
Next


End Sub



Public Sub ORGxUgaleX()
For i = 1 To Npoint
    
    Point(i).ORGx = Point(i).X
    Point(i).ORGy = Point(i).Y
    
Next

End Sub



