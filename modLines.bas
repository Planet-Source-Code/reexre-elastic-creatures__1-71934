Attribute VB_Name = "modLines"
'**************************************
' Name: Fast and easy line intersect fin
'     der
' Description:Finds the intersection of
'     two lines with straight forward code (no


'     if/then's) in any situation except when
'     they are parallel
' By: Daniel Whitmer
'
' Inputs:X1S,Y1S,X1E,Y1E,X2S,Y2S,X2E,Y2E
'
' Returns:the two points or an error mes
'     sage
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=50394&lngWId=1'for details.'**************************************

''Calculate the Deltas (distance of X2 - X1 or Y2 - Y1 of any 2 points)
'Delta(0).X = Line1.ptStart.X - Line2.ptStart.X   'Line1-Line2.ptStart X-Cross-Delta
'Delta(0).Y = Line1.ptStart.Y - Line2.ptStart.Y   'Line1-Line2.ptStart Y-Cross-Delta
'Delta(1).X = Line1.ptEnd.X - Line1.ptStart.X   'Line1 X-Delta
'Delta(1).Y = Line1.ptEnd.Y - Line1.ptStart.Y   'Line1 Y-Delta
'Delta(2).X = Line2.ptEnd.X - Line2.ptStart.X   'Line2 X-Delta
'Delta(2).Y = Line2.ptEnd.Y - Line2.ptStart.Y   'Line2 Y-Delta
Type PointDbl
    x As Double
    y As Double
End Type





'Function LineIntersect(Line1 As LineDbl, Line2 As LineDbl, ptIntersect As PointDbl) As Integer
Function LineIntersect(X1S, Y1S, X1E, Y1E, X2S, Y2S, X2E, Y2E, DestX As Double, DestY As Double) As Integer

'Calculate the intersection point of any two given non-parallel lines.
'
'Returns:  -1 = lines are parallel (no intersection).
'           0 = Neither line contains the intersect point between its points.**
'           1 = Line1 contains the intersect point between its points.**
'           2 = Line2 contains the intersect point between its points.**
'           3 = Both Lines contain the intersect point between their points.**
'           ** Lines Do intersect; Also fills in the ptIntersect point.
'
'BTW:       There are 18 lines of pure code, 25 lines of pure comments and 6
'           mixed lines in this function, just in case you were wondering. (:oþ}

Dim bIntersect  As Boolean
Dim iReturn     As Integer
Dim dDenom      As Double
Dim dPctDelta1  As Double
Dim dPctDelta2  As Double
Dim Delta(2)    As PointDbl

''Calculate the Deltas (distance of X2 - X1 or Y2 - Y1 of any 2 points)
'Delta(0).X = X1S - X2S   'Line1-Line2.ptStart X-Cross-Delta
'Delta(0).Y = Y1S - Y2S   'Line1-Line2.ptStart Y-Cross-Delta
'Delta(1).X = X1E - X1S   'Line1 X-Delta
'Delta(1).Y = Y1E - Y1S   'Line1 Y-Delta
'Delta(2).X = X2E - X2S   'Line2 X-Delta
'Delta(2).Y = Y2E - Y2S   'Line2 Y-Delta

'Calculate the Deltas (distance of X2 - X1 or Y2 - Y1 of any 2 points)
Delta(0).x = X1S - X2S 'Line1-Line2.ptStart X-Cross-Delta
Delta(0).y = Y1S - Y2S 'Line1-Line2.ptStart Y-Cross-Delta
Delta(1).x = X1E - X1S 'Line1 X-Delta
Delta(1).y = Y1E - Y1S 'Line1 Y-Delta
Delta(2).x = X2E - X2S 'Line2 X-Delta
Delta(2).y = Y2E - Y2S 'Line2 Y-Delta



'Calculate the denominator (zero = parallel (no intersection))
'Formula: (L2Dy * L1Dx) - (L2Dx * L1Dy)
iReturn = -1
dDenom = (Delta(2).y * Delta(1).x) - (Delta(2).x * Delta(1).y)
bIntersect = (dDenom <> 0)

If bIntersect Then
    'The lines will intersect somewhere.
    'Solve for both lines using the Cross-Deltas (Delta(0))
    
    'This yields percentage (0.1 = 10%; 1 = 100%) of the distance
    'between ptStart and ptEnd, of the opposite line, where the line used
    'in the calculation will cross it.
    '0 = ptStart direct hit; 1 = ptEnd direct hit; 0.5 = Centered between Pts; etc.
    'If < 0 or > 1 then the lines still intersect, just not between the points.
    
    'Solve for Line1 where Line2 will cross it.
    dPctDelta1 = ((Delta(2).x * Delta(0).y) - (Delta(2).y * Delta(0).x)) / dDenom
    
    'Solve for Line2 where Line1 will cross it.
    dPctDelta2 = ((Delta(1).x * Delta(0).y) - (Delta(1).y * Delta(0).x)) / dDenom
    
    'Check for absolute intersection. If the percentage is not between
    '0 and 1 then the lines will not intersect between their points.
    'Returns 0, 1, 2 or 3.
    iReturn = IIf(IsBetween(dPctDelta1, 0#, 1#), 1, 0) _
            Or IIf(IsBetween(dPctDelta2, 0#, 1#), 2, 0)
    
    'Calculate point of intersection on Line1 and fill ptIntersect.
    DestX = X1S + (dPctDelta1 * Delta(1).x) 'ptIntersect.X
    DestY = Y1S + (dPctDelta1 * Delta(1).y) ' ptIntersect.Y
    
End If
'        Stop

'Return the results.
LineIntersect = iReturn
'If (DestX = X1S) And (DestY = Y1S) Then LineIntersect = 3
'If (DestX = X1E) And (DestY = Y1E) Then LineIntersect = 3
'If (DestX = X2S) And (DestY = Y2S) Then LineIntersect = 3
'If (DestX = X2E) And (DestY = Y2E) Then LineIntersect = 3
End Function
Public Function IsBetween(ByVal vTestData As Variant, ByVal vLowerBound As Variant, ByVal vUpperBound As Variant, Optional ByVal bInclusive As Boolean = True) As Boolean

'Returns True if vTestData is between vLowerBound and vUpperBound.
'bInclusive = Are the bounds included in the test?

Dim vTemp   As Variant

If vLowerBound = vUpperBound Then
    Exit Function 'Returns false if upper and lower bounds are equal.
Else
    If vLowerBound > vUpperBound Then
        'If bounds are reversed, swap them.
        vTemp = vLowerBound
        vLowerBound = vUpperBound
        vUpperBound = vTemp
    End If
    If bInclusive Then
        'If bounds are included in test (use >= and <=).
        IsBetween = (vTestData >= vLowerBound) And (vTestData <= vUpperBound)
    Else
        'If bounds are not included in test (use > and <).
        IsBetween = (vTestData > vLowerBound) And (vTestData < vUpperBound)
    End If
End If

End Function

