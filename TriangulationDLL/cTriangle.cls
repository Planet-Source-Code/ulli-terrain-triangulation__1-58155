VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTriangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'represents a triangle

'''''''''''''''''''''''''''''''''''''''''''''
'property variables
Private mySideA         As cSide
Private mySideB         As cSide
Private mySideC         As cSide
Private myCoG           As cVertex 'center of gravity
Private myDegenerated   As Boolean
'''''''''''''''''''''''''''''''''''''''''''''

'circumcircle
Private ccCenterX       As Double
Private ccCenterY       As Double
Private ccSquaredRadius As Double
Private ccReach         As Double

'pointers into the terrain vertex set
Private CornerAPtr      As Long
Private CornerBPtr      As Long
Private CornerCPtr      As Long
Private Type Vector
    vx As Double
    vy As Double
    vz As Double
    l  As Double
End Type
Private vSideA          As Vector
Private vSideB          As Vector
Private vSideC          As Vector

Public Property Get Area() As Double
Attribute Area.VB_Description = "Area of the triangle."

  'computes my area using heron's fomula

  Dim a     As Double
  Dim b     As Double
  Dim c     As Double

    a = mySideA.Length
    b = mySideB.Length
    c = mySideC.Length

    Area = Sqr((a + b + c) * (a + b - c) * (a + c - b) * (b + c - a)) / 4 'Herons formula

End Property

Friend Property Get CCIncludes(Vertex As cVertex) As Boolean

  'true if the vertex is within (or on the border of) my circumcircle

  Dim X As Double
  Dim Y As Double

    If myDegenerated Then
        CCIncludes = True
      Else 'MYDEGENERATED = FALSE/0
        With Vertex
            X = .X - ccCenterX
            Y = .Y - ccCenterY
        End With 'VERTEX
        CCIncludes = (X * X + Y * Y <= ccSquaredRadius)
    End If

End Property

Public Property Get CenterOfGravity() As cVertex

    Set CenterOfGravity = myCoG

End Property

Private Sub Class_Initialize()

    Set mySideA = New cSide
    Set mySideB = New cSide
    Set mySideC = New cSide
    Set myCoG = New cVertex

End Sub

Private Sub ComputeCC()

  'circumcircle

  Dim Slope1        As Double
  Dim Slope2        As Double
  Dim SideACenterX  As Double
  Dim SideACenterY  As Double
  Dim SideCCenterX  As Double
  Dim SideCCenterY  As Double
  Dim X As Double
  Dim Y As Double

    With SideA
        Slope1 = .InvSlope
        With .Center
            SideACenterX = .X
            SideACenterY = .Y
        End With '.CENTER
    End With 'SIDEA
    With SideC
        Slope2 = .InvSlope
        With .Center
            SideCCenterX = .X
            SideCCenterY = .Y
        End With '.CENTER
    End With 'SIDEC

    If Abs(Slope1 - Slope2) <= InvInfinity Then 'not a triangle (or rather a degenerated triangle)
        myDegenerated = True
      Else 'NOT ABS(SLOPE1...
        ccCenterX = (Slope1 * SideACenterX - Slope2 * SideCCenterX + SideCCenterY - SideACenterY) / (Slope1 - Slope2)
        ccCenterY = Slope1 * (ccCenterX - SideACenterX) + SideACenterY
        With mySideA.FromPoint
            X = .X - ccCenterX
            Y = .Y - ccCenterY
        End With 'MYSIDEA.FROMPOINT
        ccSquaredRadius = X * X + Y * Y
        ccReach = Sqr(ccSquaredRadius) + ccCenterX
    End If

End Sub

Private Sub ComputeCoG()

  'center of gravity

    Set myCoG = myCoG.MakeVertex((CornerA3D.X + CornerB3D.X + CornerC3D.X) / 3, (CornerA3D.Y + CornerB3D.Y + CornerC3D.Y) / 3, (CornerA3D.Z + CornerB3D.Z + CornerC3D.Z) / 3, -1)

End Sub

Public Property Get CornerA2D() As cVertex

  Dim rx        As Double
  Dim ry        As Double
  Dim Number    As Long

    FlattenCorner CornerA3D, rx, ry, Number
    Set CornerA2D = New cVertex
    Set CornerA2D = CornerA2D.MakeVertex(rx, ry, 0, Number)

End Property

Public Property Get CornerA3D() As cVertex
Attribute CornerA3D.VB_Description = "Triangle's corner A coordinates."

    Set CornerA3D = mySideC.FromPoint

End Property

Public Property Get CornerB2D() As cVertex

  Dim rx        As Double
  Dim ry        As Double
  Dim Number    As Long

    FlattenCorner CornerB3D, rx, ry, Number
    Set CornerB2D = New cVertex
    Set CornerB2D = CornerB2D.MakeVertex(rx, ry, 0, Number)

End Property

Public Property Get CornerB3D() As cVertex
Attribute CornerB3D.VB_Description = "Triangle's corner B coordinates."

    Set CornerB3D = mySideA.FromPoint

End Property

Public Property Get CornerC2D() As cVertex

  Dim rx        As Double
  Dim ry        As Double
  Dim Number    As Long

    FlattenCorner CornerC3D, rx, ry, Number
    Set CornerC2D = New cVertex
    Set CornerC2D = CornerC2D.MakeVertex(rx, ry, 0, Number)

End Property

Public Property Get CornerC3D() As cVertex
Attribute CornerC3D.VB_Description = "Triangle's corner C coordinates."

    Set CornerC3D = mySideB.FromPoint

End Property

Private Function CrossProduct(V As Vector, W As Vector) As Vector

  'returns a vector that is perpendicular to the plane defined by the input sides

    With V
        CrossProduct.vx = .vy * W.vz - .vz * W.vy
        CrossProduct.vy = .vz * W.vx - .vx * W.vz
        CrossProduct.vz = .vx * W.vy - .vy * W.vx
    End With 'V
    With CrossProduct
        .l = Sqr(.vx * .vx + .vy * .vy + .vz * .vz)
    End With 'CROSSPRODUCT

End Function

Private Function Diff(V As cVertex, W As cVertex) As Vector

  'subtracts to vectors

    With Diff
        .vx = V.xRot - W.xRot
        .vy = V.yRot - W.yRot
        .vz = V.zRot - W.zRot
    End With 'DIFF
    With Diff
        .l = Sqr(.vx * .vx + .vy * .vy + .vz * .vz)
    End With 'DIFF

End Function

Private Function DotProduct(V As Vector, W As Vector) As Double

  'returns a value representing the cosine of the angle between two vectors

    With V
        DotProduct = .vx * W.vx + .vy * W.vy + .vz * W.vz
        'we're not interested in negative values - thus:
        If DotProduct < 0 Then
            DotProduct = 0
        End If
    End With 'V

End Function

Private Sub FlattenCorner(Corner As cVertex, ByRef rx As Double, ByRef ry As Double, ByRef rNumber As Long)

  Dim ScaleFactor As Double

    With Corner
        ScaleFactor = Sqr(1 + (myCoG.Z - .Z) ^ 2 / ((myCoG.X - .X) ^ 2 + (myCoG.Y - .Y) ^ 2))
        rx = (.X - myCoG.X) * ScaleFactor + myCoG.X
        ry = (.Y - myCoG.Y) * ScaleFactor + myCoG.Y
        rNumber = .Number
    End With 'CORNER

End Sub

Friend Property Get IsOutOfReach(Vertex As cVertex) As Boolean

    IsOutOfReach = (Vertex.X > ccReach) And Not myDegenerated

End Property

Friend Property Get IsTemp() As Boolean

    IsTemp = mySideA.FromPoint.IsTemp Or mySideB.FromPoint.IsTemp Or mySideC.FromPoint.IsTemp

End Property

Public Property Get LightIntensity() As Double

    vSideB = MakeVector(mySideB)
    vSideC = MakeVector(mySideC)
    LightIntensity = DotProduct(Normalize(CrossProduct(vSideB, vSideC)), Normalize(Diff(LightPos, myCoG)))

End Property

Friend Function MakeTriangle(Terrain As cTerrain, ByVal CornerA As Long, ByVal CornerB As Long, ByVal CornerC As Long) As cTriangle

    CornerAPtr = CornerA
    CornerBPtr = CornerB
    CornerCPtr = CornerC
    Set mySideA = mySideA.MakeSide(Terrain, CornerBPtr, CornerCPtr)
    Set mySideB = mySideB.MakeSide(Terrain, CornerCPtr, CornerAPtr)
    Set mySideC = mySideC.MakeSide(Terrain, CornerAPtr, CornerBPtr)
    ComputeCC
    ComputeCoG
    Set MakeTriangle = Me

End Function

Private Function MakeVector(Side As cSide) As Vector

    With Side
        MakeVector.vx = (.ToPoint.xRot - .FromPoint.xRot)
        MakeVector.vy = (.ToPoint.yRot - .FromPoint.yRot)
        MakeVector.vz = (.ToPoint.zRot - .FromPoint.zRot)
        MakeVector.l = .Length
    End With 'SIDE

End Function

Private Function Normalize(V As Vector) As Vector

  'makes a side's length = 1 (unless the lenght of the side is 0 anyway)

    With V
        If .l Then
            Normalize.vx = .vx / .l
            Normalize.vy = .vy / .l
            Normalize.vz = .vz / .l
          Else '.L = FALSE/0
            Normalize = V
        End If
    End With 'SIDE 'V

End Function

Friend Property Get SideA() As cSide

    Set SideA = mySideA

End Property

Friend Property Get SideB() As cSide

    Set SideB = mySideB

End Property

Friend Property Get SideC() As cSide

    Set SideC = mySideC

End Property

':) Ulli's VB Code Formatter V2.18.3 (2005-Jan-07 14:09)  Decl: 32  Code: 294  Total: 326 Lines
':) CommentOnly: 15 (4,6%)  Commented: 20 (6,1%)  Empty: 96 (29,4%)  Max Logic Depth: 3
