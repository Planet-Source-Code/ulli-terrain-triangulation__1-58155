VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'represents a side of a triangle

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'property variables
Private myFromPtr   As Long
Private myToPtr     As Long
Private myFromPoint As cVertex
Private myToPoint   As cVertex
Private myCenter    As cVertex
Private myLength    As Double
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Friend Property Get Center() As cVertex

    Set Center = myCenter

End Property

Private Sub Class_Initialize()

    Set myFromPoint = New cVertex
    Set myToPoint = New cVertex
    Set myCenter = New cVertex

End Sub

Friend Property Get CoincidesWith(Side As cSide) As Boolean

  'compare the projection of two sides on the horizontal plane

    CoincidesWith = (myFromPoint.CoincidesWith(Side.ToPoint) And myToPoint.CoincidesWith(Side.FromPoint))

End Property

Friend Property Get FromPoint() As cVertex

    Set FromPoint = myFromPoint

End Property

Friend Property Get FromPtr() As Long

    FromPtr = myFromPtr

End Property

Friend Property Get InvSlope() As Double

  'inverted slope, ie the slope of a line that is perpendicular to this side

    If Abs(myToPoint.Y - myFromPoint.Y) <= InvInfinity Then
        InvSlope = Infinity
      Else 'NOT ABS(MYTOPOINT.Y...
        InvSlope = (myFromPoint.X - myToPoint.X) / (myToPoint.Y - myFromPoint.Y)
    End If

End Property

Friend Property Get Length() As Double

    Length = myLength

End Property

Friend Function MakeSide(Terrain As cTerrain, ByVal pFromPoint As Long, ByVal pToPoint As Long) As cSide

  Dim X As Double
  Dim Y As Double
  Dim Z As Double

    myFromPtr = pFromPoint
    myToPtr = pToPoint
    With Terrain
        Set myFromPoint = .GetVertex(myFromPtr)
        Set myToPoint = .GetVertex(myToPtr)
    End With 'TERRAIN
    With myToPoint
        Set myCenter = myCenter.MakeVertex((myFromPoint.X + .X) / 2, (myFromPoint.Y + .Y) / 2, (myFromPoint.Z + .Z), -1)
        X = .X - myFromPoint.X
        Y = .Y - myFromPoint.Y
        Z = .Z - myFromPoint.Z
    End With 'MYTOPOINT
    myLength = Sqr(X * X + Y * Y + Z * Z)
    Set MakeSide = Me

End Function

Friend Property Get ToPoint() As cVertex

    Set ToPoint = myToPoint

End Property

Friend Property Get ToPtr() As Long

    ToPtr = myToPtr

End Property

':) Ulli's VB Code Formatter V2.18.3 (2005-Jan-07 14:09)  Decl: 13  Code: 90  Total: 103 Lines
':) CommentOnly: 6 (5,8%)  Commented: 3 (2,9%)  Empty: 35 (34%)  Max Logic Depth: 2
