VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVertex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'represents a point in 3d space

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'property variables
Private myX        As Double
Private myY        As Double
Private myZ        As Double
Private myNumber   As Long
Private myIsTemp   As Boolean
Private myIsInside As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Friend Function CoincidesWith(Vertex As cVertex) As Boolean

  'compares to vertices (we're only interested in their x and y coordinates so z is not compared)

    With Vertex
        CoincidesWith = (myX = .X And myY = .Y)
    End With 'VERTEX

End Function

Friend Property Get IsInside() As Boolean

    IsInside = myIsInside

End Property

Friend Property Get IsTemp() As Boolean

    IsTemp = myIsTemp

End Property

Friend Function MakeVertex(X As Double, Y As Double, Z As Double, Number As Long, Optional IsInside As Boolean = True, Optional IsTemp As Boolean = False) As cVertex

    myX = X
    myY = Y
    myZ = Z
    myNumber = Number
    myIsInside = IsInside
    myIsTemp = IsTemp
    Set MakeVertex = Me

End Function

Public Property Get Number() As Long

    Number = myNumber

End Property

Public Property Get X() As Double
Attribute X.VB_Description = "x-coordinate of vertex."

  'projection onto horiz plane

    X = myX

End Property

Public Property Get xRot() As Double

  'rotated as seen by the camera

    xRot = myX * RotationMatrix(1) + myY * RotationMatrix(2) + myZ * RotationMatrix(3)

End Property

Public Property Get Y() As Double
Attribute Y.VB_Description = "y-coordinate of vertex."

  'projection onto horiz plane

    Y = myY

End Property

Public Property Get yRot() As Double

  'rotated as seen by the camera

    yRot = myX * RotationMatrix(4) + myY * RotationMatrix(5) + myZ * RotationMatrix(6)

End Property

Public Property Get Z() As Double
Attribute Z.VB_Description = "z-coordinate of vertex."

  'elevation above / below horiz plane

    Z = myZ

End Property

Public Property Get zRot() As Double

  'rotated as seen by the camera

    zRot = myX * RotationMatrix(7) + myY * RotationMatrix(8) + myZ * RotationMatrix(9)

End Property

':) Ulli's VB Code Formatter V2.18.3 (2005-Jan-07 14:09)  Decl: 13  Code: 91  Total: 104 Lines
':) CommentOnly: 11 (10,6%)  Commented: 1 (1%)  Empty: 42 (40,4%)  Max Logic Depth: 2
