VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTerrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'represents the triangulated terrain

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'property variables
Private myVertices      As Collection
Private myTriangles     As Collection
Private myHighest       As cVertex
Private myLowest        As cVertex
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Busy            As Boolean
Private LastX           As Double
Private SortedX         As Boolean 'indicates that the vertices are sorted by their x-component
Private Type tPoint
    px                  As Double
    py                  As Double
End Type

Public Event Progress(ByVal PercentCompleted As Long)

Public Function AddVertex(X As Double, Y As Double, Z As Double, Optional ByVal IsInsideVertex As Boolean = False, Optional ByVal Number As Long) As Long
Attribute AddVertex.VB_Description = "Adds a vertex to the terrain. Returns zero if the vertex was successfully added to the set; returns a postive number for a duplication and a negative number for a contradiction, the number being that of the conflicting vertex."

  'returns zero if the vertex was successfully added to the set
  'returns a postive number for a duplication and a negative number for a contradiction
  'the number in both cases being that of the conflicting vertex

  Dim Vertex    As cVertex
  Dim i         As Long

    If Busy Then
        Err.Raise 10, ErrorSource
      Else 'BUSY = FALSE/0
        For Each Vertex In myVertices
            i = i + 1
            If i = 1 Then
                Set myHighest = Vertex
                Set myLowest = Vertex
              Else 'NOT I...
                If Vertex.Z > myHighest.Z Then
                    Set myHighest = Vertex
                End If
                If Vertex.Z < myLowest.Z Then
                    Set myLowest = Vertex
                End If
            End If
            If Vertex.X = X And Vertex.Y = Y Then
                If Vertex.Z = Z Then 'duplicatiom
                    AddVertex = i
                  Else 'contradiction 'NOT VERTEX.Z...
                    AddVertex = -i
                End If
                Exit For 'loop varying vertex
            End If
        Next Vertex
        If AddVertex = 0 Then 'no dup - can add
            Set Vertex = New cVertex
            If Number = 0 Then
                Number = myVertices.Count + 1
            End If
            myVertices.Add Vertex.MakeVertex(X, Y, Z, Number, IsInsideVertex, False)
            If myVertices.Count > 1 Then 'check whether "outside"-sorted by x-coord
                If X < LastX Then
                    SortedX = False
                  Else 'NOT X...
                    LastX = X
                End If
              Else 'NOT MYVERTICES.COUNT...
                LastX = X
            End If
        End If
    End If

End Function

Private Sub Class_Initialize()

    Reset
    SortedX = True
    SetCameraPosition 180, 180, 180

    'Busy = (MsgBox( _
            "Copyright © 2004 UMGEDV GmbH" & vbCrLf & _
            vbCrLf & _
            "You may use this software free of charge for evaluation purposes only " & vbCrLf & _
            "for a period of 21 days after receipt. After that period you must " & vbCrLf & _
            vbCrLf & _
            "either" & vbCrLf & _
            vbCrLf & _
            "Unrecoverably destroy this software and all files you received (including" & vbCrLf & _
            "but not limited to all backup files and output on paper or elsewhere)" & vbCrLf & _
            vbCrLf & _
            "or" & vbCrLf & _
            vbCrLf & _
            "Buy a license from the copyright owner which entitles you to use this" & vbCrLf & _
            "software for any purpose you consider it fit." & vbCrLf & _
            vbCrLf & _
            "If you accept these conditions click  [Yes] ;  otherwise click  [No]", vbInformation Or vbYesNo, "Terrain Triangulation") = vbNo)

End Sub

Private Function DistanceBetween(PtA As tPoint, PtB As tPoint) As Double

  Dim X As Double
  Dim Y As Double

    With PtA
        X = .px - PtB.px
        X = .py - PtB.py
    End With 'PTA
    DistanceBetween = Sqr(X * X + Y * Y)

End Function

Friend Property Get GetVertex(ByVal Index As Long) As cVertex

  'for internal use while busy is on

    Set GetVertex = myVertices(Index)

End Property

Public Property Get HighestVertex() As cVertex

    If Busy Then
        Err.Raise 10, ErrorSource
      Else 'BUSY = FALSE/0
        Set HighestVertex = myHighest
    End If

End Property

Public Property Get LightPosition() As cVertex

    Set LightPosition = LightPos

End Property

Public Property Get LowestVertex() As cVertex

    If Busy Then
        Err.Raise 10, ErrorSource
      Else 'BUSY = FALSE/0
        Set LowestVertex = myLowest
    End If

End Property

Private Function MakeInitialTriangle() As cTriangle

  'make the initial triangle - this triangle encompasses all vertices

  Dim Vertex        As cVertex
  Dim Triangle      As cTriangle

  Dim idxVertex     As Long
  Dim Point1        As tPoint
  Dim Point2        As tPoint
  Dim InitialCenter As tPoint
  Dim Radius        As Double
  Dim Cosine        As Double

    idxVertex = myVertices.Count
    'find the two vertices most far apart
    Do While idxVertex
        Set Vertex = GetVertex(idxVertex)
        With Vertex
            If idxVertex = myVertices.Count Then
                Point1.px = .X
                Point1.py = .Y
                Point2 = Point1
              Else 'NOT IDXVERTEX...
                If DistanceBetween(Point1, MakePoint(.X, .Y)) > DistanceBetween(Point1, Point2) Then
                    Point2.px = .X
                    Point2.py = .Y
                End If
                If DistanceBetween(Point2, MakePoint(.X, .Y)) > DistanceBetween(Point2, Point1) Then
                    Point1.px = .X
                    Point1.py = .Y
                End If
            End If
        End With 'VERTEX
        idxVertex = idxVertex - 1
    Loop

    Radius = DistanceBetween(Point1, Point2) * 512 'circumcircle radius of the initial triangle
    With InitialCenter
        .px = (Point1.px + Point2.px) / 2
        .py = (Point1.py + Point2.py) / 2
        Cosine = Radius * Cos(Atn(1) * 2 / 3)
        Set Vertex = New cVertex
        myVertices.Add Vertex.MakeVertex(.px - Radius, .py, 0, -1, False, True)
        Set Vertex = New cVertex
        myVertices.Add Vertex.MakeVertex(.px + Radius / 2, .py - Cosine, 0, -1, False, True)
        Set Vertex = New cVertex
        myVertices.Add Vertex.MakeVertex(.px + Radius / 2, .py + Cosine, 0, -1, False, True)
    End With 'INITIALCENTER
    Set Triangle = New cTriangle
    With myVertices
        Set MakeInitialTriangle = Triangle.MakeTriangle(Me, .Count - 2, .Count - 1, .Count)
    End With 'MYVERTICES

End Function

Private Function MakePoint(X As Double, Y As Double) As tPoint

    With MakePoint
        .px = X
        .py = Y
    End With 'MAKEPOINT

End Function

Public Sub PresortVertices()

  Dim Vertex        As cVertex
  Dim idxVertex     As Long

    If Not SortedX Then
        If Busy Then
            Err.Raise 10, ErrorSource
          Else 'BUSY = FALSE/0
            With myVertices
                ReDim SortElems(0 To .Count)
                For idxVertex = 1 To .Count
                    Set Vertex = myVertices(idxVertex)
                    With SortElems(idxVertex)
                        .X = Vertex.X
                        .Y = Vertex.Y
                        .Z = Vertex.Z
                        .ii = Vertex.IsInside
                    End With 'SORTELEMS(IDXVERTEX)
                Next idxVertex
                QuickSort SortByX, 1, .Count
            End With 'MYVERTICES
            Set myVertices = New Collection
            For idxVertex = 1 To UBound(SortElems)
                With SortElems(idxVertex)
                    AddVertex .X, .Y, .Z, .ii
                    LastX = .X 'in case new vertices are added after sorting
                End With 'SORTELEMS(IDXVERTEX)
            Next idxVertex
            SortedX = True
            Erase SortElems
        End If
    End If

End Sub

Public Sub Reset()
Attribute Reset.VB_Description = "Removes all vertices an triangles from the terrain."

    If Busy Then
        Err.Raise 10, ErrorSource
      Else 'BUSY = FALSE/0
        Set myTriangles = New Collection
        Set myVertices = New Collection
    End If

End Sub

Public Sub SetCameraPosition(X As Double, Y As Double, Z As Double)

  'build the rotation matrix with x, y and z = camera position angles in °

  Dim SinX      As Double
  Dim CosX      As Double
  Dim SinY      As Double
  Dim CosY      As Double
  Dim SinZ      As Double
  Dim CosZ      As Double
  Dim Deg°      As Double

    Deg° = Atn(1) / 45

    SinX = Sin(X * Deg°)
    CosX = Cos(X * Deg°)
    SinY = Sin(Y * Deg°)
    CosY = Cos(Y * Deg°)
    SinZ = Sin(Z * Deg°)
    CosZ = Cos(Z * Deg°)

    'create the rotation matrix for the camera
    'x-axis
    RotationMatrix(1) = CosY * CosZ
    RotationMatrix(2) = CosX * SinZ + SinX * SinY * CosZ
    RotationMatrix(3) = SinX * SinZ - CosX * SinY * CosZ
    'y-axis
    RotationMatrix(4) = -CosY * SinZ
    RotationMatrix(5) = CosX * CosZ - SinX * SinX * SinZ
    RotationMatrix(6) = SinX * CosZ + CosX * SinY * SinZ
    'z-axis
    RotationMatrix(7) = SinY
    RotationMatrix(8) = -SinX * CosY
    RotationMatrix(9) = CosX * CosY

End Sub

Public Sub SetLightPosition(X As Double, Y As Double, Z As Double)

    Set LightPos = New cVertex
    Set LightPos = LightPos.MakeVertex(X, Y, Z, -1)

End Sub

Public Property Get TotalArea() As Double
Attribute TotalArea.VB_Description = "The total area of all triangles in the terrain."

  Dim Triangle As cTriangle

    If Busy Then
        Err.Raise 10, ErrorSource
      Else 'BUSY = FALSE/0
        For Each Triangle In myTriangles
            TotalArea = TotalArea + Triangle.Area
        Next Triangle
    End If

End Property

Public Property Get TriangleCount() As Long
Attribute TriangleCount.VB_Description = "The number of triangles in the terrain."

    If Busy Then
        Err.Raise 10, ErrorSource
      Else 'BUSY = FALSE/0
        TriangleCount = myTriangles.Count
    End If

End Property

Public Property Get Triangles(ByVal Index As Long) As cTriangle
Attribute Triangles.VB_Description = "Unordered set of triangles."

    If Busy Then
        Err.Raise 10, ErrorSource
      Else 'BUSY = FALSE/0
        If Index > 0 And Index <= myTriangles.Count Then
            Set Triangles = myTriangles(Index)
          Else 'NOT INDEX...
            Err.Raise 9, ErrorSource
        End If
    End If

End Property

Public Function Triangulate() As Long
Attribute Triangulate.VB_Description = "Triangulates the terrain and returns the number of triangles created."

  'triangulate the terrain and return the number of triangles created

  Dim colSides          As Collection
  Dim TempTriangles     As Collection
  Dim CurrentVertex     As cVertex
  Dim Side1             As cSide
  Dim Side2             As cSide
  Dim Triangle          As cTriangle

  Dim idxTriangle       As Long
  Dim idxSide1          As Long
  Dim idxSide2          As Long
  Dim idxVertex         As Long
  Dim ProgressBase      As Long
  Dim ProgressPcnt      As Long
  Dim PrevPcnt          As Long

    If Busy Then
        Err.Raise 10, ErrorSource
      Else 'BUSY = FALSE/0
        If myVertices.Count > 2 Then
            ProgressBase = myVertices.Count

            'here we go
            Set TempTriangles = New Collection
            Set myTriangles = New Collection
            TempTriangles.Add MakeInitialTriangle
            Busy = True
            RaiseEvent Progress(0)

            idxVertex = 0

            Do While idxVertex < myVertices.Count - 3 'for each user vertex... the last 3 belong to the initial triangle
                idxVertex = idxVertex + 1
                Set CurrentVertex = myVertices(idxVertex)

                'include the current vertex into the triangular mesh

                Set colSides = New Collection
                With colSides

                    'identify all triangles which have the current vertex inside their circumcircle
                    idxTriangle = TempTriangles.Count
                    Do While idxTriangle
                        Set Triangle = TempTriangles(idxTriangle)
                        If Triangle.CCIncludes(CurrentVertex) Then 'the current vertex is inside circumcircle
                            'add this triangle's sides to the sides collection...
                            .Add Triangle.SideA
                            .Add Triangle.SideB
                            .Add Triangle.SideC
                            '...and remove the triangle
                            TempTriangles.Remove idxTriangle
                          ElseIf SortedX Then 'TRIANGLE.CCINCLUDES(CURRENTVERTEX) = FALSE/0
                            If Triangle.IsOutOfReach(CurrentVertex) Then
                                myTriangles.Add Triangle 'ready - add triangle to the output set
                                TempTriangles.Remove idxTriangle 'and remove it from the temporary set
                            End If
                        End If
                        idxTriangle = idxTriangle - 1
                    Loop

                    'remove all duplicated sides
                    idxSide1 = .Count - 1
                    Do While idxSide1 > 0
                        Set Side1 = colSides(idxSide1)
                        idxSide2 = .Count
                        Do Until idxSide2 = idxSide1
                            Set Side2 = colSides(idxSide2)
                            If Side1.CoincidesWith(Side2) Then
                                .Remove idxSide2
                                .Remove idxSide1
                                Exit Do 'loop 
                              Else 'SIDE1.COINCIDESWITH(SIDE2) = FALSE/0
                                idxSide2 = idxSide2 - 1
                            End If
                        Loop
                        idxSide1 = idxSide1 - 1
                    Loop

                    'use the remaining sides and the current vertex to make new triangles
                    For Each Side1 In colSides
                        Set Triangle = New cTriangle
                        With Side1
                            TempTriangles.Add Triangle.MakeTriangle(Me, .FromPtr, .ToPtr, idxVertex)
                        End With 'SIDE1
                    Next Side1

                End With 'COLSIDES

                ProgressPcnt = idxVertex * 100 \ ProgressBase
                If ProgressPcnt <> PrevPcnt Then
                    RaiseEvent Progress(ProgressPcnt)
                    PrevPcnt = ProgressPcnt
                End If
            Loop 'next vertex

            If PrevPcnt < 100 Then
                RaiseEvent Progress(100)
            End If
            Busy = False

            'remove 3 temporary vertices
            With myVertices
                .Remove .Count
                .Remove .Count
                .Remove .Count
            End With 'MYVERTICES

            With myTriangles
                'add the (remaining) triangles
                For Each Triangle In TempTriangles
                    .Add Triangle
                Next Triangle
                Set TempTriangles = Nothing

                'remove temporary triangles
                idxTriangle = .Count
                Do While idxTriangle
                    Set Triangle = Triangles(idxTriangle)
                    If Triangle.IsTemp Then
                        .Remove idxTriangle
                    End If
                    idxTriangle = idxTriangle - 1
                Loop
            End With 'MYTRIANGLES

        End If
        Triangulate = myTriangles.Count
    End If

End Function

Public Property Get VertexCount() As Long
Attribute VertexCount.VB_Description = "The number of vertices in the terrain."

    If Busy Then
        Err.Raise 10, ErrorSource
      Else 'BUSY = FALSE/0
        VertexCount = myVertices.Count
    End If

End Property

Public Property Get Vertices(ByVal Index As Long) As cVertex
Attribute Vertices.VB_Description = "Unordered set of vertices."

    If Busy Then
        Err.Raise 10, ErrorSource
      Else 'BUSY = FALSE/0
        If Index > 0 And Index <= myVertices.Count Then
            Set Vertices = myVertices(Index)
          Else 'NOT INDEX...
            Err.Raise 9, ErrorSource
        End If
    End If

End Property

':) Ulli's VB Code Formatter V2.18.3 (2005-Jan-07 14:09)  Decl: 21  Code: 482  Total: 503 Lines
':) CommentOnly: 27 (5,4%)  Commented: 43 (8,5%)  Empty: 94 (18,7%)  Max Logic Depth: 8
