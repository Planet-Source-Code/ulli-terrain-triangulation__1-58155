VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fTriangulate 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "3D Terrain Triangulation"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12705
   FillStyle       =   0  'Ausgefüllt
   Icon            =   "fTriangulate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox ckPerspective 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Perspective"
      Height          =   195
      Left            =   6960
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Add perspective distortion"
      Top             =   300
      Width           =   1140
   End
   Begin VB.CheckBox ckNumbers 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Numbers"
      Height          =   195
      Left            =   7155
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Show vertex numbers"
      Top             =   525
      Value           =   1  'Aktiviert
      Width           =   945
   End
   Begin VB.HScrollBar scrY 
      Height          =   270
      LargeChange     =   9
      Left            =   165
      Max             =   360
      SmallChange     =   3
      TabIndex        =   19
      Top             =   9900
      Value           =   180
      Width           =   12090
   End
   Begin VB.VScrollBar scrX 
      Height          =   9090
      LargeChange     =   9
      Left            =   12270
      Max             =   360
      SmallChange     =   3
      TabIndex        =   18
      Top             =   810
      Value           =   180
      Width           =   270
   End
   Begin VB.CheckBox ckSort 
      Alignment       =   1  'Rechts ausgerichtet
      BackColor       =   &H00E0E0E0&
      Caption         =   "Presort"
      Height          =   300
      Left            =   8280
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Presort vertices"
      Top             =   285
      Value           =   1  'Aktiviert
      Width           =   795
   End
   Begin VB.CheckBox ckDefer 
      Caption         =   "&Defer"
      Height          =   360
      Left            =   9285
      Style           =   1  'Grafisch
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Defer triangulation"
      Top             =   240
      Width           =   795
   End
   Begin VB.CheckBox ckMesh 
      Alignment       =   1  'Rechts ausgerichtet
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Mesh"
      Height          =   195
      Left            =   7410
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Show wire mesh"
      Top             =   75
      Value           =   1  'Aktiviert
      Width           =   690
   End
   Begin VB.TextBox txNum 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   11790
      MaxLength       =   4
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "100"
      ToolTipText     =   "Number of random vertices to generate"
      Top             =   285
      Width           =   480
   End
   Begin MSComctlLib.ProgressBar pgb 
      Align           =   2  'Unten ausrichten
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   10290
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton btRandom 
      Caption         =   "&Random"
      Height          =   360
      Left            =   10950
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Generate random vertices"
      Top             =   240
      Width           =   795
   End
   Begin VB.CommandButton btReset 
      Caption         =   "&Erase"
      Height          =   360
      Left            =   10125
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Reset"
      Top             =   240
      Width           =   795
   End
   Begin VB.PictureBox picCanvas 
      BackColor       =   &H00FFFFFF&
      DrawMode        =   9  'Stift maskieren invers
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Ausgefüllt
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   9075
      Left            =   187
      ScaleHeight     =   601
      ScaleLeft       =   -400
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleTop        =   -300
      ScaleWidth      =   801
      TabIndex        =   0
      Top             =   810
      Width           =   12075
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   4
      Left            =   2910
      TabIndex        =   17
      Top             =   555
      Width           =   45
   End
   Begin VB.Label lbPosY 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1950
      TabIndex        =   16
      Top             =   555
      Width           =   45
   End
   Begin VB.Label lbPosX 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1080
      TabIndex        =   15
      Top             =   555
      Width           =   45
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Click into the area below (left to add vertex)  (right to position light)."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   600
      Index           =   3
      Left            =   4830
      TabIndex        =   14
      Top             =   120
      Width           =   2130
   End
   Begin VB.Image imgUMG 
      Height          =   630
      Left            =   150
      Picture         =   "fTriangulate.frx":08CA
      Top             =   60
      Width           =   675
   End
   Begin VB.Label lb 
      Caption         =   "Area in Square Units"
      Height          =   195
      Index           =   2
      Left            =   2925
      TabIndex        =   13
      Top             =   105
      Width           =   1455
   End
   Begin VB.Label lb 
      Caption         =   "Triangles"
      Height          =   195
      Index           =   1
      Left            =   1950
      TabIndex        =   12
      Top             =   90
      Width           =   645
   End
   Begin VB.Label lb 
      Caption         =   "Vertices"
      Height          =   195
      Index           =   0
      Left            =   1095
      TabIndex        =   11
      Top             =   90
      Width           =   570
   End
   Begin VB.Label lbTotalArea 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2910
      TabIndex        =   9
      Top             =   315
      Width           =   45
   End
   Begin VB.Label lbTriangles 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1950
      TabIndex        =   8
      Top             =   315
      Width           =   45
   End
   Begin VB.Label lbVertices 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1095
      TabIndex        =   7
      Top             =   315
      Width           =   45
   End
End
Attribute VB_Name = "fTriangulate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

'drawing
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As tPoint2D, ByVal nCount As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As tPoint2D, ByVal nCount As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'for polygon and polyline api call
Private hDCCanvas           As Long
Private Type tPoint2D
    x   As Long
    y   As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Type MEMORYSTATUS
    dwLength                As Long
    dwMemoryLoad            As Long
    dwTotalPhys             As Long
    dwAvailPhys             As Long
    dwTotalPageFile         As Long
    dwAvailPageFile         As Long
    dwTotalVirtual          As Long
    dwAvailVirtual          As Long
End Type
Private Memstat             As MEMORYSTATUS

'the terrain
Private WithEvents Terrain  As cTerrain
Attribute Terrain.VB_VarHelpID = -1

Private DotRad              As Single 'red dot radius

'canvas origin and size
Private CanvasLeft          As Double
Private CanvasTop           As Double
Private CanvasWidth         As Double
Private CanvasHeight        As Double

Private LightZ              As Double 'light elevation

'some indicators
Private Rotated             As Boolean
Private Perspective         As Boolean
Private Internal            As Boolean

Private Function AvailableMemoryPercentage() As Long

    With Memstat
        .dwLength = Len(Memstat)
        GlobalMemoryStatus Memstat
        AvailableMemoryPercentage = .dwMemoryLoad
    End With 'MEMSTAT

End Function

Private Sub btRandom_Click()

  'creates random vertices in the terrain

  Dim i     As Long
  Dim n     As Long
  Dim x     As Double
  Dim y     As Double
  Dim Corners2D() As tPoint2D

    btReset_Click
    Enabled = False
    Screen.MousePointer = vbHourglass
    n = Val(txNum)
    If n >= 1 Then
        ReDim SortElems(1 To n)
        With picCanvas
            For i = 1 To n
                x = CDbl(Rnd * (CanvasWidth - 6)) + CanvasLeft + 3
                y = CDbl(Rnd * (CanvasHeight - 6)) + CanvasTop + 3
                If Terrain.AddVertex(x, y, Rnd * 100 - 50) = 0 Then
                    'If Terrain.AddVertex(x, y, Rnd * 100 - 2000) = 0 Then 'for test
                    .FillColor = vbRed
                    picCanvas.Circle (x, y), DotRad, vbRed
                End If
            Next i
        End With 'PICCANVAS

        If ckSort = vbChecked Then
            Terrain.PresortVertices
        End If

        lbVertices = Terrain.VertexCount
        DoEvents
        Render True
    End If
    Screen.MousePointer = vbDefault
    Enabled = True

End Sub

Private Sub btReset_Click()

  'what it says - it resets variables

    picCanvas.SetFocus
    Enabled = False
    Screen.MousePointer = vbHourglass
    Terrain.Reset
    Internal = True
    scrX = 180
    scrY = 180
    Internal = False
    picCanvas.Cls
    DrawGrid
    lbTriangles = vbNullString
    lbVertices = vbNullString
    lbTotalArea = vbNullString
    Screen.MousePointer = vbDefault
    Enabled = True

End Sub

Private Sub ckDefer_Click()

  'option to defer triangulation

    If ckDefer = vbUnchecked Then
        Render True
    End If
    picCanvas.SetFocus

End Sub

Private Sub ckMesh_Click()

  'option to defer draw wire mesh or filled triangles

    picCanvas.SetFocus
    If ckMesh = vbUnchecked Then
        Internal = True
        scrX = 180
        scrY = 180
        Internal = False
        scrX.Enabled = False
        scrY.Enabled = False
      Else 'NOT CKMESH...
        scrX.Enabled = True
        scrY.Enabled = True
    End If
    Render False

End Sub

Private Sub ckNumbers_Click()

  'option to show vertex nmbers

    If ckMesh = vbChecked Then
        picCanvas.SetFocus
        Render False
    End If

End Sub

Private Sub ckPerspective_Click()

  'option to add perspective distortion

    Perspective = (ckPerspective = vbChecked)
    picCanvas.SetFocus
    Render False

End Sub

Private Sub ckSort_Click()

  'option to presort vertices

    picCanvas.SetFocus

End Sub

Private Sub Render(NewSamples As Boolean)

  'render the triangulated terrain

  Dim Triangle          As cTriangle
  Dim i                 As Long
  Dim j                 As Long
  Dim k                 As Long
  Const a               As Long = 0
  Const b               As Long = 1
  Const c               As Long = 2
  Const aa              As Long = 3

  Dim Intensity         As Long
  Dim Corners2D(0 To 3) As tPoint2D

    Enabled = False
    If Not Rotated Then
        Screen.MousePointer = vbHourglass
    End If
    If ckDefer = vbUnchecked Then
        With Terrain
            If NewSamples Then
                lbTriangles = .Triangulate
                lbTotalArea = Format$(Round(.TotalArea, 5), "#,0.0####")
                If Not .HighestVertex Is Nothing Then
                    .SetLightPosition .HighestVertex.x, .HighestVertex.y, LightZ
                End If
            End If
            If .TriangleCount Then
                picCanvas.ForeColor = vbBlack
                picCanvas.DrawMode = vbCopyPen
                picCanvas.Cls
                For i = 1 To .TriangleCount
                    Set Triangle = .Triangles(i)
                    With Triangle
                        Corners2D(a) = MakePoint2D(.CornerA3D) '3 points for polygon drawing
                        Corners2D(b) = MakePoint2D(.CornerB3D)
                        Corners2D(c) = MakePoint2D(.CornerC3D)
                        If ckMesh = vbUnchecked Then 'colorize
                            'draw the shaded triangles
                            Intensity = 255 * Sqr(.LightIntensity)
                            picCanvas.FillColor = RGB(Intensity / 2, Intensity, Intensity / 2)
                            Polygon hDCCanvas, Corners2D(a), 3 'a-b-c and closed automatically
                          Else 'NOT CKMESH...
                            'draw the 3D wire mesh
                            Corners2D(aa) = Corners2D(a) '4th point for wire mesh
                            Polyline hDCCanvas, Corners2D(a), 4 'a-b-c-a
                            If Not (Rotated Or Perspective) Then
                                'draw the red dots and print vertex numbers
                                picCanvas.FillColor = vbRed
                                With .CornerA3D
                                    picCanvas.Circle (.x, .y), DotRad, vbRed
                                    If ckNumbers = vbChecked Then
                                        picCanvas.ForeColor = &HB00000
                                        picCanvas.Print .Number
                                    End If
                                End With '.CORNERA3D
                                With .CornerB3D
                                    picCanvas.Circle (.x, .y), DotRad, vbRed
                                    If ckNumbers = vbChecked Then
                                        picCanvas.Print .Number
                                    End If
                                End With '.CORNERB3D
                                With .CornerC3D
                                    picCanvas.Circle (.x, .y), DotRad, vbRed
                                    If ckNumbers = vbChecked Then
                                        picCanvas.Print .Number
                                        picCanvas.ForeColor = vbBlack
                                    End If
                                End With '.CORNERC3D
                            End If
                        End If
                    End With 'TRIANGLE
                Next i
                If ckMesh = vbUnchecked Then
                    'draw light
                    picCanvas.FillColor = vbYellow
                    picCanvas.Circle (.LightPosition.xRot, .LightPosition.yRot), DotRad + DotRad, vbRed
                End If
                DrawGrid
            End If
        End With 'TERRAIN
        pgb = 0
    End If
    Screen.MousePointer = vbDefault
    Enabled = True

End Sub

Private Sub DrawGrid()

  'draws the grid

  Dim i     As Long

    If Not Rotated Then
        picCanvas.DrawMode = vbMaskPen
        For i = CanvasLeft To CanvasWidth - CanvasLeft Step 10
            If i Mod 100 Then
                picCanvas.Line (i, CanvasTop)-(i, CanvasHeight - CanvasTop), &HE8E8E8
                picCanvas.Line (CanvasLeft, i)-(CanvasWidth - CanvasLeft, i), &HE8E8E8
              Else 'NOT I...
                picCanvas.Line (i, CanvasTop)-(i, CanvasHeight), &HD0D0D0
                picCanvas.Line (CanvasLeft, i)-(CanvasWidth, i), &HD0D0D0
            End If
        Next i
        lb(4) = "Memory used " & AvailableMemoryPercentage & "%"
    End If

End Sub

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Form_Load()

    Set Terrain = New cTerrain
    DotRad = 2
    LightZ = 500
    Rnd -1 'make sure the randomizer has initial seed
    With picCanvas
        hDCCanvas = GetDC(.hWnd)
        CanvasLeft = .ScaleLeft
        CanvasTop = .ScaleTop
        CanvasWidth = .ScaleWidth
        CanvasHeight = .ScaleHeight
    End With 'PICCANVAS

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    lbPosX = vbNullString
    lbPosY = vbNullString

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)

  'tidy up

    ReleaseDC picCanvas.hWnd, hDCCanvas
    Set Terrain = Nothing

End Sub

Private Function MakePoint2D(Vert As cVertex) As tPoint2D

  'makes a 2D point from a 3D vertex optionally taking into account perspective distortion
  'according to z-coord

  'the 2D points are then used in API-calls

  Dim PerspectiveDistortionFactor As Double
  Const Sqr2    As Double = 1.4142135623731

    With Vert
        If Perspective Then
            PerspectiveDistortionFactor = 1# + .zRot / CanvasWidth / Sqr2
            MakePoint2D.x = .xRot * PerspectiveDistortionFactor - CanvasLeft
            MakePoint2D.y = .yRot * PerspectiveDistortionFactor - CanvasTop
          Else 'PERSPECTIVE = FALSE/0
            MakePoint2D.x = .xRot - CanvasLeft
            MakePoint2D.y = .yRot - CanvasTop
        End If
    End With 'VERT

End Function

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  'adds a vertex to the vertex set or sets the light position

    Internal = True
    scrX = 180
    scrY = 180
    Internal = False
    If Button = vbLeftButton Then
        Terrain.AddVertex CDbl(x), CDbl(y), 0, False
        lbVertices = Terrain.VertexCount
        picCanvas.FillColor = vbRed
        picCanvas.Circle (x, y), DotRad, vbRed
        Render True
      Else 'NOT BUTTON...
        If ckMesh = vbUnchecked Then
            Terrain.SetLightPosition CDbl(x), CDbl(y), LightZ
        End If
        Render False
    End If

End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    lbPosX = "x " & Round(x)
    lbPosY = "y " & Round(y)

End Sub

Private Sub picCanvas_Paint()

    DrawGrid

End Sub

Private Sub scrX_Change()

  'rotate around x-axis - sets camera postion / viewing angle

    Rotated = Not Internal
    Terrain.SetCameraPosition scrX, scrY, 180
    If Rotated Then
        Render False
    End If

End Sub

Private Sub scrX_Scroll()

    scrX_Change

End Sub

Private Sub scrY_Change()

  'rotate around y-axis - sets camera postion / viewing angle

    Rotated = Not Internal
    Terrain.SetCameraPosition scrX, scrY, 180
    If Rotated Then
        Render False
    End If

End Sub

Private Sub scrY_Scroll()

    scrY_Change

End Sub

Private Sub Terrain_Progress(ByVal PercentCompleted As Long)

  'events fired during triangulation

    pgb = PercentCompleted
    lb(4) = "Memory used " & AvailableMemoryPercentage & "%"
    DoEvents

End Sub

Private Sub txNum_GotFocus()

    txNum.SelStart = 0
    txNum.SelLength = 4

End Sub

':) Ulli's VB Code Formatter V2.18.3 (2005-Jan-07 13:59)  Decl: 46  Code: 404  Total: 450 Lines
':) CommentOnly: 27 (6%)  Commented: 22 (4,9%)  Empty: 106 (23,6%)  Max Logic Depth: 10
