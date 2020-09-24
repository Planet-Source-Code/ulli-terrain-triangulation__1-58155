Attribute VB_Name = "mMisc"
'Sorting and some common variables / constants

Option Explicit

Public Const Infinity       As Double = 10 ^ 11
Public Const InvInfinity    As Double = 1 / Infinity

Public Type tVertex
    X   As Double
    Y   As Double
    Z   As Double
    ii  As Boolean
End Type

Public LightPos             As cVertex 'light position

Public RotationMatrix(1 To 9) As Double 'camera viewing angle transformation matrix

Public Enum SortType
    SortByX = 1
    SortByAngle = 2 'not yet used
End Enum
#If False Then
Private SortByX, SortByAngle
#End If

Public Const ErrorSource    As String = "Triangulation"
Public SortElems()          As tVertex
Private TempElem            As tVertex

Private Function Compare(ByVal sType As SortType, p1 As tVertex, p2 As tVertex) As Long

    Select Case sType
      Case SortByX
        'returns +1 if p1.x > p2.x
        'returns 0 if p1.x = p2.x
        'returns -1 if p1.x < p2.x
        Compare = Sgn(p1.X - p2.X)
    End Select

End Function

Public Sub QuickSort(ByVal sType As SortType, Optional ByVal xFrom As Long = 0, Optional ByVal xThru As Long = 0)

  'sorts a table of Vertices

  Dim xLeft As Long, xRite As Long

    If xFrom < xThru Then 'we have something to sort (@ least two elements)
        xLeft = xFrom
        xRite = xThru
        TempElem = SortElems(xLeft) 'get ref element and make room
        Do
            Do Until xRite = xLeft
                If Compare(sType, SortElems(xRite), TempElem) = -1 Then
                    'is smaller than ref so move it to the left...
                    SortElems(xLeft) = SortElems(xRite)
                    xLeft = xLeft + 1 '...and leave the item just moved alone for now
                    Exit Do 'loop 
                  Else 'NOT COMPARE(STYPE,...
                    xRite = xRite - 1
                End If
            Loop
            Do Until xLeft = xRite
                If Compare(sType, SortElems(xLeft), TempElem) = 1 Then
                    'is greater than ref so move it to the right...
                    SortElems(xRite) = SortElems(xLeft)
                    xRite = xRite - 1 '...and leave the item just moved alone for now
                    Exit Do 'loop 
                  Else 'NOT COMPARE(STYPE,...
                    xLeft = xLeft + 1
                End If
            Loop
        Loop Until xLeft = xRite
        DoEvents
        'now the indexes have met and all bigger items are to the right and all smaller items are left
        SortElems(xRite) = TempElem 'insert ref elem in proper place and sort the two areas left and right of it
        If xLeft - xFrom > xThru - xRite Then 'smaller part 1st to reduce recursion depth
            QuickSort sType, xRite + 1, xThru
            QuickSort sType, xFrom, xLeft - 1
          Else 'NOT XLEFT...
            QuickSort sType, xFrom, xLeft - 1
            QuickSort sType, xRite + 1, xThru
        End If
    End If

End Sub

':) Ulli's VB Code Formatter V2.18.3 (2005-Jan-07 14:09)  Decl: 29  Code: 61  Total: 90 Lines
':) CommentOnly: 8 (8,9%)  Commented: 12 (13,3%)  Empty: 15 (16,7%)  Max Logic Depth: 5
