VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMesh 
   Caption         =   "Mesh Editor"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPoints 
      Caption         =   "Show Vertices"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Object"
      Height          =   375
      Left            =   7560
      TabIndex        =   20
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Object"
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   6360
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   3120
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save/Load Object"
      Filter          =   "BAA 3D files|*.b3d"
   End
   Begin VB.CommandButton cmdEditTriangle 
      Caption         =   "Edit"
      Height          =   375
      Left            =   8040
      TabIndex        =   18
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdRemoveTriangle 
      Caption         =   "Remove"
      Height          =   375
      Left            =   7080
      TabIndex        =   17
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdAddTriangle 
      Caption         =   "Add"
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   3240
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   16777215
      DialogTitle     =   "Triangle Color"
   End
   Begin VB.TextBox txtTVertex3 
      Height          =   285
      Left            =   8280
      TabIndex        =   15
      Text            =   "0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtTVertex2 
      Height          =   285
      Left            =   7680
      TabIndex        =   14
      Text            =   "0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtTVertex1 
      Height          =   285
      Left            =   7080
      TabIndex        =   13
      Text            =   "0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtZVertex 
      Height          =   285
      Left            =   8040
      TabIndex        =   11
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtYVertex 
      Height          =   285
      Left            =   7200
      TabIndex        =   10
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtXVertex 
      Height          =   285
      Left            =   6360
      TabIndex        =   9
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdEditVertex 
      Caption         =   "Edit"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRemoveVertex 
      Caption         =   "Remove"
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdAddVertex 
      Caption         =   "Add"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.ListBox lstTriangles 
      Height          =   2205
      Left            =   6360
      TabIndex        =   5
      Top             =   4080
      Width           =   2415
   End
   Begin VB.ListBox lstVertices 
      Height          =   2205
      Left            =   6360
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.PictureBox picFront 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   3240
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   3240
      Width           =   3015
   End
   Begin VB.PictureBox picFree 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   3240
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox picSide 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   3240
      Width           =   3015
   End
   Begin VB.PictureBox picTop 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblTColor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   3720
      Width           =   615
   End
End
Attribute VB_Name = "frmMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CRed, CGreen, CBlue, Color 'for color extraction
Dim Rx, Ry, Rz 'for free rotation
Dim mObject As TDObject
Dim LightVector As TDPoint
Dim View As TDPoint
Dim WorldX As Integer
Dim WorldY As Integer

Private Sub chkPoints_Click()
DrawObject
End Sub

Private Sub cmdAddTriangle_Click()
'add a triangle to the list
lstTriangles.AddItem lstTriangles.ListCount & ": " & txtTVertex1.Text & " " & txtTVertex2.Text & " " & txtTVertex3.Text
mObject.TriangleCount = mObject.TriangleCount + 1
mObject.Triangle(lstTriangles.ListCount - 1).Vertex(0) = txtTVertex1.Text
mObject.Triangle(lstTriangles.ListCount - 1).Vertex(1) = txtTVertex2.Text
mObject.Triangle(lstTriangles.ListCount - 1).Vertex(2) = txtTVertex3.Text
Color = lblTColor.BackColor
GetColors 'seperates Color into CRed, CGreen and CBlue
mObject.Triangle(lstTriangles.ListCount - 1).Color.Red = CRed
mObject.Triangle(lstTriangles.ListCount - 1).Color.Green = CGreen
mObject.Triangle(lstTriangles.ListCount - 1).Color.Blue = CBlue
DrawObject
End Sub

Private Sub cmdAddVertex_Click()
'add a vertex to the list
lstVertices.AddItem lstVertices.ListCount & ": " & txtXVertex.Text & " " & txtYVertex.Text & " " & txtZVertex.Text
mObject.VertexCount = mObject.VertexCount + 1
mObject.Vertex(lstVertices.ListCount - 1).X = txtXVertex.Text
mObject.Vertex(lstVertices.ListCount - 1).Y = txtYVertex.Text
mObject.Vertex(lstVertices.ListCount - 1).Z = txtZVertex.Text
DrawObject
End Sub

Private Sub cmdEditTriangle_Click()
'change the value of a triangle
If lstTriangles.ListIndex = -1 Then Exit Sub
mObject.Triangle(lstTriangles.ListIndex).Vertex(0) = txtTVertex1.Text
mObject.Triangle(lstTriangles.ListIndex).Vertex(1) = txtTVertex2.Text
mObject.Triangle(lstTriangles.ListIndex).Vertex(2) = txtTVertex3.Text
Color = lblTColor.BackColor
GetColors
mObject.Triangle(lstTriangles.ListIndex).Color.Red = CRed
mObject.Triangle(lstTriangles.ListIndex).Color.Green = CGreen
mObject.Triangle(lstTriangles.ListIndex).Color.Blue = CBlue
lstTriangles.Clear
For i = 0 To mObject.TriangleCount - 1
lstTriangles.AddItem lstTriangles.ListCount & ": " & mObject.Triangle(i).Vertex(0) & " " & mObject.Triangle(i).Vertex(1) & " " & mObject.Triangle(i).Vertex(2)
Next i
DrawObject
End Sub

Private Sub cmdEditVertex_Click()
'change the value of a vertex
If lstVertices.ListIndex = -1 Then Exit Sub
mObject.Vertex(lstVertices.ListIndex).X = txtXVertex.Text
mObject.Vertex(lstVertices.ListIndex).Y = txtYVertex.Text
mObject.Vertex(lstVertices.ListIndex).Z = txtZVertex.Text
lstVertices.Clear
For i = 0 To mObject.VertexCount - 1
lstVertices.AddItem lstVertices.ListCount & ": " & mObject.Vertex(i).X & " " & mObject.Vertex(i).Y & " " & mObject.Vertex(i).Z
Next i
DrawObject
End Sub

Private Sub cmdLoad_Click()
CommonDialog2.ShowOpen
Open CommonDialog2.filename For Binary As 1
Get #1, 1, mObject
Close #1
'refill all the lists and redraw the pictures
lstVertices.Clear
lstTriangles.Clear
For i = 0 To mObject.VertexCount - 1
lstVertices.AddItem lstVertices.ListCount & ": " & mObject.Vertex(i).X & " " & mObject.Vertex(i).Y & " " & mObject.Vertex(i).Z
Next i
For i = 0 To mObject.TriangleCount - 1
lstTriangles.AddItem lstTriangles.ListCount & ": " & mObject.Triangle(i).Vertex(0) & " " & mObject.Triangle(i).Vertex(1) & " " & mObject.Triangle(i).Vertex(2)
Next i
DrawObject
End Sub

Private Sub cmdRemoveTriangle_Click()
'remove a triangle by shifting all triangles above it down one
If lstTriangles.ListIndex = -1 Then Exit Sub

For i = lstTriangles.ListIndex + 1 To lstTriangles.ListCount
mObject.Triangle(i - 1).Vertex(0) = mObject.Triangle(i).Vertex(0)
mObject.Triangle(i - 1).Vertex(1) = mObject.Triangle(i).Vertex(1)
mObject.Triangle(i - 1).Vertex(2) = mObject.Triangle(i).Vertex(2)

mObject.Triangle(i - 1).Color.Red = mObject.Triangle(i).Color.Red
mObject.Triangle(i - 1).Color.Green = mObject.Triangle(i).Color.Green
mObject.Triangle(i - 1).Color.Blue = mObject.Triangle(i).Color.Blue
Next i

mObject.TriangleCount = mObject.TriangleCount - 1
lstTriangles.Clear
For i = 0 To mObject.TriangleCount - 1
lstTriangles.AddItem lstTriangles.ListCount & ": " & mObject.Triangle(i).Vertex(0) & " " & mObject.Triangle(i).Vertex(1) & " " & mObject.Triangle(i).Vertex(2)
Next i

DrawObject
End Sub

Private Sub cmdRemoveVertex_Click()
'remove a vertex
If lstVertices.ListIndex = -1 Then Exit Sub
For i = 0 To mObject.TriangleCount - 1
'shift all points down in the triangles
If mObject.Triangle(i).Vertex(0) > lstVertices.ListIndex Then mObject.Triangle(i).Vertex(0) = mObject.Triangle(i).Vertex(0) - 1
If mObject.Triangle(i).Vertex(1) > lstVertices.ListIndex Then mObject.Triangle(i).Vertex(1) = mObject.Triangle(i).Vertex(1) - 1
If mObject.Triangle(i).Vertex(2) > lstVertices.ListIndex Then mObject.Triangle(i).Vertex(2) = mObject.Triangle(i).Vertex(2) - 1
Next i
For i = lstVertices.ListIndex + 1 To lstVertices.ListCount
mObject.Vertex(i - 1).X = mObject.Vertex(i).X
mObject.Vertex(i - 1).Y = mObject.Vertex(i).Y
mObject.Vertex(i - 1).Z = mObject.Vertex(i).Z
Next i

mObject.VertexCount = mObject.VertexCount - 1
lstVertices.Clear
For i = 0 To mObject.VertexCount - 1
lstVertices.AddItem lstVertices.ListCount & ": " & mObject.Vertex(i).X & " " & mObject.Vertex(i).Y & " " & mObject.Vertex(i).Z
Next i
lstTriangles.Clear
For i = 0 To mObject.TriangleCount - 1
lstTriangles.AddItem lstTriangles.ListCount & ": " & mObject.Triangle(i).Vertex(0) & " " & mObject.Triangle(i).Vertex(1) & " " & mObject.Triangle(i).Vertex(2)
Next i

DrawObject
End Sub

Private Sub DrawObject()
If chkPoints.Value = 1 Then 'send to the point draw sub
DrawPoints
Exit Sub
End If
Dim Normal As TDPoint 'the normal of a polygon
Dim Light As Double 'the lighting of a polygon
Dim RedLight As Integer 'the red light hitting a surface
Dim GreenLight As Integer 'the green light hitting a surface
Dim BlueLight As Integer 'guess
Dim Ambience As Integer 'minimum light
Dim Rotated As TDObject 'the rotated object
ReDim PointList(0 To 2) As PointAPI 'list of points to send to Polygon
Ambience = 32
picTop.Cls 'clear out the surfaces
picFront.Cls
picSide.Cls
picFree.Cls
'I only commented the first section that draws the top view, because
'the rest are pretty much the same except rendered to different windows
BuildMatrix 90, 0, 0 'build rotation (x, y, z degrees)
Rotated = RotateObject(mObject, LightVector, View) 'rotates the object
If lstTriangles.ListIndex <> -1 Then 'highlight selected polygon
    Rotated.Triangle(lstTriangles.ListIndex).Color.Red = 255 'make the selected
    Rotated.Triangle(lstTriangles.ListIndex).Color.Green = 0 'polygon red
    Rotated.Triangle(lstTriangles.ListIndex).Color.Blue = 0
End If
SortTriangles Rotated 'sort them from furthest to closest to draw them correctly

For i = 0 To Rotated.TriangleCount - 1 'draw each triangle
    If Rotated.Triangle(i).Visible = True Then 'only if visible
        'fill out the PointList, which is what we have to use the Polygon call with
        PointList(0).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(0)).X + WorldX
        PointList(0).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(0)).Y + WorldY
        PointList(1).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(1)).X + WorldX
        PointList(1).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(1)).Y + WorldY
        PointList(2).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(2)).X + WorldX
        PointList(2).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(2)).Y + WorldY
        Light = Rotated.Triangle(i).Shade 'grab the shade value
        'figure out the red, green and blue shades
        'Shading is a value from 0 to 1.  So I subtracted .5 to make it
        '-.5 to .5 and multiplied to get -128 to 128 range, which is the
        'maximum color difference you can get.  Ambience makes sure it doesn't
        'go completely black
        RedLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Red + Ambience
        GreenLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Green + Ambience
        BlueLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Blue + Ambience
        'some limit checking
        If RedLight > 255 Then RedLight = 255
        If GreenLight > 255 Then GreenLight = 255
        If BlueLight > 255 Then BlueLight = 255
        If RedLight < 0 Then RedLight = 0
        If GreenLight < 0 Then GreenLight = 0
        If BlueLight < 0 Then BlueLight = 0
        'se the fill and fore color to the shade values
        picTop.FillColor = RGB(RedLight, GreenLight, BlueLight)
        picTop.ForeColor = picTop.FillColor
        Polygon picTop.HDC, PointList(0), 3 'draw that polygon
    End If
Next i

BuildMatrix 0, 0, 0
Rotated = RotateObject(mObject, LightVector, View)
If lstTriangles.ListIndex <> -1 Then 'highlight selected polygon
Rotated.Triangle(lstTriangles.ListIndex).Color.Red = 255
Rotated.Triangle(lstTriangles.ListIndex).Color.Green = 0
Rotated.Triangle(lstTriangles.ListIndex).Color.Blue = 0
End If
SortTriangles Rotated
For i = 0 To Rotated.TriangleCount - 1
If Rotated.Triangle(i).Visible = True Then
PointList(0).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(0)).X + WorldX
PointList(0).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(0)).Y + WorldY
PointList(1).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(1)).X + WorldX
PointList(1).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(1)).Y + WorldY
PointList(2).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(2)).X + WorldX
PointList(2).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(2)).Y + WorldY
Light = Rotated.Triangle(i).Shade
RedLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Red + Ambience
GreenLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Green + Ambience
BlueLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Blue + Ambience
If RedLight > 255 Then RedLight = 255
If GreenLight > 255 Then GreenLight = 255
If BlueLight > 255 Then BlueLight = 255
If RedLight < 0 Then RedLight = 0
If GreenLight < 0 Then GreenLight = 0
If BlueLight < 0 Then BlueLight = 0

picFront.FillColor = RGB(RedLight, GreenLight, BlueLight)

picFront.ForeColor = picFront.FillColor
Polygon picFront.HDC, PointList(0), 3
End If
Next i

BuildMatrix 0, 90, 0
Rotated = RotateObject(mObject, LightVector, View)
If lstTriangles.ListIndex <> -1 Then 'highlight selected polygon
Rotated.Triangle(lstTriangles.ListIndex).Color.Red = 255
Rotated.Triangle(lstTriangles.ListIndex).Color.Green = 0
Rotated.Triangle(lstTriangles.ListIndex).Color.Blue = 0
End If
SortTriangles Rotated
For i = 0 To Rotated.TriangleCount - 1
If Rotated.Triangle(i).Visible = True Then
PointList(0).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(0)).X + WorldX
PointList(0).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(0)).Y + WorldY
PointList(1).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(1)).X + WorldX
PointList(1).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(1)).Y + WorldY
PointList(2).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(2)).X + WorldX
PointList(2).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(2)).Y + WorldY
Light = Rotated.Triangle(i).Shade
RedLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Red + Ambience
GreenLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Green + Ambience
BlueLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Blue + Ambience
If RedLight > 255 Then RedLight = 255
If GreenLight > 255 Then GreenLight = 255
If BlueLight > 255 Then BlueLight = 255
If RedLight < 0 Then RedLight = 0
If GreenLight < 0 Then GreenLight = 0
If BlueLight < 0 Then BlueLight = 0

picSide.FillColor = RGB(RedLight, GreenLight, BlueLight)
picSide.ForeColor = picSide.FillColor
Polygon picSide.HDC, PointList(0), 3
End If
Next i

BuildMatrix Rx, Ry, Rz
Rotated = RotateObject(mObject, LightVector, View)
If lstTriangles.ListIndex <> -1 Then 'highlight selected polygon
Rotated.Triangle(lstTriangles.ListIndex).Color.Red = 255
Rotated.Triangle(lstTriangles.ListIndex).Color.Green = 0
Rotated.Triangle(lstTriangles.ListIndex).Color.Blue = 0
End If
SortTriangles Rotated
For i = 0 To Rotated.TriangleCount - 1
If Rotated.Triangle(i).Visible = True Then
PointList(0).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(0)).X + WorldX
PointList(0).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(0)).Y + WorldY
PointList(1).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(1)).X + WorldX
PointList(1).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(1)).Y + WorldY
PointList(2).X = Rotated.Vertex(Rotated.Triangle(i).Vertex(2)).X + WorldX
PointList(2).Y = Rotated.Vertex(Rotated.Triangle(i).Vertex(2)).Y + WorldY
Light = Rotated.Triangle(i).Shade
RedLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Red + Ambience
GreenLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Green + Ambience
BlueLight = (Light - 0.5) * 256 + Rotated.Triangle(i).Color.Blue + Ambience
If RedLight > 255 Then RedLight = 255
If GreenLight > 255 Then GreenLight = 255
If BlueLight > 255 Then BlueLight = 255
If RedLight < 0 Then RedLight = 0
If GreenLight < 0 Then GreenLight = 0
If BlueLight < 0 Then BlueLight = 0

picFree.FillColor = RGB(RedLight, GreenLight, BlueLight)
picFree.ForeColor = picFree.FillColor
Polygon picFree.HDC, PointList(0), 3
End If
Next i
picFront.Refresh
picTop.Refresh
picSide.Refresh
picFree.Refresh
End Sub

Private Sub cmdSave_Click()
CommonDialog2.ShowSave
Open CommonDialog2.filename For Binary As 1
Put #1, 1, mObject
Close #1
End Sub

Private Sub Form_Load()
WorldX = 96
WorldY = 96
BuildLookup
LightVector.X = 64
LightVector.Y = 64
LightVector.Z = -64
View.X = 0
View.Y = 0
View.Z = -1
LightVector = Normalize(LightVector)
End Sub

Private Sub lblTColor_Click()
CommonDialog1.ShowColor
lblTColor.BackColor = CommonDialog1.Color
End Sub

Private Sub GetColors()
CRed = Int(Color Mod 256)
CBlue = Int(Color / 65536)
CGreen = Int((Color - (CBlue * 65536) - CRed) / 256)
End Sub


Private Sub lstTriangles_Click()
DrawObject
End Sub

Private Sub lstTriangles_DblClick()
txtTVertex1.Text = mObject.Triangle(lstTriangles.ListIndex).Vertex(0)
txtTVertex2.Text = mObject.Triangle(lstTriangles.ListIndex).Vertex(1)
txtTVertex3.Text = mObject.Triangle(lstTriangles.ListIndex).Vertex(2)
lblTColor.BackColor = RGB(mObject.Triangle(lstTriangles.ListIndex).Color.Red, _
        mObject.Triangle(lstTriangles.ListIndex).Color.Green, _
        mObject.Triangle(lstTriangles.ListIndex).Color.Blue)
End Sub

Private Sub lstVertices_Click()
DrawObject
End Sub

Private Sub lstVertices_DblClick()
txtXVertex.Text = mObject.Vertex(lstVertices.ListIndex).X
txtYVertex.Text = mObject.Vertex(lstVertices.ListIndex).Y
txtZVertex.Text = mObject.Vertex(lstVertices.ListIndex).Z
End Sub

Private Sub picFree_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static OldX, OldY
If Button = 1 Then
Rx = Rx + (OldY - Y) / 3
Ry = Ry + (OldX - X) / 3
DrawObject
End If
OldX = X
OldY = Y

End Sub

Private Sub DrawPoints()
Dim Rotated As TDObject

picTop.Cls
picFront.Cls
picSide.Cls
picFree.Cls

BuildMatrix 90, 0, 0
Rotated = RotateObject(mObject, LightVector, View)
For i = 0 To Rotated.VertexCount - 1
If i = lstVertices.ListIndex Then
SetPixel picTop.HDC, Rotated.Vertex(i).X + WorldX, Rotated.Vertex(i).Y + WorldY, RGB(255, 0, 0)
Else
SetPixel picTop.HDC, Rotated.Vertex(i).X + WorldX, Rotated.Vertex(i).Y + WorldY, RGB(255, 255, 255)
End If
Next i

BuildMatrix 0, 0, 0
Rotated = RotateObject(mObject, LightVector, View)
For i = 0 To Rotated.VertexCount - 1
If i = lstVertices.ListIndex Then
SetPixel picFront.HDC, Rotated.Vertex(i).X + WorldX, Rotated.Vertex(i).Y + WorldY, RGB(255, 0, 0)
Else
SetPixel picFront.HDC, Rotated.Vertex(i).X + WorldX, Rotated.Vertex(i).Y + WorldY, RGB(255, 255, 255)
End If
Next i

BuildMatrix 0, 90, 0
Rotated = RotateObject(mObject, LightVector, View)
For i = 0 To Rotated.VertexCount - 1
If i = lstVertices.ListIndex Then
SetPixel picSide.HDC, Rotated.Vertex(i).X + WorldX, Rotated.Vertex(i).Y + WorldY, RGB(255, 0, 0)
Else
SetPixel picSide.HDC, Rotated.Vertex(i).X + WorldX, Rotated.Vertex(i).Y + WorldY, RGB(255, 255, 255)
End If
Next i

BuildMatrix Rx, Ry, Rz
Rotated = RotateObject(mObject, LightVector, View)
For i = 0 To Rotated.VertexCount - 1
If i = lstVertices.ListIndex Then
SetPixel picFree.HDC, Rotated.Vertex(i).X + WorldX, Rotated.Vertex(i).Y + WorldY, RGB(255, 0, 0)
Else
SetPixel picFree.HDC, Rotated.Vertex(i).X + WorldX, Rotated.Vertex(i).Y + WorldY, RGB(255, 255, 255)
End If
Next i

End Sub
