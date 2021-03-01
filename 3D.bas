Attribute VB_Name = "Mod3D"
'3D Polygon Module
'By Alan Buzbee
'If you want to use this module for anything, here is a brief list of what you need
'to do to initialize and use it.
'First, make a call to the BuildLookup() sub.  This only needs to be done once,
'and it simply initializes a lookup table to make things run faster.
'You need to have a TDObject filled out, using the MeshEdit program, loaded from
'a file or whatever.

'Before you rotate, send the rotation values to the BuildMatrix
'sub.  It builds the rotation matrix.  This only has to be done when you are rotating
'by different angles.  If you wish to rotate every object by 45 degrees x, 20 y
'and 15 z, you only have to make one BuildMatrix 45, 20, 15 call beforehand
'after that you can make your calls to RotateObject.

'RotateMatrix parameters are the object
'it needs to rotate from, a viewing vector to determine which surfaces are visible,
'and a Light vector to determine shading.  The viewing vector is a TDPoint that
'is a point that the camera is looking directly at.  Usually it's best to define
'this as 0, 0, -1.  The Light vector is a point that light is eminating from.
'when you define the Light and View point, be sure to call Normalize on
'them before using them.

'One last thing before you begin drawing:  For more complex objects with peices
'jutting out and whatnot, it's good to make a call to SortTriangles.  This sorts all
'the triangles in a rotated polygon and sorts them from furthest to closest.  This
'way when you draw the triangles the further ones back will be covered by the ones
'in front so it looks normal.

'Right now this module only supports objects up to 50 vertices and 50 polygons.
'this can be changed by changing the array values in the TDObject.  I left a
'brief description of each function right above it in case you have any questions,
'otherwise just E-mail any questions or comments to Alan_Buzbee@hotmail.com

'One special note:  If you are defining triangles and they are facing the wrong
'way than what you want, it means you are doing them in the wrong order.  Specifying
'vertices in a clockwise order will have the triangle face one way, while
'counterclockwise will face the other.  To fix a backward polygon just switch
'the first and last vertices.

Public Type TDPoint 'A single 3D point
X As Double
Y As Double
Z As Double
End Type

Public Type PointAPI 'A 2D point for polygon calls
X As Long
Y As Long
End Type

Public Type TColor 'a polygon color (for faster shading)
Red As Byte
Green As Byte
Blue As Byte
End Type

Public Type TDTriangle 'a 3D triangle
Vertex(0 To 2) As Integer
Color As TColor
Visible As Integer
Shade As Double
End Type

Public Type TDObject 'a full 3D object
VertexCount As Integer
TriangleCount As Integer
Vertex(0 To 50) As TDPoint
Triangle(0 To 50) As TDTriangle
End Type

'some important DLL functions
Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, ByVal dx As Long, ByVal dy As Long, ByVal Color As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal dx As Long, ByVal dy As Long) As Long
Declare Function Polygon Lib "gdi32" _
(ByVal HDC As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long

Global RM(0 To 3, 0 To 3) As Double 'the rotation matrix
Global Sine(0 To 359) As Double 'array of sine values
Global Cosine(0 To 359) As Double 'array of cosine values
Global Const pi = 3.14159265358979 'pi!
Global Const Deg2Rad = pi / 180 'for a little more speed

Public Sub BuildLookup()
'Build the lookup table for all Sines and Cosines of 0 to 359 degrees
For i = 0 To 359
Sine(i) = Sin(i * Deg2Rad)
Cosine(i) = Cos(i * Deg2Rad)
Next i
End Sub

Public Sub BuildMatrix(ByVal Ax As Integer, ByVal Ay As Integer, ByVal Az As Integer)
' this sub builds the rotation matrix with x, y and z as axis angles
Dim SinX, CosX, SinY, CosY, SinZ, CosZ

'since we're using a lookup table, we have to make sure our angles are
'between 0 and 359.  If you're sending values above 32700 or below -32700 then
'it will crash because of integer use, but you can check for that yourself.
While Ax > 359
Ax = Ax - 360
Wend
While Ax < 0
Ax = Ax + 360
Wend
While Ay > 359
Ay = Ay - 360
Wend
While Ay < 0
Ay = Ay + 360
Wend
While Az > 359
Az = Az - 360
Wend
While Az < 0
Az = Az + 360
Wend
'obtain the values we need
SinX = Sine(Ax)
CosX = Cosine(Ax)
SinY = Sine(Ay)
CosY = Cosine(Ay)
SinZ = Sine(Az)
CosZ = Cosine(Az)
'fill out the rotation matrix.  Now we can multiply any point by this matrix
'to rotate it by the current set of angles!
RM(0, 0) = (CosZ * CosY)
RM(0, 1) = (CosZ * -SinY * -SinX + SinZ * CosX)
RM(0, 2) = (CosZ * -SinY * CosX + SinZ * SinX)
RM(1, 0) = (-SinZ * CosY)
RM(1, 1) = (-SinZ * -SinY * -SinX + CosZ * CosX)
RM(1, 2) = (-SinZ * -SinY * CosX + CosZ * SinX)
RM(2, 0) = SinY
RM(2, 1) = CosY * -SinX
RM(2, 2) = CosY * CosX
End Sub

Public Function RotatePoint(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As TDPoint
'rotate the point by the current matrix
Dim ZPoint As TDPoint
'This is how you multiply by a matrix.  Just don't ask.
ZPoint.X = (X * RM(0, 0)) + (Y * RM(0, 1)) + (Z * RM(0, 2)) + RM(0, 3)
ZPoint.Y = (X * RM(1, 0)) + (Y * RM(1, 1)) + (Z * RM(1, 2)) + RM(1, 3)
ZPoint.Z = (X * RM(2, 0)) + (Y * RM(2, 1)) + (Z * RM(2, 2)) + RM(2, 3)
RotatePoint = ZPoint
End Function

Public Function RotateObject(tObject As TDObject, Light As TDPoint, View As TDPoint) As TDObject
'rotate an entire object, and then shade and set the visibility values
Dim tmpObject As TDObject 'object to change
Dim Normal As TDPoint 'normal of the current polygon
Dim Luminance As Double 'brightness
Dim Visibility As Double ' if the polygon is visible or not
tmpObject = tObject 'set the object to a temporary one
For i = 0 To tmpObject.VertexCount
'rotate all the vertices
tmpObject.Vertex(i) = RotatePoint(tmpObject.Vertex(i).X, tmpObject.Vertex(i).Y, tmpObject.Vertex(i).Z)
Next i
For i = 0 To tmpObject.TriangleCount 'shade and check visibility for all the triangles
Normal = GetNormal(tmpObject.Vertex(tmpObject.Triangle(i).Vertex(0)), _
    tmpObject.Vertex(tmpObject.Triangle(i).Vertex(1)), _
    tmpObject.Vertex(tmpObject.Triangle(i).Vertex(2)))
Normal = Normalize(Normal) 'normalizing saves us lots of trouble
Visibility = DotProduct(Normal, View) 'finds the angle between the view and the triangle
If Visibility >= 0 Then 'only figure stuff out for visible gons
Luminance = DotProduct(Normal, Light)
tmpObject.Triangle(i).Shade = Abs(Luminance)
tmpObject.Triangle(i).Visible = True
Else
tmpObject.Triangle(i).Visible = False
End If
Next i
RotateObject = tmpObject 'send back the data!
End Function
Public Function DotProduct(vector1 As TDPoint, vector2 As TDPoint) As Double
'the Dot Product between two points, for finding if a surface is visible, etc
'You shouldn't have to worry about this, RotateObject makes all the calls to here
DotProduct = vector1.X * vector2.X + vector1.Y * vector2.Y + vector1.Z * vector2.Z
End Function

Public Function AbsV(vector As TDPoint) As Double
'absolute value of a 3D point
'Don't worry about this, it's used to Normalize a vector in the Normalize call
AbsV = (vector.X ^ 2 + vector.Y ^ 2 + vector.Z ^ 2) ^ 0.5
End Function

Public Function GetNormal(Point1 As TDPoint, Point2 As TDPoint, Point3 As TDPoint) As TDPoint
'get the normal of a surface
'taken care of by the RotateObject function for determining visibility and shading
Dim vector1 As TDPoint
Dim vector2 As TDPoint
Dim tmpPoint As TDPoint

vector1.X = Point2.X - Point1.X 'find two vectors of the polygon
vector1.Y = Point2.Y - Point1.Y
vector1.Z = Point2.Z - Point1.Z
vector2.X = Point3.X - Point1.X
vector2.Y = Point3.Y - Point1.Y
vector2.Z = Point3.Z - Point1.Z

tmpPoint.X = vector1.Y * vector2.Z - vector1.Z * vector2.Y
tmpPoint.Y = vector1.X * vector2.Z - vector1.Z * vector2.X
tmpPoint.Z = vector1.X * vector2.Y - vector1.Y * vector2.X
GetNormal = tmpPoint
End Function

Public Function Normalize(vector As TDPoint) As TDPoint
'set a vector to a length of 1
'the only time you need to worry about this is when you make
'a light or viewing vector.  If you don't normalize them then you may
'get some funny happenings.
Dim tmpPoint As TDPoint
Dim length As Double
tmpPoint = vector
length = AbsV(tmpPoint)
If length < 0.05 Then length = 0.05
tmpPoint.X = tmpPoint.X / length
tmpPoint.Y = tmpPoint.Y / length
tmpPoint.Z = tmpPoint.Z / length
Normalize = tmpPoint
End Function

Public Sub SortTriangles(tObject As TDObject)
'use an insertion sort to sort triangles from furthest to closest
'Use it, but don't try to understand it, unless you really like
'figuring out sorting algorithms.
'Actually, in most 3D games you have to do an object level sort so that
'all the objects are drawn in order.  If that's what you're going for,
'you'll probably need to know how this, or some other sorting method, works.
Dim j As Integer
Dim v As TDTriangle
Dim AvgZ1 As Double
Dim AvgZ2 As Double

For i = 1 To tObject.TriangleCount - 1 'check every triangle
v = tObject.Triangle(i) 'current triangle
j = i 'current position
AvgZ1 = (tObject.Vertex(tObject.Triangle(j - 1).Vertex(0)).Z + _
    tObject.Vertex(tObject.Triangle(j - 1).Vertex(1)).Z + _
    tObject.Vertex(tObject.Triangle(j - 1).Vertex(2)).Z) / 3 'get average z's
AvgZ2 = (tObject.Vertex(v.Vertex(0)).Z + _
    tObject.Vertex(v.Vertex(1)).Z + _
    tObject.Vertex(v.Vertex(2)).Z) / 3
Do While AvgZ1 > AvgZ2
tObject.Triangle(j) = tObject.Triangle(j - 1) 'shift to the right
j = j - 1 'move down
If j = 0 Then Exit Do 'if zero then we're done with this one
AvgZ1 = (tObject.Vertex(tObject.Triangle(j - 1).Vertex(0)).Z + _
    tObject.Vertex(tObject.Triangle(j - 1).Vertex(1)).Z + _
    tObject.Vertex(tObject.Triangle(j - 1).Vertex(2)).Z) / 3 'recalculate average z's
AvgZ2 = (tObject.Vertex(v.Vertex(0)).Z + _
    tObject.Vertex(v.Vertex(1)).Z + _
    tObject.Vertex(v.Vertex(2)).Z) / 3
Loop
tObject.Triangle(j) = v 'set to the triangle
Next i
End Sub
