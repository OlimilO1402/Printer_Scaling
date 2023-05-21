Attribute VB_Name = "MPGeom"
Option Explicit

Public Type Point
    X As Double
    Y As Double
    Z As Double
End Type

Public Type Size
    Width  As Double
    Height As Double
    Depth  As Double
End Type

Public Type AARect 'Axis-aligned
    Pt As Point
    Sz As Size
End Type

Public Type PPRect '2 Points
    P1 As Point
    P2 As Point
End Type

' v ############################## v '    Point    ' v ############################## v '
Public Function New_Point(ByVal X As Double, ByVal Y As Double, Optional ByVal Z As Double = 0#) As Point
    With New_Point: .X = X: .Y = Y: .Z = Z: End With
End Function

Public Function Point_ToStr(this As Point) As String
    Point_ToStr = "Point(X: " & this.X & "; Y: " & this.Y & "; Z: " & this.Z & ")"
End Function

Public Function Point_ToSize(this As Point) As Size
    With Point_ToSize: .Height = this.X: .Width = this.Y: End With
End Function

Public Function Point_Add(this As Point, other As Point) As Point
    With Point_Add: .X = this.X + other.X: .Y = this.Y + other.Y: .Z = this.Z + other.Z: End With
End Function

Public Function Point_Dif(this As Point, other As Point) As Point
    With Point_Dif: .X = this.X - other.X: .Y = this.Y - other.Y: .Z = this.Z - other.Z: End With
End Function

Public Function Point_Mul(this As Point, ByVal Scalar As Double) As Point
    With Point_Mul: .X = this.X * Scalar: .Y = this.Y * Scalar: .Z = this.Z * Scalar: End With
End Function

Public Function Point_Div(this As Point, ByVal Divisor As Double) As Point
    With Point_Div: .X = this.X / Divisor: .Y = this.Y / Divisor: .Z = this.Z / Divisor: End With
End Function

Public Function Point_Neg(this As Point) As Point
    With Point_Neg: .X = -this.X: .Y = -this.Y: .Z = -this.Z: End With
End Function

Public Function Point_Equals(this As Point, other As Point) As Boolean
    With this: Point_Equals = (.X = other.X) And (.Y = other.Y): End With
End Function
' ^ ############################## ^ '    Point    ' ^ ############################## ^ '

' v ############################## v '    Size     ' v ############################## v '
Public Function New_Size(ByVal Width As Double, ByVal Height As Double, Optional ByVal Depth As Double = 0#) As Size
    With New_Size: .Width = Width: .Height = Height: .Depth = Depth: End With
End Function

Public Function Size_Mul(this As Size, ByVal Scalar As Double) As Size
    With Size_Mul: .Width = this.Width * Scalar: .Height = this.Height * Scalar: .Depth = this.Depth * Scalar: End With
End Function

Public Function Size_ToStr(this As Size) As String
    Size_ToStr = "Size(W: " & this.Width & "; H: " & this.Height & "; D: " & this.Depth & ")"
End Function

Public Function Size_ToPoint(this As Size) As Point
    With Size_ToPoint: .X = this.Width: .Y = this.Height: .Z = this.Depth: End With
End Function

Public Function Size_Equals(this As Size, other As Size) As Boolean
    With this: Size_Equals = (.Width = other.Width) And (.Height = other.Height) And (.Depth = other.Depth): End With
End Function
' ^ ############################## ^ '    Size     ' ^ ############################## ^ '

' v ############################## v '   AARect    ' v ############################## v '
Public Function New_AARect(P As Point, S As Size) As AARect
    With New_AARect: .Pt = P: .Sz = S: End With
End Function

Public Function AARect_Mul(this As AARect, ByVal Scalar As Double) As AARect
    AARect_Mul = New_AARect(Point_Mul(this.Pt, Scalar), Size_Mul(this.Sz, Scalar))
End Function

Public Function AARect_ToStr(this As AARect) As String
    AARect_ToStr = "AARect(Pt: " & Point_ToStr(this.Pt) & "; Sz: " & Size_ToStr(this.Sz)
End Function

Public Function AARect_ToPPRect(this As AARect) As PPRect
    AARect_ToPPRect = New_PPRect(this.Pt, Point_Add(this.Pt, Size_ToPoint(this.Sz)))
End Function

Public Function AARect_Equals(this As AARect, other As AARect) As Boolean
    With this: AARect_Equals = Point_Equals(.Pt, other.Pt) And Size_Equals(.Sz, other.Sz): End With
End Function
' ^ ############################## ^ '   AARect    ' ^ ############################## ^ '

' v ############################## v '   PPRect    ' v ############################## v '
Public Function New_PPRect(Point1 As Point, Point2 As Point) As PPRect
    With New_PPRect: .P1 = Point1: .P2 = Point2: End With
End Function

Public Function PPRect_Mul(this As PPRect, ByVal Scalar As Double) As PPRect
    PPRect_Mul = New_PPRect(Point_Mul(this.P1, Scalar), Point_Mul(this.P2, Scalar))
End Function

Public Function PPRect_ToStr(this As PPRect) As String
    PPRect_ToStr = "PPRect(P1: " & Point_ToStr(this.P1) & "; P2: " & Point_ToStr(this.P2) & ")"
End Function

Public Function PPRect_Equals(this As PPRect, other As PPRect) As Boolean
    With this: PPRect_Equals = Point_Equals(.P1, other.P1) And Point_Equals(.P2, other.P2): End With
End Function

'Test:
'?Point_ToStr(New_Point(50, 100, 30))
'?Point_ToStr(Point_Add(New_Point(50, 75, 10), New_Point(25, 35, 15)))
'?Point_ToStr(Point_Dif(New_Point(50, 75, 10), New_Point(25, 35, 5)))
'?Point_ToStr(Point_Mul(New_Point(50, 75, 10), 1.5))
'?Point_ToStr(Point_Div(New_Point(50, 75, 10), 1.5))
'?Point_ToStr(Point_Neg(New_Point(50, 100, 30)))
'?Size_ToStr(New_Size(50, 62, 80))
'?Point_ToStr(Size_ToPoint(New_Size(50, 62, 80)))
'?AARect_ToStr(New_AARect(New_Point(50, 100), New_Size(50, 62)))
'?AARect_ToStr(AARect_Mul(New_AARect(New_Point(50, 100), New_Size(50, 62)), 1.25))
'?PPRect_ToStr(AARect_ToPPRect(New_AARect(New_Point(50, 100), New_Size(50, 62))))

