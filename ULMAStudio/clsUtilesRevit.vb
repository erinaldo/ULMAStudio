#Region "Imported Namespaces"
Imports System.Runtime.InteropServices
Imports System.Runtime.CompilerServices
Imports System.Reflection
Imports Autodesk.Revit.ApplicationServices
#End Region

Namespace clsUtilesRevit
    Class Util
        Public Const _eps As Double = 0.000000001

        Public Shared ReadOnly Property Eps As Double
            Get
                Return _eps
            End Get
        End Property

        Public Shared ReadOnly Property MinLineLength As Double
            Get
                Return _eps
            End Get
        End Property

        Public Shared ReadOnly Property TolPointOnPlane As Double
            Get
                Return _eps
            End Get
        End Property

        Public Shared Sub ExecuteCommandId(queComando As Object)
            'ID_REVIT_FILE_OPEN (OpenRevitFile)
            'ID_WINDOW_CLOSE_HIDDEN
            Dim cId As RevitCommandId = Nothing
            If TypeOf queComando Is PostableCommand Then
                cId = RevitCommandId.LookupPostableCommandId(CType(queComando, PostableCommand))
            ElseIf TypeOf queComando Is String Then
                cId = RevitCommandId.LookupCommandId(CType(queComando, String))
            End If
            If cId IsNot Nothing Then
                Try
                    evRevit.evAppUI.PostCommand(cId)
                Catch ex As Exception
                    MsgBox("Error --> Command " & queComando.ToString & "Is Not exist")
                End Try
            End If
        End Sub

        Public Shared Sub ExecuteCommandIdRevit(queComando As PostableCommand)
            'ID_REVIT_FILE_OPEN (OpenRevitFile)
            Dim cId As RevitCommandId = RevitCommandId.LookupPostableCommandId(queComando)
            If cId IsNot Nothing Then
                Try
                    evRevit.evAppUI.PostCommand(cId)
                Catch ex As Exception
                    MsgBox("Error --> Command " & queComando.ToString & "Is Not exist")
                End Try
            End If
        End Sub

        Public Shared Sub ExecuteCommandIdAddIn(queComando As String)
            'ID_REVIT_FILE_OPEN (OpenRevitFile)
            Dim cId As RevitCommandId = RevitCommandId.LookupCommandId(queComando)
            If cId IsNot Nothing Then
                Try
                    evRevit.evAppUI.PostCommand(cId)
                Catch ex As Exception
                    MsgBox("Error --> Command " & queComando.ToString & "Is Not exist")
                End Try
            End If
        End Sub
        '
        Public Shared Function IsZero(ByVal a As Double, ByVal Optional tolerance As Double = _eps) As Boolean
            Return tolerance > Math.Abs(a)
        End Function

        Public Shared Function IsEqual(ByVal a As Double, ByVal b As Double, ByVal Optional tolerance As Double = _eps) As Boolean
            Return IsZero(b - a, tolerance)
        End Function

        Public Shared Function Compare(ByVal a As Double, ByVal b As Double, ByVal Optional tolerance As Double = _eps) As Integer
            Return If(IsEqual(a, b, tolerance), 0, (If(a < b, -1, 1)))
        End Function

        Public Shared Function Compare(ByVal p As XYZ, ByVal q As XYZ) As Integer
            Dim d As Integer = Compare(p.X, q.X)

            If 0 = d Then
                d = Compare(p.Y, q.Y)

                If 0 = d Then
                    d = Compare(p.Z, q.Z)
                End If
            End If

            Return d
        End Function

        Public Shared Function Compare(ByVal a As Line, ByVal b As Line) As Integer
            Dim pa As XYZ = a.GetEndPoint(0)
            Dim qa As XYZ = a.GetEndPoint(1)
            Dim pb As XYZ = b.GetEndPoint(0)
            Dim qb As XYZ = b.GetEndPoint(1)
            Dim va As XYZ = qa - pa
            Dim vb As XYZ = qb - pb
            Dim ang_a As Double = Math.Atan2(va.Y, va.X)
            Dim ang_b As Double = Math.Atan2(vb.Y, vb.X)
            Dim d As Integer = Compare(ang_a, ang_b)

            If 0 = d Then
                Dim da As Double = (qa.X * pa.Y - qa.Y * pa.Y) / va.GetLength()
                Dim db As Double = (qb.X * pb.Y - qb.Y * pb.Y) / vb.GetLength()
                d = Compare(da, db)

                If 0 = d Then
                    d = Compare(pa.GetLength(), pb.GetLength())

                    If 0 = d Then
                        d = Compare(qa.GetLength(), qb.GetLength())
                    End If
                End If
            End If

            Return d
        End Function

        'Public Shared Function Compare(ByVal a As Plane, ByVal b As Plane) As Integer
        '    Dim d As Integer = Compare(a.Normal, b.Normal)

        '    If 0 = d Then
        '        d = Compare(a.SignedDistanceTo(XYZ.Zero), b.SignedDistanceTo(XYZ.Zero))

        '        If 0 = d Then
        '            d = Compare(a.XVec.AngleOnPlaneTo(b.XVec, b.Normal), 0)
        '        End If
        '    End If

        '    Return d
        'End Function

        Public Shared Function IsEqual(ByVal p As XYZ, ByVal q As XYZ) As Boolean
            Return 0 = Compare(p, q)
        End Function

        Public Function BoundingBoxXyzContains(ByVal bb As BoundingBoxXYZ, ByVal p As XYZ) As Boolean
            Return 0 < Compare(bb.Min, p) AndAlso 0 < Compare(p, bb.Max)
        End Function

        Private Function IsPerpendicular(ByVal v As XYZ, ByVal w As XYZ) As Boolean
            Dim a As Double = v.GetLength()
            Dim b As Double = v.GetLength()
            Dim c As Double = Math.Abs(v.DotProduct(w))
            Return _eps < a AndAlso _eps < b AndAlso _eps > c
        End Function

        Public Shared Function IsParallel(ByVal p As XYZ, ByVal q As XYZ) As Boolean
            Return p.CrossProduct(q).IsZeroLength()
        End Function

        Public Shared Function IsCollinear(ByVal a As Line, ByVal b As Line) As Boolean
            Dim v As XYZ = a.Direction
            Dim w As XYZ = b.Origin - a.Origin
            Return IsParallel(v, b.Direction) AndAlso IsParallel(v, w)
        End Function

        Public Shared Function IsHorizontal(ByVal v As XYZ) As Boolean
            Return IsZero(v.Z)
        End Function

        Public Shared Function IsVertical(ByVal v As XYZ) As Boolean
            Return IsZero(v.X) AndAlso IsZero(v.Y)
        End Function

        Public Shared Function IsVertical(ByVal v As XYZ, ByVal tolerance As Double) As Boolean
            Return IsZero(v.X, tolerance) AndAlso IsZero(v.Y, tolerance)
        End Function

        Public Shared Function IsHorizontal(ByVal e As Edge) As Boolean
            Dim p As XYZ = e.Evaluate(0)
            Dim q As XYZ = e.Evaluate(1)
            Return IsHorizontal(q - p)
        End Function

        Public Shared Function IsHorizontal(ByVal f As PlanarFace) As Boolean
            Return IsVertical(f.FaceNormal)
        End Function

        Public Shared Function IsVertical(ByVal f As PlanarFace) As Boolean
            Return IsHorizontal(f.FaceNormal)
        End Function

        Public Shared Function IsVertical(ByVal f As CylindricalFace) As Boolean
            Return IsVertical(f.Axis)
        End Function

        Const _minimumSlope As Double = 0.3

        Public Shared Function PointsUpwards(ByVal v As XYZ) As Boolean
            Dim horizontalLength As Double = v.X * v.X + v.Y * v.Y
            Dim verticalLength As Double = v.Z * v.Z
            Return 0 < v.Z AndAlso _minimumSlope < verticalLength / horizontalLength
        End Function

        Public Shared Function Max(ByVal a As Double()) As Double
            Debug.Assert(1 = a.Rank, "expected one-dimensional array")
            Debug.Assert(0 = a.GetLowerBound(0), "expected zero-based array")
            Debug.Assert(0 < a.GetUpperBound(0), "expected non-empty array")
            Dim maxi As Double = a(0)

            For i As Integer = 1 To a.GetUpperBound(0)

                If maxi < a(i) Then
                    maxi = a(i)
                End If
            Next

            Return maxi
        End Function

        Public Shared Sub GetArbitraryAxes(ByVal normal As XYZ, <Out> ByRef ax As XYZ, <Out> ByRef ay As XYZ)
            Dim limit As Double = 1.0 / 64
            Dim pick_cardinal_axis As XYZ = If((IsZero(normal.X, limit) AndAlso IsZero(normal.Y, limit)), XYZ.BasisY, XYZ.BasisZ)
            ax = pick_cardinal_axis.CrossProduct(normal).Normalize()
            ay = normal.CrossProduct(ax).Normalize()
        End Sub

        Public Shared Function Midpoint(ByVal p As XYZ, ByVal q As XYZ) As XYZ
            Return 0.5 * (p + q)
        End Function

        Public Shared Function Midpoint(ByVal line As Line) As XYZ
            Return Midpoint(line.GetEndPoint(0), line.GetEndPoint(1))
        End Function

        Public Shared Function Normal(ByVal line As Line) As XYZ
            Dim p As XYZ = line.GetEndPoint(0)
            Dim q As XYZ = line.GetEndPoint(1)
            Dim v As XYZ = q - p
            Return v.CrossProduct(XYZ.BasisZ).Normalize()
        End Function

        Public Shared Function GetBottomCorners(ByVal b As BoundingBoxXYZ, ByVal z As Double) As XYZ()
            Return New XYZ() {New XYZ(b.Min.X, b.Min.Y, z), New XYZ(b.Max.X, b.Min.Y, z), New XYZ(b.Max.X, b.Max.Y, z), New XYZ(b.Min.X, b.Max.Y, z)}
        End Function

        Public Shared Function GetBottomCorners(ByVal b As BoundingBoxXYZ) As XYZ()
            Return GetBottomCorners(b, b.Min.Z)
        End Function

        Public Shared Function Intersection(ByVal c1 As Curve, ByVal c2 As Curve) As XYZ
            Dim p1 As XYZ = c1.GetEndPoint(0)
            Dim q1 As XYZ = c1.GetEndPoint(1)
            Dim p2 As XYZ = c2.GetEndPoint(0)
            Dim q2 As XYZ = c2.GetEndPoint(1)
            Dim v1 As XYZ = q1 - p1
            Dim v2 As XYZ = q2 - p2
            Dim w As XYZ = p2 - p1
            Dim p5 As XYZ = Nothing
            Dim c As Double = (v2.X * w.Y - v2.Y * w.X) / (v2.X * v1.Y - v2.Y * v1.X)

            If Not Double.IsInfinity(c) Then
                Dim x As Double = p1.X + c * v1.X
                Dim y As Double = p1.Y + c * v1.Y
                p5 = New XYZ(x, y, 0)
            End If

            Return p5
        End Function

        Public Shared Function CalculateMatrixForGlobalToLocalCoordinateSystem(ByVal face As Face) As Double(,)
            Dim originDisplacementVectorUV As XYZ = face.Evaluate(UV.Zero)
            Dim unitVectorUWithDisplacement As XYZ = face.Evaluate(UV.BasisU)
            Dim unitVectorVWithDisplacement As XYZ = face.Evaluate(UV.BasisV)
            Dim unitVectorU As XYZ = unitVectorUWithDisplacement - originDisplacementVectorUV
            Dim unitVectorV As XYZ = unitVectorVWithDisplacement - originDisplacementVectorUV
            Dim a11i = unitVectorU.X
            Dim a12i = unitVectorU.Y
            Dim a21i = unitVectorV.X
            Dim a22i = unitVectorV.Y
            Return New Double(1, 1) {
        {a11i, a12i},
        {a21i, a22i}}
        End Function

        Public Shared Function CreateArc2dFromRadiusStartAndEndPoint(ByVal ps As XYZ, ByVal pe As XYZ, ByVal radius As Double, ByVal Optional largeSagitta As Boolean = False, ByVal Optional clockwise As Boolean = False) As Arc
            Dim midPointChord As XYZ = 0.5 * (ps + pe)
            Dim v As XYZ = pe - ps
            Dim d As Double = 0.5 * v.GetLength()
            Dim s As Double = If(largeSagitta, radius + Math.Sqrt(radius * radius - d * d), radius - Math.Sqrt(radius * radius - d * d))
            Dim midPointOffset As XYZ = Transform.CreateRotation(XYZ.BasisZ, 0.5 * Math.PI).OfVector(v.Normalize().Multiply(s))
            Dim midPointArc As XYZ = If(clockwise, midPointChord + midPointOffset, midPointChord - midPointOffset)
            Return Arc.Create(ps, pe, midPointArc)
        End Function

        Public Shared Function CreateSphereAt(ByVal centre As XYZ, ByVal radius As Double) As Solid
            Dim frame As Frame = New Frame(centre, XYZ.BasisX, XYZ.BasisY, XYZ.BasisZ)
            Dim arc As Arc = Arc.Create(centre - radius * XYZ.BasisZ, centre + radius * XYZ.BasisZ, centre + radius * XYZ.BasisX)
            Dim line As Line = Line.CreateBound(arc.GetEndPoint(1), arc.GetEndPoint(0))
            Dim halfCircle As CurveLoop = New CurveLoop()
            halfCircle.Append(arc)
            halfCircle.Append(line)
            Dim loops As List(Of CurveLoop) = New List(Of CurveLoop)(1)
            loops.Add(halfCircle)
            Return GeometryCreationUtilities.CreateRevolvedGeometry(frame, loops, 0, 2 * Math.PI)
        End Function

        Public Shared Function CreateCone(ByVal center As XYZ, ByVal axis_vector As XYZ, ByVal radius As Double, ByVal height As Double) As Solid
            Dim az As XYZ = axis_vector.Normalize()
            Dim ax As XYZ = Nothing
            Dim ay As XYZ = Nothing
            GetArbitraryAxes(az, ax, ay)
            Dim px As XYZ = center + radius * ax
            Dim pz As XYZ = center + height * az
            Dim profile As List(Of Curve) = New List(Of Curve)()
            profile.Add(Line.CreateBound(center, px))
            profile.Add(Line.CreateBound(px, pz))
            profile.Add(Line.CreateBound(pz, center))
            Dim curveLoop As CurveLoop = CurveLoop.Create(profile)
            Dim frame As Frame = New Frame(center, ax, ay, az)
            Dim cone As Solid = GeometryCreationUtilities.CreateRevolvedGeometry(frame, New CurveLoop() {curveLoop}, 0, 2 * Math.PI)
            Return cone
        End Function

        Private Shared Function CreateCube(ByVal d As Double) As Solid
            Return CreateRectangularPrism(XYZ.Zero, d, d, d)
        End Function

        Private Shared Function CreateRectangularPrism(ByVal center As XYZ, ByVal d1 As Double, ByVal d2 As Double, ByVal d3 As Double) As Solid
            Dim profile As List(Of Curve) = New List(Of Curve)()
            Dim profile00 As XYZ = New XYZ(-d1 / 2, -d2 / 2, -d3 / 2)
            Dim profile01 As XYZ = New XYZ(-d1 / 2, d2 / 2, -d3 / 2)
            Dim profile11 As XYZ = New XYZ(d1 / 2, d2 / 2, -d3 / 2)
            Dim profile10 As XYZ = New XYZ(d1 / 2, -d2 / 2, -d3 / 2)
            profile.Add(Line.CreateBound(profile00, profile01))
            profile.Add(Line.CreateBound(profile01, profile11))
            profile.Add(Line.CreateBound(profile11, profile10))
            profile.Add(Line.CreateBound(profile10, profile00))
            Dim curveLoop As CurveLoop = CurveLoop.Create(profile)
            Dim options As SolidOptions = New SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId)
            Return GeometryCreationUtilities.CreateExtrusionGeometry(New CurveLoop() {curveLoop}, XYZ.BasisZ, d3, options)
        End Function

        Public Shared Function CreateSolidFromBoundingBox(ByVal inputSolid As Solid) As Solid
            Dim bbox As BoundingBoxXYZ = inputSolid.GetBoundingBox()
            Dim pt0 As XYZ = New XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z)
            Dim pt1 As XYZ = New XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z)
            Dim pt2 As XYZ = New XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z)
            Dim pt3 As XYZ = New XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z)
            Dim edge0 As Line = Line.CreateBound(pt0, pt1)
            Dim edge1 As Line = Line.CreateBound(pt1, pt2)
            Dim edge2 As Line = Line.CreateBound(pt2, pt3)
            Dim edge3 As Line = Line.CreateBound(pt3, pt0)
            Dim edges As List(Of Curve) = New List(Of Curve)()
            edges.Add(edge0)
            edges.Add(edge1)
            edges.Add(edge2)
            edges.Add(edge3)
            Dim height As Double = bbox.Max.Z - bbox.Min.Z
            Dim baseLoop As CurveLoop = CurveLoop.Create(edges)
            Dim loopList As List(Of CurveLoop) = New List(Of CurveLoop)()
            loopList.Add(baseLoop)
            Dim preTransformBox As Solid = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, height)
            Dim transformBox As Solid = SolidUtils.CreateTransformed(preTransformBox, bbox.Transform)
            Return transformBox
        End Function

        Public Shared Function ConvexHull(ByVal points As List(Of XYZ)) As List(Of XYZ)
            If points Is Nothing Then Throw New ArgumentNullException(NameOf(points))
            Dim startPoint As XYZ = points.MinBy(Function(p) p.X)
            Dim convexHullPoints = New List(Of XYZ)()
            Dim walkingPoint As XYZ = startPoint
            Dim refVector As XYZ = XYZ.BasisY.Negate()

            Do
                convexHullPoints.Add(walkingPoint)
                Dim wp As XYZ = walkingPoint
                Dim rv As XYZ = refVector
                walkingPoint = points.MinBy(Function(p)
                                                Dim angle As Double = (p - wp).AngleOnPlaneTo(rv, XYZ.BasisZ)
                                                If angle < 0.0000000001 Then angle = 2 * Math.PI
                                                Return angle
                                            End Function)
                refVector = wp - walkingPoint
            Loop While walkingPoint.Equals(startPoint) = False

            convexHullPoints.Reverse()
            Return convexHullPoints
        End Function

        Enum BaseUnit
            BU_Length = 0
            BU_Angle
            BU_Mass
            BU_Time
            BU_Electric_Current
            BU_Temperature
            BU_Luminous_Intensity
            BU_Solid_Angle
            NumBaseUnits
        End Enum

        Const _inchToMm As Double = 25.4
        Const _footToMm As Double = 12 * _inchToMm
        Const _footToMeter As Double = _footToMm * 0.001
        Const _cubicFootToCubicMeter As Double = _footToMeter * _footToMeter * _footToMeter

        Public Shared Function FootToMm(ByVal length As Double) As Double
            Return length * _footToMm
        End Function

        Public Shared Function FootToMetre(ByVal length As Double) As Double
            Return length * _footToMeter
        End Function

        Public Shared Function MmToFoot(ByVal length As Double) As Double
            Return length / _footToMm
        End Function

        Public Shared Function MmToFoot(ByVal v As XYZ) As XYZ
            Return v.Divide(_footToMm)
        End Function

        Public Shared Function CubicFootToCubicMeter(ByVal volume As Double) As Double
            Return volume * _cubicFootToCubicMeter
        End Function

        Public Shared DisplayUnitTypeAbbreviation As String() = New String() {"m", "cm", "mm", "ft", "N/A", "N/A", "in", "ac", "ha", "N/A", "y^3", "ft^2", "m^2", "ft^3", "m^3", "deg", "N/A", "N/A", "N/A", "%", "in^2", "cm^2", "mm^2", "in^3", "cm^3", "mm^3", "l"}

        Public Shared Function PluralSuffix(ByVal n As Integer) As String
            Return If(1 = n, "", "s")
        End Function

        Public Shared Function PluralSuffixY(ByVal n As Integer) As String
            Return If(1 = n, "y", "ies")
        End Function

        Public Shared Function DotOrColon(ByVal n As Integer) As String
            Return If(0 < n, ": ", ".")
        End Function

        Public Shared Function RealString(ByVal a As Double) As String
            Return a.ToString("0.##")
        End Function

        Public Shared Function HashString(ByVal a As Double) As String
            Return a.ToString("0.#########")
        End Function

        Public Shared Function AngleString(ByVal angle As Double) As String
            Return RealString(angle * 180 / Math.PI) & " degrees"
        End Function

        Public Shared Function MmString(ByVal length As Double) As String
            Return Math.Round(FootToMm(length)).ToString() & " mm"
        End Function

        Public Shared Function PointString(ByVal p As UV) As String
            Return String.Format("({0},{1})", RealString(p.U), RealString(p.V))
        End Function

        Public Shared Function PointString(ByVal p As XYZ) As String
            Return String.Format("({0},{1},{2})", RealString(p.X), RealString(p.Y), RealString(p.Z))
        End Function

        Public Shared Function HashString(ByVal p As XYZ) As String
            Return String.Format("({0},{1},{2})", HashString(p.X), HashString(p.Y), HashString(p.Z))
        End Function

        Public Shared Function BoundingBoxString(ByVal bb As BoundingBoxUV) As String
            Return String.Format("({0},{1})", PointString(bb.Min), PointString(bb.Max))
        End Function

        Public Shared Function BoundingBoxString(ByVal bb As BoundingBoxXYZ) As String
            Return String.Format("({0},{1})", PointString(bb.Min), PointString(bb.Max))
        End Function

        Public Shared Function PlaneString(ByVal p As Plane) As String
            Return String.Format("plane origin {0}, plane normal {1}", PointString(p.Origin), PointString(p.Normal))
        End Function

        Public Shared Function TransformString(ByVal t As Transform) As String
            Return String.Format("({0},{1},{2},{3})", PointString(t.Origin), PointString(t.BasisX), PointString(t.BasisY), PointString(t.BasisZ))
        End Function

        Public Shared Function DoubleArrayString(ByVal a As IList(Of Double)) As String
            Return String.Join(", ", a.Select(Function(x) RealString(x)))
        End Function

        Public Shared Function PointArrayString(ByVal pts As IList(Of XYZ)) As String
            Return String.Join(", ", pts.Select(Function(p) PointString(p)))
        End Function

        Public Shared Function CurveString(ByVal c As Curve) As String
            Dim s As String = c.[GetType]().Name.ToLower()
            Dim p As XYZ = c.GetEndPoint(0)
            Dim q As XYZ = c.GetEndPoint(1)
            s += String.Format(" {0} --> {1}", PointString(p), PointString(q))
            Dim arc As Arc = TryCast(c, Arc)

            If arc IsNot Nothing Then
                s += String.Format(" center {0} radius {1}", PointString(arc.Center), arc.Radius)
            End If

            Return s
        End Function

        Public Shared Function CurveTessellateString(ByVal curve As Curve) As String
            Return "curve tessellation " & PointArrayString(curve.Tessellate())
        End Function

        Public Shared Function UnitSymbolTypeString(ByVal u As UnitSymbolType) As String
            Dim s As String = u.ToString()
            Debug.Assert(s.StartsWith("UST_"), "expected UnitSymbolType enumeration value " & "to begin with 'UST_'")
            s = s.Substring(4).Replace("_SUP_", "^").ToLower()
            Return s
        End Function

        Const _caption As String = "The Building Coder"

        Public Shared Sub InfoMsg(ByVal msg As String)
            Debug.WriteLine(msg)
            MsgBox(msg, MsgBoxStyle.Information)
        End Sub

        Public Shared Sub InfoMsg2(ByVal instruction As String, ByVal content As String)
            Debug.WriteLine(instruction & vbCrLf & content)
            Dim d As TaskDialog = New TaskDialog(_caption)
            d.MainInstruction = instruction
            d.MainContent = content
            d.Show()
        End Sub

        Public Shared Sub ErrorMsg(ByVal msg As String)
            Debug.WriteLine(msg)
            MsgBox(msg, MsgBoxStyle.Critical)
        End Sub

        Public Shared Function ElementDescription(ByVal e As Element) As String
            If e Is Nothing Then
                Return "<null>"
            End If

            Dim fi As FamilyInstance = TryCast(e, FamilyInstance)
            Dim typeName As String = e.[GetType]().Name
            Dim categoryName As String = If((e.Category Is Nothing), String.Empty, e.Category.Name & " ")
            Dim familyName As String = If((fi Is Nothing), String.Empty, fi.Symbol.Family.Name & " ")
            Dim symbolName As String = If((fi Is Nothing OrElse e.Name.Equals(fi.Symbol.Name)), String.Empty, fi.Symbol.Name & " ")
            Return String.Format("{0} {1}{2}{3}<{4} {5}>", typeName, categoryName, familyName, symbolName, e.Id.IntegerValue, e.Name)
        End Function

        Public Shared Function GetElementLocation(<Out> ByRef p As XYZ, ByVal e As Element) As Boolean
            p = XYZ.Zero
            Dim rc As Boolean = False
            Dim loc As Location = e.Location

            If loc IsNot Nothing Then
                Dim lp As LocationPoint = TryCast(loc, LocationPoint)

                If lp IsNot Nothing Then
                    p = lp.Point
                    rc = True
                Else
                    Dim lc As LocationCurve = TryCast(loc, LocationCurve)
                    Debug.Assert(lc IsNot Nothing, "expected location to be either point or curve")
                    p = lc.Curve.GetEndPoint(0)
                    rc = True
                End If
            End If

            Return rc
        End Function

        Public Shared Function GetFamilyInstanceLocation(ByVal fi As FamilyInstance) As XYZ
            Return (CType(fi?.Location, LocationPoint))?.Point
        End Function

        Public Shared Function SelectSingleElement(ByVal uidoc As UIDocument, ByVal description As String) As Element
            If ViewType.Internal = uidoc.ActiveView.ViewType Then
                TaskDialog.Show("Error", "Cannot pick element in this view: " & uidoc.ActiveView.Name)
                Return Nothing
            End If

            Try
                Dim r As Reference = uidoc.Selection.PickObject(ObjectType.Element, "Please select " & description)
                Return uidoc.Document.GetElement(r)
            Catch __unusedOperationCanceledException1__ As Autodesk.Revit.Exceptions.OperationCanceledException
                Return Nothing
            End Try
        End Function

        Public Shared Function GetSingleSelectedElement(ByVal uidoc As UIDocument) As Element
            Dim ids As ICollection(Of ElementId) = uidoc.Selection.GetElementIds()
            Dim e As Element = Nothing

            If 1 = ids.Count Then

                For Each id As ElementId In ids
                    e = uidoc.Document.GetElement(id)
                Next
            End If

            Return e
        End Function

        Private Shared Function HasRequestedType(ByVal e As Element, ByVal t As Type, ByVal acceptDerivedClass As Boolean) As Boolean
            Dim rc As Boolean = e IsNot Nothing

            If rc Then
                Dim t2 As Type = e.[GetType]()
                rc = t2.Equals(t)

                If Not rc AndAlso acceptDerivedClass Then
                    rc = t2.IsSubclassOf(t)
                End If
            End If

            Return rc
        End Function

        Public Shared Function SelectSingleElementOfType(ByVal uidoc As UIDocument, ByVal t As Type, ByVal description As String, ByVal acceptDerivedClass As Boolean) As Element
            Dim e As Element = GetSingleSelectedElement(uidoc)

            If Not HasRequestedType(e, t, acceptDerivedClass) Then
                e = Util.SelectSingleElement(uidoc, description)
            End If

            Return If(HasRequestedType(e, t, acceptDerivedClass), e, Nothing)
        End Function

        Public Shared Function GetSelectedElementsOrAll(ByVal a As List(Of Element), ByVal uidoc As UIDocument, ByVal t As Type) As Boolean
            Dim doc As Document = uidoc.Document
            Dim ids As ICollection(Of ElementId) = uidoc.Selection.GetElementIds()

            If 0 < ids.Count Then
                a.AddRange(ids.Select(Function(id) doc.GetElement(id)).Where(Function(e) t.IsInstanceOfType(e)))
            Else
                a.AddRange(New FilteredElementCollector(doc).OfClass(t))
            End If

            Return 0 < a.Count
        End Function

        Public Shared Function GetElementsOfType(ByVal doc As Document, ByVal type As Type, ByVal bic As BuiltInCategory) As FilteredElementCollector
            Dim collector As FilteredElementCollector = New FilteredElementCollector(doc)
            collector.OfCategory(bic)
            collector.OfClass(type)
            Return collector
        End Function

        Public Shared Function GetFirstElementOfTypeNamed(ByVal doc As Document, ByVal type As Type, ByVal name As String) As Element
            Dim collector As FilteredElementCollector = New FilteredElementCollector(doc).OfClass(type)
            'Dim nameEquals As Func(Of Element, Boolean) = Function(e) e.Name.Equals(name)
            ''Return If(collector.Any(nameEquals(Of Element, true), collector.First(Of Element)(nameEquals), Nothing)
            'Return If(collector.Any(nameEquals(Element, True), collector.First(x), Nothing)
            Dim resultado As Element = Nothing
            '
            Dim query1 As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.Element)
            query1 = From element In collector
                     Where CType(element, Element).Name.Equals(name)
                     Select CType(element, Element)
            '
            If query1 IsNot Nothing AndAlso query1.Count > 0 Then
                resultado = query1.First
            End If
            Return resultado
        End Function

        Public Shared Function GetFirstNonTemplate3dView(ByVal doc As Document) As Element
            Dim collector As FilteredElementCollector = New FilteredElementCollector(doc)
            collector.OfClass(GetType(View3D))
            Return collector.Cast(Of View3D).First(Function(v3) Not v3.IsTemplate)
        End Function

        Public Shared Function FindFamilySymbol(ByVal doc As Document, ByVal familyName As String, ByVal symbolName As String) As FamilySymbol
            Dim collector As FilteredElementCollector = New FilteredElementCollector(doc).OfClass(GetType(Family))

            For Each f As Family In collector

                If f.Name.Equals(familyName) Then
                    Dim ids As ISet(Of ElementId) = f.GetFamilySymbolIds()

                    For Each id As ElementId In ids
                        Dim symbol As FamilySymbol = TryCast(doc.GetElement(id), FamilySymbol)

                        If symbol.Name = symbolName Then
                            Return symbol
                        End If
                    Next
                End If
            Next

            Return Nothing
        End Function

        Private Shared Function GetConnectorManager(ByVal e As Element) As ConnectorManager
            Dim mc As MEPCurve = TryCast(e, MEPCurve)
            Dim fi As FamilyInstance = TryCast(e, FamilyInstance)

            If mc Is Nothing AndAlso fi Is Nothing Then
                Throw New ArgumentException("Element is neither an MEP curve nor a fitting.")
            End If

            Return If(mc Is Nothing, fi.MEPModel.ConnectorManager, mc.ConnectorManager)
        End Function

        Private Shared Function GetConnectorAt(ByVal e As Element, ByVal location As XYZ, <Out> ByRef otherConnector As Connector) As Connector
            otherConnector = Nothing
            Dim targetConnector As Connector = Nothing
            Dim cm As ConnectorManager = GetConnectorManager(e)
            Dim hasTwoConnectors As Boolean = 2 = cm.Connectors.Size

            For Each c As Connector In cm.Connectors

                If c.Origin.IsAlmostEqualTo(location) Then
                    targetConnector = c

                    If Not hasTwoConnectors Then
                        Exit For
                    End If
                ElseIf hasTwoConnectors Then
                    otherConnector = c
                End If
            Next

            Return targetConnector
        End Function

        Private Shared Function GetConnectorClosestTo(ByVal connectors As ConnectorSet, ByVal p As XYZ) As Connector
            Dim targetConnector As Connector = Nothing
            Dim minDist As Double = Double.MaxValue

            For Each c As Connector In connectors
                Dim d As Double = c.Origin.DistanceTo(p)

                If d < minDist Then
                    targetConnector = c
                    minDist = d
                End If
            Next

            Return targetConnector
        End Function

        Public Shared Function GetConnectorClosestTo(ByVal e As Element, ByVal p As XYZ) As Connector
            Dim cm As ConnectorManager = GetConnectorManager(e)
            Return If(cm Is Nothing, Nothing, GetConnectorClosestTo(cm.Connectors, p))
        End Function

        Public Shared Sub Connect(ByVal p As XYZ, ByVal a As Element, ByVal b As Element)
            Dim cm As ConnectorManager = GetConnectorManager(a)

            If cm Is Nothing Then
                Throw New ArgumentException("Element a has no connectors.")
            End If

            Dim ca As Connector = GetConnectorClosestTo(cm.Connectors, p)
            cm = GetConnectorManager(b)

            If cm Is Nothing Then
                Throw New ArgumentException("Element b has no connectors.")
            End If

            Dim cb As Connector = GetConnectorClosestTo(cm.Connectors, p)
            ca.ConnectTo(cb)
        End Sub

        Public Class SpellingErrorCorrector
            Shared _in_revit_2015_or_earlier As Boolean
            Shared _external_definition_creation_options_type As Type

            Public Sub New(ByVal app As Application)
                _in_revit_2015_or_earlier = 0 <= app.VersionNumber.CompareTo("2015")
                Dim s As String = If(_in_revit_2015_or_earlier, "ExternalDefinitonCreationOptions", "ExternalDefinitionCreationOptions")
                _external_definition_creation_options_type = System.Reflection.Assembly.GetExecutingAssembly().[GetType](s)
            End Sub

            Private Function NewExternalDefinitionCreationOptions(ByVal name As String, ByVal parameterType As ParameterType) As Object
                Dim args As Object() = New Object() {name, parameterType}
                Return _external_definition_creation_options_type.GetConstructor(New Type() {_external_definition_creation_options_type}).Invoke(args)
            End Function

            Public Function NewDefinition(ByVal definitions As Definitions, ByVal name As String, ByVal parameterType As ParameterType) As Definition
                Dim opt As Object = NewExternalDefinitionCreationOptions(name, parameterType)
                Return TryCast(GetType(Definitions).InvokeMember("Create", BindingFlags.InvokeMethod, Nothing, definitions, New Object() {opt}), Definition)
            End Function
        End Class
    End Class

    Module IEnumerableExtensions
        <Extension()>
        Function MinBy(Of tsource, tkey)(ByVal source As IEnumerable(Of tsource), ByVal selector As Func(Of tsource, tkey)) As tsource
            Return source.MinBy(selector, Comparer(Of tkey).[Default])
        End Function

        <Extension()>
        Function MinBy(Of tsource, tkey)(ByVal source As IEnumerable(Of tsource), ByVal selector As Func(Of tsource, tkey), ByVal comparer As IComparer(Of tkey)) As tsource
            If source Is Nothing Then Throw New ArgumentNullException(NameOf(source))
            If selector Is Nothing Then Throw New ArgumentNullException(NameOf(selector))
            If comparer Is Nothing Then Throw New ArgumentNullException(NameOf(comparer))

            Using sourceIterator As IEnumerator(Of tsource) = source.GetEnumerator()
                If Not sourceIterator.MoveNext() Then Throw New InvalidOperationException("Sequence was empty")
                Dim min As tsource = sourceIterator.Current
                Dim minKey As tkey = selector(min)

                While sourceIterator.MoveNext()
                    Dim candidate As tsource = sourceIterator.Current
                    Dim candidateProjected As tkey = selector(candidate)

                    If comparer.Compare(candidateProjected, minKey) < 0 Then
                        min = candidate
                        minKey = candidateProjected
                    End If
                End While

                Return min
            End Using
        End Function

        <Extension()>
        Function ToHashSet(Of TSource, TElement)(ByVal source As IEnumerable(Of TSource), ByVal elementSelector As Func(Of TSource, TElement), ByVal comparer As IEqualityComparer(Of TElement)) As HashSet(Of TElement)
            If source Is Nothing Then Throw New ArgumentNullException("source")
            If elementSelector Is Nothing Then Throw New ArgumentNullException("elementSelector")
            Return New HashSet(Of TElement)(source.[Select](elementSelector), comparer)
        End Function

        <Extension()>
        Function ToHashSet(Of TSource)(ByVal source As IEnumerable(Of TSource)) As HashSet(Of TSource)
            Return source.ToHashSet(Function(item) item, Nothing)
        End Function

        <Extension()>
        Function ToHashSet(Of TSource)(ByVal source As IEnumerable(Of TSource), ByVal comparer As IEqualityComparer(Of TSource)) As HashSet(Of TSource)
            Return source.ToHashSet(Function(item) item, comparer)
        End Function

        <Extension()>
        Function ToHashSet(Of TSource, TElement)(ByVal source As IEnumerable(Of TSource), ByVal elementSelector As Func(Of TSource, TElement)) As HashSet(Of TElement)
            Return source.ToHashSet(elementSelector, Nothing)
        End Function
    End Module

    Module JtElementExtensionMethods
        <Extension()>
        Function IsPhysicalElement(ByVal e As Element) As Boolean
            If e.Category Is Nothing Then Return False
            If e.ViewSpecific Then Return False

            If (CType(e.Category.Id.IntegerValue, BuiltInCategory)) = BuiltInCategory.OST_HVAC_Zones Then
                Return False
            End If

            Return e.Category.CategoryType = CategoryType.Model AndAlso e.Category.CanAddSubcategory
        End Function

        <Extension()>
        Function GetCurve(ByVal e As Element) As Curve
            Debug.Assert(e.Location IsNot Nothing, "expected an element with a valid Location")
            Dim lc As LocationCurve = TryCast(e.Location, LocationCurve)
            Debug.Assert(lc IsNot Nothing, "expected an element with a valid LocationCurve")
            Return lc.Curve
        End Function
    End Module

    Module JtElementIdExtensionMethods
        <Extension()>
        Function IsInvalid(ByVal id As ElementId) As Boolean
            Return ElementId.InvalidElementId = id
        End Function

        <Extension()>
        Function IsValid(ByVal id As ElementId) As Boolean
            Return Not IsInvalid(id)
        End Function
    End Module

    Module JtLineExtensionMethods
        <Extension()>
        Function Contains(ByVal line As Line, ByVal p As XYZ, ByVal Optional tolerance As Double = Util._eps) As Boolean
            Dim a As XYZ = line.GetEndPoint(0)
            Dim b As XYZ = line.GetEndPoint(1)
            Dim f As Double = a.DistanceTo(b)
            Dim da As Double = a.DistanceTo(p)
            Dim db As Double = p.DistanceTo(b)
            Return ((da + db) - f) * f < tolerance
        End Function
    End Module

    Module JtBoundingBoxXyzExtensionMethods
        <Extension()>
        Sub ExpandToContain(ByVal bb As BoundingBoxXYZ, ByVal p As XYZ)
            bb.Min = New XYZ(Math.Min(bb.Min.X, p.X), Math.Min(bb.Min.Y, p.Y), Math.Min(bb.Min.Z, p.Z))
            bb.Max = New XYZ(Math.Max(bb.Max.X, p.X), Math.Max(bb.Max.Y, p.Y), Math.Max(bb.Max.Z, p.Z))
        End Sub

        <Extension()>
        Sub ExpandToContain(ByVal bb As BoundingBoxXYZ, ByVal other As BoundingBoxXYZ)
            bb.ExpandToContain(other.Min)
            bb.ExpandToContain(other.Max)
        End Sub
    End Module

    Module JtPlaneExtensionMethods
        <Extension()>
        Function SignedDistanceTo(ByVal plane As Plane, ByVal p As XYZ) As Double
            Debug.Assert(Util.IsEqual(plane.Normal.GetLength(), 1), "expected normalised plane normal")
            Dim v As XYZ = p - plane.Origin
            Return plane.Normal.DotProduct(v)
        End Function

        <Extension()>
        Function ProjectOnto(ByVal plane As Plane, ByVal p As XYZ) As XYZ
            Dim d As Double = plane.SignedDistanceTo(p)
            Dim q As XYZ = p - d * plane.Normal
            Debug.Assert(Util.IsZero(plane.SignedDistanceTo(q)), "expected point on plane to have zero distance to plane")
            Return q
        End Function

        <Extension()>
        Function ProjectInto(ByVal plane As Plane, ByVal p As XYZ) As UV
            Dim q As XYZ = plane.ProjectOnto(p)
            Dim o As XYZ = plane.Origin
            Dim d As XYZ = q - o
            Dim u As Double = d.DotProduct(plane.XVec)
            Dim v As Double = d.DotProduct(plane.YVec)
            Return New UV(u, v)
        End Function
    End Module

    Module JtEdgeArrayExtensionMethods
        <Extension()>
        Function GetPolygon(ByVal ea As EdgeArray) As List(Of XYZ)
            Dim n As Integer = ea.Size
            Dim polygon As List(Of XYZ) = New List(Of XYZ)(n)

            For Each e As Edge In ea
                Dim pts As IList(Of XYZ) = e.Tessellate()
                n = polygon.Count

                If 0 < n Then
                    Debug.Assert(pts(0).IsAlmostEqualTo(polygon(n - 1)), "expected last edge end point to " & "equal next edge start point")
                    polygon.RemoveAt(n - 1)
                End If

                polygon.AddRange(pts)
            Next

            n = polygon.Count
            Debug.Assert(polygon(0).IsAlmostEqualTo(polygon(n - 1)), "expected first edge start point to " & "equal last edge end point")
            polygon.RemoveAt(n - 1)
            Return polygon
        End Function
    End Module

    Module JtFamilyParameterExtensionMethods
        <Extension()>
        Function IsShared(ByVal familyParameter As FamilyParameter) As Boolean
            Dim mi As MethodInfo = familyParameter.[GetType]().GetMethod("getParameter", BindingFlags.Instance Or BindingFlags.NonPublic)

            If mi Is Nothing Then
                Throw New InvalidOperationException("Could not find getParameter method")
            End If

            Dim parameter = TryCast(mi.Invoke(familyParameter, New Object() {}), Parameter)
            Return parameter.IsShared
        End Function
    End Module

    Module JtFilteredElementCollectorExtensions
        <Extension()>
        Function OfClass(Of T As Element)(ByVal collector As FilteredElementCollector) As FilteredElementCollector
            Return collector.OfClass(GetType(T))
        End Function

        <Extension()>
        Function OfType(Of T As Element)(ByVal collector As FilteredElementCollector) As IEnumerable(Of T)
            Return Enumerable.OfType(Of T)(collector.OfClass(Of T)())
        End Function
    End Module

    Module JtBuiltInCategoryExtensionMethods
        <Extension()>
        Function Description(ByVal bic As BuiltInCategory) As String
            Dim s As String = bic.ToString().ToLower()
            s = s.Substring(4)
            Debug.Assert(s.EndsWith("s"), "expected plural suffix 's'")
            s = s.Substring(0, s.Length - 1)
            Return s
        End Function
    End Module

    Module CompatibilityMethods
        <Extension()>
        Function GetPoint2(ByVal curva As Curve, ByVal i As Integer) As XYZ
            Dim value As XYZ = Nothing
            Dim met As MethodInfo = curva.[GetType]().GetMethod("GetEndPoint", New Type() {GetType(Integer)})

            If met Is Nothing Then
                met = curva.[GetType]().GetMethod("get_EndPoint", New Type() {GetType(Integer)})
            End If

            value = TryCast(met.Invoke(curva, New Object() {i}), XYZ)
            Return value
        End Function

        <Extension()>
        Function Create2(ByVal definitions As Definitions, ByVal doc As Document, ByVal nome As String, ByVal tipo As ParameterType, ByVal visibilidade As Boolean) As Definition
            Dim value As Definition = Nothing
            Dim ls As List(Of Type) = doc.[GetType]().Assembly.GetTypes().Where(Function(a) a.IsClass AndAlso a.Name = "ExternalDefinitonCreationOptions").ToList()

            If ls.Count > 0 Then
                Dim t As Type = ls(0)
                Dim c As ConstructorInfo = t.GetConstructor(New Type() {GetType(String), GetType(ParameterType)})
                Dim ed As Object = c.Invoke(New Object() {nome, tipo})
                ed.[GetType]().GetProperty("Visible").SetValue(ed, visibilidade, Nothing)
                value = TryCast(definitions.[GetType]().GetMethod("Create", New Type() {t}).Invoke(definitions, New Object() {ed}), Definition)
            Else
                value = TryCast(definitions.[GetType]().GetMethod("Create", New Type() {GetType(String), GetType(ParameterType), GetType(Boolean)}).Invoke(definitions, New Object() {nome, tipo, visibilidade}), Definition)
            End If

            Return value
        End Function

        <Extension()>
        Function GetElement2(ByVal doc As Document, ByVal id As ElementId) As Element
            Dim value As Element = Nothing
            Dim met As MethodInfo = doc.[GetType]().GetMethod("get_Element", New Type() {GetType(ElementId)})
            If met Is Nothing Then met = doc.[GetType]().GetMethod("GetElement", New Type() {GetType(ElementId)})
            value = TryCast(met.Invoke(doc, New Object() {id}), Element)
            Return value
        End Function

        <Extension()>
        Function GetElement2(ByVal doc As Document, ByVal refe As Reference) As Element
            Dim value As Element = Nothing
            value = doc.GetElement(refe)
            Return value
        End Function

        <Extension()>
        Function CreateLine2(ByVal doc As Document, ByVal p1 As XYZ, ByVal p2 As XYZ, ByVal Optional bound As Boolean = True) As Line
            Dim value As Line = Nothing
            Dim parametros As Object() = New Object() {p1, p2}
            Dim tipos As Type() = parametros.[Select](Function(a) a.[GetType]()).ToArray()
            Dim metodo As String = "CreateBound"
            If bound = False Then metodo = "CreateUnbound"
            Dim met As MethodInfo = GetType(Line).GetMethod(metodo, tipos)

            If met IsNot Nothing Then
                value = TryCast(met.Invoke(Nothing, parametros), Line)
            Else
                parametros = New Object() {p1, p2, bound}
                tipos = parametros.[Select](Function(a) a.[GetType]()).ToArray()
                value = TryCast(doc.Application.Create.[GetType]().GetMethod("NewLine", tipos).Invoke(doc.Application.Create, parametros), Line)
            End If

            Return value
        End Function

        <Extension()>
        Function CreateWall2(ByVal doc As Document, ByVal curve As Curve, ByVal wallTypeId As ElementId, ByVal levelId As ElementId, ByVal height As Double, ByVal offset As Double, ByVal flip As Boolean, ByVal structural As Boolean) As Wall
            Dim value As Wall = Nothing
            Dim parametros As Object() = New Object() {doc, curve, wallTypeId, levelId, height, offset, flip, structural}
            Dim tipos As Type() = parametros.[Select](Function(a) a.[GetType]()).ToArray()
            Dim met As MethodInfo = GetType(Wall).GetMethod("Create", tipos)

            If met IsNot Nothing Then
                value = TryCast(met.Invoke(Nothing, parametros), Wall)
            Else
                parametros = New Object() {curve, CType(doc.GetElement2(wallTypeId), WallType), CType(doc.GetElement2(levelId), Level), height, offset, flip, structural}
                tipos = parametros.[Select](Function(a) a.[GetType]()).ToArray()
                value = TryCast(doc.Create.[GetType]().GetMethod("NewWall", tipos).Invoke(doc.Create, parametros), Wall)
            End If

            Return value
        End Function

        <Extension()>
        Function CreateArc2(ByVal doc As Document, ByVal p1 As XYZ, ByVal p2 As XYZ, ByVal p3 As XYZ) As Arc
            Dim value As Arc = Nothing
            Dim parametros As Object() = New Object() {p1, p2, p3}
            Dim tipos As Type() = parametros.[Select](Function(a) a.[GetType]()).ToArray()
            Dim metodo As String = "Create"
            Dim met As MethodInfo = GetType(Arc).GetMethod(metodo, tipos)

            If met IsNot Nothing Then
                value = TryCast(met.Invoke(Nothing, parametros), Arc)
            Else
                value = TryCast(doc.Application.Create.[GetType]().GetMethod("NewArc", tipos).Invoke(doc.Application.Create, parametros), Arc)
            End If

            Return value
        End Function

        <Extension()>
        Function GetDecimalSymbol2(ByVal doc As Document) As Char
            Dim valor As Char = ","c
            Dim met As MethodInfo = doc.[GetType]().GetMethod("GetUnits")

            If met IsNot Nothing Then
                Dim temp As Object = met.Invoke(doc, Nothing)
                Dim prop As PropertyInfo = temp.[GetType]().GetProperty("DecimalSymbol")
                Dim o As Object = prop.GetValue(temp, Nothing)

                If o.ToString() = "Comma" Then
                    valor = ","c
                Else
                    valor = "."c
                End If
            Else
                Dim temp As Object = doc.[GetType]().GetProperty("ProjectUnit").GetValue(doc, Nothing)
                Dim prop As PropertyInfo = temp.[GetType]().GetProperty("DecimalSymbolType")
                Dim o As Object = prop.GetValue(temp, Nothing)

                If o.ToString() = "DST_COMMA" Then
                    valor = ","c
                Else
                    valor = "."c
                End If
            End If

            Return valor
        End Function

        <Extension()>
        Sub UnjoinGeometry2(ByVal doc As Document, ByVal firstElement As Element, ByVal secondElement As Element)
            Dim ls As List(Of Type) = doc.[GetType]().Assembly.GetTypes().Where(Function(a) a.IsClass AndAlso a.Name = "JoinGeometryUtils").ToList()
            Dim parametros As Object() = New Object() {doc, firstElement, secondElement}
            Dim tipos As Type() = parametros.[Select](Function(a) a.[GetType]()).ToArray()

            If ls.Count > 0 Then
                Dim t As Type = ls(0)
                Dim met As MethodInfo = t.GetMethod("UnjoinGeometry", tipos)
                met.Invoke(Nothing, parametros)
            End If
        End Sub

        <Extension()>
        Sub JoinGeometry2(ByVal doc As Document, ByVal firstElement As Element, ByVal secondElement As Element)
            Dim ls As List(Of Type) = doc.[GetType]().Assembly.GetTypes().Where(Function(a) a.IsClass AndAlso a.Name = "JoinGeometryUtils").ToList()
            Dim parametros As Object() = New Object() {doc, firstElement, secondElement}
            Dim tipos As Type() = parametros.[Select](Function(a) a.[GetType]()).ToArray()

            If ls.Count > 0 Then
                Dim t As Type = ls(0)
                Dim met As MethodInfo = t.GetMethod("JoinGeometry", tipos)
                met.Invoke(Nothing, parametros)
            End If
        End Sub

        <Extension()>
        Function IsJoined2(ByVal doc As Document, ByVal firstElement As Element, ByVal secondElement As Element) As Boolean
            Dim value As Boolean = False
            Dim ls As List(Of Type) = doc.[GetType]().Assembly.GetTypes().Where(Function(a) a.IsClass AndAlso a.Name = "JoinGeometryUtils").ToList()
            Dim parametros As Object() = New Object() {doc, firstElement, secondElement}
            Dim tipos As Type() = parametros.[Select](Function(a) a.[GetType]()).ToArray()

            If ls.Count > 0 Then
                Dim t As Type = ls(0)
                Dim met As MethodInfo = t.GetMethod("AreElementsJoined", tipos)
                value = CBool(met.Invoke(Nothing, parametros))
            End If

            Return value
        End Function

        <Extension()>
        Function CalculateVolumeArea2(ByVal doc As Document, ByVal value As Boolean) As Boolean
            Dim ls As List(Of Type) = doc.[GetType]().Assembly.GetTypes().Where(Function(a) a.IsClass AndAlso a.Name = "AreaVolumeSettings").ToList()

            If ls.Count > 0 Then
                Dim t As Type = ls(0)
                Dim parametros As Object() = New Object() {doc}
                Dim tipos As Type() = parametros.[Select](Function(a) a.[GetType]()).ToArray()
                Dim met As MethodInfo = t.GetMethod("GetAreaVolumeSettings", tipos)
                Dim temp As Object = met.Invoke(Nothing, parametros)
                temp.[GetType]().GetProperty("ComputeVolumes").SetValue(temp, value, Nothing)
            Else
                Dim prop As PropertyInfo = doc.Settings.[GetType]().GetProperty("VolumeCalculationSetting")
                Dim temp As Object = prop.GetValue(doc.Settings, Nothing)
                prop = temp.[GetType]().GetProperty("VolumeCalculationOptions")
                temp = prop.GetValue(temp, Nothing)
                prop = temp.[GetType]().GetProperty("VolumeComputationEnable")
                prop.SetValue(temp, value, Nothing)
            End If

            Return value
        End Function

        <Extension()>
        Function CreateGroup2(ByVal doc As Document, ByVal elementos As List(Of Element)) As Group
            Dim value As Group = Nothing
            Dim eleset As ElementSet = New ElementSet()

            For Each ele As Element In elementos
                eleset.Insert(ele)
            Next

            Dim col As ICollection(Of ElementId) = elementos.[Select](Function(a) a.Id).ToList()
            Dim obj As Object = doc.Create
            Dim met As MethodInfo = obj.[GetType]().GetMethod("NewGroup", New Type() {col.[GetType]()})

            If met IsNot Nothing Then
                met.Invoke(obj, New Object() {col})
            Else
                met = obj.[GetType]().GetMethod("NewGroup", New Type() {eleset.[GetType]()})
                met.Invoke(obj, New Object() {eleset})
            End If

            Return value
        End Function

        <Extension()>
        Sub Delete2(ByVal doc As Document, ByVal ele As Element)
            Dim obj As Object = doc
            Dim met As MethodInfo = obj.[GetType]().GetMethod("Delete", New Type() {GetType(Element)})

            If met IsNot Nothing Then
                met.Invoke(obj, New Object() {ele})
            Else
                met = obj.[GetType]().GetMethod("Delete", New Type() {GetType(ElementId)})
                met.Invoke(obj, New Object() {ele.Id})
            End If
        End Sub

        <Extension()>
        Function Level2(ByVal ele As Element) As Element
            Dim value As Element = Nothing
            Dim doc As Document = ele.Document
            Dim t As Type = ele.[GetType]()

            If t.GetProperty("Level") IsNot Nothing Then
                value = TryCast(t.GetProperty("Level").GetValue(ele, Nothing), Element)
            Else
                value = doc.GetElement2(CType(t.GetProperty("LevelId").GetValue(ele, Nothing), ElementId))
            End If

            Return value
        End Function

        <Extension()>
        Function Materiais2(ByVal ele As Element) As List(Of Material)
            Dim value As List(Of Material) = New List(Of Material)()
            Dim doc As Document = ele.Document
            Dim t As Type = ele.[GetType]()

            If t.GetProperty("Materials") IsNot Nothing Then
                value = (CType(t.GetProperty("Materials").GetValue(ele, Nothing), IEnumerable)).Cast(Of Material)().ToList()
            Else
                Dim met As MethodInfo = t.GetMethod("GetMaterialIds", New Type() {GetType(Boolean)})
                value = (CType(met.Invoke(ele, New Object() {False}), ICollection(Of ElementId))).[Select](Function(a) doc.GetElement2(a)).Cast(Of Material)().ToList()
            End If

            Return value
        End Function

        <Extension()>
        Function GetParameter2(ByVal ele As Element, ByVal nome_paramentro As String) As Parameter
            Dim value As Parameter = Nothing
            Dim t As Type = ele.[GetType]()
            Dim met As MethodInfo = t.GetMethod("LookupParameter", New Type() {GetType(String)})
            If met Is Nothing Then met = t.GetMethod("get_Parameter", New Type() {GetType(String)})
            value = TryCast(met.Invoke(ele, New Object() {nome_paramentro}), Parameter)

            If value Is Nothing Then
                Dim pas = ele.Parameters.Cast(Of Parameter)().ToList()
                If pas.Exists(Function(a) a.Definition.Name.ToLower() = nome_paramentro.Trim().ToLower()) Then value = pas.First(Function(a) a.Definition.Name.ToLower() = nome_paramentro.Trim().ToLower())
            End If

            Return value
        End Function

        <Extension()>
        Function GetParameter2(ByVal ele As Element, ByVal builtInParameter As BuiltInParameter) As Parameter
            Dim value As Parameter = Nothing
            Dim t As Type = ele.[GetType]()
            Dim met As MethodInfo = t.GetMethod("LookupParameter", New Type() {GetType(BuiltInParameter)})
            If met Is Nothing Then met = t.GetMethod("get_Parameter", New Type() {GetType(BuiltInParameter)})
            value = TryCast(met.Invoke(ele, New Object() {builtInParameter}), Parameter)
            Return value
        End Function

        <Extension()>
        Function GetMaterialArea2(ByVal ele As Element, ByVal m As Material) As Double
            Dim value As Double = 0
            Dim t As Type = ele.[GetType]()
            Dim met As MethodInfo = t.GetMethod("GetMaterialArea", New Type() {GetType(ElementId), GetType(Boolean)})

            If met IsNot Nothing Then
                value = CDbl(met.Invoke(ele, New Object() {m.Id, False}))
            Else
                met = t.GetMethod("GetMaterialArea", New Type() {GetType(Element)})
                value = CDbl(met.Invoke(ele, New Object() {m}))
            End If

            Return value
        End Function

        <Extension()>
        Function GetMaterialVolume2(ByVal ele As Element, ByVal m As Material) As Double
            Dim value As Double = 0
            Dim t As Type = ele.[GetType]()
            Dim met As MethodInfo = t.GetMethod("GetMaterialVolume", New Type() {GetType(ElementId), GetType(Boolean)})

            If met IsNot Nothing Then
                value = CDbl(met.Invoke(ele, New Object() {m.Id, False}))
            Else
                met = t.GetMethod("GetMaterialVolume", New Type() {GetType(ElementId)})
                value = CDbl(met.Invoke(ele, New Object() {m.Id}))
            End If

            Return value
        End Function

        <Extension()>
        Function GetGeometricObjects2(ByVal ele As Element) As List(Of GeometryObject)
            Dim value As List(Of GeometryObject) = New List(Of GeometryObject)()
            Dim op As Options = New Options()
            Dim obj As Object = ele.Geometry(op)
            Dim prop As PropertyInfo = obj.[GetType]().GetProperty("Objects")

            If prop IsNot Nothing Then
                obj = prop.GetValue(obj, Nothing)
                Dim arr As IEnumerable = TryCast(obj, IEnumerable)

                For Each geo As GeometryObject In arr
                    value.Add(geo)
                Next
            Else
                Dim geos As IEnumerable(Of GeometryObject) = TryCast(obj, IEnumerable(Of GeometryObject))

                For Each geo In geos
                    value.Add(geo)
                Next
            End If

            Return value
        End Function

        <Extension()>
        Sub EnableFamilySymbol2(ByVal fsymbol As FamilySymbol)
            Dim met As MethodInfo = fsymbol.[GetType]().GetMethod("Activate")

            If met IsNot Nothing Then
                met.Invoke(fsymbol, Nothing)
            End If
        End Sub

        <Extension()>
        Sub VaryGroup2(ByVal def As InternalDefinition, ByVal doc As Document)
            Dim parametros As Object() = New Object() {doc, True}
            Dim tipos As Type() = parametros.[Select](Function(a) a.[GetType]()).ToArray()
            Dim met As MethodInfo = def.[GetType]().GetMethod("SetAllowVaryBetweenGroups", tipos)

            If met IsNot Nothing Then
                met.Invoke(def, parametros)
            End If
        End Sub

        <Extension()>
        Function GetSource2(ByVal part As Part) As ElementId
            Dim value As ElementId = Nothing
            Dim prop As PropertyInfo = part.[GetType]().GetProperty("OriginalDividedElementId")

            If prop IsNot Nothing Then
                value = TryCast(prop.GetValue(part, Nothing), ElementId)
            Else
                Dim met As MethodInfo = part.[GetType]().GetMethod("GetSourceElementIds")
                Dim temp As Object = met.Invoke(part, Nothing)
                met = temp.[GetType]().GetMethod("First")
                temp = met.Invoke(temp, Nothing)
                prop = temp.[GetType]().GetProperty("HostElementId")
                value = TryCast(prop.GetValue(temp, Nothing), ElementId)
            End If

            Return value
        End Function

        <Extension()>
        Function GetSelection2(ByVal sel As Selection, ByVal doc As Document) As List(Of Element)
            Dim value As List(Of Element) = New List(Of Element)()
            sel.GetElementIds()
            Dim t As Type = sel.[GetType]()

            If t.GetMethod("GetElementIds") IsNot Nothing Then
                Dim met As MethodInfo = t.GetMethod("GetElementIds")
                value = (CType(met.Invoke(sel, Nothing), ICollection(Of ElementId))).[Select](Function(a) doc.GetElement2(a)).ToList()
            Else
                value = (CType(t.GetProperty("Elements").GetValue(sel, Nothing), IEnumerable)).Cast(Of Element)().ToList()
            End If

            Return value
        End Function

        <Extension()>
        Sub SetSelection2(ByVal sel As Selection, ByVal doc As Document, ByVal elementos As ICollection(Of ElementId))
            sel.ClearSelection2()
            Dim parametros As Object() = New Object() {elementos}
            Dim tipos As Type() = parametros.[Select](Function(a) a.[GetType]()).ToArray()
            Dim met As MethodInfo = sel.[GetType]().GetMethod("SetElementIds", tipos)

            If met IsNot Nothing Then
                met.Invoke(sel, parametros)
            Else
                Dim prop As PropertyInfo = sel.[GetType]().GetProperty("Elements")
                Dim temp As Object = prop.GetValue(sel, Nothing)

                If elementos.Count = 0 Then
                    met = temp.[GetType]().GetMethod("Clear")
                    met.Invoke(temp, Nothing)
                Else

                    For Each id As ElementId In elementos
                        Dim elemento As Element = doc.GetElement2(id)
                        parametros = New Object() {elemento}
                        tipos = parametros.[Select](Function(a) a.[GetType]()).ToArray()
                        met = temp.[GetType]().GetMethod("Add", tipos)
                        met.Invoke(temp, parametros)
                    Next
                End If
            End If
        End Sub

        <Extension()>
        Sub ClearSelection2(ByVal sel As Selection)
            Dim prop As PropertyInfo = sel.[GetType]().GetProperty("Elements")

            If prop IsNot Nothing Then
                Dim obj As Object = prop.GetValue(sel, Nothing)
                Dim met As MethodInfo = obj.[GetType]().GetMethod("Clear")
                met.Invoke(obj, Nothing)
            Else
                Dim ids As ICollection(Of ElementId) = New List(Of ElementId)()
                Dim met As MethodInfo = sel.[GetType]().GetMethod("SetElementIds", New Type() {ids.[GetType]()})
                met.Invoke(sel, New Object() {ids})
            End If
        End Sub

        <Extension()>
        Function GetDrawingArea2(ByVal ui As UIApplication) As System.Drawing.Rectangle
            Dim value As System.Drawing.Rectangle = System.Windows.Forms.Screen.PrimaryScreen.Bounds
            Return value
        End Function

        <Extension()>
        Function Duplicate2(ByVal view As View) As ElementId
            Dim value As ElementId = Nothing
            Dim doc As Document = view.Document
            Dim ls As List(Of Type) = doc.[GetType]().Assembly.GetTypes().Where(Function(a) a.IsEnum AndAlso a.Name = "ViewDuplicateOption").ToList()

            If ls.Count > 0 Then
                Dim t As Type = ls(0)
                Dim obj As Object = view
                Dim met As MethodInfo = view.[GetType]().GetMethod("Duplicate", New Type() {t})

                If met IsNot Nothing Then
                    value = TryCast(met.Invoke(obj, New Object() {2}), ElementId)
                End If
            End If

            Return value
        End Function

        <Extension()>
        Sub SetOverlayView2(ByVal view As View, ByVal ids As List(Of ElementId), ByVal Optional cor As Color = Nothing, ByVal Optional espessura As Integer = -1)
            Dim doc As Document = view.Document
            Dim ls As List(Of Type) = doc.[GetType]().Assembly.GetTypes().Where(Function(a) a.IsClass AndAlso a.Name = "OverrideGraphicSettings").ToList()

            If ls.Count > 0 Then
                Dim t As Type = ls(0)
                Dim construtor As ConstructorInfo = t.GetConstructor(New Type() {})
                construtor.Invoke(New Object() {})
                Dim obj As Object = construtor.Invoke(New Object() {})
                Dim met As MethodInfo = obj.[GetType]().GetMethod("SetProjectionLineColor", New Type() {cor.[GetType]()})
                met.Invoke(obj, New Object() {cor})
                met = obj.[GetType]().GetMethod("SetProjectionLineWeight", New Type() {espessura.[GetType]()})
                met.Invoke(obj, New Object() {espessura})
                met = view.[GetType]().GetMethod("SetElementOverrides", New Type() {GetType(ElementId), obj.[GetType]()})

                For Each id As ElementId In ids
                    met.Invoke(view, New Object() {id, obj})
                Next
            Else
                Dim met As MethodInfo = view.[GetType]().GetMethod("set_ProjColorOverrideByElement", New Type() {GetType(ICollection(Of ElementId)), GetType(Color)})
                met.Invoke(view, New Object() {ids, cor})
                met = view.[GetType]().GetMethod("set_ProjLineWeightOverrideByElement", New Type() {GetType(ICollection(Of ElementId)), GetType(Integer)})
                met.Invoke(view, New Object() {ids, espessura})
            End If
        End Sub

        <Extension()>
        Function GetViewTemplateId2(ByVal view As ViewPlan) As ElementId
            Dim value As ElementId = Nothing
            Dim prop As PropertyInfo = view.[GetType]().GetProperty("ViewTemplateId")

            If prop IsNot Nothing Then
                value = TryCast(prop.GetValue(view, Nothing), ElementId)
            End If

            Return value
        End Function

        <Extension()>
        Sub SetViewTemplateId2(ByVal view As ViewPlan, ByVal id As ElementId)
            Dim prop As PropertyInfo = view.[GetType]().GetProperty("ViewTemplateId")

            If prop IsNot Nothing Then
                prop.SetValue(view, id, Nothing)
            End If
        End Sub

        <Extension()>
        Sub FlipWall2(ByVal wall As Wall)
            Dim metodo As String = "Flip"
            Dim met As MethodInfo = GetType(Wall).GetMethod(metodo)

            If met IsNot Nothing Then
                met.Invoke(wall, Nothing)
            Else
                metodo = "flip"
                met = GetType(Wall).GetMethod(metodo)
                met.Invoke(wall, Nothing)
            End If
        End Sub
    End Module
End Namespace