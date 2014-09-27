' (C) Copyright 2011 by  
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(WJL_PolylineTools.MyCommands))> 
Namespace WJL_PolylineTools

    Public Class MyCommands

        ' Modal Command with pickfirst selection
        <CommandMethod("PlineReduce", CommandFlags.Modal + CommandFlags.UsePickSet)> _
        Public Sub PlineReduce()

            '' Counters
            Dim iPLineCount As Integer = 0
            Dim iVertecesRemoved As Integer = 0

            '' get document, database & editor
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database

            '' Select objects
            Dim pSelOpt As New PromptSelectionOptions()
            pSelOpt.MessageForAdding = "Select polylines to simplify: "
            Dim pSelRes As PromptSelectionResult = ed.GetSelection(pSelOpt)

            
                If (pSelRes.Status = PromptStatus.OK) Then
                    '' There are selected entities

                    '' Get tolerance
                    Dim Tolerance As Double
                    Dim GetDoubleMessage As String = vbLf & "Enter maximum vertex deviation distance for reduction:"
                    Dim GetDoubleOpts As New PromptDoubleOptions(GetDoubleMessage)
                    GetDoubleOpts.AllowZero = False
                    GetDoubleOpts.AllowNegative = False
                    GetDoubleOpts.UseDefaultValue = True
                    Dim GetDoubleResult As PromptDoubleResult = ed.GetDouble(GetDoubleOpts)

                    If GetDoubleResult.Status = PromptStatus.OK Then
                        Tolerance = GetDoubleResult.Value
                    Else
                        Exit Sub
                End If

                Dim tr As Transaction = db.TransactionManager.StartTransaction()
                Using tr
                    '' Simplify polylines
                    Dim acSSet As SelectionSet = pSelRes.Value
                    For Each acSSObj As SelectedObject In acSSet
                        Dim obj As DBObject = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)

                        Dim lwp As Polyline = TryCast(obj, Polyline)
                        If lwp IsNot Nothing Then
                            Dim acPoly As Polyline = ReducePolyline(lwp, Tolerance)
                            If acPoly IsNot Nothing Then
                                iPLineCount += 1

                                '' Open the Block table for read
                                Dim acBlkTbl As BlockTable
                                acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead)

                                '' Open the Block table record Model space for write
                                Dim acBlkTblRec As BlockTableRecord
                                acBlkTblRec = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                                '' Add the new object to the block table record and the transaction
                                acBlkTblRec.AppendEntity(acPoly)
                                tr.AddNewlyCreatedDBObject(acPoly, True)

                                '' Tally the number of points removed
                                iVertecesRemoved += lwp.NumberOfVertices - acPoly.NumberOfVertices

                                '' Erase the original polyline
                                lwp.UpgradeOpen()
                                lwp.Erase()
                            End If
                        End If
                    Next
                    tr.Commit()
                End Using
                ed.WriteMessage(vbLf & "{0} vertices removed from {1} polylines in total.", iVertecesRemoved, iPLineCount)
            Else
                ' There are no selected entities
                Exit Sub
            End If

        End Sub

        Function ReducePolyline(lwp As Polyline, Tolerance As Double) As Polyline
            ' Use a for loop to get each vertex, one by one
            Dim Points As New List(Of Point2d)
            Dim vn As Integer = lwp.NumberOfVertices
            For i As Integer = 0 To vn - 1
                ' Could also get the 3D point here
                Dim vertex As Point2d = lwp.GetPoint2dAt(i)
                Points.Add(vertex)
                'ed.WriteMessage(vbLf & vertex.ToString())
            Next
            If Points IsNot Nothing Then

                Dim ReducedPoints As New List(Of Point2d)
                ReducedPoints = DouglasPeuckerReduction(Points, Tolerance)

                '' Clone the selected polyline
                Dim acPoly As Polyline = lwp.Clone

                '' Remove all but the last vertex from the cloned polyline
                Dim iVerticesInSelected As Integer = acPoly.NumberOfVertices
                For index As Integer = 1 To iVerticesInSelected - 1
                    acPoly.RemoveVertexAt(1)
                Next

                Dim count As Integer = 0
                For Each Vertex As Point2d In ReducedPoints
                    acPoly.AddVertexAt(count, Vertex, 0, 0, 0)
                    count += 1
                Next

                'Remove the last original point from the cloned polyline
                Dim iLastVertex As Integer = acPoly.NumberOfVertices - 1
                acPoly.RemoveVertexAt(iLastVertex)

                '' Re-apply global width
                acPoly.ConstantWidth = lwp.ConstantWidth

                Return acPoly
            Else
                Return Nothing
            End If

        End Function

        Function ExtractVertices(ByVal obj As DBObject) As List(Of Point2d)
            Dim Points As New List(Of Point2d)
            ' If a "lightweight" (or optimized) polyline
            Dim lwp As Polyline = TryCast(obj, Polyline)

            If lwp IsNot Nothing Then
                ' Use a for loop to get each vertex, one by one
                Dim vn As Integer = lwp.NumberOfVertices
                For i As Integer = 0 To vn - 1
                    ' Could also get the 3D point here
                    Dim vertex As Point2d = lwp.GetPoint2dAt(i)
                    Points.Add(vertex)
                    'ed.WriteMessage(vbLf & vertex.ToString())
                Next
                Return Points
            Else
                '' If an old-style, 2D polyline
                'Dim p2d As Polyline2d = TryCast(obj, Polyline2d)
                'If p2d IsNot Nothing Then
                '    ' Use foreach to get each contained vertex
                '    For Each vId As ObjectId In p2d
                '        Dim v2d As Vertex2d = DirectCast(tr.GetObject(vId, OpenMode.ForRead), Vertex2d)
                '        'ed.WriteMessage(vbLf & v2d.Position.ToString())
                '        Dim vertex As New Point2d(v2d.Position.X, v2d.Position.Y)
                '        Points.Add(vertex)
                '    Next
                'Else

                '    ' If an old-style, 3D polyline
                '    Dim p3d As Polyline3d = TryCast(obj, Polyline3d)
                '    If p3d IsNot Nothing Then
                '        ' Use foreach to get each contained vertex
                '        For Each vId As ObjectId In p3d
                '            Dim v3d As PolylineVertex3d = DirectCast(tr.GetObject(vId, OpenMode.ForRead), PolylineVertex3d)
                '            ed.WriteMessage(vbLf & v3d.Position.ToString())
                '        Next
                '    End If
                'End If
                Return Nothing
            End If

        End Function


        Public Shared Function DouglasPeuckerReduction(ByVal Points As List(Of Point2d), ByVal Tolerance As [Double]) As List(Of Point2d)
            If Points Is Nothing OrElse Points.Count < 3 Then
                Return Points
            End If

            Dim firstPoint As Integer = 0
            Dim lastPoint As Integer = Points.Count - 1
            Dim pointIndexsToKeep As New List(Of Int32)()

            'Add the first and last index to the keepers
            pointIndexsToKeep.Add(firstPoint)
            pointIndexsToKeep.Add(lastPoint)

            'The first and the last Point2d cannot be the same
            While Points(firstPoint).Equals(Points(lastPoint))
                lastPoint -= 1
            End While

            DouglasPeuckerReduction(Points, firstPoint, lastPoint, Tolerance, pointIndexsToKeep)

            Dim returnPoints As New List(Of Point2d)()
            pointIndexsToKeep.Sort()
            For Each index As Int32 In pointIndexsToKeep
                returnPoints.Add(Points(index))
            Next

            Return returnPoints
        End Function


        Private Shared Sub DouglasPeuckerReduction(ByVal points As List(Of Point2d), ByVal firstPoint As Int32, ByVal lastPoint As Int32, ByVal tolerance As [Double], ByRef pointIndexsToKeep As List(Of Int32))
            Dim maxDistance As Double = 0
            Dim indexFarthest As Integer = 0

            For index As Integer = firstPoint To lastPoint - 1
                Dim distance As Double = PerpendicularDistance(points(firstPoint), points(lastPoint), points(index))
                If distance > maxDistance Then
                    maxDistance = distance
                    indexFarthest = index
                End If
            Next

            If maxDistance > tolerance AndAlso indexFarthest <> 0 Then
                'Add the largest Point2d that exceeds the tolerance
                pointIndexsToKeep.Add(indexFarthest)

                DouglasPeuckerReduction(points, firstPoint, indexFarthest, tolerance, pointIndexsToKeep)
                DouglasPeuckerReduction(points, indexFarthest, lastPoint, tolerance, pointIndexsToKeep)
            End If
        End Sub

        Public Shared Function PerpendicularDistance(ByVal Point1 As Point2d, ByVal Point2 As Point2d, ByVal Point2d As Point2d) As Double
            'Area = |(1/2)(x1y2 + x2y3 + x3y1 - x2y1 - x3y2 - x1y3)|   *Area of triangle
            'Base = v((x1-x2)²+(x1-x2)²)                               *Base of Triangle*
            'Area = .5*Base*H                                          *Solve for height
            'Height = Area/.5/Base

            Dim area As Double = Math.Abs(0.5 * (Point1.X * Point2.Y + Point2.X * Point2d.Y + Point2d.X * Point1.Y - Point2.X * Point1.Y - Point2d.X * Point2.Y - Point1.X * Point2d.Y))
            Dim bottom As Double = Math.Sqrt(Math.Pow(Point1.X - Point2.X, 2) + Math.Pow(Point1.Y - Point2.Y, 2))
            Dim height As Double = area / bottom * 2

            Return height


        End Function

    End Class

End Namespace