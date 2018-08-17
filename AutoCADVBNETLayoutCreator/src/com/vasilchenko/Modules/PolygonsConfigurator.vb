Imports AutoCADVBNETLayoutCreator.com.vasilchenko.classes
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.Linq

Namespace com.vasilchenko.modules

    Module PolygonsConfigurator
        Dim acDocument As Document
        Dim acDatabase As Database
        Dim acEditor As Editor
        Dim acTransaction As Transaction
        Public Function GetAllLayers() As List(Of String)
            acDocument = Application.DocumentManager.MdiActiveDocument
            acDatabase = acDocument.Database
            acEditor = acDocument.Editor

            Dim acLayersList As New List(Of String)

            Using acTransaction = acDatabase.TransactionManager.StartTransaction()
                Dim acLayers As LayerTable = CType(acTransaction.GetObject(acDatabase.LayerTableId, OpenMode.ForRead), LayerTable)
                For Each itemID In acLayers
                    acLayersList.Add(CType(acTransaction.GetObject(itemID, OpenMode.ForRead), LayerTableRecord).Name)
                Next
                acTransaction.Abort()
            End Using
            Return acLayersList
        End Function
        Public Sub CreateLayoutsList(strLayerName As String)

            Using docLock As DocumentLock = acDocument.LockDocument()
                acTransaction = acDatabase.TransactionManager.StartTransaction()

                Dim acTypedValue As TypedValue() = New TypedValue(0) {New TypedValue(DxfCode.LayerName, strLayerName)}
                Dim acSelFilter As SelectionFilter = New SelectionFilter(acTypedValue)
                Dim acPromSelResult As PromptSelectionResult = acEditor.SelectAll(acSelFilter)

                If acPromSelResult.Value Is Nothing Then
                    MsgBox("В данном слое нет объектов", MsgBoxStyle.Critical)
                    Exit Sub
                End If

                Dim acObjIdCol As ObjectIdCollection = New ObjectIdCollection(acPromSelResult.Value.GetObjectIds)

                Dim acPoligonList As New List(Of PrintedPoligonClass)
                Dim acSortedPoligonList As SortedList(Of Double, SortedList(Of Double, PrintedPoligonClass))

                Try
                    'Получаем id всех полигонов (в т.ч. блоков)
                    For Each acCurId As ObjectId In acObjIdCol
                        'добавляем в список и попутно конвертируем в класс
                        acPoligonList.Add(EntityConverter(DirectCast(acTransaction.GetObject(acCurId, OpenMode.ForRead), Entity)))
                    Next

                    'Отсортируем по x и y осям
                    acSortedPoligonList = SortEntitysOnAxis(acPoligonList)

                    'Создаем Layout для каждого из полигонов
                    LayoutsCreator.CreateLayouts(acSortedPoligonList, acDatabase, acTransaction, acEditor)

                    acTransaction.Commit()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                    acTransaction.Abort()
                Finally
                    acTransaction.Dispose()
                End Try
            End Using
        End Sub
        Private Function EntityConverter(acEntity As Entity) As PrintedPoligonClass
            Dim objPrintedPoligon As New PrintedPoligonClass
            Dim acMin As New Point2d
            Dim acMax As New Point2d

            If acEntity.GetType.Name = "BlockReference" Then
                acMin = New Point2d(CInt(acEntity.GeometricExtents.MinPoint.X), CInt(acEntity.GeometricExtents.MinPoint.Y))
                acMax = New Point2d(CInt(acEntity.GeometricExtents.MaxPoint.X), CInt(acEntity.GeometricExtents.MaxPoint.Y))
                objPrintedPoligon.Normal = DirectCast(acEntity, BlockReference).Normal
            ElseIf acEntity.GetType.Name = "Polyline" Then
                Dim acPolyline As Polyline = DirectCast(acEntity, Polyline)
                Dim acPoint As Point2d
                For i As Integer = 0 To acPolyline.NumberOfVertices - 1
                    acPoint = New Point2d(CInt(acPolyline.GetPoint2dAt(i).X), CInt(acPolyline.GetPoint2dAt(i).Y))
                    If i = 0 Then
                        acMin = acPoint
                        acMax = acPoint
                    Else
                        If (acPoint.X <= acMin.X And acPoint.Y <= acMin.Y) Then acMin = acPoint
                        If (acPoint.X >= acMax.X And acPoint.Y >= acMax.Y) Then acMax = acPoint
                    End If
                Next
                objPrintedPoligon.Normal = acPolyline.Normal
            End If

            objPrintedPoligon.ID = acEntity.ObjectId
            objPrintedPoligon.BottonLeft = acMin
            objPrintedPoligon.TopRight = acMax
            objPrintedPoligon.XSize = Math.Round(acMax.X - acMin.X)
            objPrintedPoligon.YSize = Math.Round(acMax.Y - acMin.Y)
            objPrintedPoligon.Area = objPrintedPoligon.XSize * objPrintedPoligon.YSize

            Return objPrintedPoligon
        End Function
        Private Function SortEntitysOnAxis(acPoligonList As List(Of PrintedPoligonClass)) As SortedList(Of Double, SortedList(Of Double, PrintedPoligonClass))
            Dim descendingComparer = Comparer(Of Double).Create(Function(x, y) y.CompareTo(x))
            Dim acSortedPoligonList As New SortedList(Of Double, SortedList(Of Double, PrintedPoligonClass))(descendingComparer)

            For Each acItem In acPoligonList
                Dim queryResult = From key In acSortedPoligonList.Keys
                                  Where key > acItem.BottonLeft.Y - 25 And key < acItem.BottonLeft.Y + 25
                                  Select key
                If queryResult.Count = 0 Then
                    Dim tempList = New SortedList(Of Double, PrintedPoligonClass)
                    tempList.Add(acItem.BottonLeft.X, acItem)
                    acSortedPoligonList.Add(acItem.BottonLeft.Y, tempList)
                Else
                    acSortedPoligonList.Item(queryResult(0)).Add(acItem.BottonLeft.X, acItem)
                End If
            Next

            Return acSortedPoligonList
        End Function

    End Module
End Namespace