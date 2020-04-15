Imports System.Windows.Forms
Imports AutoCADVBNETLayoutCreator.com.vasilchenko.classes
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ApplicationServices

Namespace com.vasilchenko.modules
    Module LayoutsCreator

        Private strPlotDeviceName As String = "AutoCAD PDF (General Documentation).pc3"
        Public Sub CreateLayouts(acSortedPoligonList As SortedList(Of Double, SortedList(Of Double, PrintedPoligonClass)),
                                 acDatabase As Database, acTransaction As Transaction, acEditor As Editor)
            Dim blnTest As Boolean = True
            Dim dblOffset As Double() = New Double(1) {0, 0}
            Dim acPrintedArea As New Extents2d

            Dim acLayoutMgr As LayoutManager = LayoutManager.Current
            Dim acLayoutObjID As ObjectId

            'If acLayoutMgr.LayoutExists("TempLayout") Then
            Dim lays As DBDictionary = acTransaction.GetObject(acDatabase.LayoutDictionaryId, OpenMode.ForRead)
            If lays.Contains("TempLayout") Then
                acLayoutObjID = acLayoutMgr.GetLayoutId("TempLayout")
            Else
                acLayoutObjID = acLayoutMgr.CreateLayout("TempLayout")
            End If
            acLayoutMgr.CurrentLayout = "TempLayout"

            DeleteExistingLayouts(acDatabase, acTransaction, acEditor, acLayoutMgr)

            Dim intLayoutNumber As Integer = 1
            For Each curItemLine In acSortedPoligonList.Values
                For Each curItem In curItemLine.Values
                    Dim acLayout As Layout
                    acLayoutObjID = acLayoutMgr.CreateLayout("Layout" & intLayoutNumber)
                    acLayout = acTransaction.GetObject(acLayoutObjID, OpenMode.ForRead)
                    acLayoutMgr.CurrentLayout = acLayout.LayoutName
                                    
                    'Apply plot settings to the provided layout
                    SetPlotSettings(acEditor, curItem, acLayout, acDatabase, dblOffset, acPrintedArea)
                    If blnTest Then
                        Dim acPrmtStrOptns As PromptStringOptions = New PromptStringOptions(vbCrLf & "Всё хорошо?" & vbCrLf _
                                                                                          & "Да(Д)/Yes(Y) - продолжаем печать" & vbCrLf _
                                                                                          & "Нет(Н)/No(N) - выбрать левую нижнюю точку ПЕРВОЙ рамки для определения смещения" & vbCrLf _
                                                                                          & "Выход(В)/Exit(E) - Завершить выполение" & vbCrLf _
                                                                                          & "Choose your destiny")

                        acPrmtStrOptns.AllowSpaces = True
                        Dim acPrmtRslt As PromptResult = acEditor.GetString(acPrmtStrOptns)
                        If acPrmtRslt.StringResult.ToLower.Equals("нет") OrElse acPrmtRslt.StringResult.ToLower.Equals("н") OrElse
                            acPrmtRslt.StringResult.ToLower.Equals("no") OrElse acPrmtRslt.StringResult.ToLower.Equals("n") Then
                            dblOffset = GetOffset(acEditor, acPrintedArea, "Выберите базовую точку для первого ряда")
                            SetPlotSettings(acEditor, curItem, acLayout, acDatabase, dblOffset, acPrintedArea)
                        ElseIf acPrmtRslt.StringResult.ToLower.Equals("выход") OrElse acPrmtRslt.StringResult.ToLower.Equals("в") OrElse
                            acPrmtRslt.StringResult.ToLower.Equals("exit") OrElse acPrmtRslt.StringResult.ToLower.Equals("e") Then
                            Exit Sub
                        End If
                        blnTest = False
                    End If
                                      
                    'Core.Application.SetSystemVariable("PSLTSCALE", 0)
                    ApplicationServices.Application.AcadApplication.ActiveDocument.SendCommand("PSLTSCALE 0 " & vbCr)
                    'ApplicationServices.Application.AcadApplication.ActiveDocument.SendCommand(vbCr)
                    acEditor.Regen()

                    intLayoutNumber = intLayoutNumber + 1
                Next
            Next

            acLayoutMgr.DeleteLayout("TempLayout")
        End Sub

        Private Function GetOffset(acEditor As Editor, acPrintedArea As Extents2d, strPromtMessage As String) As Double()
            Dim dblOffset As Double()
            Dim acPrmtPntOptns As PromptPointOptions = New PromptPointOptions(strPromtMessage)
            Dim acPrmtPntRslt As PromptPointResult = acEditor.GetPoint(acPrmtPntOptns)
            Dim acBasicPoint As Point3d = acPrmtPntRslt.Value

            dblOffset = New Double(1) {acPrintedArea.MinPoint.X - acBasicPoint.X, acPrintedArea.MinPoint.Y - acBasicPoint.Y}
            Return dblOffset
        End Function

        Private Sub SetPlotSettings(acEditor As Editor, curItem As PrintedPoligonClass, acLayout As Layout, acDatabase As Database, ByRef dblOffset As Double(), ByRef acPrintedArea As Extents2d)
            Dim strCanonicalName As String = SelectCanonicalName(curItem, acLayout.LayoutName, acEditor)
            Using acPlotStg As PlotSettings = New PlotSettings(acLayout.ModelType)
                acPlotStg.CopyFrom(acLayout)

                ' Specify if plot styles should be displayed on the layout
                acPlotStg.ShowPlotStyles = True
                acPlotStg.ScaleLineweights = False

                Dim acPlotStgValidator As PlotSettingsValidator = PlotSettingsValidator.Current

                ' проверяем существование принтера и устанавливаем его
                Dim acDevice = acPlotStgValidator.GetPlotDeviceList
                If acDevice.Contains(strPlotDeviceName) Then
                    acPlotStgValidator.SetPlotConfigurationName(acPlotStg, strPlotDeviceName, Nothing)
                    acPlotStgValidator.RefreshLists(acPlotStg)
                Else
                    acEditor.WriteMessage("Принтера " & strPlotDeviceName & " не существует")
                    Throw New Exception("PlotDeviceNameError", New ArgumentNullException)
                End If

                ' проверяем существование формата и устанавливаем его
                Dim acCanMediaName = acPlotStgValidator.GetCanonicalMediaNameList(acPlotStg)
                If acCanMediaName.Contains(strCanonicalName) Then
                    acPlotStgValidator.SetCanonicalMediaName(acPlotStg, strCanonicalName)
                Else
                    acEditor.WriteMessage("Формата " & strCanonicalName & " не существует")
                    Throw New Exception("CanonicalMediaNameError", New ArgumentNullException)
                End If

                ' устанавливаем стиль печати
                acPlotStgValidator.SetCurrentStyleSheet(acPlotStg, "monochrome.ctb")

                'Устанавливаем ориентацию
                acPlotStgValidator.SetPlotRotation(acPlotStg, SelectPlotRotation(curItem))
                'переганяем координаты
                acPrintedArea = TranslateCoordinates(curItem, acDatabase, dblOffset)
                acPlotStgValidator.SetPlotWindowArea(acPlotStg, acPrintedArea)

                'выбираем окно для печати
                acPlotStgValidator.SetPlotType(acPlotStg, DatabaseServices.PlotType.Window)

                'выставляем заполнение по листу
                acPlotStgValidator.SetUseStandardScale(acPlotStg, True)
                acPlotStgValidator.SetStdScaleType(acPlotStg, StdScaleType.ScaleToFit)

                'центрируем
                acPlotStgValidator.SetPlotCentered(acPlotStg, True)

                ' Rebuild plotter, plot style, and canonical media lists 
                ' (must be called before setting the plot style)
                acPlotStgValidator.RefreshLists(acPlotStg)

                'Copy the PlotSettings data back to the Layout
                Dim blnUpgraded = False
                If Not acLayout.IsWriteEnabled Then
                    acLayout.UpgradeOpen()
                    blnUpgraded = True
                End If

                acLayout.CopyFrom(acPlotStg)
                If blnUpgraded Then acLayout.DowngradeOpen()

                Using view As ViewTableRecord = acEditor.GetCurrentView()
                    view.Width = acPrintedArea.MaxPoint.X - acPrintedArea.MinPoint.X
                    view.Height = acPrintedArea.MaxPoint.Y - acPrintedArea.MinPoint.Y
                    view.CenterPoint = New Point2d((acPrintedArea.MaxPoint.X + acPrintedArea.MinPoint.X) / 2.0, (acPrintedArea.MaxPoint.Y + acPrintedArea.MinPoint.Y) / 2.0)
                    acEditor.SetCurrentView(view)
                End Using

            End Using

        End Sub

        Private Function TranslateCoordinates(curItem As PrintedPoligonClass, acDatabase As Database, dblOffset As Double()) As Extents2d

            Dim acMSBotton As Point3d = New Point3d(curItem.BottonLeft.X, curItem.BottonLeft.Y, 0)
            Dim acMSTop As Point3d = New Point3d(curItem.TopRight.X, curItem.TopRight.Y, 0)
            'translate to the DCS of Paper Space (PSDCS) RTSHORT=3 from
            'the DCS of the current model space viewport RTSHORT=2
            Dim acRBFrom As New ResultBuffer(New TypedValue(5003, 2)), acRBTo As New ResultBuffer(New TypedValue(5003, 3))
            Dim acPSBotton As Double() = New Double() {0, 0, 0}
            Dim acPSTop As Double() = New Double() {0, 0, 0}

            ' Transform points...
            Commands.acedTrans(acMSBotton.ToArray(), acRBFrom.UnmanagedObject, acRBTo.UnmanagedObject, 0, acPSBotton)
            Commands.acedTrans(acMSTop.ToArray(), acRBFrom.UnmanagedObject, acRBTo.UnmanagedObject, 0, acPSTop)

            Return New Extents2d(New Point2d(acPSBotton(0) - dblOffset(0), acPSBotton(1) - dblOffset(1)),
                                 New Point2d(acPSTop(0) - dblOffset(0), acPSTop(1) - dblOffset(1)))
        End Function

        Private Sub DeleteExistingLayouts(acDatabase As Database, acTransaction As Transaction, acEditor As Editor, acLayoutMgr As LayoutManager)
            Dim lays As DBDictionary = acTransaction.GetObject(acDatabase.LayoutDictionaryId, OpenMode.ForRead)
            For Each item As DBDictionaryEntry In lays
                If item.Key <> "Model" AndAlso item.Key <> "TempLayout" Then
                    acLayoutMgr.DeleteLayout(item.Key)
                    acEditor.WriteMessage("Лист " & item.Key & " удален" & "\n")
                End If
            Next
        End Sub
        Private Function SelectPlotRotation(curItem As PrintedPoligonClass) As PlotRotation
            If curItem.XSize > curItem.YSize Then
                Return PlotRotation.Degrees000
            Else
                Return PlotRotation.Degrees180
            End If
        End Function
        Private Function SelectCanonicalName(curItem As PrintedPoligonClass, layoutName As String, acEditor As Editor) As String
            Select Case curItem.Area
                Case 62360 To 62379
                    Return "ISO_full_bleed_A4_(" & curItem.XSize & ".00_x_" & curItem.YSize & ".00_MM)"
                Case 124730 To 124749
                    Return "ISO_full_bleed_A3_(" & curItem.XSize & ".00_x_" & curItem.YSize & ".00_MM)"
                Case 249470 To 249489
                    Return "ISO_full_bleed_A2_(" & curItem.XSize & ".00_x_" & curItem.YSize & ".00_MM)"
                Case 499546 To 499564
                    Return "ISO_full_bleed_A1_(" & curItem.XSize & ".00_x_" & curItem.YSize & ".00_MM)"
                Case Else
                    Return "ISO_full_bleed_A4_(210.00_x_297.00_MM)"
                    acEditor.WriteMessage("На листе " & layoutName & " рамка установлена по-умолчанию")
            End Select
        End Function
    End Module
End Namespace