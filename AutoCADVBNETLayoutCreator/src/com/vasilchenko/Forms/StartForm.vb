Imports AutoCADVBNETLayoutCreator.com.vasilchenko.modules

Public Class ufStartForm
    Private Sub StartForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblText.Text = "Данная программа предназначена для автоматического создания слоев печати. " & vbCrLf _
            & vbCrLf _
            & "Для получения ожидаемого результата придерживайтель следующих правил:" & vbCrLf _
            & "1. Все рамки (полилинии или блоки) должны быть в одном 'печатном' слое, слой выбирается при старте программы" & vbCrLf _
            & "2. Все рамки (полилинии или блоки) должны соответствовать ГОСТ-овским размерам" & vbCrLf _
            & "3. Нумерация листов происходит слева-направо/сверху-вниз, согласно листам модели" & vbCrLf _
            & "4. Используется принтер AutoCAD PDF (General Documentation).pc3 (ACAD2016 и выше)" & vbCrLf _
            & "5. Обратите внимание, если файл редактировался в других программах (например NanoCAD)" & vbCrLf _
            & "     данная программа может работать некорректно. " & vbCrLf _
            & "     Для этого встроена функция определения разности координат (консольный запрос после печати первого листа)." & vbCrLf _
            & vbCrLf _
            & "Детали уточняйте у разработчика или работайте методом тыка =)" & vbCrLf _
            & "Программа работает в тестовом режиме. Баги возможны! " & vbCrLf

        Label1.Text = "Для предложений и пожеланий можно писать сюда:" & vbCrLf _
            & "@e-mail: vasylchenko.s@gmail.com"

        lblText.AutoSize = True
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Me.Dispose()
        Dim ufSelector As New ufLayerSelector
        Try
            ufSelector.ShowDialog()
            PolygonsConfigurator.CreateLayoutsList(ufSelector.cbLayers.SelectedItem)
        Catch ex As Exception
            MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
        Finally
            ufSelector.Dispose()
        End Try
    End Sub

End Class