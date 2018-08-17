Imports AutoCADVBNETLayoutCreator.com.vasilchenko.modules
Public Class ufLayerSelector
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cbLayers.Items.AddRange(PolygonsConfigurator.GetAllLayers().ToArray)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        If cbLayers.SelectedItem = "0" Then
            MsgBox("Слой 0 является базовым и не подходящий для слоя рамок", MsgBoxStyle.Critical)
        ElseIf cbLayers.SelectedItem = "" Then
            MsgBox("Слой не выбран", MsgBoxStyle.Critical)
        Else
            Me.Close()
        End If
    End Sub
End Class