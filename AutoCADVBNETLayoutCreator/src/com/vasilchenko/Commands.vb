Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime

Namespace com.vasilchenko

    Public Class Commands
        '<DllImport("acad.exe", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="acedTrans")>
        <DllImport("accore.dll", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="acedTrans")>
        Public Shared Function acedTrans(ByVal point As Double(), ByVal fromRb As IntPtr, ByVal toRb As IntPtr, ByVal disp As Integer, ByVal result As Double()) As Integer

        End Function

        <CommandMethod("ASU_LayoutCreator", CommandFlags.Session)>
        Public Shared Sub Main()
            Dim swTimer = New Stopwatch
            swTimer.Start()

            Dim strMessage As String = ""

            Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            Application.AcadApplication.ActiveDocument.SendCommand("(command ""PSLTSCALE"")(command ""0"")" & vbCr)

            Dim ufStart As New ufStartForm
            Try
                ufStart.ShowDialog()
            Catch ex As Exception
                MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                strMessage = "[ERROR] Не удалось завершить программу корректно"
            Finally
                swTimer.Stop()
                If strMessage = "" Then strMessage = "Успех! Программа выполнилась за {HH:MM:SS.ms}" & swTimer.Elapsed.ToString
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("---------------------------------------------")
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(strMessage)
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("---------------------------------------------")
                ufStart.Dispose()
            End Try
        End Sub
    End Class

End Namespace
