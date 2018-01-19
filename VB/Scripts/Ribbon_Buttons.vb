Option Strict On
Option Explicit On

Imports System.IO

Namespace Scripts

    Public Class Ribbon_Buttons

        Public Shared Sub ExportShapeValues()
            Dim outFile As StreamWriter = Nothing
            Try
                Dim shape As Visio.Shape
                Dim cell As Visio.Cell
                Dim nRows As Short = 0
                Dim i As Short = 0
                Dim shpNo As Integer = 0
                Dim promptName As String = String.Empty
                Dim cellName As String = String.Empty
                Dim cellValue As String = String.Empty
                Dim line As String = String.Empty

                If Globals.ThisAddIn.Application.ActiveDocument IsNot Nothing Then
                    Exit Try
                End If

                Dim fileName As String = My.Settings.App_PathExportLocation & "\" & Globals.ThisAddIn.Application.ActiveDocument.Name.ToString & "_exported_" & System.DateTime.Now().ToString("yyyyMMdd_hhmmss") & ".csv"
                outFile = My.Computer.FileSystem.OpenTextFileWriter(fileName, False, System.Text.Encoding.Default)

                outFile.WriteLine("Shape Type, Shape ID, Shape Name, Cell Name, Prompt Name, Cell Value, Label Text")
                For shpNo = 1 To Globals.ThisAddIn.Application.ActivePage.Shapes.Count
                    shape = Globals.ThisAddIn.Application.ActivePage.Shapes(shpNo)
                    nRows = shape.RowCount(CShort(Visio.VisSectionIndices.visSectionProp))
                    For i = 0 To CShort(nRows - 1)
                        cell = shape.CellsSRC(CShort(Visio.VisSectionIndices.visSectionProp), i, 0)
                        cellValue = cell.ResultStr(Visio.VisSectionIndices.visSectionNone)
                        cell = shape.CellsSRC(CShort(Visio.VisSectionIndices.visSectionProp), i, 1)
                        promptName = cell.ResultStr(Visio.VisSectionIndices.visSectionNone)
                        cell = shape.CellsSRC(CShort(Visio.VisSectionIndices.visSectionProp), i, 2)
                        cellName = cell.ResultStr(Visio.VisSectionIndices.visSectionNone)
                        line = ErrorHandler.RemoveString(shape.Name.ToString, ".") & ", " & ErrorHandler.ValidateString(shape.ID.ToString) & ", " & ErrorHandler.ValidateString(shape.Name.ToString) & ", " & ErrorHandler.ValidateString(cellName) & ", " & ErrorHandler.ValidateString(promptName) & ", " & ErrorHandler.ValidateString(cellValue) & ", " & ErrorHandler.ValidateString(shape.Text.ToString)
                        outFile.WriteLine(line)
                    Next i
                Next shpNo
                My.Settings.App_FileExport = fileName

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                outFile.Close()

            End Try

        End Sub

        Public Shared Sub OpenFile(ByVal fileName As String)
            Try
                Dim pStart As New System.Diagnostics.Process
                If fileName = String.Empty Then Exit Try
                pStart.StartInfo.FileName = fileName
                pStart.Start()

            Catch ex As System.ComponentModel.Win32Exception
                Exit Try

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Exit Try

            End Try

        End Sub

        Public Shared Sub OpenReadMe()
            System.Diagnostics.Process.Start(My.Settings.App_PathReadMe)
        End Sub

        Public Shared Sub OpenNewIssue()
            System.Diagnostics.Process.Start(My.Settings.App_PathReportIssue)

        End Sub

        Public Shared Sub OpenSettings()
            Try
                Dim formSettings As New Forms.Settings
                formSettings.ShowDialog()
                Scripts.Ribbon.ribbonref.InvalidateRibbon()

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

    End Class

End Namespace