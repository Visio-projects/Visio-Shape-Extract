Option Strict On
Option Explicit On

Imports System.IO
Imports ShapeExtract
Imports ShapeExtract.Forms

Namespace Scripts

    ''' <summary>
    ''' The ribbon code used for the addin
    ''' </summary>
    ''' <remarks></remarks>
    <Runtime.InteropServices.ComVisible(True)> _
    Public Class Ribbon
        Implements Office.IRibbonExtensibility
        Private ribbon As Office.IRibbonUI

#Region "| Ribbon Events |"

        Public Sub New()
        End Sub

        ''' <summary>
        ''' Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
        ''' </summary>
        ''' <param name="ribbonID">Represents the XML customization file</param>
        ''' <returns>A method that returns a bitmap image for the control id.</returns>
        ''' <remarks></remarks>
        Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
            Return GetResourceText("ShapeExtract.Ribbon.xml")
        End Function
		
        ''' <summary>
        ''' Get resource text
        ''' </summary>
        ''' <param name="resourceName"></param>
        ''' <returns>string</returns>
        ''' <remarks></remarks>
        Private Shared Function GetResourceText(ByVal resourceName As String) As String
            Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
            Dim resourceNames() As String = asm.GetManifestResourceNames()
            For i As Integer = 0 To resourceNames.Length - 1
                If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                    Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                        If resourceReader IsNot Nothing Then
                            Return resourceReader.ReadToEnd()
                        End If
                    End Using
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Load the ribbon
        ''' </summary>
        ''' <param name="ribbonUI">Represents the IRibbonUI instance that is provided by the Microsoft Office application to the Ribbon extensibility code.</param>
        ''' <remarks></remarks>
        Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
            Me.ribbon = ribbonUI
            Call SetAddRemoveProgramsIcon("VisioAddin.ico")
        End Sub

        ''' <summary>
        '''To assign a images to the controls on the ribbon in the xml file
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <returns>A method that returns a bitmap image for the control id.</returns>
        ''' <remarks></remarks>
        Public Function GetButtonImage(ByVal control As Office.IRibbonControl) As System.Drawing.Bitmap
            Try
                Select Case control.Id.ToString
                    Case Is = "btnExtractShapes"
                        Return My.Resources.Resources.Export
                    Case Is = "btnOpenFolder"
                        Return My.Resources.Resources.Folder
                    Case Is = "btnOpenFile"
                        Return My.Resources.Resources.File
                    Case Is = "btnSettings"
                        Return My.Resources.Resources.Settings
                    Case Else
                        Return Nothing
                End Select

            Catch ex As Exception
                Call DisplayErrorMessage(ex)
                Return Nothing

            End Try

        End Function

        ''' <summary>
        ''' To assign text to controls on the ribbon from the xml file
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <returns>A method that returns a string for a label. </returns>
        ''' <remarks></remarks>
        Public Function GetLabelText(ByVal control As Office.IRibbonControl) As String
            Try
                Select Case control.Id.ToString
                    Case Is = "tabShapeExtract"
                        Return My.Application.Info.Title
                    Case Is = "txtCopyright"
                        Return "© " & My.Application.Info.Copyright.ToString
                    Case Is = "txtDescription"
                        Dim strVersion As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
                        Return My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & strVersion
                    Case Is = "txtInstallDate"
                        Dim dteCreateDate As DateTime = System.IO.File.GetLastWriteTime(My.Application.Info.DirectoryPath.ToString & "\" & My.Application.Info.AssemblyName.ToString & ".dll") 'get creation date 
                        Return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt")
                    Case Else
                        Return String.Empty
                End Select

            Catch ex As Exception
                Call DisplayErrorMessage(ex)
                'Console.WriteLine(ex.Message.ToString)
                Return String.Empty

            End Try

        End Function

        ''' <summary>
        ''' To assign the visiblity to controls
        ''' </summary>
        ''' <param name="Control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <returns>A method that returns true or false if the control is visible</returns>
        ''' <remarks></remarks>
        Public Function GetVisible(ByVal control As Office.IRibbonControl) As Boolean
            Try
                Select Case control.Id.ToString
                    Case Is = "ComAddInsDialog"
                        Return My.Settings.Visible_ComAddInsDialog
                    Case Else
                        Return False
                End Select

            Catch ex As Exception
                Call DisplayErrorMessage(ex)
                Return False

            End Try

        End Function

#End Region

#Region "| Ribbon Buttons |"

        ''' <summary>
        ''' Extract entity attributes from file
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <remarks></remarks>
        Public Sub BtnExtractShapes(ByVal control As Office.IRibbonControl)
            Try
                If Globals.ThisAddIn.Application.ActiveDocument IsNot Nothing Then
                    Dim filePath As String = My.Settings.ExportLocation & "\" & Globals.ThisAddIn.Application.ActiveDocument.Name.ToString & "_exported_" & System.DateTime.Now().ToString("yyyyMMdd_hhmmss") & ".csv"
                    Call ExportShapeValues(filePath)
                End If

            Catch ex As Exception
                Call DisplayErrorMessage(ex)

            End Try

        End Sub

        ''' <summary>
        ''' Open a selected folder
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <remarks></remarks>
        Public Sub BtnOpenFolder(ByVal control As Office.IRibbonControl)
            Call OpenFile(My.Settings.ExportLocation)
        End Sub

        ''' <summary>
        ''' Open the exported attributes file
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <remarks></remarks>
        Public Sub BtnOpenFile(ByVal control As Office.IRibbonControl)
            Call OpenFile(My.Settings.ExportFile)
        End Sub

        ''' <summary>
        ''' show the settings form
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <remarks></remarks>
        Public Sub BtnSettings(ByVal control As Office.IRibbonControl)
            Try
                Dim formSettings As New Settings
                formSettings.ShowDialog()
                ribbon.Invalidate()

            Catch ex As Exception
                Call DisplayErrorMessage(ex)

            End Try

        End Sub

        ''' <summary>
        ''' Opens a api help file
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <remarks></remarks>
        Public Sub OpenHelpApiFile(ByVal control As Office.IRibbonControl)
            Try
                Dim fileName As String = Path.Combine(GetCurrentLocation("ClickOnceLocation"), "Help\\Api Help.chm")
                Call OpenFile(fileName)

            Catch ex As Exception
                Call DisplayErrorMessage(ex)

            End Try
        End Sub

        ''' <summary>
        ''' Opens an as built help file
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <remarks></remarks>
        Public Sub OpenHelpAsBuiltFile(ByVal control As Office.IRibbonControl)
            Try
                Dim fileName As String = Path.Combine(GetCurrentLocation("ClickOnceLocation"), "Help\\As Built.docx")
                Call OpenFile(fileName)

            Catch ex As Exception
                Call DisplayErrorMessage(ex)

            End Try
        End Sub

#End Region

#Region "| Subroutines |"

        ''' <summary>
        ''' Export attribute values from Visio shapes
        ''' Example: Call ExportShapeValues("C:\Temp\YourFileNameHere_ExportEntityAttributes.txt")
        ''' </summary>
        ''' <param name="fileName">The selected file name</param>
        ''' <returns>Has there been an error during the export?</returns>
        ''' <remarks></remarks>
        Public Function ExportShapeValues(ByVal fileName As String) As Boolean
            Dim outFile As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(fileName, False, System.Text.Encoding.Default)
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
                        line = RemoveString(shape.Name.ToString, ".") & ", " & ValidateString(shape.ID.ToString) & ", " & ValidateString(shape.Name.ToString) & ", " & ValidateString(cellName) & ", " & ValidateString(promptName) & ", " & ValidateString(cellValue) & ", " & ValidateString(shape.Text.ToString)
                        outFile.WriteLine(line)
                    Next i
                Next shpNo
                My.Settings.ExportFile = fileName
                Return True

            Catch ex As Exception
                Call DisplayErrorMessage(ex)
                Return False

            Finally
                outFile.Close()

            End Try

        End Function

        ''' <summary>
        ''' Remove any characters that will crash the row
        ''' </summary>
        ''' <param name="text">The string to validate</param>
        ''' <returns>The corrected string</returns>
        ''' <remarks></remarks>
        Public Function ValidateString(ByVal text As String) As String
            Try
                If String.IsNullOrEmpty(text) Then
                    text = String.Empty
                Else
                    text = text.Replace(vbCr, "").Replace(vbLf, "")
                    text = text.Replace(",", ";")
                    text = text.Trim
                End If
                Return text

            Catch ex As Exception
                'Call DisplayErrorMessage(ex)
                Return text

            Finally

            End Try

        End Function

        ''' <summary>
        ''' Remove any characters that will crash the row
        ''' </summary>
        ''' <param name="textValue">The string to evaluate</param>
        ''' <param name="charValue">The string to remove</param>
        ''' <returns>The corrected string</returns>
        ''' <remarks></remarks>
        Public Function RemoveString(ByVal textValue As String, Optional ByVal charValue As String = ".") As String
            Try
                If textValue.Contains(charValue) Then
                    Return textValue.Substring(0, textValue.IndexOf("."))
                Else
                    Return textValue
                End If

            Catch ex As Exception
                'Call DisplayErrorMessage(ex)
                Return textValue

            Finally

            End Try

        End Function

        ''' <summary>
        ''' open a file from the source list
        ''' </summary>
        ''' <param name="fileName">The selected file to open</param>
        ''' <remarks></remarks>
        Public Sub OpenFile(ByVal fileName As String)
            Try
                Dim pStart As New System.Diagnostics.Process
                If fileName = String.Empty Then Exit Try
                pStart.StartInfo.FileName = fileName
                pStart.Start()
                'MsgBox(pstrFile, vbCritical, "file path")

            Catch ex As System.ComponentModel.Win32Exception
                'MessageBox.Show("No application is assicated to this file type." & vbCrLf & vbCrLf & pstrFile, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try

            Catch ex As Exception
                Call DisplayErrorMessage(ex)
                Exit Try

                'Finally
                '    pStart = Nothing

            End Try

        End Sub

        ''' <summary> 
        ''' Returns the assembly location string based on the type of location
        ''' </summary>
        ''' <param name="locationType">Represents assembly location type </param>
        ''' <returns>A method that returns a string of the current location </returns> 
        ''' <remarks></remarks>
        Public Shared Function GetCurrentLocation(locationType As String) As String
            Try
                'Get the assembly information
                Dim assemblyInfo As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()

                'CodeBase is the location of the ClickOnce deployment files
                Dim uriCodeBase As New Uri(assemblyInfo.CodeBase)

                Select Case locationType
                    Case "AssemblyLocation"
                        Return assemblyInfo.Location
                        'Location is where the assembly is run from 
                    Case "ClickOnceLocation"
                        Return Path.GetDirectoryName(uriCodeBase.LocalPath.ToString())
                    Case Else
                        Return String.Empty

                End Select
            Catch generatedExceptionName As Exception

                Return String.Empty
            End Try

        End Function

#End Region

    End Class

End Namespace
