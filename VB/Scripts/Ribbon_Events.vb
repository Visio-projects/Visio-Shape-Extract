Option Strict On
Option Explicit On

Imports System.Windows.Forms

Namespace Scripts

    <Runtime.InteropServices.ComVisible(True)>
    Public Class Ribbon
        Implements Office.IRibbonExtensibility
        Private ribbon As Office.IRibbonUI
        Public Shared ribbonref As Ribbon

        Public Sub New()
        End Sub

        Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
            Return GetResourceText("ShapeExtract.Ribbon.xml")
        End Function

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

        Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
            Me.ribbon = ribbonUI
            ribbonref = Me
            AssemblyInfo.SetAddRemoveProgramsIcon("VisioAddin.ico")
        End Sub

        Public Function GetLabelText(ByVal control As Office.IRibbonControl) As String
            Try
                Select Case control.Id.ToString
                    Case Is = "tabShapeExtract"
                        If Application.ProductVersion.Substring(0, 2) = "15" Then
                            Return My.Application.Info.Title.ToUpper()
                        Else
                            Return My.Application.Info.Title
                        End If
                    Case Is = "txtCopyright"
                        Return "© " & My.Application.Info.Copyright.ToString
                    Case Is = "txtDescription", "btnDescription"
                        Dim version As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
                        Return My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & version
                    Case Is = "txtReleaseDate"
                        Return My.Settings.App_ReleaseDate.ToString("dd-MMM-yyyy hh:mm tt")
                    Case Else
                        Return String.Empty
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty

            End Try

        End Function

        Public Sub OnAction(ByVal Control As Office.IRibbonControl)
            Try
                Select Case Control.Id
                    Case = "btnExtractShapes"
                        Ribbon_Buttons.ExportShapeValues()
                    Case = "btnOpenFolder"
                        Ribbon_Buttons.OpenFile(My.Settings.App_PathExportLocation)
                    Case = "btnOpenFile"
                        Ribbon_Buttons.OpenFile(My.Settings.App_FileExport)
                    Case "btnSettings"
                        Ribbon_Buttons.OpenSettings()
                    Case "btnOpenReadMe"
                        Ribbon_Buttons.OpenReadMe()
                    Case "btnOpenNewIssue"
                        Ribbon_Buttons.OpenNewIssue()
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub InvalidateRibbon()
            ribbon.Invalidate()
        End Sub

    End Class

End Namespace
