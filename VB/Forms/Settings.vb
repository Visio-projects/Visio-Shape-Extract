Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports System.Reflection
Imports ShapeExtract.Scripts

Namespace Forms

    Public Class Settings

        Private Sub FrmSettingsLoad(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            Try
                Me.pgdSettings.SelectedObject = My.Settings
                SetLabelColumnWidth(Me.pgdSettings, 200)
                AssemblyInfo.SetFormIcon(Me, My.Resources.Settings)
                Me.Text = "Settings for " & My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Exit Try

            End Try

        End Sub

        Private Sub FrmSettingsFormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
            Try
                My.Settings.Save()

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Exit Try

            End Try

        End Sub

        Public Sub SetLabelColumnWidth(grid As PropertyGrid, width As Integer)
            Try

                If grid Is Nothing Then
                    Return
                End If

                Dim fi As FieldInfo = grid.[GetType]().GetField("gridView", BindingFlags.Instance Or BindingFlags.NonPublic)
                If fi Is Nothing Then
                    Return
                End If

                Dim view As Control = TryCast(fi.GetValue(grid), Control)
                If view Is Nothing Then
                    Return
                End If

                Dim mi As MethodInfo = view.[GetType]().GetMethod("MoveSplitterTo", BindingFlags.Instance Or BindingFlags.NonPublic)
                If mi Is Nothing Then
                    Return
                End If
                mi.Invoke(view, New Object() {width})

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Exit Try

            End Try
        End Sub

    End Class

End Namespace