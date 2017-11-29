Option Strict On
Option Explicit On

Imports System.Environment
Imports System.Windows.Forms

Namespace Scripts

    Public Class ErrorHandler

        Public Shared Sub DisplayMessage(ex As Exception)
            Dim sf As New System.Diagnostics.StackFrame(1)
            Dim caller As System.Reflection.MethodBase = sf.GetMethod()
            Dim currentProcedure As String = (caller.Name).Trim()
            Dim currentFileName As String = "" 'AssemblyInfo.GetCurrentFileName()
            Dim errorMessageDescription As String = ex.ToString()
            errorMessageDescription = System.Text.RegularExpressions.Regex.Replace(errorMessageDescription, "\r\n+", " ")
            Dim msg As String = "Contact your system administrator. A record has been created in the log file." + Environment.NewLine
            msg += (Convert.ToString("Procedure: ") & currentProcedure) + Environment.NewLine
            msg += "Description: " + ex.ToString() + Environment.NewLine
            MessageBox.Show(msg, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])

        End Sub

        Public Shared Function RemoveString(ByVal text As String, Optional ByVal charValue As String = ".") As String
            Try
                If text.Contains(charValue) Then
                    Return text.Substring(0, text.IndexOf("."))
                Else
                    Return text
                End If

            Catch ex As Exception
                Return text

            Finally

            End Try

        End Function

        Public Shared Function ValidateString(ByVal text As String) As String
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
                Return text

            Finally

            End Try

        End Function

    End Class

End Namespace
