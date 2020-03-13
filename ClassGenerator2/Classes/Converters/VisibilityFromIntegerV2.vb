Imports System.Windows
Imports System.Windows.Data

Namespace WPFConverters

    Public Class VisibilityFromIntegerV2
        Implements IValueConverter

        Public Property number As Integer = 1
        Public Property collapsed As Boolean = False

        Public Function IValueConverter_Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert

            Try
                Dim result As Integer = Integer.Parse(value)

                Select Case result
                    Case number
                        Return Visibility.Visible
                    Case Else

                        If (collapsed) Then
                            Return Visibility.Collapsed
                        Else
                            Return Visibility.Hidden
                        End If

                End Select

            Catch ex As Exception

                If (collapsed) Then
                    Return Visibility.Collapsed
                Else
                    Return Visibility.Hidden
                End If

            End Try

        End Function

        Public Function IValueConverter_ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function

    End Class
End Namespace
