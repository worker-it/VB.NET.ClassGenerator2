Imports System.Windows
Imports System.Windows.Data

Namespace WPFConverters

    Public Class VisibilityFromListIntegerV2
        Implements IValueConverter

        Public Property numbers As List(Of Integer) = New List(Of Integer) From {1}
        Public Property collapsed As Boolean = False

        Public Function IValueConverter_Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert

            Try
                Dim valeur As Integer = Integer.Parse(value)
                Dim result As Visibility

                If (collapsed) Then
                    result = Visibility.Collapsed
                Else
                    result = Visibility.Hidden
                End If

                For Each number As Integer In numbers
                    If (number = valeur) Then
                        result = Visibility.Visible
                        Exit For
                    End If
                Next

                Return result

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
