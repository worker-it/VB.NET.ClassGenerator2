Imports System.Windows
Imports System.Windows.Data

Namespace WPFConverters

    Public Class VisibilityInverterConverter
        Implements IValueConverter

        Public Property collapsed As Boolean = False

        Public Function IValueConverter_Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert

            Try
                Dim aVisibility As Visibility = CType(value, Visibility)
                If (aVisibility = Visibility.Visible) Then

                    If (collapsed) Then
                        Return Visibility.Collapsed
                    Else
                        Return Visibility.Hidden
                    End If

                Else
                    Return Visibility.Visible
                End If
            Catch ex As Exception
                Return value
            End Try

        End Function

        Public Function IValueConverter_ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack

            Try
                Dim aVisibility As Visibility = CType(value, Visibility)
                If (aVisibility = Visibility.Visible) Then

                    If (collapsed) Then
                        Return Visibility.Collapsed
                    Else
                        Return Visibility.Hidden
                    End If

                Else
                    Return Visibility.Visible
                End If
            Catch ex As Exception
                Return value
            End Try

        End Function

    End Class
End Namespace
