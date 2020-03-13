Imports System.Windows
Imports System.Windows.Data

Namespace WPFConverters

    Public Class VisibilityFromBoolean
        Implements IValueConverter

        Public Property CollapseControl As Boolean = False

        Public Function IValueConverter_Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert

            Try
                Dim aVisibility As Visibility = CType(CType(value, Boolean), Visibility)
                If CBool(aVisibility) Then
                    Return Visibility.Visible
                ElseIf (CollapseControl) Then
                    Return Visibility.Collapsed
                Else
                    Return Visibility.Hidden
                End If
            Catch ex As Exception
                Return Visibility.Visible
            End Try

        End Function

        Public Function IValueConverter_ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack

            Try
                Dim aVisibility As Visibility = CType(value, Visibility)
                If (aVisibility = Visibility.Visible) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Return True
            End Try

        End Function

    End Class
End Namespace
