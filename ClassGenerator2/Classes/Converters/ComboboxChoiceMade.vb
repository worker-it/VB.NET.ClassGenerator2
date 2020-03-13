Imports System.Windows.Data

Namespace WPFConverters
    Public Class ComboboxChoiceMade
        Implements IValueConverter

        Public Function IValueConverter_Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert

            Try
                If (CInt(value) < 1) Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                Return False
            End Try

        End Function

        Public Function IValueConverter_ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack

            Try
                If (CType(value, Boolean)) Then
                    Return 1
                Else
                    Return -1
                End If
            Catch ex As Exception
                Return 0
            End Try

        End Function
    End Class
End Namespace