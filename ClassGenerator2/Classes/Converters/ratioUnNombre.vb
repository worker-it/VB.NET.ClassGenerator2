Imports System.Globalization
Imports System.Windows.Data

Namespace WPFConverters


    Public Class ratioUnNombre
        Implements IValueConverter

        Public Property A As Double

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
            Dim a1 As Double = GetDoubleValue(parameter, A)
            Dim x As Double = GetDoubleValue(value, 0)
            Return x * a1
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Dim a1 As Double = GetDoubleValue(parameter, A)
            Dim y As Double = GetDoubleValue(value, 0)

            Try
                Return y / a1
            Catch ex As Exception
                Throw
            End Try

        End Function

        Private Function GetDoubleValue(ByVal parameter As Object, ByVal defaultValue As Double) As Double
            Dim a As Double
            If parameter IsNot Nothing Then
                Try
                    a = System.Convert.ToDouble(parameter)
                Catch
                    a = defaultValue
                End Try
            Else
                a = defaultValue
            End If

            Return a
        End Function

    End Class
End Namespace
