Imports System.Globalization
Imports System.Windows.Data
Imports System.Windows.Shapes

Namespace WPFConverters

    Public Class ratioWidthMoinsHeight
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert

            Return GetDoubleValue(value, 0) - GetDoubleValue(TryCast(parameter, Rectangle).ActualHeight, 0)

        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotImplementedException()
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
