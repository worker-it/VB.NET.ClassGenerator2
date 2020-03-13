Imports System.Windows.Data

Namespace WPFConverters

    Public Class FirstDegreeFunctionConverter
        Implements IValueConverter

        Public Property A As Double

        Public Property B As Double

        Public Function IValueConverter_Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
            Dim a1 As Double = GetDoubleValue(parameter, A)
            Dim b1 As Double = GetDoubleValue(parameter, B)
            Dim x As Double = GetDoubleValue(value, 0)
            Return (a1 * x) + b1
        End Function

        Public Function IValueConverter_ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
            Dim a1 As Double = GetDoubleValue(parameter, A)
            Dim b1 As Double = GetDoubleValue(parameter, B)
            Dim y As Double = GetDoubleValue(value, 0)
            Return (y - b1) / a1
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
