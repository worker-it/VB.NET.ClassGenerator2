Imports System.Globalization
Imports System.Windows.Data

Namespace WPFConverters

    Public Class FieldsEqualsSelectedNumber
        Implements IValueConverter

        Public Property Number As Double

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
            Dim a1 As Double = GetDoubleValue(parameter, Number)
            Dim x As Double = GetDoubleValue(value, 0)
            Return (a1 = x)
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
