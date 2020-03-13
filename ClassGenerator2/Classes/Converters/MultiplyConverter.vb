Imports System.Globalization
Imports System.Windows.Data

Namespace WPFConverters

    Public Class MultiplyConverter
        Implements IMultiValueConverter

        Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert

            Dim result As Double = 1.0

            For i As Integer = 0 To values.Length - 1
                If (values(i).GetType = result.GetType) Then
                    result *= CType(values(i), Double)
                End If
            Next

            Return result

        End Function

        Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace
