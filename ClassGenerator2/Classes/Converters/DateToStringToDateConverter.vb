Imports System.Globalization
Imports System.Windows.Data

Namespace WPFConverters

    Public Class DateToStringToDateConverter
        Implements IValueConverter

        Private Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert

            Dim strValue As String = System.Convert.ToString(value)
            Dim resultDateTime As DateTime

            If (DateTime.TryParse(strValue, resultDateTime)) Then
                Return resultDateTime
            End If

            Return value

        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack

            Dim dtValue As DateTime = System.Convert.ToDateTime(value)

            Return dtValue.ToShortDateString

        End Function

    End Class

End Namespace