Imports System.Globalization
Imports System.Windows.Data
Imports System.Windows.Media

Namespace WPFConverters

    Public Class transformSolidBrushToColor
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
            Dim aSolidBrush As SolidColorBrush = TryCast(value, SolidColorBrush)
            If (Not aSolidBrush Is Nothing) Then
                Return aSolidBrush.Color
            Else
                Return Nothing
            End If

        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace
