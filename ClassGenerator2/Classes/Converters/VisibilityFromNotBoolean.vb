Imports System.Windows
Imports System.Windows.Data
Imports System.Windows.Markup

Namespace WPFConverters

    <MarkupExtensionReturnType(GetType(IValueConverter))>
    <ValueConversion(GetType(Boolean), GetType(Visibility))>
    Public Class VisibilityFromNotBoolean
        Inherits MarkupExtension
        Implements IValueConverter

        Private Shared _Converter As VisibilityFromNotBoolean
        Public Property collapsed As Boolean = False

        Public Function Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert

            Try
                Dim aVisibility As Visibility = CType(CType(value, Boolean), Visibility)
                If CBool(aVisibility) Then

                    If (collapsed) Then
                        Return Visibility.Collapsed
                    Else
                        Return Visibility.Hidden
                    End If

                Else
                    Return Visibility.Visible
                End If
            Catch ex As Exception

                If (collapsed) Then
                    Return Visibility.Collapsed
                Else
                    Return Visibility.Hidden
                End If

            End Try

        End Function

        Public Function ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack

            Try
                Dim aVisibility As Visibility = CType(value, Visibility)
                If (aVisibility = Visibility.Visible) Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                Return False
            End Try

        End Function

        Public Overrides Function ProvideValue(serviceProvider As IServiceProvider) As Object

            If (_Converter Is Nothing) Then
                _Converter = New VisibilityFromNotBoolean()
            End If

            Return _Converter

        End Function
    End Class
End Namespace
