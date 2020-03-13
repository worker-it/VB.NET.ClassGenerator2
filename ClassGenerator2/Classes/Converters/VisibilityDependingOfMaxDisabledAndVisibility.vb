Imports System.Windows
Imports System.Windows.Data

Namespace WPFConverters

    Public Class VisibilityDependingOfMaxDisabledAndVisibility
        Implements IMultiValueConverter

        Public Property Collapsed As Boolean = False

        Public Function IValueConverter_Convert(ByVal value() As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IMultiValueConverter.Convert


            Try

                If (value.Count > 2) Then
                    Throw New ArgumentException("Too many arguments for this converter.")
                ElseIf (value.Count < 2) Then
                    Throw New ArgumentException("Too few argument for this converter.")
                Else
                    If (CType(value(0), Visibility) = Visibility.Visible) Then

                        If (Collapsed) Then
                            Return Visibility.Collapsed
                        Else
                            Return Visibility.Hidden
                        End If

                    Else
                        If (CType(value(1), Visibility) = Visibility.Visible) Then
                            Return Visibility.Visible
                        Else

                            If (Collapsed) Then
                                Return Visibility.Collapsed
                            Else
                                Return Visibility.Hidden
                            End If

                        End If
                    End If
                End If

            Catch ex As Exception
                Return Visibility.Visible
            End Try

        End Function

        Public Function IValueConverter_ConvertBack(ByVal value As Object, ByVal targetType As Type(), ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack

            Try

                ReDim Preserve targetType(2)
                targetType(0) = GetType(Visibility)
                targetType(1) = GetType(Visibility)

                Dim result(2) As Object

                If (CType(value, Visibility) = Visibility.Visible) Then
                    result(0) = Visibility.Hidden
                    result(1) = Visibility.Visible
                Else
                    result(0) = Visibility.Visible
                    result(1) = Visibility.Hidden
                End If

                Return result
            Catch ex As Exception
                Return {Visibility.Hidden, Visibility.Visible}
            End Try

        End Function

    End Class
End Namespace
