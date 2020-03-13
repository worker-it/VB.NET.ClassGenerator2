Imports System.Globalization
Imports System.Windows
Imports System.Windows.Data

Namespace WPFConverters

    Public Class MultipleBoolean
        Implements IMultiValueConverter

        Public Property OnlyAnd As Boolean = False
        Public Property OperatorsList As String()

        Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert

            Dim result As Boolean

            If (OperatorsList IsNot Nothing) AndAlso (OperatorsList.Count > 0) Then

                For i As Integer = 0 To values.Length - 1

                    If (values(i) Is DependencyProperty.UnsetValue) Then
                        values(i) = False
                    End If

                    If (i = 0) Then
                        result = values(i)
                    ElseIf (values(i).GetType = result.GetType) Then
                        Select Case OperatorsList(i - 1)
                            Case "And"
                                result = result And CBool(values(i))
                            Case "Or"
                                result = result Or CBool(values(i))
                            Case "Xor"
                                result = result Xor CBool(values(i))
                            Case Else
                                result = result And CBool(values(i))
                        End Select
                    End If
                Next

                Return result
            Else

                For i As Integer = 0 To values.Length - 1

                    If (values(i) Is DependencyProperty.UnsetValue) Then
                        values(i) = False
                    End If

                    If (i = 0) Then
                        result = values(i)
                    ElseIf (values(i).GetType = result.GetType) Then
                        If (OnlyAnd) Then
                            result = result And CBool(values(i))
                        Else
                            result = result Or CBool(values(i))
                        End If
                    End If
                Next

                Return result

            End If

        End Function

        Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace
