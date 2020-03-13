'************************************************************************************
'* DÉVELOPPÉ PAR :  Jean-Claude Frigon                                               *
'* DATE :           Juin 07                                                         *
'* MODIFIÉ :                                                                        *
'* PAR :                                                                            *
'* DESCRIPTION :                                                                    *
'*      Public [Function | Sub] nomProcFunct( paramètres)                           *
'************************************************************************************

'************************************************************************************
'                                                                                   *
'                           L I B R A R Y  I M P O R T S                            *
'                                                                                   *
'************************************************************************************
#Region "Library Imports"

Imports System.Globalization
Imports System.Windows.Data
Imports System.Windows.Media
Imports ClassGenerator2.Graphical.Modifications


#End Region

Namespace WPFConverters
    Public Class modifyColorBrightness
        Implements IDisposable, IValueConverter

        '************************************************************************************
        '                            V  A  R  I  A  B  L  E  S                              *
        '                        D E C L A R E   F U N C T I O N S                          *
        '                                    T Y P E S                                      *
        '************************************************************************************
#Region "Variables, Declare Functions and Types"

        '------- ------
        '------- ------
        'Section privée
        '------- ------
        '------- ------


        ' Field to handle multiple calls to Dispose gracefully.
        Dim disposed As Boolean = False

        'Classe Variables



        '------- --------
        '------- --------
        'Section publique
        '------- --------
        '------- --------

#End Region
        '************************************************************************************
        '                    C  O  N  S  T  R  U  C  T  E  U  R                             *
        '                    ----------------------------------                             *
        '                      D  E  S  T  R  U  C  T  E  U  R                              *
        '************************************************************************************
#Region "Constructors"

        Public Sub New()

            ' This call is required by the designer.
            'InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.

        End Sub

#End Region

        ' Implement IDisposable.
#Region "IDisposable implementation"
        Public Overloads Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Overloads Sub Dispose(disposing As Boolean)
            If disposed = False Then
                If disposing Then
                    ' Free other state (managed objects).
                    disposed = True
                End If
            End If

            ' Free your own state (unmanaged objects).
            ' Set large fields to null.

        End Sub

        Protected Overrides Sub Finalize()

            ' Simply call Dispose(False).
            Dispose(False)

        End Sub

#End Region

        '************************************************************************************
        '                           P  R  O  P  R  I  É  T  É  S                            *
        '************************************************************************************
#Region "Properties"

        '------- ------
        '------- ------
        'Section privée
        '------- ------
        '------- ------

        '------- --------
        '------- --------
        'Section publique
        '------- --------
        '------- --------

        Public Property BrightnessModifier As Single

#End Region

        '************************************************************************************
        '                           P  R  O  C  É  D  U  R  E  S                            *
        '************************************************************************************
#Region "Procédures"

        '------- ------
        '------- ------
        'Section privée
        '------- ------
        '------- ------

        '------- --------
        '------- --------
        'Section publique
        '------- --------
        '------- --------

#End Region

        '************************************************************************************
        '                             F  O  N  C  T  I  O  N  S                             *
        '************************************************************************************
#Region "Functions"

        '------- ------
        '------- ------
        'Section privée
        '------- ------
        '------- ------

        '------- --------
        '------- --------
        'Section publique
        '------- --------
        '------- --------

#End Region

        '************************************************************************************
        '                                  E V E N T S                                      *
        '************************************************************************************
#Region "Events"

        '------- ------
        '------- ------
        'Section privée
        '------- ------
        '------- ------

        '------- --------
        '------- --------
        'Section publique
        '------- --------
        '------- --------

#End Region

        '************************************************************************************
        '                I N T E R F A C E S  I M P L  E M E N T A T J O N S                *
        '************************************************************************************
#Region "Interfaces implementations"

        '------- ------
        '------- ------
        'Section privée
        '------- ------
        '------- ------

        '------- --------
        '------- --------
        'Section publique
        '------- --------
        '------- --------

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert

            If (value Is Nothing) Then
                Return Nothing
            ElseIf (value.GetType() Is GetType(Color)) Then

                Return RGBColor.FromMediaColor(CType(value, Color)).modifyBrightness(BrightnessModifier).RGBToMediaColor

            ElseIf ((value.GetType() Is GetType(Brush)) Or (value.GetType() Is GetType(SolidColorBrush))) Then

                Return RGBColor.FromMediaBrush(CType(value, Brush)).modifyBrightness(BrightnessModifier).RGBToSolidColorBrush

            Else

                Return Nothing

            End If

        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack

            If (value.GetType() Is GetType(Color)) Then

                Return RGBColor.FromMediaColor(CType(value, Color)).modifyBrightness(-BrightnessModifier).RGBToMediaColor

            ElseIf ((value.GetType() Is GetType(Brush)) Or (value.GetType() Is GetType(SolidColorBrush))) Then

                Return RGBColor.FromMediaBrush(CType(value, Brush)).modifyBrightness(-BrightnessModifier).RGBToSolidColorBrush

            Else

                Return Nothing

            End If

        End Function

#End Region

    End Class
End Namespace