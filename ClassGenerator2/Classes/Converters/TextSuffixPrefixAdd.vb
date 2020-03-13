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

#End Region

Namespace WPFConverters

    Public Class TextSuffixPrefixAdd
        Implements IValueConverter

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

        Public Property AddSuffix As String
        Public Property AddPrefix As String

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
            Dim mot As String = TryCast(value, String)

            If (IsNumeric(value)) Then
                mot = CStr(value)
            End If

            If (mot Is Nothing) Then
                Return AddPrefix & AddSuffix
            Else
                Return AddPrefix & mot & AddSuffix
            End If

        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack

            Dim mot As String = TryCast(value, String)

            If (IsNumeric(value)) Then
                mot = CStr(value)
            End If

            If (mot Is Nothing) Then
                Return Nothing
            Else
                Return mot.Replace(AddPrefix, "").Replace(AddSuffix, "")
            End If

        End Function

#End Region

    End Class
End Namespace