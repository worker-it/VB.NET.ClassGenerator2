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



#End Region

Namespace Graphical.Modifications
    Public Class HSLColor
        Implements IDisposable

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

        Private _h As Integer
        Private _s As Single
        Private _l As Single

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

            Me.New(0, 0, 1)

        End Sub

        Public Sub New(ByVal h As Integer, ByVal s As Single, ByVal l As Single)

            ' This call is required by the designer.
            'InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.

            _h = h
            _s = s
            _l = l

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

        Public Property H() As Integer
            Get
                Return Me._h
            End Get
            Set(value As Integer)
                Me._h = value
            End Set
        End Property

        Public Property S() As Single
            Get
                Return Me._s
            End Get
            Set(value As Single)
                Me._s = value
            End Set
        End Property

        Public Property L() As Single
            Get
                Return Me._l
            End Get
            Set(value As Single)
                Me._l = value
            End Set
        End Property

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

        Private Function HueToRGB(v1 As Single, v2 As Single, vH As Single) As Single
            If vH < 0 Then
                vH += 1
            End If

            If vH > 1 Then
                vH -= 1
            End If

            If (6 * vH) < 1 Then
                Return (v1 + (v2 - v1) * 6 * vH)
            End If

            If (2 * vH) < 1 Then
                Return v2
            End If

            If (3 * vH) < 2 Then
                Return (v1 + (v2 - v1) * ((2.0F / 3) - vH) * 6)
            End If

            Return v1
        End Function

        '------- --------
        '------- --------
        'Section publique
        '------- --------
        '------- --------

        Public Overrides Function Equals(ByVal hsl As Object) As Boolean

            Dim _HSL As HSLColor = TryCast(hsl, HSLColor)

            If (_HSL Is Nothing) Then
                Return False
            Else
                Return (Me.H = _HSL.H) AndAlso (Me.S = _HSL.S) AndAlso (Me.L = _HSL.L)
            End If

        End Function

        Public Function HSLToRGB() As RGBColor
            Dim r As Byte = 0
            Dim g As Byte = 0
            Dim b As Byte = 0

            If Me.S = 0 Then
                r = CByte(Math.Truncate(Me.L * 255))
                g = CByte(Math.Truncate(Me.L * 255))
                b = CByte(Math.Truncate(Me.L * 255))
            Else
                Dim v1 As Single, v2 As Single
                Dim hue As Single = CSng(Me.H) / 360

                v2 = If((Me.L < 0.5), (Me.L * (1 + Me.S)), ((Me.L + Me.S) - (Me.L * Me.S)))
                v1 = 2 * Me.L - v2

                r = CByte(Math.Truncate(255 * HueToRGB(v1, v2, hue + (1.0F / 3))))
                g = CByte(Math.Truncate(255 * HueToRGB(v1, v2, hue)))
                b = CByte(Math.Truncate(255 * HueToRGB(v1, v2, hue - (1.0F / 3))))
            End If

            Return New RGBColor(r, g, b)
        End Function

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

#End Region

    End Class
End Namespace