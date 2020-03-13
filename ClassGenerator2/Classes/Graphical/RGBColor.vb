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

Imports System.Windows.Media

Namespace Graphical.Modifications
    Public Class RGBColor
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

        Private red As Byte
        Private green As Byte
        Private blue As Byte
        Private Alpha As Byte

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

            Me.New(255, 255, 255)

        End Sub

        Public Sub New(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255)

            ' This call is required by the designer.
            'InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.
            red = R
            green = G
            blue = B
            Alpha = A

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

        Public Property R() As Byte
            Get
                Return Me.red
            End Get
            Set(value As Byte)
                Me.red = value
            End Set
        End Property

        Public Property G() As Byte
            Get
                Return Me.green
            End Get
            Set(value As Byte)
                Me.green = value
            End Set
        End Property

        Public Property B() As Byte
            Get
                Return Me.blue
            End Get
            Set(value As Byte)
                Me.blue = value
            End Set
        End Property

        Public Property A() As Byte
            Get
                Return Me.Alpha
            End Get
            Set(value As Byte)
                Me.Alpha = value
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

        Private Function DecimalToHexadecimal(ByVal dec As Integer) As String
            If dec <= 0 Then
                Return "00"
            End If

            Dim hex As Integer = dec
            Dim hexStr As String = String.Empty

            While dec > 0
                hex = dec Mod 16

                If hex < 10 Then
                    hexStr = hexStr.Insert(0, Convert.ToChar(hex + 48).ToString())
                Else
                    hexStr = hexStr.Insert(0, Convert.ToChar(hex + 55).ToString())
                End If

                dec \= 16
            End While

            Return hexStr
        End Function

        '------- --------
        '------- --------
        'Section publique
        '------- --------
        '------- --------

        Public Shadows Function Equals(obj As Object) As Boolean

            Dim RGB As RGBColor = TryCast(obj, RGBColor)

            If (RGB Is Nothing) Then
                Return False
            Else
                Return (Me.R = RGB.R) AndAlso (Me.G = RGB.G) AndAlso (Me.B = RGB.B) AndAlso (Me.A = RGB.A)
            End If

        End Function

        Public Function RGBToHSL() As HSLColor
            Dim hsl As New HSLColor()

            Dim r As Single = (Me.R / 255.0F)
            Dim g As Single = (Me.G / 255.0F)
            Dim b As Single = (Me.B / 255.0F)

            Dim min As Single = Math.Min(Math.Min(r, g), b)
            Dim max As Single = Math.Max(Math.Max(r, g), b)
            Dim delta As Single = max - min

            hsl.L = (max + min) / 2

            If delta = 0 Then
                hsl.H = 0
                hsl.S = 0.0F
            Else
                hsl.S = If((hsl.L <= 0.5), (delta / (max + min)), (delta / (2 - max - min)))

                Dim hue As Single

                If r = max Then
                    hue = ((g - b) / 6) / delta
                ElseIf g = max Then
                    hue = (1.0F / 3) + ((b - r) / 6) / delta
                Else
                    hue = (2.0F / 3) + ((r - g) / 6) / delta
                End If

                If hue < 0 Then
                    hue += 1
                End If
                If hue > 1 Then
                    hue -= 1
                End If

                hsl.H = CInt(Math.Truncate(hue * 360))
            End If

            Return hsl
        End Function

        Public Function RGBToHexadecimal() As String
            Dim rs As String = DecimalToHexadecimal(Me.R)
            Dim gs As String = DecimalToHexadecimal(Me.G)
            Dim bs As String = DecimalToHexadecimal(Me.B)

            Return "#"c & rs & gs & bs
        End Function

        Public Function modifyBrightness(ByVal addRemovePercent As Single) As RGBColor

            Dim HSL As HSLColor = Me.RGBToHSL()
            Dim tempL As Single = HSL.L + addRemovePercent


            If (tempL > 1) Then
                HSL.L = 1
            ElseIf (tempL < -1) Then
                HSL.L = 0
            Else
                HSL.L = tempL
            End If

            Return HSL.HSLToRGB()

        End Function

        Public Function RGBToMediaColor() As Color

            Dim c As New Color()

            c.A = Me.A
            c.R = Me.R
            c.G = Me.G
            c.B = Me.B

            Return c

        End Function

        Public Function RGBToSolidColorBrush() As SolidColorBrush

            Dim c As New Color()

            c.A = Me.A
            c.R = Me.R
            c.G = Me.G
            c.B = Me.B

            Return New SolidColorBrush(c)

        End Function

        Public Shared Function FromMediaColor(ByVal c As Color) As RGBColor
            Return New RGBColor(c.R, c.G, c.B, c.A)
        End Function

        Public Shared Function FromMediaBrush(ByVal c As Brush) As RGBColor
            Dim scb As SolidColorBrush = TryCast(c, SolidColorBrush)
            If (scb Is Nothing) Then
                Return Nothing
            Else
                Dim couleur As Color = scb.Color
                Return New RGBColor(couleur.R, couleur.G, couleur.B, couleur.A)
            End If
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