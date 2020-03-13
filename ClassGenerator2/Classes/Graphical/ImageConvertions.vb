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

Imports System.Drawing
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

#End Region


Namespace Graphical
    Public Class ImageConvertions
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

        Public Shared Function BitmapImage2Bitmap(bitmapImage As BitmapImage) As Bitmap
            ' BitmapImage bitmapImage = new BitmapImage(new Uri("../Images/test.png", UriKind.Relative));

            Using outStream As New IO.MemoryStream()

                Dim enc As BitmapEncoder = New BmpBitmapEncoder()

                enc.Frames.Add(BitmapFrame.Create(bitmapImage))
                enc.Save(outStream)

                Dim bitmap As New System.Drawing.Bitmap(outStream)

                Return New Bitmap(bitmap)

            End Using
        End Function

        <System.Runtime.InteropServices.DllImport("gdi32.dll")>
        Public Shared Function DeleteObject(hObject As IntPtr) As Boolean
        End Function

        Public Shared Function Bitmap2BitmapImage(bmp As Bitmap) As BitmapImage

            Dim ms As New IO.MemoryStream
            Dim bi As New BitmapImage
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
            bi.BeginInit()
            bi.StreamSource = ms
            bi.EndInit()
            Return bi

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
