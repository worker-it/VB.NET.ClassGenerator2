'************************************************************************************
'* DÉVELOPPÉ PAR :  Jean-Claude Frigon -> 0087378                                   *
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

Imports MahApps.Metro

Public Class AccentColorMenuData
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

    Private RcchangeAccentCommand As ICommand

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

    Public Property Name() As String
    Public Property BorderColorBrush() As Brush
    Public Property ColorBrush() As Brush

    Public ReadOnly Property ChangeAccentCommand() As ICommand
        Get
            If (RcchangeAccentCommand Is Nothing) Then
                RcchangeAccentCommand = New RelayCommand(AddressOf DoChangeTheme, Function()
                                                                                      Return True
                                                                                  End Function)
            End If
            Return RcchangeAccentCommand
        End Get
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

    Protected Friend Overridable Sub DoChangeTheme(sender As Object)
        Dim theme As Tuple(Of AppTheme, Accent) = ThemeManager.DetectAppStyle(Application.Current)
        Dim accent As MahApps.Metro.Accent = ThemeManager.GetAccent(Me.Name)
        ThemeManager.ChangeAppStyle(Application.Current, accent, theme.Item1)
    End Sub

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

#End Region

End Class
