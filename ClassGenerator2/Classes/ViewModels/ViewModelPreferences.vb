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

Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports MahApps.Metro

#End Region

Public Class ViewModelPreferences
    Implements IDisposable, INotifyPropertyChanged

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

    Private m_AccentColors As List(Of AccentColorMenuData)
    Private m_AppThemes As List(Of AppThemeMenuData)

    Private TmdSelectedTheme As AppThemeMenuData
    Private AcmdSelectedAccent As AccentColorMenuData

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

        ' create accent color menu items for the demo
        Me.AccentColors = ThemeManager.Accents.[Select](Function(a)
                                                            Return New AccentColorMenuData() With {
                                                                .Name = a.Name,
                                                                .ColorBrush = TryCast(a.Resources("AccentColorBrush"), Brush)
                                                            }
                                                        End Function).ToList()

        ' create metro theme color menu items for the demo
        Me.AppThemes = ThemeManager.AppThemes.[Select](Function(a)
                                                           Return New AppThemeMenuData() With {
                                                                .Name = a.Name,
                                                                .BorderColorBrush = TryCast(a.Resources("BlackColorBrush"), Brush),
                                                                .ColorBrush = TryCast(a.Resources("WhiteColorBrush"), Brush)
                                                           }
                                                       End Function).ToList()

        Dim ApplicationTheme As Tuple(Of AppTheme, Accent) = ThemeManager.DetectAppStyle(Application.Current)

        For Each at As AppThemeMenuData In Me.AppThemes
            If (at.Name = ApplicationTheme.Item1.Name) Then
                TmdSelectedTheme = at
                Exit For
            End If
        Next

        For Each ac As AccentColorMenuData In Me.AccentColors
            If (ac.Name = ApplicationTheme.Item2.Name) Then
                AcmdSelectedAccent = ac
                Exit For
            End If
        Next

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

    Public Property SelectedTheme As AppThemeMenuData
        Get
            Return TmdSelectedTheme
        End Get
        Set(value As AppThemeMenuData)
            If (value IsNot Nothing) Then
                TmdSelectedTheme = value
                TmdSelectedTheme.DoChangeTheme(Nothing)
                My.Settings.THEME_BASE = TmdSelectedTheme.Name
                My.Settings.Save()
                OnPropertyChanged()
            End If
        End Set
    End Property

    Public Property SelectedAccent As AccentColorMenuData
        Get
            Return AcmdSelectedAccent
        End Get
        Set(value As AccentColorMenuData)
            If (value IsNot Nothing) Then
                AcmdSelectedAccent = value
                AcmdSelectedAccent.DoChangeTheme(Nothing)
                My.Settings.THEME_ACCENT = AcmdSelectedAccent.Name
                My.Settings.Save()
                OnPropertyChanged()
            End If
        End Set
    End Property

    Public Property AccentColors() As List(Of AccentColorMenuData)
        Get
            Return m_AccentColors
        End Get
        Set
            m_AccentColors = Value
        End Set
    End Property

    Public Property AppThemes() As List(Of AppThemeMenuData)
        Get
            Return m_AppThemes
        End Get
        Set
            m_AppThemes = Value
        End Set
    End Property

    Public Property NomPromgrammeur As String
        Get
            Return My.Settings.NOM_PROGRAMMEUR
        End Get
        Set(value As String)
            My.Settings.NOM_PROGRAMMEUR = value
            OnPropertyChanged()
        End Set
    End Property

#End Region

    '************************************************************************************
    '                           P  R  O  C  É  D  U  R  E  S                            *
    '************************************************************************************
#Region "Procdures"

    '------- ------
    '------- ------
    'Section privée
    '------- ------
    '------- ------

    Protected Friend Sub OnPropertyChanged(<CallerMemberName()> Optional ByVal propertyName As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
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
    '                I N T E R F A C E S  I M P L  E M E N T A T I O N S                *
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

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

#End Region

End Class
