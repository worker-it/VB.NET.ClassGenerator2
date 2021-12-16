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
Imports ControlzEx.Theming
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

    Private WdwAssociated As Preferences

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

    Public Sub New(ByRef _WdwAssociated As Preferences)

        ' This call is required by the designer.
        'InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        ' create accent color menu items for the demo
        Me.m_AccentColors = ThemeManager.Current.Themes.GroupBy(Function(x) x.ColorScheme).OrderBy(Function(x) x.Key).Select(Function(x) New AccentColorMenuData() With {.Name = x.Key, .ColorBrush = x.First().ShowcaseBrush, .BorderColorBrush = CType(x.First().Resources("MahApps.Brushes.ThemeBackground"), Brush)}).ToList()

        ' create metro theme color menu items for the demo
        Me.m_AppThemes = ThemeManager.Current.Themes.GroupBy(Function(x) x.BaseColorScheme).OrderBy(Function(x) x.Key).Select(Function(x) New AppThemeMenuData() With {.Name = x.Key, .BorderColorBrush = CType(x.First().Resources("MahApps.Brushes.ThemeForeground"), Brush), .ColorBrush = CType(x.First().Resources("MahApps.Brushes.ThemeBackground"), Brush)}).ToList()


        TmdSelectedTheme = m_AppThemes.Where(Function(x) x.Name = My.Settings.THEME_BASE).FirstOrDefault()
        AcmdSelectedAccent = m_AccentColors.Where(Function(x) x.Name = My.Settings.THEME_ACCENT).FirstOrDefault()

        WdwAssociated = _WdwAssociated

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

                ThemeManager.Current.ChangeTheme(WdwAssociated, value.Name & "." & SelectedAccent.Name)
                ThemeManager.Current.ChangeTheme(Application.Current.MainWindow, value.Name & "." & SelectedAccent.Name)
                ThemeManager.Current.ThemeSyncMode = ThemeSyncMode.SyncAll
                ThemeManager.Current.SyncTheme()

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

                ThemeManager.Current.ChangeTheme(WdwAssociated, SelectedTheme.Name & "." & value.Name)
                ThemeManager.Current.ChangeTheme(Application.Current.MainWindow, SelectedTheme.Name & "." & value.Name)
                ThemeManager.Current.ThemeSyncMode = ThemeSyncMode.SyncAll
                ThemeManager.Current.SyncTheme()

                My.Settings.THEME_ACCENT = AcmdSelectedAccent.Name
                My.Settings.Save()
                OnPropertyChanged()
            End If
        End Set
    End Property

    Public ReadOnly Property AccentColors() As List(Of AccentColorMenuData)
        Get
            Return m_AccentColors
        End Get
    End Property

    Public ReadOnly Property AppThemes() As List(Of AppThemeMenuData)
        Get
            Return m_AppThemes
        End Get
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
