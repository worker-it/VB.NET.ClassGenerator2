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

Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports ClassGenerator2.Debugging
Imports ControlzEx.Theming
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports System.Globalization
Imports System.Threading
Imports System.Windows.Markup


#End Region

Public Class ViewModelMainWindow
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

    Private LstListeDesConnections As New ObservableCollection(Of ConnectionInfos)(ConnectionInfos.getAllConnectionInfos(Me))
    Private CiSelectedConnection As ConnectionInfos

    Private StrConnectionName As String

    Private StrWindowTitle As String = "VB.Net Class Generator"

    Private RcAbout As RelayCommand
    Private RcSettings As RelayCommand
    Private RcRetrieveDB As RelayCommand
    Private RcRetrieveDBInfos As RelayCommand
    Private RcCreateFiles As RelayCommand

    Private strAnEventLogger As New EventLogger("C:\temp", "ClassGenerator.log", "MainWindow", ".", "MainWindow", True)

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
        strAnEventLogger.writeLog("Voilà, c'est un bon début!", "", EventLogEntryType.Information)

        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        '@    D E B U T P R O G R A M M E     @
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

        '----------
        'Force le format "ENGLISH - US" pour date, nombres, etc
        Dim newCulture As CultureInfo = New CultureInfo("en-US")
        CultureInfo.DefaultThreadCurrentCulture = newCulture
        CultureInfo.DefaultThreadCurrentUICulture = newCulture
        Thread.CurrentThread.CurrentCulture = newCulture

        Dim lang As XmlLanguage = XmlLanguage.GetLanguage(CultureInfo.CurrentCulture.IetfLanguageTag)
        FrameworkElement.LanguageProperty.OverrideMetadata(GetType(FrameworkElement), New FrameworkPropertyMetadata(lang))
        FrameworkContentElement.LanguageProperty.OverrideMetadata(GetType(System.Windows.Documents.TextElement), New FrameworkPropertyMetadata(lang))

        'Gestion du thème - initialisation
        ThemeManager.Current.ThemeSyncMode = ThemeSyncMode.SyncAll
        ThemeManager.Current.SyncTheme()

        ThemeManager.Current.ChangeThemeBaseColor(Application.Current, My.Settings.THEME_BASE)
        ThemeManager.Current.ChangeThemeColorScheme(Application.Current, My.Settings.THEME_ACCENT)

        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        '@ F I N  D E B U T  P R O G R A M M E@
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        '@     V E R I F.  V E R S I O N      @
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

        Try
            UpdateApplication.VersionVerification()
        Catch ex As Exception

        Finally

        End Try

        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        '@  F I N  V E R I F.  V E R S I O N  @
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


        LstListeDesConnections.Insert(0, New ConnectionInfos(Me))
        ConnectionSelectionnee = LstListeDesConnections(0)
        SelectedConnectionName = ConnectionSelectionnee.ConnectionName
        OnPropertyChanged("SelectedConnectionName")

        strAnEventLogger.writeLog("Bon, on a passé la première étape!", "", EventLogEntryType.Information)

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

        LstListeDesConnections = Nothing

        Try
            RemoveHandler CiSelectedConnection.PropertyChanged, AddressOf ConnectionInfosPropertyChanged
        Catch ex As Exception

        End Try

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

    Public Property WindowTitle As String
        Get
            strAnEventLogger.writeLog("On a lu le titre!", "", EventLogEntryType.Information)
            Return StrWindowTitle & " - V" & My.Application.Info.Version.ToString
        End Get
        Set(value As String)
            StrWindowTitle = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property ConnectionSelectionnee As ConnectionInfos
        Get
            strAnEventLogger.writeLog(CiSelectedConnection.ToString(), "", EventLogEntryType.Information)
            Return CiSelectedConnection
        End Get
        Set(value As ConnectionInfos)

            strAnEventLogger.writeLog("Propriété ConnectionSelectionnee!", "", EventLogEntryType.Information)

            If (value IsNot Nothing) Then

                CiSelectedConnection = value

            Else

                CiSelectedConnection = New ConnectionInfos(Me)
                AddHandler CiSelectedConnection.PropertyChanged, AddressOf ConnectionInfosPropertyChanged

            End If



            RcRetrieveDB = New RelayCommand(AddressOf CiSelectedConnection.BtnRetrieveDbsCommand, Function()
                                                                                                      Return True
                                                                                                  End Function)
            RcRetrieveDBInfos = New RelayCommand(AddressOf CiSelectedConnection.RetrieveDBInfosCommand, Function()
                                                                                                            Return True
                                                                                                        End Function)
            RcCreateFiles = New RelayCommand(AddressOf CiSelectedConnection.CreateFilesCommand, Function()
                                                                                                    Return True
                                                                                                End Function)

            OnPropertyChanged("RetrieveDBs")
            OnPropertyChanged("RetrieveDBInfos")
            OnPropertyChanged("CreateFiles")
            OnPropertyChanged("AdresseNomServeurOuNomInstance")
            OnPropertyChanged()
        End Set
    End Property

    Public Property SelectedConnectionName As String
        Get
            strAnEventLogger.writeLog(CiSelectedConnection.ConnectionName, "", EventLogEntryType.Information)
            Return CiSelectedConnection.ConnectionName
        End Get
        Set(value As String)
            strAnEventLogger.writeLog("Remplace " & CiSelectedConnection.ConnectionName & " par " & value & ".", "", EventLogEntryType.Information)
            CiSelectedConnection.ConnectionName = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property ListeConnections As ObservableCollection(Of ConnectionInfos)
        Get
            strAnEventLogger.writeLog("Lecture de la liste des connections", "", EventLogEntryType.Information)
            Return LstListeDesConnections
        End Get
        Set(value As ObservableCollection(Of ConnectionInfos))
            strAnEventLogger.writeLog("Changement de la liste des connections", "", EventLogEntryType.Information)
            LstListeDesConnections = value
            OnPropertyChanged()
        End Set
    End Property

    Public ReadOnly Property BrowseButtonVisibility As Visibility
        Get
            Select Case CiSelectedConnection.TypeBaseDonnees
                Case ConnectionInfos.databaseType.FLAT_FILE, ConnectionInfos.databaseType.MS_ACCESS_2007_2019, ConnectionInfos.databaseType.MS_ACCESS_97_2003, ConnectionInfos.databaseType.MS_EXCEL
                    Return Visibility.Visible
                Case Else
                    Return Visibility.Hidden
            End Select
        End Get
    End Property

    Public ReadOnly Property TCPPortVisibility As Visibility
        Get
            Select Case CiSelectedConnection.TypeBaseDonnees
                Case ConnectionInfos.databaseType.FLAT_FILE, ConnectionInfos.databaseType.MS_ACCESS_2007_2019, ConnectionInfos.databaseType.MS_ACCESS_97_2003, ConnectionInfos.databaseType.MS_EXCEL, ConnectionInfos.databaseType.NONE
                    Return Visibility.Hidden
                Case Else
                    Return Visibility.Visible
            End Select
        End Get
    End Property

    Public ReadOnly Property AdresseNomServeurOuNomInstance As String
        Get
            Select Case CiSelectedConnection.TypeBaseDonnees
                Case ConnectionInfos.databaseType.FLAT_FILE, ConnectionInfos.databaseType.MS_ACCESS_2007_2019, ConnectionInfos.databaseType.MS_ACCESS_97_2003, ConnectionInfos.databaseType.MS_EXCEL
                    Return "Nom du fichier :"
                Case ConnectionInfos.databaseType.MYSQL, ConnectionInfos.databaseType.ORACLE, ConnectionInfos.databaseType.POSTGRE_SQL, ConnectionInfos.databaseType.SQL_SERVER
                    Return "Nom ou Adresse IP du serveur :"
                Case ConnectionInfos.databaseType.SQL_SERVER_LOCALDB
                    Return "Nom de l'instance :"
                Case Else
                    Return ""
            End Select
        End Get
    End Property

    Public Property ShowAboutDlg As ICommand
        Get
            If (RcAbout Is Nothing) Then
                RcAbout = New RelayCommand(AddressOf CommandAbout, Function()
                                                                       Return True
                                                                   End Function)
            End If
            Return RcAbout
        End Get
        Set(value As ICommand)
            RcAbout = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property ShowPreferences As ICommand
        Get
            If (RcSettings Is Nothing) Then
                RcSettings = New RelayCommand(AddressOf CommandSettings, Function()
                                                                             Return True
                                                                         End Function)
            End If
            Return RcSettings
        End Get
        Set(value As ICommand)
            RcSettings = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property RetrieveDBs As ICommand
        Get
            RcRetrieveDB = New RelayCommand(AddressOf CiSelectedConnection.BtnRetrieveDbsCommand, Function()
                                                                                                      Return True
                                                                                                  End Function)
            Return RcRetrieveDB
        End Get
        Set(value As ICommand)
            RcRetrieveDB = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property RetrieveDBInfos As ICommand
        Get
            RcRetrieveDBInfos = New RelayCommand(AddressOf CiSelectedConnection.RetrieveDBInfosCommand, Function()
                                                                                                            Return True
                                                                                                        End Function)
            Return RcRetrieveDBInfos
        End Get
        Set(value As ICommand)
            RcRetrieveDBInfos = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property CreateFiles As ICommand
        Get
            RcCreateFiles = New RelayCommand(AddressOf CiSelectedConnection.CreateFilesCommand, Function()
                                                                                                    Return True
                                                                                                End Function)
            Return RcCreateFiles
        End Get
        Set(value As ICommand)
            RcCreateFiles = value
            OnPropertyChanged()
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

    Private Sub ConnectionInfosPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)

        Select Case e.PropertyName
            Case "TypeBaseDonnees"
                OnPropertyChanged("AdresseNomServeurOuNomInstance")
            Case Else

        End Select

    End Sub

    Protected Friend Sub OnPropertyChanged(<CallerMemberName()> Optional ByVal propertyName As String = Nothing)
        'MsgBox(propertyName & vbCrLf & ViewModelPeriodeDePaie.instance & vbCrLf & "Lecture :" & Employes.LectureDePropriete & vbCrLf & "Ecriture :" & Employes.EcriturePropriete & vbCrLf & "Fonction :" & ViewModelPeriodeDePaie.AppelDeFonction)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Friend Sub CommandAbout(ByVal params As Object)



    End Sub

    Protected Friend Sub CommandSettings(ByVal params As Object)
        strAnEventLogger.writeLog("On veut voir les préférences!", "", EventLogEntryType.Information)
        Dim pref As New Preferences
        pref.DataContext = New ViewModelPreferences(pref)
        pref.ShowDialog()
        pref.Close()
        pref = Nothing
    End Sub

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

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged


#End Region


End Class
