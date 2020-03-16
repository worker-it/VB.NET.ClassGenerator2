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
Imports System.Data
Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Globalization
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms
Imports ClassGenerator2.Debugging
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports MySql.Data.MySqlClient
Imports Npgsql

#End Region

Public Class ViewModelMainWindow
    Implements IDisposable, IDataErrorInfo, INotifyPropertyChanged

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

    'Déclaration de la variable pour la connection      
    Private cnxTables As DbConnection
    'Déclaration de la variable pour la connection      
    Private cnxColonnes As DbConnection
    'Déclaration de la variable pour la connectionstring      
    Private cnxstr As String
    'Déclaration de la variable pour la requête      
    Private sqlTables As String
    'Déclaration de la variable pour la requête      
    Private sqlColonnes As String
    'Déclaration de la variable pour la commande       
    Private cmdTables As DbCommand
    'Déclaration de la variable pour le dataadapter
    Private dtrTables As DbDataReader
    'Déclaration de la variable pour la commande       
    Private cmdColonnes As DbCommand
    'Déclaration de la variable pour le dataadapter
    Private dtrColonnes As DbDataReader

    Private lngTypeBaseDonnees As Long
    Private strServerAddresseName As String
    Private IntTCPPort As Integer
    Private lngDatabaseCatalog As Long = 0
    Private strUsername As String
    Private strPassword As String
    Private booTrustedConnection As Boolean
    Private strSelectedDB As String
    Private lstListeTables As ObservableCollection(Of TreeView.Noeud)
    Private listeDeClasses As List(Of ClassCodeVb)
    Private LstListeDesBaseDeDonnees As New ObservableCollection(Of String)

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

    Public Enum databaseType
        NONE = 0
        SQL_SERVER = 1
        ORACLE = 2
        MYSQL = 3
        POSTGRE_SQL = 4
        MS_ACCESS_97_2003 = 5
        MS_ACCESS_2007_2019 = 6
        MS_EXCEL = 7
        FLAT_FILE = 8

    End Enum



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

        MahApps.Metro.ThemeManager.ChangeAppStyle(Application.Current, MahApps.Metro.ThemeManager.GetAccent(My.Settings.ACCENT_COLOR), MahApps.Metro.ThemeManager.GetAppTheme(My.Settings.APPLICATION_THEME))

        TypeBaseDonnees = 0
        LstListeDesBaseDeDonnees.Add("Charger liste des BDs")
        DatabaseCatalog = 0
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

    Public Property WindowTitle As String
        Get
            Return StrWindowTitle & " - V" & My.Application.Info.Version.ToString
        End Get
        Set(value As String)
            StrWindowTitle = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property TypeBaseDonnees As Long
        Get
            Return lngTypeBaseDonnees
        End Get
        Set(value As Long)
            lngTypeBaseDonnees = value
            Select Case lngTypeBaseDonnees
                Case databaseType.SQL_SERVER
                    Me.IntTCPPort = 1433
                Case databaseType.POSTGRE_SQL
                    Me.IntTCPPort = 5432
                Case databaseType.MYSQL
                    Me.IntTCPPort = 3306
                Case databaseType.ORACLE
                    Me.IntTCPPort = 1521
                Case Else
                    Me.IntTCPPort = -1
            End Select
            DatabaseCatalog = 0
            OnPropertyChanged("TCPPort")
            OnPropertyChanged("BrowseButtonVisibility")
            OnPropertyChanged("TCPPortVisibility")
            OnPropertyChanged()
        End Set
    End Property

    Public Property ServerAddresseName As String
        Get
            Return strServerAddresseName
        End Get
        Set(value As String)
            strServerAddresseName = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property TCPPort As Integer
        Get
            Return IntTCPPort
        End Get
        Set(value As Integer)
            IntTCPPort = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property DatabaseCatalog As Long
        Get
            Return lngDatabaseCatalog
        End Get
        Set(value As Long)
            lngDatabaseCatalog = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Username As String
        Get
            Return strUsername
        End Get
        Set(value As String)
            strUsername = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property TrustedConnection As Boolean
        Get
            Return booTrustedConnection
        End Get
        Set(value As Boolean)
            booTrustedConnection = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property SelectedDB As String
        Get
            Return strSelectedDB
        End Get
        Set(value As String)
            strSelectedDB = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property ListeDesBaseDeDonnees As ObservableCollection(Of String)
        Get
            Return LstListeDesBaseDeDonnees
        End Get
        Set(value As ObservableCollection(Of String))
            LstListeDesBaseDeDonnees = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property ListeTablesEtChamps As ObservableCollection(Of TreeView.Noeud)
        Get
            If (lstListeTables Is Nothing) Then
                lstListeTables = New ObservableCollection(Of TreeView.Noeud)
            End If
            Return lstListeTables
        End Get
        Set(value As ObservableCollection(Of TreeView.Noeud))
            lstListeTables = value
            OnPropertyChanged()
        End Set
    End Property

    Public ReadOnly Property BrowseButtonVisibility As Visibility
        Get
            Select Case TypeBaseDonnees
                Case databaseType.FLAT_FILE, databaseType.MS_ACCESS_2007_2019, databaseType.MS_ACCESS_97_2003, databaseType.MS_EXCEL
                    Return Visibility.Visible
                Case Else
                    Return Visibility.Hidden
            End Select
        End Get
    End Property

    Public ReadOnly Property TCPPortVisibility As Visibility
        Get
            Select Case TypeBaseDonnees
                Case databaseType.FLAT_FILE, databaseType.MS_ACCESS_2007_2019, databaseType.MS_ACCESS_97_2003, databaseType.MS_EXCEL, databaseType.NONE
                    Return Visibility.Hidden
                Case Else
                    Return Visibility.Visible
            End Select
        End Get
    End Property

    Public Property ListeTables As ObservableCollection(Of TreeView.Noeud)
        Get
            Return lstListeTables
        End Get
        Set(value As ObservableCollection(Of TreeView.Noeud))
            lstListeTables = value
            OnPropertyChanged()
        End Set
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
            If (RcRetrieveDB Is Nothing) Then
                RcRetrieveDB = New RelayCommand(AddressOf BtnRetrieveDbsCommand, Function()
                                                                                     Return True
                                                                                 End Function)
            End If
            Return RcRetrieveDB
        End Get
        Set(value As ICommand)
            RcRetrieveDB = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property RetrieveDBInfos As ICommand
        Get
            If (RcRetrieveDBInfos Is Nothing) Then
                RcRetrieveDBInfos = New RelayCommand(AddressOf RetrieveDBInfosCommand, Function()
                                                                                           Return True
                                                                                       End Function)
            End If
            Return RcRetrieveDBInfos
        End Get
        Set(value As ICommand)
            RcRetrieveDBInfos = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property CreateFiles As ICommand
        Get
            If (RcCreateFiles Is Nothing) Then
                RcCreateFiles = New RelayCommand(AddressOf CreateFilesCommand, Function()
                                                                                   Return True
                                                                               End Function)
            End If
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

    Private Sub BtnRetrieveDbsCommand(ByVal params As Object)
        Try

            Dim pwdbox As PasswordBox = TryCast(params, PasswordBox)
            strPassword = pwdbox.Password
            pwdbox = Nothing

            If (strPassword <> String.Empty) Then

                strAnEventLogger.writeLog("Initiation de la recherche des bases de données", "", EventLogEntryType.Information)
                If (Me.ServerAddresseName <> "") And (((Me.Username <> "") And (Me.strPassword <> "")) Or (Me.TrustedConnection)) Then

                    strAnEventLogger.writeLog("Si les infos de bases de données sont bien entrées.", "", EventLogEntryType.Information)

                    Dim connString As String = ""
                    Dim commString As String = ""

                    Dim connection As DbConnection
                    Dim commande As DbCommand
                    Dim reader As DbDataReader

                    strAnEventLogger.writeLog("Vide la collection de bases de données.", "", EventLogEntryType.Information)
                    Me.ListeDesBaseDeDonnees.Clear()

                    Select Case Me.TypeBaseDonnees
                        Case databaseType.SQL_SERVER
                            Try

                                strAnEventLogger.writeLog("MS SQL Server.", "", EventLogEntryType.Information)

                                commString = "SELECT name FROM master.sys.databases;"
                                If (Me.TrustedConnection) Then
                                    strAnEventLogger.writeLog("Connection MS SQL integrated Security.", "", EventLogEntryType.Information)
                                    connection = New OleDbConnection("Provider=sqloledb;Server=" & Me.ServerAddresseName & ";Integrated Security=SSPI")
                                Else
                                    strAnEventLogger.writeLog("Connection MS SQL username and password.", "", EventLogEntryType.Information)
                                    connection = New OleDbConnection("Provider=sqloledb;Server=" & Me.ServerAddresseName & ";User Id=" & Me.strUsername & ";Password=" & Me.strPassword & ";")
                                End If

                                strAnEventLogger.writeLog("Definition de la commande.", "", EventLogEntryType.Information)
                                commande = connection.CreateCommand()
                                commande.CommandText = commString
                                commande.CommandType = System.Data.CommandType.Text

                                strAnEventLogger.writeLog("Exécution de la commande.", "", EventLogEntryType.Information)
                                connection.Open()

                                strAnEventLogger.writeLog("Lecture des résultats.", "", EventLogEntryType.Information)
                                reader = commande.ExecuteReader()

                                Me.LstListeDesBaseDeDonnees.Add("Veuillez sélectionner une BD")

                                While reader.Read()
                                    strAnEventLogger.writeLog("Lecture ligne : " & reader.GetString(0) & ".", "", EventLogEntryType.Information)
                                    Me.LstListeDesBaseDeDonnees.Add(reader.GetString(0))
                                End While

                                strAnEventLogger.writeLog("Fermeture de la connection.", "", EventLogEntryType.Information)
                                reader.Close()
                                connection.Close()

                                reader = Nothing
                                connection = Nothing
                                commande = Nothing

                                lngDatabaseCatalog = 0
                                OnPropertyChanged("DatabaseCatalog")

                            Catch oledbEx As OleDbException
                                strAnEventLogger.writeLog(oledbEx.Message, oledbEx.StackTrace, EventLogEntryType.Error)
                                MsgBox(oledbEx.Message)
                            Catch ex As Exception
                                strAnEventLogger.writeLog(ex.Message, ex.StackTrace, EventLogEntryType.Error)
                                MsgBox(ex.Message)
                            End Try

                        Case databaseType.ORACLE

                            strAnEventLogger.writeLog("Orable Database.", "", EventLogEntryType.Information)

                            commString = "SELECT table_name FROM user_tables;"

                            lngDatabaseCatalog = 0
                            OnPropertyChanged("DatabaseCatalog")

                        Case databaseType.POSTGRE_SQL

                            Try

                                strAnEventLogger.writeLog("PostgreSQL.", "", EventLogEntryType.Information)

                                commString = "SELECT datname FROM pg_database;"

                                strAnEventLogger.writeLog("Connection MS SQL username and password.", "", EventLogEntryType.Information)
                                connection = New NpgsqlConnection("User ID=" & Me.strUsername & ";Password=" & Me.strPassword & ";Server=" & Me.ServerAddresseName & ";Port=" & IntTCPPort)

                                strAnEventLogger.writeLog("Definition de la commande.", "", EventLogEntryType.Information)
                                commande = New NpgsqlCommand(commString, CType(connection, NpgsqlConnection))
                                commande.CommandType = System.Data.CommandType.Text

                                strAnEventLogger.writeLog("Exécution de la commande.", "", EventLogEntryType.Information)
                                connection.Open()

                                strAnEventLogger.writeLog("Lecture des résultats.", "", EventLogEntryType.Information)
                                reader = commande.ExecuteReader()

                                Me.LstListeDesBaseDeDonnees.Add("Veuillez sélectionner une BD")

                                While reader.Read()
                                    strAnEventLogger.writeLog("Lecture ligne : " & reader.GetString(0) & ".", "", EventLogEntryType.Information)
                                    Me.LstListeDesBaseDeDonnees.Add(reader.GetString(0))
                                End While

                                strAnEventLogger.writeLog("Fermeture de la connection.", "", EventLogEntryType.Information)
                                reader.Close()
                                connection.Close()

                                reader = Nothing
                                connection = Nothing
                                commande = Nothing

                                lngDatabaseCatalog = 0
                                OnPropertyChanged("DatabaseCatalog")

                            Catch ex As Exception
                                strAnEventLogger.writeLog(ex.Message, ex.StackTrace, EventLogEntryType.Error)
                                MsgBox(ex.Message)
                            End Try

                        Case databaseType.MYSQL

                            Try

                                strAnEventLogger.writeLog("MySQL.", "", EventLogEntryType.Information)

                                commString = "SHOW DATABASES;"

                                strAnEventLogger.writeLog("Connection MS SQL username and password.", "", EventLogEntryType.Information)
                                connection = New MySqlConnection("Server=" & Me.ServerAddresseName & ";Port=" & Me.IntTCPPort & ";Uid=" & Me.strUsername & ";Pwd=" & Me.strPassword & ";")

                                strAnEventLogger.writeLog("Definition de la commande.", "", EventLogEntryType.Information)
                                commande = New MySqlCommand(commString, CType(connection, MySqlConnection))
                                commande.CommandType = System.Data.CommandType.Text

                                strAnEventLogger.writeLog("Exécution de la commande.", "", EventLogEntryType.Information)
                                connection.Open()

                                strAnEventLogger.writeLog("Lecture des résultats.", "", EventLogEntryType.Information)
                                reader = commande.ExecuteReader()

                                Me.LstListeDesBaseDeDonnees.Add("Veuillez sélectionner une BD")

                                While reader.Read()
                                    strAnEventLogger.writeLog("Lecture ligne : " & reader.GetString(0) & ".", "", EventLogEntryType.Information)
                                    Me.LstListeDesBaseDeDonnees.Add(reader.GetString(0))
                                End While

                                reader.Close()
                                connection.Close()

                                reader = Nothing
                                connection = Nothing
                                commande = Nothing

                                lngDatabaseCatalog = 0
                                OnPropertyChanged("DatabaseCatalog")

                            Catch oledbEx As OleDbException
                                strAnEventLogger.writeLog(oledbEx.Message, oledbEx.StackTrace, EventLogEntryType.Error)
                                MsgBox(oledbEx.Message)
                            Catch ex As Exception
                                strAnEventLogger.writeLog(ex.Message, ex.StackTrace, EventLogEntryType.Error)
                                MsgBox(ex.Message)
                            End Try

                        Case Else

                            strAnEventLogger.writeLog("Autres : Ne devrait pas arriver ou devrait être implémenté.", "", EventLogEntryType.Information)


                    End Select
                End If

            End If

        Catch ex As Exception

            strAnEventLogger.writeLog(ex.Message, ex.StackTrace, EventLogEntryType.Error)

            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub RetrieveDBInfosCommand(ByVal params As Object)

        listeDeClasses = New List(Of ClassCodeVb)

        'ouverture de la connection (à partir du répertoire de l'application) sur la même ligne      
        cnxstr = buildConnectionString()

        Dim principale As New ObservableCollection(Of TreeView.Noeud)
        Dim lst As New ObservableCollection(Of TreeView.Noeud)

        Select Case Me.TypeBaseDonnees
            Case databaseType.SQL_SERVER

                cnxTables = New OleDbConnection
                cnxTables.ConnectionString = cnxstr
                cnxTables.Open()

                'Création de la requête sql      
                sqlTables = "SELECT " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_NAME, " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_SCHEMA " &
                            "FROM " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES " &
                            "WHERE (LEFT(" & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_NAME,3)<>'SYS');"

                'sqlTables = "SELECT " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_NAME, " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_SCHEMA " &
                '            "FROM " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES " &
                '            "WHERE (" & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_TYPE<>'VIEW') AND (LEFT(" & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_NAME,3)<>'SYS');"

                'Création de la commande et on l'instancie (sql)       
                cmdTables = New OleDbCommand(sqlTables, CType(cnxTables, OleDbConnection))

                'Création du datareader (dta)
                dtrTables = cmdTables.ExecuteReader()

                While dtrTables.Read()

                    Dim colonnes As New ObservableCollection(Of TreeView.Noeud)

                    cnxColonnes = New OleDbConnection
                    cnxColonnes.ConnectionString = cnxstr
                    cnxColonnes.Open()

                    'Création de la requête sql      
                    sqlColonnes = "SELECT " & Me.SelectedDB & ".INFORMATION_SCHEMA.COLUMNS.COLUMN_NAME, " &
                                                Me.SelectedDB & ".INFORMATION_SCHEMA.COLUMNS.DATA_TYPE, " &
                                                "CONSTRAINTS_COLUMNS.CONSTRAINT_TYPE " &
                                    "FROM " & Me.SelectedDB & ".INFORMATION_SCHEMA.COLUMNS LEFT JOIN " &
                                        "(" & Me.SelectedDB & ".INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE LEFT JOIN (SELECT * " &
                                                                                                                        "FROM " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLE_CONSTRAINTS " &
                                                                                                                        "WHERE CONSTRAINT_TYPE='PRIMARY KEY') AS CONSTRAINTS_COLUMNS ON " &
                                        "CONSTRAINT_COLUMN_USAGE.TABLE_SCHEMA=CONSTRAINTS_COLUMNS.TABLE_SCHEMA And " &
                                        "CONSTRAINT_COLUMN_USAGE.TABLE_NAME=CONSTRAINTS_COLUMNS.TABLE_NAME And " &
                                        "CONSTRAINT_COLUMN_USAGE.CONSTRAINT_NAME=CONSTRAINTS_COLUMNS.CONSTRAINT_NAME" &
                                        ") ON " &
                                    "COLUMNS.TABLE_SCHEMA=CONSTRAINTS_COLUMNS.TABLE_SCHEMA And " &
                                    "COLUMNS.TABLE_NAME=CONSTRAINTS_COLUMNS.TABLE_NAME And " &
                                    "COLUMNS.COLUMN_NAME=CONSTRAINT_COLUMN_USAGE.COLUMN_NAME " &
                                    "WHERE COLUMNS.TABLE_SCHEMA='" & dtrTables.Item("TABLE_SCHEMA").ToString() & "' AND " &
                                        "COLUMNS.TABLE_NAME='" & dtrTables.Item("TABLE_NAME").ToString() & "';"

                    'Création de la commande et on l'instancie (sql)       
                    cmdColonnes = New OleDbCommand(sqlColonnes, CType(cnxColonnes, OleDbConnection))

                    'Création du datareader (dta)
                    dtrColonnes = cmdColonnes.ExecuteReader()

                    While dtrColonnes.Read()

                        Dim uneColonne As New DbChampTable(dtrColonnes.Item("COLUMN_NAME").ToString, dtrColonnes.Item("DATA_TYPE").ToString)

                        uneColonne.IsPrimaryKey = (dtrColonnes.Item("CONSTRAINT_TYPE").ToString().ToUpper = "PRIMARY KEY")

                        colonnes.Add(uneColonne)
                    End While

                    lst.Add(New DbTable(dtrTables.Item("TABLE_NAME").ToString(), colonnes) With {.tableSchema = dtrTables.Item("TABLE_SCHEMA").ToString()})

                    dtrColonnes.Close()
                    dtrColonnes = Nothing

                    cmdColonnes = Nothing
                    cnxColonnes.Close()
                    cnxColonnes = Nothing

                End While

            Case databaseType.ORACLE
            Case databaseType.MYSQL

                cnxTables = New MySqlConnection()
                cnxTables.ConnectionString = cnxstr
                cnxTables.Open()

                'Création de la requête sql      
                sqlTables = "SHOW TABLES;"

                'Création de la commande et on l'instancie (sql)       
                cmdTables = New MySqlCommand(sqlTables, CType(cnxTables, MySqlConnection))

                'Création du datareader (dta)
                dtrTables = cmdTables.ExecuteReader()

                While dtrTables.Read()

                    Dim colonnes As New ObservableCollection(Of TreeView.Noeud)

                    cnxColonnes = New MySqlConnection
                    cnxColonnes.ConnectionString = cnxstr
                    cnxColonnes.Open()

                    'Création de la requête sql      
                    sqlColonnes = "SHOW COLUMNS FROM " & dtrTables.Item(Me.SelectedDB).ToString()

                    'Création de la commande et on l'instancie (sql)       
                    cmdColonnes = New MySqlCommand(sqlColonnes, CType(cnxColonnes, MySqlConnection))

                    'Création du datareader (dta)
                    dtrColonnes = cmdColonnes.ExecuteReader()

                    While dtrColonnes.Read()

                        Dim uneColonne As New DbChampTable(dtrColonnes.Item("Field").ToString, dtrColonnes.Item("Type").ToString)

                        uneColonne.IsPrimaryKey = (dtrColonnes.Item("Key").ToString().ToUpper = "PRI")

                        colonnes.Add(uneColonne)
                    End While

                    lst.Add(New DbTable(dtrTables.Item(Me.SelectedDB).ToString(), colonnes))

                    dtrColonnes.Close()
                    dtrColonnes = Nothing

                    cmdColonnes = Nothing
                    cnxColonnes.Close()
                    cnxColonnes = Nothing

                End While

            Case databaseType.POSTGRE_SQL

                cnxTables = New NpgsqlConnection()
                cnxTables.ConnectionString = cnxstr
                cnxTables.Open()

                'Création de la requête sql      
                sqlTables = "SELECT schemaname, tablename as object_name " &
                            "FROM """ & Me.SelectedDB & """.""pg_catalog"".""pg_tables"" " &
                            "WHERE ""schemaname"" Not IN ('pg_catalog','information_schema') " &
                            "UNION ALL " &
                            "SELECT schemaname, viewname as object_name " &
                            "FROM """ & Me.SelectedDB & """.""pg_catalog"".""pg_views"" " &
                            "WHERE ""schemaname"" Not IN ('pg_catalog','information_schema');"

                'Création de la commande et on l'instancie (sql)       
                cmdTables = New NpgsqlCommand(sqlTables, CType(cnxTables, NpgsqlConnection))

                'Création du datareader (dta)
                dtrTables = cmdTables.ExecuteReader()

                While dtrTables.Read()

                    Dim colonnes As New ObservableCollection(Of TreeView.Noeud)

                    cnxColonnes = New NpgsqlConnection
                    cnxColonnes.ConnectionString = cnxstr
                    cnxColonnes.Open()

                    'Création de la requête sql  
                    'sqlColonnes = "SELECT current_database();"
                    sqlColonnes = "SELECT ""X"".""table_catalog"", " &
                                            """X"".""table_schema"", " &
                                            """X"".""table_name"", " &
                                            """X"".""column_name"", " &
                                            """X"".""data_type"", " &
                                            """Z"".""constraint_type"" " &
                                    "FROM """ & Me.SelectedDB & """.""information_schema"".""columns"" AS ""X"" " &
                                    "LEFT JOIN (SELECT ""Y"".""table_name"", ""Y"".""column_name"", ""Y"".""table_schema"", ""Y"".""table_catalog"", ""A"".""constraint_type"" " &
                                                "FROM """ & Me.SelectedDB & """.""information_schema"".""constraint_column_usage"" AS ""Y"" " &
                                                    "LEFT JOIN """ & Me.SelectedDB & """.""information_schema"".""table_constraints"" AS ""A"" " &
                                                    "ON ""Y"".""constraint_catalog""=""A"".""constraint_catalog"" And " &
                                                        """Y"".""constraint_schema""=""A"".""constraint_schema"" And " &
                                                        """Y"".""constraint_name""=""A"".""constraint_name"" " &
                                                "WHERE ""A"".""constraint_type""='PRIMARY KEY') AS ""Z"" " &
                                        "ON ""X"".""table_catalog""=""Z"".""table_catalog"" AND " &
                                            """X"".""table_schema""=""Z"".""table_schema"" AND " &
                                            """X"".""table_name""=""Z"".""table_name"" AND " &
                                            """X"".""column_name""=""Z"".""column_name"" " &
                                    "WHERE ""X"".""table_schema"" = '" & dtrTables.Item("schemaname").ToString() & "' AND " &
                                        """X"".""table_name"" = '" & dtrTables.Item("object_name").ToString() & "';"

                    'Création de la commande et on l'instancie (sql)       
                    cmdColonnes = New NpgsqlCommand(sqlColonnes, CType(cnxColonnes, NpgsqlConnection))

                    'Création du datareader (dta)
                    dtrColonnes = cmdColonnes.ExecuteReader()

                    While dtrColonnes.Read()

                        Dim uneColonne As New DbChampTable(dtrColonnes.Item("column_name").ToString, dtrColonnes.Item("data_type").ToString)

                        uneColonne.IsPrimaryKey = (dtrColonnes.Item("constraint_type").ToString().ToUpper = "PRIMARY KEY")

                        colonnes.Add(uneColonne)
                    End While

                    lst.Add(New DbTable(dtrTables.Item("object_name").ToString(), colonnes) With {.tableSchema = dtrTables.Item("schemaname").ToString()})

                    dtrColonnes.Close()
                    dtrColonnes = Nothing

                    cmdColonnes = Nothing
                    cnxColonnes.Close()
                    cnxColonnes = Nothing

                End While

            Case databaseType.MS_ACCESS_97_2003
            Case databaseType.MS_ACCESS_2007_2019

            Case databaseType.MS_EXCEL

            Case databaseType.FLAT_FILE

        End Select

        lst = New ObservableCollection(Of TreeView.Noeud)(lst.OrderBy(Of String)(Function(tbl) CType(tbl, DbTable).Name))

        principale.Add(New DbDatabase(Me.SelectedDB, lst, False) With {.IsExpanded = True})
        ListeTables = principale

        dtrTables.Close()
        dtrTables = Nothing

        cmdTables = Nothing
        cnxTables.Close()
        cnxTables = Nothing

        OnPropertyChanged("ListeTablesEtChamps")
    End Sub

    Public Async Sub CreateFilesCommand(ByVal params As Object)

        Dim wdw As MetroWindow = TryCast(params, MetroWindow)

        If (wdw IsNot Nothing) Then
            Dim folderBrowser As New FolderBrowserDialog()

            If (CreateClasses() = 0) Then
                Await DialogManager.ShowMessageAsync(wdw, "Erreur", "Aucune table ou colonne n'a été sélectionné!", MessageDialogStyle.Affirmative)
            Else
                With folderBrowser

                    .RootFolder = Environment.SpecialFolder.MyComputer
                    If (IO.Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\source\repos")) Then
                        .SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\source\repos"
                    Else
                        .SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer)
                    End If

                    If (.ShowDialog() = DialogResult.OK) Then

                        For Each classe As ClassCodeVb In Me.listeDeClasses

                            classe.CreateVbFile(.SelectedPath)

                        Next

                        listeDeClasses.Clear()
                    End If
                End With
            End If
        Else

        End If

    End Sub

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

    Private Function buildConnectionString() As String

        Dim cnctstr As String

        Select Case Me.TypeBaseDonnees
            Case databaseType.SQL_SERVER
                If (Me.TrustedConnection) Then
                    cnctstr = "Provider=sqloledb;Server=" & Me.ServerAddresseName & ";Database=" & Me.SelectedDB & ";Trusted_Connection=yes;"
                Else
                    cnctstr = "Provider=sqloledb;Server=" & Me.ServerAddresseName & ";Database=" & Me.SelectedDB & ";Uid=" & Me.Username & ";Pwd=" & Me.strPassword & ";"
                End If
            Case databaseType.ORACLE
                If (Me.TrustedConnection) Then
                    cnctstr = "Provider=msdaora;Data Source=" & Me.ServerAddresseName & If(TCPPort.ToString() <> "", ";Port=" & Me.TCPPort.ToString(), "") & ";Persist Security Info=False;Integrated Security=Yes;"
                Else
                    cnctstr = "Provider=msdaora;Data Source=" & Me.ServerAddresseName & If(TCPPort.ToString() <> "", ";Port=" & Me.TCPPort.ToString(), "") & ";User Id=" & Me.Username & ";Password=" & Me.strPassword & ";Integrated Security=no;"
                End If
            Case databaseType.MYSQL
                cnctstr = "Server=" & Me.ServerAddresseName & ";Port=" & Me.IntTCPPort.ToString() & ";Database=" & Me.SelectedDB & ";Uid=" & Me.Username & ";Pwd=" & Me.strPassword & ";"
            Case databaseType.POSTGRE_SQL
                cnctstr = "Server=" & Me.ServerAddresseName & If(TCPPort.ToString() <> "", ";Port=" & Me.TCPPort.ToString(), ";Port=5432") & ";Database=" & Me.SelectedDB & ";User Id=" & Me.Username & ";Password=" & Me.strPassword & ";"
            Case databaseType.MS_ACCESS_97_2003
                cnctstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Me.ServerAddresseName & ";User Id=admin;Password=;"
            Case databaseType.MS_ACCESS_2007_2019
                cnctstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Me.ServerAddresseName & ";Persist Security Info=False;"
            Case databaseType.MS_EXCEL
                'ADO .Net
                cnctstr = "Excel File=" & Me.ServerAddresseName & ";"
            Case databaseType.FLAT_FILE
                cnctstr = Me.ServerAddresseName
            Case Else
                cnctstr = ""
        End Select

        Return cnctstr

    End Function

    Private Function getTypeFromString(ByVal typeDonnee As String) As String

        Dim convertedDataType As String


        Select Case Me.TypeBaseDonnees
            Case databaseType.SQL_SERVER

                Select Case typeDonnee.ToUpper
                    Case "BIGINT"
                        convertedDataType = "Long"
                    Case "INT", "SMALLINT"
                        convertedDataType = "Integer"
                    Case "BINARY", "VARBINARY", "TINYINT"
                        convertedDataType = "Byte"
                    Case "CHAR", "NCHAR"
                        convertedDataType = "Char"
                    Case "DATE", "DATETIME", "DATETIME2", "DATETIMEOFFSET", "SMALLDATETIME", "TIME", "TIMESTAMP"
                        convertedDataType = "Date"
                    Case "FLOAT", "REAL"
                        convertedDataType = "Double"
                    Case "DECIMAL", "NUMERIC", "UNIQUEIDENTIFIER", "MONEY", "SMALLMONEY"
                        convertedDataType = "Decimal"
                    Case "NVARCHAR", "VARCHAR", "NTEXT", "TEXT", "XML"
                        convertedDataType = "String"
                    Case "BIT"
                        convertedDataType = "Boolean"
                    Case Else
                        Return Nothing
                End Select
            Case databaseType.ORACLE
                convertedDataType = ""
            Case databaseType.MYSQL
                Select Case typeDonnee.ToUpper
                    Case "MEDIUMINT", "INT", "BIGINT"
                        convertedDataType = "Long"
                    Case "TINYINT", "SMALLINT"
                        convertedDataType = "Integer"
                    Case "BINARY", "VARBINARY", "TINYINT", "BIT", "TINYBLOB", "BLOB", "MEDIUMBLOB", "LONGBLOB"
                        convertedDataType = "Byte"
                    Case "CHAR"
                        convertedDataType = "Char"
                    Case "DATE", "DATETIME", "YEAR", "TIME", "TIMESTAMP"
                        convertedDataType = "Date"
                    Case "FLOAT", "DOUBLE"
                        convertedDataType = "Double"
                    Case "DECIMAL", "NUMERIC"
                        convertedDataType = "Decimal"
                    Case "VARCHAR", "TINYTEXT", "TEXT", "MEDIUMTEXT", "LONGTEXT", "XML"
                        convertedDataType = "String"
                    Case Else
                        Return Nothing
                End Select
            Case databaseType.POSTGRE_SQL
                Select Case typeDonnee.ToUpper
                    Case "BIGINT"
                        convertedDataType = "Long"
                    Case "INTEGER", "SMALLINT"
                        convertedDataType = "Integer"
                    Case "BIT", "BIT VARYING", "BYTEA"
                        convertedDataType = "Byte"
                    Case "BOOLEAN"
                        convertedDataType = "Boolean"
                    Case "CHARACTER"
                        convertedDataType = "Char"
                    Case "DATE", "TIMESTAMPTZ", "INTERVAL", "TIME", "TIMESTAMP", "TIMESTAMP WITHOUT TIME ZONE"
                        convertedDataType = "Date"
                    Case "DOUBLE PRECISION", "MONEY"
                        convertedDataType = "Double"
                    Case "REAL", "NUMERIC"
                        convertedDataType = "Decimal"
                    Case "TEXT", "TSQUERY", "TSVECTOR", "UUID", "XML", "CHARACTER VARYING"
                        convertedDataType = "String"
                    Case Else
                        Return Nothing
                End Select
            Case databaseType.MS_ACCESS_97_2003
                convertedDataType = ""
            Case databaseType.MS_ACCESS_2007_2019
                convertedDataType = ""
            Case databaseType.MS_EXCEL
                convertedDataType = ""
            Case databaseType.FLAT_FILE
                convertedDataType = ""
            Case Else
                convertedDataType = ""
        End Select

        Return convertedDataType

    End Function

    Private Function getDataTypePrefix(ByVal dataType As String) As String

        Select Case dataType.ToUpper
            Case "LONG", "INT64"
                Return "Lng"
            Case "INTEGER", "INT32"
                Return "Int"
            Case "BYTE", "INT16"
                Return "Byte"
            Case "CHAR"
                Return "Chr"
            Case "DATE", "TIME"
                Return "Dte"
            Case "DOUBLE"
                Return "Dbl"
            Case "SINGLE"
                Return "Sng"
            Case "DECIMAL"
                Return "Dec"
            Case "STRING"
                Return "Str"
            Case Else
                Return UppercaseFirstLetter(Strings.Left(dataType, 3))
        End Select

    End Function

    Private Function getDefaultValue(ByVal dataType As String) As String

        Select Case dataType.ToUpper
            Case "LONG", "INT64", "INTEGER", "INT32", "BYTE", "INT16", "DOUBLE", "SINGLE", "DECIMAL"
                Return "0"
            Case "CHAR", "STRING"
                Return """"""
            Case "DATE", "TIME"
                Return "Now()"
            Case Else
                Return "Nothing"
        End Select

    End Function

    Private Function UppercaseFirstLetter(ByVal val As String) As String
        ' Test for nothing or empty.
        If String.IsNullOrEmpty(val) Then
            Return val
        End If

        ' Convert to character array.
        Dim array() As Char = val.ToCharArray

        ' Uppercase first character.
        array(0) = Char.ToUpper(array(0))

        ' Return new string.
        Return New String(array)
    End Function

    Private Function getNumberTab(ByVal numberOfTab As Integer) As String
        Dim result As String = ""
        For i = 1 To numberOfTab
            result += vbTab
        Next
        Return result
    End Function

    Private Function getPKSignature(ByVal uneListe As List(Of UneColonne), ByVal paramByValue As Boolean) As String

        Dim uneDef As String = ""
        Dim result As String = ""
        Dim byValByRef As String = ""

        If (paramByValue) Then
            byValByRef = "ByVal"
        Else
            byValByRef = "ByRef"
        End If


        For Each col As UneColonne In uneListe
            If (col.IsPrimaryKey) Then
                uneDef = byValByRef & " _" & col.NomColonneOriginal & " AS " & col.TypeDeDonnees
            End If
            result = uneDef & ", "
        Next

        Return Strings.Left(result, result.Length - 2)

    End Function

    Private Function CreateClasses() As Integer

        Dim result As Integer = 0
        Dim leNomDeLaTable As TableName

        Dim qryDbName As String = ""
        Dim qryTblName As String = ""
        Dim qryOneField As String = ""

        Try

            For Each db As DbDatabase In Me.ListeTables

                For Each tbl As DbTable In db.Childrens

                    Dim uneClasse As String = ""
                    Dim listeDesColonnes As New List(Of UneColonne)

                    If (tbl.IsChecked Is Nothing) Or (tbl.IsChecked) Then

                        leNomDeLaTable = New TableName(CType(tbl, TreeView.Noeud).Name)

                        '************************************************************************************
                        'Entête de la classe
                        '************************************************************************************

                        uneClasse = "'************************************************************************************" & vbCrLf
                        uneClasse &= "'* DÉVELOPPÉ PAR : " & My.Settings.NOM_PROGRAMMEUR & Strings.Space(66 - My.Settings.NOM_PROGRAMMEUR.Length) & "*" & vbCrLf
                        uneClasse &= "'* DATE : " & Now().ToLongDateString() & Strings.Space(75 - Now().ToLongDateString().Length) & "*" & vbCrLf
                        uneClasse &= "'* MODIFIÉ : " & Now().ToLongDateString() & Strings.Space(72 - Now().ToLongDateString().Length) & "*" & vbCrLf
                        uneClasse &= "'* PAR :                                                                            *" & vbCrLf
                        uneClasse &= "'* DESCRIPTION :                                                                    *" & vbCrLf
                        uneClasse &= "'*      Public [Function | Sub] nomProcFunct( paramètres)                           *" & vbCrLf
                        uneClasse &= "'*      Public Function ToString() as String                                        *" & vbCrLf
                        uneClasse &= "'*      Public Function Equals() as Boolean                                         *" & vbCrLf
                        uneClasse &= "'************************************************************************************" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= "'************************************************************************************" & vbCrLf
                        uneClasse &= "'                                                                                   *" & vbCrLf
                        uneClasse &= "'                           L I B R A R Y  I M P O R T S                            *" & vbCrLf
                        uneClasse &= "'                                                                                   *" & vbCrLf
                        uneClasse &= "'************************************************************************************" & vbCrLf
                        uneClasse &= "#Region ""Library Imports""" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= "Imports System.ComponentModel" & vbCrLf
                        uneClasse &= "Imports System.Data.Common" & vbCrLf
                        Select Case Me.TypeBaseDonnees
                            Case databaseType.SQL_SERVER
                                uneClasse &= "Imports System.Data.OleDb" & vbCrLf
                            Case databaseType.MYSQL
                                uneClasse &= "Imports MySql.Data.MySqlClient" & vbCrLf
                            Case databaseType.ORACLE
                            Case databaseType.POSTGRE_SQL
                                uneClasse &= "Imports NpgSql" & vbCrLf
                            Case databaseType.MS_ACCESS_2007_2019
                            Case databaseType.MS_ACCESS_97_2003
                            Case databaseType.MS_EXCEL
                            Case databaseType.FLAT_FILE
                            Case Else

                        End Select
                        uneClasse &= "Imports System.Runtime.CompilerServices" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= "#End Region" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= "Public Class " & leNomDeLaTable.ClassName & vbCrLf
                        uneClasse &= getNumberTab(1) & "Implements IDisposable, INotifyPropertyChanged" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & " '************************************************************************************" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                            V  A  R  I  A  B  L  E  S                              *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                        D E C L A R E   F U N C T I O N S                          *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                                    T Y P E S                                      *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= "#Region ""Variables, Declare Functions And Types""" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section privée" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "' Field to handle multiple calls to Dispose gracefully." & vbCrLf
                        uneClasse &= getNumberTab(1) & "Private disposed As Boolean = False" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Private BooIsSaved As Boolean = True" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Classe Variables" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Liste des variables
                        '************************************************************************************

                        Dim constructeurHeader As String = "Public Sub New("
                        Dim constructeurContenu As String = ""
                        Dim constructeurContenu2 As String = ""
                        Dim proprietes As String = ""
                        Dim equalsFunction As String = "Return ("
                        Dim toStringFunction As String = ""
                        Dim toStringValueFound As Boolean = False

                        For Each col As DbChampTable In tbl.Childrens

                            If (col.IsChecked Is Nothing) Or (col.IsChecked) Then

                                Dim nombreEntreParentheses As String = ""
                                Dim typeDonnee As String = col.DataType
                                Dim typeDonneeVB As String
                                Dim infosColonne As UneColonne

                                If (InStr(typeDonnee, "(") > 0) Then
                                    nombreEntreParentheses = Mid(typeDonnee, InStr(typeDonnee, "(") + 1, InStr(typeDonnee, ")") - 1)
                                    typeDonnee = Left(typeDonnee, InStr(typeDonnee, "(") - 1)
                                End If

                                typeDonneeVB = getTypeFromString(typeDonnee)

                                If (typeDonneeVB IsNot Nothing) Then

                                    infosColonne = New UneColonne(col.Name,
                                                                  getDataTypePrefix(typeDonneeVB) & CultureInfo.CurrentCulture.TextInfo.ToTitleCase(col.Name.Replace("_", " ").ToLower).Replace(" ", ""),
                                                                  CultureInfo.CurrentCulture.TextInfo.ToTitleCase(col.Name.Replace("_", " ").ToLower).Replace(" ", ""),
                                                                  typeDonneeVB,
                                                                  col.IsPrimaryKey)

                                    Select Case typeDonneeVB.ToUpper
                                        Case "BINARY", "VARBINARY", "CHAR", "NCHAR"

                                            If (nombreEntreParentheses <> "") Then

                                                infosColonne.VarDeClasse &= "(" & nombreEntreParentheses & ")"
                                                infosColonne.VarDePropriete &= "(" & nombreEntreParentheses & ")"
                                                infosColonne.VarDeConstructeur &= "()"

                                            End If

                                        Case Else

                                    End Select

                                    uneClasse &= getNumberTab(1) & "Private " & infosColonne.VarDeClasse & " As " & typeDonneeVB & vbCrLf
                                    proprietes &= getNumberTab(1) & "Public Property " & infosColonne.VarDePropriete & " As " & typeDonneeVB & vbCrLf
                                    constructeurHeader &= "Byval " & infosColonne.VarDeConstructeur & " As " & typeDonneeVB & ", "

                                    If (Not toStringValueFound) And
                                       (infosColonne.VarDeClasse.ToUpper.Contains("NAME") Or
                                        infosColonne.VarDeClasse.ToUpper.Contains("NOM") Or
                                        infosColonne.VarDeClasse.ToUpper.Contains("DESC") Or
                                        infosColonne.VarDeClasse.ToUpper.Contains("CODE") Or
                                        infosColonne.VarDeClasse.ToUpper.Contains("ID")) Then

                                        If (MsgBox("Est-ce que le champ """ & infosColonne.VarDeClasse & """ est le meilleur champ pour la fonction ""ToString()"" de la table " & leNomDeLaTable.TableName & "?", CType(MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, Global.Microsoft.VisualBasic.MsgBoxStyle), "Fonction ToString()") = vbYes) Then

                                            Select Case infosColonne.TypeDeDonnees
                                                Case "Boolean", "Byte", "Char", "Integer", "Long", "Single", "Double", "Date", "Currency", "Decimal"
                                                    toStringFunction = "CStr(" & infosColonne.VarDeClasse & ")"
                                                Case "String"
                                                    toStringFunction = infosColonne.VarDeClasse
                                                Case Else
                                                    toStringFunction = infosColonne.VarDeClasse & ".ToString()"
                                            End Select
                                            toStringValueFound = True

                                        End If
                                    End If

                                    Select Case typeDonneeVB.ToUpper
                                        Case "BINARY", "VARBINARY", "CHAR", "NCHAR"

                                            If (nombreEntreParentheses <> "") Then

                                                For i As Integer = 0 To CInt(nombreEntreParentheses)
                                                    constructeurContenu2 &= getNumberTab(2) & infosColonne.VarDeClasse & "(" & i & ")"
                                                    constructeurContenu2 &= " = " & getDefaultValue(typeDonneeVB) & vbCrLf
                                                Next

                                            Else

                                                constructeurContenu2 &= getNumberTab(2) & infosColonne.VarDeClasse
                                                constructeurContenu2 &= " = " & getDefaultValue(typeDonneeVB) & vbCrLf

                                            End If

                                        Case Else

                                            constructeurContenu2 &= getNumberTab(2) & infosColonne.VarDeClasse
                                            constructeurContenu2 &= " = " & getDefaultValue(typeDonneeVB) & vbCrLf

                                    End Select

                                    equalsFunction &= "(Me." & infosColonne.VarDeClasse
                                    equalsFunction &= "= value." & infosColonne.VarDePropriete & ") And "
                                    constructeurContenu &= getNumberTab(2) & infosColonne.VarDeClasse
                                    constructeurContenu &= " = "
                                    constructeurContenu &= "_" & infosColonne.VarDePropriete & vbCrLf

                                    proprietes &= getNumberTab(2) & "Get" & vbCrLf
                                    proprietes &= getNumberTab(3) & vbTab & "Return " & infosColonne.VarDeClasse & vbCrLf
                                    proprietes &= getNumberTab(2) & "End Get" & vbCrLf
                                    proprietes &= getNumberTab(2) & "Set(value As " & typeDonneeVB & ")" & vbCrLf
                                    proprietes &= getNumberTab(3) & infosColonne.VarDeClasse & " = value" & vbCrLf
                                    proprietes &= getNumberTab(3) & "BooIsSaved = False" & vbCrLf
                                    proprietes &= getNumberTab(3) & "OnPropertyChanged()" & vbCrLf
                                    proprietes &= getNumberTab(2) & "End Set" & vbCrLf
                                    proprietes &= getNumberTab(1) & "End Property" & vbCrLf
                                    proprietes &= "" & vbCrLf

                                    listeDesColonnes.Add(infosColonne)

                                End If

                            End If

                        Next

                        constructeurHeader = Strings.Left(constructeurHeader, constructeurHeader.Length - 2) & ")" & vbCrLf

                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section publique" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf

                        '************************************************************************************
                        'Ajouter la connection string à la classe
                        '************************************************************************************

                        uneClasse &= "" & vbCrLf
                        If (cnxstr <> "") Then
                            uneClasse &= "Public Const OLEDB_CONN_STRING = """ & cnxstr & """" & vbCrLf
                        End If
                        uneClasse &= "" & vbCrLf

                        uneClasse &= "#End Region" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                    C  O  N  S  T  R  U  C  T  E  U  R                             *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                    ----------------------------------                             *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                      D  E  S  T  R  U  C  T  E  U  R                              *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= getNumberTab(1) & "#Region ""Constructors""" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Constructeur sans paramètre
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "''' <summary>" & vbCrLf
                        uneClasse &= getNumberTab(1) & " ''' Constructeur de base sans paramètre" & vbCrLf
                        uneClasse &= getNumberTab(1) & "''' </summary>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Public Sub New()" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "' This call is required by the designer." & vbCrLf
                        uneClasse &= getNumberTab(2) & "'InitializeComponent()" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "' Add any initialization after the InitializeComponent() call." & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= constructeurContenu2
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Sub" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        ' Constructeur avec paramètres d'identification (ID, PK)
                        ' À faire
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "''' <summary>" & vbCrLf
                        uneClasse &= getNumberTab(1) & " ''' Constructeur de base sans paramètre" & vbCrLf
                        uneClasse &= getNumberTab(1) & "''' </summary>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Public Sub New(" & getPKSignature(listeDesColonnes, True) & ")" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(2) & " ' This call is required by the designer." & vbCrLf
                        uneClasse &= getNumberTab(2) & "'InitializeComponent()" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "' Add any initialization after the InitializeComponent() call." & vbCrLf
                        uneClasse &= "" & vbCrLf
                        'uneClasse &= constructeurContenu2
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Sub" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Constructeur avec paramètres
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "''' <summary>" & vbCrLf
                        uneClasse &= getNumberTab(1) & " ''' Constructeur de base avec paramètres" & vbCrLf
                        uneClasse &= getNumberTab(1) & "''' </summary>" & vbCrLf
                        uneClasse &= getNumberTab(1) & constructeurHeader
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "' This call is required by the designer." & vbCrLf
                        uneClasse &= getNumberTab(1) & "'InitializeComponent()" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "' Add any initialization after the InitializeComponent() call." & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= constructeurContenu
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Sub" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= "#End Region" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Destructeurs
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "' Implement IDisposable." & vbCrLf
                        uneClasse &= "#Region ""IDisposable implementation""" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Public Overloads Sub Dispose() Implements IDisposable.Dispose" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dispose(True)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "GC.SuppressFinalize(Me)" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Sub" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Protected Overridable Overloads Sub Dispose(disposing As Boolean)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "If disposed = False Then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "If disposing Then" & vbCrLf
                        uneClasse &= getNumberTab(4) & "' Free other state (managed objects)." & vbCrLf
                        uneClasse &= getNumberTab(4) & "disposed = True" & vbCrLf
                        uneClasse &= getNumberTab(3) & "End If" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "' Free your own state (unmanaged objects)." & vbCrLf
                        uneClasse &= getNumberTab(2) & "' Set large fields to null." & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Sub" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Protected Overrides Sub Finalize()" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "' Simply call Dispose(False)." & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dispose(False)" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Sub" & vbCrLf
                        uneClasse &= "#End Region" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Entêtre propriétés
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                           P  R  O  P  R  I  É  T  É  S                            *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= "#Region ""Properties""" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section privée" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section publique" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Propriétés
                        '************************************************************************************

                        uneClasse &= proprietes

                        uneClasse &= "#End Region" & vbCrLf

                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                           P  R  O  C  É  D  U  R  E  S                            *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= "#Region ""Procédures""" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section privée" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf

                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section publique" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= "#End Region" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                             F  O  N  C  T  I  O  N  S                             *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= "#Region ""Functions""" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        Dim sqlTables As String = ""
                        Dim CommandeSQL As String = ""
                        Dim NewConnections As String = ""
                        Dim sqlConditions As String = ""
                        Dim optionalConnection As String = ", Optional Byref AnOpenConneciton as DbConnection = Nothing"

                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section privée" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Création de la fonction insert de la classe
                        '************************************************************************************

                        uneClasse &= getNumberTab(2) & "''' <summary>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' Fonction permettant d'insérer dans la table de la classe l'objet passé en paramètre." & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' </summary>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' <param name=""_" & leNomDeLaTable.NomSingulier & """>Un objet de type " & leNomDeLaTable.TableName & "</param>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Private Shared Function Insert" & leNomDeLaTable.NomSingulier & "(ByVal _" & leNomDeLaTable.NomSingulier & " As " & leNomDeLaTable.ClassName & optionalConnection & ") As Integer" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim resultat As Integer = 0" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim sqlTables As String" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aConn As DbConnection" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aCommand As DbCommand" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        Select Case Me.TypeBaseDonnees
                            Case databaseType.SQL_SERVER, databaseType.MYSQL, databaseType.ORACLE, databaseType.MS_ACCESS_2007_2019, databaseType.MS_ACCESS_97_2003, databaseType.MS_EXCEL, databaseType.FLAT_FILE
                                qryDbName = db.Name
                                qryTblName = tbl.Name
                            Case databaseType.POSTGRE_SQL
                                qryDbName = """""" & db.Name & """"""
                                qryTblName = ("""""" & tbl.Name & """""").Replace(".", """"".""""")
                            Case Else

                        End Select                    'Création de la requête sql      
                        sqlTables = "INSERT INTO " & qryDbName & "." & qryTblName & " ("
                        For Each col As UneColonne In listeDesColonnes
                            sqlTables &= col.NomColonneOriginal & ", "
                        Next

                        sqlTables = Strings.Left(sqlTables, sqlTables.Length - 2) & ") VALUES ("

                        For Each col As UneColonne In listeDesColonnes

                            If (col Is Nothing) OrElse (col.TypeDeDonnees Is Nothing) OrElse (col.TypeDeDonnees = "") Then
                                Continue For
                            Else
                                Select Case col.TypeDeDonnees
                                    Case "Long", "Integer", "Byte", "Double", "Decimal"
                                        sqlTables &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "", "
                                    Case "Char", "DateDate", "String"
                                        sqlTables &= "'"" & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & ""', "
                                    Case Else
                                        sqlTables &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "", "
                                End Select
                            End If

                        Next

                        sqlTables = Strings.Left(sqlTables, sqlTables.Length - 2) & ");"

                        Select Case CType(Me.lngTypeBaseDonnees, databaseType)
                            Case databaseType.SQL_SERVER

                                NewConnections = "New OleDbConnection()"
                                CommandeSQL = "New OleDbCommand(sqlTables, CType(aConn, OleDbConnection))"

                            Case databaseType.ORACLE
                            Case databaseType.MYSQL

                                NewConnections = "New MySqlConnection()"
                                CommandeSQL = "New MySqlCommand(sqlTables, CType(aConn, MySqlConnection))"

                            Case databaseType.POSTGRE_SQL

                                NewConnections = "New NpgSqlConnection()"
                                CommandeSQL = "New NpgSqlCommand(sqlTables, CType(aConn, NpgSqlConnection))"

                            Case databaseType.MS_ACCESS_97_2003
                            Case databaseType.MS_ACCESS_2007_2019
                            Case databaseType.MS_EXCEL
                            Case databaseType.FLAT_FILE
                        End Select

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "'Ouverture de la connection SQL" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = " & NewConnections & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.ConnectionString = OLEDB_CONN_STRING" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Open()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Else" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = AnOpenConneciton" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf

                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création de la requête sql" & vbCrLf
                        uneClasse &= getNumberTab(2) & "sqlTables = """ & sqlTables & """" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création de la commande et on l'exécute" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aCommand = " & CommandeSQL & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "resultat = aCommand.ExecuteNonQuery()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aCommand = Nothing" & vbCrLf

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Close()" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = Nothing" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf

                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Return resultat" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Function" & vbCrLf
                        uneClasse &= getNumberTab(1) & "" & vbCrLf

                        '************************************************************************************
                        'Création de la fonction update de la classe
                        '************************************************************************************

                        uneClasse &= getNumberTab(2) & "''' <summary>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' Fonction permettant de mettre à jour l'objet passé en paramètre dans la table de la classe s'il existe." & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' </summary>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' <param name=""_" & leNomDeLaTable.NomSingulier & """>Un objet de type " & leNomDeLaTable.TableName & "</param>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Private Shared Function Update" & leNomDeLaTable.NomSingulier & "(ByVal _" & leNomDeLaTable.NomSingulier & " As " & leNomDeLaTable.ClassName & optionalConnection & ") As Integer" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim resultat As Integer = 0" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim sqlTables As String" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aConn As DbConnection" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aCommand As DbCommand" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        'Création de la requête sql      
                        sqlTables = "UPDATE " & qryDbName & "." & qryTblName & " SET "
                        sqlConditions = "WHERE "
                        For Each col As UneColonne In listeDesColonnes

                            If (col.IsPrimaryKey) Then
                                sqlConditions &= col.NomColonneOriginal & " = "
                                Select Case col.TypeDeDonnees
                                    Case "Long", "Integer", "Byte", "Double", "Decimal"
                                        sqlConditions &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "", "
                                    Case "Char", "DateDate", "String"
                                        sqlConditions &= "'"" & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & ""', "
                                    Case Else
                                        sqlConditions &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "", "
                                End Select
                            Else

                                sqlTables &= col.NomColonneOriginal & " = "
                                If (col Is Nothing) OrElse (col.TypeDeDonnees Is Nothing) OrElse (col.TypeDeDonnees = "") Then
                                    Continue For
                                Else
                                    Select Case col.TypeDeDonnees
                                        Case "Long", "Integer", "Byte", "Double", "Decimal"
                                            sqlTables &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "", "
                                        Case "Char", "DateDate", "String"
                                            sqlTables &= "'"" & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & ""', "
                                        Case Else
                                            sqlTables &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "", "
                                    End Select
                                End If

                            End If

                        Next

                        sqlTables = Strings.Left(sqlTables, sqlTables.Length - 2) & " " & Strings.Left(sqlConditions, sqlConditions.Length - 2) & ";"

                        Select Case CType(Me.lngTypeBaseDonnees, databaseType)
                            Case databaseType.SQL_SERVER

                                NewConnections = "New OleDbConnection()"
                                CommandeSQL = "New OleDbCommand(sqlTables, CType(aConn, OleDbConnection))"

                            Case databaseType.ORACLE
                            Case databaseType.MYSQL

                                NewConnections = "New MySqlConnection()"
                                CommandeSQL = "New MySqlCommand(sqlTables, CType(aConn, MySqlConnection))"

                            Case databaseType.POSTGRE_SQL

                                NewConnections = "New NpgSqlConnection()"
                                CommandeSQL = "New NpgSqlCommand(sqlTables, CType(aConn, NpgSqlConnection))"

                            Case databaseType.MS_ACCESS_97_2003
                            Case databaseType.MS_ACCESS_2007_2019
                            Case databaseType.MS_EXCEL
                            Case databaseType.FLAT_FILE
                        End Select

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "'Ouverture de la connection SQL" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = " & NewConnections & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.ConnectionString = OLEDB_CONN_STRING" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Open()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Else" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = AnOpenConneciton" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf

                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création de la requête sql" & vbCrLf
                        uneClasse &= getNumberTab(2) & "sqlTables = """ & sqlTables & """" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création de la commande et on l'exécute" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aCommand = " & CommandeSQL & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "resultat = aCommand.ExecuteNonQuery()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aCommand = Nothing" & vbCrLf

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Close()" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = Nothing" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf

                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Return resultat" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Function" & vbCrLf

                        uneClasse &= getNumberTab(1) & "" & vbCrLf

                        '************************************************************************************
                        'Création de la fonction objet existe de la classe
                        '************************************************************************************

                        uneClasse &= getNumberTab(2) & "''' <summary>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' Fonction permettant de déterminer si l'objet passé en paramètre existe dans la base de données ou non." & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' </summary>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' <param name=""_" & leNomDeLaTable.NomSingulier & """>Un objet de type " & leNomDeLaTable.TableName & "</param>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' <returns>Retourne Vrai si l'objet existe dans la base de données ou False s'il n'existe pas.</returns>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Private Shared Function " & leNomDeLaTable.NomSingulier & "Exists(ByVal _" & leNomDeLaTable.NomSingulier & " As " & leNomDeLaTable.ClassName & optionalConnection & ") As Boolean" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim resultat As Boolean = False" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim sqlTables As String" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aConn As DbConnection" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aCommand As DbCommand" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        'Création de la requête sql      
                        sqlTables = "SELECT COUNT(*) AS EST_EXISTANT FROM " & qryDbName & "." & qryTblName & " "
                        sqlConditions = "WHERE "
                        For Each col As UneColonne In listeDesColonnes

                            If (col.IsPrimaryKey) Then
                                sqlConditions &= col.NomColonneOriginal & " = "
                                Select Case col.TypeDeDonnees
                                    Case "Long", "Integer", "Byte", "Double", "Decimal"
                                        sqlConditions &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "" AND "
                                    Case "Char", "DateDate", "String"
                                        sqlConditions &= "'"" & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & ""' AND "
                                    Case Else
                                        sqlConditions &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "" AND "
                                End Select
                            End If

                        Next

                        sqlTables = sqlTables & Strings.Left(sqlConditions, sqlConditions.Length - 5) & ";"

                        Select Case CType(Me.lngTypeBaseDonnees, databaseType)
                            Case databaseType.SQL_SERVER

                                NewConnections = "New OleDbConnection()"
                                CommandeSQL = "New OleDbCommand(sqlTables, CType(aConn, OleDbConnection))"

                            Case databaseType.ORACLE
                            Case databaseType.MYSQL

                                NewConnections = "New MySqlConnection()"
                                CommandeSQL = "New MySqlCommand(sqlTables, CType(aConn, MySqlConnection))"

                            Case databaseType.POSTGRE_SQL

                                NewConnections = "New NpgSqlConnection()"
                                CommandeSQL = "New NpgSqlCommand(sqlTables, CType(aConn, NpgSqlConnection))"

                            Case databaseType.MS_ACCESS_97_2003
                            Case databaseType.MS_ACCESS_2007_2019
                            Case databaseType.MS_EXCEL
                            Case databaseType.FLAT_FILE
                        End Select

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "'Ouverture de la connection SQL" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = " & NewConnections & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.ConnectionString = OLEDB_CONN_STRING" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Open()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Else" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = AnOpenConneciton" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf

                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création de la requête sql" & vbCrLf
                        uneClasse &= getNumberTab(2) & "sqlTables = """ & sqlTables & """" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création de la commande et on l'exécute" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aCommand = " & CommandeSQL & vbCrLf
                        uneClasse &= getNumberTab(2) & "Try" & vbCrLf
                        uneClasse &= getNumberTab(3) & "resultat = (CType(aCommand.ExecuteScalar(), Integer) > 0)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Catch ex As Exception" & vbCrLf
                        uneClasse &= getNumberTab(3) & "resultat = False" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End Try" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aCommand = Nothing" & vbCrLf

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Close()" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = Nothing" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf

                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Return resultat" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Function" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        '************************************************************************************
                        'Création de la fonction delete objet de la classe
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "''' <summary>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "''' Fonction permettant de supprimer l'objet passé en paramètre de la base de données." & vbCrLf
                        uneClasse &= getNumberTab(1) & "''' </summary>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "''' <param name=""_" & leNomDeLaTable.NomSingulier & """>Un objet de type " & leNomDeLaTable.TableName & "</param>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Private Shared Function Delete" & leNomDeLaTable.NomSingulier & "(ByVal _" & leNomDeLaTable.NomSingulier & " As " & leNomDeLaTable.ClassName & optionalConnection & ") As Integer" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim resultat As Integer = 0" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(3) & "Dim sqlTables As String" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aConn As DbConnection" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aCommand As DbCommand" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "If (" & leNomDeLaTable.NomSingulier & "Exists(_" & leNomDeLaTable.NomSingulier & ")) Then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "" & vbCrLf

                        'Création de la requête sql      
                        sqlTables = "DELETE FROM " & qryDbName & "." & qryTblName & " "
                        sqlConditions = "WHERE "
                        For Each col As UneColonne In listeDesColonnes

                            If (col.IsPrimaryKey) Then
                                sqlConditions &= col.NomColonneOriginal & " = "
                                Select Case col.TypeDeDonnees
                                    Case "Long", "Integer", "Byte", "Double", "Decimal"
                                        sqlConditions &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "" AND "
                                    Case "Char", "DateDate", "String"
                                        sqlConditions &= "'"" & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & ""' AND "
                                    Case Else
                                        sqlConditions &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "" AND "
                                End Select
                            End If

                        Next

                        sqlTables = sqlTables & Strings.Left(sqlConditions, sqlConditions.Length - 5) & ";"

                        Select Case CType(Me.lngTypeBaseDonnees, databaseType)
                            Case databaseType.SQL_SERVER

                                NewConnections = "New OleDbConnection()"
                                CommandeSQL = "New OleDbCommand(sqlTables, CType(aConn, OleDbConnection))"


                            Case databaseType.ORACLE
                            Case databaseType.MYSQL

                                NewConnections = "New MySqlConnection()"
                                CommandeSQL = "New MySqlCommand(sqlTables, CType(aConn, MySqlConnection))"

                            Case databaseType.POSTGRE_SQL

                                NewConnections = "New NpgSqlConnection()"
                                CommandeSQL = "New NpgSqlCommand(sqlTables, CType(aConn, NpgSqlConnection))"

                            Case databaseType.MS_ACCESS_97_2003
                            Case databaseType.MS_ACCESS_2007_2019
                            Case databaseType.MS_EXCEL
                            Case databaseType.FLAT_FILE
                        End Select

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "'Ouverture de la connection SQL" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = " & NewConnections & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.ConnectionString = OLEDB_CONN_STRING" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Open()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Else" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = AnOpenConneciton" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf

                        uneClasse &= getNumberTab(3) & "" & vbCrLf
                        uneClasse &= getNumberTab(3) & "'Création de la requête sql" & vbCrLf
                        uneClasse &= getNumberTab(3) & "sqlTables = """ & sqlTables & """" & vbCrLf
                        uneClasse &= getNumberTab(3) & "" & vbCrLf
                        uneClasse &= getNumberTab(3) & "'Création de la commande et on l'exécute" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aCommand = " & CommandeSQL & vbCrLf
                        uneClasse &= getNumberTab(3) & "Try" & vbCrLf
                        uneClasse &= getNumberTab(3) & "resultat =  CType(aCommand.ExecuteScalar(), Integer)" & vbCrLf
                        uneClasse &= getNumberTab(3) & "Catch ex As Exception" & vbCrLf
                        uneClasse &= getNumberTab(3) & "resultat = -1" & vbCrLf
                        uneClasse &= getNumberTab(3) & "End Try" & vbCrLf
                        uneClasse &= getNumberTab(3) & "" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aCommand = Nothing" & vbCrLf

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Close()" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = Nothing" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf

                        uneClasse &= getNumberTab(3) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Else" & vbCrLf
                        uneClasse &= getNumberTab(3) & "" & vbCrLf
                        uneClasse &= getNumberTab(3) & "MsgBox(""L'objet """"" & leNomDeLaTable.TableName & """"" à supprimer n'existe pas."")" & vbCrLf
                        uneClasse &= getNumberTab(3) & "resultat = -1" & vbCrLf
                        uneClasse &= getNumberTab(3) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Return resultat" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Function" & vbCrLf

                        uneClasse &= getNumberTab(1) & "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section publique" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Création de la fonction getAll objets de la classe
                        '************************************************************************************

                        uneClasse &= getNumberTab(2) & "''' <summary>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' Retourne une liste de tous les " & leNomDeLaTable.TableName & " de la table." & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' </summary>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' <returns>Retourne une liste de tous les " & leNomDeLaTable.TableName & " de la table</returns>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Public Shared Function getAll" & leNomDeLaTable.TableName & "(" & Strings.Right(optionalConnection, optionalConnection.Length - 2) & ") As List(Of " & leNomDeLaTable.ClassName & ")" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        uneClasse &= getNumberTab(2) & "Dim lst As New List(Of " & leNomDeLaTable.ClassName & ")" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim sqlTables As String" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aConn As DbConnection" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aCommand as DbCommand" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aDtr as DbDataReader" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        'Création de la requête sql      
                        sqlTables = "SELECT * " &
                                    "FROM " & qryDbName & "." & qryTblName & ";"

                        Select Case CType(Me.lngTypeBaseDonnees, databaseType)
                            Case databaseType.SQL_SERVER

                                NewConnections = "New OleDbConnection()"
                                CommandeSQL = "New OleDbCommand(sqlTables, CType(aConn, OleDbConnection))"

                            Case databaseType.ORACLE
                            Case databaseType.MYSQL

                                NewConnections = "New MySqlConnection()"
                                CommandeSQL = "New MySqlCommand(sqlTables, CType(aConn, MySqlConnection))"

                            Case databaseType.POSTGRE_SQL

                                NewConnections = "New NpgSqlConnection()"
                                CommandeSQL = "New NpgSqlCommand(sqlTables, CType(aConn, NpgSqlConnection))"

                            Case databaseType.MS_ACCESS_97_2003
                            Case databaseType.MS_ACCESS_2007_2019
                            Case databaseType.MS_EXCEL
                            Case databaseType.FLAT_FILE
                        End Select

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "'Ouverture de la connection SQL" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = " & NewConnections & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.ConnectionString = OLEDB_CONN_STRING" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Open()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Else" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = AnOpenConneciton" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf

                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création de la requête sql" & vbCrLf
                        uneClasse &= getNumberTab(2) & "sqlTables = """ & sqlTables & """" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création de la commande et on l'instancie (sql)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aCommand = " & CommandeSQL & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création du datareader (aDtr)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aDtr = aCommand.ExecuteReader()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "While aDtr.Read()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        uneClasse &= getNumberTab(3) & "Dim uneTable as new " & leNomDeLaTable.ClassName & "()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        For Each col As UneColonne In listeDesColonnes


                            If (col Is Nothing) OrElse (col.TypeDeDonnees Is Nothing) OrElse (col.TypeDeDonnees = "") Then
                                Continue For
                            Else
                                uneClasse &= getNumberTab(3) & "uneTable." & col.VarDePropriete & " = " & ' aDtr.Item(""" & col.NomColonneOriginal & """)" & vbCrLf
                                                               "If(IsDBNull(aDtr.Item(""" & col.NomColonneOriginal & """)), Nothing, CType(aDtr.Item(""" & col.NomColonneOriginal & """), " & col.TypeDeDonnees & "))" & vbCrLf
                            End If

                        Next

                        uneClasse &= getNumberTab(3) & "" & vbCrLf
                        uneClasse &= getNumberTab(3) & "lst.Add(uneTable)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End While" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        uneClasse &= getNumberTab(2) & "aDtr.Close()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aDtr = Nothing" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aCommand = Nothing" & vbCrLf

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Close()" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = Nothing" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf



                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Return lst" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Function" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Création de la fonction get 1 objet de la classe à partir de la Primary Key
                        '************************************************************************************

                        uneClasse &= getNumberTab(2) & "''' <summary>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' Retourne une liste de tous les " & leNomDeLaTable.TableName & " de la table." & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' </summary>" & vbCrLf
                        uneClasse &= getNumberTab(2) & "''' <returns>Retourne une liste de tous les " & leNomDeLaTable.TableName & " de la table</returns>" & vbCrLf
                        uneClasse &= getNumberTab(1) & "Public Shared Function get" & leNomDeLaTable.TableName & "FromID("

                        For Each col As UneColonne In listeDesColonnes
                            If (col.IsPrimaryKey) Then
                                uneClasse &= "Byval _" & col.NomColonneOriginal & " as " & col.TypeDeDonnees & ", "
                            End If
                        Next

                        uneClasse &=  Strings.Right(optionalConnection, optionalConnection.Length - 2) & ") As " & leNomDeLaTable.ClassName & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        uneClasse &= getNumberTab(2) & "Dim result As New " & leNomDeLaTable.ClassName &  vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim sqlTables As String" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aConn As DbConnection" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aCommand as DbCommand" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim aDtr as DbDataReader" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        'Création de la requête sql      
                        sqlTables = "SELECT * " &
                                    "FROM " & qryDbName & "." & qryTblName & " "
                        sqlConditions = "WHERE "
                        For Each col As UneColonne In listeDesColonnes

                            If (col.IsPrimaryKey) Then
                                sqlConditions &= col.NomColonneOriginal & " = "
                                Select Case col.TypeDeDonnees
                                    Case "Long", "Integer", "Byte", "Double", "Decimal"
                                        sqlConditions &= """ & _" & col.NomColonneOriginal & " & "" AND "
                                    Case "Char", "DateDate", "String"
                                        sqlConditions &= "'"" & _" & col.NomColonneOriginal & " & ""' AND "
                                    Case Else
                                        sqlConditions &= """ & _" & col.NomColonneOriginal & " & "" AND "
                                End Select
                            End If

                        Next

                        sqlTables = sqlTables & Strings.Left(sqlConditions, sqlConditions.Length - 5) & ";"

                        Select Case CType(Me.lngTypeBaseDonnees, databaseType)
                            Case databaseType.SQL_SERVER

                                NewConnections = "New OleDbConnection()"
                                CommandeSQL = "New OleDbCommand(sqlTables, CType(aConn, OleDbConnection))"

                            Case databaseType.ORACLE
                            Case databaseType.MYSQL

                                NewConnections = "New MySqlConnection()"
                                CommandeSQL = "New MySqlCommand(sqlTables, CType(aConn, MySqlConnection))"

                            Case databaseType.POSTGRE_SQL

                                NewConnections = "New NpgSqlConnection()"
                                CommandeSQL = "New NpgSqlCommand(sqlTables, CType(aConn, NpgSqlConnection))"

                            Case databaseType.MS_ACCESS_97_2003
                            Case databaseType.MS_ACCESS_2007_2019
                            Case databaseType.MS_EXCEL
                            Case databaseType.FLAT_FILE
                        End Select

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "'Ouverture de la connection SQL" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = " & NewConnections & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.ConnectionString = OLEDB_CONN_STRING" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Open()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Else" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = AnOpenConneciton" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf

                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création de la requête sql" & vbCrLf
                        uneClasse &= getNumberTab(2) & "sqlTables = """ & sqlTables & """" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création de la commande et on l'instancie (sql)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aCommand = " & CommandeSQL & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "'Création du datareader (aDtr)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aDtr = aCommand.ExecuteReader()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "While aDtr.Read()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        For Each col As UneColonne In listeDesColonnes


                            If (col Is Nothing) OrElse (col.TypeDeDonnees Is Nothing) OrElse (col.TypeDeDonnees = "") Then
                                Continue For
                            Else
                                uneClasse &= getNumberTab(3) & "result." & col.VarDePropriete & " = " & ' aDtr.Item(""" & col.NomColonneOriginal & """)" & vbCrLf
                                                               "If(IsDBNull(aDtr.Item(""" & col.NomColonneOriginal & """)), Nothing, CType(aDtr.Item(""" & col.NomColonneOriginal & """), " & col.TypeDeDonnees & "))" & vbCrLf
                            End If

                        Next

                        uneClasse &= getNumberTab(3) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End While" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        uneClasse &= getNumberTab(2) & "aDtr.Close()" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aDtr = Nothing" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "aCommand = Nothing" & vbCrLf

                        uneClasse &= getNumberTab(2) & "If (AnOpenConneciton Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn.Close()" & vbCrLf
                        uneClasse &= getNumberTab(3) & "aConn = Nothing" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf



                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Return result" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Function" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Création de la fonction ToString() de la classe
                        '************************************************************************************

                        If (toStringValueFound) Then
                            uneClasse &= getNumberTab(1) & "Public Overrides Function ToString() As String" & vbCrLf
                            uneClasse &= getNumberTab(2) & "Return " & toStringFunction & vbCrLf
                            uneClasse &= getNumberTab(1) & "End Function" & vbCrLf
                        End If

                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Création de la fonction Equals de la classe
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "Public Overrides Function Equals(obj As Object) As Boolean" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "dim value as " & leNomDeLaTable.ClassName & " = trycast(obj, " & leNomDeLaTable.ClassName & ")" & vbCrLf

                        uneClasse &= getNumberTab(2) & "If (value Is Nothing) then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "Return False" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Else" & vbCrLf
                        uneClasse &= getNumberTab(3) & Strings.Left(equalsFunction, equalsFunction.Length - 5) & ")" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Function" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'Création de la fonction Save un objet de la classe
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "Public Shared Function Save" & leNomDeLaTable.NomSingulier & "(ByVal _" & leNomDeLaTable.NomSingulier & " As " & leNomDeLaTable.ClassName & optionalConnection & ") As Boolean" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Dim resultat As Boolean = False" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "If (" & leNomDeLaTable.NomSingulier & "Exists(_" & leNomDeLaTable.NomSingulier & ", AnOpenConneciton)) Then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "resultat = (Update" & leNomDeLaTable.NomSingulier & "(_" & leNomDeLaTable.NomSingulier & ", AnOpenConneciton) > 0)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Else" & vbCrLf
                        uneClasse &= getNumberTab(3) & "resultat = (Insert" & leNomDeLaTable.NomSingulier & "(_" & leNomDeLaTable.NomSingulier & ", AnOpenConneciton) > 0)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Return resultat" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Function" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        '************************************************************************************
                        'Création de la fonction Save lui-même
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "Public Function Save() As Boolean" & vbCrLf
                        uneClasse &= getNumberTab(2) & "If (Not BooIsSaved) Then" & vbCrLf
                        uneClasse &= getNumberTab(3) & "BooIsSaved = " & leNomDeLaTable.ClassName & ".Save" & leNomDeLaTable.NomSingulier & "(me)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "End If" & vbCrLf
                        uneClasse &= getNumberTab(2) & "Return BooIsSaved" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Function" & vbCrLf
                        uneClasse &= getNumberTab(2) & "" & vbCrLf

                        uneClasse &= "" & vbCrLf
                        uneClasse &= "#End Region" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                                  E V E N T S                                      *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= "#Region ""Events""" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section privée" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section publique" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        '************************************************************************************
                        'PropertyChanged Event
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        uneClasse &= "#End Region" & vbCrLf

                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'                I N T E R F A C E S  I M P L  E M E N T A T J O N S                *" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'************************************************************************************" & vbCrLf
                        uneClasse &= "#Region ""Interfaces implementations""" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section privée" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- ------" & vbCrLf
                        uneClasse &= "" & vbCrLf


                        '************************************************************************************
                        'OnPropertyChanged function
                        '************************************************************************************

                        uneClasse &= getNumberTab(1) & "' This method is called by the Set accessor of each property." & vbCrLf
                        uneClasse &= getNumberTab(1) & "' The CallerMemberName attribute that is applied to the optional propertyName" & vbCrLf
                        uneClasse &= getNumberTab(1) & "' parameter causes the property name of the caller to be substituted as an argument." & vbCrLf
                        uneClasse &= getNumberTab(1) & "Private Sub OnPropertyChanged(<CallerMemberName()> Optional ByVal propertyName As String = Nothing)" & vbCrLf
                        uneClasse &= getNumberTab(2) & "    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))" & vbCrLf
                        uneClasse &= getNumberTab(1) & "End Sub" & vbCrLf
                        uneClasse &= "" & vbCrLf

                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'Section publique" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= getNumberTab(1) & "'------- --------" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= "#End Region" & vbCrLf
                        uneClasse &= "" & vbCrLf
                        uneClasse &= "End Class" & vbCrLf

                        listeDeClasses.Add(New ClassCodeVb(leNomDeLaTable.TableName, uneClasse))
                        result += 1

                    End If

                Next

            Next

            Return result
        Catch ex As Exception
            MsgBox(ex.Message)
            Return -1
        End Try

    End Function

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

    Protected Friend Sub OnPropertyChanged(<CallerMemberName()> Optional ByVal propertyName As String = Nothing)
        'MsgBox(propertyName & vbCrLf & ViewModelPeriodeDePaie.instance & vbCrLf & "Lecture :" & Employes.LectureDePropriete & vbCrLf & "Ecriture :" & Employes.EcriturePropriete & vbCrLf & "Fonction :" & ViewModelPeriodeDePaie.AppelDeFonction)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Friend Sub CommandAbout(ByVal params As Object)

        Dim bmpImg As New BitmapImage(New Uri("pack://application:,,,/Images/LogoExacad.png"))

        Dim about As New AboutDialogBox.AboutDialogViewModel(My.Application.Info,
                                                             "EULA",
                                                             "",
                                                             "http:\\www.exacad.com",
                                                             bmpImg,
                                                             "BaseDark",
                                                             "Crimson")

        about.ShowDialog()

        about.Dispose()

    End Sub

    Protected Friend Sub CommandSettings(ByVal params As Object)
        Dim pref As New Preferences
        pref.DataContext = New ViewModelPreferences()
        pref.ShowDialog()
        pref.Close()
        pref = Nothing
    End Sub

    '------- --------
    '------- --------
    'Section publique
    '------- --------
    '------- --------

    Default Public ReadOnly Property Item(columnName As String) As String Implements IDataErrorInfo.Item
        Get

            Dim resultat As String

            Select Case columnName
                Case "TypeBaseDonnees"
                    If (Me.TypeBaseDonnees <= 0) Then
                        resultat = "Vous devez sélectionner un type de base de données!"
                    Else
                        resultat = String.Empty
                    End If

                Case "ServerAddresseName"
                    If (Not Me.ServerAddresseName Is Nothing) Then
                        If (Me.ServerAddresseName = String.Empty) Then
                            resultat = "Vous devez entrer le nom du serveur ou son adresse IP."
                        Else
                            resultat = String.Empty
                        End If
                    Else
                        resultat = "Vous devez entrer le nom du serveur ou son adresse IP."
                    End If

                Case "TCPPort"
                    If (Me.TypeBaseDonnees = databaseType.MYSQL) Then
                        If (IsNumeric(TCPPort)) Then
                            If (InStr(TCPPort, ",") <> 0) Or (InStr(TCPPort, ".") <> 0) Then
                                resultat = "Le port TCP ne supporte que des entiers (Sans décimales)."
                            ElseIf ((CLng(TCPPort) < 0) Or (CLng(TCPPort) > 65535)) Then
                                resultat = "Le port TCP ne supporte que des entier compris entre 0 et 65535."
                            Else
                                resultat = String.Empty
                            End If
                        Else
                            resultat = "Le port TCP doit être numérique."
                        End If
                    Else
                        resultat = String.Empty
                    End If
                Case "DatabaseCatalog"
                    resultat = String.Empty
                Case "Username"
                    If (TypeBaseDonnees = databaseType.SQL_SERVER) Or (TypeBaseDonnees = databaseType.ORACLE) Then
                        If (Not TrustedConnection) And (Username = String.Empty) Then
                            resultat = "Le nom d'usager est un champ obligatoire."
                        Else
                            resultat = String.Empty
                        End If
                    ElseIf (Username = String.Empty) Then
                        resultat = "Le nom d'usager est un champ obligatoire."
                    Else
                        resultat = String.Empty
                    End If

                Case "Password"
                    If (TypeBaseDonnees = databaseType.SQL_SERVER) Or (TypeBaseDonnees = databaseType.ORACLE) Then
                        If (Not TrustedConnection) And (strPassword = String.Empty) Then
                            resultat = "Le nom d'usager est un champ obligatoire."
                        Else
                            resultat = String.Empty
                        End If
                    ElseIf (strPassword = String.Empty) Then
                        resultat = "Le nom d'usager est un champ obligatoire."
                    Else
                        resultat = String.Empty
                    End If

                Case "TrustedConnection"
                    resultat = String.Empty
                Case "SelectedDB"
                    If ((TypeBaseDonnees = databaseType.SQL_SERVER) Or
                        (TypeBaseDonnees = databaseType.ORACLE) Or
                        (TypeBaseDonnees = databaseType.MYSQL) Or
                        (TypeBaseDonnees = databaseType.POSTGRE_SQL)) Then

                        If (SelectedDB <> "Veuillez sélectionner une BD") Then
                            resultat = String.Empty
                        Else
                            resultat = ""
                        End If
                    Else
                        resultat = String.Empty
                    End If

                Case "listeTables"
                    resultat = String.Empty
                Case Else
                    resultat = String.Empty
            End Select

            Return resultat

        End Get
    End Property

    Public ReadOnly Property [Error] As String Implements IDataErrorInfo.Error
        Get
            Return Nothing
        End Get
    End Property

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

    '************************************************************************************
    '            P R I V A T E  C L A S S E S  I M P L E M E N T A T I O N              *
    '************************************************************************************
#Region "Private Classes Implementations"

    '------- ------
    '------- ------
    'Section privée
    '------- ------
    '------- ------

    Private Class TableName

        Private tblName As String
        Private nomSingulierTable As String = ""
        Private lstReservedWords As New List(Of String)


        Public Sub New(ByVal name As String)

            tblName = name


            tblName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(tblName.Replace("_", " ").ToLower).Replace(" ", "")
            If (Strings.Left(tblName, 3).ToUpper = "TBL") Then
                tblName = Strings.Right(tblName, tblName.Length - 3)
            End If

            nomSingulierTable = If(Strings.Right(tblName, 1).ToUpper = "S", Strings.Left(tblName, tblName.Length - 1), tblName)

            lstReservedWords.Add("AddHandler".ToUpper)
            lstReservedWords.Add("AddressOf".ToUpper)
            lstReservedWords.Add("Alias".ToUpper)
            lstReservedWords.Add("And".ToUpper)
            lstReservedWords.Add("AndAlso".ToUpper)
            lstReservedWords.Add("As".ToUpper)
            lstReservedWords.Add("Boolean".ToUpper)
            lstReservedWords.Add("ByRef".ToUpper)
            lstReservedWords.Add("Byte".ToUpper)
            lstReservedWords.Add("ByVal".ToUpper)
            lstReservedWords.Add("Call".ToUpper)
            lstReservedWords.Add("Case".ToUpper)
            lstReservedWords.Add("Catch".ToUpper)
            lstReservedWords.Add("CBool".ToUpper)
            lstReservedWords.Add("CByte".ToUpper)
            lstReservedWords.Add("CChar".ToUpper)
            lstReservedWords.Add("CDate".ToUpper)
            lstReservedWords.Add("CDbl".ToUpper)
            lstReservedWords.Add("CDec".ToUpper)
            lstReservedWords.Add("Char".ToUpper)
            lstReservedWords.Add("CInt".ToUpper)
            lstReservedWords.Add("Class".ToUpper)
            lstReservedWords.Add("CLng".ToUpper)
            lstReservedWords.Add("CObj".ToUpper)
            lstReservedWords.Add("Const".ToUpper)
            lstReservedWords.Add("Continue".ToUpper)
            lstReservedWords.Add("CSByte".ToUpper)
            lstReservedWords.Add("CShort".ToUpper)
            lstReservedWords.Add("CSng".ToUpper)
            lstReservedWords.Add("CStr".ToUpper)
            lstReservedWords.Add("CType".ToUpper)
            lstReservedWords.Add("CUInt".ToUpper)
            lstReservedWords.Add("CULng".ToUpper)
            lstReservedWords.Add("CUShort".ToUpper)
            lstReservedWords.Add("Date".ToUpper)
            lstReservedWords.Add("Decimal".ToUpper)
            lstReservedWords.Add("Declare".ToUpper)
            lstReservedWords.Add("Default".ToUpper)
            lstReservedWords.Add("Delegate".ToUpper)
            lstReservedWords.Add("Dim".ToUpper)
            lstReservedWords.Add("DirectCast".ToUpper)
            lstReservedWords.Add("Do".ToUpper)
            lstReservedWords.Add("Double".ToUpper)
            lstReservedWords.Add("Each".ToUpper)
            lstReservedWords.Add("Else".ToUpper)
            lstReservedWords.Add("ElseIf".ToUpper)
            lstReservedWords.Add("End".ToUpper)
            lstReservedWords.Add("EndIf".ToUpper)
            lstReservedWords.Add("Enum".ToUpper)
            lstReservedWords.Add("Erase".ToUpper)
            lstReservedWords.Add("Error".ToUpper)
            lstReservedWords.Add("Event".ToUpper)
            lstReservedWords.Add("Exit".ToUpper)
            lstReservedWords.Add("Finally".ToUpper)
            lstReservedWords.Add("For".ToUpper)
            lstReservedWords.Add("Friend".ToUpper)
            lstReservedWords.Add("Function".ToUpper)
            lstReservedWords.Add("Get".ToUpper)
            lstReservedWords.Add("GetType".ToUpper)
            lstReservedWords.Add("GetXmlNamespace.ToUpper")
            lstReservedWords.Add("Global".ToUpper)
            lstReservedWords.Add("GoSub".ToUpper)
            lstReservedWords.Add("GoTo".ToUpper)
            lstReservedWords.Add("Handles".ToUpper)
            lstReservedWords.Add("If".ToUpper)
            lstReservedWords.Add("Implements.ToUpper")
            lstReservedWords.Add("Imports".ToUpper)
            lstReservedWords.Add("In".ToUpper)
            lstReservedWords.Add("Inherits".ToUpper)
            lstReservedWords.Add("Integer".ToUpper)
            lstReservedWords.Add("Interface".ToUpper)
            lstReservedWords.Add("Is".ToUpper)
            lstReservedWords.Add("IsNot".ToUpper)
            lstReservedWords.Add("Let".ToUpper)
            lstReservedWords.Add("Lib".ToUpper)
            lstReservedWords.Add("Like".ToUpper)
            lstReservedWords.Add("Long".ToUpper)
            lstReservedWords.Add("Loop".ToUpper)
            lstReservedWords.Add("Me".ToUpper)
            lstReservedWords.Add("Mod".ToUpper)
            lstReservedWords.Add("Module".ToUpper)
            lstReservedWords.Add("MustInherit".ToUpper)
            lstReservedWords.Add("MustOverride".ToUpper)
            lstReservedWords.Add("MyBase".ToUpper)
            lstReservedWords.Add("MyClass".ToUpper)
            lstReservedWords.Add("Namespace".ToUpper)
            lstReservedWords.Add("Narrowing".ToUpper)
            lstReservedWords.Add("New".ToUpper)
            lstReservedWords.Add("Next".ToUpper)
            lstReservedWords.Add("Not".ToUpper)
            lstReservedWords.Add("Nothing".ToUpper)
            lstReservedWords.Add("NotInheritable".ToUpper)
            lstReservedWords.Add("NotOverridable".ToUpper)
            lstReservedWords.Add("Object".ToUpper)
            lstReservedWords.Add("Of".ToUpper)
            lstReservedWords.Add("On".ToUpper)
            lstReservedWords.Add("Operator".ToUpper)
            lstReservedWords.Add("Option".ToUpper)
            lstReservedWords.Add("Optional".ToUpper)
            lstReservedWords.Add("Or".ToUpper)
            lstReservedWords.Add("OrElse".ToUpper)
            lstReservedWords.Add("Overloads".ToUpper)
            lstReservedWords.Add("Overridable".ToUpper)
            lstReservedWords.Add("Overrides".ToUpper)
            lstReservedWords.Add("ParamArray".ToUpper)
            lstReservedWords.Add("Partial".ToUpper)
            lstReservedWords.Add("Private".ToUpper)
            lstReservedWords.Add("Property".ToUpper)
            lstReservedWords.Add("Protected".ToUpper)
            lstReservedWords.Add("Public".ToUpper)
            lstReservedWords.Add("RaiseEvent".ToUpper)
            lstReservedWords.Add("ReadOnly".ToUpper)
            lstReservedWords.Add("ReDim".ToUpper)
            lstReservedWords.Add("REM".ToUpper)
            lstReservedWords.Add("RemoveHandler".ToUpper)
            lstReservedWords.Add("Resume".ToUpper)
            lstReservedWords.Add("Return".ToUpper)
            lstReservedWords.Add("SByte".ToUpper)
            lstReservedWords.Add("Select".ToUpper)
            lstReservedWords.Add("Set".ToUpper)
            lstReservedWords.Add("Shadows".ToUpper)
            lstReservedWords.Add("Shared".ToUpper)
            lstReservedWords.Add("Short".ToUpper)
            lstReservedWords.Add("Single".ToUpper)
            lstReservedWords.Add("Static".ToUpper)
            lstReservedWords.Add("Step".ToUpper)
            lstReservedWords.Add("Stop".ToUpper)
            lstReservedWords.Add("String".ToUpper)
            lstReservedWords.Add("Structure".ToUpper)
            lstReservedWords.Add("Sub".ToUpper)
            lstReservedWords.Add("SyncLock".ToUpper)
            lstReservedWords.Add("Then".ToUpper)
            lstReservedWords.Add("Throw".ToUpper)
            lstReservedWords.Add("To".ToUpper)
            lstReservedWords.Add("Try".ToUpper)
            lstReservedWords.Add("TryCast".ToUpper)
            lstReservedWords.Add("TypeOf".ToUpper)
            lstReservedWords.Add("UInteger".ToUpper)
            lstReservedWords.Add("ULong".ToUpper)
            lstReservedWords.Add("UShort".ToUpper)
            lstReservedWords.Add("Using".ToUpper)
            lstReservedWords.Add("Variant".ToUpper)
            lstReservedWords.Add("Wend".ToUpper)
            lstReservedWords.Add("When".ToUpper)
            lstReservedWords.Add("While".ToUpper)
            lstReservedWords.Add("Widening".ToUpper)
            lstReservedWords.Add("With".ToUpper)
            lstReservedWords.Add("WithEvents".ToUpper)
            lstReservedWords.Add("WriteOnly".ToUpper)
            lstReservedWords.Add("Xor".ToUpper)
            lstReservedWords.Add("FALSE".ToUpper)
            lstReservedWords.Add("TRUE".ToUpper)

        End Sub

        Public Property TableName As String
            Get
                Return tblName
            End Get
            Set(value As String)
                tblName = value
            End Set
        End Property

        Public ReadOnly Property ClassName As String
            Get
                Return IfReservedWord(Me.UppercaseFirstLetter(tblName))
            End Get
        End Property

        Public ReadOnly Property NomSingulier As String
            Get
                Return nomSingulierTable
            End Get
        End Property


        Private Function UppercaseFirstLetter(ByVal val As String) As String
            ' Test for nothing or empty.
            If String.IsNullOrEmpty(val) Then
                Return val
            End If

            ' Convert to character array.
            Dim array() As Char = val.ToCharArray

            ' Uppercase first character.
            array(0) = Char.ToUpper(array(0))

            ' Return new string.
            Return New String(array)
        End Function

        Private Function IfReservedWord(ByVal className As String) As String
            If (lstReservedWords.Contains(className.ToUpper)) Then
                Return "[" & className & "]"
            Else
                Return className
            End If
        End Function

    End Class

    '//////////////////////////////////////////////////
    Private Class UneColonne

        Public Property NomColonneOriginal As String = ""
        Public Property VarDeClasse As String = ""
        Public Property VarDePropriete As String = ""
        Public Property VarDeConstructeur As String = ""
        Public Property TypeDeDonnees As String = ""
        Public Property IsPrimaryKey As Boolean = False

        Public Sub New(ByVal variableColOriginal As String,
                       ByVal variableDeClasse As String,
                       ByVal variableDePropriete As String,
                       ByVal variableTypeDonnees As String,
                       ByVal _isPrimaryKey As Boolean)
            NomColonneOriginal = variableColOriginal
            VarDeClasse = variableDeClasse
            VarDePropriete = variableDePropriete
            VarDeConstructeur = "_" & variableDePropriete
            TypeDeDonnees = variableTypeDonnees
            IsPrimaryKey = _isPrimaryKey

        End Sub

    End Class

    '------- --------
    '------- --------
    'Section publique
    '------- --------
    '------- --------

#End Region

End Class
