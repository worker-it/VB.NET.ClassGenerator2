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

Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports ClassGenerator2.Debugging
Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Globalization
Imports System.Windows.Forms
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports MySql.Data.MySqlClient
Imports Npgsql

#End Region


Public Class ConnectionInfos
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
    Private disposed As Boolean = False
    Private BooIsSaved As Boolean = True

    'Classe Variables

    Private IntId As Integer
    Private StrConnectionName As String
    Private IntTypeBaseDonnees As Long
    Private StrServerAddressName As String
    Private IntTCPPort As Integer
    Private IntDatabaseCatalog As Long = 0
    Private strUsername As String
    Private strPassword As String
    Private booTrustedConnection As Boolean
    Private StrSelectedDb As String

    Private vmmw As ViewModelMainWindow

    Private lstListeSchemas As ObservableCollection(Of TreeView.Noeud)
    Private listeDeClasses As List(Of ClassCodeVb)
    Private LstListeDesBaseDeDonnees As New ObservableCollection(Of String)

    Private strAnEventLogger As New EventLogger("C:\temp", "ClassGenerator.log", "MainWindow", ".", "MainWindow", True)

    '------- --------
    '------- --------
    'Section publique
    '------- --------
    '------- --------

    Public Enum databaseType
        NONE = 0
        SQL_SERVER = 1
        SQL_SERVER_LOCALDB = 2
        ORACLE = 3
        MYSQL = 4
        POSTGRE_SQL = 5
        MS_ACCESS_97_2003 = 6
        MS_ACCESS_2007_2019 = 7
        MS_EXCEL = 8
        FLAT_FILE = 9
    End Enum

    Public Const OLEDB_CONN_STRING = "Server=(LocalDb)\MSSQLLocalDB;Database=ConnectionListDB;Uid=ConnectionList;Pwd=X043rMiVpOlbAlGT9UKZ;"

#End Region
    '************************************************************************************
    '                    C  O  N  S  T  R  U  C  T  E  U  R                             *
    '                    ----------------------------------                             *
    '                      D  E  S  T  R  U  C  T  E  U  R                              *
    '************************************************************************************
#Region "Constructors"

    Public Sub New(ByRef _vmmw As ViewModelMainWindow)

        ' This call is required by the designer.
        'InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        vmmw = _vmmw

        StrConnectionName = ""
        IntTypeBaseDonnees = 0
        StrServerAddressName = ""
        IntTCPPort = 0
        IntDatabaseCatalog = 0
        strUsername = ""
        strPassword = ""
        booTrustedConnection = False
        StrSelectedDb = ""

        LstListeDesBaseDeDonnees.Add("Charger liste des BDs")

    End Sub

    ''' <summary>
    ''' Constructeur de base avec paramètres
    ''' </summary>
    Public Sub New(ByVal _Id As Integer,
                   ByVal _ConnectionName As String,
                   ByVal _TypeBaseDonnees As Integer,
                   ByVal _ServerAddressName As String,
                   ByVal _TcpPort As Integer,
                   ByVal _DatabaseCatalog As Integer,
                   ByVal _SelectedDb As String,
                   ByVal _Username As String,
                   ByVal _Password As String,
                   ByVal _TrustedConnection As Boolean,
                   ByRef _vmmw As ViewModelMainWindow)

        ' This call is required by the designer.
        'InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        IntId = _Id
        StrConnectionName = _ConnectionName
        IntTypeBaseDonnees = _TypeBaseDonnees
        StrServerAddressName = _ServerAddressName
        IntTCPPort = _TcpPort
        IntDatabaseCatalog = _DatabaseCatalog
        StrSelectedDb = _SelectedDb
        strUsername = _Username
        strPassword = _Password
        booTrustedConnection = _TrustedConnection

        LstListeDesBaseDeDonnees.Add("Charger liste des BDs")
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

    Public Property Id As Integer
        Get
            Return IntId
        End Get
        Set(value As Integer)
            IntId = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property ConnectionName As String
        Get
            Return StrConnectionName
        End Get
        Set(value As String)
            StrConnectionName = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property TypeBaseDonnees As Long
        Get
            Return IntTypeBaseDonnees
        End Get
        Set(value As Long)
            IntTypeBaseDonnees = value
            Select Case IntTypeBaseDonnees
                Case databaseType.SQL_SERVER, databaseType.SQL_SERVER_LOCALDB
                    TCPPort = 1433
                Case databaseType.POSTGRE_SQL
                    TCPPort = 5432
                Case databaseType.MYSQL
                    TCPPort = 3306
                Case databaseType.ORACLE
                    TCPPort = 1521
                Case Else
                    TCPPort = -1
            End Select
            DatabaseCatalog = 0
            BooIsSaved = False
            vmmw.OnPropertyChanged("BrowseButtonVisibility")
            vmmw.OnPropertyChanged("TCPPortVisibility")
            OnPropertyChanged()
        End Set
    End Property

    Public Property ServerAddressName As String
        Get
            Return StrServerAddressName
        End Get
        Set(value As String)
            StrServerAddressName = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property TCPPort As Integer
        Get
            Return IntTCPPort
        End Get
        Set(value As Integer)
            IntTCPPort = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property DatabaseCatalog As Long
        Get
            Return IntDatabaseCatalog
        End Get
        Set(value As Long)
            IntDatabaseCatalog = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property Username As String
        Get
            Return strUsername
        End Get
        Set(value As String)
            strUsername = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property TrustedConnection As Boolean
        Get
            Return booTrustedConnection
        End Get
        Set(value As Boolean)
            booTrustedConnection = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property SelectedDB As String
        Get
            Return StrSelectedDb
        End Get
        Set(value As String)
            StrSelectedDb = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property ListeDesBaseDeDonnees As ObservableCollection(Of String)
        Get
            Return LstListeDesBaseDeDonnees
        End Get
        Set(value As ObservableCollection(Of String))
            LstListeDesBaseDeDonnees = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property ListeTablesEtChamps As ObservableCollection(Of TreeView.Noeud)
        Get
            If (lstListeSchemas Is Nothing) Then
                lstListeSchemas = New ObservableCollection(Of TreeView.Noeud)
            End If
            Return lstListeSchemas
        End Get
        Set(value As ObservableCollection(Of TreeView.Noeud))
            lstListeSchemas = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property ListeSchemas As ObservableCollection(Of TreeView.Noeud)
        Get
            Return lstListeSchemas
        End Get
        Set(value As ObservableCollection(Of TreeView.Noeud))
            lstListeSchemas = value
            BooIsSaved = False
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

    Public Sub BtnRetrieveDbsCommand(ByVal params As Object)
        Try

            Dim pwdbox As PasswordBox = TryCast(params, PasswordBox)
            strPassword = pwdbox.Password
            pwdbox = Nothing

            If (strPassword <> String.Empty) Then

                strAnEventLogger.writeLog("Initiation de la recherche des bases de données", "", EventLogEntryType.Information)
                If (Me.ServerAddressName <> "") And (((Me.Username <> "") And (Me.strPassword <> "")) Or (Me.TrustedConnection)) Then

                    strAnEventLogger.writeLog("Si les infos de bases de données sont bien entrées.", "", EventLogEntryType.Information)

                    Dim connString As String = ""
                    Dim commString As String = ""

                    Dim connection As DbConnection
                    Dim commande As DbCommand
                    Dim reader As DbDataReader

                    strAnEventLogger.writeLog("Vide la collection de bases de données.", "", EventLogEntryType.Information)
                    Me.ListeDesBaseDeDonnees.Clear()

                    Select Case Me.TypeBaseDonnees
                        Case databaseType.SQL_SERVER, databaseType.SQL_SERVER_LOCALDB
                            Try

                                strAnEventLogger.writeLog("MS SQL Server.", "", EventLogEntryType.Information)

                                commString = "SELECT name FROM master.sys.databases;"

                                If (Me.TypeBaseDonnees = databaseType.SQL_SERVER_LOCALDB) Then

                                    connection = New SqlClient.SqlConnection(buildConnectionString(False))

                                Else

                                    connection = New OleDbConnection(buildConnectionString(False))

                                End If

                                strAnEventLogger.writeLog("Exécution de la commande.", "", EventLogEntryType.Information)
                                connection.Open()

                                strAnEventLogger.writeLog("Definition de la commande.", "", EventLogEntryType.Information)
                                commande = connection.CreateCommand()
                                commande.CommandText = commString
                                commande.CommandType = System.Data.CommandType.Text

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

                                IntDatabaseCatalog = 0
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

                            IntDatabaseCatalog = 0
                            OnPropertyChanged("DatabaseCatalog")

                        Case databaseType.POSTGRE_SQL

                            Try

                                strAnEventLogger.writeLog("PostgreSQL.", "", EventLogEntryType.Information)

                                commString = "SELECT datname FROM pg_database;"

                                strAnEventLogger.writeLog("Connection MS SQL username and password.", "", EventLogEntryType.Information)
                                connection = New NpgsqlConnection(buildConnectionString(False))

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

                                IntDatabaseCatalog = 0
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
                                connection = New MySqlConnection(buildConnectionString(False))

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

                                IntDatabaseCatalog = 0
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

            OnPropertyChanged("DatabaseCatalog")
            OnPropertyChanged("ListeDesBaseDeDonnees")

        Catch ex As Exception

            strAnEventLogger.writeLog(ex.Message, ex.StackTrace, EventLogEntryType.Error)

            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub RetrieveDBInfosCommand(ByVal params As Object)

        'Déclaration de la variable pour la connection      
        Dim cnxSchemas As DbConnection
        'Déclaration de la variable pour la connection      
        Dim cnxTables As DbConnection
        'Déclaration de la variable pour la connection      
        Dim cnxColonnes As DbConnection
        'Déclaration de la variable pour la connectionstring      
        Dim cnxstr As String
        'Déclaration de la variable pour la requête      
        Dim sqlSchemas As String
        'Déclaration de la variable pour la requête      
        Dim sqlTables As String
        'Déclaration de la variable pour la requête      
        Dim sqlColonnes As String
        'Déclaration de la variable pour la commande       
        Dim cmdSchemas As DbCommand
        'Déclaration de la variable pour la commande       
        Dim cmdTables As DbCommand
        'Déclaration de la variable pour le dataadapter
        Dim dtrSchemas As DbDataReader
        'Déclaration de la variable pour le dataadapter
        Dim dtrTables As DbDataReader
        'Déclaration de la variable pour la commande       
        Dim cmdColonnes As DbCommand
        'Déclaration de la variable pour le dataadapter
        Dim dtrColonnes As DbDataReader

        listeDeClasses = New List(Of ClassCodeVb)

        'ouverture de la connection (à partir du répertoire de l'application) sur la même ligne      
        cnxstr = buildConnectionString()

        Dim principale As New ObservableCollection(Of TreeView.Noeud)
        Dim lst As ObservableCollection(Of TreeView.Noeud)
        Dim Schemas As New ObservableCollection(Of TreeView.Noeud)

        Select Case Me.TypeBaseDonnees
            Case databaseType.SQL_SERVER


                cnxSchemas = New OleDbConnection
                cnxSchemas.ConnectionString = cnxstr
                cnxSchemas.Open()

                'Création de la requête sql pour les schemas
                sqlSchemas = "SELECT " & Me.SelectedDB & ".INFORMATION_SCHEMA.SCHEMATA.SCHEMA_NAME " &
                             "FROM " & Me.SelectedDB & ".INFORMATION_SCHEMA.SCHEMATA " &
                             "WHERE (" & Me.SelectedDB & ".INFORMATION_SCHEMA.SCHEMATA.SCHEMA_NAME<>'guest') AND " &
                                   "(" & Me.SelectedDB & ".INFORMATION_SCHEMA.SCHEMATA.SCHEMA_NAME<>'sys') AND " &
                                   "(left(" & Me.SelectedDB & ".information_schema.schemata.schema_name,3)<>'db_') and " &
                                   "(" & Me.SelectedDB & ".INFORMATION_SCHEMA.SCHEMATA.SCHEMA_NAME<>'INFORMATION_SCHEMA') " &
                             "ORDER BY 1;"

                'Création de la commande et on l'instancie (sql)       
                cmdSchemas = New OleDbCommand(sqlSchemas, CType(cnxSchemas, OleDbConnection))

                'Création du datareader (dta)
                dtrSchemas = cmdSchemas.ExecuteReader()

                While dtrSchemas.Read()

                    lst = New ObservableCollection(Of TreeView.Noeud)

                    cnxTables = New OleDbConnection
                    cnxTables.ConnectionString = cnxstr
                    cnxTables.Open()

                    sqlTables = "SELECT " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_NAME " &
                                "FROM " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES " &
                                "WHERE " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_SCHEMA='" & dtrSchemas.Item("SCHEMA_NAME").ToString() & "' " &
                                "ORDER BY 1;"

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
                                        "WHERE COLUMNS.TABLE_SCHEMA='" & dtrSchemas.Item("SCHEMA_NAME").ToString() & "' AND " &
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

                        lst.Add(New DbTable(dtrTables.Item("TABLE_NAME").ToString(), colonnes))

                        dtrColonnes.Close()
                        dtrColonnes = Nothing

                        cmdColonnes = Nothing
                        cnxColonnes.Close()
                        cnxColonnes = Nothing

                    End While

                    Schemas.Add(New DbSchemas(dtrSchemas.Item("SCHEMA_NAME"), lst))

                End While

            Case databaseType.SQL_SERVER_LOCALDB


                cnxSchemas = New SqlClient.SqlConnection
                cnxSchemas.ConnectionString = cnxstr
                cnxSchemas.Open()

                'Création de la requête sql pour les schemas
                sqlSchemas = "SELECT " & Me.SelectedDB & ".INFORMATION_SCHEMA.SCHEMATA.SCHEMA_NAME " &
                             "FROM " & Me.SelectedDB & ".INFORMATION_SCHEMA.SCHEMATA " &
                             "WHERE (" & Me.SelectedDB & ".INFORMATION_SCHEMA.SCHEMATA.SCHEMA_NAME<>'guest') AND " &
                                   "(" & Me.SelectedDB & ".INFORMATION_SCHEMA.SCHEMATA.SCHEMA_NAME<>'sys') AND " &
                                   "(left(" & Me.SelectedDB & ".information_schema.schemata.schema_name,3)<>'db_') and " &
                                   "(" & Me.SelectedDB & ".INFORMATION_SCHEMA.SCHEMATA.SCHEMA_NAME<>'INFORMATION_SCHEMA') " &
                             "ORDER BY 1;"

                'Création de la commande et on l'instancie (sql)       
                cmdSchemas = New SqlClient.SqlCommand(sqlSchemas, CType(cnxSchemas, SqlClient.SqlConnection))

                'Création du datareader (dta)
                dtrSchemas = cmdSchemas.ExecuteReader()

                While dtrSchemas.Read()

                    lst = New ObservableCollection(Of TreeView.Noeud)

                    cnxTables = New SqlClient.SqlConnection
                    cnxTables.ConnectionString = cnxstr
                    cnxTables.Open()

                    sqlTables = "SELECT " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_NAME " &
                                "FROM " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES " &
                                "WHERE " & Me.SelectedDB & ".INFORMATION_SCHEMA.TABLES.TABLE_SCHEMA='" & dtrSchemas.Item("SCHEMA_NAME").ToString() & "' " &
                                "ORDER BY 1;"

                    'Création de la commande et on l'instancie (sql)       
                    cmdTables = New SqlClient.SqlCommand(sqlTables, CType(cnxTables, SqlClient.SqlConnection))

                    'Création du datareader (dta)
                    dtrTables = cmdTables.ExecuteReader()

                    While dtrTables.Read()

                        Dim colonnes As New ObservableCollection(Of TreeView.Noeud)

                        cnxColonnes = New SqlClient.SqlConnection
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
                                        "WHERE COLUMNS.TABLE_SCHEMA='" & dtrSchemas.Item("SCHEMA_NAME").ToString() & "' AND " &
                                            "COLUMNS.TABLE_NAME='" & dtrTables.Item("TABLE_NAME").ToString() & "';"

                        'Création de la commande et on l'instancie (sql)       
                        cmdColonnes = New SqlClient.SqlCommand(sqlColonnes, CType(cnxColonnes, SqlClient.SqlConnection))

                        'Création du datareader (dta)
                        dtrColonnes = cmdColonnes.ExecuteReader()

                        While dtrColonnes.Read()

                            Dim uneColonne As New DbChampTable(dtrColonnes.Item("COLUMN_NAME").ToString, dtrColonnes.Item("DATA_TYPE").ToString)

                            uneColonne.IsPrimaryKey = (dtrColonnes.Item("CONSTRAINT_TYPE").ToString().ToUpper = "PRIMARY KEY")

                            colonnes.Add(uneColonne)
                        End While

                        lst.Add(New DbTable(dtrTables.Item("TABLE_NAME").ToString(), colonnes))

                        dtrColonnes.Close()
                        dtrColonnes = Nothing

                        cmdColonnes = Nothing
                        cnxColonnes.Close()
                        cnxColonnes = Nothing

                    End While

                    Schemas.Add(New DbSchemas(dtrSchemas.Item("SCHEMA_NAME"), lst))

                End While

            Case databaseType.ORACLE
            Case databaseType.MYSQL

                lst = New ObservableCollection(Of TreeView.Noeud)

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
                    sqlColonnes = "SHOW COLUMNS FROM " & dtrTables.Item("Tables_in_" & Me.SelectedDB).ToString()

                    'Création de la commande et on l'instancie (sql)       
                    cmdColonnes = New MySqlCommand(sqlColonnes, CType(cnxColonnes, MySqlConnection))

                    'Création du datareader (dta)
                    dtrColonnes = cmdColonnes.ExecuteReader()

                    While dtrColonnes.Read()

                        Dim uneColonne As New DbChampTable(dtrColonnes.Item("Field").ToString, dtrColonnes.Item("Type").ToString)

                        uneColonne.IsPrimaryKey = (dtrColonnes.Item("Key").ToString().ToUpper = "PRI")

                        colonnes.Add(uneColonne)
                    End While

                    lst.Add(New DbTable(dtrTables.Item("Tables_in_" & Me.SelectedDB).ToString(), colonnes))

                    dtrColonnes.Close()
                    dtrColonnes = Nothing

                    cmdColonnes = Nothing
                    cnxColonnes.Close()
                    cnxColonnes = Nothing

                End While

                Schemas.Add(New DbSchemas(Me.SelectedDB, lst))

            Case databaseType.POSTGRE_SQL

                cnxSchemas = New NpgsqlConnection
                cnxSchemas.ConnectionString = cnxstr
                cnxSchemas.Open()

                'création de la requête sql pour les schemas
                sqlSchemas = "SELECT " & Me.SelectedDB & ".information_schema.schemata.schema_name " &
                             "FROM " & Me.SelectedDB & ".information_schema.schemata " &
                             "WHERE (left(" & Me.SelectedDB & ".information_schema.schemata.schema_name,3)<>'sys') and " &
                                   "(left(" & Me.SelectedDB & ".information_schema.schemata.schema_name,3)<>'pg_') and " &
                                   "" & Me.SelectedDB & ".information_schema.schemata.schema_name<>'information_schema' " &
                             "ORDER BY 1;"

                'Création de la commande et on l'instancie (sql)       
                cmdSchemas = New NpgsqlCommand(sqlSchemas, CType(cnxSchemas, NpgsqlConnection))

                'Création du datareader (dta)
                dtrSchemas = cmdSchemas.ExecuteReader()

                While dtrSchemas.Read()

                    lst = New ObservableCollection(Of TreeView.Noeud)

                    cnxTables = New NpgsqlConnection()
                    cnxTables.ConnectionString = cnxstr
                    cnxTables.Open()

                    'Création de la requête sql      
                    sqlTables = "SELECT schemaname, tablename as object_name, table_type " &
                                "FROM """ & Me.SelectedDB & """.""pg_catalog"".""pg_tables"" " &
                                "WHERE schemaname='" & dtrSchemas.Item("schema_name").ToString() & "' " &
                                "UNION ALL " &
                                "SELECT schemaname, viewname as object_name, table_type " &
                                "FROM """ & Me.SelectedDB & """.""pg_catalog"".""pg_views"" " &
                                "WHERE schemaname='" & dtrSchemas.Item("schema_name").ToString() & "';"

                    sqlTables = "SELECT table_schema AS schemaname, table_name AS object_name, table_type " &
                                "FROM """ & Me.SelectedDB & """.""information_schema"".""tables"" " &
                                "WHERE table_schema='" & dtrSchemas.Item("schema_name").ToString() & "';"

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

                        lst.Add(New DbTable(dtrTables.Item("object_name").ToString(), colonnes, dtrTables.Item("table_type").ToString().ToUpper() <> "VIEW"))

                        dtrColonnes.Close()
                        dtrColonnes = Nothing

                        cmdColonnes = Nothing
                        cnxColonnes.Close()
                        cnxColonnes = Nothing

                    End While

                    Schemas.Add(New DbSchemas(dtrSchemas.Item("schema_name"), lst))

                End While

            Case databaseType.MS_ACCESS_97_2003
            Case databaseType.MS_ACCESS_2007_2019

            Case databaseType.MS_EXCEL

            Case databaseType.FLAT_FILE

        End Select

        principale.Add(New DbDatabase(Me.SelectedDB, Schemas, False) With {.IsExpanded = True})
        ListeSchemas = principale

        cmdSchemas = Nothing
        If (dtrSchemas IsNot Nothing) Then
            dtrSchemas.Close()
        End If
        dtrSchemas = Nothing

        If (dtrTables IsNot Nothing) Then
            dtrTables.Close()
        End If
        dtrTables = Nothing

        cmdTables = Nothing
        If (cnxTables IsNot Nothing) Then
            cnxTables.Close()
        End If
        cnxTables = Nothing

        OnPropertyChanged("ListeTablesEtChamps")
    End Sub

    Public Async Sub CreateFilesCommand(ByVal params As Object)

        Dim wdw As MetroWindow = TryCast(params, MetroWindow)

        If (wdw IsNot Nothing) Then
            Dim folderBrowser As New FolderBrowserDialog()

            If (CreateClasses() <= 0) Then
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

    ''' <summary>
    ''' Fonction permettant d'insérer dans la table de la classe l'objet passé en paramètre.
    ''' </summary>
    ''' <param name="_ConnectionInfos">Un objet de type ConnectionInfos</param>
    ''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>
    Private Shared Function InsertConnectionInfos(ByVal _ConnectionInfos As ConnectionInfos, Optional ByRef AnOpenConneciton As DbConnection = Nothing) As Integer

        Dim resultat As Integer = 0

        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand

        If (AnOpenConneciton Is Nothing) Then
            'Ouverture de la connection SQL
            aConn = New SqlClient.SqlConnection()
            aConn.ConnectionString = OLEDB_CONN_STRING
            aConn.Open()
        Else
            aConn = AnOpenConneciton
        End If

        'Création de la requête sql
        sqlTables = "INSERT INTO ConnectionListDB.CL_SCHEMA.TBL_CONNECTION_LIST (CONNECTION_NAME, TYPE_BASE_DONNEES, SERVER_ADDRESS_NAME, TCP_PORT, DATABASE_CATALOG, SELECTED_DB, USERNAME, TRUSTED_CONNECTION) VALUES ('" & _ConnectionInfos.ConnectionName & "', " & _ConnectionInfos.TypeBaseDonnees & ", '" & _ConnectionInfos.ServerAddressName & "', " & _ConnectionInfos.TCPPort & ", " & _ConnectionInfos.DatabaseCatalog & ", '" & _ConnectionInfos.SelectedDB & "', '" & _ConnectionInfos.Username & "', " & If(_ConnectionInfos.TrustedConnection, 1, 0) & ");"

        'Création de la commande et on l'exécute
        aCommand = New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))

        resultat = aCommand.ExecuteNonQuery()

        aCommand = Nothing
        If (AnOpenConneciton Is Nothing) Then
            aConn.Close()
            aConn = Nothing
        End If

        Return resultat
    End Function

    ''' <summary>
    ''' Fonction permettant de mettre à jour l'objet passé en paramètre dans la table de la classe s'il existe.
    ''' </summary>
    ''' <param name="_ConnectionInfos">Un objet de type ConnectionInfos</param>
    ''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>
    Private Shared Function UpdateConnectionInfos(ByVal _ConnectionInfos As ConnectionInfos, Optional ByRef AnOpenConneciton As DbConnection = Nothing) As Integer

        Dim resultat As Integer = 0

        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand

        If (AnOpenConneciton Is Nothing) Then
            'Ouverture de la connection SQL
            aConn = New SqlClient.SqlConnection()
            aConn.ConnectionString = OLEDB_CONN_STRING
            aConn.Open()
        Else
            aConn = AnOpenConneciton
        End If

        'Création de la requête sql
        sqlTables = "UPDATE ConnectionListDB.CL_SCHEMA.TBL_CONNECTION_LIST SET CONNECTION_NAME = '" & _ConnectionInfos.ConnectionName & "', TYPE_BASE_DONNEES = " & _ConnectionInfos.TypeBaseDonnees & ", SERVER_ADDRESS_NAME = '" & _ConnectionInfos.ServerAddressName & "', TCP_PORT = " & _ConnectionInfos.TCPPort & ", DATABASE_CATALOG = " & _ConnectionInfos.DatabaseCatalog & ", SELECTED_DB = '" & _ConnectionInfos.SelectedDB & "', USERNAME = '" & _ConnectionInfos.Username & "', TRUSTED_CONNECTION = '" & _ConnectionInfos.TrustedConnection & "' WHERE Id = " & _ConnectionInfos.Id & ";"

        'Création de la commande et on l'exécute
        aCommand = New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))

        resultat = aCommand.ExecuteNonQuery()

        aCommand = Nothing
        If (AnOpenConneciton Is Nothing) Then
            aConn.Close()
            aConn = Nothing
        End If

        Return resultat
    End Function

    ''' <summary>
    ''' Fonction permettant de déterminer si l'objet passé en paramètre existe dans la base de données ou non.
    ''' </summary>
    ''' <param name="_ConnectionInfos">Un objet de type ConnectionInfos</param>
    ''' <returns>Retourne Vrai si l'objet existe dans la base de données ou False s'il n'existe pas.</returns>
    Private Shared Function ConnectionInfosExists(ByVal _ConnectionInfos As ConnectionInfos, Optional ByRef AnOpenConneciton As DbConnection = Nothing) As Boolean

        Dim resultat As Boolean = False

        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand

        If (AnOpenConneciton Is Nothing) Then
            'Ouverture de la connection SQL
            aConn = New SqlClient.SqlConnection()
            aConn.ConnectionString = OLEDB_CONN_STRING
            aConn.Open()
        Else
            aConn = AnOpenConneciton
        End If

        'Création de la requête sql
        sqlTables = "SELECT COUNT(*) AS EST_EXISTANT FROM ConnectionListDB.CL_SCHEMA.TBL_CONNECTION_LIST WHERE Id = " & _ConnectionInfos.Id & ";"

        'Création de la commande et on l'exécute
        aCommand = New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))
        Try
            resultat = (CType(aCommand.ExecuteScalar(), Integer) > 0)
        Catch ex As Exception
            resultat = False
        End Try

        aCommand = Nothing
        If (AnOpenConneciton Is Nothing) Then
            aConn.Close()
            aConn = Nothing
        End If

        Return resultat
    End Function

    ''' <summary>
    ''' Fonction permettant de supprimer l'objet passé en paramètre de la base de données.
    ''' </summary>
    ''' <param name="_ConnectionInfos">Un objet de type ConnectionInfos</param>
    ''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>
    Private Shared Function DeleteConnectionInfos(ByVal _ConnectionInfos As ConnectionInfos, Optional ByRef AnOpenConneciton As DbConnection = Nothing) As Integer

        Dim resultat As Integer = 0

        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand

        If (ConnectionInfosExists(_ConnectionInfos)) Then

            If (AnOpenConneciton Is Nothing) Then
                'Ouverture de la connection SQL
                aConn = New SqlClient.SqlConnection()
                aConn.ConnectionString = OLEDB_CONN_STRING
                aConn.Open()
            Else
                aConn = AnOpenConneciton
            End If

            'Création de la requête sql
            sqlTables = "DELETE FROM ConnectionListDB.CL_SCHEMA.TBL_CONNECTION_LIST WHERE Id = " & _ConnectionInfos.Id & ";"

            'Création de la commande et on l'exécute
            aCommand = New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))
            Try
                resultat = CType(aCommand.ExecuteScalar(), Integer)
            Catch ex As Exception
                resultat = -1
            End Try

            aCommand = Nothing
            If (AnOpenConneciton Is Nothing) Then
                aConn.Close()
                aConn = Nothing
            End If

        Else

            MsgBox("L'objet ""ConnectionInfos"" à supprimer n'existe pas.")
            resultat = -1

        End If

        Return resultat
    End Function

    Private Function buildConnectionString(Optional ByVal WithDatabase As Boolean = True) As String

        Dim cnctstr As String

        Select Case Me.TypeBaseDonnees
            Case databaseType.SQL_SERVER
                If (Me.TrustedConnection) Then
                    cnctstr = "Provider=sqloledb;Server=" & Me.ServerAddressName & If(WithDatabase, ";Database=" & Me.SelectedDB, "") & ";Trusted_Connection=yes;"
                Else
                    cnctstr = "Provider=sqloledb;Server=" & Me.ServerAddressName & If(WithDatabase, ";Database=" & Me.SelectedDB, "") & ";Uid=" & Me.Username & ";Pwd=" & Me.strPassword & ";"
                End If
            Case databaseType.SQL_SERVER_LOCALDB
                If (Me.TrustedConnection) Then
                    cnctstr = "Server=" & Me.ServerAddressName & If(WithDatabase, ";Database=" & Me.SelectedDB, "") & ";Trusted_Connection=yes;"
                Else
                    cnctstr = "Server=" & Me.ServerAddressName & If(WithDatabase, ";Database=" & Me.SelectedDB, "") & ";Uid=" & Me.Username & ";Pwd=" & Me.strPassword & ";"
                End If
            Case databaseType.ORACLE
                If (Me.TrustedConnection) Then
                    cnctstr = "Provider=msdaora;Data Source=" & Me.ServerAddressName & If(TCPPort.ToString() <> "", ";Port=" & Me.TCPPort.ToString(), "") & ";Persist Security Info=False;Integrated Security=Yes;"
                Else
                    cnctstr = "Provider=msdaora;Data Source=" & Me.ServerAddressName & If(TCPPort.ToString() <> "", ";Port=" & Me.TCPPort.ToString(), "") & ";User Id=" & Me.Username & ";Password=" & Me.strPassword & ";Integrated Security=no;"
                End If
            Case databaseType.MYSQL
                cnctstr = "Server=" & Me.ServerAddressName & ";Port=" & Me.IntTCPPort.ToString() & If(WithDatabase, ";Database=" & Me.SelectedDB, "") & ";Uid=" & Me.Username & ";Pwd=" & Me.strPassword & ";"
            Case databaseType.POSTGRE_SQL
                cnctstr = "Server=" & Me.ServerAddressName & If(TCPPort.ToString() <> "", ";Port=" & Me.TCPPort.ToString(), ";Port=5432") & If(WithDatabase, ";Database=" & Me.SelectedDB, "") & ";User Id=" & Me.Username & ";Password=" & Me.strPassword & ";"
            Case databaseType.MS_ACCESS_97_2003
                cnctstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Me.ServerAddressName & ";User Id=admin;Password=;"
            Case databaseType.MS_ACCESS_2007_2019
                cnctstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Me.ServerAddressName & ";Persist Security Info=False;"
            Case databaseType.MS_EXCEL
                'ADO .Net
                cnctstr = "Excel File=" & Me.ServerAddressName & ";"
            Case databaseType.FLAT_FILE
                cnctstr = Me.ServerAddressName
            Case Else
                cnctstr = ""
        End Select

        Return cnctstr

    End Function

    Private Function getTypeFromString(ByVal typeDonnee As String) As String

        Dim convertedDataType As String


        Select Case Me.TypeBaseDonnees
            Case databaseType.SQL_SERVER, databaseType.SQL_SERVER_LOCALDB

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

        'Déclaration de la variable pour la connectionstring      
        Dim cnxstr As String

        Dim result As Integer = 0
        Dim leNomDeLaTable As TableName

        Dim qryDbName As String = ""
        Dim qrySchema As String = ""
        Dim qryTblName As String = ""
        Dim qryOneField As String = ""

        Try

            'ouverture de la connection (à partir du répertoire de l'application) sur la même ligne      
            cnxstr = buildConnectionString()

            For Each db As DbDatabase In ListeSchemas
                For Each Schem As DbSchemas In db.Childrens
                    For Each tbl As DbTable In Schem.Childrens

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

                            'uneClasse &= getNumberTab(1) & "''' <summary>" & vbCrLf
                            'uneClasse &= getNumberTab(1) & " ''' Constructeur de base sans paramètre" & vbCrLf
                            'uneClasse &= getNumberTab(1) & "''' </summary>" & vbCrLf
                            'uneClasse &= getNumberTab(1) & "Public Sub New(" & getPKSignature(listeDesColonnes, True) & ")" & vbCrLf
                            'uneClasse &= "" & vbCrLf
                            'uneClasse &= getNumberTab(2) & " ' This call is required by the designer." & vbCrLf
                            'uneClasse &= getNumberTab(2) & "'InitializeComponent()" & vbCrLf
                            'uneClasse &= "" & vbCrLf
                            'uneClasse &= getNumberTab(2) & "' Add any initialization after the InitializeComponent() call." & vbCrLf
                            'uneClasse &= "" & vbCrLf
                            ''uneClasse &= constructeurContenu2
                            'uneClasse &= "" & vbCrLf
                            'uneClasse &= getNumberTab(1) & "End Sub" & vbCrLf
                            'uneClasse &= "" & vbCrLf

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

                            Select Case Me.TypeBaseDonnees
                                Case databaseType.SQL_SERVER,
                                     databaseType.SQL_SERVER_LOCALDB,
                                     databaseType.MYSQL,
                                     databaseType.ORACLE,
                                     databaseType.MS_ACCESS_2007_2019,
                                     databaseType.MS_ACCESS_97_2003,
                                     databaseType.MS_EXCEL,
                                     databaseType.FLAT_FILE
                                    qryDbName = db.Name
                                    qrySchema = Schem.Name
                                    qryTblName = tbl.Name
                                Case databaseType.POSTGRE_SQL
                                    qryDbName = """""" & db.Name & """"""
                                    qrySchema = """""" & Schem.Name & """""".Replace(".", """"".""""")
                                    qryTblName = ("""""" & tbl.Name & """""").Replace(".", """"".""""")
                                Case Else

                            End Select

                            If (tbl.IsTable) Then

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

                                'Création de la requête sql      
                                sqlTables = "INSERT INTO " & qryDbName & "." & qrySchema & "." & qryTblName & " ("
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

                                Select Case CType(Me.IntTypeBaseDonnees, databaseType)
                                    Case databaseType.SQL_SERVER

                                        NewConnections = "New OleDbConnection()"
                                        CommandeSQL = "New OleDbCommand(sqlTables, CType(aConn, OleDbConnection))"

                                    Case databaseType.SQL_SERVER_LOCALDB

                                        NewConnections = "New SqlClient.SqlConnection()"
                                        CommandeSQL = "New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))"

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

                            End If


                            '************************************************************************************
                            'Création de la fonction update de la classe
                            '************************************************************************************

                            If (tbl.IsTable) Then

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
                                sqlTables = "UPDATE " & qryDbName & "." & qrySchema & "." & qryTblName & " SET "
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

                                Select Case CType(Me.IntTypeBaseDonnees, databaseType)
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

                            End If

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
                            sqlTables = "SELECT COUNT(*) AS EST_EXISTANT FROM " & qryDbName & "." & qrySchema & "." & qryTblName & " "
                            sqlConditions = "WHERE "
                            If (tbl.IsTable) Then
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
                            Else
                                For Each col As UneColonne In listeDesColonnes

                                    sqlConditions &= col.NomColonneOriginal & " = "
                                    Select Case col.TypeDeDonnees
                                        Case "Long", "Integer", "Byte", "Double", "Decimal"
                                            sqlConditions &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "" AND "
                                        Case "Char", "DateDate", "String"
                                            sqlConditions &= "'"" & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & ""' AND "
                                        Case Else
                                            sqlConditions &= """ & _" & leNomDeLaTable.NomSingulier & "." & col.VarDePropriete & " & "" AND "
                                    End Select

                                Next

                            End If

                            If (sqlConditions = "WHERE ") Then
                                sqlConditions = ""
                            Else
                                sqlConditions = Strings.Left(sqlConditions, sqlConditions.Length - 5)
                            End If

                            sqlTables = sqlTables & sqlConditions & ";"

                            Select Case CType(Me.IntTypeBaseDonnees, databaseType)
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

                            If (tbl.IsTable) Then

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
                                sqlTables = "DELETE FROM " & qryDbName & "." & qrySchema & "." & qryTblName & " "
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

                                If (sqlConditions = "WHERE ") Then
                                    sqlConditions = ""
                                Else
                                    sqlConditions = Strings.Left(sqlConditions, sqlConditions.Length - 5)
                                End If

                                sqlTables = sqlTables & sqlConditions & ";"

                                Select Case CType(Me.IntTypeBaseDonnees, databaseType)
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

                            End If

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
                            "FROM " & qryDbName & "." & qrySchema & "." & qryTblName & ";"

                            Select Case CType(Me.IntTypeBaseDonnees, databaseType)
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

                            If (tbl.IsTable) Then

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

                                uneClasse &= Strings.Right(optionalConnection, optionalConnection.Length - 2) & ") As " & leNomDeLaTable.ClassName & vbCrLf
                                uneClasse &= getNumberTab(2) & "" & vbCrLf

                                uneClasse &= getNumberTab(2) & "Dim result As New " & leNomDeLaTable.ClassName & vbCrLf
                                uneClasse &= getNumberTab(2) & "Dim sqlTables As String" & vbCrLf
                                uneClasse &= getNumberTab(2) & "Dim aConn As DbConnection" & vbCrLf
                                uneClasse &= getNumberTab(2) & "Dim aCommand as DbCommand" & vbCrLf
                                uneClasse &= getNumberTab(2) & "Dim aDtr as DbDataReader" & vbCrLf
                                uneClasse &= getNumberTab(2) & "" & vbCrLf

                                'Création de la requête sql      
                                sqlTables = "SELECT * " &
                            "FROM " & qryDbName & "." & qrySchema & "." & qryTblName & " "
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

                                Select Case CType(Me.IntTypeBaseDonnees, databaseType)
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

                            End If

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

                            If (tbl.IsTable) Then

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

                            End If

                            '************************************************************************************
                            'Création de la fonction Save lui-même
                            '************************************************************************************

                            If (tbl.IsTable) Then

                                uneClasse &= getNumberTab(1) & "Public Function Save() As Boolean" & vbCrLf
                                uneClasse &= getNumberTab(2) & "If (Not BooIsSaved) Then" & vbCrLf
                                uneClasse &= getNumberTab(3) & "BooIsSaved = " & leNomDeLaTable.ClassName & ".Save" & leNomDeLaTable.NomSingulier & "(me)" & vbCrLf
                                uneClasse &= getNumberTab(2) & "End If" & vbCrLf
                                uneClasse &= getNumberTab(2) & "Return BooIsSaved" & vbCrLf
                                uneClasse &= getNumberTab(1) & "End Function" & vbCrLf
                                uneClasse &= getNumberTab(2) & "" & vbCrLf

                            End If

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

            Next

            Save()

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

    ''' <summary>
    ''' Retourne une liste de tous les ConnectionInfos de la table.
    ''' </summary>
    ''' <returns>Retourne une liste de tous les ConnectionInfos de la table</returns>
    Public Shared Function getAllConnectionInfos(ByRef _vmmw As ViewModelMainWindow,
                                                 Optional ByRef AnOpenConneciton As DbConnection = Nothing) As List(Of ConnectionInfos)

        Dim lst As New List(Of ConnectionInfos)
        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand
        Dim aDtr As DbDataReader

        If (AnOpenConneciton Is Nothing) Then
            'Ouverture de la connection SQL
            aConn = New SqlClient.SqlConnection()
            aConn.ConnectionString = OLEDB_CONN_STRING
            aConn.Open()
        Else
            aConn = AnOpenConneciton
        End If

        'Création de la requête sql
        sqlTables = "SELECT * FROM ConnectionListDB.CL_SCHEMA.TBL_CONNECTION_LIST;"

        'Création de la commande et on l'instancie (sql)
        aCommand = New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))

        'Création du datareader (aDtr)
        aDtr = aCommand.ExecuteReader()

        While aDtr.Read()

            Dim uneTable As New ConnectionInfos(_vmmw)

            uneTable.Id = If(IsDBNull(aDtr.Item("Id")), Nothing, CType(aDtr.Item("Id"), Integer))
            uneTable.ConnectionName = If(IsDBNull(aDtr.Item("CONNECTION_NAME")), Nothing, CType(aDtr.Item("CONNECTION_NAME"), String))
            uneTable.TypeBaseDonnees = If(IsDBNull(aDtr.Item("TYPE_BASE_DONNEES")), Nothing, CType(aDtr.Item("TYPE_BASE_DONNEES"), Integer))
            uneTable.ServerAddressName = If(IsDBNull(aDtr.Item("SERVER_ADDRESS_NAME")), Nothing, CType(aDtr.Item("SERVER_ADDRESS_NAME"), String))
            uneTable.TCPPort = If(IsDBNull(aDtr.Item("TCP_PORT")), Nothing, CType(aDtr.Item("TCP_PORT"), Integer))
            uneTable.DatabaseCatalog = If(IsDBNull(aDtr.Item("DATABASE_CATALOG")), Nothing, CType(aDtr.Item("DATABASE_CATALOG"), Integer))
            uneTable.SelectedDB = If(IsDBNull(aDtr.Item("SELECTED_DB")), Nothing, CType(aDtr.Item("SELECTED_DB"), String))
            uneTable.Username = If(IsDBNull(aDtr.Item("USERNAME")), Nothing, CType(aDtr.Item("USERNAME"), String))
            uneTable.strPassword = If(IsDBNull(aDtr.Item("PASSWORD")), Nothing, CType(aDtr.Item("PASSWORD"), String))
            uneTable.TrustedConnection = If(IsDBNull(aDtr.Item("TRUSTED_CONNECTION")), Nothing, CType(aDtr.Item("TRUSTED_CONNECTION"), Boolean))

            lst.Add(uneTable)

        End While

        aDtr.Close()
        aDtr = Nothing

        aCommand = Nothing
        If (AnOpenConneciton Is Nothing) Then
            aConn.Close()
            aConn = Nothing
        End If

        Return lst

    End Function

    ''' <summary>
    ''' Retourne une liste de tous les ConnectionInfos de la table.
    ''' </summary>
    ''' <returns>Retourne une liste de tous les ConnectionInfos de la table</returns>
    Public Shared Function getConnectionInfosFromID(ByRef _vmmw As ViewModelMainWindow,
                                                 ByVal _Id As Integer,
                                                    Optional ByRef AnOpenConneciton As DbConnection = Nothing) As ConnectionInfos

        Dim result As New ConnectionInfos(_vmmw)
        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand
        Dim aDtr As DbDataReader

        If (AnOpenConneciton Is Nothing) Then
            'Ouverture de la connection SQL
            aConn = New SqlClient.SqlConnection()
            aConn.ConnectionString = OLEDB_CONN_STRING
            aConn.Open()
        Else
            aConn = AnOpenConneciton
        End If

        'Création de la requête sql
        sqlTables = "SELECT * FROM ConnectionListDB.CL_SCHEMA.TBL_CONNECTION_LIST WHERE Id = " & _Id & ";"

        'Création de la commande et on l'instancie (sql)
        aCommand = New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))

        'Création du datareader (aDtr)
        aDtr = aCommand.ExecuteReader()

        While aDtr.Read()

            result.Id = If(IsDBNull(aDtr.Item("Id")), Nothing, CType(aDtr.Item("Id"), Integer))
            result.ConnectionName = If(IsDBNull(aDtr.Item("CONNECTION_NAME")), Nothing, CType(aDtr.Item("CONNECTION_NAME"), String))
            result.TypeBaseDonnees = If(IsDBNull(aDtr.Item("TYPE_BASE_DONNEES")), Nothing, CType(aDtr.Item("TYPE_BASE_DONNEES"), Integer))
            result.ServerAddressName = If(IsDBNull(aDtr.Item("SERVER_ADDRESS_NAME")), Nothing, CType(aDtr.Item("SERVER_ADDRESS_NAME"), String))
            result.TCPPort = If(IsDBNull(aDtr.Item("TCP_PORT")), Nothing, CType(aDtr.Item("TCP_PORT"), Integer))
            result.DatabaseCatalog = If(IsDBNull(aDtr.Item("DATABASE_CATALOG")), Nothing, CType(aDtr.Item("DATABASE_CATALOG"), Integer))
            result.SelectedDB = If(IsDBNull(aDtr.Item("SELECTED_DB")), Nothing, CType(aDtr.Item("SELECTED_DB"), String))
            result.Username = If(IsDBNull(aDtr.Item("USERNAME")), Nothing, CType(aDtr.Item("USERNAME"), String))
            result.strPassword = If(IsDBNull(aDtr.Item("PASSWORD")), Nothing, CType(aDtr.Item("PASSWORD"), String))
            result.TrustedConnection = If(IsDBNull(aDtr.Item("TRUSTED_CONNECTION")), Nothing, CType(aDtr.Item("TRUSTED_CONNECTION"), Boolean))


        End While

        aDtr.Close()
        aDtr = Nothing

        aCommand = Nothing
        If (AnOpenConneciton Is Nothing) Then
            aConn.Close()
            aConn = Nothing
        End If

        Return result

    End Function

    Public Overrides Function ToString() As String
        Return StrConnectionName
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean

        Dim value As ConnectionInfos = TryCast(obj, ConnectionInfos)
        If (value Is Nothing) Then
            Return False
        Else
            Return ((Me.IntId = value.Id) And
                    (Me.StrConnectionName = value.ConnectionName) And
                    (Me.IntTypeBaseDonnees = value.TypeBaseDonnees) And
                    (Me.StrServerAddressName = value.ServerAddressName) And
                    (Me.IntTCPPort = value.TCPPort) And
                    (Me.IntDatabaseCatalog = value.DatabaseCatalog) And
                    (Me.StrSelectedDb = value.SelectedDB) And
                    (Me.strUsername = value.Username) And
                    (Me.strPassword = value.strPassword) And
                    (Me.booTrustedConnection = value.TrustedConnection))
        End If

    End Function

    Public Shared Function SaveConnectionInfos(ByVal _ConnectionInfos As ConnectionInfos, Optional ByRef AnOpenConneciton As DbConnection = Nothing) As Boolean

        Dim resultat As Boolean = False

        If (ConnectionInfosExists(_ConnectionInfos, AnOpenConneciton)) Then
            resultat = (UpdateConnectionInfos(_ConnectionInfos, AnOpenConneciton) > 0)
        Else
            resultat = (InsertConnectionInfos(_ConnectionInfos, AnOpenConneciton) > 0)
        End If

        Return resultat
    End Function

    Public Function Save() As Boolean
        If (Not BooIsSaved) Then
            BooIsSaved = ConnectionInfos.SaveConnectionInfos(Me)
        End If
        Return BooIsSaved
    End Function


    Protected Friend Sub OnPropertyChanged(<CallerMemberName()> Optional ByVal propertyName As String = Nothing)
        'MsgBox(propertyName & vbCrLf & ViewModelPeriodeDePaie.instance & vbCrLf & "Lecture :" & Employes.LectureDePropriete & vbCrLf & "Ecriture :" & Employes.EcriturePropriete & vbCrLf & "Fonction :" & ViewModelPeriodeDePaie.AppelDeFonction)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub


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

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

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
                    If (Not Me.ServerAddressName Is Nothing) Then
                        If (Me.ServerAddressName = String.Empty) Then
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
