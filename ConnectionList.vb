'************************************************************************************
'* DÉVELOPPÉ PAR : Jean-Claude Frigon                                                *
'* DATE : Friday, March 27, 2020                                                     *
'* MODIFIÉ : Friday, March 27, 2020                                                  *
'* PAR :                                                                            *
'* DESCRIPTION :                                                                    *
'*      Public [Function | Sub] nomProcFunct( paramètres)                           *
'*      Public Function ToString() as String                                        *
'*      Public Function Equals() as Boolean                                         *
'************************************************************************************

'************************************************************************************
'                                                                                   *
'                           L I B R A R Y  I M P O R T S                            *
'                                                                                   *
'************************************************************************************
#Region "Library Imports"

Imports System.ComponentModel
Imports System.Data.Common
Imports System.Runtime.CompilerServices

#End Region

Public Class ConnectionList
	Implements IDisposable, INotifyPropertyChanged

	 '************************************************************************************
	'                            V  A  R  I  A  B  L  E  S                              *
	'                        D E C L A R E   F U N C T I O N S                          *
	'                                    T Y P E S                                      *
	'************************************************************************************
#Region "Variables, Declare Functions And Types"

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
	Private IntTypeBaseDonnees As Integer
	Private StrServerAddressName As String
	Private IntTcpPort As Integer
	Private IntDatabaseCatalog As Integer
	Private StrSelectedDb As String
	Private StrUsername As String
	Private StrPassword As String
	Private BooTrustedConnection As Boolean

	'------- --------
	'------- --------
	'Section publique
	'------- --------
	'------- --------

Public Const OLEDB_CONN_STRING = "Server=(LocalDb)\MSSQLLocalDB;Database=ConnectionListDB;Uid=ConnectionList;Pwd=X043rMiVpOlbAlGT9UKZ;"

#End Region
	'************************************************************************************
	'                    C  O  N  S  T  R  U  C  T  E  U  R                             *
	'                    ----------------------------------                             *
	'                      D  E  S  T  R  U  C  T  E  U  R                              *
	'************************************************************************************
	#Region "Constructors"

	''' <summary>
	 ''' Constructeur de base sans paramètre
	''' </summary>
	Public Sub New()

		' This call is required by the designer.
		'InitializeComponent()

		' Add any initialization after the InitializeComponent() call.

		IntId = 0
		StrConnectionName = ""
		IntTypeBaseDonnees = 0
		StrServerAddressName = ""
		IntTcpPort = 0
		IntDatabaseCatalog = 0
		StrSelectedDb = ""
		StrUsername = ""
		StrPassword = ""
		BooTrustedConnection = Nothing

	End Sub

	''' <summary>
	 ''' Constructeur de base avec paramètres
	''' </summary>
	Public Sub New(Byval _Id As Integer, Byval _ConnectionName As String, Byval _TypeBaseDonnees As Integer, Byval _ServerAddressName As String, Byval _TcpPort As Integer, Byval _DatabaseCatalog As Integer, Byval _SelectedDb As String, Byval _Username As String, Byval _Password As String, Byval _TrustedConnection As Boolean)

	' This call is required by the designer.
	'InitializeComponent()

	' Add any initialization after the InitializeComponent() call.

		IntId = _Id
		StrConnectionName = _ConnectionName
		IntTypeBaseDonnees = _TypeBaseDonnees
		StrServerAddressName = _ServerAddressName
		IntTcpPort = _TcpPort
		IntDatabaseCatalog = _DatabaseCatalog
		StrSelectedDb = _SelectedDb
		StrUsername = _Username
		StrPassword = _Password
		BooTrustedConnection = _TrustedConnection

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

	Public Property TypeBaseDonnees As Integer
		Get
				Return IntTypeBaseDonnees
		End Get
		Set(value As Integer)
			IntTypeBaseDonnees = value
			BooIsSaved = False
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

	Public Property TcpPort As Integer
		Get
				Return IntTcpPort
		End Get
		Set(value As Integer)
			IntTcpPort = value
			BooIsSaved = False
			OnPropertyChanged()
		End Set
	End Property

	Public Property DatabaseCatalog As Integer
		Get
				Return IntDatabaseCatalog
		End Get
		Set(value As Integer)
			IntDatabaseCatalog = value
			BooIsSaved = False
			OnPropertyChanged()
		End Set
	End Property

	Public Property SelectedDb As String
		Get
				Return StrSelectedDb
		End Get
		Set(value As String)
			StrSelectedDb = value
			BooIsSaved = False
			OnPropertyChanged()
		End Set
	End Property

	Public Property Username As String
		Get
				Return StrUsername
		End Get
		Set(value As String)
			StrUsername = value
			BooIsSaved = False
			OnPropertyChanged()
		End Set
	End Property

	Public Property Password As String
		Get
				Return StrPassword
		End Get
		Set(value As String)
			StrPassword = value
			BooIsSaved = False
			OnPropertyChanged()
		End Set
	End Property

	Public Property TrustedConnection As Boolean
		Get
				Return BooTrustedConnection
		End Get
		Set(value As Boolean)
			BooTrustedConnection = value
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
		''' <param name="_ConnectionList">Un objet de type ConnectionList</param>
		''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>
	Private Shared Function InsertConnectionList(ByVal _ConnectionList As ConnectionList, Optional Byref AnOpenConneciton as DbConnection = Nothing) As Integer

		Dim resultat As Integer = 0
		
		Dim sqlTables As String
		Dim aConn As DbConnection
		Dim aCommand As DbCommand
		
		If (AnOpenConneciton Is Nothing) then
			'Ouverture de la connection SQL
			aConn = New SqlClient.SqlConnection()
			aConn.ConnectionString = OLEDB_CONN_STRING
			aConn.Open()
		Else
			aConn = AnOpenConneciton
		End If
		
		'Création de la requête sql
		sqlTables = "INSERT INTO ConnectionListDB.CL_SCHEMA.TBL_CONNECTION_LIST (Id, CONNECTION_NAME, TYPE_BASE_DONNEES, SERVER_ADDRESS_NAME, TCP_PORT, DATABASE_CATALOG, SELECTED_DB, USERNAME, PASSWORD, TRUSTED_CONNECTION) VALUES (" & _ConnectionList.Id & ", '" & _ConnectionList.ConnectionName & "', " & _ConnectionList.TypeBaseDonnees & ", '" & _ConnectionList.ServerAddressName & "', " & _ConnectionList.TcpPort & ", " & _ConnectionList.DatabaseCatalog & ", '" & _ConnectionList.SelectedDb & "', '" & _ConnectionList.Username & "', '" & _ConnectionList.Password & "', " & _ConnectionList.TrustedConnection & ");"
		
		'Création de la commande et on l'exécute
		aCommand = New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))
		
		resultat = aCommand.ExecuteNonQuery()
		
		aCommand = Nothing
		If (AnOpenConneciton Is Nothing) then
			aConn.Close()
			aConn = Nothing
		End If
		
		Return resultat
	End Function
	
		''' <summary>
		''' Fonction permettant de mettre à jour l'objet passé en paramètre dans la table de la classe s'il existe.
		''' </summary>
		''' <param name="_ConnectionList">Un objet de type ConnectionList</param>
		''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>
	Private Shared Function UpdateConnectionList(ByVal _ConnectionList As ConnectionList, Optional Byref AnOpenConneciton as DbConnection = Nothing) As Integer

		Dim resultat As Integer = 0
		
		Dim sqlTables As String
		Dim aConn As DbConnection
		Dim aCommand As DbCommand
		
		If (AnOpenConneciton Is Nothing) then
			'Ouverture de la connection SQL
			aConn = New SqlClient.SqlConnection()
			aConn.ConnectionString = OLEDB_CONN_STRING
			aConn.Open()
		Else
			aConn = AnOpenConneciton
		End If
		
		'Création de la requête sql
		sqlTables = "UPDATE ConnectionListDB.CL_SCHEMA.TBL_CONNECTION_LIST SET CONNECTION_NAME = '" & _ConnectionList.ConnectionName & "', TYPE_BASE_DONNEES = " & _ConnectionList.TypeBaseDonnees & ", SERVER_ADDRESS_NAME = '" & _ConnectionList.ServerAddressName & "', TCP_PORT = " & _ConnectionList.TcpPort & ", DATABASE_CATALOG = " & _ConnectionList.DatabaseCatalog & ", SELECTED_DB = '" & _ConnectionList.SelectedDb & "', USERNAME = '" & _ConnectionList.Username & "', PASSWORD = '" & _ConnectionList.Password & "', TRUSTED_CONNECTION = " & _ConnectionList.TrustedConnection & " WHERE Id = " & _ConnectionList.Id & ";"
		
		'Création de la commande et on l'exécute
		aCommand = New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))
		
		resultat = aCommand.ExecuteNonQuery()
		
		aCommand = Nothing
		If (AnOpenConneciton Is Nothing) then
			aConn.Close()
			aConn = Nothing
		End If
		
		Return resultat
	End Function
	
		''' <summary>
		''' Fonction permettant de déterminer si l'objet passé en paramètre existe dans la base de données ou non.
		''' </summary>
		''' <param name="_ConnectionList">Un objet de type ConnectionList</param>
		''' <returns>Retourne Vrai si l'objet existe dans la base de données ou False s'il n'existe pas.</returns>
	Private Shared Function ConnectionListExists(ByVal _ConnectionList As ConnectionList, Optional Byref AnOpenConneciton as DbConnection = Nothing) As Boolean
		
		Dim resultat As Boolean = False
		
		Dim sqlTables As String
		Dim aConn As DbConnection
		Dim aCommand As DbCommand
		
		If (AnOpenConneciton Is Nothing) then
			'Ouverture de la connection SQL
			aConn = New SqlClient.SqlConnection()
			aConn.ConnectionString = OLEDB_CONN_STRING
			aConn.Open()
		Else
			aConn = AnOpenConneciton
		End If
		
		'Création de la requête sql
		sqlTables = "SELECT COUNT(*) AS EST_EXISTANT FROM ConnectionListDB.CL_SCHEMA.TBL_CONNECTION_LIST WHERE Id = " & _ConnectionList.Id & ";"
		
		'Création de la commande et on l'exécute
		aCommand = New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))
		Try
			resultat = (CType(aCommand.ExecuteScalar(), Integer) > 0)
		Catch ex As Exception
			resultat = False
		End Try
		
		aCommand = Nothing
		If (AnOpenConneciton Is Nothing) then
			aConn.Close()
			aConn = Nothing
		End If
		
		Return resultat
	End Function
		
	''' <summary>
	''' Fonction permettant de supprimer l'objet passé en paramètre de la base de données.
	''' </summary>
	''' <param name="_ConnectionList">Un objet de type ConnectionList</param>
	''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>
	Private Shared Function DeleteConnectionList(ByVal _ConnectionList As ConnectionList, Optional Byref AnOpenConneciton as DbConnection = Nothing) As Integer
		
		Dim resultat As Integer = 0
		
			Dim sqlTables As String
		Dim aConn As DbConnection
		Dim aCommand As DbCommand
		
		If (ConnectionListExists(_ConnectionList)) Then
			
		If (AnOpenConneciton Is Nothing) then
			'Ouverture de la connection SQL
			aConn = New SqlClient.SqlConnection()
			aConn.ConnectionString = OLEDB_CONN_STRING
			aConn.Open()
		Else
			aConn = AnOpenConneciton
		End If
			
			'Création de la requête sql
			sqlTables = "DELETE FROM ConnectionListDB.CL_SCHEMA.TBL_CONNECTION_LIST WHERE Id = " & _ConnectionList.Id & ";"
			
			'Création de la commande et on l'exécute
			aCommand = New SqlClient.SqlCommand(sqlTables, CType(aConn, SqlClient.SqlConnection))
			Try
			resultat =  CType(aCommand.ExecuteScalar(), Integer)
			Catch ex As Exception
			resultat = -1
			End Try
			
			aCommand = Nothing
		If (AnOpenConneciton Is Nothing) then
			aConn.Close()
			aConn = Nothing
		End If
			
		Else
			
			MsgBox("L'objet ""ConnectionList"" à supprimer n'existe pas.")
			resultat = -1
			
		End If
		
		Return resultat
	End Function
	
	'------- --------
	'------- --------
	'Section publique
	'------- --------
	'------- --------

		''' <summary>
		''' Retourne une liste de tous les ConnectionList de la table.
		''' </summary>
		''' <returns>Retourne une liste de tous les ConnectionList de la table</returns>
	Public Shared Function getAllConnectionList(Optional Byref AnOpenConneciton as DbConnection = Nothing) As List(Of ConnectionList)
		
		Dim lst As New List(Of ConnectionList)
		Dim sqlTables As String
		Dim aConn As DbConnection
		Dim aCommand as DbCommand
		Dim aDtr as DbDataReader
		
		If (AnOpenConneciton Is Nothing) then
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
		
			Dim uneTable as new ConnectionList()
		
			uneTable.Id = If(IsDBNull(aDtr.Item("Id")), Nothing, CType(aDtr.Item("Id"), Integer))
			uneTable.ConnectionName = If(IsDBNull(aDtr.Item("CONNECTION_NAME")), Nothing, CType(aDtr.Item("CONNECTION_NAME"), String))
			uneTable.TypeBaseDonnees = If(IsDBNull(aDtr.Item("TYPE_BASE_DONNEES")), Nothing, CType(aDtr.Item("TYPE_BASE_DONNEES"), Integer))
			uneTable.ServerAddressName = If(IsDBNull(aDtr.Item("SERVER_ADDRESS_NAME")), Nothing, CType(aDtr.Item("SERVER_ADDRESS_NAME"), String))
			uneTable.TcpPort = If(IsDBNull(aDtr.Item("TCP_PORT")), Nothing, CType(aDtr.Item("TCP_PORT"), Integer))
			uneTable.DatabaseCatalog = If(IsDBNull(aDtr.Item("DATABASE_CATALOG")), Nothing, CType(aDtr.Item("DATABASE_CATALOG"), Integer))
			uneTable.SelectedDb = If(IsDBNull(aDtr.Item("SELECTED_DB")), Nothing, CType(aDtr.Item("SELECTED_DB"), String))
			uneTable.Username = If(IsDBNull(aDtr.Item("USERNAME")), Nothing, CType(aDtr.Item("USERNAME"), String))
			uneTable.Password = If(IsDBNull(aDtr.Item("PASSWORD")), Nothing, CType(aDtr.Item("PASSWORD"), String))
			uneTable.TrustedConnection = If(IsDBNull(aDtr.Item("TRUSTED_CONNECTION")), Nothing, CType(aDtr.Item("TRUSTED_CONNECTION"), Boolean))
			
			lst.Add(uneTable)
		
		End While
		
		aDtr.Close()
		aDtr = Nothing
		
		aCommand = Nothing
		If (AnOpenConneciton Is Nothing) then
			aConn.Close()
			aConn = Nothing
		End If
		
		Return lst
		
	End Function

		''' <summary>
		''' Retourne une liste de tous les ConnectionList de la table.
		''' </summary>
		''' <returns>Retourne une liste de tous les ConnectionList de la table</returns>
	Public Shared Function getConnectionListFromID(Byval _Id as Integer, Optional Byref AnOpenConneciton as DbConnection = Nothing) As ConnectionList
		
		Dim result As New ConnectionList
		Dim sqlTables As String
		Dim aConn As DbConnection
		Dim aCommand as DbCommand
		Dim aDtr as DbDataReader
		
		If (AnOpenConneciton Is Nothing) then
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
			result.TcpPort = If(IsDBNull(aDtr.Item("TCP_PORT")), Nothing, CType(aDtr.Item("TCP_PORT"), Integer))
			result.DatabaseCatalog = If(IsDBNull(aDtr.Item("DATABASE_CATALOG")), Nothing, CType(aDtr.Item("DATABASE_CATALOG"), Integer))
			result.SelectedDb = If(IsDBNull(aDtr.Item("SELECTED_DB")), Nothing, CType(aDtr.Item("SELECTED_DB"), String))
			result.Username = If(IsDBNull(aDtr.Item("USERNAME")), Nothing, CType(aDtr.Item("USERNAME"), String))
			result.Password = If(IsDBNull(aDtr.Item("PASSWORD")), Nothing, CType(aDtr.Item("PASSWORD"), String))
			result.TrustedConnection = If(IsDBNull(aDtr.Item("TRUSTED_CONNECTION")), Nothing, CType(aDtr.Item("TRUSTED_CONNECTION"), Boolean))
			
		
		End While
		
		aDtr.Close()
		aDtr = Nothing
		
		aCommand = Nothing
		If (AnOpenConneciton Is Nothing) then
			aConn.Close()
			aConn = Nothing
		End If
		
		Return result
		
	End Function

	Public Overrides Function ToString() As String
		Return StrConnectionName
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean

		dim value as ConnectionList = trycast(obj, ConnectionList)
		If (value Is Nothing) then
			Return False
		Else
			Return ((Me.IntId= value.Id) And (Me.StrConnectionName= value.ConnectionName) And (Me.IntTypeBaseDonnees= value.TypeBaseDonnees) And (Me.StrServerAddressName= value.ServerAddressName) And (Me.IntTcpPort= value.TcpPort) And (Me.IntDatabaseCatalog= value.DatabaseCatalog) And (Me.StrSelectedDb= value.SelectedDb) And (Me.StrUsername= value.Username) And (Me.StrPassword= value.Password) And (Me.BooTrustedConnection= value.TrustedConnection))
		End If

	End Function

	Public Shared Function SaveConnectionList(ByVal _ConnectionList As ConnectionList, Optional Byref AnOpenConneciton as DbConnection = Nothing) As Boolean
		
		Dim resultat As Boolean = False
		
		If (ConnectionListExists(_ConnectionList, AnOpenConneciton)) Then
			resultat = (UpdateConnectionList(_ConnectionList, AnOpenConneciton) > 0)
		Else
			resultat = (InsertConnectionList(_ConnectionList, AnOpenConneciton) > 0)
		End If
		
		Return resultat
	End Function
		
	Public Function Save() As Boolean
		If (Not BooIsSaved) Then
			BooIsSaved = ConnectionList.SaveConnectionList(me)
		End If
		Return BooIsSaved
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

	' This method is called by the Set accessor of each property.
	' The CallerMemberName attribute that is applied to the optional propertyName
	' parameter causes the property name of the caller to be substituted as an argument.
	Private Sub OnPropertyChanged(<CallerMemberName()> Optional ByVal propertyName As String = Nothing)
		    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
	End Sub

	'------- --------
	'------- --------
	'Section publique
	'------- --------
	'------- --------

#End Region

End Class
