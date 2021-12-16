'************************************************************************************
'* DÉVELOPPÉ PAR : Jean-Claude Frigon                                                *
'* DATE : 14 octobre 2020                                                            *
'* MODIFIÉ : 14 octobre 2020                                                         *
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
Imports NpgSql
Imports System.Runtime.CompilerServices

#End Region

Public Class AppVersion
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

	Private IntAppId As Integer
	Private StrAppName As String
	Private StrAppVersion As String
	Private StrAppSetupFilePath As String
	Private DteLastCleanup As Date

    '------- --------
    '------- --------
    'Section publique
    '------- --------
    '------- --------

    Public Const ConnectionString = "Server=192.168.0.30;Port=5432;Database=new_exacad_db;User Id=postgres;Password=uLOAddL8;"

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

		IntAppId = 0
		StrAppName = ""
		StrAppVersion = ""
		StrAppSetupFilePath = ""
		DteLastCleanup = Now()

	End Sub

	''' <summary>
	 ''' Constructeur de base avec paramètres
	''' </summary>
	Public Sub New(Byval _AppId As Integer, Byval _AppName As String, Byval _AppVersion As String, Byval _AppSetupFilePath As String, Byval _LastCleanup As Date)

	' This call is required by the designer.
	'InitializeComponent()

	' Add any initialization after the InitializeComponent() call.

		IntAppId = _AppId
		StrAppName = _AppName
		StrAppVersion = _AppVersion
		StrAppSetupFilePath = _AppSetupFilePath
		DteLastCleanup = _LastCleanup

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

	Public Property AppId As Integer
		Get
				Return IntAppId
		End Get
		Set(value As Integer)
			IntAppId = value
			BooIsSaved = False
			OnPropertyChanged()
		End Set
	End Property

	Public Property AppName As String
		Get
				Return StrAppName
		End Get
		Set(value As String)
			StrAppName = value
			BooIsSaved = False
			OnPropertyChanged()
		End Set
	End Property

    Public Property ApplicationVersion As String
        Get
            Return StrAppVersion
        End Get
        Set(value As String)
            StrAppVersion = value
            BooIsSaved = False
            OnPropertyChanged()
        End Set
    End Property

    Public Property AppSetupFilePath As String
		Get
				Return StrAppSetupFilePath
		End Get
		Set(value As String)
			StrAppSetupFilePath = value
			BooIsSaved = False
			OnPropertyChanged()
		End Set
	End Property

	Public Property LastCleanup As Date
		Get
				Return DteLastCleanup
		End Get
		Set(value As Date)
			DteLastCleanup = value
			BooIsSaved = False
			OnPropertyChanged()
		End Set
	End Property

    Public ReadOnly Property Major As Integer
        Get
            Dim ver As String() = ApplicationVersion.Split("."c)

            If (ver.Count = 4) Then

                Dim mjr As Integer

                If (Integer.TryParse(ver(0), mjr)) Then

                    Return mjr

                Else

                    Throw New Exception("Le numéro de version n'est pas conforme.")

                End If

            Else

                Throw New Exception("Le numéro de version n'est pas conforme.")

            End If

        End Get
    End Property

    Public ReadOnly Property Minor As Integer
        Get
            Dim ver As String() = ApplicationVersion.Split("."c)

            If (ver.Count = 4) Then

                Dim mnr As Integer

                If (Integer.TryParse(ver(1), mnr)) Then

                    Return mnr

                Else

                    Throw New Exception("Le numéro de version n'est pas conforme.")

                End If

            Else

                Throw New Exception("Le numéro de version n'est pas conforme.")

            End If

        End Get
    End Property

    Public ReadOnly Property Build As Integer
        Get
            Dim ver As String() = ApplicationVersion.Split("."c)

            If (ver.Count = 4) Then

                Dim mjrRev As Integer

                If (Integer.TryParse(ver(2), mjrRev)) Then

                    Return mjrRev

                Else

                    Throw New Exception("Le numéro de version n'est pas conforme.")

                End If

            Else

                Throw New Exception("Le numéro de version n'est pas conforme.")

            End If

        End Get
    End Property

    Public ReadOnly Property Revision As Integer
        Get
            Dim ver As String() = ApplicationVersion.Split("."c)

            If (ver.Count = 4) Then

                Dim mnrRev As Integer

                If (Integer.TryParse(ver(3), mnrRev)) Then

                    Return mnrRev

                Else

                    Throw New Exception("Le numéro de version n'est pas conforme.")

                End If

            Else

                Throw New Exception("Le numéro de version n'est pas conforme.")

            End If

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
    ''' <param name="_AppVersion">Un objet de type AppVersion</param>
    ''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>
    Private Shared Function InsertAppVersion(ByVal _AppVersion As AppVersion, Optional Byref AnOpenConneciton as DbConnection = Nothing) As Integer

		Dim resultat As Integer = 0
		
		Dim sqlTables As String
		Dim aConn As DbConnection
		Dim aCommand As DbCommand
		
		If (AnOpenConneciton Is Nothing) then
			'Ouverture de la connection SQL
			aConn = New NpgSqlConnection()
            aConn.ConnectionString = ConnectionString
            aConn.Open()
        Else
            aConn = AnOpenConneciton
        End If

        'Création de la requête sql
        sqlTables = "INSERT INTO ""new_exacad_db"".""applications"".""tbl_app_version"" (app_id, app_name, app_version, app_setup_file_path, last_cleanup) VALUES (" & _AppVersion.AppId & ", '" & _AppVersion.AppName & "', '" & _AppVersion.ApplicationVersion & "', '" & _AppVersion.AppSetupFilePath & "', " & _AppVersion.LastCleanup & ");"

        'Création de la commande et on l'exécute
        aCommand = New NpgsqlCommand(sqlTables, CType(aConn, NpgsqlConnection))

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
    ''' <param name="_AppVersion">Un objet de type AppVersion</param>
    ''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>
    Private Shared Function UpdateAppVersion(ByVal _AppVersion As AppVersion, Optional ByRef AnOpenConneciton As DbConnection = Nothing) As Integer

        Dim resultat As Integer = 0

        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand

        If (AnOpenConneciton Is Nothing) Then
            'Ouverture de la connection SQL
            aConn = New NpgsqlConnection()
            aConn.ConnectionString = ConnectionString
            aConn.Open()
        Else
            aConn = AnOpenConneciton
        End If

        'Création de la requête sql
        sqlTables = "UPDATE ""new_exacad_db"".""applications"".""tbl_app_version"" SET app_name = '" & _AppVersion.AppName & "', app_version = '" & _AppVersion.ApplicationVersion & "', app_setup_file_path = '" & _AppVersion.AppSetupFilePath & "', last_cleanup = '" & _AppVersion.LastCleanup & "' WHERE app_id = " & _AppVersion.AppId & ";"

        'Création de la commande et on l'exécute
        aCommand = New NpgsqlCommand(sqlTables, CType(aConn, NpgsqlConnection))

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
    ''' <param name="_AppVersion">Un objet de type AppVersion</param>
    ''' <returns>Retourne Vrai si l'objet existe dans la base de données ou False s'il n'existe pas.</returns>
    Private Shared Function AppVersionExists(ByVal _AppVersion As AppVersion, Optional ByRef AnOpenConneciton As DbConnection = Nothing) As Boolean

        Dim resultat As Boolean = False

        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand

        If (AnOpenConneciton Is Nothing) Then
            'Ouverture de la connection SQL
            aConn = New NpgsqlConnection()
            aConn.ConnectionString = ConnectionString
            aConn.Open()
        Else
            aConn = AnOpenConneciton
        End If

        'Création de la requête sql
        sqlTables = "SELECT COUNT(*) AS EST_EXISTANT FROM ""new_exacad_db"".""applications"".""tbl_app_version"" WHERE app_id = " & _AppVersion.AppId & ";"

        'Création de la commande et on l'exécute
        aCommand = New NpgsqlCommand(sqlTables, CType(aConn, NpgsqlConnection))
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
    ''' <param name="_AppVersion">Un objet de type AppVersion</param>
    ''' <returns>Retourne un integer représentant le nombre d'enregistrements affectés : devrait toujours être 1 ou 0</returns>
    Private Shared Function DeleteAppVersion(ByVal _AppVersion As AppVersion, Optional ByRef AnOpenConneciton As DbConnection = Nothing) As Integer

        Dim resultat As Integer = 0

        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand

        If (AppVersionExists(_AppVersion)) Then

            If (AnOpenConneciton Is Nothing) Then
                'Ouverture de la connection SQL
                aConn = New NpgsqlConnection()
                aConn.ConnectionString = ConnectionString
                aConn.Open()
            Else
                aConn = AnOpenConneciton
            End If

            'Création de la requête sql
            sqlTables = "DELETE FROM ""new_exacad_db"".""applications"".""tbl_app_version"" WHERE app_id = " & _AppVersion.AppId & ";"

            'Création de la commande et on l'exécute
            aCommand = New NpgsqlCommand(sqlTables, CType(aConn, NpgsqlConnection))
            Try
                resultat = CType(aCommand.ExecuteNonQuery(), Integer)
            Catch ex As Exception
                resultat = -1
            End Try

            aCommand = Nothing
            If (AnOpenConneciton Is Nothing) Then
                aConn.Close()
                aConn = Nothing
            End If

        Else

            'MsgBox("L'objet ""AppVersion"" à supprimer n'existe pas.")
            resultat = 0

        End If

        Return resultat
    End Function

    '------- --------
    '------- --------
    'Section publique
    '------- --------
    '------- --------

    ''' <summary>
    ''' Retourne une liste de tous les AppVersion de la table.
    ''' </summary>
    ''' <returns>Retourne une liste de tous les AppVersion de la table</returns>
    Public Shared Function getAllAppVersion(Optional ByRef AnOpenConneciton As DbConnection = Nothing) As List(Of AppVersion)

        Dim lst As New List(Of AppVersion)
        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand
        Dim aDtr As DbDataReader

        If (AnOpenConneciton Is Nothing) Then
            'Ouverture de la connection SQL
            aConn = New NpgsqlConnection()
            aConn.ConnectionString = ConnectionString
            aConn.Open()
        Else
            aConn = AnOpenConneciton
        End If

        'Création de la requête sql
        sqlTables = "SELECT * FROM ""new_exacad_db"".""applications"".""tbl_app_version"";"

        'Création de la commande et on l'instancie (sql)
        aCommand = New NpgsqlCommand(sqlTables, CType(aConn, NpgsqlConnection))

        'Création du datareader (aDtr)
        aDtr = aCommand.ExecuteReader()

        While aDtr.Read()

            Dim uneTable As New AppVersion()

            uneTable.AppId = If(IsDBNull(aDtr.Item("app_id")), Nothing, CType(aDtr.Item("app_id"), Integer))
            uneTable.AppName = If(IsDBNull(aDtr.Item("app_name")), Nothing, CType(aDtr.Item("app_name"), String))
            uneTable.ApplicationVersion = If(IsDBNull(aDtr.Item("app_version")), Nothing, CType(aDtr.Item("app_version"), String))
            uneTable.AppSetupFilePath = If(IsDBNull(aDtr.Item("app_setup_file_path")), Nothing, CType(aDtr.Item("app_setup_file_path"), String))
            uneTable.LastCleanup = If(IsDBNull(aDtr.Item("last_cleanup")), Nothing, CType(aDtr.Item("last_cleanup"), Date))

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
    ''' Retourne une liste de tous les AppVersion de la table.
    ''' </summary>
    ''' <returns>Retourne une liste de tous les AppVersion de la table</returns>
    Public Shared Function getAppVersionFromID(ByVal _app_id As Integer, Optional ByRef AnOpenConneciton As DbConnection = Nothing) As AppVersion

        Dim result As New AppVersion
        Dim sqlTables As String
        Dim aConn As DbConnection
        Dim aCommand As DbCommand
        Dim aDtr As DbDataReader

        If (AnOpenConneciton Is Nothing) Then
            'Ouverture de la connection SQL
            aConn = New NpgsqlConnection()
            aConn.ConnectionString = ConnectionString
            aConn.Open()
		Else
			aConn = AnOpenConneciton
		End If
		
		'Création de la requête sql
		sqlTables = "SELECT * FROM ""new_exacad_db"".""applications"".""tbl_app_version"" WHERE app_id = " & _app_id & ";"
		
		'Création de la commande et on l'instancie (sql)
		aCommand = New NpgSqlCommand(sqlTables, CType(aConn, NpgSqlConnection))
		
		'Création du datareader (aDtr)
		aDtr = aCommand.ExecuteReader()
		
		While aDtr.Read()
		
			result.AppId = If(IsDBNull(aDtr.Item("app_id")), Nothing, CType(aDtr.Item("app_id"), Integer))
			result.AppName = If(IsDBNull(aDtr.Item("app_name")), Nothing, CType(aDtr.Item("app_name"), String))
            result.ApplicationVersion = If(IsDBNull(aDtr.Item("app_version")), Nothing, CType(aDtr.Item("app_version"), String))
            result.AppSetupFilePath = If(IsDBNull(aDtr.Item("app_setup_file_path")), Nothing, CType(aDtr.Item("app_setup_file_path"), String))
			result.LastCleanup = If(IsDBNull(aDtr.Item("last_cleanup")), Nothing, CType(aDtr.Item("last_cleanup"), Date))
			
		
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
		Return StrAppName
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean

		dim value as AppVersion = trycast(obj, AppVersion)
		If (value Is Nothing) then
			Return False
		Else
            Return ((Me.IntAppId = value.AppId) And (Me.StrAppName = value.AppName) And (Me.StrAppVersion = value.ApplicationVersion) And (Me.StrAppSetupFilePath = value.AppSetupFilePath) And (Me.DteLastCleanup = value.LastCleanup))
        End If

	End Function

	Public Shared Function SaveAppVersion(ByVal _AppVersion As AppVersion, Optional Byref AnOpenConneciton as DbConnection = Nothing) As Boolean
		
		Dim resultat As Boolean = False
		
		If (AppVersionExists(_AppVersion, AnOpenConneciton)) Then
			resultat = (UpdateAppVersion(_AppVersion, AnOpenConneciton) > 0)
		Else
			resultat = (InsertAppVersion(_AppVersion, AnOpenConneciton) > 0)
		End If
		
		Return resultat
	End Function
		
	Public Function Save(Optional ByVal whatever As Boolean = False) As Boolean
		If (Not BooIsSaved) Or (whatever) Then
			BooIsSaved = AppVersion.SaveAppVersion(me)
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
