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
Imports System.Data


#End Region

Public Class DbDatabase
    Inherits TreeView.Noeud
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

    Public Shared Function createDbDatabase() As ObservableCollection(Of TreeView.Noeud)

        Dim lst As New ObservableCollection(Of TreeView.Noeud)

        For i As Integer = 1 To 5
            Dim tbls As New list(Of TreeView.Noeud)
            For j As Integer = 1 To 5
                Dim cols As New List(Of String)
                For k As Integer = 1 To 12
                    cols.Add("Colonne_" & k)
                Next
                tbls.Add(New DbTable("Table_" & j, cols, False))
            Next
            lst.Add(New DbSchemas("Schema_" & i, tbls, False))
        Next


        Dim lstdb As New ObservableCollection(Of TreeView.Noeud)
        lstdb.Add(New DbDatabase("Test DB", lst, False))

        Return lstdb

    End Function

    Public Sub New(ByVal database As String)

        ' This call is required by the designer.
        'InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Me.New(database, False)

    End Sub

    Public Sub New(ByVal database As String, ByVal listOfNoeud As ObservableCollection(Of TreeView.Noeud))
        Me.New(database, listOfNoeud, False)
    End Sub

    Public Sub New(ByVal database As String, ByVal isSelected As Boolean)
        Me.New(database, New ObservableCollection(Of TreeView.Noeud), False)
    End Sub

    Public Sub New(ByVal database As String, ByVal listOfNoeud As ObservableCollection(Of TreeView.Noeud), ByVal _isSelected As Boolean)

        ' This call is required by the designer.
        'InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        _childrens = listOfNoeud
        Me.parent = Nothing
        Name = database
        IsChecked = False
        IsSelected = _isSelected

        setParentsToChilds()

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

    Public Overrides Property Name As String
        Get
            Return TreeViewItemBaseText
        End Get
        Set(value As String)
            TreeViewItemBaseText = value
        End Set
    End Property

    Public Overrides Property Childrens As ObservableCollection(Of TreeView.Noeud)
        Get
            Return MyBase.Childrens
        End Get
        Set(value As ObservableCollection(Of TreeView.Noeud))
            MyBase.Childrens = value
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

    Public Sub addSchema(ByVal UnSchema As DbSchemas)
        Childrens.Add(UnSchema)
    End Sub

    Public Sub clear()
        Childrens.Clear()
    End Sub

    Public Sub insert(ByVal index As Int32, ByVal UnSchema As DbSchemas)
        Childrens.Insert(index, UnSchema)
    End Sub

    Public Sub SetItem(ByVal index As Int32, ByVal UnSchema As DbSchemas)
        Childrens.Item(index) = UnSchema
    End Sub

    Public Sub RemoveAt(index As Integer)
        Childrens.RemoveAt(index)
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

    '------- --------
    '------- --------
    'Section publique
    '------- --------
    '------- --------

    Public Function Contains(item As DbSchemas) As Boolean
        Return Childrens.Contains(item)
    End Function

    Public Function IndexOf(item As DbSchemas) As Integer
        Return Childrens.IndexOf(item)
    End Function

    Public Function Remove(item As DbSchemas) As Boolean
        Return Childrens.Remove(item)
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
