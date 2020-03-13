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
Imports ModifiedControls.TreeView

#End Region

Public Class DbChampTable
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

    Private _DataType As String
    Private _IsPrimaryKey As Boolean = False

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

    Public Sub New(ByVal laColonne As String)

        ' This call is required by the designer.
        'InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Me.New(laColonne, "String")

    End Sub

    Public Sub New(ByVal laColonne As String, ByVal typeDonnees As String)

        ' This call is required by the designer.
        'InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Me.New(laColonne, typeDonnees, False)

    End Sub

    Public Sub New(ByVal laColonne As String, ByVal typeDonnees As String, ByVal _isSelected As Boolean)

        Me.New(laColonne, typeDonnees, _isSelected, False)

    End Sub

    Public Sub New(ByVal laColonne As String, ByVal typeDonnees As String, ByVal _isSelected As Boolean, ByVal _IsPK As Boolean)

        Childrens = Nothing
        DataType = typeDonnees
        parent = Nothing
        Name = laColonne
        IsChecked = False
        IsSelected = _isSelected
        _IsPrimaryKey = _IsPK

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

    Public Overloads Property Name As String
        Get
            Return TreeViewItemBaseText
        End Get
        Set(value As String)
            TreeViewItemBaseText = value
        End Set
    End Property

    Public Overloads Property Childrens As ObservableCollection(Of TreeView.Noeud)
        Get
            Return MyBase.Childrens
        End Get
        Set(value As ObservableCollection(Of TreeView.Noeud))
            MyBase.Childrens = value
        End Set
    End Property

    Public Property DataType As String
        Get
            Return _DataType
        End Get
        Set(value As String)
            _DataType = value
        End Set
    End Property

    Public Property IsPrimaryKey As Boolean
        Get
            Return _IsPrimaryKey
        End Get
        Set(value As Boolean)
            _IsPrimaryKey = value
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
