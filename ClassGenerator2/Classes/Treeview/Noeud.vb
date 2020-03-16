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

#End Region

Namespace TreeView
    Public Class Noeud
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

        Protected _childrens As New ObservableCollection(Of Noeud)
        Private _TreeViewItemBaseText As String
        Private _parent As Noeud
        Private _isSelected As Boolean
        Private _isExpanded As Boolean
        Private _IsChecked As Boolean?

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

        Public Overridable Property Childrens As ObservableCollection(Of Noeud)
            Get
                If (_childrens Is Nothing) Then
                    Return Nothing
                Else
                    Return _childrens
                End If
            End Get
            Set(value As ObservableCollection(Of Noeud))
                _childrens = value
            End Set
        End Property

        Protected Property parent As Noeud
            Get
                Return _parent
            End Get
            Set(value As Noeud)
                _parent = value
                NotifyPropertyChanged("parent")
            End Set
        End Property

        Public Property IsSelected As Boolean
            Get
                Return Me._isSelected
            End Get
            Set(ByVal value As Boolean)

                If value <> Me._isSelected Then
                    Me._isSelected = value
                    NotifyPropertyChanged("IsSelected")
                End If
            End Set
        End Property

        Public Property IsExpanded As Boolean
            Get
                Return Me._isExpanded
            End Get
            Set(ByVal value As Boolean)

                If value <> Me._isExpanded Then
                    Me._isExpanded = value
                    NotifyPropertyChanged("IsExpanded")
                End If
            End Set
        End Property

        Public Property IsChecked As Boolean?
            Get
                Return _IsChecked
            End Get
            Set(value As Boolean?)
                SetIsChecked(value, True, True)
            End Set
        End Property

        Protected Property TreeViewItemBaseText As String
            Get
                Return _TreeViewItemBaseText
            End Get
            Set(value As String)
                _TreeViewItemBaseText = value
                NotifyPropertyChanged("TreeViewItemBaseText")
            End Set
        End Property

        Public ReadOnly Property HasItems As Boolean
            Get
                If (Childrens Is Nothing) Then
                    Return False
                Else
                    Return (Childrens.Count > 0)
                End If
            End Get
        End Property

        Public ReadOnly Property getNode(ByVal nodeName As String) As Noeud
            Get
                Dim result As Noeud = Nothing

                If (Me.TreeViewItemBaseText = nodeName) Then
                    result = Me
                Else
                    If (_childrens Is Nothing) Then
                        result = Nothing
                    Else
                        For Each element As Noeud In Childrens
                            result = element.getNode(nodeName)
                            If (result IsNot Nothing) Then
                                Exit For
                            End If
                        Next
                    End If
                End If

                Return result

            End Get
        End Property

        Public Overridable Property Name As String
            Get
                Return Me.TreeViewItemBaseText
            End Get
            Set(value As String)
                Me.TreeViewItemBaseText = value
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

        Private Sub SetIsChecked(ByVal value As Boolean?, ByVal updateChildren As Boolean, ByVal updateParent As Boolean)
            If (value <> _IsChecked) Or
                ((value Is Nothing) And (_IsChecked IsNot Nothing)) Or
                ((value IsNot Nothing) And (_IsChecked Is Nothing)) Then

                _IsChecked = value

                If (_childrens IsNot Nothing) AndAlso updateChildren AndAlso _IsChecked.HasValue Then
                    For Each item As Noeud In Me._childrens
                        item.SetIsChecked(_IsChecked, True, False)
                    Next
                End If

                If (_parent IsNot Nothing) AndAlso updateParent Then
                    _parent.VerifyCheckState()
                End If

                NotifyPropertyChanged("IsChecked")
            End If
        End Sub

        Private Sub VerifyCheckState()
            Dim state As Boolean? = Nothing

            For i As Integer = 0 To Me._childrens.Count - 1
                Dim current As Boolean? = Me._childrens(i).IsChecked

                If i = 0 Then
                    state = current
                ElseIf (state <> current) Or
                       ((state Is Nothing) And (current IsNot Nothing)) Or
                       ((state IsNot Nothing) And (current Is Nothing)) Then
                    state = Nothing
                    Exit For
                End If
            Next

            Me.SetIsChecked(state, False, True)
        End Sub

        Private Sub CollapseTreeviewItems()

            IsExpanded = False

            If (_childrens IsNot Nothing) Then
                For Each element As Noeud In Childrens
                    element.IsExpanded = False

                    If element.HasItems Then
                        element.CollapseTreeviewItems()
                    End If

                Next
            End If

        End Sub

        Private Sub ExpandTreeviewItems()

            IsExpanded = True

        End Sub

        Protected Sub setParentsToChilds()
            If (_childrens IsNot Nothing) Then
                For Each item As Noeud In _childrens
                    item.parent = Me
                Next
            End If
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

        Public Sub collapseAll()
            CollapseTreeviewItems()
        End Sub

        Public Sub expandAll()
            For Each item As Noeud In Childrens
                item.ExpandTreeviewItems()
            Next
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

        Public Sub NotifyPropertyChanged(ByVal propName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
        End Sub

#End Region

    End Class
End Namespace
