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

Imports System.Windows.Input

#End Region

Public Class RelayCommand
    Implements ICommand
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

    Private execute As Action(Of Object)
    Private canExecute As Predicate(Of Object)

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

    Public Sub New(ByVal execute As Action(Of Object))

        ' This call is required by the designer.
        'InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Me.New(execute, AddressOf DefaultCanExecute)
    End Sub

    Public Sub New(ByVal execute As Action(Of Object), ByVal canExecute As Predicate(Of Object))

        ' This call is required by the designer.
        'InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        If execute Is Nothing Then
            Throw New ArgumentNullException("execute")
        End If

        If canExecute Is Nothing Then
            Throw New ArgumentNullException("canExecute")
        End If

        Me.execute = execute
        Me.canExecute = canExecute
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

        Destroy()

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

    Private Sub ICommand_Execute(parameter As Object) Implements ICommand.Execute
        Me.execute(parameter)
    End Sub

    '------- --------
    '------- --------
    'Section publique
    '------- --------
    '------- --------

    Public Sub OnCanExecuteChanged()
        RaiseEvent CanExecuteChangedInternal(Me, EventArgs.Empty)
    End Sub

    Public Sub Destroy()
        Me.canExecute = Function() False
        Me.execute = Function()
                         Return Nothing
                     End Function
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

    Private Shared Function DefaultCanExecute(ByVal parameter As Object) As Boolean
        Return True
    End Function

    Private Function ICommand_CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
        Return Me.canExecute IsNot Nothing AndAlso Me.canExecute(parameter)
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

    Private Event CanExecuteChangedInternal As EventHandler

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

    Private Event ICommand_CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

    '------- --------
    '------- --------
    'Section publique
    '------- --------
    '------- --------

#End Region

End Class
