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

Imports System.Reflection

#End Region

Public Class Validator
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

    Private Shared PropertiesReflectionChace As Dictionary(Of Type, List(Of DependencyProperty)) = New Dictionary(Of Type, List(Of DependencyProperty))()

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

    Private Shared Function GetDPs(ByVal t As Type) As List(Of DependencyProperty)

        If PropertiesReflectionChace.ContainsKey(t) Then
            Return PropertiesReflectionChace(t)
        End If

        Dim properties As FieldInfo() = t.GetFields(BindingFlags.[Public] Or BindingFlags.GetProperty Or BindingFlags.[Static] Or BindingFlags.FlattenHierarchy)
        Dim dps As List(Of DependencyProperty) = New List(Of DependencyProperty)()

        ' we cycle and store only the dependency properties
        For Each field As FieldInfo In properties
            If field.FieldType = GetType(DependencyProperty) Then dps.Add(CType(field.GetValue(Nothing), DependencyProperty))
        Next

        PropertiesReflectionChace.Add(t, dps)
        Return dps
    End Function

    ''' <summary>
    ''' checks all the validation rule associated with objects,
    ''' forces the binding to execute all their validation rules
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function IsValid(ByVal mainObject As DependencyObject) As Boolean

        'Validate all the bindings on the parent
        Dim valid As Boolean = True

        'get the list of all the dependency properties, we can use a level of caching to avoid to use reflection
        'more than one time for each object
        'For Each dp As DependencyProperty In GetDPs(mainObject.[GetType]())

        '    If BindingOperations.IsDataBound(mainObject, dp) Then
        '        Dim binding As Binding = BindingOperations.GetBinding(mainObject, dp)

        '        If binding.ValidationRules.Count > 0 Then
        '            Dim expression As BindingExpression = BindingOperations.GetBindingExpression(mainObject, dp)

        '            Select Case binding.Mode
        '                Case BindingMode.OneTime, BindingMode.OneWay
        '                    expression.UpdateTarget()
        '                Case Else
        '                    expression.UpdateSource()
        '            End Select

        '            If expression.HasError Then
        '                valid = False
        '            End If
        '        End If
        '    End If
        'Next

        Dim i As Integer = 0

        While i <> VisualTreeHelper.GetChildrenCount(mainObject)
            Dim child As DependencyObject = VisualTreeHelper.GetChild(mainObject, i)

            If Not IsValid(child) Then
                valid = False
            End If

            i += 1
        End While

        Return valid
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
