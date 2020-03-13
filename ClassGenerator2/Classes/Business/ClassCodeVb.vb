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

Imports System.IO

#End Region

Public Class ClassCodeVb
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

    Private strClassName As String
    Private strClassCode As String

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

    Public Sub New(ByVal _ClassName As String,
                   ByVal _ClassCode As String)

        ' This call is required by the designer.
        'InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        strClassCode = _ClassCode
        strClassName = _ClassName

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

    Public Property ClassName As String
        Get
            Return strClassName
        End Get
        Set(value As String)
            strClassName = value
        End Set
    End Property

    Public Property ClassCode As String
        Get
            Return strClassCode
        End Get
        Set(value As String)
            strClassCode = value
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

    Public Sub CreateVbFile(ByVal path As String)


        Dim fileName As String = path & "\" & strClassName & ".vb"
        Dim errorMessage As String = ""
        Dim writeFile As Boolean = False

        If (File.Exists(fileName)) Then
            If (MsgBox("Le fichier (" & fileName & ") est déjà existant.  Voulez-vous le remplacer?", CType(MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton1, Global.Microsoft.VisualBasic.MsgBoxStyle), "Remplacer le fichier") = vbYes) Then
                writeFile = True
            Else
                MsgBox("Opération avortée par l'utilisateur!")
            End If
        Else
            writeFile = True
        End If

        If (writeFile) Then
            Try

                Dim sw As New StreamWriter(fileName)
                sw.Write(strClassCode)
                sw.Close()

            Catch ioex As IOException
                errorMessage = "IOException : " & ioex.Message
            Catch ex As Exception
                errorMessage = "Exception : " & ex.Message
            Finally

                If (File.Exists(fileName)) Then
                    'MsgBox("Fichier sauvegardé avec succès!")
                ElseIf (errorMessage <> "") Then
                    MsgBox(errorMessage)
                Else
                    MsgBox("Une erreur est survenue lors de la sauvegarde!")
                End If

            End Try
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
