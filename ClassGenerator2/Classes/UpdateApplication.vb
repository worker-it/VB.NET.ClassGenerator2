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



#End Region

Public Class UpdateApplication
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

    Public Shared Sub VersionVerification(Optional ByVal ConfirmationUpToDate As Boolean = False)

        Dim IfNewVersion As AppVersion = AppVersion.getAppVersionFromID(9)

        Dim ToUpdate As Boolean = False

        If (IfNewVersion.Major > My.Application.Info.Version.Major) Then
            ToUpdate = True
        ElseIf (IfNewVersion.Major = My.Application.Info.Version.Major) Then

            If (IfNewVersion.Minor > My.Application.Info.Version.Minor) Then
                ToUpdate = True
            ElseIf (IfNewVersion.Minor = My.Application.Info.Version.Minor) Then

                If (IfNewVersion.Build > My.Application.Info.Version.Build) Then
                    ToUpdate = True
                ElseIf (IfNewVersion.Build = My.Application.Info.Version.Build) Then

                    If (IfNewVersion.Revision > My.Application.Info.Version.Revision) Then
                        ToUpdate = True
                    ElseIf (IfNewVersion.Revision = My.Application.Info.Version.Revision) Then

                        If (ConfirmationUpToDate) Then

                            MsgBox("Vous êtes à jour!", MsgBoxStyle.OkOnly, "Vérification Version")


                        End If

                    End If
                End If
            End If
        End If

        If (ToUpdate) Then
            If (MsgBox("La version actuelle est la suivante :  V" & My.Application.Info.Version.ToString() & "." & vbCrLf &
                       "Une nouvelle version est disponible (V" & IfNewVersion.ApplicationVersion & ").  Voulez-vous l'installer?", vbYesNo, "Vérification Version") = vbYes) Then

                'Process.Start(IfNewVersion.AppSetupFilePath)

                'Dim fichier As String = "" & IfNewVersion.AppSetupFilePath & ""
                'If (IO.File.Exists(fichier)) Then
                '    Dim procInfos As New ProcessStartInfo("""" & IfNewVersion.AppSetupFilePath & """")

                '    procInfos.UseShellExecute = True
                '    procInfos.WindowStyle = ProcessWindowStyle.Normal

                '    procInfos.Verb = "runas"

                '    Process.Start(procInfos)

                Process.Start("C:\Windows\System32\msiexec.exe", "/i """ & IfNewVersion.AppSetupFilePath.Replace("setup.exe", "GammeDeFabSetup.msi") & """ REINSTALLMODE=amus")
                Application.Current.MainWindow.Close()
                'End If

            End If

        ElseIf (ConfirmationUpToDate) Then

            MsgBox("Vous êtes à jour!", MsgBoxStyle.OkOnly, "Vérification Version")

        End If

    End Sub

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
