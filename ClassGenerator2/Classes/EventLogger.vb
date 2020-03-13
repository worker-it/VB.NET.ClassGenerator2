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

Imports System.IO
Imports System.Xml

#End Region

Namespace Debugging

    Public Class EventLogger
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

        Private cfnCompleteLogFileName As String
        Private writeStartEnd As Boolean

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

            Me.New("C:\Temp")

        End Sub

        Public Sub New(ByVal logPath As String)

            Me.New(logPath, "EventLogger.log")

        End Sub

        Public Sub New(ByVal logPath As String,
                       ByVal logFileName As String)

            Me.New(logPath, logFileName, "", ".", "")

        End Sub

        Public Sub New(ByVal logPath As String,
                       ByVal logFileName As String,
                       ByVal aLogSource As String,
                       ByVal aLogMachine As String,
                       ByVal aLogWithinClass As String,
                       Optional ByVal ClearLog As Boolean = False)

            Dim xmlFile As FileStream
            Dim textWriter As XmlTextWriter
            writeStartEnd = False

            If (Right(logFileName, 3) <> "log") Then
                logFileName += ".log"
            End If

            If (Not Directory.Exists(logPath)) Then
                Directory.CreateDirectory(logPath)
            End If

            cfnCompleteLogFileName = logPath & "\" & logFileName

            If (Not File.Exists(cfnCompleteLogFileName)) Then
                writeStartEnd = True
            End If

            Me.logMachine = aLogMachine
            Me.logSource = aLogSource
            Me.logWithinClass = aLogWithinClass

            xmlFile = New FileStream(cfnCompleteLogFileName, FileMode.Append)
            textWriter = New XmlTextWriter(xmlFile, System.Text.Encoding.UTF8)

            If (writeStartEnd) Then

                textWriter.WriteStartDocument(True)
                textWriter.Formatting = Formatting.Indented
                textWriter.Indentation = 2
                textWriter.WriteStartElement("EventsLog")
                textWriter.WriteEndElement()

                textWriter.WriteEndDocument()
                writeStartEnd = False

                textWriter.Close()
                textWriter = Nothing

            Else

                textWriter.Close()
                textWriter = Nothing

            End If

            If (textWriter IsNot Nothing) And (ClearLog) Then

                textWriter.Close()
                textWriter = Nothing

            End If


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
        Public Property logSource As String
        Public Property logMachine As String
        Public Property logWithinClass As String
        Public ReadOnly Property completeLogFileName As String
            Get
                Return cfnCompleteLogFileName
            End Get
        End Property

        Public Property LoggerActive As Boolean = True


#End Region

        '************************************************************************************
        '                           P  R  O  C  É  D  U  R  E  S                            *
        '************************************************************************************
#Region "Procdures"

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

        Public Sub clearLogFile()

            Dim myXmlDocument As XmlDocument
            Dim myNodes As XmlNodeList

            If (File.Exists(cfnCompleteLogFileName)) Then

                Try
                    myXmlDocument = New XmlDocument()
                    myXmlDocument.Load(cfnCompleteLogFileName)
                    myNodes = myXmlDocument.GetElementsByTagName("EventsLog")

                    For Each aNode As XmlNode In myNodes
                        If (aNode.Name = "EventsLog") Then

                            aNode.RemoveAll()

                        End If
                    Next

                    myXmlDocument.Save(cfnCompleteLogFileName)
                    myXmlDocument = Nothing

                Catch ex As Exception

                End Try

            Else
                Throw New IOException("Event Log File does not exists!")
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

        Public Function writeLog(ByVal logEvent As String,
                                 ByVal stackTrace As String,
                                 ByVal eventSeverity As EventLogEntryType) As Boolean

            If (LoggerActive) Then
                If (File.Exists(cfnCompleteLogFileName)) Then

                    Dim myXmlDocument As New XmlDocument
                    Dim myNodes As XmlNodeList
                    Dim node As XmlNode
                    Dim nodeChild As XmlNode

                    Try

                        myXmlDocument.Load(cfnCompleteLogFileName)
                        myNodes = myXmlDocument.GetElementsByTagName("EventsLog")

                        For Each aNode As XmlNode In myNodes
                            If (aNode.Name = "EventsLog") Then

                                node = myXmlDocument.CreateElement("Event")
                                Dim anAttribute As XmlAttribute = myXmlDocument.CreateAttribute("Machine")
                                anAttribute.Value = logMachine
                                node.Attributes.Append(anAttribute)

                                anAttribute = myXmlDocument.CreateAttribute("Source")
                                anAttribute.Value = logSource
                                node.Attributes.Append(anAttribute)

                                anAttribute = Nothing

                                anAttribute = myXmlDocument.CreateAttribute("Class")
                                anAttribute.Value = logWithinClass
                                node.Attributes.Append(anAttribute)

                                anAttribute = myXmlDocument.CreateAttribute("EventSeverity")
                                anAttribute.Value = eventSeverity.ToString
                                node.Attributes.Append(anAttribute)

                                anAttribute = myXmlDocument.CreateAttribute("LogTime")
                                anAttribute.Value = Now().ToString("yyyyMMdd_hhmmss_FFFFF")
                                node.Attributes.Append(anAttribute)

                                nodeChild = myXmlDocument.CreateElement("EventMessage")
                                nodeChild.InnerText = logEvent
                                node.AppendChild(nodeChild)

                                nodeChild = Nothing

                                If (Not stackTrace Is Nothing) Then
                                    If (stackTrace <> "") Then
                                        nodeChild = myXmlDocument.CreateElement("StackTrace")
                                        nodeChild.InnerText = stackTrace
                                        node.AppendChild(nodeChild)

                                        nodeChild = Nothing
                                    End If
                                End If

                                aNode.AppendChild(node)

                            End If
                        Next

                        myXmlDocument.Save(cfnCompleteLogFileName)
                        myXmlDocument = Nothing

                    Catch ex As Exception
                        Return False
                    End Try

                Else
                    Throw New IOException("Event Log File does not exists!")
                End If

            End If

            Return True

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

End Namespace