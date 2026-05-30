Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Threading

Namespace NDXAccess

    ''' <summary>
    ''' Vérifie l'état de santé d'une base Access : connectivité, taille du fichier
    ''' par rapport à la limite stricte de 2 Go, version du moteur et nombre de tables.
    ''' </summary>
    Public NotInheritable Class AccessHealthCheck

        Private ReadOnly _connectionFactory As IAccessConnectionFactory

        ''' <summary>Crée un health check à partir d'une factory de connexions.</summary>
        Public Sub New(connectionFactory As IAccessConnectionFactory)
            ArgumentNullException.ThrowIfNull(connectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        ''' <summary>Vérifie la santé de la base (synchrone).</summary>
        Public Function CheckHealth() As HealthCheckResult
            Dim startTime = DateTime.UtcNow
            Try
                Using connection = _connectionFactory.CreateConnection()
                    connection.Open()
                    Dim info = BuildDatabaseInfo(connection)
                    Return BuildHealthyResult(info, DateTime.UtcNow - startTime)
                End Using
            Catch ex As Exception
                Return New HealthCheckResult(False, $"Erreur de connexion : {ex.Message}", DateTime.UtcNow - startTime, ex)
            End Try
        End Function

        ''' <summary>Vérifie la santé de la base (async de façade).</summary>
        Public Async Function CheckHealthAsync(Optional cancellationToken As CancellationToken = Nothing) As Task(Of HealthCheckResult)
            Dim startTime = DateTime.UtcNow
            Try
                Using connection = _connectionFactory.CreateConnection()
                    Await connection.OpenAsync(cancellationToken).ConfigureAwait(False)
                    Dim info = BuildDatabaseInfo(connection)
                    Return BuildHealthyResult(info, DateTime.UtcNow - startTime)
                End Using
            Catch ex As Exception
                Return New HealthCheckResult(False, $"Erreur de connexion : {ex.Message}", DateTime.UtcNow - startTime, ex)
            End Try
        End Function

        ''' <summary>Récupère les informations détaillées de la base (synchrone).</summary>
        Public Function GetDatabaseInfo() As DatabaseInfo
            Using connection = _connectionFactory.CreateConnection()
                connection.Open()
                Return BuildDatabaseInfo(connection)
            End Using
        End Function

        ''' <summary>Récupère les informations détaillées de la base (async de façade).</summary>
        Public Async Function GetDatabaseInfoAsync(Optional cancellationToken As CancellationToken = Nothing) As Task(Of DatabaseInfo)
            Using connection = _connectionFactory.CreateConnection()
                Await connection.OpenAsync(cancellationToken).ConfigureAwait(False)
                Return BuildDatabaseInfo(connection)
            End Using
        End Function

        Private Shared Function BuildHealthyResult(info As DatabaseInfo, duration As TimeSpan) As HealthCheckResult
            Dim message = "Connexion Access fonctionnelle"
            Dim healthy = True

            If info.IsApproachingSizeLimit Then
                message &= $" — ATTENTION : {info.UsagePercent:F1}% de la limite de 2 Go atteinte ({info.FileSizeMegabytes:F1} Mo). Pensez à compacter ou archiver."
            End If

            Return New HealthCheckResult(healthy, message, duration) With {.DatabaseInfo = info}
        End Function

        Private Shared Function BuildDatabaseInfo(connection As IAccessConnection) As DatabaseInfo
            Dim ole = connection.Connection
            Dim info As New DatabaseInfo()

            ' Provider et version du moteur.
            Try
                Dim builder As New OleDbConnectionStringBuilder(ole.ConnectionString)
                info.Provider = builder.Provider
                If builder.ContainsKey("Data Source") Then
                    info.FilePath = TryCast(builder("Data Source"), String)
                End If
                If builder.ContainsKey("Mode") Then
                    info.IsReadOnly = String.Equals(TryCast(builder("Mode"), String), "Read", StringComparison.OrdinalIgnoreCase)
                End If
            Catch
            End Try

            ' Version du moteur ACE (DLL) et version du format de fichier (Jet/ACE).
            info.EngineVersion = AccessProviderHelper.GetEngineVersion(info.Provider)
            Try
                info.FileFormatVersion = ole.ServerVersion
            Catch
                info.FileFormatVersion = "Inconnue"
            End Try

            ' Taille du fichier et usage par rapport à la limite de 2 Go.
            If Not String.IsNullOrWhiteSpace(info.FilePath) AndAlso File.Exists(info.FilePath) Then
                info.FileSizeBytes = New FileInfo(info.FilePath).Length
            End If

            ' Nombre de tables utilisateur.
            Try
                Using tables = ole.GetSchema("Tables")
                    Dim count = 0
                    For Each row As DataRow In tables.Rows
                        Dim tableType = TryCast(row("TABLE_TYPE"), String)
                        If String.Equals(tableType, "TABLE", StringComparison.OrdinalIgnoreCase) Then
                            count += 1
                        End If
                    Next
                    info.UserTableCount = count
                End Using
            Catch
                info.UserTableCount = -1
            End Try

            Return info
        End Function

    End Class

    ''' <summary>Résultat d'un contrôle de santé.</summary>
    Public NotInheritable Class HealthCheckResult

        Public Sub New(isHealthy As Boolean, message As String, responseTime As TimeSpan, Optional exception As Exception = Nothing)
            Me.IsHealthy = isHealthy
            Me.Message = message
            Me.ResponseTime = responseTime
            Me.Exception = exception
        End Sub

        ''' <summary>La base est-elle accessible et opérationnelle ?</summary>
        Public ReadOnly Property IsHealthy As Boolean

        ''' <summary>Message décrivant le résultat (inclut l'avertissement de taille le cas échéant).</summary>
        Public ReadOnly Property Message As String

        ''' <summary>Temps de réponse mesuré.</summary>
        Public ReadOnly Property ResponseTime As TimeSpan

        ''' <summary>Exception capturée en cas d'échec (Nothing sinon).</summary>
        Public ReadOnly Property Exception As Exception

        ''' <summary>Informations détaillées sur la base (Nothing en cas d'échec).</summary>
        Public Property DatabaseInfo As DatabaseInfo

    End Class

    ''' <summary>Informations détaillées sur une base Access.</summary>
    Public NotInheritable Class DatabaseInfo

        ''' <summary>Nom du provider OLE DB utilisé.</summary>
        Public Property Provider As String = String.Empty

        ''' <summary>Version du moteur ACE (DLL), ex. "16.0.5011.1000" (ou la génération "16.0").</summary>
        Public Property EngineVersion As String = String.Empty

        ''' <summary>Version du format de fichier (Jet/ACE), ex. "04.00.0000". Différent du moteur.</summary>
        Public Property FileFormatVersion As String = String.Empty

        ''' <summary>Chemin du fichier de base.</summary>
        Public Property FilePath As String = String.Empty

        ''' <summary>Taille du fichier en octets.</summary>
        Public Property FileSizeBytes As Long

        ''' <summary>Taille du fichier en mégaoctets.</summary>
        Public ReadOnly Property FileSizeMegabytes As Double
            Get
                Return FileSizeBytes / 1024.0 / 1024.0
            End Get
        End Property

        ''' <summary>Limite stricte de taille du moteur (2 Go).</summary>
        Public ReadOnly Property MaxSizeBytes As Long
            Get
                Return AccessConnectionOptions.MaxDatabaseSizeBytes
            End Get
        End Property

        ''' <summary>Pourcentage d'utilisation par rapport à la limite de 2 Go.</summary>
        Public ReadOnly Property UsagePercent As Double
            Get
                If MaxSizeBytes = 0 Then Return 0
                Return FileSizeBytes / MaxSizeBytes * 100.0
            End Get
        End Property

        ''' <summary>Vrai si la base dépasse 90% de la limite de 2 Go.</summary>
        Public ReadOnly Property IsApproachingSizeLimit As Boolean
            Get
                Return UsagePercent >= 90.0
            End Get
        End Property

        ''' <summary>La base est-elle ouverte en lecture seule ?</summary>
        Public Property IsReadOnly As Boolean

        ''' <summary>Nombre de tables utilisateur (-1 si indéterminé).</summary>
        Public Property UserTableCount As Integer

    End Class

End Namespace
