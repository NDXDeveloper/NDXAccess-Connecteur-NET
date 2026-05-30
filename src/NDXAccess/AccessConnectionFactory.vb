Imports Microsoft.Extensions.Logging

Namespace NDXAccess

    ''' <summary>
    ''' Interface de la factory de connexions Access.
    ''' </summary>
    Public Interface IAccessConnectionFactory

        ''' <summary>Crée une connexion avec les options par défaut.</summary>
        Function CreateConnection() As IAccessConnection

        ''' <summary>Crée une connexion avec les options spécifiées.</summary>
        Function CreateConnection(options As AccessConnectionOptions) As IAccessConnection

        ''' <summary>Crée une connexion principale (jamais fermée automatiquement).</summary>
        Function CreatePrimaryConnection() As IAccessConnection

        ''' <summary>Crée une connexion en partant des options par défaut et en les ajustant.</summary>
        Function CreateConnection(configure As Action(Of AccessConnectionOptions)) As IAccessConnection

    End Interface

    ''' <summary>
    ''' Factory centralisée pour créer des instances de connexion Access avec des
    ''' options par défaut et un logging optionnel.
    ''' </summary>
    Public NotInheritable Class AccessConnectionFactory
        Implements IAccessConnectionFactory

        Private ReadOnly _defaultOptions As AccessConnectionOptions
        Private ReadOnly _loggerFactory As ILoggerFactory

        ''' <summary>Crée une factory avec des options par défaut.</summary>
        Public Sub New(defaultOptions As AccessConnectionOptions, Optional loggerFactory As ILoggerFactory = Nothing)
            ArgumentNullException.ThrowIfNull(defaultOptions)
            _defaultOptions = defaultOptions
            _loggerFactory = loggerFactory
        End Sub

        ''' <summary>Crée une factory à partir d'un chemin de base de données.</summary>
        Public Sub New(databasePath As String, Optional password As String = Nothing, Optional loggerFactory As ILoggerFactory = Nothing)
            Me.New(New AccessConnectionOptions With {
                .DatabasePath = databasePath,
                .Password = If(password, String.Empty)
            }, loggerFactory)
        End Sub

        ''' <inheritdoc/>
        Public Function CreateConnection() As IAccessConnection Implements IAccessConnectionFactory.CreateConnection
            Return CreateConnection(_defaultOptions)
        End Function

        ''' <inheritdoc/>
        Public Function CreateConnection(options As AccessConnectionOptions) As IAccessConnection Implements IAccessConnectionFactory.CreateConnection
            Dim logger = _loggerFactory?.CreateLogger(Of AccessConnection)()
            Return New AccessConnection(options, logger)
        End Function

        ''' <inheritdoc/>
        Public Function CreatePrimaryConnection() As IAccessConnection Implements IAccessConnectionFactory.CreatePrimaryConnection
            Dim options = CloneOptions(_defaultOptions)
            options.IsPrimaryConnection = True
            options.DisableAutoClose = True
            Return CreateConnection(options)
        End Function

        ''' <inheritdoc/>
        Public Function CreateConnection(configure As Action(Of AccessConnectionOptions)) As IAccessConnection Implements IAccessConnectionFactory.CreateConnection
            ArgumentNullException.ThrowIfNull(configure)
            Dim options = CloneOptions(_defaultOptions)
            configure(options)
            Return CreateConnection(options)
        End Function

        Private Shared Function CloneOptions(source As AccessConnectionOptions) As AccessConnectionOptions
            Return New AccessConnectionOptions With {
                .DatabasePath = source.DatabasePath,
                .Password = source.Password,
                .Provider = source.Provider,
                .ConnectionString = source.ConnectionString,
                .OpenExclusive = source.OpenExclusive,
                .OpenReadOnly = source.OpenReadOnly,
                .PersistSecurityInfo = source.PersistSecurityInfo,
                .SystemDatabasePath = source.SystemDatabasePath,
                .IsPrimaryConnection = source.IsPrimaryConnection,
                .AutoCloseTimeoutMs = source.AutoCloseTimeoutMs,
                .DisableAutoClose = source.DisableAutoClose,
                .ValidateProvider = source.ValidateProvider,
                .EnableRetryOnTransientErrors = source.EnableRetryOnTransientErrors,
                .MaxRetries = source.MaxRetries,
                .RetryBaseDelayMs = source.RetryBaseDelayMs,
                .TranslateErrors = source.TranslateErrors
            }
        End Function

    End Class

End Namespace
