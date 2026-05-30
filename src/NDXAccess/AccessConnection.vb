Imports System.Collections.Concurrent
Imports System.Data
Imports System.Data.OleDb
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Threading
Imports Microsoft.Extensions.Logging

Namespace NDXAccess

    ''' <summary>
    ''' Connexion Microsoft Access (.accdb / .mdb) moderne basée sur OLE DB (provider ACE).
    ''' Gère le cycle de vie, les transactions, la fermeture automatique, le logging,
    ''' les requêtes paramétrées, les requêtes enregistrées et le compactage.
    ''' </summary>
    ''' <remarks>
    ''' <para>Windows uniquement (le provider ACE n'existe pas ailleurs).</para>
    ''' <para>Auteur : Nicolas DEOUX - NDXDev 2026</para>
    ''' </remarks>
    Public NotInheritable Class AccessConnection
        Implements IAccessConnection

#Region "Constantes et champs"

        Private Const MaxActionHistorySize As Integer = 5

        Private Shared _connectionCounter As Integer

        Private ReadOnly _actionLock As New Object()
        Private ReadOnly _disposeLock As New Object()
        Private ReadOnly _logger As ILogger(Of AccessConnection)
        Private ReadOnly _options As AccessConnectionOptions
        Private ReadOnly _autoCloseTimer As Timer
        Private ReadOnly _actionHistory As New List(Of String)(MaxActionHistorySize + 1)

        Private _connection As OleDbConnection
        Private _transaction As OleDbTransaction
        Private _lastAction As String = String.Empty
        Private _isDisposed As Boolean
        Private _isTransactionActive As Boolean

#End Region

#Region "Propriétés"

        ''' <summary>
        ''' Version (informationnelle) de la bibliothèque NDXAccess, ex. "1.1.0".
        ''' Raccourci vers <see cref="NDXAccessInfo.InformationalVersion"/>. À ne pas confondre
        ''' avec <see cref="DatabaseInfo.EngineVersion"/> (version du moteur Access).
        ''' </summary>
        Public Shared ReadOnly Property Version As String
            Get
                Return NDXAccessInfo.InformationalVersion
            End Get
        End Property

        ''' <inheritdoc/>
        Public ReadOnly Property ProviderName As String Implements IAccessConnection.ProviderName
            Get
                Return _options.ResolveProviderName()
            End Get
        End Property

        ''' <inheritdoc/>
        Public ReadOnly Property EngineVersion As String Implements IAccessConnection.EngineVersion
            Get
                Return AccessProviderHelper.GetEngineVersion(_options.ResolveProviderName())
            End Get
        End Property

        ''' <inheritdoc/>
        Public ReadOnly Property Id As Integer Implements IAccessConnection.Id

        ''' <inheritdoc/>
        Public ReadOnly Property CreatedAt As DateTime Implements IAccessConnection.CreatedAt

        ''' <inheritdoc/>
        Public ReadOnly Property State As ConnectionState Implements IAccessConnection.State
            Get
                Return If(_connection Is Nothing, ConnectionState.Closed, _connection.State)
            End Get
        End Property

        ''' <inheritdoc/>
        Public ReadOnly Property IsTransactionActive As Boolean Implements IAccessConnection.IsTransactionActive
            Get
                Return _isTransactionActive
            End Get
        End Property

        ''' <inheritdoc/>
        Public ReadOnly Property IsPrimaryConnection As Boolean Implements IAccessConnection.IsPrimaryConnection
            Get
                Return _options.IsPrimaryConnection
            End Get
        End Property

        ''' <inheritdoc/>
        Public ReadOnly Property Connection As OleDbConnection Implements IAccessConnection.Connection
            Get
                Return _connection
            End Get
        End Property

        ''' <inheritdoc/>
        Public ReadOnly Property Transaction As OleDbTransaction Implements IAccessConnection.Transaction
            Get
                Return _transaction
            End Get
        End Property

        ''' <inheritdoc/>
        Public ReadOnly Property LastAction As String Implements IAccessConnection.LastAction
            Get
                SyncLock _actionLock
                    Return _lastAction
                End SyncLock
            End Get
        End Property

        ''' <inheritdoc/>
        Public ReadOnly Property ActionHistory As IReadOnlyList(Of String) Implements IAccessConnection.ActionHistory
            Get
                SyncLock _actionLock
                    Return _actionHistory.ToList().AsReadOnly()
                End SyncLock
            End Get
        End Property

#End Region

#Region "Constructeurs"

        ''' <summary>
        ''' Crée une connexion Access avec les options spécifiées.
        ''' </summary>
        ''' <param name="options">Options de configuration.</param>
        ''' <param name="logger">Logger optionnel.</param>
        Public Sub New(options As AccessConnectionOptions, Optional logger As ILogger(Of AccessConnection) = Nothing)
            ArgumentNullException.ThrowIfNull(options)

            _options = options
            _logger = logger

            ' Détection x86/x64 : échoue tôt avec un message clair en cas de mismatch.
            If options.ValidateProvider Then
                AccessProviderHelper.EnsureProviderAvailable(options.ResolveProviderName())
            End If

            Id = Interlocked.Increment(_connectionCounter)
            CreatedAt = DateTime.UtcNow

            _connection = New OleDbConnection(options.BuildConnectionString())
            AddHandler _connection.StateChange, AddressOf OnConnectionStateChanged

            If Not options.IsPrimaryConnection AndAlso Not options.DisableAutoClose Then
                _autoCloseTimer = New Timer(AddressOf OnAutoCloseTimerElapsed, Nothing, Timeout.Infinite, Timeout.Infinite)
            End If

            LogAction("New", $"Nouvelle connexion {If(options.IsPrimaryConnection, "principale", "secondaire")} créée")
        End Sub

        ''' <summary>
        ''' Crée une connexion Access à partir du chemin d'un fichier de base.
        ''' </summary>
        ''' <param name="databasePath">Chemin du fichier .accdb / .mdb.</param>
        ''' <param name="password">Mot de passe optionnel.</param>
        ''' <param name="isPrimary">Connexion principale.</param>
        ''' <param name="logger">Logger optionnel.</param>
        Public Sub New(databasePath As String, Optional password As String = Nothing, Optional isPrimary As Boolean = False, Optional logger As ILogger(Of AccessConnection) = Nothing)
            Me.New(New AccessConnectionOptions With {
                .DatabasePath = databasePath,
                .Password = If(password, String.Empty),
                .IsPrimaryConnection = isPrimary
            }, logger)
        End Sub

#End Region

#Region "Cycle de vie"

        ''' <inheritdoc/>
        Public Sub Open() Implements IAccessConnection.Open
            ThrowIfDisposed()
            If _connection Is Nothing Then
                Throw New InvalidOperationException("La connexion n'est pas initialisée.")
            End If

            ExecuteResilient(Function()
                                  OpenInternal()
                                  Return True
                              End Function)
            ResetAutoCloseTimer()
        End Sub

        ''' <inheritdoc/>
        Public Async Function OpenAsync(Optional cancellationToken As CancellationToken = Nothing) As Task Implements IAccessConnection.OpenAsync
            ThrowIfDisposed()
            If _connection Is Nothing Then
                Throw New InvalidOperationException("La connexion n'est pas initialisée.")
            End If

            Await ExecuteResilientAsync(Async Function()
                                            Await OpenInternalAsync(cancellationToken).ConfigureAwait(False)
                                            Return True
                                        End Function, cancellationToken).ConfigureAwait(False)
            ResetAutoCloseTimer()
        End Function

        ''' <summary>Ouverture brute (sans retry/traduction) — utilisée dans les wrappers résilients.</summary>
        Private Sub OpenInternal()
            If _connection.State <> ConnectionState.Open Then
                _connection.Open()
                LogAction("Open", "Connexion ouverte")
            End If
        End Sub

        Private Async Function OpenInternalAsync(cancellationToken As CancellationToken) As Task
            If _connection.State <> ConnectionState.Open Then
                Await _connection.OpenAsync(cancellationToken).ConfigureAwait(False)
                LogAction("OpenAsync", "Connexion ouverte (async de façade)")
            End If
        End Function

        ''' <inheritdoc/>
        Public Sub Close() Implements IAccessConnection.Close
            If _connection IsNot Nothing AndAlso _connection.State <> ConnectionState.Closed Then
                _connection.Close()
                LogAction("Close", "Connexion fermée")
            End If
        End Sub

        ''' <inheritdoc/>
        Public Async Function CloseAsync() As Task Implements IAccessConnection.CloseAsync
            If _connection IsNot Nothing AndAlso _connection.State <> ConnectionState.Closed Then
                Await _connection.CloseAsync().ConfigureAwait(False)
                LogAction("CloseAsync", "Connexion fermée (async de façade)")
            End If
        End Function

#End Region

#Region "Transactions"

        ''' <inheritdoc/>
        Public Function BeginTransaction(Optional isolationLevel As IsolationLevel = IsolationLevel.ReadCommitted) As Boolean Implements IAccessConnection.BeginTransaction
            ThrowIfDisposed()
            Try
                EnsureConnectionOpen()
                _transaction = _connection.BeginTransaction(isolationLevel)
                _isTransactionActive = True
                LogAction("BeginTransaction", $"Transaction démarrée (IsolationLevel: {isolationLevel})")
                Return True
            Catch ex As OleDbException
                _logger?.LogError(ex, "Erreur lors du démarrage de la transaction")
                _transaction = Nothing
                _isTransactionActive = False
                If _options.TranslateErrors Then Throw AccessErrorTranslator.Translate(ex)
                Throw
            Catch ex As Exception
                _logger?.LogError(ex, "Erreur lors du démarrage de la transaction")
                _transaction = Nothing
                _isTransactionActive = False
                Throw
            End Try
        End Function

        ''' <inheritdoc/>
        Public Async Function BeginTransactionAsync(Optional isolationLevel As IsolationLevel = IsolationLevel.ReadCommitted, Optional cancellationToken As CancellationToken = Nothing) As Task(Of Boolean) Implements IAccessConnection.BeginTransactionAsync
            ThrowIfDisposed()
            Try
                Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
                _transaction = CType(Await _connection.BeginTransactionAsync(isolationLevel, cancellationToken).ConfigureAwait(False), OleDbTransaction)
                _isTransactionActive = True
                LogAction("BeginTransactionAsync", $"Transaction démarrée (IsolationLevel: {isolationLevel})")
                Return True
            Catch ex As OleDbException
                _logger?.LogError(ex, "Erreur lors du démarrage de la transaction")
                _transaction = Nothing
                _isTransactionActive = False
                If _options.TranslateErrors Then Throw AccessErrorTranslator.Translate(ex)
                Throw
            Catch ex As Exception
                _logger?.LogError(ex, "Erreur lors du démarrage de la transaction")
                _transaction = Nothing
                _isTransactionActive = False
                Throw
            End Try
        End Function

        ''' <inheritdoc/>
        Public Sub Commit() Implements IAccessConnection.Commit
            If _transaction Is Nothing Then Return
            Try
                _transaction.Commit()
                LogAction("Commit", "Transaction validée")
            Finally
                DisposeTransaction()
            End Try
        End Sub

        ''' <inheritdoc/>
        Public Async Function CommitAsync(Optional cancellationToken As CancellationToken = Nothing) As Task Implements IAccessConnection.CommitAsync
            If _transaction Is Nothing Then Return
            Try
                Await _transaction.CommitAsync(cancellationToken).ConfigureAwait(False)
                LogAction("CommitAsync", "Transaction validée (async de façade)")
            Finally
                DisposeTransaction()
            End Try
        End Function

        ''' <inheritdoc/>
        Public Sub Rollback() Implements IAccessConnection.Rollback
            If _transaction Is Nothing Then Return
            Try
                _transaction.Rollback()
                LogAction("Rollback", "Transaction annulée")
            Finally
                DisposeTransaction()
            End Try
        End Sub

        ''' <inheritdoc/>
        Public Async Function RollbackAsync(Optional cancellationToken As CancellationToken = Nothing) As Task Implements IAccessConnection.RollbackAsync
            If _transaction Is Nothing Then Return
            Try
                Await _transaction.RollbackAsync(cancellationToken).ConfigureAwait(False)
                LogAction("RollbackAsync", "Transaction annulée (async de façade)")
            Finally
                DisposeTransaction()
            End Try
        End Function

        Private Sub DisposeTransaction()
            Try
                _transaction?.Dispose()
            Catch
            End Try
            _transaction = Nothing
            _isTransactionActive = False
        End Sub

#End Region

#Region "Exécution de requêtes"

        ''' <inheritdoc/>
        Public Function ExecuteNonQuery(sql As String, ParamArray parameters As Object()) As Integer Implements IAccessConnection.ExecuteNonQuery
            ThrowIfDisposed()
            Return ExecuteResilient(Function()
                                        EnsureConnectionOpen()
                                        Using command = CreateCommandInternal(sql, parameters)
                                            Dim result = command.ExecuteNonQuery()
                                            ResetAutoCloseTimer()
                                            Return result
                                        End Using
                                    End Function)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteNonQueryAsync(sql As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of Integer) Implements IAccessConnection.ExecuteNonQueryAsync
            ThrowIfDisposed()
            Return ExecuteResilientAsync(Async Function()
                                             Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
                                             Using command = CreateCommandInternal(sql, parameters)
                                                 Dim result = Await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(False)
                                                 ResetAutoCloseTimer()
                                                 Return result
                                             End Using
                                         End Function, cancellationToken)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteScalar(Of T)(sql As String, ParamArray parameters As Object()) As T Implements IAccessConnection.ExecuteScalar
            ThrowIfDisposed()
            Return ExecuteResilient(Function()
                                        EnsureConnectionOpen()
                                        Using command = CreateCommandInternal(sql, parameters)
                                            Dim result = command.ExecuteScalar()
                                            ResetAutoCloseTimer()
                                            Return ConvertScalar(Of T)(result)
                                        End Using
                                    End Function)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteScalarAsync(Of T)(sql As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of T) Implements IAccessConnection.ExecuteScalarAsync
            ThrowIfDisposed()
            Return ExecuteResilientAsync(Async Function()
                                             Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
                                             Using command = CreateCommandInternal(sql, parameters)
                                                 Dim result = Await command.ExecuteScalarAsync(cancellationToken).ConfigureAwait(False)
                                                 ResetAutoCloseTimer()
                                                 Return ConvertScalar(Of T)(result)
                                             End Using
                                         End Function, cancellationToken)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteQuery(sql As String, ParamArray parameters As Object()) As DataTable Implements IAccessConnection.ExecuteQuery
            ThrowIfDisposed()
            Return ExecuteResilient(Function()
                                        EnsureConnectionOpen()
                                        Using command = CreateCommandInternal(sql, parameters)
                                            Using reader = command.ExecuteReader()
                                                Dim table As New DataTable()
                                                table.Load(reader)
                                                ResetAutoCloseTimer()
                                                Return table
                                            End Using
                                        End Using
                                    End Function)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteQueryAsync(sql As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of DataTable) Implements IAccessConnection.ExecuteQueryAsync
            ThrowIfDisposed()
            Return ExecuteResilientAsync(Async Function()
                                             Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
                                             Using command = CreateCommandInternal(sql, parameters)
                                                 Using reader = Await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(False)
                                                     Dim table As New DataTable()
                                                     table.Load(reader)
                                                     ResetAutoCloseTimer()
                                                     Return table
                                                 End Using
                                             End Using
                                         End Function, cancellationToken)
        End Function

        ''' <inheritdoc/>
        ''' <remarks>
        ''' La commande sous-jacente n'est pas disposée explicitement (le reader en a besoin) :
        ''' elle est récupérée par le GC à la fermeture du reader. Pour une libération
        ''' déterministe, préférez <see cref="ExecuteQuery"/> (DataTable) ou
        ''' <c>ExecuteQuery(Of T)</c> (objets typés).
        ''' </remarks>
        Public Function ExecuteReader(sql As String, ParamArray parameters As Object()) As OleDbDataReader Implements IAccessConnection.ExecuteReader
            ThrowIfDisposed()
            Return ExecuteResilient(Function()
                                        EnsureConnectionOpen()
                                        Dim command = CreateCommandInternal(sql, parameters)
                                        Dim reader = command.ExecuteReader()
                                        ResetAutoCloseTimer()
                                        Return reader
                                    End Function)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteReaderAsync(sql As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of OleDbDataReader) Implements IAccessConnection.ExecuteReaderAsync
            ThrowIfDisposed()
            Return ExecuteResilientAsync(Async Function()
                                             Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
                                             Dim command = CreateCommandInternal(sql, parameters)
                                             Dim reader = CType(Await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(False), OleDbDataReader)
                                             ResetAutoCloseTimer()
                                             Return reader
                                         End Function, cancellationToken)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteQuery(Of T As New)(sql As String, ParamArray parameters As Object()) As List(Of T) Implements IAccessConnection.ExecuteQuery
            ThrowIfDisposed()
            Return ExecuteResilient(Function()
                                        EnsureConnectionOpen()
                                        Using command = CreateCommandInternal(sql, parameters)
                                            Using reader = command.ExecuteReader()
                                                Dim list = MapReader(Of T)(reader)
                                                ResetAutoCloseTimer()
                                                Return list
                                            End Using
                                        End Using
                                    End Function)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteQueryAsync(Of T As New)(sql As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of List(Of T)) Implements IAccessConnection.ExecuteQueryAsync
            ThrowIfDisposed()
            Return ExecuteResilientAsync(Async Function()
                                             Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
                                             Using command = CreateCommandInternal(sql, parameters)
                                                 Using reader = Await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(False)
                                                     Dim list = MapReader(Of T)(reader)
                                                     ResetAutoCloseTimer()
                                                     Return list
                                                 End Using
                                             End Using
                                         End Function, cancellationToken)
        End Function

#End Region

#Region "Paramètres nommés (@nom -> ?)"

        ''' <inheritdoc/>
        Public Function ExecuteNonQueryNamed(sql As String, parameters As IDictionary(Of String, Object)) As Integer Implements IAccessConnection.ExecuteNonQueryNamed
            Dim t = NamedParameterParser.Translate(sql, parameters)
            Return ExecuteNonQuery(t.Sql, t.Values)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteNonQueryNamedAsync(sql As String, parameters As IDictionary(Of String, Object), Optional cancellationToken As CancellationToken = Nothing) As Task(Of Integer) Implements IAccessConnection.ExecuteNonQueryNamedAsync
            Dim t = NamedParameterParser.Translate(sql, parameters)
            Return ExecuteNonQueryAsync(t.Sql, t.Values, cancellationToken)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteScalarNamed(Of T)(sql As String, parameters As IDictionary(Of String, Object)) As T Implements IAccessConnection.ExecuteScalarNamed
            Dim parsed = NamedParameterParser.Translate(sql, parameters)
            Return ExecuteScalar(Of T)(parsed.Sql, parsed.Values)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteScalarNamedAsync(Of T)(sql As String, parameters As IDictionary(Of String, Object), Optional cancellationToken As CancellationToken = Nothing) As Task(Of T) Implements IAccessConnection.ExecuteScalarNamedAsync
            Dim parsed = NamedParameterParser.Translate(sql, parameters)
            Return ExecuteScalarAsync(Of T)(parsed.Sql, parsed.Values, cancellationToken)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteQueryNamed(sql As String, parameters As IDictionary(Of String, Object)) As DataTable Implements IAccessConnection.ExecuteQueryNamed
            Dim t = NamedParameterParser.Translate(sql, parameters)
            Return ExecuteQuery(t.Sql, t.Values)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteQueryNamedAsync(sql As String, parameters As IDictionary(Of String, Object), Optional cancellationToken As CancellationToken = Nothing) As Task(Of DataTable) Implements IAccessConnection.ExecuteQueryNamedAsync
            Dim t = NamedParameterParser.Translate(sql, parameters)
            Return ExecuteQueryAsync(t.Sql, t.Values, cancellationToken)
        End Function

#End Region

#Region "Requêtes enregistrées (stored queries)"

        ''' <inheritdoc/>
        Public Function ExecuteStoredQuery(queryName As String, ParamArray parameters As Object()) As DataTable Implements IAccessConnection.ExecuteStoredQuery
            ThrowIfDisposed()
            Return ExecuteResilient(Function()
                                        EnsureConnectionOpen()
                                        Using command = CreateStoredQueryCommand(queryName, parameters)
                                            Using reader = command.ExecuteReader()
                                                Dim table As New DataTable()
                                                table.Load(reader)
                                                ResetAutoCloseTimer()
                                                Return table
                                            End Using
                                        End Using
                                    End Function)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteStoredQueryAsync(queryName As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of DataTable) Implements IAccessConnection.ExecuteStoredQueryAsync
            ThrowIfDisposed()
            Return ExecuteResilientAsync(Async Function()
                                             Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
                                             Using command = CreateStoredQueryCommand(queryName, parameters)
                                                 Using reader = Await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(False)
                                                     Dim table As New DataTable()
                                                     table.Load(reader)
                                                     ResetAutoCloseTimer()
                                                     Return table
                                                 End Using
                                             End Using
                                         End Function, cancellationToken)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteStoredQueryNonQuery(queryName As String, ParamArray parameters As Object()) As Integer Implements IAccessConnection.ExecuteStoredQueryNonQuery
            ThrowIfDisposed()
            Return ExecuteResilient(Function()
                                        EnsureConnectionOpen()
                                        Using command = CreateStoredQueryCommand(queryName, parameters)
                                            Dim result = command.ExecuteNonQuery()
                                            ResetAutoCloseTimer()
                                            Return result
                                        End Using
                                    End Function)
        End Function

        ''' <inheritdoc/>
        Public Function ExecuteStoredQueryNonQueryAsync(queryName As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of Integer) Implements IAccessConnection.ExecuteStoredQueryNonQueryAsync
            ThrowIfDisposed()
            Return ExecuteResilientAsync(Async Function()
                                             Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
                                             Using command = CreateStoredQueryCommand(queryName, parameters)
                                                 Dim result = Await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(False)
                                                 ResetAutoCloseTimer()
                                                 Return result
                                             End Using
                                         End Function, cancellationToken)
        End Function

#End Region

#Region "Helpers de schéma"

        ''' <inheritdoc/>
        Public Function TableExists(tableName As String) As Boolean Implements IAccessConnection.TableExists
            ThrowIfDisposed()
            EnsureConnectionOpen()
            Return TableExistsCore(tableName)
        End Function

        ''' <inheritdoc/>
        Public Async Function TableExistsAsync(tableName As String, Optional cancellationToken As CancellationToken = Nothing) As Task(Of Boolean) Implements IAccessConnection.TableExistsAsync
            ThrowIfDisposed()
            Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
            Return TableExistsCore(tableName)
        End Function

        ''' <inheritdoc/>
        Public Function GetTableNames() As IReadOnlyList(Of String) Implements IAccessConnection.GetTableNames
            ThrowIfDisposed()
            EnsureConnectionOpen()
            Return GetTableNamesCore()
        End Function

        ''' <inheritdoc/>
        Public Async Function GetTableNamesAsync(Optional cancellationToken As CancellationToken = Nothing) As Task(Of IReadOnlyList(Of String)) Implements IAccessConnection.GetTableNamesAsync
            ThrowIfDisposed()
            Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
            Return GetTableNamesCore()
        End Function

        ''' <inheritdoc/>
        Public Function GetQueryNames() As IReadOnlyList(Of String) Implements IAccessConnection.GetQueryNames
            ThrowIfDisposed()
            EnsureConnectionOpen()
            Return GetQueryNamesCore()
        End Function

        ''' <inheritdoc/>
        Public Async Function GetQueryNamesAsync(Optional cancellationToken As CancellationToken = Nothing) As Task(Of IReadOnlyList(Of String)) Implements IAccessConnection.GetQueryNamesAsync
            ThrowIfDisposed()
            Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
            Return GetQueryNamesCore()
        End Function

        ''' <inheritdoc/>
        Public Function GetColumns(tableName As String) As DataTable Implements IAccessConnection.GetColumns
            ThrowIfDisposed()
            EnsureConnectionOpen()
            Return _connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, tableName, Nothing})
        End Function

        Private Function TableExistsCore(tableName As String) As Boolean
            For Each name In GetTableNamesCore()
                If String.Equals(name, tableName, StringComparison.OrdinalIgnoreCase) Then Return True
            Next
            Return False
        End Function

        Private Function GetTableNamesCore() As IReadOnlyList(Of String)
            Dim names As New List(Of String)()
            Using schema = _connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
                If schema IsNot Nothing Then
                    For Each row As DataRow In schema.Rows
                        names.Add(CStr(row("TABLE_NAME")))
                    Next
                End If
            End Using
            names.Sort(StringComparer.OrdinalIgnoreCase)
            Return names.AsReadOnly()
        End Function

        Private Function GetQueryNamesCore() As IReadOnlyList(Of String)
            Dim names As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' Les requêtes SELECT enregistrées apparaissent comme des vues.
            Using views = _connection.GetOleDbSchemaTable(OleDbSchemaGuid.Views, Nothing)
                If views IsNot Nothing Then
                    For Each row As DataRow In views.Rows
                        names.Add(CStr(row("TABLE_NAME")))
                    Next
                End If
            End Using

            ' Les requêtes d'action enregistrées apparaissent comme des procédures.
            Try
                Using procs = _connection.GetOleDbSchemaTable(OleDbSchemaGuid.Procedures, Nothing)
                    If procs IsNot Nothing Then
                        For Each row As DataRow In procs.Rows
                            names.Add(CStr(row("PROCEDURE_NAME")))
                        Next
                    End If
                End Using
            Catch
                ' Certains providers n'exposent pas les procédures : non bloquant.
            End Try

            Return names.ToList().AsReadOnly()
        End Function

#End Region

#Region "Insertion en masse"

        ''' <inheritdoc/>
        Public Function BulkInsert(tableName As String, columns As String(), rows As IEnumerable(Of Object())) As Integer Implements IAccessConnection.BulkInsert
            ThrowIfDisposed()
            ArgumentNullException.ThrowIfNull(columns)
            ArgumentNullException.ThrowIfNull(rows)
            Dim sql = BuildInsertSql(tableName, columns)

            Return ExecuteResilient(Function()
                                        EnsureConnectionOpen()
                                        Dim ownTransaction = Not _isTransactionActive
                                        If ownTransaction Then
                                            _transaction = _connection.BeginTransaction()
                                            _isTransactionActive = True
                                        End If

                                        Try
                                            Dim total = 0
                                            Using command = CreateCommand(sql)
                                                For Each row In rows
                                                    command.Parameters.Clear()
                                                    AddParameters(command, row)
                                                    total += command.ExecuteNonQuery()
                                                Next
                                            End Using
                                            If ownTransaction Then
                                                _transaction.Commit()
                                                DisposeTransaction()
                                            End If
                                            ResetAutoCloseTimer()
                                            Return total
                                        Catch
                                            If ownTransaction Then
                                                Try : _transaction.Rollback() : Catch : End Try
                                                DisposeTransaction()
                                            End If
                                            Throw
                                        End Try
                                    End Function)
        End Function

        ''' <inheritdoc/>
        Public Function BulkInsertAsync(tableName As String, columns As String(), rows As IEnumerable(Of Object()), Optional cancellationToken As CancellationToken = Nothing) As Task(Of Integer) Implements IAccessConnection.BulkInsertAsync
            ThrowIfDisposed()
            ArgumentNullException.ThrowIfNull(columns)
            ArgumentNullException.ThrowIfNull(rows)
            Dim sql = BuildInsertSql(tableName, columns)

            Return ExecuteResilientAsync(Async Function()
                                             Await EnsureConnectionOpenAsync(cancellationToken).ConfigureAwait(False)
                                             Dim ownTransaction = Not _isTransactionActive
                                             If ownTransaction Then
                                                 _transaction = CType(Await _connection.BeginTransactionAsync(IsolationLevel.ReadCommitted, cancellationToken).ConfigureAwait(False), OleDbTransaction)
                                                 _isTransactionActive = True
                                             End If

                                             Dim failure As Exception = Nothing
                                             Dim total = 0
                                             Try
                                                 Using command = CreateCommand(sql)
                                                     For Each row In rows
                                                         cancellationToken.ThrowIfCancellationRequested()
                                                         command.Parameters.Clear()
                                                         AddParameters(command, row)
                                                         total += Await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(False)
                                                     Next
                                                 End Using
                                             Catch ex As Exception
                                                 failure = ex
                                             End Try

                                             If failure IsNot Nothing Then
                                                 If ownTransaction Then
                                                     Try : Await _transaction.RollbackAsync(cancellationToken).ConfigureAwait(False) : Catch : End Try
                                                     DisposeTransaction()
                                                 End If
                                                 Throw failure
                                             End If

                                             If ownTransaction Then
                                                 Await _transaction.CommitAsync(cancellationToken).ConfigureAwait(False)
                                                 DisposeTransaction()
                                             End If
                                             ResetAutoCloseTimer()
                                             Return total
                                         End Function, cancellationToken)
        End Function

        Private Shared Function BuildInsertSql(tableName As String, columns As String()) As String
            If String.IsNullOrWhiteSpace(tableName) Then
                Throw New ArgumentException("Le nom de la table est requis.", NameOf(tableName))
            End If
            If columns Is Nothing OrElse columns.Length = 0 Then
                Throw New ArgumentException("Au moins une colonne est requise.", NameOf(columns))
            End If
            Dim cols = String.Join(", ", columns.Select(Function(c) $"[{c}]"))
            Dim placeholders = String.Join(", ", columns.Select(Function(c) "?"))
            Return $"INSERT INTO [{tableName}] ({cols}) VALUES ({placeholders})"
        End Function

#End Region

#Region "Création de commandes"

        ''' <inheritdoc/>
        Public Function CreateCommand(Optional commandText As String = Nothing) As OleDbCommand Implements IAccessConnection.CreateCommand
            ThrowIfDisposed()
            Dim command = _connection.CreateCommand()
            command.Transaction = _transaction
            If Not String.IsNullOrEmpty(commandText) Then
                command.CommandText = commandText
            End If
            Return command
        End Function

        Private Function CreateCommandInternal(sql As String, parameters As Object()) As OleDbCommand
            Dim command = CreateCommand(sql)
            AddParameters(command, parameters)
            Return command
        End Function

        Private Function CreateStoredQueryCommand(queryName As String, parameters As Object()) As OleDbCommand
            If String.IsNullOrWhiteSpace(queryName) Then
                Throw New ArgumentException("Le nom de la requête enregistrée est requis.", NameOf(queryName))
            End If
            Dim command = CreateCommand(queryName)
            command.CommandType = CommandType.StoredProcedure
            AddParameters(command, parameters)
            Return command
        End Function

        ''' <summary>
        ''' Ajoute des paramètres positionnels à la commande. OLE DB est positionnel :
        ''' l'ORDRE des paramètres compte, pas leur nom. Utilisez '?' dans le SQL.
        ''' Un élément déjà typé <see cref="OleDbParameter"/> est ajouté tel quel.
        ''' </summary>
        ''' <remarks>
        ''' Certains types .NET sont typés explicitement pour éviter les erreurs
        ''' "Data type mismatch" classiques d'Access :
        ''' <list type="bullet">
        ''' <item><see cref="DateTime"/> -> <see cref="OleDbType.Date"/> (sinon DBTimeStamp incompatible avec DATETIME).</item>
        ''' <item><see cref="Decimal"/> -> <see cref="OleDbType.Currency"/> (type monétaire standard d'Access).
        ''' Pour une colonne DECIMAL(p,s) haute précision, passez plutôt un
        ''' <see cref="OleDbParameter"/> explicite.</item>
        ''' </list>
        ''' </remarks>
        Private Shared Sub AddParameters(command As OleDbCommand, parameters As Object())
            If parameters Is Nothing Then Return
            For i = 0 To parameters.Length - 1
                Dim value = parameters(i)
                Dim typed = TryCast(value, OleDbParameter)
                If typed IsNot Nothing Then
                    command.Parameters.Add(typed)
                Else
                    Dim p = command.CreateParameter()
                    p.ParameterName = $"p{i}"
                    If value Is Nothing Then
                        p.Value = DBNull.Value
                    ElseIf TypeOf value Is Date Then
                        p.OleDbType = OleDbType.[Date]
                        p.Value = value
                    ElseIf TypeOf value Is Decimal Then
                        p.OleDbType = OleDbType.Currency
                        p.Value = value
                    Else
                        p.Value = value
                    End If
                    command.Parameters.Add(p)
                End If
            Next
        End Sub

#End Region

#Region "Maintenance (compactage)"

        ''' <inheritdoc/>
        Public Sub CompactDatabase(Optional targetPath As String = Nothing) Implements IAccessConnection.CompactDatabase
            ThrowIfDisposed()
            If _isTransactionActive Then
                Throw New InvalidOperationException("Impossible de compacter pendant une transaction active.")
            End If

            ' Le compactage exige une base fermée et sans verrou (.laccdb).
            Close()
            OleDbConnection.ReleaseObjectPool()

            Dim source = _options.ResolveDatabasePath()
            AccessMaintenance.Compact(source, _options.Password, targetPath)
            LogAction("CompactDatabase", $"Base compactée ({If(String.IsNullOrWhiteSpace(targetPath), "en place", targetPath)})")
        End Sub

        ''' <inheritdoc/>
        Public Function CompactDatabaseAsync(Optional targetPath As String = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task Implements IAccessConnection.CompactDatabaseAsync
            ' Vrai async : opération bloquante (COM/IO) déportée sur le pool de threads.
            Return Task.Run(Sub() CompactDatabase(targetPath), cancellationToken)
        End Function

#End Region

#Region "Création de base de données"

        ''' <summary>
        ''' Crée un nouveau fichier de base Access (.accdb) vide via ADOX (liaison tardive).
        ''' Le fichier ne doit pas déjà exister.
        ''' </summary>
        ''' <param name="databasePath">Chemin du fichier à créer.</param>
        ''' <param name="password">Mot de passe optionnel à appliquer à la base.</param>
        ''' <param name="provider">Provider OLE DB (par défaut Microsoft.ACE.OLEDB.16.0).</param>
        Public Shared Sub CreateDatabase(databasePath As String, Optional password As String = Nothing, Optional provider As String = AccessConnectionOptions.DefaultProvider)
            AccessDatabaseManager.Create(databasePath, password, provider)
        End Sub

        ''' <summary>Crée un nouveau fichier de base Access (.accdb) vide (vrai async, déporté sur un thread).</summary>
        Public Shared Function CreateDatabaseAsync(databasePath As String, Optional password As String = Nothing, Optional provider As String = AccessConnectionOptions.DefaultProvider, Optional cancellationToken As CancellationToken = Nothing) As Task
            Return Task.Run(Sub() AccessDatabaseManager.Create(databasePath, password, provider), cancellationToken)
        End Function

#End Region

#Region "Timer de fermeture automatique"

        ''' <inheritdoc/>
        Public Sub ResetAutoCloseTimer() Implements IAccessConnection.ResetAutoCloseTimer
            If _autoCloseTimer Is Nothing OrElse _isTransactionActive Then Return
            _autoCloseTimer.Change(_options.AutoCloseTimeoutMs, Timeout.Infinite)
            _logger?.LogDebug("Timer de fermeture automatique réinitialisé pour la connexion {ConnectionId}", Id)
        End Sub

        Private Sub OnAutoCloseTimerElapsed(state As Object)
            If _isTransactionActive OrElse _options.DisableAutoClose OrElse _options.IsPrimaryConnection Then Return
            Try
                If _connection IsNot Nothing AndAlso _connection.State = ConnectionState.Open Then
                    Close()
                    LogAction("AutoClose", "Connexion fermée automatiquement (timeout)")
                End If
            Catch ex As Exception
                _logger?.LogError(ex, "Erreur lors de la fermeture automatique de la connexion {ConnectionId}", Id)
            End Try
        End Sub

#End Region

#Region "Méthodes utilitaires privées"

        ' Utilisé À L'INTÉRIEUR des wrappers résilients : appelle l'ouverture brute
        ' (OpenInternal) pour éviter un double retry/traduction imbriqué.
        Private Sub EnsureConnectionOpen()
            If _connection?.State <> ConnectionState.Open Then
                OpenInternal()
            End If
        End Sub

        Private Async Function EnsureConnectionOpenAsync(cancellationToken As CancellationToken) As Task
            If _connection?.State <> ConnectionState.Open Then
                Await OpenInternalAsync(cancellationToken).ConfigureAwait(False)
            End If
        End Function

        ' --- Résilience (retry + traduction) ---------------------------------

        Private Function ExecuteResilient(Of T)(operation As Func(Of T)) As T
            Dim attempt = 0
            Do
                Try
                    Return operation()
                Catch ex As OleDbException
                    If ShouldRetry(ex, attempt) Then
                        attempt += 1
                        Thread.Sleep(ComputeBackoffMs(attempt))
                    ElseIf _options.TranslateErrors Then
                        Throw AccessErrorTranslator.Translate(ex)
                    Else
                        Throw
                    End If
                End Try
            Loop
        End Function

        Private Async Function ExecuteResilientAsync(Of T)(operation As Func(Of Task(Of T)), cancellationToken As CancellationToken) As Task(Of T)
            Dim attempt = 0
            Do
                cancellationToken.ThrowIfCancellationRequested()
                Dim toThrow As Exception = Nothing
                Dim delayMs = -1
                Try
                    Return Await operation().ConfigureAwait(False)
                Catch ex As OleDbException
                    If ShouldRetry(ex, attempt) Then
                        attempt += 1
                        delayMs = ComputeBackoffMs(attempt)
                    ElseIf _options.TranslateErrors Then
                        toThrow = AccessErrorTranslator.Translate(ex)
                    Else
                        toThrow = ex
                    End If
                End Try

                ' Lever / temporiser HORS du bloc Catch (Await interdit dans Catch en VB).
                If toThrow IsNot Nothing Then Throw toThrow
                If delayMs >= 0 Then Await Task.Delay(delayMs, cancellationToken).ConfigureAwait(False)
            Loop
        End Function

        Private Function ShouldRetry(ex As OleDbException, attempt As Integer) As Boolean
            If Not _options.EnableRetryOnTransientErrors Then Return False
            If _isTransactionActive Then Return False   ' jamais de retry au sein d'une transaction
            If attempt >= _options.MaxRetries Then Return False
            Return AccessErrorTranslator.IsTransient(ex)
        End Function

        Private Function ComputeBackoffMs(attempt As Integer) As Integer
            ' attempt >= 1 ; back-off exponentiel : base * 2^(attempt-1)
            Dim factor = CInt(Math.Pow(2, Math.Max(0, attempt - 1)))
            Return _options.RetryBaseDelayMs * factor
        End Function

        ' --- Mapping objet (micro-ORM par réflexion) -------------------------

        Private Shared ReadOnly _propertyCache As New ConcurrentDictionary(Of Type, PropertyInfo())()

        Private Shared Function MapReader(Of T As New)(reader As IDataReader) As List(Of T)
            Dim byName As New Dictionary(Of String, PropertyInfo)(StringComparer.OrdinalIgnoreCase)
            For Each p In GetWritableProperties(GetType(T))
                byName(p.Name) = p
            Next

            Dim mapping As New List(Of KeyValuePair(Of Integer, PropertyInfo))()
            For i = 0 To reader.FieldCount - 1
                Dim prop As PropertyInfo = Nothing
                If byName.TryGetValue(reader.GetName(i), prop) Then
                    mapping.Add(New KeyValuePair(Of Integer, PropertyInfo)(i, prop))
                End If
            Next

            Dim result As New List(Of T)()
            While reader.Read()
                Dim item As New T()
                For Each m In mapping
                    Dim value = reader.GetValue(m.Key)
                    If value Is Nothing OrElse Convert.IsDBNull(value) Then Continue For
                    Dim targetType = If(Nullable.GetUnderlyingType(m.Value.PropertyType), m.Value.PropertyType)
                    If targetType.IsInstanceOfType(value) Then
                        m.Value.SetValue(item, value)
                    Else
                        m.Value.SetValue(item, Convert.ChangeType(value, targetType))
                    End If
                Next
                result.Add(item)
            End While
            Return result
        End Function

        Private Shared Function GetWritableProperties(t As Type) As PropertyInfo()
            Return _propertyCache.GetOrAdd(t, Function(tt) tt.GetProperties(BindingFlags.Public Or BindingFlags.Instance).Where(Function(p) p.CanWrite).ToArray())
        End Function

        Private Shared Function ConvertScalar(Of T)(result As Object) As T
            If result Is Nothing OrElse Convert.IsDBNull(result) Then
                Return Nothing
            End If

            Dim targetType = GetType(T)
            Dim underlying = Nullable.GetUnderlyingType(targetType)
            If underlying IsNot Nothing Then targetType = underlying

            If targetType.IsInstanceOfType(result) Then
                Return CType(result, T)
            End If

            Return CType(Convert.ChangeType(result, targetType), T)
        End Function

        Private Sub OnConnectionStateChanged(sender As Object, e As StateChangeEventArgs)
            If e.CurrentState = ConnectionState.Open Then
                ResetAutoCloseTimer()
            End If
            _logger?.LogDebug("État de la connexion {ConnectionId} changé : {OldState} -> {NewState}", Id, e.OriginalState, e.CurrentState)
        End Sub

        Private Sub LogAction(action As String, description As String, <CallerMemberName> Optional callerName As String = Nothing)
            SyncLock _actionLock
                Dim logEntry = $"[{DateTime.UtcNow:HH:mm:ss.fff}] [{action}] {description}"

                If _actionHistory.Count >= MaxActionHistorySize Then
                    _actionHistory.RemoveAt(_actionHistory.Count - 1)
                End If

                If Not String.IsNullOrEmpty(_lastAction) Then
                    _actionHistory.Insert(0, _lastAction)
                End If

                _lastAction = logEntry
            End SyncLock

            _logger?.LogDebug("Connexion {ConnectionId} - {Action}: {Description}", Id, action, description)
        End Sub

        Private Sub ThrowIfDisposed()
            ObjectDisposedException.ThrowIf(_isDisposed, Me)
        End Sub

#End Region

#Region "IDisposable / IAsyncDisposable"

        ''' <inheritdoc/>
        Public Sub Dispose() Implements IDisposable.Dispose
            DisposeCore(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>
        ''' VB.NET n'autorise pas une fonction Async à retourner ValueTask : on enveloppe
        ''' donc une méthode interne Async retournant Task dans un ValueTask.
        ''' </remarks>
        Public Function DisposeAsync() As ValueTask Implements IAsyncDisposable.DisposeAsync
            Return New ValueTask(DisposeAsyncInternal())
        End Function

        Private Async Function DisposeAsyncInternal() As Task
            Await DisposeAsyncCore().ConfigureAwait(False)
            DisposeCore(disposing:=False)
            GC.SuppressFinalize(Me)
        End Function

        Private Sub DisposeCore(disposing As Boolean)
            SyncLock _disposeLock
                If _isDisposed Then Return

                If disposing Then
                    _autoCloseTimer?.Dispose()
                    DisposeTransaction()

                    If _connection IsNot Nothing Then
                        If _connection.State <> ConnectionState.Closed Then
                            Close()
                        End If
                        RemoveHandler _connection.StateChange, AddressOf OnConnectionStateChanged
                        _connection.Dispose()
                        _connection = Nothing
                    End If

                    ' Libère le pool OLE DB pour relâcher le verrou .laccdb si connexion principale.
                    If _options.IsPrimaryConnection Then
                        OleDbConnection.ReleaseObjectPool()
                    End If

                    LogAction("Dispose", "Ressources libérées")
                End If

                _isDisposed = True
            End SyncLock
        End Sub

        Private Async Function DisposeAsyncCore() As Task
            If _autoCloseTimer IsNot Nothing Then
                Await _autoCloseTimer.DisposeAsync().ConfigureAwait(False)
            End If

            If _transaction IsNot Nothing Then
                Await _transaction.DisposeAsync().ConfigureAwait(False)
                _transaction = Nothing
                _isTransactionActive = False
            End If

            If _connection IsNot Nothing Then
                If _connection.State <> ConnectionState.Closed Then
                    Await CloseAsync().ConfigureAwait(False)
                End If
                RemoveHandler _connection.StateChange, AddressOf OnConnectionStateChanged
                Await _connection.DisposeAsync().ConfigureAwait(False)
                _connection = Nothing
            End If

            If _options.IsPrimaryConnection Then
                OleDbConnection.ReleaseObjectPool()
            End If

            LogAction("DisposeAsync", "Ressources libérées (async)")
        End Function

#End Region

    End Class

End Namespace
