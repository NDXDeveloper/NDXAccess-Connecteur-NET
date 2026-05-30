Imports System.Data
Imports System.Data.OleDb
Imports System.Threading

Namespace NDXAccess

    ''' <summary>
    ''' Interface de gestion d'une connexion Microsoft Access via OLE DB (ACE).
    ''' </summary>
    ''' <remarks>
    ''' <para>
    ''' Les méthodes asynchrones sont fournies pour la cohérence d'API et l'intégration
    ''' avec du code async/await. ATTENTION : le provider OLE DB n'expose pas de vraies
    ''' opérations asynchrones — il s'agit d'un async "de façade" qui s'exécute de manière
    ''' synchrone sur le thread courant (sauf <see cref="CompactDatabaseAsync"/> qui est
    ''' réellement déporté sur un thread). Ne comptez pas dessus pour du parallélisme.
    ''' </para>
    ''' <para>
    ''' THREAD-SAFETY : une instance n'est PAS thread-safe. N'utilisez pas la même
    ''' connexion depuis plusieurs threads simultanément ; créez une connexion par
    ''' thread/opération (via <see cref="IAccessConnectionFactory"/>).
    ''' </para>
    ''' </remarks>
    Public Interface IAccessConnection
        Inherits IDisposable
        Inherits IAsyncDisposable

        ''' <summary>Identifiant unique de la connexion.</summary>
        ReadOnly Property Id As Integer

        ''' <summary>Nom du provider OLE DB utilisé (ex. "Microsoft.ACE.OLEDB.16.0").</summary>
        ReadOnly Property ProviderName As String

        ''' <summary>
        ''' Version du moteur ACE (DLL, ex. "16.0.5011.1000", ou la génération "16.0").
        ''' À ne pas confondre avec la version du format de fichier (DatabaseInfo.FileFormatVersion).
        ''' </summary>
        ReadOnly Property EngineVersion As String

        ''' <summary>Date de création de la connexion (UTC).</summary>
        ReadOnly Property CreatedAt As DateTime

        ''' <summary>État actuel de la connexion.</summary>
        ReadOnly Property State As ConnectionState

        ''' <summary>Indique si une transaction est en cours.</summary>
        ReadOnly Property IsTransactionActive As Boolean

        ''' <summary>Indique si c'est la connexion principale.</summary>
        ReadOnly Property IsPrimaryConnection As Boolean

        ''' <summary>Connexion OLE DB sous-jacente (Nothing si non initialisée).</summary>
        ReadOnly Property Connection As OleDbConnection

        ''' <summary>Transaction en cours (Nothing si aucune).</summary>
        ReadOnly Property Transaction As OleDbTransaction

        ''' <summary>Dernière action effectuée sur la connexion.</summary>
        ReadOnly Property LastAction As String

        ''' <summary>Historique des 5 dernières actions.</summary>
        ReadOnly Property ActionHistory As IReadOnlyList(Of String)

        ' ---- Cycle de vie -------------------------------------------------------

        ''' <summary>Ouvre la connexion (synchrone).</summary>
        Sub Open()

        ''' <summary>Ouvre la connexion (async de façade).</summary>
        Function OpenAsync(Optional cancellationToken As CancellationToken = Nothing) As Task

        ''' <summary>Ferme la connexion (synchrone).</summary>
        Sub Close()

        ''' <summary>Ferme la connexion (async de façade).</summary>
        Function CloseAsync() As Task

        ''' <summary>Réinitialise le timer de fermeture automatique.</summary>
        Sub ResetAutoCloseTimer()

        ' ---- Transactions -------------------------------------------------------

        ''' <summary>Démarre une transaction (synchrone).</summary>
        Function BeginTransaction(Optional isolationLevel As IsolationLevel = IsolationLevel.ReadCommitted) As Boolean

        ''' <summary>Démarre une transaction (async de façade).</summary>
        Function BeginTransactionAsync(Optional isolationLevel As IsolationLevel = IsolationLevel.ReadCommitted, Optional cancellationToken As CancellationToken = Nothing) As Task(Of Boolean)

        ''' <summary>Valide la transaction en cours (synchrone).</summary>
        Sub Commit()

        ''' <summary>Valide la transaction en cours (async de façade).</summary>
        Function CommitAsync(Optional cancellationToken As CancellationToken = Nothing) As Task

        ''' <summary>Annule la transaction en cours (synchrone).</summary>
        Sub Rollback()

        ''' <summary>Annule la transaction en cours (async de façade).</summary>
        Function RollbackAsync(Optional cancellationToken As CancellationToken = Nothing) As Task

        ' ---- Exécution de requêtes (paramètres positionnels '?') ----------------

        ''' <summary>
        ''' Exécute une requête et retourne le nombre de lignes affectées (synchrone).
        ''' Les paramètres sont positionnels : utilisez '?' dans le SQL, dans l'ordre.
        ''' </summary>
        Function ExecuteNonQuery(sql As String, ParamArray parameters As Object()) As Integer

        ''' <summary>Exécute une requête et retourne le nombre de lignes affectées (async de façade).</summary>
        Function ExecuteNonQueryAsync(sql As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of Integer)

        ''' <summary>Exécute une requête et retourne une valeur scalaire typée (synchrone).</summary>
        Function ExecuteScalar(Of T)(sql As String, ParamArray parameters As Object()) As T

        ''' <summary>Exécute une requête et retourne une valeur scalaire typée (async de façade).</summary>
        Function ExecuteScalarAsync(Of T)(sql As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of T)

        ''' <summary>Exécute une requête et retourne un DataTable (synchrone).</summary>
        Function ExecuteQuery(sql As String, ParamArray parameters As Object()) As DataTable

        ''' <summary>Exécute une requête et retourne un DataTable (async de façade).</summary>
        Function ExecuteQueryAsync(sql As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of DataTable)

        ''' <summary>Exécute une requête et retourne un OleDbDataReader (synchrone). Le reader doit être disposé par l'appelant.</summary>
        Function ExecuteReader(sql As String, ParamArray parameters As Object()) As OleDbDataReader

        ''' <summary>Exécute une requête et retourne un OleDbDataReader (async de façade).</summary>
        Function ExecuteReaderAsync(sql As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of OleDbDataReader)

        ' ---- Mapping objet (micro-ORM par réflexion) ---------------------------

        ''' <summary>
        ''' Exécute une requête et mappe chaque ligne vers un objet <typeparamref name="T"/>
        ''' (correspondance colonne -> propriété par nom, insensible à la casse). Synchrone.
        ''' </summary>
        Function ExecuteQuery(Of T As New)(sql As String, ParamArray parameters As Object()) As List(Of T)

        ''' <summary>Exécute une requête et mappe les lignes vers des objets typés (async de façade).</summary>
        Function ExecuteQueryAsync(Of T As New)(sql As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of List(Of T))

        ' ---- Paramètres nommés (@nom -> ?) -------------------------------------

        ''' <summary>Comme ExecuteNonQuery mais avec des paramètres nommés (@nom) traduits en positionnels.</summary>
        Function ExecuteNonQueryNamed(sql As String, parameters As IDictionary(Of String, Object)) As Integer

        ''' <summary>Version async de façade de ExecuteNonQueryNamed.</summary>
        Function ExecuteNonQueryNamedAsync(sql As String, parameters As IDictionary(Of String, Object), Optional cancellationToken As CancellationToken = Nothing) As Task(Of Integer)

        ''' <summary>Comme ExecuteScalar mais avec des paramètres nommés (@nom).</summary>
        Function ExecuteScalarNamed(Of T)(sql As String, parameters As IDictionary(Of String, Object)) As T

        ''' <summary>Version async de façade de ExecuteScalarNamed.</summary>
        Function ExecuteScalarNamedAsync(Of T)(sql As String, parameters As IDictionary(Of String, Object), Optional cancellationToken As CancellationToken = Nothing) As Task(Of T)

        ''' <summary>Comme ExecuteQuery (DataTable) mais avec des paramètres nommés (@nom).</summary>
        Function ExecuteQueryNamed(sql As String, parameters As IDictionary(Of String, Object)) As DataTable

        ''' <summary>Version async de façade de ExecuteQueryNamed.</summary>
        Function ExecuteQueryNamedAsync(sql As String, parameters As IDictionary(Of String, Object), Optional cancellationToken As CancellationToken = Nothing) As Task(Of DataTable)

        ' ---- Requêtes enregistrées (stored queries) ----------------------------

        ''' <summary>
        ''' Exécute une requête enregistrée (QueryDef) avec paramètres IN positionnels et
        ''' retourne un DataTable (synchrone).
        ''' </summary>
        ''' <remarks>
        ''' Access ne possède pas de vraies procédures stockées : seules les requêtes
        ''' enregistrées paramétrées existent, et uniquement avec des paramètres d'ENTRÉE.
        ''' Aucun paramètre OUT/INOUT n'est possible.
        ''' </remarks>
        Function ExecuteStoredQuery(queryName As String, ParamArray parameters As Object()) As DataTable

        ''' <summary>Exécute une requête enregistrée et retourne un DataTable (async de façade).</summary>
        Function ExecuteStoredQueryAsync(queryName As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of DataTable)

        ''' <summary>Exécute une requête enregistrée d'action (INSERT/UPDATE/DELETE) et retourne les lignes affectées (synchrone).</summary>
        Function ExecuteStoredQueryNonQuery(queryName As String, ParamArray parameters As Object()) As Integer

        ''' <summary>Exécute une requête enregistrée d'action et retourne les lignes affectées (async de façade).</summary>
        Function ExecuteStoredQueryNonQueryAsync(queryName As String, Optional parameters As Object() = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task(Of Integer)

        ''' <summary>Crée une nouvelle commande liée à cette connexion (et à la transaction en cours).</summary>
        Function CreateCommand(Optional commandText As String = Nothing) As OleDbCommand

        ' ---- Helpers de schéma --------------------------------------------------

        ''' <summary>Indique si une table utilisateur existe (insensible à la casse).</summary>
        Function TableExists(tableName As String) As Boolean

        ''' <summary>Version async de façade de TableExists.</summary>
        Function TableExistsAsync(tableName As String, Optional cancellationToken As CancellationToken = Nothing) As Task(Of Boolean)

        ''' <summary>Retourne la liste des tables utilisateur (triée).</summary>
        Function GetTableNames() As IReadOnlyList(Of String)

        ''' <summary>Version async de façade de GetTableNames.</summary>
        Function GetTableNamesAsync(Optional cancellationToken As CancellationToken = Nothing) As Task(Of IReadOnlyList(Of String))

        ''' <summary>Retourne la liste des requêtes enregistrées (vues + procédures).</summary>
        Function GetQueryNames() As IReadOnlyList(Of String)

        ''' <summary>Version async de façade de GetQueryNames.</summary>
        Function GetQueryNamesAsync(Optional cancellationToken As CancellationToken = Nothing) As Task(Of IReadOnlyList(Of String))

        ''' <summary>Retourne le schéma des colonnes d'une table (DataTable OLE DB).</summary>
        Function GetColumns(tableName As String) As DataTable

        ' ---- Insertion en masse -------------------------------------------------

        ''' <summary>
        ''' Insère plusieurs lignes dans une transaction unique (bien plus rapide qu'en
        ''' auto-commit). Chaque ligne fournit les valeurs dans l'ordre des colonnes.
        ''' Retourne le nombre total de lignes insérées. Synchrone.
        ''' </summary>
        Function BulkInsert(tableName As String, columns As String(), rows As IEnumerable(Of Object())) As Integer

        ''' <summary>Version async de façade de BulkInsert.</summary>
        Function BulkInsertAsync(tableName As String, columns As String(), rows As IEnumerable(Of Object()), Optional cancellationToken As CancellationToken = Nothing) As Task(Of Integer)

        ' ---- Maintenance --------------------------------------------------------

        ''' <summary>
        ''' Compacte et répare la base de données (synchrone). La connexion est fermée
        ''' au préalable. Si <paramref name="targetPath"/> est Nothing, compactage en place.
        ''' </summary>
        Sub CompactDatabase(Optional targetPath As String = Nothing)

        ''' <summary>
        ''' Compacte et répare la base de données. C'est la SEULE opération réellement
        ''' asynchrone (déportée sur un thread du pool car bloquante et coûteuse).
        ''' </summary>
        Function CompactDatabaseAsync(Optional targetPath As String = Nothing, Optional cancellationToken As CancellationToken = Nothing) As Task

    End Interface

End Namespace
