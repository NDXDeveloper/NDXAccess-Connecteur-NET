Imports System.Data.OleDb

Namespace NDXAccess

    ''' <summary>
    ''' Options de configuration pour une connexion Microsoft Access (.accdb / .mdb)
    ''' via le provider OLE DB Microsoft ACE.
    ''' </summary>
    ''' <remarks>
    ''' <para>Windows uniquement : le provider ACE n'existe pas sur Linux/macOS.</para>
    ''' <para>Auteur : Nicolas DEOUX - NDXDev 2026</para>
    ''' </remarks>
    Public NotInheritable Class AccessConnectionOptions

        ''' <summary>
        ''' Nom du provider OLE DB ACE par défaut (Access Database Engine 2016+).
        ''' </summary>
        Public Const DefaultProvider As String = "Microsoft.ACE.OLEDB.16.0"

        ''' <summary>
        ''' Taille maximale d'un fichier Access en octets (limite stricte du moteur : 2 Go).
        ''' </summary>
        Public Const MaxDatabaseSizeBytes As Long = 2_147_483_648L

        ''' <summary>
        ''' Chemin complet du fichier de base de données (.accdb ou .mdb).
        ''' </summary>
        Public Property DatabasePath As String = String.Empty

        ''' <summary>
        ''' Mot de passe de la base (encapsulé dans "Jet OLEDB:Database Password").
        ''' Laisser vide si la base n'est pas protégée.
        ''' </summary>
        Public Property Password As String = String.Empty

        ''' <summary>
        ''' Nom du provider OLE DB à utiliser. Par défaut <see cref="DefaultProvider"/>.
        ''' Utilisez "Microsoft.ACE.OLEDB.12.0" pour un moteur plus ancien.
        ''' </summary>
        Public Property Provider As String = DefaultProvider

        ''' <summary>
        ''' Chaîne de connexion complète. Si définie, elle surcharge toutes les autres propriétés.
        ''' </summary>
        Public Property ConnectionString As String = Nothing

        ''' <summary>
        ''' Ouvre la base en mode exclusif (Mode=Share Exclusive).
        ''' Nécessaire pour certaines opérations de maintenance.
        ''' </summary>
        Public Property OpenExclusive As Boolean = False

        ''' <summary>
        ''' Ouvre la base en lecture seule (Mode=Read).
        ''' </summary>
        Public Property OpenReadOnly As Boolean = False

        ''' <summary>
        ''' Conserve les informations de sécurité (mot de passe) dans la chaîne de connexion
        ''' après ouverture. Déconseillé. Par défaut False.
        ''' </summary>
        Public Property PersistSecurityInfo As Boolean = False

        ''' <summary>
        ''' Chemin d'un fichier de groupe de travail (.mdw) pour la sécurité utilisateur (Jet legacy).
        ''' Optionnel.
        ''' </summary>
        Public Property SystemDatabasePath As String = String.Empty

        ''' <summary>
        ''' Indique s'il s'agit de la connexion principale (jamais fermée automatiquement).
        ''' </summary>
        Public Property IsPrimaryConnection As Boolean = False

        ''' <summary>
        ''' Délai d'inactivité avant fermeture automatique, en millisecondes (60000 = 1 minute).
        ''' </summary>
        Public Property AutoCloseTimeoutMs As Integer = 60_000

        ''' <summary>
        ''' Désactive la fermeture automatique de la connexion inactive.
        ''' </summary>
        Public Property DisableAutoClose As Boolean = False

        ''' <summary>
        ''' Vérifie la disponibilité du provider (et donc la cohérence x86/x64) à la création
        ''' de la connexion. Lève <see cref="AccessProviderNotFoundException"/> en cas de problème.
        ''' Par défaut True.
        ''' </summary>
        Public Property ValidateProvider As Boolean = True

        ''' <summary>
        ''' Réessaie automatiquement les opérations qui échouent sur une erreur de verrou
        ''' transitoire (fichier/enregistrement verrouillé par un autre utilisateur).
        ''' Les opérations au sein d'une transaction ne sont jamais réessayées. Par défaut True.
        ''' </summary>
        Public Property EnableRetryOnTransientErrors As Boolean = True

        ''' <summary>Nombre maximal de nouvelles tentatives sur erreur transitoire (par défaut 3).</summary>
        Public Property MaxRetries As Integer = 3

        ''' <summary>Délai de base (ms) du back-off exponentiel entre tentatives (par défaut 100 ms).</summary>
        Public Property RetryBaseDelayMs As Integer = 100

        ''' <summary>
        ''' Traduit les <see cref="System.Data.OleDb.OleDbException"/> en
        ''' <see cref="AccessQueryException"/> avec un message clair et le code natif.
        ''' Par défaut True. Mettez à False pour propager l'exception OLE DB brute.
        ''' </summary>
        Public Property TranslateErrors As Boolean = True

        ''' <summary>
        ''' Génère la chaîne de connexion OLE DB à partir des options.
        ''' </summary>
        Public Function BuildConnectionString() As String
            If Not String.IsNullOrWhiteSpace(ConnectionString) Then
                Return ConnectionString
            End If

            Dim builder As New OleDbConnectionStringBuilder() With {
                .Provider = Provider
            }
            builder("Data Source") = DatabasePath
            builder("Persist Security Info") = PersistSecurityInfo

            If Not String.IsNullOrEmpty(Password) Then
                builder("Jet OLEDB:Database Password") = Password
            End If

            If OpenExclusive Then
                builder("Mode") = "Share Exclusive"
            ElseIf OpenReadOnly Then
                builder("Mode") = "Read"
            End If

            If Not String.IsNullOrWhiteSpace(SystemDatabasePath) Then
                builder("Jet OLEDB:System Database") = SystemDatabasePath
            End If

            Return builder.ConnectionString
        End Function

        ''' <summary>
        ''' Extrait le nom du provider effectif (en tenant compte d'une éventuelle
        ''' chaîne de connexion personnalisée).
        ''' </summary>
        Public Function ResolveProviderName() As String
            If String.IsNullOrWhiteSpace(ConnectionString) Then
                Return Provider
            End If

            Try
                Return New OleDbConnectionStringBuilder(ConnectionString).Provider
            Catch
                Return Provider
            End Try
        End Function

        ''' <summary>
        ''' Extrait le chemin du fichier de base effectif (en tenant compte d'une
        ''' éventuelle chaîne de connexion personnalisée via "Data Source").
        ''' </summary>
        Public Function ResolveDatabasePath() As String
            If Not String.IsNullOrWhiteSpace(DatabasePath) Then
                Return DatabasePath
            End If

            If Not String.IsNullOrWhiteSpace(ConnectionString) Then
                Try
                    Dim builder As New OleDbConnectionStringBuilder(ConnectionString)
                    If builder.ContainsKey("Data Source") Then
                        Return TryCast(builder("Data Source"), String)
                    End If
                Catch
                End Try
            End If

            Return DatabasePath
        End Function

    End Class

End Namespace
