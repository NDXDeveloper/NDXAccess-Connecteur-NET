# Documentation du projet source NDXAccess

## Vue d'ensemble

NDXAccess est un connecteur Microsoft Access (.accdb / .mdb) moderne pour VB.NET,  
basé sur `System.Data.OleDb` et le provider `Microsoft.ACE.OLEDB.16.0`. Il propose une  
API fluide (connexion, transactions, CRUD, requêtes enregistrées, health checks,  
compactage) en tenant compte des spécificités du moteur Access.  

## Architecture

```
src/NDXAccess/
├── IAccessConnection.vb           # Interface de connexion (sync + async)
├── AccessConnection.vb            # Connexion principale (cœur)
├── AccessConnectionOptions.vb     # Options de configuration
├── AccessConnectionFactory.vb     # Factory + IAccessConnectionFactory
├── AccessHealthCheck.vb           # Health check + DatabaseInfo + HealthCheckResult
├── AccessProviderHelper.vb        # Détection x86/x64 du provider ACE
├── AccessMaintenance.vb           # Compactage via DAO (late binding)
├── AccessExceptions.vb            # Exceptions dédiées
└── Extensions/
    └── ServiceCollectionExtensions.vb  # Intégration DI
```

## Classes principales

### AccessConnection

Gère le cycle de vie d'une connexion OLE DB vers une base Access.

**Fonctionnalités :**
- Ouverture/fermeture synchrone et asynchrone (async de façade)
- Transactions (`OleDbTransaction`) Begin/Commit/Rollback
- Fermeture automatique des connexions inactives (timer)
- Historique des actions pour le débogage
- `IDisposable` et `IAsyncDisposable`
- Requêtes paramétrées positionnelles (`?`)
- Requêtes enregistrées (paramètres IN)
- Compactage/réparation
- Validation x86/x64 du provider à la création

**Exemple :**

```vb
Dim options As New AccessConnectionOptions With {
    .DatabasePath = "C:\data\ma_base.accdb",
    .Password = "secret"   ' optionnel
}

Using connection As IAccessConnection = New AccessConnection(options)
    Await connection.OpenAsync()

    Dim count = Await connection.ExecuteScalarAsync(Of Integer)("SELECT COUNT(*) FROM clients")

    Await connection.BeginTransactionAsync()
    Try
        Await connection.ExecuteNonQueryAsync(
            "UPDATE clients SET actif = ? WHERE id = ?", {True, 1})
        Await connection.CommitAsync()
    Catch
        Await connection.RollbackAsync()
        Throw
    End Try
End Using
```

### AccessConnectionOptions

| Propriété | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `DatabasePath` | String | "" | Chemin du fichier .accdb/.mdb |
| `Password` | String | "" | Mot de passe (Jet OLEDB:Database Password) |
| `Provider` | String | "Microsoft.ACE.OLEDB.16.0" | Provider OLE DB |
| `ConnectionString` | String | Nothing | Chaîne complète (surcharge les autres props) |
| `OpenExclusive` | Boolean | False | Mode=Share Exclusive |
| `OpenReadOnly` | Boolean | False | Mode=Read |
| `PersistSecurityInfo` | Boolean | False | Conserve les infos de sécurité |
| `SystemDatabasePath` | String | "" | Fichier de groupe de travail .mdw |
| `IsPrimaryConnection` | Boolean | False | Connexion principale (pas de fermeture auto) |
| `AutoCloseTimeoutMs` | Integer | 60000 | Fermeture auto (ms) |
| `DisableAutoClose` | Boolean | False | Désactive la fermeture auto |
| `ValidateProvider` | Boolean | True | Vérifie le provider (x86/x64) à la création |
| `EnableRetryOnTransientErrors` | Boolean | True | Retry sur verrous transitoires |
| `MaxRetries` | Integer | 3 | Nombre max de tentatives |
| `RetryBaseDelayMs` | Integer | 100 | Délai de base du back-off exponentiel |
| `TranslateErrors` | Boolean | True | OleDbException → AccessQueryException |

Constante : `MaxDatabaseSizeBytes = 2 147 483 648` (2 Go).

### AccessConnectionFactory

```vb
Dim factory As IAccessConnectionFactory = New AccessConnectionFactory(options, loggerFactory)

Using conn = factory.CreateConnection()              ' standard  
Using main = factory.CreatePrimaryConnection()       ' principale (jamais auto-fermée)  
Using cfg = factory.CreateConnection(Sub(o) o.OpenReadOnly = True)  
```

### AccessHealthCheck

```vb
Dim hc = New AccessHealthCheck(factory)

Dim result = Await hc.CheckHealthAsync()  
Console.WriteLine(result.IsHealthy & " - " & result.Message)  

Dim info = Await hc.GetDatabaseInfoAsync()  
Console.WriteLine($"{info.FileSizeMegabytes:F1} Mo / 2 Go ({info.UsagePercent:F1} %)")  
```

`DatabaseInfo` expose : `Provider`, `EngineVersion` (moteur ACE), `FileFormatVersion`
(format Jet/ACE), `FilePath`, `FileSizeBytes`, `FileSizeMegabytes`, `MaxSizeBytes`,
`UsagePercent`, `IsApproachingSizeLimit`, `IsReadOnly`, `UserTableCount`.

### AccessProviderHelper

Détection de l'architecture et de la disponibilité du provider (voir
[guide d'installation](../installation/README.md)).

## Paramètres : positionnels (`?`)

Contrairement aux SGBD à paramètres nommés (`@nom`), OLE DB/Access utilise des paramètres **positionnels**.  
Le SQL contient des `?` et les valeurs sont passées **dans l'ordre** :  

```vb
connection.ExecuteNonQuery(
    "INSERT INTO clients (nom, email, actif) VALUES (?, ?, ?)",
    "Jean", "jean@example.com", True)
```

Typage automatique : `Date` → `OleDbType.Date`, `Decimal` → `OleDbType.Currency`.  
Pour un contrôle fin, passez directement un `OleDbParameter`.  

## Requêtes enregistrées (et non procédures stockées)

Access n'a pas de procédures stockées. Il a des **requêtes enregistrées** acceptant des  
paramètres d'**entrée uniquement** (aucun OUT/INOUT) :  

```vb
' Création (une fois)
connection.ExecuteNonQuery(
    "CREATE PROCEDURE qParStatut (prmActif BIT) AS " &
    "SELECT * FROM clients WHERE actif = prmActif")

' Appel
Dim dt = connection.ExecuteStoredQuery("qParStatut", True)
```

## Async « de façade »

OLE DB n'expose pas de vraies opérations asynchrones. Les méthodes `...Async`  
retournent une `Task` mais s'exécutent **synchronement** sur le thread courant. Elles  
existent pour la cohérence d'API et l'usage de `Await`. **Seule `CompactDatabaseAsync`**  
est réellement asynchrone (opération bloquante déportée sur le pool de threads).  

## Fonctionnalités avancées

### Résilience (retry) et traduction d'erreurs

- **Retry** : les opérations qui échouent sur une erreur de **verrou transitoire** (3050,
  3260, 3218, etc., ou détectées par message) sont automatiquement réessayées avec un
  back-off exponentiel (`MaxRetries`, `RetryBaseDelayMs`). **Jamais** au sein d'une
  transaction (rejouer un statement isolé serait incorrect).
- **Traduction** : les `OleDbException` (souvent cryptiques, code natif fréquemment 0 avec
  ACE) sont converties en `AccessQueryException` avec un message clair, le code natif quand
  disponible, le flag `IsTransient`, et l'`InnerException` d'origine préservée. La détection
  combine code natif **et motifs de message** (doublon, mot de passe, verrou, « too few
  parameters », format non reconnu…). Désactivable via `TranslateErrors = False`.

### Helpers de schéma

```vb
connection.TableExists("clients")          ' insensible à la casse  
connection.GetTableNames()                 ' tables utilisateur  
connection.GetQueryNames()                 ' requêtes enregistrées (vues + procédures)  
connection.GetColumns("clients")           ' DataTable des colonnes  
```

### Création de base — `CreateDatabase`

```vb
AccessConnection.CreateDatabase("C:\data\new.accdb", password:="secret")  
Await AccessConnection.CreateDatabaseAsync("C:\data\new.accdb")  
```

### Insertion en masse — `BulkInsert`

Toutes les lignes sont insérées dans **une seule transaction** (gain majeur en Access) :

```vb
Dim n = Await connection.BulkInsertAsync("clients", {"nom", "email"}, rows)
```

### Mapping objet — `ExecuteQuery(Of T)`

Correspondance colonne → propriété par **nom** (insensible à la casse), conversions de  
type et `DBNull` gérés. `T` nécessite un constructeur sans paramètre.  

```vb
Dim list = Await connection.ExecuteQueryAsync(Of Client)("SELECT id, nom FROM clients")
```

### Paramètres nommés — `...Named`

OLE DB est positionnel, mais les méthodes `ExecuteNonQueryNamed` / `ExecuteScalarNamed` /
`ExecuteQueryNamed` (et async) acceptent `@nom` et les traduisent en `?` (littéraux chaîne
et `@@IDENTITY` préservés).

### Versions — bibliothèque, moteur ACE, format de fichier

Trois versions distinctes sont exposées :

```vb
' 1) Version de la BIBLIOTHÈQUE NDXAccess (métadonnées d'assembly)
NDXAccessInfo.Version                ' "1.2.0.0"  
NDXAccessInfo.InformationalVersion   ' "1.2.0"  
NDXAccessInfo.ProductName            ' "NDXAccess"  
AccessConnection.Version             ' raccourci vers InformationalVersion  

' 2) Version du MOTEUR ACE (DLL ACEOLEDB, via le registre ; repli "16.0")
connection.ProviderName                       ' "Microsoft.ACE.OLEDB.16.0"  
connection.EngineVersion                      ' "16.0.5011.1000"  
AccessProviderHelper.GetEngineVersion()       ' "16.0.5011.1000"  

' 3) Version du FORMAT de fichier (Jet/ACE)
healthCheck.GetDatabaseInfo().FileFormatVersion   ' "04.00.0000"
```

## Ce qui n'existe PAS dans Access (vs un SGBD client-serveur)

| Fonctionnalité | Statut Access |
|----------------|---------------|
| Event Scheduler | ❌ → Planificateur de tâches Windows |
| Procédures stockées IN/OUT/INOUT | ⚠️ Requêtes enregistrées, IN uniquement |
| Vrai async | ❌ Async de façade (sauf compactage) |
| Pooling serveur | ⚠️ Pooling OLE DB local |
| Multiplateforme | ❌ Windows uniquement |
| Forte concurrence | ⚠️ ~10-15 utilisateurs max |

## Migration depuis du code ADO.NET classique

NDXAccess remplace le `OleDbConnection`/`OleDbCommand`/`OleDbDataAdapter` répétitif par  
une API fluide avec gestion du cycle de vie, logging et health checks — tout en gardant  
l'accès au `OleDbConnection` sous-jacent via la propriété `Connection` si besoin.  
