# Sommaire - Documentation NDXAccess

## Guide de démarrage

- [README.md](README.md) - Présentation et démarrage rapide
- [Installation & architecture x86/x64](docs/installation/README.md) - Access Database Engine
- [Dépannage & FAQ](docs/troubleshooting/README.md) - erreurs courantes Access → solutions

## Documentation technique

### Bibliothèque source

- [Documentation du projet source](docs/source/README.md)
  - Architecture de la bibliothèque
  - Classes principales
  - Paramètres positionnels et requêtes enregistrées
  - Async « de façade »
  - Spécificités et limites du moteur Access

### Tests

- [Documentation des tests](docs/tests/README.md)
  - Structure des tests (98 tests)
  - Tests unitaires (37)
  - Tests d'intégration (61)
  - Skip automatique si ACE absent

### Exemples

- [Exemples d'utilisation](examples/README.md)
  - [CRUD de base](examples/BasicCrudExamples.vb)
  - [Requêtes enregistrées](examples/StoredQueryExamples.vb) - paramètres IN
  - [Transactions](examples/TransactionExamples.vb)
  - [Avancé](examples/AdvancedExamples.vb) - health checks, DI, concurrence
  - [Maintenance](examples/MaintenanceExamples.vb) - compactage, détection x86/x64
  - [Nouveautés v1.1](examples/NewFeaturesExamples.vb) - résilience, schéma, BulkInsert, mapping
  - [Entité Active Record](examples/EntityCrudExample.vb) - OleDbCommand / da.Fill / @@IDENTITY
  - [DataAdapter](examples/DataAdapterExamples.vb) - Fill/Update + CommandBuilder
  - [Démarrage tout-en-un](examples/GettingStartedExample.vb)
  - [Pagination](examples/PaginationExample.vb) - TOP n / NOT IN
  - [DDL & relations](examples/SchemaDdlExample.vb) - index unique, clé étrangère
  - [Excel / CSV via ACE](examples/ExcelCsvExample.vb)
  - [Logging](examples/LoggingExample.vb)
  - [Liaison WinForms](examples/WinFormsBindingExample.vb) - DataGridView éditable

### Projet & contribution

- [CHANGELOG](CHANGELOG.md) - Historique des versions
- [CONTRIBUTING](CONTRIBUTING.md) - Compiler, tester, conventions VB
- [SECURITY](SECURITY.md) - Politique de sécurité
- [CODE_OF_CONDUCT](CODE_OF_CONDUCT.md) - Code de conduite
- [LICENSE](LICENSE) - Licence MIT

## Références

### Classes principales

| Classe | Description |
|--------|-------------|
| `AccessConnection` | Connexion principale (sync + async) |
| `AccessConnectionOptions` | Options de configuration |
| `AccessConnectionFactory` | Factory pour créer des connexions |
| `AccessHealthCheck` | Santé de la base + surveillance des 2 Go |
| `AccessProviderHelper` | Détection x86/x64 + version du moteur ACE (`GetEngineVersion`) |
| `NDXAccessInfo` | Version de la bibliothèque (Version / InformationalVersion / ProductName) |

### Interfaces

| Interface | Description |
|-----------|-------------|
| `IAccessConnection` | Interface de connexion |
| `IAccessConnectionFactory` | Interface de factory |

### Exceptions

| Exception | Description |
|-----------|-------------|
| `AccessConnectionException` | Erreur générale du connecteur |
| `AccessProviderNotFoundException` | Provider ACE absent / mauvaise architecture |
| `AccessQueryException` | OleDbException traduite (message clair + code natif + IsTransient) |

### Options de connexion

| Propriété | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `DatabasePath` | String | "" | Chemin du fichier .accdb / .mdb |
| `Password` | String | "" | Mot de passe (Jet OLEDB:Database Password) |
| `Provider` | String | Microsoft.ACE.OLEDB.16.0 | Provider OLE DB |
| `ConnectionString` | String | Nothing | Chaîne complète (surcharge) |
| `OpenExclusive` | Boolean | False | Mode exclusif |
| `OpenReadOnly` | Boolean | False | Lecture seule |
| `PersistSecurityInfo` | Boolean | False | Conserver les infos de sécurité |
| `SystemDatabasePath` | String | "" | Fichier de groupe de travail .mdw |
| `IsPrimaryConnection` | Boolean | False | Connexion principale |
| `AutoCloseTimeoutMs` | Integer | 60000 | Fermeture auto (ms) |
| `DisableAutoClose` | Boolean | False | Désactiver fermeture auto |
| `ValidateProvider` | Boolean | True | Vérifier le provider (x86/x64) |
| `EnableRetryOnTransientErrors` | Boolean | True | Retry sur verrous transitoires |
| `MaxRetries` | Integer | 3 | Tentatives maximales |
| `RetryBaseDelayMs` | Integer | 100 | Back-off de base (ms) |
| `TranslateErrors` | Boolean | True | Traduire les OleDbException |

## Fonctionnalités

### CRUD (paramètres positionnels `?`)

- INSERT, SELECT (DataTable / DataReader / scalaire), UPDATE, DELETE
- Variantes synchrones (ParamArray) et asynchrones (de façade)
- Typage auto : Date → OleDbType.Date, Decimal → OleDbType.Currency

### Transactions

- BeginTransaction / Commit / Rollback (+ variantes async)
- Niveau d'isolation ReadCommitted (limite d'Access)

### Requêtes enregistrées

- Paramètres d'**entrée uniquement** (pas de OUT/INOUT)
- `ExecuteStoredQuery` / `ExecuteStoredQueryNonQuery` (+ async)

### Maintenance

- `CompactDatabase` (synchrone)
- `CompactDatabaseAsync` (vrai async)
- `CreateDatabase` / `CreateDatabaseAsync` (création d'un .accdb vide)

### Résilience & productivité (v1.1)

- Retry automatique sur verrous transitoires (back-off exponentiel)
- Traduction des erreurs ACE → `AccessQueryException`
- Helpers de schéma : `TableExists`, `GetTableNames`, `GetQueryNames`, `GetColumns`
- `BulkInsert` / `BulkInsertAsync` (transaction unique)
- Mapping objet `ExecuteQuery(Of T)` (sync + async)
- Paramètres nommés : `ExecuteNonQueryNamed`, `ExecuteScalarNamed`, `ExecuteQueryNamed` (+ async)

### Robustesse Access

- Détection x86/x64 du provider (exception claire)
- Gestion du verrou `.laccdb` (fermeture + libération du pool)
- Surveillance de la limite des 2 Go (HealthCheck)
- Mot de passe encapsulé

### Autres

- Fermeture automatique des connexions inactives
- Historique des actions
- Logging avec ILogger
- Injection de dépendances (`AddNDXAccess`)

## Limites du moteur Access (et alternatives)

| Fonctionnalité | Raison / alternative |
|----------------|----------------------|
| Event Scheduler | N'existe pas dans Access → Planificateur de tâches Windows |
| Procédures stockées OUT/INOUT | Access n'a que des requêtes enregistrées (IN) |
| Vrai async | OLE DB n'a pas d'asynchronie native |
| Multiplateforme Linux | Provider ACE = Windows uniquement |
| Docker | Sans objet (base fichier locale) |

## Historique

### Version 1.2.0 (2026)

- API de version : bibliothèque (`NDXAccessInfo`, `AccessConnection.Version`), **moteur ACE**
  (`connection.EngineVersion`, `AccessProviderHelper.GetEngineVersion`) et **format de fichier**
  (`DatabaseInfo.FileFormatVersion`)
- `connection.ProviderName` exposé
- 98 tests (37 unitaires + 61 d'intégration)

### Version 1.1.0 (2026)

- Résilience : retry automatique sur verrous transitoires (back-off exponentiel)
- Traduction des erreurs ACE en `AccessQueryException` (message clair + code natif + IsTransient)
- Helpers de schéma : `TableExists`, `GetTableNames`, `GetQueryNames`, `GetColumns`
- `CreateDatabase` / `CreateDatabaseAsync` (création programmatique via ADOX)
- `BulkInsert` / `BulkInsertAsync` (insertion en masse en une transaction)
- Mapping objet `ExecuteQuery(Of T)` (micro-ORM par réflexion)
- Paramètres nommés (`@nom` → `?`) via les méthodes `...Named`
- `CancellationToken` honoré, note thread-safety, dispose de commande documenté
- Intégration continue (GitHub Actions) + packaging NuGet/SourceLink
- Exemple d'entité « Active Record » (OleDbCommand / da.Fill / @@IDENTITY) testé de bout en bout
- Tests de résilience (retry sur verrou), mot de passe, injection de dépendances, lecture seule, NULL, auto-close
- Exemples de patterns réels (DataAdapter Fill/Update, pagination, DDL/relations, CSV via ACE, WinForms)
- API de version de la bibliothèque (`NDXAccessInfo`, `AccessConnection.Version`)
- 92 tests (33 unitaires + 59 d'intégration)

### Version 1.0.0 (2026)

- Connecteur Access initial pour VB.NET (.NET 8)
- API CRUD paramétrée (positionnelle), transactions
- Requêtes enregistrées (paramètres IN)
- Compactage via DAO (sync + async réel)
- Health checks avec surveillance des 2 Go
- Détection x86/x64 du provider ACE
- 44 tests (22 unitaires + 22 d'intégration)
- Support des builds AnyCPU / x86 / x64

## Ressources externes

- [Access Database Engine 2016 Redistributable](https://www.microsoft.com/download/details.aspx?id=54920)
- [Référence SQL Access (Microsoft)](https://support.microsoft.com/office/access-sql-basic-concepts-vocabulary-and-syntax-444d0303-cde1-424e-9a74-e8dc3e460671)
- [System.Data.OleDb](https://learn.microsoft.com/dotnet/api/system.data.oledb)
- [Planificateur de tâches Windows](https://learn.microsoft.com/windows/win32/taskschd/task-scheduler-start-page)
