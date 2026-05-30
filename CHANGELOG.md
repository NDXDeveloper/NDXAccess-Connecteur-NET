# Changelog

Toutes les modifications notables de ce projet sont documentées dans ce fichier.

Le format s'inspire de [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/),  
et le projet suit le [Versioning Sémantique](https://semver.org/lang/fr/).  

## [1.2.0] - 2026-05-30

### Ajouté
- API de version :
  - `NDXAccessInfo` (`Version`, `InformationalVersion`, `ProductName`) et raccourci `AccessConnection.Version` — version de la **bibliothèque**.
  - `IAccessConnection.EngineVersion` et `AccessProviderHelper.GetEngineVersion()` — version du **moteur ACE** (DLL `ACEOLEDB`, via le registre ; repli sur la génération « 16.0 »).
  - `IAccessConnection.ProviderName`.
  - `DatabaseInfo.FileFormatVersion` — version du **format de fichier** (ex. « 04.00.0000 »).
- Traits xUnit `Category=Unit` / `Category=Integration` pour filtrer l'exécution des tests.

### Modifié
- `DatabaseInfo.EngineVersion` renvoie désormais la version réelle du moteur ACE
  (auparavant la version du format de fichier, déplacée vers `FileFormatVersion`).

## [1.1.0] - 2026-05-29

### Ajouté
- **Résilience** : retry automatique avec back-off sur erreurs de verrou transitoires
  (`EnableRetryOnTransientErrors`, `MaxRetries`, `RetryBaseDelayMs`).
- **Traduction des erreurs** : `OleDbException` → `AccessQueryException` (message clair, code
  natif, `IsTransient`) ; désactivable via `TranslateErrors`.
- **Helpers de schéma** : `TableExists`, `GetTableNames`, `GetQueryNames`, `GetColumns`.
- **Création de base** : `AccessConnection.CreateDatabase` / `CreateDatabaseAsync` (ADOX).
- **Insertion en masse** : `BulkInsert` / `BulkInsertAsync` (transaction unique).
- **Mapping objet** : `ExecuteQuery(Of T)` / `ExecuteQueryAsync(Of T)`.
- **Paramètres nommés** : `ExecuteNonQueryNamed`, `ExecuteScalarNamed`, `ExecuteQueryNamed` (+ async).
- Prise en compte du `CancellationToken`, note thread-safety.
- Exemples supplémentaires (entité Active Record, DataAdapter, démarrage tout-en-un,
  pagination, DDL/relations, Excel/CSV, logging, WinForms) — projet d'exemples **compilé**.
- Intégration continue (GitHub Actions) + packaging NuGet/SourceLink (symboles `.snupkg`).

## [1.0.0] - 2026-05-29

### Ajouté
- Connecteur Microsoft Access initial pour VB.NET (`.NET 8`, provider `Microsoft.ACE.OLEDB.16.0`).
- `AccessConnection` / `IAccessConnection`, `AccessConnectionOptions`, `AccessConnectionFactory`,
  `AccessHealthCheck`, `AccessProviderHelper`, extensions DI (`AddNDXAccess`).
- CRUD paramétré (positionnel `?`), transactions, requêtes enregistrées (paramètres IN),
  compactage (`CompactDatabase`), détection x86/x64, surveillance de la limite 2 Go.
- API synchrone et asynchrone (async « de façade », sauf `CompactDatabaseAsync`).
- Support des builds AnyCPU / x86 / x64.

[1.2.0]: https://github.com/NDXDeveloper/NDXAccess-Connecteur-NET/releases/tag/v1.2.0
[1.1.0]: https://github.com/NDXDeveloper/NDXAccess-Connecteur-NET/releases/tag/v1.1.0
[1.0.0]: https://github.com/NDXDeveloper/NDXAccess-Connecteur-NET/releases/tag/v1.0.0
