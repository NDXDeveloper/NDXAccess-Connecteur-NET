# Documentation des tests NDXAccess

Le projet comporte **98 tests** (xUnit) répartis entre tests unitaires (sans base) et  
tests d'intégration (sur une vraie base `.accdb`).  

## Prérequis pour les tests d'intégration

Les tests d'intégration s'exécutent sur une vraie base `.accdb`. Ils nécessitent donc le  
provider **Microsoft Access Database Engine (ACE)** installé **dans l'architecture du  
processus de test** (x64 par défaut). Voir [docs/installation](../installation/README.md).  

- **ACE présent (bonne architecture)** → les tests d'intégration s'exécutent (une base
  `.accdb` temporaire est créée via ADOX puis supprimée).
- **ACE absent / mauvaise architecture** → ils sont **ignorés** (Skip, via
  `<SkippableFact>`), avec un message indiquant l'architecture attendue — **jamais en
  échec**. Les tests unitaires, eux, passent toujours.

> ⚠️ Aucune installation, aucune configuration de connexion ni aucun serveur n'est requis :  
> tout est géré par la fixture `AccessDatabaseFixture`.

## Exécution

```powershell
# Tous les tests (unitaires + intégration)
dotnet test

# Uniquement les tests d'intégration
dotnet test --filter "Category=Integration"

# Uniquement les tests unitaires (aucune base / aucun ACE requis)
dotnet test --filter "Category=Unit"

# Une classe précise
dotnet test --filter "FullyQualifiedName~AccessConnectionTests"

# Un test précis
dotnet test --filter "FullyQualifiedName~AccessConnectionTests.Open_ShouldOpenConnection"

# Sortie détaillée (affiche les tests ignorés et la raison du Skip)
dotnet test --logger "console;verbosity=normal"

# Avec couverture de code
dotnet test --collect:"XPlat Code Coverage"
```

Les catégories reposent sur l'attribut `<Trait("Category", "Integration")>` /
`<Trait("Category", "Unit")>` posé sur chaque classe de test.

### Vérifier qu'ils se sont bien exécutés (et non ignorés)

Dans la ligne de résumé :

- `… ignorée(s) : 0 …` → les tests d'intégration se sont **exécutés** (ACE présent).
- `… ignorée(s) : 61 …` → ACE absent ou mauvaise architecture → ils ont été **ignorés**
  (le message de Skip précise l'architecture attendue).

### Choisir l'architecture (x86 / x64)

L'architecture du test doit correspondre à celle de l'ACE installé :

```powershell
dotnet test -p:Platform=x64   # nécessite l'ACE x64 (défaut)  
dotnet test -p:Platform=x86   # nécessite l'ACE x86  
```

### Depuis l'IDE

- **Visual Studio 2022** : *Test ▸ Explorateur de tests*, puis « Exécuter tout » (ou clic
  droit sur une classe/un test). Les catégories apparaissent comme traits (regroupement
  possible par « Trait »).
- **VS Code** : tâche `test` préconfigurée (`.vscode/tasks.json`), ou la configuration de
  débogage « Tests NDXAccess » (`.vscode/launch.json`).

## Tests unitaires (37) — aucune base requise

### AccessConnectionOptions (11)
- Valeurs par défaut, constante 2 Go, défauts de résilience
- `BuildConnectionString` : propriétés, mot de passe (présent/absent), mode exclusif,
  mode lecture seule, surcharge par `ConnectionString`
- `ResolveProviderName`, `ResolveDatabasePath`

### AccessConnectionFactory (7)
- Constructeur avec options nulles (exception), avec chemin
- Création de connexion standard / principale / configurée
- IDs uniques, état initial

### AccessProviderHelper (9)
- Architecture du processus (x86/x64)
- Liste des providers (jamais nulle)
- Provider inconnu / nul → False
- `EnsureProviderAvailable` → `AccessProviderNotFoundException` détaillée
- `GetProviderGeneration` parse la génération ("16.0", "12.0", "4.0")
- `GetEngineVersion` retourne la version du moteur ACE (DLL ou génération)

### NamedParameterParser (6)
- Traduction `@nom` → `?` et valeurs ordonnées
- Paramètre répété, littéraux chaîne ignorés, `@@IDENTITY` préservé
- Clés avec/sans préfixe `@`, absence de paramètres

### NDXAccessInfo / version (4)
- `Version`, `InformationalVersion`, `ProductName`
- `AccessConnection.Version` = `InformationalVersion`

## Tests d'intégration (61) — base `.accdb` réelle

### AccessConnection (16)
- Ouverture/fermeture (sync + async)
- `ProviderName` et `EngineVersion` disponibles sans ouverture
- Scalaire, scalaire paramétré
- INSERT/SELECT/UPDATE/DELETE paramétrés (`?`)
- DataTable, DataReader
- Types multiples (texte, entier, CURRENCY, DATETIME)
- `@@IDENTITY` après insertion
- Transactions commit / rollback
- Historique d'actions, IDs uniques, DisposeAsync

### AccessHealthCheck (4)
- Base saine
- Fichier manquant → résultat « unhealthy »
- Informations base (taille, limite 2 Go, nombre de tables)
- Versions moteur ACE et format de fichier (`EngineVersion` / `FileFormatVersion`)

### Requêtes enregistrées (2)
- SELECT paramétré (paramètre IN)
- Requête d'action (DELETE) → lignes affectées

### Maintenance / compactage (3)
- `CompactDatabaseAsync` préserve les données (in-place)
- Compactage refusé pendant une transaction
- `CompactDatabase` (synchrone) vers un nouveau fichier

### Schéma & création de base (6)
- `TableExists` (insensible à la casse), `GetTableNames`, `GetColumns`, `GetQueryNames`
- `CreateDatabase` crée un `.accdb` utilisable
- `CreateDatabase` échoue si le fichier existe

### Insertion en masse & mapping (6)
- `BulkInsert` / `BulkInsertAsync` (200 lignes en une transaction)
- `ExecuteQuery(Of T)` mappe les lignes vers des objets (sync + async)
- `ExecuteQueryNamed` / `ExecuteScalarNamed` (paramètres `@nom`)

### Traduction des erreurs (3)
- SQL invalide → `AccessQueryException`
- Doublon de clé → exception traduite (OleDbException préservée)
- `TranslateErrors = False` → `OleDbException` brute

### Entité « Active Record » de l'exemple (7)
- Exerce `examples/EntityCrudExample.vb` de bout en bout (référence le projet d'exemples)
- `Save` (INSERT + `@@IDENTITY`), `LoadFromID`, `LoadFrom`, `Save` (UPDATE), `Remove`
- `ListeModules` (chargement paresseux de la collection enfant)
- `LoadFromIDViaConnector` (variante API haut niveau)
- Crée/supprime son propre schéma (`tb_Etapes`, `tb_Modules`, `tb_Etape_Modules`)

### Résilience (2)
- Retry sur verrou : ouverture exclusive concurrente → exception transitoire après back-off
  (comparaison du temps écoulé avec/sans retry)
- `CancellationToken` déjà annulé → `OperationCanceledException`

### Mot de passe (1)
- Création d'une base protégée, ouverture OK avec le bon mot de passe, échec (traduit)
  avec le mauvais / sans mot de passe

### Injection de dépendances (1)
- `AddNDXAccess` → résolution et utilisation de la factory, de la connexion et du health check

### Mode lecture seule & valeurs NULL (2)
- `OpenReadOnly` : SELECT autorisé, écriture refusée, `DatabaseInfo.IsReadOnly = True`
- NULL → défaut du type (scalaire et mapping `ExecuteQuery(Of T)`)

### Fermeture automatique (2)
- Connexion inactive fermée après `AutoCloseTimeoutMs`
- Connexion principale jamais fermée automatiquement

### Exemples « patterns réels » (6)
- `DataAdapterExamples` : `Fill` + modifications déconnectées + `Update` par lot
- `GettingStartedExample` : flux complet de bout en bout
- `PaginationExample` : pagination `TOP n` / `NOT IN`
- `SchemaDdlExample` : index unique et clé étrangère réellement appliqués
- `ExcelCsvExample` : lecture d'un CSV via ACE
- `LoggingExample` : un logger personnalisé capture l'activité du connecteur

## Remarques d'implémentation

- VB.NET interdit `Await` dans un bloc `Finally` : le nettoyage (`DROP TABLE`) est fait
  de façon **synchrone**.
- Chaque test d'intégration crée ses propres tables (`t_<guid>`) pour rester isolé,
  la collection `"Access"` garantissant une exécution séquentielle sur le fichier partagé.
