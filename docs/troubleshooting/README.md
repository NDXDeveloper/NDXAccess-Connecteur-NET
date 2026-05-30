# Dépannage & FAQ — NDXAccess

Guide des erreurs les plus fréquentes avec Access/ACE et leur solution. NDXAccess
**traduit** la plupart de ces erreurs en `AccessQueryException` avec un message clair
(le code natif ACE est souvent `0`, la détection se fait alors par motif de message).

---

## 1. Provider / installation (x86 vs x64)

### `AccessProviderNotFoundException` à la création de la connexion
**Message** : « Le provider OLE DB 'Microsoft.ACE.OLEDB.16.0' n'est pas disponible pour un processus xNN… »

**Cause** : l'Access Database Engine (ACE) n'est pas installé, ou pas dans la **même
architecture** que le processus (un ACE x64 est invisible depuis un process x86, et inversement).

**Solution** :
1. Vérifier l'architecture du process et les providers visibles :
   ```vb
   Console.WriteLine(AccessProviderHelper.CurrentProcessArchitecture)         ' x86 / x64
   For Each p In AccessProviderHelper.GetAvailableProviders() : Console.WriteLine(p) : Next
   ```
2. Installer **Microsoft Access Database Engine 2016** dans la bonne architecture, OU forcer
   le bitness de l'appli (`<PlatformTarget>x64</PlatformTarget>` ou `x86`).
3. Voir [docs/installation](../installation/README.md) (dont l'installation x86 **et** x64 côte à côte).

### « The 'Microsoft.ACE.OLEDB.16.0' provider is not registered on the local machine »
Même cause que ci-dessus (ACE absent ou mauvaise architecture). Même solution.  
Astuce : si vous n'avez que l'ACE 12.0, définissez `options.Provider = "Microsoft.ACE.OLEDB.12.0"`.  

### « Could not find installable ISAM »
**Cause** : chaîne de connexion mal formée — typiquement `Extended Properties` mal échappé
(lecture Excel/CSV), ou guillemets manquants.

**Solution** : vérifier la chaîne. Exemple correct (CSV) :
```
Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\dossier;Extended Properties="text;HDR=Yes;FMT=Delimited"
```

---

## 2. Verrouillage & fichier `.laccdb`

### 3051 — « …cannot open or write to the file… already opened exclusively / need permission »
**Causes possibles** :
- Le fichier est ouvert **en exclusif** par une autre connexion/application.
- Le dossier n'autorise pas l'écriture (ACE doit créer le fichier `.laccdb` à côté du `.accdb`).
- Le fichier est en lecture seule sur le disque.

**Solution** : fermer les autres accès exclusifs ; donner les droits d'écriture sur le  
**dossier** ; éviter `OpenExclusive = True` si plusieurs accès sont nécessaires.  

### 3050 — « Could not lock file » (contention `.laccdb`)
**Cause** : verrou transitoire en accès concurrent.

**Solution** : NDXAccess **réessaie automatiquement** ces erreurs transitoires
(`EnableRetryOnTransientErrors`, `MaxRetries`, `RetryBaseDelayMs`). Si ça persiste :
réduire la concurrence (Access tient ~10-15 utilisateurs max), vérifier les droits du dossier.

### Le compactage échoue : « …already opened by user 'Admin' »
**Cause** : une connexion (ou une session OLE DB encore poolée) détient toujours le fichier.

**Solution** : disposer toutes les connexions avant de compacter. `CompactDatabase` ferme la
connexion courante et appelle `OleDbConnection.ReleaseObjectPool()` ; mais une session ADOX  
fraîchement créée (ex. `CreateDatabase`) peut rester ~60 s. Évitez de créer puis compacter  
immédiatement le même fichier dans le même process.  

---

## 3. Requêtes & paramètres

### « Too few parameters. Expected N »
**Cause** : un nom de **colonne ou de table est mal orthographié**. Access interprète tout
identifiant inconnu comme un paramètre attendu.

**Solution** : vérifier l'orthographe des colonnes/tables ; mettre entre crochets les noms
contenant des espaces : `SELECT [Mon Champ] FROM [Ma Table]`.

### « Data type mismatch in criteria expression »
**Causes** :
- L'**ordre** des `?` ne correspond pas à l'ordre des valeurs (OLE DB est **positionnel**).
- Type incompatible (ex. `DateTime` envoyé en DBTimeStamp, `Decimal` vers `CURRENCY`).

**Solution** : aligner l'ordre des paramètres sur les `?`. NDXAccess type déjà
automatiquement `Date → OleDbType.Date` et `Decimal → OleDbType.Currency`. Pour une colonne
`DECIMAL(p,s)` haute précision, passez un `OleDbParameter` explicite.

### Paramètres nommés `@nom` qui « ne marchent pas »
**Cause** : OLE DB n'accepte pas les paramètres nommés dans le SQL — uniquement `?` positionnels.

**Solution** : utiliser `?` et l'ordre, OU les méthodes `ExecuteQueryNamed` /
`ExecuteScalarNamed` / `ExecuteNonQueryNamed` qui traduisent `@nom` → `?` automatiquement.

### Une procédure stockée avec paramètre OUT/INOUT ne fonctionne pas
**Cause** : Access n'a **pas** de vraies procédures stockées — seulement des requêtes
enregistrées à paramètres d'**entrée**.

**Solution** : utiliser une requête enregistrée (IN), puis lire le résultat via un `SELECT`
(ex. `SELECT @@IDENTITY` après une insertion).

---

## 4. Sécurité / mot de passe

### 3031 — « Not a valid password »
**Cause** : mot de passe absent ou incorrect pour une base protégée.

**Solution** : `options.Password = "…"`. Pour créer une base protégée :
`AccessConnection.CreateDatabase(path, password)`.

---

## 5. Format & taille de fichier

### 3343 — « Unrecognized database format »
**Causes** : fichier corrompu, ou créé par un moteur incompatible, ou `.accdb` ouvert avec
un provider trop ancien.

**Solution** : utiliser ACE 16, **compacter/réparer** (`CompactDatabase`), restaurer une
sauvegarde si corrompu.

### Limite des 2 Go atteinte
**Cause** : un fichier Access ne peut pas dépasser **2 Go** (limite stricte du moteur).

**Solution** : surveiller via le health check —
```vb
Dim info = New AccessHealthCheck(factory).GetDatabaseInfo()  
If info.IsApproachingSizeLimit Then ' > 90 %  
    ' compacter, archiver, ou scinder ; au-delà : migrer vers un SGBD client-serveur
End If
```
Compacter régulièrement libère l'espace des enregistrements supprimés.

---

## 6. Performances & concurrence

- **Concurrence** : Access reste fiable jusqu'à ~10-15 utilisateurs simultanés. Au-delà,
  envisager une migration.
- **Réseau** : un `.accdb` sur partage réseau est sensible à la latence et aux coupures
  (corruption possible). Préférer le local quand c'est possible.
- **Async « de façade »** : les méthodes `...Async` ne sont pas réellement parallèles
  (OLE DB n'a pas d'asynchronie native). Seule `CompactDatabaseAsync` est déportée sur un thread.
- **Insertion en masse** : utiliser `BulkInsert` (une seule transaction) plutôt que des
  `INSERT` en auto-commit.

---

## 7. Questions fréquentes

**NDXAccess fonctionne-t-il sous Linux/macOS ?** Non — le provider ACE est Windows uniquement.

**Faut-il Microsoft Office/Access installé ?** Non — seul l'**Access Database Engine**
redistribuable (gratuit) est requis.

**.mdb (ancien Jet) est-il supporté ?** Oui, via le provider ACE 16 (`Data Source` pointant le `.mdb`).

**Comment connaître les versions ?**
```vb
AccessConnection.Version       ' version de la bibliothèque NDXAccess  
connection.EngineVersion       ' version du moteur ACE (ex. 16.0.5011.1000)  
info.FileFormatVersion         ' version du format de fichier (ex. 04.00.0000)  
```

**Une `AccessQueryException` est levée au lieu d'une `OleDbException` ?** C'est voulu (message
clair). L'`OleDbException` d'origine reste dans `InnerException`. Pour la brute :
`options.TranslateErrors = False`.

---

Un cas non couvert ? Ouvrez une *issue* (voir [CONTRIBUTING](../../CONTRIBUTING.md)).
