# Contribuer à NDXAccess

Merci de votre intérêt ! Ce guide explique comment compiler, tester et proposer des  
modifications.  

## Prérequis

- **Windows** (le provider ACE est Windows uniquement).
- **.NET SDK 8.0+** (le projet cible `net8.0-windows`).
- **Microsoft Access Database Engine 2016** dans l'architecture utilisée — requis seulement
  pour les **tests d'intégration** (voir [docs/installation](docs/installation/README.md)).
- Optionnel : **Visual Studio 2022** ou **VS Code**.

## Compiler

```powershell
dotnet build NDXAccess.sln                 # AnyCPU  
dotnet build NDXAccess.sln -p:Platform=x86 # 32 bits  
dotnet build NDXAccess.sln -p:Platform=x64 # 64 bits  
```

La solution doit compiler **sans avertissement**.

## Tester

```powershell
dotnet test                                  # tout  
dotnet test --filter "Category=Unit"         # unitaires (aucun ACE requis)  
dotnet test --filter "Category=Integration"  # intégration (ACE requis, sinon ignorés)  
```

- Les tests d'intégration créent une base `.accdb` temporaire (ADOX) puis la suppriment.
- Sans ACE installé, ils sont **ignorés** (Skip), pas en échec.
- Détails : [docs/tests](docs/tests/README.md).

## Conventions de code (VB.NET)

- `Option Strict On`, `Option Explicit On` — aucune liaison tardive.
- **Doc XML** (`'''`) sur **tous** les membres publics.
- Pièges VB à connaître (déjà rencontrés dans ce projet) :
  - Pas de `Await` dans un bloc `Catch`/`Finally`/`SyncLock`.
  - Pas de `Await Using` (utiliser `Using` synchrone ou `Await x.DisposeAsync()` manuel).
  - Une fonction `Async` ne peut pas retourner `ValueTask` (envelopper une `Task`).
  - VB est **insensible à la casse** : éviter une variable `x` homonyme d'un type/classe `X`.
  - `RootNamespace` est vide ; les fichiers déclarent explicitement `Namespace NDXAccess`.
- **Paramètres positionnels** : OLE DB/Access utilise `?` (l'ordre compte), pas `@nom`.

## Spécificités Access à respecter

Ne réintroduisez pas de fonctionnalités qui n'existent pas dans Access :  
pas d'Event Scheduler, pas de procédures stockées OUT/INOUT, pas de vrai async, Windows uniquement.  
Voir la section « Ce qu'Access sait et ne sait pas faire » du [README](README.md).  

## Ajouter un exemple

Les exemples vivent dans `examples/` et forment un **projet compilé**
(`NDXAccess.Examples.vbproj`) : tout nouvel exemple doit donc compiler, et idéalement être
couvert par un test d'intégration (voir `tests/.../Integration/MoreExamplesIntegrationTests.vb`).

## Processus de Pull Request

1. Créez une branche depuis la branche par défaut.
2. Ajoutez/mettez à jour les **tests** et la **documentation** (dont le [CHANGELOG](CHANGELOG.md)).
3. Vérifiez `dotnet build` (0 avertissement) et `dotnet test` (tout vert).
4. Décrivez clairement le « pourquoi » dans la PR.

Merci !
