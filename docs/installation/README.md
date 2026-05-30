# Installation et prérequis

NDXAccess s'appuie sur le **provider OLE DB Microsoft ACE** (Access Database Engine).  
Ce composant n'est **pas** fourni avec Windows ni avec .NET : il doit être installé,  
**dans la bonne architecture (x86 ou x64)**.

## 1. Prérequis

| Élément | Détail |
|---------|--------|
| OS | Windows uniquement (le provider ACE n'existe pas sur Linux/macOS) |
| .NET | .NET 8.0 (cible `net8.0-windows`) |
| Provider | `Microsoft.ACE.OLEDB.16.0` (Access Database Engine 2016 Redistributable, gratuit) |
| IDE | Visual Studio Community 2022 (ou `dotnet` CLI) |

> Le provider ACE 16.0 lit aussi bien les `.accdb` (Access 2007+) que les `.mdb` (Jet, legacy).

## 2. Installer l'Access Database Engine

Téléchargez le redistribuable gratuit :
**Microsoft Access Database Engine 2016 Redistributable**
https://www.microsoft.com/download/details.aspx?id=54920

Deux fichiers sont proposés :

- `accessdatabaseengine.exe` → **x86 (32 bits)**
- `accessdatabaseengine_X64.exe` → **x64 (64 bits)**

### Quelle version choisir ?

**L'architecture du provider doit correspondre à celle de votre application**, PAS à
celle de Windows :

| Votre application tourne en… | Installez l'ACE… |
|------------------------------|------------------|
| 64 bits (`x64`) | x64 |
| 32 bits (`x86`) | x86 |
| `AnyCPU` sur Windows 64 bits | x64 (le process sera 64 bits par défaut) |
| `AnyCPU` avec « Préférer 32 bits » | x86 |

NDXAccess détecte ce problème automatiquement et lève une
`AccessProviderNotFoundException` avec un message explicite si l'architecture ne
correspond pas (voir §4).

### Installer x86 ET x64 côte à côte

Office bloque par défaut l'installation simultanée des deux architectures. Pour les  
installer côte à côte (utile sur un poste de dev), utilisez l'option silencieuse :  

```powershell
.\accessdatabaseengine_X64.exe /quiet
.\accessdatabaseengine.exe /quiet
```

## 3. Construire l'application en x86 ou x64

Le binaire NDXAccess est **AnyCPU** (neutre). C'est l'**exécutable hôte** qui décide du  
bitness, donc de l'ACE requis. Le projet expose les plateformes `AnyCPU`, `x86` et `x64`.  

```powershell
# Build par défaut (AnyCPU)
dotnet build NDXAccess.sln

# Build 32 bits (nécessite ACE x86)
dotnet build NDXAccess.sln -p:Platform=x86

# Build 64 bits (nécessite ACE x64)
dotnet build NDXAccess.sln -p:Platform=x64
```

Dans Visual Studio 2022 : **Générer ▸ Gestionnaire de configurations** pour choisir
`x86` ou `x64`.

Pour forcer le bitness d'un exécutable applicatif :

```xml
<!-- Forcer 64 bits -->
<PlatformTarget>x64</PlatformTarget>

<!-- Forcer 32 bits -->
<PlatformTarget>x86</PlatformTarget>
```

## 4. Vérifier l'installation du provider

```vb
Imports NDXAccess

Console.WriteLine(AccessProviderHelper.CurrentProcessArchitecture)      ' x86 ou x64  
Console.WriteLine(AccessProviderHelper.IsProviderAvailable("Microsoft.ACE.OLEDB.16.0"))  

' Lève une exception détaillée si le provider manque pour cette architecture :
AccessProviderHelper.EnsureProviderAvailable("Microsoft.ACE.OLEDB.16.0")
```

Si le provider est absent ou dans la mauvaise architecture, le message indique :  
l'architecture du processus, les providers détectés, et le lien de téléchargement.  

## 5. Le fichier de verrou `.laccdb`

À l'ouverture d'une base, ACE crée un fichier `.laccdb` (ou `.ldb` pour `.mdb`) à côté
du `.accdb`. Il est supprimé quand **toutes** les connexions sont fermées. NDXAccess  
ferme proprement les connexions (Dispose/auto-close) et libère le pool OLE DB pour la  
connexion principale, ce qui relâche le verrou.  

## 6. Limites à connaître

- **Taille** : 2 Go maximum par fichier `.accdb` (surveillé par le HealthCheck).
- **Concurrence** : Access reste fiable jusqu'à ~10-15 utilisateurs simultanés.
- **Async** : OLE DB n'a pas de vraie asynchronie (async « de façade »).
- **Planification** : pas d'Event Scheduler → utilisez le Planificateur de tâches Windows.
