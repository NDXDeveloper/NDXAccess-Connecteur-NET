---
name: Rapport de bug  
about: Signaler un dysfonctionnement  
title: "[Bug] "  
labels: bug  
---

## Description
Description claire et concise du problème.

## Reproduction
Étapes pour reproduire (idéalement un extrait de code minimal) :

```vb
' …
```

## Comportement attendu
Ce qui devrait se passer.

## Comportement observé
Ce qui se passe réellement (message d'erreur complet, `AccessQueryException.Message`,
`InnerException`, code natif `NativeError` si disponible).

## Environnement
- **NDXAccess** : <!-- AccessConnection.Version -->
- **Moteur ACE** : <!-- connection.EngineVersion, ex. 16.0.5011.1000 -->
- **Architecture du process** : <!-- x86 / x64 (AccessProviderHelper.CurrentProcessArchitecture) -->
- **ACE installé** : <!-- x86 / x64 / les deux -->
- **.NET** : <!-- ex. .NET 8 -->
- **OS** : <!-- ex. Windows 11 23H2 -->
- **Type de base** : <!-- .accdb / .mdb, protégée par mot de passe ? réseau ou local ? -->

## Vérifications préalables
- [ ] J'ai consulté le [guide de dépannage](../../docs/troubleshooting/README.md).
- [ ] Le provider ACE est installé dans la **même architecture** que l'application.
