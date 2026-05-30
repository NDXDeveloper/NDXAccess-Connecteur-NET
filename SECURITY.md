# Politique de sécurité

## Versions prises en charge

| Version | Prise en charge |
|---------|-----------------|
| 1.2.x   | ✅ |
| < 1.2   | ❌ (mettre à niveau) |

## Signaler une vulnérabilité

Merci de **ne pas** ouvrir d'issue publique pour une faille de sécurité.

Envoyez un e-mail à **NDXDev@gmail.com** avec :
- une description du problème et de son impact ;
- les étapes de reproduction ;
- la version de NDXAccess et l'environnement (OS, architecture, version d'ACE).

Vous recevrez un accusé de réception, et un correctif sera publié dès que possible avec  
mention (si vous le souhaitez) dans le [CHANGELOG](CHANGELOG.md).  

## Bonnes pratiques côté utilisateur

- **Mots de passe** : ne codez pas en dur le mot de passe de la base. Utilisez la
  configuration / un coffre de secrets, et évitez `PersistSecurityInfo = True`.
- **Chaînes de connexion** : ne les journalisez pas en clair (elles peuvent contenir
  `Jet OLEDB:Database Password`).
- **Fichiers `.accdb`** : protégez l'accès au fichier et au dossier (le chiffrement du
  mot de passe Access est faible — pour des données sensibles, chiffrez le support ou
  utilisez un SGBD adapté).
- **Permissions** : accordez les droits minimaux nécessaires sur le dossier de la base
  (ACE doit pouvoir créer le fichier de verrou `.laccdb`).

## Portée

NDXAccess est une bibliothèque Windows s'appuyant sur le provider Microsoft ACE OLE DB.  
Les vulnérabilités du moteur ACE lui-même relèvent de Microsoft (maintenez l'Access  
Database Engine à jour).  
