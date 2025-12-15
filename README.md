# Rise Outlook Add-in - LinkedIn Search

Add-in Outlook qui permet de rechercher rapidement une personne sur LinkedIn depuis un email ou un contact.

## Fonctionnalités

- **Recherche LinkedIn depuis un email** : Cliquez sur le bouton dans un email pour rechercher l'expéditeur sur LinkedIn
- **Recherche LinkedIn depuis un contact** : Utilisez le bouton dans la section People pour rechercher un contact sur LinkedIn

## Installation et développement

### Prérequis

- Node.js (version 14 ou supérieure)
- npm ou yarn

### Installation des dépendances

```bash
npm install
```

### Configuration HTTPS (requis pour les add-ins Office)

Les add-ins Office nécessitent HTTPS. Pour le développement local, créez un certificat auto-signé :

```bash
openssl req -x509 -newkey rsa:2048 -keyout localhost-key.pem -out localhost.pem -days 365 -nodes
```

**Note Windows** : Si vous n'avez pas OpenSSL installé, vous pouvez utiliser Git Bash ou installer OpenSSL pour Windows.

### Démarrage du serveur

```bash
npm start
```

Le serveur démarrera sur `https://localhost:3000` (ou `http://localhost:3000` si les certificats ne sont pas disponibles).

### Chargement de l'add-in dans Outlook

1. Ouvrez Outlook
2. Allez dans **Fichier** > **Gérer les compléments** (ou **Options** > **Compléments**)
3. Cliquez sur **Mes compléments** > **Ajouter un complément personnalisé** > **Ajouter depuis un fichier**
4. Sélectionnez le fichier `manifest.xml`
5. Acceptez les avertissements de sécurité si nécessaire

### Test

1. Ouvrez un email dans Outlook
2. Vous devriez voir le bouton "Rechercher sur LinkedIn" dans le ruban
3. Cliquez sur le bouton pour rechercher l'expéditeur sur LinkedIn

## Structure du projet

```
.
├── manifest.xml          # Manifest de l'add-in
├── commands.html         # Page HTML pour les commandes
├── commands.js           # Logique JavaScript pour la recherche LinkedIn
├── server.js             # Serveur Express pour le développement
├── package.json          # Dépendances et scripts
└── assets/               # Icônes de l'add-in
```

## Fonctionnement

L'add-in utilise les ExtensionPoints suivants :
- **MessageReadCommandSurface** : Pour les emails
- **AppointmentOrganizerCommandSurface** : Pour les rendez-vous (section People)
- **AppointmentAttendeeCommandSurface** : Pour les participants aux rendez-vous

Quand vous cliquez sur le bouton, l'add-in :
1. Récupère le nom complet (displayName) de l'expéditeur ou du contact
2. Encode le nom pour l'URL
3. Ouvre une recherche LinkedIn dans une nouvelle fenêtre

## Validation du manifest

Pour valider le manifest avant le déploiement :

```bash
npm run validate
```

## Déploiement sur GitHub Pages

L'add-in est hébergé sur GitHub Pages à l'adresse : **https://rise-4.github.io/all--outlook-linkedin--addin/**

Toutes les URLs dans `manifest.xml` pointent vers cette adresse. Pour déployer :

1. Poussez tous les fichiers du projet dans le dépôt GitHub
2. Activez GitHub Pages dans les paramètres du dépôt (Settings > Pages)
3. Sélectionnez la branche principale comme source
4. Les fichiers seront disponibles à l'URL : `https://rise-4.github.io/all--outlook-linkedin--addin/`

### Structure des fichiers sur GitHub Pages

Assurez-vous que la structure suivante est respectée :
```
/
├── manifest.xml
├── commands.html
├── commands.js
├── index.html (optionnel, pour FormSettings)
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
└── ...
```

## Développement local (optionnel)

Pour tester localement avant de déployer sur GitHub Pages, vous pouvez utiliser le serveur local :

### Configuration HTTPS (requis pour les add-ins Office)

Les add-ins Office nécessitent HTTPS. Pour le développement local, créez un certificat auto-signé :

```bash
openssl req -x509 -newkey rsa:2048 -keyout localhost-key.pem -out localhost.pem -days 365 -nodes
```

**Note Windows** : Si vous n'avez pas OpenSSL installé, vous pouvez utiliser Git Bash ou installer OpenSSL pour Windows.

### Démarrage du serveur local

```bash
npm start
```

Le serveur démarrera sur `https://localhost:3000` (ou `http://localhost:3000` si les certificats ne sont pas disponibles).

**Important** : Si vous testez localement, vous devrez temporairement modifier les URLs dans `manifest.xml` pour pointer vers `https://localhost:3000` au lieu de l'URL GitHub Pages.

