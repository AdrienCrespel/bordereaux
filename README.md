# Bordereaux 

`bordereaux.py` est un script Python conçu pour automatiser le téléchargement, le rognage et la fusion des bordereaux d'expédition reçus par email.

## Fonctionnalités

- Connexion à une boîte mail IMAP pour récupérer les emails non lus avec un sujet spécifique.
- Téléchargement des pièces jointes PDF des emails et organisation dans des répertoires par transporteur.
- Rognage des bordereaux d'expédition en fonction de règles spécifiques à chaque transporteur.
- Fusion des images rognées en un seul fichier PDF.

## Prérequis

- Python 3.x
- Un compte email avec accès IMAP (testé avec Gmail)
- Les bibliothèques Python suivantes : `python-dotenv`, `imaplib`, `email`, `pathlib`, `PyPDF2`

## Configuration du Compte Gmail

Pour utiliser ce script avec un compte Gmail, vous devez configurer un mot de passe d'application spécifique ou activer l'accès pour les applications moins sécurisées. Voici comment procéder :

### Option 1 : Mot de passe d'application (recommandé)

1. **Activer l'authentification à deux facteurs** : Si ce n'est pas déjà fait, activez l'authentification à deux facteurs sur votre compte Google.

2. **Générer un mot de passe d'application** :
   - Allez sur [cette page](https://myaccount.google.com/apppasswords).
   - Sélectionnez "Mail" pour l'application et choisissez votre appareil.
   - Cliquez sur "Générer" pour obtenir un mot de passe d'application.

3. **Utiliser le mot de passe d'application** : Utilisez ce mot de passe dans votre fichier `.env` à la place de votre mot de passe habituel.

### Option 2 : Accès pour les applications moins sécurisées

1. **Activer l'accès** : Allez sur [cette page](https://myaccount.google.com/lesssecureapps) et activez l'option "Autoriser les applications moins sécurisées".

**Remarque** : L'utilisation d'un mot de passe d'application est recommandée pour des raisons de sécurité.

## Installation

1. Clonez ce dépôt ou téléchargez le script `bordereaux.py`.

2. Installez les dépendances nécessaires :

    ```bash
    pip install python-dotenv PyPDF2
    ```

3. Créez un fichier `.env` dans le même répertoire que le script avec les informations d'identification de votre compte email :

    ```
    EMAIL_ADDRESS=votre_email@example.com
    EMAIL_PASSWORD=votre_mot_de_passe
    IMAP_SERVER=imap.gmail.com
    ```

## Utilisation

1. Assurez-vous que votre fichier `.env` est correctement configuré avec vos informations d'identification.

2. Exécutez le script :

   ```bash
   python bordereaux.py
   ```

3. Le script va :
    - Se connecter à votre boîte mail.
    - Télécharger les bordereaux d'expédition des emails non lus avec le sujet spécifié.
    - Rogner les bordereaux en fonction du transporteur.
    - Fusionner les images rognées en un seul fichier PDF dans le répertoire bordereaux.

## Structure du Projet

- `bordereaux.py` : Le script principal.
- `.env` : Fichier contenant les variables d'environnement pour l'email et le serveur IMAP.
- `bordereaux/` : Répertoire où les bordereaux téléchargés et traités seront stockés.

## Personnalisation

- **Transporteurs et règles de rognage** : Vous pouvez ajouter ou modifier les transporteurs et leurs règles de rognage dans le dictionnaire `crop_rules` du script.
- **Sujet des emails** : Modifiez la variable `subject_keyword` pour cibler des emails avec un autre sujet.