# Bot Discord / Assistant de bureau Windows

Ce projet est un bot Discord développé en Python qui sert d’assistant local sur un ordinateur Windows. Il permet d’interagir avec un salon Discord depuis une fenêtre Tkinter locale, de recevoir des images, d’ouvrir des fenêtres de chat, de capturer le bureau et d’envoyer des captures dans Discord.

Il a été conçu pour fonctionner en environnement Windows avec des fonctionnalités spécifiques au système d’exploitation (Win32, capture d’écran, ouverture de fichiers Excel, etc.).

## Fonctionnalités

- Bot Discord avec commandes personnalisées
- Ouverture d’une fenêtre locale de discussion par salon Discord
- Fermeture distante de la fenêtre de chat
- Relay des messages texte dans une fenêtre Tkinter
- Relay des images jointes dans la fenêtre locale
- Envoi de messages depuis la fenêtre locale vers Discord
- Envoi d’images depuis l’ordinateur vers Discord
- Capture du bureau Windows
- Envoi de la capture dans un salon Discord
- Son d’ouverture et de notification avec pygame
- Ouverture automatique d’un fichier Excel au démarrage
- Message de démarrage dans un salon configuré
- Interface visuelle “Matrix / hacker” avec animation

## Prérequis

Avant de lancer le projet, assure-toi d’avoir :

- Python 3.9+
- Windows 10 ou Windows 11
- Un bot Discord créé sur le portail développeur Discord
- Un serveur Discord avec les permissions nécessaires
- Les dépendances Python suivantes installées

## Dépendances

Installe les bibliothèques suivantes :

```bash
pip install discord.py python-dotenv pillow pygame pywin32
