# Créer une version Windows (.exe)

Objectif : créer un seul fichier **ElternTermine.exe** à transmettre aux collègues. Ils n’ont rien à installer : ils double-cliquent, se connectent une fois à Microsoft avec le code, puis utilisent l’application.

## Étapes

1. Installer Python sur le PC Windows si nécessaire.
2. Copier le dossier **ElternTermineWeb_v81** sur le PC Windows.
3. Double-cliquer sur **build_windows.bat**.
4. Le fichier généré se trouve dans **dist\ElternTermine.exe**.

Commande utilisée :

```powershell
pyinstaller --onefile --name ElternTermine --add-data "graph.html;." --add-data "index.html;." --add-data "schullogo.png;." --hidden-import certifi server.py
```

## Distribution

- Transmettre uniquement **ElternTermine.exe** si une version autonome est souhaitée.
- Au premier lancement, Windows peut afficher un avertissement de sécurité. Dans ce cas : **Informations complémentaires** → **Exécuter quand même**.
- Chaque collègue se connecte avec son propre compte Microsoft et envoie depuis sa propre boîte mail.
