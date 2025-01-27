# Gestionnaire de fichier Office

Ce projet est un script PowerShell conçu pour gérer et convertir des anciens fichiers Office (à savoir .doc, .xls et .ppt) vers des formats modernes (à savoir .docx, .xlsx et .pptx). Il inclut des fonctionnalités pour copier les fichiers, les convertir, générer des rapports d'inventaire, et nettoyer les dossiers.

---

## Fonctionnalités principales

### 1. **Journal d'événements**
- Enregistre toutes les actions effectuées par le programme.
- Affiche les étapes importantes et les erreurs potentielles dans une zone de logs redimensionnable.

### 2. **Copie de fichiers**
- Recherche et copie tous les fichiers avec les extensions .doc, .xls et .ppt depuis un dossier source (y compris ses sous-dossiers) vers un dossier de destination.

### 3. **Rapport E1**
- Génère un fichier Excel (à savoir `Liste_fichiers_e1.xlsx`) listant les chemins complets des fichiers copiés.

### 4. **Conversion des fichiers**
- Convertit les fichiers copiés vers leurs formats modernes associés (à savoir .docx, .xlsx, .pptx).

### 5. **Rapport E2**
- Génère un fichier Excel (à savoir `Liste_fichiers_e2.xlsx`) contenant :
  - **Nm_Orig** : chemins complets des fichiers convertis.
  - **Nm_Tmp** : noms des fichiers convertis sans leur dernière lettre (utile pour certaines compatibilités).

### 6. **Nettoyage des fichiers originaux**
- Permet de supprimer les fichiers originaux (à savoir .doc, .xls, .ppt) après leur conversion.

---

## Prérequis

- **Système d'exploitation** : Windows avec PowerShell 5.0 ou version ultérieure (pré-installé sur Windows 10).
- **Suite Office** : Version moderne de Word, Excel et PowerPoint pour la conversion des fichiers.

---

## Utilisation

1. **Choix des dossiers**
   - Utiliser les boutons "Parcourir" pour sélectionner :
     - Le dossier source à scanner.
     - Le dossier de destination pour copier les fichiers et générer les rapports.

2. **Options supplémentaires**
   - **Conversion des fichiers** : Activer la case à cocher pour lancer la conversion et générer le rapport E2.
   - **Nettoyage des fichiers originaux** : Disponible uniquement après la conversion.

3. **Rapports générés**
   - `Liste_fichiers_e1.xlsx` : Liste des fichiers originaux copiés.
   - `Liste_fichiers_e2.xlsx` : Liste des fichiers convertis avec leurs nouveaux noms et chemins.

---

## Structure des rapports

### **Liste_fichiers_e1.xlsx**
- Contient une colonne **FilePath** listant les chemins complets et noms des fichiers originaux.

### **Liste_fichiers_e2.xlsx**
- Contient deux colonnes :
  - **Nm_Orig** : Chemins complets des fichiers convertis.
  - **Nm_Tmp** : Noms des fichiers convertis sans leur dernière lettre.

---

## Notes importantes

- Veillez à avoir des droits suffisants sur les dossiers pour éviter les erreurs d'accès.
- Les logs affichent en temps réel les étapes importantes et permettent de débugger facilement.

---

## Auteur

Ce projet a été réalisé par **Thomas Heriaud** pour le compte du groupe **Mayr-Melnhof**.

---

## Licence

Ce projet est distribué sous une licence libre. Vous êtes libre de le modifier et de le redistribuer en mentionnant l'auteur original.

