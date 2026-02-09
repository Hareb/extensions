# Comparateur de Répertoire Téléphonique

Interface graphique PowerShell pour comparer les extensions téléphoniques des utilisateurs entre Active Directory et un fichier Excel.

## Fonctionnalités

### 🔍 Détection Complète des Changements
- **Nouveaux employés** : Détecte les utilisateurs présents dans AD mais absents du fichier Excel
- **Employés partis** : Identifie les utilisateurs dans le fichier Excel mais plus dans AD
- **Modifications** : Détecte automatiquement les changements de :
  - Extensions téléphoniques
  - Adresses
  - Villes
  - Succursales
  - Codes postaux
  - Emails

### 📊 Interface Graphique Améliorée
- **3 onglets de résultats** :
  - Nouveaux employés (fond vert)
  - Employés partis (fond rouge)
  - **Modifications** (fond jaune) - Affiche côte à côte les anciennes et nouvelles valeurs

### 🔎 Filtres de Recherche en Temps Réel
- Filtrage dynamique dans les panneaux AD et Fichier
- Recherche par :
  - Nom
  - Prénom
  - Succursale
  - Extension

### 💾 Export des Résultats
- Export CSV avec encodage UTF-8 (compatible Excel)
- Inclut tous les types de changements :
  - Nouveaux
  - Partis
  - Modifications avec détails des changements
- Nom de fichier automatique avec horodatage

### ⚡ Performance Optimisée
- **Cache intelligent des données AD** :
  - Valide pendant 5 minutes
  - Évite les rechargements inutiles
  - Indication visuelle "(depuis cache)"
- **Barre de progression** pour toutes les opérations longues
- **Normalisation des extensions** : Compare correctement même avec espaces/tirets différents

### 🛠️ Améliorations Techniques
- Code refactorisé avec fonction helper `New-CustomDataGrid`
- Comparaison insensible à la casse des SamAccountName
- Gestion d'erreurs robuste
- Interface responsive et moderne

## Utilisation

1. Lancer le script PowerShell
2. Cliquer sur "CHARGER DEPUIS AD" pour récupérer les données Active Directory
3. Cliquer sur "CHARGER FICHIER EXCEL" pour importer le fichier de référence
4. La comparaison s'effectue automatiquement
5. Consulter les résultats dans les 3 onglets
6. Utiliser les filtres pour rechercher des utilisateurs spécifiques
7. Cliquer sur "EXPORTER LES RESULTATS (CSV)" pour sauvegarder

## Configuration

Le script utilise les paramètres suivants (modifiables dans le code) :
- `$OUPath` : Chemin de l'OU dans Active Directory
- `$locationMapping` : Mapping des codes postaux vers les succursales
- `$cacheValidityMinutes` : Durée de validité du cache AD (5 minutes par défaut)

## Prérequis

- Windows PowerShell 5.1+
- Module Active Directory
- Microsoft Excel (pour l'import de fichiers Excel)
- Droits de lecture sur l'OU Active Directory configurée

## Scripts Disponibles

### 1. Extensions GUI v2.ps1
Interface graphique pour comparer les extensions entre AD et fichier Excel.

**Utilisation:**
```powershell
.\Extensions GUI v2.ps1
```

### 2. Generate-SuccursaleReport.ps1 ⭐ NOUVEAU
Génère un rapport HTML professionnel classant les employés par succursale.

**Fonctionnalités:**
- 📊 Classification intelligente par succursale (14 succursales + 7 Espaces Plombérium)
- 🎯 Matching tolérant basé sur les adresses AD
- 🎨 Rapport HTML avec design professionnel (gradients, badges, tables interactives)
- 📋 Table des matières cliquable
- 📈 Statistiques détaillées
- 🏢 Distinction visuelle Succursales vs Espaces Plombérium
- ❓ Section pour employés non classés

**Utilisation:**
```powershell
.\Generate-SuccursaleReport.ps1
```

Le script génère un fichier `Rapport_Succursales_YYYYMMDD_HHmmss.html` avec:
- En-tête avec gradient bleu/violet
- Cartes statistiques (Total employés, Succursales, Espaces, Non classés)
- Sections par succursale avec headers colorés
- Tableaux d'employés triés par nom
- Design responsive et imprimable

**Algorithme de Classification:**
1. Extrait les mots-clés des adresses (numéros de rue, noms, villes)
2. Compare avec les adresses de référence des succursales
3. Score basé sur les correspondances (fuzzy matching)
4. Attribution à la succursale avec le meilleur score (seuil: 10+)

## Fichiers Requis

- `Succursales addresses.xlsx` : Correspondance adresses ↔ succursales (21 lignes)
  - Colonnes: Nom succursale, Adresse, Numéro succursale
  - 14 succursales (#1-9, #20, #40, #42-44)
  - 7 Espaces Plombérium (#21, #23-27, #50)

## Améliorations Version 2

### Nouvelles fonctionnalités (février 2026)
✅ Détection des changements d'extension
✅ Onglet "Modifications" avec vue détaillée ancien/nouveau
✅ Export CSV complet des résultats
✅ Filtres de recherche en temps réel
✅ Barre de progression pour les opérations longues
✅ Cache intelligent des données AD
✅ Refactorisation du code (évite la duplication)
✅ Normalisation des extensions pour comparaison précise
✅ **Générateur de rapport par succursale** (HTML professionnel)
✅ **Matching intelligent d'adresses** (classification automatique)
