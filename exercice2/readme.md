# Rapport d'extraction – Archives Départementales de l'Aisne   
Temps: 5hr

## 1. Objectif

Extraire l'ensemble des registres paroissiaux et d'état civil visibles pour la commune
d'**Abbécourt (Aisne)** depuis le portail des Archives Départementales de l'Aisne, et
livrer un fichier structuré (CSV + Excel) contenant :

| Champ attendu           | Colonne livrée           |
|-------------------------|--------------------------|
| Commune                 | `commune`                |
| Type d'acte             | `type_acte` + `type_acte_libelle` |
| Années couvertes        | `annee_debut` / `annee_fin` |
| Cote / référence        | `cote`                   |
| Lien vers l'image       | `lien_image` (ARK direct) + `lien_viewer` |

---

## 2. Démarche et outils utilisés

### 2.1 Analyse préliminaire

L'URL fournie est une page de résultats d'une recherche sur le portail
**archives.aisne.fr**, un système d'archives numériques basé sur un CMS propriétaire.
Le filtre appliqué dans l'URL est :
- Commune : Abbécourt (Aisne)
- Depuis 1860 (`RECH_annee_debut=1860`)
- Limite : 50 résultats par page

**Outils utilisés :**
- **Python 3.12.7** (Anaconda)
- **requests 2.32.3** : HTTP client
- **BeautifulSoup4 4.12.3** + **lxml 5.2.1** : parsing HTML
- **pandas 2.2.2** : structuration des données
- **openpyxl 3.1.5** : export Excel avec mise en forme
- **hashlib** (stdlib) : SHA-256 pour le solveur anti-bot

---

## Déroulement de l'extraction

L'extraction est réalisée en plusieurs étapes afin de garantir la récupération complète et fiable des données.

### 1. Initialisation

Le script commence par définir l'URL de recherche, les paramètres de connexion ainsi que les en-têtes HTTP (User-Agent, langue, référent, etc.). Ces informations permettent d'envoyer des requêtes similaires à celles d'un navigateur web.

---

### 2. Récupération de la page

Le script tente ensuite de récupérer la page contenant les registres à l'aide de la bibliothèque **requests**.

Lors de cette étape, il vérifie automatiquement si le site renvoie directement la page des résultats ou une page de protection anti-bot. Si un challenge Anubis est détecté, le script le résout avant de demander à nouveau la page afin d'obtenir le contenu réel.

Une solution de secours avec **Playwright** est également prévue si le contenu devait être généré entièrement en JavaScript.

---

### 3. Analyse du contenu HTML

Une fois la page récupérée, elle est analysée avec **BeautifulSoup**.

Le script vérifie d'abord que le HTML contient bien les données attendues (commune, registres, cotes, etc.). Si le contenu est valide, il lance l'extraction des informations.

---

### 4. Extraction des registres

Le parseur parcourt chaque registre présent dans la page et extrait automatiquement les informations demandées :

- la commune ;
- la cote du registre ;
- les années couvertes ;
- le type d'acte ;
- la collection archivistique ;
- le lien vers l'image du registre.

Lorsque certaines informations sont regroupées dans une seule cellule HTML, le script utilise des expressions régulières afin d'isoler chaque champ.

---

### 5. Nettoyage et validation

Une fois l'ensemble des registres extraits, les données sont nettoyées avant leur export.

Cette étape comprend notamment :

- la suppression des doublons ;
- la normalisation des types d'actes ;
- la conversion des années dans un format homogène ;
- la transformation des liens relatifs en liens absolus.

Ces traitements garantissent un jeu de données cohérent et directement exploitable.

---

### 6. Export des résultats

Les données sont ensuite converties en **DataFrame Pandas** puis exportées dans deux formats :

- **CSV**, pour une utilisation dans d'autres applications ;
- **Excel**, avec une mise en forme automatique (en-têtes, filtres, largeurs de colonnes et liens cliquables).

---

### 7. Journalisation

Tout au long de l'exécution, le script enregistre les principales étapes dans un fichier de log.

Ce journal facilite le suivi de l'extraction et permet d'identifier rapidement un éventuel problème lors de l'exécution du script.
