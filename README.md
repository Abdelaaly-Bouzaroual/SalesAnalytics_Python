# SalesAnalytics_Python

# 📊 Projet d’Analyse des Ventes – Data Science avec Python

## 🎯 Objectif
Ce projet a pour but de :
- Nettoyer et transformer un jeu de données de ventes issues d’un fichier Excel (`sales.xlsx`)
- Calculer des indicateurs de performance clés (KPIs) tels que le chiffre d’affaires, le bénéfice et la marge bénéficiaire
- Comparer les performances de vente par produit, client, canal et ville d’une année sur l’autre
- Créer des visualisations claires et interactives des résultats

## 🧾 Description des Données
Le fichier Excel contient 4 feuilles :
- **Sales Orders** : détails des commandes (quantité, prix unitaire, coût, date, etc.)
- **Customers** : données sur les clients
- **Regions** : informations géographiques (ville, latitude, longitude)
- **Products** : nom des produits

## 🔧 Étapes du Projet
1. **Chargement et nettoyage** des données (`pandas`)
2. **Fusion** des différentes sources via les index
3. **Création d’une table de dates** pour permettre l’analyse temporelle
4. **Calcul des mesures** : ventes totales, profits, marges, comparaisons annuelles
5. **Création de visualisations** via `matplotlib` et `seaborn`
6. **Export** d’un rapport DOCX et d’une présentation PowerPoint avec extraits de code

## 📈 KPIs calculés
- Total Sales
- Total Profit
- Profit Margin (%)
- Total Orders
- Comparaison des ventes/profits avec l’année précédente
- Top/Bottom 5 clients, produits, villes

## 📊 Visualisations Générées
- Ventes par produit (année en cours vs précédente)
- Ventes mensuelles par région
- Top 5 des villes les plus lucratives
- Profits par canal de vente
- Top 5 / Last 5 clients

## 📦 Technologies utilisées
- Python 3.x
- `pandas`, `numpy` pour les calculs
- `matplotlib`, `seaborn` pour les graphiques
- `python-pptx`, `python-docx` pour les livrables
