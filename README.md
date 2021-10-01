# Carnets OCS

Un script python pour automatiser la construction des carnets OCS.

## usage

```Python
>>> import dossierspot as spot
>>> rap = spot.Report(r'/path/to/depcoms.txt')
>>> rap.process()
>>>
```

## requires

Le script tourne sous arcgis desktop 10.7, python 2.7.

Modules python:

- arcpy
- pandas
- win32com.client

## fichiers inclus

- templates arcgis desktop
- templates excel

## fonctionnement

Le script accepte deux paramètres:

- un fichier texte depcoms.txt qui contient la liste des communes à traiter.
- le nom du dossier racine (conteant le fichier depcoms.txt) qui définit le nom du périmètre.

### exécuter le script sous arcgis desktop

Dans boîte à outils arcgis desktop, pointer sur le fichier depcoms.txt et lancer le script.

Le script lit la liste de codes communaux, effectue un ensemble de définition sur la base source OCS et peuple les templates cartographiques et excel avec cet extrait.

Les templates sont exportées en pdf dans le dossier racine sous une arborescence complète.

### usage

todo...
