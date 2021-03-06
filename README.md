[![Windows](https://svgshare.com/i/ZhY.svg)](https://svgshare.com/i/ZhY.svg)
[![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)](https://www.python.org/)

# Carnets OCS

Un script python pour automatiser la construction des carnets OCS.

Le script est lancé à partir d'une boîte à outils arcgis desktop acceptant un seul contrôle pointant sur un fichier ```depcoms.txt``` qui contient la liste des codes communes à traiter. Le nom du dossier contenant le fichier définit le nom du périmètre dans les pdf produits. Le script peut donc très facilement être lancé en mode batch dans arcgis desktop.

## avertissement

La source de données mos est privée et accessible uniquement localement. La source du script doit être modifiée pour pointer sur une ressource accessible au script.

```Python
MOS_PATH = r'J:\Etudes\laufma\Python26\site-packages\mezcal\data\mos.gdb\mos_urba3_2010_2020'
```

## usage

Console python

```Python
>>> import dossierspot as spot
>>> rap = spot.Report(r'/path/to/depcoms.txt')
>>> rap.process()
>>>
```

Terminal

```powershell
python dossierocs /path/to/depcoms.txt
```



## requires

Le script tourne sous arcgis desktop >10.7, python 2.7.

Modules python:

- arcpy
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
