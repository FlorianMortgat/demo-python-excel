Ce dossier contient une démonstration de manipulation en python de fichiers Excel :

* nettoyage-exemple.py : le script de démonstration

* exemple-source.xlsx : le fichier "sale" à nettoyer

* bijnum.py : ma bibliothèque pour convertir des nombres en noms de colonne
  Excel et vice versa (ex : 'XA' = colonne 625 en partant de 1). Elle n’est pas
  utilisée par mon script de démo, mais elle peut servir.
  Exemple :
  >>> from bijnum import AZ
  >>> AZ.aaa2n('XJ')
  634
  >>> AZ.n2aaa(5)
  'E'

* TSV.py : ma bibliothèque pour travailler avec des données CSV du type
  tab-separated telles que celles contenues par le presse-papier quand on copie
  des cellules depuis Excel. Cette bibliothèque n’est pas utilisée par mon
  script de démo, mais elle peut parfois servir (ça va parfois plus vite de
  travailler sur des dumps du presse-papier collés dans un bloc-notes que de
  créer un script openpyxl complet).  Exemple :
  >>> from TSV import parseTSV, exportTSV
  >>> donnees_test = [['A1', 'B1 avec\nfin de ligne'], ['A2', 'B2'], ['A3', 'B3']]
  >>> print(exportTSV(donnees_test))
  "A1"	"B1 avec
  fin de ligne"
  "A2"	"B2"
  "A3"	"B3"
  >>> parseTSV('''
  ... "A1"	"B1 avec
  ... fin de ligne"
  ... "A2"	"B2"
  ... "A3"	"B3"''')
  [['A1', 'B1 avec\nfin de ligne'], ['A2', 'B2'], ['A3', 'B3']]

Version de python :
* ces scripts sont prévus pour fonctionner avec python 3.*.
* avec python 2.7, ça fonctionne assez bien, MAIS : quand on utilise python 2
  avec openpyxl, il vaut mieux ne pas mettre de caractères non ASCII dans les
  identifiants Excel (noms des feuilles, noms des plages de cellules) car
  openpyxl pour python 2 a quelques bugs non corrigés (très faciles à corriger
  dans le code source, d’ailleurs, c’est ce que j’avais fait pour
  <nom-du-projet> car j’avais commencé avec python 2).

Version d’openpyxl :
* J’ai utilisé la version 2.5.12 d’openpyxl, que j’ai légèrement modifiée pour
  résoudre des bugs mineurs ; mes scripts peuvent ou non fonctionner avec des
  versions plus récentes.

Note sur les formules avec openpyxl :
* l’évaluation des formules dans Excel est faite par Excel. Openpyxl n’inclut
  pas de module d’évaluation de formules (autrement dit, openpyxl est incapable
  de calculer le contenu d’une cellule ayant une formule). Openpyxl peut mettre
  des formules dans les cellules, mais ne peut pas les "appliquer", ni même
  vérifier qu’elles sont valides.

Contact :
    Florian Mortgat
