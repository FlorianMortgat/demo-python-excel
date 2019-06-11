#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Permet de décoder une chaîne de caractères issue du presse-papier (lui-même issu
d’un tableur Excel ou autre) en une liste de listes représentant ces données
tabulaires.

Permet d’encoder une liste de données tabulaires en une chaîne interprétable par
Excel (ou autre tableur).

POURQUOI CE FICHIER ?

Un décodage naïf se contenterait de faire :
    parseTSV = lambda tsv: [ '\t'.split(row) for row in [tsv.split('\n')] ]

Ça fonctionne très bien du moment qu’il n’y a aucun retour à la ligne à
l’intérieur d’une cellule. Le module TSV permet de prendre en compte la façon
dont ces particularités (les guillemets doubles et les retours ligne) sont
encodées dans le format "presse-papier" reconnu par Excel et LibreOffice.

Dans ce format (très peu documenté à ma connaissance) :
    * les cellules contenant des fins de ligne sont mises entre double quotes 
    * les doubles quotes sont échappées en double-double-quotes ("").

Note : mon implémentation parseTSV(tsv) n’est pas parfaite non plus, mais elle
est meilleure :
    * elle produira des erreurs si les cellules contiennent des
      double-double-quotes
    * elle produira des erreurs si les cellules contiennent les chaînes de
      caractères arbitraires utilisées ici comme tokens : 
      "[DOUBLEDOUBLEQUOTETOKEN]" et "[EOLTOKEN]".
"""
import re

RE_EMPTY_ROW = re.compile(r'^\s*$')
DBL2_TOKEN = '[DOUBLEDOUBLEQUOTETOKEN]'
EOL_TOKEN = '[EOLTOKEN]'

def parseTSV(tsv, sep='\t'):
    RE_INSIDE_DOUBLE_QUOTES = re.compile('"([^%s]*?)"'%sep)
    def cb(m):
        return m.group(1).replace('\n', EOL_TOKEN)
    tsv = tsv.replace('\r\n', '\n').replace('""', DBL2_TOKEN)
    tsv = RE_INSIDE_DOUBLE_QUOTES.sub(cb, tsv)
    tsv = tsv.replace(DBL2_TOKEN, '"')
    ret = [ row.replace(EOL_TOKEN, '\n').split(sep) for row in tsv.split('\n') if not RE_EMPTY_ROW.match(row) ]
    return ret

def exportTSV(rows_of_cells, sep='\t'):
    ret = []
    for row in rows_of_cells:
        ret.append(sep.join(('"%s"'%cell.replace('"', '""') for cell in row)))
    return '\n'.join(ret)
