# dataquieR
The “dataquieR format” is a format for storing metadata for controlling automatic data quality assessments with the R package dataquieR. In this repositroy: Converting XLSX to ODM.

## Infos
Missinglist-Tabellen werden als Codelists hinter allen andern Codelists angehängt.

Aufteilung in StudyEvent, FormDef, etc. aufgrund der Spalte HIERARCHIE (danach DCE, STUDY_SEGMENT)

Maximal 5700 Variablen pro ODM-Datei

## Zuordnung 
VARNAMES/VAR_NAMES: ItemDef (Name)

VALUE_LABELS: Codelists (Mehrfach verwendete VALUE_LABELS werden nur einmal als Codelist verwendet, aber mehrfach verlinkt)

## Starten des Programms
Linux

$ python3 dataquieR2ODM.py /path/to/your/file.xlsx

Windows

$ python dataquieR2ODM.py "C:\path\to\your\file.xlsx"

Start with flag force_single_odm:

$ python3 dataquieR2ODM.py /Users/.../x0.xlsx --force_single_odm

=> You want to have all items of the xlsx in just one ODM. No mather how big it will be.

## Generierte ODM-Dateien
Diese landen im neuen Ordner "output" des gleichen Verzeichnisses wie das Python-Skript.

## Warnings / Future Warnings
Bitte ignorieren, das Skript funktioniert trotzdem.
