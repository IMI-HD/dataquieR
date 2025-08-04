# dataquieR
Converting XLSX into ODM

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

## Generierte ODM-Dateien
Diese landen im neuen Ordner "output" des gleichen Verzeichnisses wie das Python-Skript.

## Warnings / Future Warnings
Bitte ignorieren, das Skript funktioniert trotzdem.
