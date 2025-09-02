# dataquieR
The “dataquieR format” is a format for storing metadata for controlling automatic data quality assessments with the R package dataquieR. In this repositroy: Converting XLSX to ODM.

## Infos
Missinglist-Tables are atteched as Codelists after all other Codelists.

Destributed in StudyEvent, FormDef, etc. because of the column HIERARCHIE (danach DCE, STUDY_SEGMENT)

Maximum of 5700 Variables of each ODM-file

## Zuordnung 
VARNAMES/VAR_NAMES: ItemDef (Name)

VALUE_LABELS: Codelists (Multiple uses of VALUE_LABELS are just saved once as Codelist, but linked each time)

## Starten des Programms
Linux

$ python3 dataquieR2ODM.py /path/to/your/file.xlsx

Windows

$ python dataquieR2ODM.py "C:\path\to\your\file.xlsx"

Start with flag force_single_odm:

$ python3 dataquieR2ODM.py /Users/.../x0.xlsx --force_single_odm

=> You want to have all items of the xlsx in just one ODM. No mather how big it will be.

## Output: ODM-Files
The output is placed in a new folder "output" in this path. 

## Warnings / Future Warnings
Please ignore, the script is working.