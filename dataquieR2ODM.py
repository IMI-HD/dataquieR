#!/usr/bin/python3
import argparse
import pandas as pd
import numpy as np
from lxml import etree as ET
from datetime import datetime
import os
import sys
import html
from itertools import zip_longest
import ast
from pathlib import Path
import hashlib

"""
Codelist represents the number for the OID, the list of names which
use the codelist, and the codelist itself in English and German as a dictionary.
"""
class CodeList:
    def __init__(self, number, name, codelist_en, codelist_de):
        """
        :param number: CodeList number (int)
        :param name: A varname that uses this base codelist (str)
        :param codelist_en: A dictionary with the English codelist (dict)
        :param codelist_de: A dictionary with the German codelist (dict)
        """
        self.number = number
        self.names = [name]
        self.codelist_en = codelist_en or {}
        self.codelist_de = codelist_de or {}

    def add_name(self, name):
        """
        Add a name to the list
        :param name: name to be added (str)
        """
        self.names.append(name)


""" 
Extracts the number of the label of the sheet from a missing list.
(Kept for compatibility; currently not used.)
"""
def extract_number(missing_list_name):
    return missing_list_name.split("_")[-1]


""" 
Check if the given codelist already exists.
"""
def check_codelist(codelist_en, codelist_de, name, CodeLists):
    # search for existing codelist with the exact codelist
    for codelist in CodeLists:
        if codelist.codelist_de == codelist_de and codelist.codelist_en == codelist_en:
            # add name to the list of the codelist
            codelist.add_name(name)
            return True
    # the codelist was not available
    return False


""" 
Split the inserts of the cell to have an array with numbers and strings.
"""
def process_codelist(language):
    if pd.notna(language):
        CodeDict = {}
        # Split the string with "|"
        pairs = str(language).split("|")
        for pair in pairs:
            # split the pairs with "=" at the first "="
            # e.g. "1=<=2cm" => 1 = "<= 2cm"
            if "=" in pair:
                key, value = pair.split("=", 1)
                # clear the key from spaces with strip
                CodeDict[key.strip()] = value
            else:
                # Add sentinel key "-999999" to strings without a key
                CodeDict["-999999"] = pair.strip()
        # return the calculated dictionary
        return CodeDict

    # there is no codelist
    return {}


"""
Save all the column names in a dictionary.
"""
def dictionary_column_names(lst):
    dictionary = {}
    column = 0
    for name in lst:
        dictionary[name] = column
        column += 1

    return dictionary


"""
Extract the value from line and catch the type error None.
"""
def extract_from_line(line, dict_get):
    value = None
    try:
        value = line[dict_get]
    except TypeError:
        value = None

    return value


"""
Check the datatypes if they are integers.
"""
def check_datatype(codelist: CodeList):
    # check the keys and their Datatype
    datatype = "integer"
    # codelist German
    if len(codelist.codelist_de) > 0:
        for key in codelist.codelist_de.keys():
            try:
                int(key)
            except ValueError:
                datatype = "string"
                break
    # codelist English
    if datatype == "integer" and len(codelist.codelist_en) > 0:
        for key in codelist.codelist_en.keys():
            try:
                int(key)
            except ValueError:
                datatype = "string"
                break

    return datatype


"""
LEGACY (not called anymore after merge-to-union change)
Create explicit Missing CodeLists (ML.*).

Example:
<CodeList OID="ML.number" Name="MISSING_LIST_TABLE" DataType="integer">
    <CodeListItem CodedValue="CODED_VALUE">
        <Decode>
            <TranslatedText xml:lang="en">CODE_LABEL</TranslatedText>
        </Decode>
    </CodeListItem>
    <Alias Context="CODE_CLASS" Name="CODE_CLASS"/>
</CodeList>
"""
def calculate_missinglists(metadata, df, all_sheets, list_ml, ml):
    # iteration over the sheets from the second sheet on
    for sheet_name, df in all_sheets.items():
        dictionary = dictionary_column_names(list(df.columns))
        if sheet_name in list_ml and pd.notna(sheet_name):
            # add a codelist as missinglist
            codelist_element = ET.SubElement(
                metadata,
                "CodeList",
                OID="ML." + str(ml),
                Name=str(sheet_name),
                DataType="integer",
            )
            for index, _ in df.iterrows():
                line = df.iloc[index]
                code_value = extract_from_line(line, dictionary.get("CODE_VALUE", None))
                code_label = extract_from_line(line, dictionary.get("CODE_LABEL", None))
                # CODE_VALUE
                codelist_item = ET.SubElement(
                    codelist_element, "CodeListItem", CodedValue=str(code_value)
                )
                # CODE_LABEL
                decode = ET.SubElement(codelist_item, "Decode")
                translated_text_en = ET.SubElement(
                    decode,
                    "TranslatedText",
                    attrib={"{http://www.w3.org/XML/1998/namespace}lang": "en"},
                )
                translated_text_en.text = code_label
                # Alias
                for context, number in dictionary.items():
                    if pd.notna(line[number]):
                        ET.SubElement(
                            codelist_item,
                            "Alias",
                            Context=str(context),
                            Name=str(line[number]),
                        )
                        if context == "CODE_VALUE":
                            try:
                                int(line[number])
                            except ValueError:
                                codelist_element.set("DataType", "string")
            ml = ml + 1


def _stable_combo_oid(base_number: int, sheet: str | None) -> tuple[str, str]:
    """
    Build a stable, deterministic OID/Name for a (base, sheet) combination.
    - No missing sheet: OID = CL.<base>
    - With missing sheet: OID = CL.<base>__M_<sha10>
    """
    if not sheet:
        return (f"CL.{base_number}", f"CL.{base_number}")
    h = hashlib.sha256(f"{base_number}||{sheet}".encode("utf-8")).hexdigest()[:10]
    return (f"CL.{base_number}__M_{h}", f"CL.{base_number}__WITH_{sheet}")

################
# Two-phase approach (compute mapping, then emit codelists)
################

def compute_final_ref_map(CodeLists, group, varname_number, all_sheets, missing_map):
    """
    Phase 1 (no writing): compute for each varname the final CodeListOID based on
    (base CodeList.number, missing sheet). Returns:
      - final_ref_map: dict varname -> CodeListOID
      - combos_used:   set of (base.number, sheet_or_None) actually needed
    """
    final_ref_map = {}
    combos_used = set()

    # Map varname -> base CodeList
    varname_to_base = {}
    for cl in CodeLists:
        for n in cl.names:
            varname_to_base[str(n)] = cl

    for _, lines in group.items():
        for row in lines:
            varname = str(row[varname_number])
            base = varname_to_base.get(varname)
            if not base:
                continue
            sheet = str(missing_map.get(varname)) if missing_map.get(varname) else None
            oid, _ = _stable_combo_oid(base.number, sheet)
            final_ref_map[varname] = oid
            combos_used.add((base.number, sheet))
    return final_ref_map, combos_used


def emit_union_codelists(CodeLists, combos_used, metadata, all_sheets, missing_map):
    """
    Phase 2 (writing): emit exactly one CodeList per needed (base.number, sheet) combo.
    - Base codes are emitted first (DE + optional EN decode).
    - If a missing sheet is present, append its rows as CodeListItem with full alias set
      and add Alias Context="ORIGIN_CODELIST" Name="<sheet>".
    - Final DataType is promoted to 'string' if any missing CODE_VALUE is non-integer.
    """
    # Build lookup: base.number -> CodeList object
    base_by_number = {cl.number: cl for cl in CodeLists}

    def _promote_dtype(a: str, b: str) -> str:
        # simple dominance: presence of 'string' yields 'string', else 'integer'
        return "string" if (a == "string" or b == "string") else "integer"

    for base_number, sheet in sorted(combos_used, key=lambda x: (x[0], str(x[1]))):
        base = base_by_number.get(base_number)
        if base is None:
            continue

        union_dtype = check_datatype(base)
        oid, name = _stable_combo_oid(base_number, sheet)
        cl_el = ET.SubElement(
            metadata, "CodeList",
            OID=oid, Name=name, DataType=union_dtype
        )
        used = set()

        # 1) Emit base codes
        if len(base.codelist_de) > 0:
            for k, v in base.codelist_de.items():
                item_el = ET.SubElement(cl_el, "CodeListItem", CodedValue=str(k))
                dec = ET.SubElement(item_el, "Decode")
                t_de = ET.SubElement(
                    dec, "TranslatedText",
                    attrib={"{http://www.w3.org/XML/1998/namespace}lang": "de"}
                )
                t_de.text = v
                if k in base.codelist_en:
                    t_en = ET.SubElement(
                        dec, "TranslatedText",
                        attrib={"{http://www.w3.org/XML/1998/namespace}lang": "en"}
                    )
                    t_en.text = base.codelist_en[k]
                used.add(str(k))
        elif len(base.codelist_en) > 0:
            for k, v in base.codelist_en.items():
                item_el = ET.SubElement(cl_el, "CodeListItem", CodedValue=str(k))
                dec = ET.SubElement(item_el, "Decode")
                t_en = ET.SubElement(
                    dec, "TranslatedText",
                    attrib={"{http://www.w3.org/XML/1998/namespace}lang": "en"}
                )
                t_en.text = v
                used.add(str(k))
        # else: empty base list is allowed

        # 2) Append missing codes if a sheet is specified
        if sheet:
            if sheet in all_sheets:
                mdf = all_sheets[sheet]
                mcols = {c: i for i, c in enumerate(mdf.columns)}
                col_code = mcols.get("CODE_VALUE")
                col_label = mcols.get("CODE_LABEL")

                for _, mrow in mdf.iterrows():
                    code = None if col_code is None else mrow.iloc[col_code]
                    if pd.isna(code):
                        continue
                    code = str(code)
                    # Promote dtype if missing code is not integer
                    try:
                        int(code)
                    except (TypeError, ValueError):
                        union_dtype = _promote_dtype(union_dtype, "string")

                    if code in used:
                        continue  # skip duplicates

                    item_el = ET.SubElement(cl_el, "CodeListItem", CodedValue=code)
                    dec = ET.SubElement(item_el, "Decode")
                    txt = None if col_label is None else mrow.iloc[col_label]
                    t_en = ET.SubElement(
                        dec, "TranslatedText",
                        attrib={"{http://www.w3.org/XML/1998/namespace}lang": "en"}
                    )
                    t_en.text = str(txt) if pd.notna(txt) else "Missing/Reason"

                    # Add all columns as alias, then mark origin sheet
                    for cname, cidx in mcols.items():
                        v = mrow.iloc[cidx]
                        if pd.notna(v):
                            ET.SubElement(item_el, "Alias", Context=str(cname), Name=str(v))
                    ET.SubElement(item_el, "Alias", Context="ORIGIN_CODELIST", Name=str(sheet))
                    used.add(code)

        # Persist final datatype (may have been promoted)
        cl_el.set("DataType", union_dtype)

###########
# Itemdef
###########

"""
Creates the ItemDefs.

Example (kept from your original comment, adapted to English heading):
<ItemDef OID="I.count_id" Name="VARNAMES" DataType="DATA_TYPE">
    <Description>
        <TranslatedText xml:lang="en">NOTE</TranslatedText>
        <TranslatedText xml:lang="de">NOTE_DE</TranslatedText>
    </Description>
    <Question>
        <TranslatedText xml:lang="en">LABEL</TranslatedText>
        <TranslatedText xml:lang="de">LABEL_DE</TranslatedText>
    </Question>
    <CodeListRef CodeListOID="CL.count_cl"/>
    <Alias Context="MISSING_LIST_TABLE" Name="MISSING_LIST_TABLE" />
    <Alias Context="NOTE_TYPE" Name="NOTE_TYPE" />
    <Alias Context="NOTE" Name="NOTE" />
    <Alias Context="NOTE_DE" Name="NOTE_DE" />
    <Alias Context="VALUE_LABELS" Name="VALUE_LABELS" />
    <Alias Context="VALUE_LABELS_DE" Name="VALUE_LABELS_DE" />
    <Alias Context="LONG_LABEL_DE" Name="LONG_LABEL_DE" />
    <Alias Context="LONG_LABEL" Name="LONG_LABEL" />
    <Alias Context="VARIABLE_ORDER" Name="VARIABLE_ORDER" />
    <Alias Context="GROUP_VAR_OBSERVER" Name="GROUP_VAR_OBSERVER" />
    <Alias Context="TIME_VAR" Name="TIME_VAR"/>
    <Alias Context="GROUP_VAR_DEVICE" Name="GROUP_VAR_DEVICE" />
</ItemDef>
"""
def calculate_itemdef(metadata, line, count_id, CodeLists, dictionary, final_ref_map):
    # variables named
    # varname
    varname_number = 0
    try:
        varname_number = dictionary["VARNAMES"]
    except KeyError:
        varname_number = dictionary.get("VAR_NAMES", None)
    varname = extract_from_line(line, varname_number)
    label = extract_from_line(line, dictionary.get("LABEL", None))
    long_label = extract_from_line(line, dictionary.get("LONG_LABEL", None))
    label_de = extract_from_line(line, dictionary.get("LABEL_DE", None))
    long_label_de = extract_from_line(line, dictionary.get("LONG_LABEL_DE", None))
    codelist_en = extract_from_line(line, dictionary.get("VALUE_LABELS", None))
    codelist_de = extract_from_line(line, dictionary.get("VALUE_LABELS_DE", None))
    group_var_observer = extract_from_line(
        line, dictionary.get("GROUP_VAR_OBSERVER", None)
    )
    time_var = extract_from_line(line, dictionary.get("TIME_VAR", None))
    group_var_device = extract_from_line(line, dictionary.get("GROUP_VAR_DEVICE", None))
    variable_order = extract_from_line(line, dictionary.get("VARIABLE_ORDER", None))
    note_type = extract_from_line(line, dictionary.get("NOTE_TYPE", None))
    note = extract_from_line(line, dictionary.get("NOTE", None))
    note_de = extract_from_line(line, dictionary.get("NOTE_DE", None))
    data_type = extract_from_line(line, dictionary.get("DATA_TYPE", None))
    # missing_table_list intentionally not used here (handled via final_ref_map)

    # itemdef
    itemdef = None
    # list of the datatypes
    valid_data_types = {
        "integer",
        "float",
        "double",
        "date",
        "time",
        "datetime",
        "string",
        "boolean",
    }
    # add non valid datatype
    if data_type not in valid_data_types:
        itemdef = ET.SubElement(
            metadata,
            "ItemDef",
            OID="I." + str(count_id),
            Name=str(varname),
            DataType="string",
        )
    else:
        itemdef = ET.SubElement(
            metadata,
            "ItemDef",
            OID="I." + str(count_id),
            Name=str(varname),
            DataType=data_type,
        )

    # description in the item as note and note_de
    if pd.notna(note_de) or pd.notna(note):
        description = ET.SubElement(itemdef, "Description")
        if pd.notna(note_de):
            translatedtext = ET.SubElement(
                description,
                "TranslatedText",
                attrib={"{http://www.w3.org/XML/1998/namespace}lang": "de"},
            )
            translatedtext.text = str(note_de)
        if pd.notna(note):
            translatedtext = ET.SubElement(
                description,
                "TranslatedText",
                attrib={"{http://www.w3.org/XML/1998/namespace}lang": "en"},
            )
            translatedtext.text = str(note)

    # question in the item
    if pd.notna(label) or pd.notna(label_de):
        question = ET.SubElement(itemdef, "Question")
        if pd.notna(label_de):
            translatedtext = ET.SubElement(
                question,
                "TranslatedText",
                attrib={"{http://www.w3.org/XML/1998/namespace}lang": "de"},
            )
            translatedtext.text = str(label_de)
        if pd.notna(label):
            translatedtext = ET.SubElement(
                question,
                "TranslatedText",
                attrib={"{http://www.w3.org/XML/1998/namespace}lang": "en"},
            )
            translatedtext.text = str(label)
    else:
        question = ET.SubElement(itemdef, "Question")
        translatedtext = ET.SubElement(
            question,
            "TranslatedText",
            attrib={"{http://www.w3.org/XML/1998/namespace}lang": "de"},
        )
        translatedtext.text = str(None)
        translatedtext = ET.SubElement(
            question,
            "TranslatedText",
            attrib={"{http://www.w3.org/XML/1998/namespace}lang": "en"},
        )
        translatedtext.text = str(None)

    # CodeListRef: per-(base,missing) final OID
    final_oid = final_ref_map.get(str(varname))
    if final_oid:
        ET.SubElement(
            itemdef, "CodeListRef", CodeListOID=final_oid
        )
    else:
        # Fallback legacy behavior
        for codelist in CodeLists:
            # match the name in the codelist to add the reference
            if varname in codelist.names:
                ET.SubElement(
                    itemdef, "CodeListRef", CodeListOID="CL." + str(codelist.number)
                )
                break

    # Alias (all columns in the source line)
    for context, number in dictionary.items():
        val = extract_from_line(line, number)
        if pd.notna(val):
            ET.SubElement(
                itemdef, "Alias", Context=str(context), Name=str(val)
            )


"""
Creates the ItemGroups

Example:
<ItemGroupDef OID="IG.count_ig" Name="s3" Repeating="No">
    <Description>
        <TranslatedText xml:lang="de">Item Group Nummer key</TranslatedText>
        <TranslatedText xml:lang="en">Item Group Number key</TranslatedText>
    </Description>
    <ItemRef ItemOID="I.count_id" Mandatory="No"/>
    ...
</ItemGroupDef>
"""
def calculate_itemgroups_event(metadata, group):
    count_id = 1
    count_ig = 1
    for key, values in group.items():
        itemgroupdef = ET.SubElement(
            metadata,
            "ItemGroupDef",
            OID="IG." + str(count_ig),
            Name=str(key),
            Repeating="No",
        )
        # Description
        description = ET.SubElement(itemgroupdef, "Description")
        translatedtext = ET.SubElement(
            description,
            "TranslatedText",
            attrib={"{http://www.w3.org/XML/1998/namespace}lang": "de"},
        )
        translatedtext.text = "Item Group " + str(key)
        translatedtext = ET.SubElement(
            description,
            "TranslatedText",
            attrib={"{http://www.w3.org/XML/1998/namespace}lang": "en"},
        )
        translatedtext.text = "Item Group " + str(key)
        # Item references
        for _line in values:
            ET.SubElement(
                itemgroupdef, "ItemRef", ItemOID="I." + str(count_id), Mandatory="No"
            )
            count_id += 1
        count_ig += 1
    return


"""
Creates the ODM

Example:
<ODM xmlns="http://www.cdisc.org/ns/odm/v1.3" xmlns:ns2="http://www.w3.org/2000/09/xmldsig#" FileType="Snapshot" FileOID="Project s" CreationDateTime="2024-07-29T16:16:28.641067" ODMVersion="1.3.2" SourceSystem="OpenEDC">
  <Study OID="s">
    <GlobalVariables>
      <StudyName>Study s</StudyName>
      <StudyDescription>This example study aims at providing an overview of the capabilities of OpenEDC.</StudyDescription>
      <ProtocolName>s---(List of the column names)</ProtocolName>
    </GlobalVariables>
    <MetaDataVersion OID="MDV.1" Name="MetaDataVersion">
      <Protocol>
        <StudyEventRef StudyEventOID="SE.1" Mandatory="No"/>
      </Protocol>
      <StudyEventDef OID="SE.1" Repeating="No">
        <FormRef FormOID="F.1" Mandatory="No"/>
      </StudyEventDef>
      <FormDef OID="F.1" Repeating="No">
        <ItemGroupRef ItemGroupOID="IG.1" Mandatory="No"/>
        ...
      </FormDef>
    </MetaDataVersion>
  </Study>
</ODM>
"""
# start calculating the odm
def calculate_odm(
    df,
    all_sheets,
    file_name,
    varname_groups,
    CodeLists,
    dictionary_names,
    varname_number,
    first_sheet_name,
    missing_map,
):
    # Study name
    name = file_name.split(".")[0]
    # separator
    separator = "---"

    """ Study Events """
    # go through all study events
    for key, group in varname_groups.items():

        """ Root-Element """
        odm = ET.Element(
            "ODM",
            nsmap={
                None: "http://www.cdisc.org/ns/odm/v1.3",
                "ns2": "http://www.w3.org/2000/09/xmldsig#",
            },
            FileType="Snapshot",
            FileOID="Project " + str(name),
            CreationDateTime=datetime.now().isoformat(),
            ODMVersion="1.3.2",
            SourceSystem="OpenEDC",
        )

        """ Study """
        study = ET.SubElement(odm, "Study", OID=name)

        """ Global Variables """
        global_variables = ET.SubElement(study, "GlobalVariables")
        # StudyName, StudyDescription and ProtocolName
        ET.SubElement(global_variables, "StudyName").text = (
            "Study " + name + "_" + str(key)
        )
        ET.SubElement(global_variables, "StudyDescription").text = (
            "This example study aims at providing an overview of the capabilities of OpenEDC."
        )
        # Use file base name and first sheet for traceability, but without column names.
        ET.SubElement(global_variables, "ProtocolName").text = f"{name}---{first_sheet_name}"

        """ Metadata, Study Event, Form, Item Group """
        metadata = ET.SubElement(
            study, "MetaDataVersion", OID="MDV.1", Name="MetaDataVersion"
        )
        protocol = ET.SubElement(metadata, "Protocol")

        # create studyevents and forms
        count_f = 1
        # get all studyevents with formrefs
        ET.SubElement(protocol, "StudyEventRef", StudyEventOID="SE.1", Mandatory="No")
        StudyEvent = ET.SubElement(
            metadata,
            "StudyEventDef",
            OID="SE.1",
            Name=key,
            Repeating="No",
            Type="Unscheduled",
        )
        for _, _ in group.items():
            ET.SubElement(
                StudyEvent, "FormRef", FormOID="F." + str(count_f), Mandatory="No"
            )
            count_f += 1
        # get all formdefs with itemgrouprefs
        count_f = 1
        for key_segment, _ in group.items():
            formdef = ET.SubElement(
                metadata,
                "FormDef",
                OID="F." + str(count_f),
                Name=key_segment,
                Repeating="No",
            )
            ET.SubElement(
                formdef,
                "ItemGroupRef",
                ItemGroupOID="IG." + str(count_f),
                Mandatory="No",
            )
            count_f += 1

        # create itemgroups with refs
        calculate_itemgroups_event(metadata, group)

        """ Phase 1: compute final mapping (no writing) """
        final_ref_map, combos_used = compute_final_ref_map(
            CodeLists, group, varname_number, all_sheets, missing_map
        )

        """ Items (ItemDef*) — MUST appear before CodeList* """
        count_id = 1
        for _, values in group.items():
            for line in values:
                calculate_itemdef(metadata, line, count_id, CodeLists, dictionary_names, final_ref_map)
                count_id += 1

        """ Phase 2: emit CodeLists (CodeList*) after ItemDefs """
        emit_union_codelists(CodeLists, combos_used, metadata, all_sheets, missing_map)

        """ XML """
        # Output Directory
        output_dir = Path("../output")
        output_dir.mkdir(parents=True, exist_ok=True)

        # create the name for the xml
        whole_name = output_dir / f"Study_{name}_{key}.xml"
        # create the xml with indentations
        xml_bytes = ET.tostring(
            odm, encoding="utf-8", xml_declaration=True, pretty_print=True
        )
        with open(whole_name, "wb") as xml_file:
            xml_file.write(xml_bytes)


"""
Sort all lines and columns in a 2D-dictionary along HIERARCHY column.
"""
def sort_new_hierarchy(
    varname_groups, list_keys, study_segment_column, hierarchy_column, hierarchy
):
    # go through all listed keys which have too many items
    for key in list_keys:
        # save the dictionary and delete the previous one
        dictionary = varname_groups[key]
        del varname_groups[key]
        # new sort of the dictionary
        for _, group in dictionary.items():
            for item in group:
                studyevent = ""
                if pd.notna(item[hierarchy_column]):
                    studyevent = item[hierarchy_column]
                    parts = item[hierarchy_column].split("|")
                    # calculate the new key
                    if len(parts) > hierarchy:
                        save = ""
                        count = 0
                        for item_string in parts:
                            save = save + item_string + "_"
                            if count == hierarchy:
                                break
                            count += 1
                        studyevent = save[:-1]
                study_segment = ""
                if pd.notna(item[study_segment_column]):
                    study_segment = item[study_segment_column]

                # 2D dictionary for studyevent and formdef
                if studyevent not in varname_groups:
                    varname_groups[studyevent] = {}
                if study_segment not in varname_groups[studyevent]:
                    varname_groups[studyevent][study_segment] = []

                # add the whole line as a list to the key itemgroup
                varname_groups[studyevent][study_segment].append(item)

    return varname_groups


"""
Sort all lines and columns in a 2D-dictionary along HIERARCHY column (chunking).
"""
def sort_new_hierarchy2(
    varname_groups, list_keys, study_segment_column, hierarchy_column, hierarchy
):
    # go through all listed keys which have too many items
    for key in list_keys:
        # save the dictionary and delete the previous one
        dictionary = varname_groups[key]
        del varname_groups[key]
        # count the item per xml
        count = 0
        count_keys = 0
        # new sort of the dictionary
        for _, group in dictionary.items():
            for item in group:
                if count >= 4500:
                    count = 0
                # if count 0 calculate a new study_segment key
                if count == 0:
                    studyevent = str(key) + "_" + str(count_keys)
                    study_segment = str(key) + "_" + str(count_keys)
                    count_keys += 1

                # 2D dictionary for studyevent and formdef
                if studyevent not in varname_groups:
                    varname_groups[studyevent] = {}
                if study_segment not in varname_groups[studyevent]:
                    varname_groups[studyevent][study_segment] = []

                # add the whole line as a list to the key itemgroup
                varname_groups[studyevent][study_segment].append(item)

                count += 1

    return varname_groups


"""
Sort all lines and columns in a 2D-dictionary.
First dictionary is the character before the dot in VARNAMES (s2.sdlkhre -> s2) => StudyEvent
Second dictionary is based on the entries in the column STUDY_SEGMENT => Form
"""
def sort_all_lines_and_columns(df, first_sheet_name, all_sheets, file_name, force_single_odm):
    # file_name
    name = file_name.split(".")[0]
    missing_map = {}  # varname -> missing_sheet_name

    # Build a dictionary of the column names with their column number
    column_names = list(df.columns)
    dictionary_names = dictionary_column_names(column_names)

    """ Variables """
    # count the codelists, they are unique
    count_cl = 1
    # in this dictionary save all lines in 2D
    varname_groups = {}
    # save all the codelists with important information
    CodeLists = []

    """ Process """
    # go through all rows in the xlsx
    for _, row in df.iterrows():
        """Varname/Study Event (2D Dictionary)"""
        # extract the varname number
        varname_number = 0
        try:
            varname_number = dictionary_names["VARNAMES"]
        except KeyError:
            varname_number = dictionary_names.get("VAR_NAMES", None)
        if varname_number is None:
            varname_number = 0
        # extract the varname
        varname = row.iloc[varname_number]
        # extract form name, e.g. s2
        # hierarchy and studyevent
        save = list(str(row["HIERARCHY"]).split("|"))
        string_save = save[0]
        for i in save:
            string_save = string_save + "_" + str(i)
        studyevent = string_save
        # dce
        if pd.notna(row.get("DCE", None)):
            studyevent = row["DCE"]
        # hierarchy and study_segment
        save = list(str(row["HIERARCHY"]).split("|"))
        string_save = save[0]
        for i in save:
            string_save = string_save + "_" + str(i)
        study_segment = string_save
        # study_segment
        if pd.notna(row.get("STUDY_SEGMENT", None)):
            study_segment = row["STUDY_SEGMENT"]

        # 2D dictionary for studyevent and formdef
        if studyevent not in varname_groups:
            varname_groups[studyevent] = {}
        if study_segment not in varname_groups[studyevent]:
            varname_groups[studyevent][study_segment] = []

        # add the whole line as a list to the key itemgroup
        varname_groups[studyevent][study_segment].append(row.tolist())

        """ Value Labels/Codelist """
        # first go through the process that splits the string into key-value-pairs
        # it returns a dictionary
        english = {}
        german = {}
        try:
            english = process_codelist(
                row.iloc[dictionary_names.get("VALUE_LABELS", None)]
            ) if dictionary_names.get("VALUE_LABELS", None) is not None else {}
        except Exception:
            english = {}
        try:
            german = process_codelist(
                row.iloc[dictionary_names.get("VALUE_LABELS_DE", None)]
            ) if dictionary_names.get("VALUE_LABELS_DE", None) is not None else {}
        except Exception:
            german = {}

        # Codelists
        if len(english) > 0 or len(german) > 0:
            # just add the codelist if there isn't an exact codelist yet
            if not check_codelist(english, german, varname, CodeLists):
                # of course only append existing codelists (not nulls)
                if pd.notna(english) or pd.notna(german):
                    CodeLists.append(CodeList(count_cl, varname, english, german))
                    count_cl += 1
        
        # Missing list name per varname
        idx = dictionary_names.get("MISSING_LIST_TABLE", None)
        missing_table_list_val = row.iloc[idx] if idx is not None else None
        if pd.notna(missing_table_list_val):
            # get varname (already available)
            missing_map[str(varname)] = str(missing_table_list_val)

    if not force_single_odm:  # write in more than one ODM if needed
        break_boolean = False
        # minimum of the hierarchy is 2 because 0 and 1 are SHIP and SHIPx (required)
        h = 2
        while not break_boolean:
            b = False
            # check the number of items and split by hierarchy
            hierarchy = []
            for key, group in varname_groups.items():
                length = 0
                for _, items in group.items():
                    length = length + len(items)
                    if length > 5700:
                        b = True
                        hierarchy.append(key)
                        break
            # new sort by hierarchy
            if hierarchy:
                if h == 2:
                    varname_groups = sort_new_hierarchy(
                        varname_groups,
                        hierarchy,
                        df.columns.get_loc("STUDY_SEGMENT"),
                        df.columns.get_loc("HIERARCHY"),
                        h,
                    )
                if h >= 3:
                    varname_groups = sort_new_hierarchy2(
                        varname_groups,
                        hierarchy,
                        df.columns.get_loc("STUDY_SEGMENT"),
                        df.columns.get_loc("HIERARCHY"),
                        h,
                    )
                h += 1
            # check if the key had changed
            if not b:
                break_boolean = True

    """ For each Study Event create an ODM """
    calculate_odm(
        df,
        all_sheets,
        file_name,
        varname_groups,
        CodeLists,
        dictionary_names,
        varname_number,
        first_sheet_name,
        missing_map,
    )


""" 
Extract sheets and names of the sheets.
"""
# read the files
def odm(file_path, file, force_single_odm):
    # load all sheets
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None)
        # take the first sheet
        first_sheet_name = list(all_sheets.keys())[0]
        first_sheet_df = all_sheets[first_sheet_name]
        # take the other sheets
        remaining_sheets_dict = {
            name: df for name, df in all_sheets.items() if name != first_sheet_name
        }
        # calculate the odm xml
        sort_all_lines_and_columns(
            first_sheet_df, first_sheet_name, remaining_sheets_dict, file, force_single_odm
        )
    except Exception as e:
        print(f"Error while reading the file {file}: {e}")


# read path
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert XLSX → ODM")

    parser.add_argument("file", help="Path to the XLSX file")
    parser.add_argument(
        "--force_single_odm",
        action="store_true",
        help="Write all items in just one ODM (optional flag)"
    )

    args = parser.parse_args()

    file_path = args.file
    force_single_odm = args.force_single_odm

    if len(sys.argv) < 2:
        print("Please add a path to the xlsx file.")
    else:
        # file name
        file_name = os.path.basename(file_path)
        # process odm
        print(file_name)
        odm(file_path, file_name, force_single_odm)
