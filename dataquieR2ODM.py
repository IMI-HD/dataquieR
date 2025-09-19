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

"""
Codelist represents the number for the OID, the list of names which
use the codelist, and the codelist itself in english and german as a dictionary
"""
class CodeList:
    def __init__(self, number, name, codelist_en, codelist_de):
        """
        :param number: CodeList number (int)
        :param names: A list of names (list of str)
        :param codelist_en: A dictionary with the english codelist (dict)
        :param codelist_de: A dictionary with the german codelist (dict)
        """
        self.number = number
        self.names = [name]
        self.codelist_en = codelist_en
        self.codelist_de = codelist_de

    def add_name(self, name):
        """
        Add a name to the list
        :param name: name to be added (str)
        """
        self.names.append(name)


""" 
Ectracts the number of the label of the sheet from a missinglist 
"""
def extract_number(missing_list_name):
    return missing_list_name.split("_")[-1]


""" 
Check if the given codelist already exists 
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
Split the inserts of the cell to have an array with numbers and strings 
"""
def process_codelist(language):
    if pd.notna(language):
        CodeDict = {}
        # Split the string with "|"
        pairs = language.split("|")
        for pair in pairs:
            # split the pairs with "=" at the first "="
            if "=" in pair:
                # with pair.split('=', 1) catch the values with "=" inside
                # e.g. "1=<=2cm" => 1 = "<= 2cm"
                key, value = pair.split("=", 1)
                # clear the key from spaces with strip
                CodeDict[key.strip()] = value
            else:
                # Add no key "-999999" to this string because there is just a string and no key
                # Example: BKK Stadtwerke Hannover AG in s.xlsx and t.xlsx
                CodeDict["-999999"] = pair.strip()
        # return the calculated dictionary
        return CodeDict

    # there is no codelist
    return {}


"""
Save all the column names in a dictionary
"""
def dictionary_column_names(list):
    dictionary = {}
    column = 0
    for name in list:
        dictionary[name] = column
        column += 1

    return dictionary


"""
Extract the value from line and catch the type error None
"""
def extract_from_line(line, dict_get):
    value = None
    try:
        value = line[dict_get]
    except TypeError:
        value = None

    return value


"""
Check the datatypes if they are integers
"""
def check_datatype(codelist):
    # check the keys and their Datatype
    datatype = "integer"
    # codelist german
    if len(codelist.codelist_de) > 0:
        for key in codelist.codelist_de.keys():
            try:
                int(key)
            except ValueError:
                datatype = "string"
    # codelist english
    if len(codelist.codelist_en) > 0:
        for key in codelist.codelist_en.keys():
            try:
                int(key)
            except ValueError:
                datatype = "string"

    return datatype


"""
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


"""
Creates the CodeLists for the ItemDefs
Beispiel:
<CodeList OID="CL.count_cl" Name="VARNAMES" DataType="integer">
	<CodeListItem CodedValue="0">
		<Decode>
			<TranslatedText xml:lang="en">VALUE_LABELS</TranslatedText>
			<TranslatedText xml:lang="de">VALUE_LABELS_DE</TranslatedText>
		</Decode>
	</CodeListItem>
	<CodeListItem CodedValue="1">
...
	<CodeListItem CodedValue="2">
</CodeList>
"""
# CodedValue = key
# Decode = value (de/en)
def calculate_codelists(CodeLists, group, varname_number, metadata):
    already_saved = []
    for codelist in CodeLists:
        for _, lines in group.items():
            for line in lines:
                if (
                    line[varname_number] in codelist.names
                    and codelist.number not in already_saved
                ):
                    already_saved.append(codelist.number)
                    datatype = check_datatype(codelist)
                    # codelist
                    codelist_element = ET.SubElement(
                        metadata,
                        "CodeList",
                        OID="CL." + str(codelist.number),
                        Name="CL." + str(codelist.number),
                        DataType=datatype,
                    )
                    # if there is a codelist in german
                    if len(codelist.codelist_de) > 0:
                        for key, value in codelist.codelist_de.items():
                            # codelist items
                            codelist_item = ET.SubElement(
                                codelist_element, "CodeListItem", CodedValue=key
                            )
                            decode = ET.SubElement(codelist_item, "Decode")
                            translated_text_de = ET.SubElement(
                                decode,
                                "TranslatedText",
                                attrib={
                                    "{http://www.w3.org/XML/1998/namespace}lang": "de"
                                },
                            )
                            translated_text_de.text = value
                            # add the english translation
                            if key in codelist.codelist_en:
                                translated_text_en = ET.SubElement(
                                    decode,
                                    "TranslatedText",
                                    attrib={
                                        "{http://www.w3.org/XML/1998/namespace}lang": "en"
                                    },
                                )
                                translated_text_en.text = codelist.codelist_en[key]
                    # if there is only the english translation available
                    elif len(codelist.codelist_en) > 0:
                        for key, value in codelist.codelist_en.items():
                            if key not in codelist.codelist_de:
                                # codelist items
                                codelist_item = ET.SubElement(
                                    codelist_element, "CodeListItem", CodedValue=key
                                )
                                decode = ET.SubElement(codelist_item, "Decode")
                                translated_text_en = ET.SubElement(
                                    decode,
                                    "TranslatedText",
                                    attrib={
                                        "{http://www.w3.org/XML/1998/namespace}lang": "en"
                                    },
                                )
                                translated_text_en.text = value
                    # no codelist available
                    else:
                        continue

                    # Incomment if you want to print all names who use one codelist as Alias in the codelist in the xml
                    # TODO: add count first for this
                    """
                    for varname in codelist.names:
                        ET.SubElement(codelist_element, "Alias", Context="Name " + str(count), Name=str(varname))
                        count += 1
                    """


"""
Creates the ItemDefs
Beispiel:
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
    <Alias Context="MISSING_LIST_TABLE" Name="MISSING_LIST_TABLE" </Alias>
    <Alias Context="NOTE_TYPE" Name="NOTE_TYPE" </Alias>
    <Alias Context="NOTE" Name="NOTE" </Alias>
    <Alias Context="NOTE_DE" Name="NOTE_DE" </Alias>
    <Alias Context="VALUE_LABELS" Name="VALUE_LABELS" </Alias>
    <Alias Context="VALUE_LABELS_DE" Name="VALUE_LABELS_DE" </Alias>
    <Alias Context="LONG_LABEL_DE" Name="LONG_LABEL_DE" </Alias>
    <Alias Context="LONG_LABEL" Name="LONG_LABEL" </Alias>
    <Alias Context="VARIABLE_ORDER" Name="VARIABLE_ORDER" </Alias>
    <Alias Context="GROUP_VAR_OBSERVER" Name="GROUP_VAR_OBSERVER" </Alias>
    <Alias Context="TIME_VAR" Name="TIME_VAR"</Alias>
    <Alias Context="GROUP_VAR_DEVICE" Name="GROUP_VAR_DEVICE" </Alias>
</ItemDef>
"""
def calculate_itemdef(metadata, line, count_id, CodeLists, dictionary):
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
    missing_table_list = extract_from_line(
        line, dictionary.get("MISSING_LIST_TABLE", None)
    )

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

    # code list reference
    for codelist in CodeLists:
        # math the name in the codelist to add the reference
        if varname in codelist.names:
            ET.SubElement(
                itemdef, "CodeListRef", CodeListOID="CL." + str(codelist.number)
            )
            break
    # Alias
    for context, number in dictionary.items():
        if pd.notna(line[number]):
            ET.SubElement(
                itemdef, "Alias", Context=str(context), Name=str(line[number])
            )


"""
Creates the ItemGroups
Beispiel:
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
        for line in values:
            ET.SubElement(
                itemgroupdef, "ItemRef", ItemOID="I." + str(count_id), Mandatory="No"
            )
            count_id += 1
        count_ig += 1
    return


"""
Creates the ODM
Beispiel:
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
    string_column_names,
    CodeLists,
    dictionary_names,
    varname_number,
    first_sheet_name,
):
    # Study name
    name = file_name.split(".")[0]
    # seperator
    separator = "---"

    """ Study Events """
    # go through all study events
    for key, group in varname_groups.items():
        # count the missinglists
        ml = 1
        # count the codelists
        cl = 0

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
        ET.SubElement(global_variables, "ProtocolName").text = (
            name + separator + first_sheet_name + string_column_names
        )  # save the name of the first sheet and the column names

        """ Metadata, Study Event, Formular, Item Group """
        metadata = ET.SubElement(
            study, "MetaDataVersion", OID="MDV.1", Name="MetaDataVersion"
        )
        protocol = ET.SubElement(metadata, "Protocol")

        # create studyevents and formulars
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

        """ Items """
        # calculate itemdefs
        count_id = 1
        list_ml = []
        for _, values in group.items():
            for line in values:
                # calculate the itemdefs with codelistrefs
                calculate_itemdef(metadata, line, count_id, CodeLists, dictionary_names)
                # add the needed missing list to the array
                missing_table_list = extract_from_line(
                    line, dictionary_names.get("MISSING_LIST_TABLE", None)
                )
                if missing_table_list not in list_ml:
                    list_ml.append(missing_table_list)
                count_id += 1

        """ Codelists """
        # calculate codelists
        calculate_codelists(CodeLists, group, varname_number, metadata)

        # calculate missinglists
        calculate_missinglists(metadata, df, all_sheets, list_ml, ml)

        """ XML """
        # Output Directory
        output_dir = Path("output")
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
Sort all lines and columns in a 2D-dictionary along HIERARCHY column
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
                    list = item[hierarchy_column].split("|")
                    # calculate the new key
                    if len(list) > hierarchy:
                        save = ""
                        count = 0
                        for item_string in list:
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
Sort all lines and columns in a 2D-dictionary along HIERARCHY column
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
Sort all lines and columns in a 2D-dictionary
First dictionary is the character before the dot in VARNAMES (s2.sdlkhre -> s2) => StudyEvent
Second dictionary is based on the entries in the column STUDY_SEGMENT => Formular
"""
def sort_all_lines_and_columns(df, first_sheet_name, all_sheets, file_name, force_single_odm):
    # file_name
    name = file_name.split(".")[0]

    """ Column Names """
    # Save all column names for later and to reconstruct
    column = 0
    separator = "---"
    string_column_names = separator
    column_names = list(df.columns)
    for column_name in column_names:
        string_column_names = (
            string_column_names + str(column_name) + "." + str(column) + separator
        )
        column += 1
    string_column_names = string_column_names[:-1]
    # create a dictionary of the column names with their column number
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
        save = list(row["HIERARCHY"].split("|"))
        string_save = save[0]
        for i in save:
            string_save = string_save + "_" + str(i)
        studyevent = string_save
        # dce
        if pd.notna(row["DCE"]):
            studyevent = row["DCE"]
        # hierarchy and study_segment
        save = list(row["HIERARCHY"].split("|"))
        string_save = save[0]
        for i in save:
            string_save = string_save + "_" + str(i)
        study_segment = string_save
        # study_segment
        if pd.notna(row["STUDY_SEGMENT"]):
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
        english = None
        german = None
        try:
            english = process_codelist(
                row.iloc[dictionary_names.get("VALUE_LABELS", None)]
            )
        except:
            continue
        try:
            german = process_codelist(
                row.iloc[dictionary_names.get("VALUE_LABELS_DE", None)]
            )
        except:
            continue
        # Codelists
        if len(english) > 0 or len(german) > 0:
            # just add the codelist if there isnt an exact codelist yet
            if not check_codelist(english, german, varname, CodeLists):
                # of course only append existing codelists (not nulls)
                if pd.notna(english) or pd.notna(german):
                    CodeLists.append(CodeList(count_cl, varname, english, german))
                    count_cl += 1
    
    if not force_single_odm: # write in more than odm
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
        string_column_names,
        CodeLists,
        dictionary_names,
        varname_number,
        first_sheet_name,
    )


""" 
Extract sheets and names of the sheets 
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

    parser.add_argument("file", help="Pfad zur XLSX-Datei")
    parser.add_argument(
        "--force_single_odm",
        action="store_true",
        help="Write all items in just one ODM (optional Flag)"
    )

    args = parser.parse_args()

    file_path = args.file
    force_single_odm = args.force_single_odm

    if len(sys.argv) < 2:
        print("Please add a path to the xlsx file.")
    else:
        #file_path = sys.argv[2]
        #force_single_odm = sys.argv[1]
        # file name
        file_name = os.path.basename(file_path)
        # process odm
        print(file_name)
        odm(file_path, file_name, force_single_odm)
