# pyRevit script
# Save as a .py file in your pyRevit scripts folder and run from pyRevit
from pyrevit import revit, DB, forms, script
import csv
import sys

doc = revit.doc

# Prompt user for CSV save location and file name
output_path = forms.save_file(file_ext='csv',
                             default_name='shared_parameters.csv',
                             title='Save Shared Parameters CSV')
if not output_path:
    script.exit()

# Collect all SharedParameterElement objects in the project
sp_collector = DB.FilteredElementCollector(doc).OfClass(DB.SharedParameterElement)
shared_params = list(sp_collector)

# Prepare mapping from GUID string to set of family names
param_families = {}

# Initialize mapping keys for all shared parameters
for sp in shared_params:
    try:
        ext_def = sp.GetDefinition()
        guid_str = str(ext_def.GUID)
    except Exception:
        guid_str = ''
    param_families[guid_str] = set()

# Inspect loaded Family elements and their symbols for shared parameters
families = DB.FilteredElementCollector(doc).OfClass(DB.Family).ToElements()
for fam in families:
    try:
        symbol_ids = fam.GetFamilySymbolIds()
    except Exception:
        symbol_ids = []
    for sid in symbol_ids:
        symbol = doc.GetElement(sid)
        if symbol is None:
            continue
        for p in symbol.Parameters:
            try:
                if p.IsShared:
                    defn = p.Definition
                    # Only ExternalDefinition has GUID for shared parameters
                    if isinstance(defn, DB.ExternalDefinition):
                        guid = str(defn.GUID)
                        if guid in param_families:
                            param_families[guid].add(fam.Name)
            except Exception:
                # ignore parameters that raise on access
                continue

# Inspect placed FamilyInstance elements for instance-level family parameters
fi_collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).ToElements()
for fi in fi_collector:
    try:
        fam_name = fi.Symbol.Family.Name if fi.Symbol and fi.Symbol.Family else ''
    except Exception:
        fam_name = ''
    for p in fi.Parameters:
        try:
            if p.IsShared:
                defn = p.Definition
                if isinstance(defn, DB.ExternalDefinition):
                    guid = str(defn.GUID)
                    if guid in param_families:
                        param_families[guid].add(fam_name)
        except Exception:
            continue

# Write CSV
try:
import codecs
import csv

# Ensure unicode name exists in case of different runtimes
try:
    unicode
except NameError:
    unicode = str

with open(output_path, 'wb') as csvfile:
    # Write UTF-8 BOM so Excel recognizes UTF-8
    csvfile.write(codecs.BOM_UTF8)

    writer = csv.writer(csvfile)
    writer.writerow(['Parameter Name', 'GUID', 'Parameter Type', 'Families'])

    for sp in shared_params:
        # compute name, guid, ptype, families_list as unicode strings
        try:
            ext_def = sp.GetDefinition()
            name = ext_def.Name
            guid = str(ext_def.GUID)
            ptype = ext_def.ParameterType.ToString()
            families_list = ';'.join(sorted(param_families.get(guid, []))) or 'Project'
        except Exception:
            continue

        # Encode each cell to UTF-8 bytes for Python 2 csv.writer
        row = []
        for cell in (name, guid, ptype, families_list):
            if isinstance(cell, unicode):
                row.append(cell.encode('utf-8'))
            else:
                row.append(str(cell))
        writer.writerow(row)
except Exception as e:
    forms.alert('Failed to write CSV:\n{}'.format(e), title='Error')
    script.exit()

forms.alert('Shared parameters exported to:\n{}'.format(output_path), title='Export Complete')