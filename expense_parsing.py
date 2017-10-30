from openpyxl import load_workbook

from sets import Set

# TODO CONFIGURE: Must set to be located in the same folder as this file
XLSX_NAME = 'Expenses.xlsx'

# Workbook that has all expenses in it
WB = load_workbook(XLSX_NAME)

#######################################################################################################################
################################################### STEP 1 !!!!!!!  ####################################################
# 1. Configure TODO
#######################################################################################################################

# TODO CONFIGURE: Must set to be all the sheet names you want to skip
SKIPPED_SHEETS = ['Totals', 'Before-Work', 'LabelMap']

def sheet_date_names():
  arr = WB.get_sheet_names()
  for sheet_name in SKIPPED_SHEETS:
    arr.remove(sheet_name)
  return arr

# [Array] an array of sheet names that have item names and costs on them
SHEET_DATE_NAMES_ARRAY = sheet_date_names()

# A set of every item name that is entered in as an expense
all_item_names = Set()

# @return a list of all unique item names across all sheets
def get_all_unique_item_names():
  for sheet_name in SHEET_DATE_NAMES_ARRAY:
    add_item_names(WB[sheet_name])
  return all_item_names

def add_item_names(sheet):
  for col_row in sheet['B']:
    all_item_names.add(col_row.value.strip())
  return True

#######################################################################################################################
################################################### STEP 2 !!!!!!!  ####################################################
# 1. Configure TODO
#######################################################################################################################

# TODO CONFIGURE: Must be set to the index of the last row with an item name and label pair on sheet LabelMap
LABEL_MAP_LAST_ROW_INDEX = 182

# This is a list of all labels that we use to label a type of expense.  Note that the actual label may not be in this
# set however every actual label should map to something in this set.
all_possible_item_labels = Set()

# @return a map of each item name to a label
def map_item_names_to_label():
  label_map = {}
  label_map_sheet = WB['LabelMap']
  for index in range(1, LABEL_MAP_LAST_ROW_INDEX + 1):
    item_name = label_map_sheet[index][0].value.strip()
    label = label_map_sheet[index][1].value
    label_map[item_name] = label
    all_possible_item_labels.add(label)
  return label_map

# Map of every time name to a label
LABEL_MAP = map_item_names_to_label()


#######################################################################################################################
################################################### STEP 3 !!!!!!!  ####################################################
# 1. Configure TODO
# 2. Run cost_highest_to_lowest
#######################################################################################################################

# @return a map where the keys are a label and value are an array of row data
# eg:
# grocery: [
#            {
#              name: 'haight street market',
#              cost: '130',
#              sheet: 'meow cat meow sheet name',
#            },
#            ...
#          ]
def map_label_to_row_on_sheet():
  cray_zay_map = {}
  for sheet_name in SHEET_DATE_NAMES_ARRAY:
    sheet = WB[sheet_name]
    last_index = 100
    for index in range(1, last_index):
      item_name = sheet[index][1].value
      if item_name is None:
        continue
      key_label
      if item_name.strip() in LABEL_MAP:
        key_label = LABEL_MAP[item_name.strip()]
      elif item_name.strip() in all_possible_item_labels:
        key_label = item_name.strip()
      else:
        print "Error!!!! The following item was not mapped to a label: " + item_name
        print "Make sure you add that item_name to LabelMaps in Google Sheets"
      cost = sheet[index][2].value
      if (key_label not in cray_zay_map):
        cray_zay_map[key_label] = []
      cray_zay_map[key_label].append({'name': item_name, 'cost': cost, 'sheet': sheet_name })
  return cray_zay_map

MAP_OF_LABELS_TO_ITEM_ROWS = map_label_to_row_on_sheet()

def calculate_costs_per_label():
  mappy = {}
  for key in MAP_OF_LABELS_TO_ITEM_ROWS.keys():
    yolo = sum(MAP_OF_LABELS_TO_ITEM_ROWS[key][i]['cost'] for i in range(0, len(MAP_OF_LABELS_TO_ITEM_ROWS[key])))
    mappy[key] = yolo
  return mappy

MEOW = calculate_costs_per_label()

print "TODO: Check for negatives and go through ? and categorize"

def cost_highest_to_lowest():
  return sorted(MEOW, key=MEOW.get, reverse=True)
