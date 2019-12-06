Attribute VB_Name = "C_ResourceString"
Option Explicit

'公共模块所用资源串
Public Const ERR_TITLE_PROMPT = "Prompt"
Public Const ERR_MSG_RANGE = "Range[@@1..@@2]"
Public Const ERR_MSG_LENGTH = "Length[@@1..@@2]"
Public Const ERR_MSG_ENUM = "Range[@@1]"
Public Const RSC_STR_FRESHING_MOC_ATTR = "Freshing template '@@1' '@@2'......"
Public Const RSC_STR_FINISHED = "Finished."
Public Const RSC_STR_INSERTED_DATA_INTO_SHEET = "Inserted @@1 row(s) data into the sheet '@@2'."
Public Const RSC_STR_FILE_SAVED = "File '@@1' was saved."

'DoubleFrequencyCell模块所用资源串
Public Const RSC_STR_FORMULA = "Formula"
Public Const RSC_STR_INTRANCELL = "IntraNCell"
Public Const RSC_STR_GSMNCELL = "GSMNCell"
Public Const RSC_STR_INTERNCELL_SAME_SECTOR = "InterNCell SameSector"
Public Const RSC_STR_INTRANCELL_DIFF_SECTOR = "InterNCell DifferentSector"
Public Const RSC_STR_FORMULA_ERR = "Formula isn't correct."
Public Const RSC_STR_CELLID_NOT_NUMERIC = "CellID must be numeric."
Public Const RSC_STR_SECTORID_RANGE_ERR = "Sector ID is out of range [0-5]"
Public Const RSC_STR_SELECT_AT_LEAST_1_ROW = "Please select at least one row data."
Public Const RSC_STR_CANNOT_SELECT_EMPTY_ROW = "Can't select empty data row or title row."
Public Const RSC_STR_ATTR_CANNOT_EMPTY = "Column '@@1' of Sheet '@@2' can't be empty."
Public Const RSC_STR_NO_DIFF_FREQ = "No different frequency can be selected."
Public Const RSC_STR_NO_DIFF_FREQ_CELL = "No different frequency can be selected. Cell ID=@@1"
Public Const RSC_STR_NO_SECTOR_MAPPING_CELL = "The cell doesn't have sector mapping cell."
Public Const RSC_STR_SECTOR_MAPPING_CELL = "The sector mapping cell(" + ATTR_CELLID + "=@@1, " + ATTR_BSCNAME + "=@@2) was found."
Public Const RSC_STR_NO_INTRA_FREQ_NCELL = "The cell(" + ATTR_CELLID + "=@@1, " + ATTR_BSCNAME + "=@@2) doesn't have Intra-frequency Neighboring Cell."
Public Const RSC_STR_CELL_WAS_NOT_FOUND = "The cell(" + ATTR_CELLID + "=@@1, " + ATTR_BSCNAME + "=@@2) was not found in sheet '" + MOC_WHOLE_NETWORK_CELL + "'."
Public Const RSC_STR_CELL_WAS_NOT_FOUND_2 = "The cell(@@1) was not found in sheet '" + MOC_WHOLE_NETWORK_CELL + "'."
Public Const RSC_STR_NODEB_WAS_NOT_FOUND = "The NodeB(" + ATTR_NODEBNAME + "=@@1, " + ATTR_BSCNAME + "=@@2) was not found in sheet '" + MOC_WHOLE_NETWORK_CELL + "'."
Public Const RSC_STR_BSC_WAS_NOT_FOUND = "The BSC(" + ATTR_BSCNAME + "=@@1) was not found in sheet '" + MOC_WHOLE_NETWORK_CELL + "'."
Public Const RSC_STR_PROCESSING_BSC = "Processing BSC '@@1'......"
Public Const RSC_STR_PROCESSING_NODEB = "Processing BSC '@@1' NodeB '@@2'......"
Public Const RSC_STR_PROCESSING_CELL = "Processing BSC '@@1' NodeB '@@2' Cell '@@3'......"
Public Const RSC_STR_CANCELLED = "Cancelled."
Public Const RSC_STR_SELECTED_CELL = ATTR_BSCNAME + "=@@1, " + ATTR_NODEBNAME + "=@@2, " + ATTR_CELLID + "=@@3"
Public Const RSC_STR_RPT_SELECTED_CELL = "--Cell(" + ATTR_CELLID + "=@@1, " + ATTR_NODEBNAME + "=@@2, " + ATTR_BSCNAME + "=@@3, RowIndex=@@4): "
Public Const RSC_STR_SELECT_ONE_FREQ = "Please selecte one frequency."
Public Const RSC_STR_SHEET_HAS_NO_RECORD = "The sheet of '@@1' has no record."
Public Const RSC_STR_SHEET_NO_RECORD = "The sheet '@@1' has no record."
Public Const RSC_STR_RNC_NOT_FOUND = "RNC ID of BSC(" + ATTR_BSCNAME + "=@@1) was not found in sheet '@@2'."
Public Const RSC_STR_BSC_NOT_FOUND = "BSC Name of RNC(" + ATTR_RNCID + "=@@1) was not found in sheet '@@2'."
Public Const RSC_STR_PEER_CELL_NOT_FOUND = "Cell '@@1' is peer cell of selected cell's InterFrequencyNCell, but it was not found in sheet '@@2' and '@@3'."
Public Const RSC_STR_SELECTED_FREQ = "The selected frequency: ULFrequency=@@1, DLFrequency=@@2."
Public Const RSC_STR_NAME_VALUE_PAIR_5 = "@@1=@@2, @@3=@@4, @@5=@@6, @@7=@@8, @@9=@@A"
Public Const RSC_STR_NAME_VALUE_PAIR_4 = "@@1=@@2, @@3=@@4, @@5=@@6, @@7=@@8"
Public Const RSC_STR_SHEET_EXISTS_SOME_DATA = "Data was found in sheet '@@1', do you want to continue?"

'ConvertTemplate模块所用资源串
Public Const RSC_STR_CONVERT_DATA = "Convert Data"
Public Const RSC_STR_TITLE_CONFIRM = "Confirm"
Public Const RSC_STR_MSG_CLEAR_SHEET = "The content of '@@1' file will be deleted before converting data, do you want to continue?"
Public Const RSC_STR_MSG_VDF_COL_NOT_FOUND = "The specified '@@1' column of '@@2' sheet was not found in '@@3' file, do you want to continue?"
Public Const RSC_STR_CONVERTING = "Converting......"
Public Const RSC_STR_CONVERTING_MOC = "Converting @@1......"
Public Const RSC_STR_CONVERTING_ATTR = "Converting @@1 @@2 @@3......"
Public Const RSC_STR_PREPARING_FREQ = "Preparing frequency data......"
Public Const RSC_STR_DELETING_MOC = "Deleting invalid @@1......"
Public Const RSC_STR_GENERATING_GSMCELLINDEX = "Generating GSM Cell Index......"
Public Const RSC_STR_CLOSING_WK = "Closing the opened workbook......"
Public Const RSC_STR_SRC_TEMPLATE_FILE_NOT_FOUND = "Source Template File was not found, please input a valid file name."
Public Const RSC_STR_SRC_HW_CME_RNP_DATA_FILE_NOT_FOUND = "Source Huawei CME RNP data file was not found, please input a valid file name."
Public Const RSC_STR_RNC_NOT_FOUND_2 = "RNC ID of BSC(" + ATTR_BSCNAME + "=@@1) was not found in sheet '@@2' workbook '@@3', do you want to continue?"
Public Const RSC_STR_BSC_NOT_FOUND_2 = "BSC Name of RNC(" + ATTR_RNCID + "=@@1) was not found in sheet '@@2' workbook '@@3', do you want to continue?"
Public Const RSC_STR_MAKE_SURE_TEMPLATE_CORRECT = "Please make sure template '@@1' is correct."
Public Const RSC_STR_MAKE_SURE_TEMPLATE_SAME_VERSION = "Please make sure template '@@1' and '@@2' are the same version."
Public Const RSC_STR_FREQ_NOT_FOUND = "The frequency of cell(" + ATTR_CELLID + "=@@1, " + ATTR_RNCID + "=@@2) was not found."
Public Const RSC_STR_FREQ_INFO = "The frequency info: " + ATTR_RNCID + "=@@1, " + ATTR_CELLID + "=@@2, " + ATTR_UARFCNUPLINK + "=@@3, " + ATTR_UARFCNDOWNLINK + "=@@4."
