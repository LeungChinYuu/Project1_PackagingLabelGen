'''
Install the following libraries in the cmd window before running this script:
pip install pandas
pip install reportlab
pip install barcode
pip install pip install python-barcode
python -m pip install --upgrade setuptools
pip install openpyxl
'''

import pandas as pd
import RFLabelGen
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Lookup Dict for N.W. and G.W. info given RevSku entry
WeightDict = {
    'R31LD23TN':[108,115],
    'R31LD23TX':[110,117],
    'R31LD23FL':[125,133],
    'R31LD23QN':[131,140],
    'R31LD23SQ':[99,105],    
    'R31LD23KG':[194,205],
    'R31LD23TK':[202,215],

    'R3TLD24TX':[110,117],
    'R3TLD24FL':[124,133],
    'R3TLD24QN':[133,142],
    'R3TLD24SQ':[98,104],    
    'R3TLD24KG':[197,208],
    'R3TLD24TK':[199,204],

    'R41LD21TX':[112,119],
    'R41LD21FL':[128,139],
    'R41LD21QN':[135,145],
    'R41LD21SQ':[100,108],    
    'R41LD21KG':[191,206],
    'R41LD21SC':[113,120],

    'M101D20TN':[38,45],
    'M101D20TX':[39,47],
    'M101D20FL':[44,51],
    'M101D20QN':[47,60],
    'M101D20KG':[75,85],

    'R31TD03TK':[202,213],
    'R31TD03TC':[199,212],

    'R310D02TK':[202,213],

    'DT70.12-E78-F052.1-TXL':[169,182],
    'DT70.12-E78-F052.1-Q':[205,223],
    'DT70.12-E78-F052.1-SCK':[138,160],

    'R220D25TX':[50,64],
    'R220D25QN':[78,88],

}

# Lookup Dict for RFSku given RevSku
SkuDict = {
    'R31LD23TN':'483132305',
    'R31LD23TX':'484132306',
    'R31LD23FL':'485132307',
    'R31LD23QN':'486132308',
    'R31LD23SQ':'487132309',    
    'R31LD23KG':'488132300',
    'R31LD23TK':'489132301',

    'R3TLD24TX':'484132407',
    'R3TLD24FL':'485132408',
    'R3TLD24QN':'486132409',
    'R3TLD24SQ':'487132400',    
    'R3TLD24KG':'488132401',
    'R3TLD24TK':'489132402',

    'R41LD21TX':'484142105',
    'R41LD21FL':'485142106',
    'R41LD21QN':'486142107',
    'R41LD21SQ':'487142108',    
    'R41LD21KG':'488142109',
    'R41LD21SC':'524901411',

    'M101D20TN':'467140124',
    'M101D20TX':'468140125',
    'M101D20FL':'469140126',
    'M101D20QN':'470140129',
    'M101D20KG':'472140121',

    'R31TD03TK':'489033109',
    'R31TD03TC':'460902020',

    'R310D02TK':'489023108',

    'DT70.12-E78-F052.1-TXL':'460900001',
    'DT70.12-E78-F052.1-Q':'TBD',
    'DT70.12-E78-F052.1-SCK':'TBD',

    'R220D25TX':'496182760',
    'R220D25QN':'498182760',

}

# Load the Excel file

'''*********************************************************************************************************************************************'''
'''______________________ !!!Make Sure the Target Folder Path and the R&F PO Info Excel File Name are Correct!!! _______________________________'''

TargetFilePath = 'C:/Python Codes/RnFLabelGen/'
ExcelFileName = 'Raymour & Flanigan - NEW lables needed.xlsx'

'''*********************************************************************************************************************************************'''


FilePath_xls = TargetFilePath + ExcelFileName
xls = pd.ExcelFile(FilePath_xls)

# Get the sheet names
sheet_names = xls.sheet_names
print(sheet_names)

# Get the number of sheets
num_sheets = len(xls.sheet_names)
print(f'There are {num_sheets} POs (sheets) in the Excel file.\n\n')

##_______________________ Start to Generate Label PDF for each PO ____________________
## Extract info from every single Excel sheet
# for x in range(num_sheets):
for i in range(num_sheets):
    df = pd.read_excel(FilePath_xls, sheet_name=sheet_names[i])

    df.dropna(axis=0, how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)

    # Set the first row as the header
    new_header = df.iloc[0]  # Take the first row for the header
    df = df[1:]  # Take the data less the header row
    df.columns = new_header  # Set the header row as the df header
    df.reset_index(drop=True, inplace=True)

    RFPO = df.iloc[0,0]
    print('This PO number is ',RFPO,'\n\n')

    Add = df['Ship To Address'].tolist()
    not_nan_condition = df['Ship To Address'].notna()

    Ship2Add = df.iloc[0,1]

    for index, item in enumerate(Add):
        if index != 0 and not_nan_condition[index]:
            Ship2Add += '\n' + item

    print('The Ship To Address is:\n',Ship2Add,'\n\n')


    SN_List = df['Serial number'].tolist()
    # Count non-NaN items (Sku type Qty)
    SN_count = len([item for item in SN_List if pd.notna(item)])
    print('There are ', SN_count, 'Units in this PO', RFPO, '\n\n')

    RevSku_List = df['SKU'].tolist()
    #RFSku_List = df['R&F SKU'].tolist()

    # Count non-NaN items (Sku type Qty)
    Sku_count = len([item for item in RevSku_List if pd.notna(item)])
    print('There are ', Sku_count, 'Skus in this PO', RFPO, '\n\n')

    ## Print one page PDF of 2 label copy for each unit in the PO
    OutputFilePath = TargetFilePath + 'Generated_Labels/'
    OutputFileName = RFPO + '.pdf'
    Contents = {'RFPO': RFPO, 'Ship2Add': Ship2Add, 
    'RevSku': 'TBD',
    'RFSku':'TBD',
    'SN':'TBD',
    'NW': 'TBD',
    'GW': 'TBD'}

    c = canvas.Canvas(OutputFilePath+OutputFileName, pagesize=letter)
    NewCanvas = False
    LabelPos_Top = [90,410] # Default Top Label Position (canvas cordinates of the label bottom left corner)
    LabelPos_Btm = [90, 95] # Default Bottom Label Position (canvas cordinates of the label bottom left corner)

    for i in range(SN_count):
        Contents['SN'] = SN_List[i]
        if pd.notna(RevSku_List[i]):
            Contents['RevSku'] = RevSku_List[i]
            try:
                Contents['RFSku'] = SkuDict[RevSku_List[i]]
            except:
                print('\n**Error**: This Reverie Sku of ', RevSku_List[i], ' is not assigned to an existing R&F Sku!!\n\n')
            try:
                Contents['NW'] = str(WeightDict[RevSku_List[i]][0])
                Contents['GW'] = str(WeightDict[RevSku_List[i]][1])
            except:
                print('\n**Error**: Could not access the N.W. or G.W. info of this Reverie Sku ', RevSku_List[i], '!!\n\n')
        if i != 0:
            NewCanvas = True
        print(Contents)
        
        c = RFLabelGen.RFLabel_Print(c, Contents, LabelPos_Top, NewCanvas, TargetFilePath)
        c = RFLabelGen.RFLabel_Print(c, Contents, LabelPos_Btm, False, TargetFilePath)

    c.save()