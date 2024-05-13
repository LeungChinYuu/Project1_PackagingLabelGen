
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

import barcode
from barcode.writer import ImageWriter


## Function to draw rectangular frame: 
# Canvas: Target canvas
# (X_BL,Y_BL),(X_TR,Y_TR): XY cords for the Frame bottom left and top right corners
# T: Line thickness in point
def DrawFrame(Canvas,X_BL,Y_BL,X_TR,Y_TR,T=1.0):
    Canvas.setLineWidth(T) 

    Canvas.line(X_BL,Y_BL,X_BL,Y_TR)
    Canvas.line(X_BL,Y_TR,X_TR,Y_TR)
    Canvas.line(X_TR,Y_TR,X_TR,Y_BL)
    Canvas.line(X_TR,Y_BL,X_BL,Y_BL)

## Function to write text on label with cord X, Y: 
def WriteText(Canvas,X,Y,Text,font = 'Helvetica',size = 12.5):
    Canvas.setFont(font, size)
    Canvas.drawString(X,Y,Text)

## Function to create a Code128 PNG Image:
def Create_128BarcodePNG(Data, SavePath, Width=0.4, Height=20, write_text=False):
    
    # Default parameters:
    Dip=600    # You can adjust the dpi (dots per inch) to change the size of the barcode
    font_size= 12
    text_distance=5

    output_path=SavePath #'C:/Users/charlie.liang/PythonCodes/SN_Barcodes/'
    module_width=Width 
    module_height=Height
    
    # Create the writer with custom options
    writer = ImageWriter()

    writer.dpi = Dip  # for a larger barcode image
    
    CODE128 = barcode.get_barcode_class('code128')

    # Create the barcode using the custom writer
    code128_barcode = CODE128(Data, writer=writer)
    
    # Set barcode options
    Options = {'module_width':module_width, 'module_height':module_height, 'font_size': font_size,'text_distance':text_distance,'write_text' : write_text}
    
    # Save the barcode as an image
    barcode_filename = code128_barcode.save(output_path+Data, Options)

    # Print the path to the saved file
    print(barcode_filename)


## Main Function to create RF Label

## Function to print one RF Label on specified position of a given canvas
'''
Function RFLabel_Print(Canvas, Content, Cord, NewCanvas=False)

    >> Canvas: Passing specified canvas for adding the label

    >> Content = {'RFPO': '01174E8IWKN', 
    'Ship2Add': 'NEW JERSEY DISTRIBUTION CENTER\n341 SWEDESBORO AVENUE\nGIBBSTOWN, NJ 08027\nUS',
    'RevSku': 'R31TD03TK',
    'RFSku':'489033109',
    'SN':'20230920-Y000623',
    'NW': '202',
    'GW': '214'}

    # The contents above are the example
    
    >> Cords = [X,Y]: X and Y cordinates of the label bottom left corner on the Top label PDF

    >> NewCanvas: Set True for creating label from blank PDF

    >> TargetFilePath: Assign a default file path for storing all the generated barcode images

    >> Del_BC_PNG: Remove all barcode PNG files after Label Generation
'''
def RFLabel_Print(Canvas, Content, Cords = [90,410], NewCanvas=False, TargetFilePath = 'C:/Program Files/R&FLabelGen/', Del_BC_PNG = True):
    # Import Canvas
    c = Canvas
    
    # Create new page if needed
    if NewCanvas == True:
        c.showPage()

    '''____________________ Draw Label Frame Lines ______________________________'''
    #Label Frame Line Cords Reference
    [X,Y] = Cords
    Xref = [x-90 for x in [90,97,235,314,405,433,515,523]]
    Yref = [x-410 for x in [410,417,490,538,588,645,688,698]]
    Xcord = [x+X for x in Xref]
    Ycord = [x+Y for x in Yref]

    # Draw label frame
    DrawFrame(c,Xcord[0],Ycord[0],Xcord[-1],Ycord[-1])
    DrawFrame(c,Xcord[1],Ycord[1],Xcord[-2],Ycord[2])
    DrawFrame(c,Xcord[1],Ycord[2],Xcord[-2],Ycord[4])
    DrawFrame(c,Xcord[1],Ycord[2],Xcord[-2],Ycord[3])
    DrawFrame(c,Xcord[1],Ycord[2],Xcord[2],Ycord[4])
    DrawFrame(c,Xcord[1],Ycord[4],Xcord[-2],Ycord[-2])
    DrawFrame(c,Xcord[1],Ycord[4],Xcord[-2],Ycord[-3])
    DrawFrame(c,Xcord[1],Ycord[4],Xcord[-2],Ycord[-2])
    DrawFrame(c,Xcord[1],Ycord[4],Xcord[3],Ycord[-3])
    DrawFrame(c,Xcord[1],Ycord[4],Xcord[4],Ycord[-3])
    DrawFrame(c,Xcord[1],Ycord[-3],Xcord[5],Ycord[-2])

    '''____________________ Extract Customized Information______________________________'''
    #RFPO = RevPO = '01174E8IWKN'
    #Ship2Add = 'NEW JERSEY DISTRIBUTION CENTER\n341 SWEDESBORO AVENUE\nGIBBSTOWN, NJ 08027\nUS'
    #RevSku = 'R31TD03TK'
    #RFSku = '489033109'
    #SN = '20230920-Y000623'
    #NW = '202'
    #GW = '214'
    RFPO = Content['RFPO']
    RevPO = RFPO                # Assign with Reverie PO if needed        
    Ship2Add = Content['Ship2Add']
    RevSku = Content['RevSku']
    RFSku = Content['RFSku']
    SN = Content['SN']
    NW = Content['NW']
    GW = Content['GW']
  
    '''____________________ Attach All Text Contents on Label ______________________________'''
    # Setup title text in labels
    
    # Customize text position and font size according to length of RevSku:
    [Xc, Fc] = [40, 13]
    if len(RevSku) > 9:
        [Xc, Fc] = [15, 11]

    # LabelText_Dict： 
    # Keys are the fill-in row and column number on the label;  
    # Velues are a list of [content text, [xcord,ycord],[font, size]] for the title or fixed text content
    LabelText_Dict = {
        '1A,1':     ['Vendor Name & Address',                                   [X+110,Y+263],      ['Helvetica', 12.5]],
        '1B,1':     ['Reverie, 750 Denison Ct, Bloomfield Hills, MI 48302',     [X+18, Y+245],      ['Helvetica-Bold', 13]],
        '1A,2':     ['Vendor Code',                                             [X+348,Y+263],      ['Helvetica', 12.5]],
        '1B,2':     ['REVR',                                                    [X+360,Y+245],      ['Helvetica-Bold', 16]],
        '2A,1':     ['Raymour & Flanigan Ship to Address',                      [X+15, Y+220],      ['Helvetica', 12.5]],
        '2A,2':     ['Net Weight',                                              [X+237,Y+220],      ['Helvetica', 12.5]],
        '2B,2A':    [NW,                                                        [X+240,Y+195],      ['Helvetica-Bold', 15.5]],
        '2B,2B':    ['LBS',                                                     [X+280,Y+195],      ['Helvetica-Bold', 15.5]],
        '2A,3':     ['Gross Weight',                                            [X+330,Y+220],      ['Helvetica', 12.5]],
        '2B,3A':    [GW,                                                        [X+340,Y+195],      ['Helvetica-Bold', 15.5]],
        '2B,3B':    ['LBS',                                                     [X+380,Y+195],      ['Helvetica-Bold', 15.5]],  
        '3A,1':     ['Reverie SKU',                                             [X+42, Y+163],      ['Helvetica', 12.5]],
        '3B,1':     [RevSku,                                                    [X+Xc, Y+140],      ['Helvetica-Bold', Fc]],  #40  13
        '3A,2A':    ['R&F SKU',                                                 [X+150,Y+163],      ['Helvetica', 12.5]],
        '3B,2A':    [RFSku,                                                     [X+150,Y+140],      ['Helvetica-Bold', 13]],
        '3B,2B':    [RFSku,                                                     [X+305,Y+138],      ['Helvetica', 12]],
        '4A,1':     ['PO Number',                                               [X+42, Y+113],      ['Helvetica', 12.5]],
        '4B,1':     [RevPO,                                                     [X+35, Y+90],       ['Helvetica-Bold', 13]],
        '4A,2A':    ['R&F PO',                                                  [X+150,Y+113],      ['Helvetica', 12.5]],
        '4B,2A':    [RFPO,                                                      [X+150,Y+90],       ['Helvetica-Bold', 13]],
        '4B,2B':    [RFPO,                                                      [X+297,Y+90],       ['Helvetica', 12]],
        '5A,1':     ['Serial Number',                                           [X+46, Y+65],       ['Helvetica', 12.5]],
        '5B,1':     [SN,                                                        [X+28, Y+40],       ['Helvetica-Bold', 15]],
        '5C,1':     ['Made in Vietnam',                                         [X+45, Y+15],       ['Helvetica-Bold', 12.5]],
        '5B,2':     [SN,                                                        [X+250,Y+20],       ['Helvetica', 12]]
    }

    DictKey = list(LabelText_Dict.keys())
    
    for loc in LabelText_Dict:
        Cont = LabelText_Dict[loc][0]
        Xt = LabelText_Dict[loc][1][0]
        Yt= LabelText_Dict[loc][1][1]
        Ft = LabelText_Dict[loc][2][0]
        FtSz = LabelText_Dict[loc][2][1]                    

        WriteText(c,Xt,Yt,Cont,font = Ft,size = FtSz)


    '''Enter Ship2Address with multiple line text'''
    # Print Ship to Address
    # Set starting position
    x = X+30
    y = Y+208
    line_height = 9 # Typically the font size

    # Split the text into lines
    lines = Ship2Add.split('\n')

    # Draw each line
    for line in lines:
        WriteText(c,x,y,line,font = 'Helvetica-Bold',size = 9)
        c.drawString(x, y, line)
        y -= line_height  # Move down for the next line


    '''____________________ Generate and Attach Barcodes on Label ______________________________'''

    ## ************************ !!!Make Sure the Barcode PNG File Save Path are Correct!!! *****************************
    # Replace with your own target folder path for geneterated barcode image storage
    FolderPath_Sku = TargetFilePath + 'RnF_Sku_Barcodes/'       ##'C:/Program Files/R&FLabelGen/RnF_Sku_Barcodes/'
    FolderPath_PO = TargetFilePath + 'RnF_PO_Barcodes/'         ##'C:/Program Files/R&FLabelGen/RnF_PO_Barcodes/'
    FolderPath_SN = TargetFilePath + 'SN_Barcodes/'             ##'C:/Program Files/R&FLabelGen/SN_Barcodes/'

    # Create barcode PNGs
    Create_128BarcodePNG(RFSku, FolderPath_Sku, 0.45, 13)
    Create_128BarcodePNG(RFPO, FolderPath_PO, 0.5, 10)
    Create_128BarcodePNG(SN, FolderPath_SN, 0.4, 20)

    PicPath_Sku = FolderPath_Sku + RFSku + '.png'
    PicPath_PO =  FolderPath_PO + RFPO + '.png'
    PicPath_SN = FolderPath_SN + SN + '.png'

    # Attach barcode
    c.drawImage(PicPath_Sku, X+290, Y+148, width=95, height=25)
    c.drawImage(PicPath_PO, X+265, Y+102, width=140, height=22)
    c.drawImage(PicPath_SN, X+231, Y+30, width=145, height=40)

    if Del_BC_PNG:
        folder_paths = [FolderPath_Sku, FolderPath_PO, FolderPath_SN]
        for folder_path in folder_paths:
            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.remove(file_path)

    return c
