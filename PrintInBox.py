from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

#'''
def PrintTextInBox(Canvas,BoxCords,Margins,Text,Font = 'Helvetica-Bold',Alignment = 'L'):

    # Import Canvas
    c = Canvas
    [X_BL,Y_TR,X_TR,Y_BL] = BoxCords
    [MarginLeft,MarginTop] = Margins

    lines = Text.split('\n')
    n_line = len(lines)

    Width = []
    Size = float(Y_TR - Y_BL - 2*MarginTop)/n_line
    MaxWidth = float(X_TR - X_BL - 2*MarginLeft)

    '''!!!Make Sure the Font File .ttf Below is Located in the Target Folder Path!!!'''
    FontFile = 'Helvetica Bold.ttf'

    # Register a TrueType font
    pdfmetrics.registerFont(TTFont(Font, FontFile))

    for x in lines:
        Width.append(pdfmetrics.stringWidth(x, Font, Size))

    Wmax = max(Width)
    if Wmax > MaxWidth:
        Size *= (MaxWidth/Wmax)
        MarginTop = (Y_TR - Y_BL - n_line*Size) / 2.0
        for i in range(n_line):
            Width[i] = pdfmetrics.stringWidth(lines[i], Font, Size)

    Y_line = Y_TR - MarginTop - Size
    X_lines = [X_BL+MarginLeft] * n_line

    if Alignment == 'R':
        X_lines = [a - b for a, b in zip([X_TR-MarginLeft] * n_line, Width)]
    elif Alignment == 'C':
        X_lines = [a - b/2.0 for a, b in zip([(X_TR+X_BL)/2.0] * n_line, Width)]

    c.setFont(Font, Size)

    # Draw each line
    for i in range(n_line):
        c.drawString(X_lines[i],Y_line,lines[i])
        Y_line -= Size  # Move down for the next line
#'''

# Register a TrueType font
pdfmetrics.registerFont(TTFont('Helvetica-Bold', 'Helvetica Bold.ttf'))

# Create a PDF canvas
c = canvas.Canvas('Test.pdf', pagesize=letter)

LabelPos_Top = [90,410] # Default Top Label Position (canvas cordinates of the label bottom left corner)
LabelPos_Btm = [90, 95] # Default Bottom Label Position (canvas cordinates of the label bottom left corner)

[X,Y]=LabelPos_Top

BoxCord = [X+7,Y+220,X+224,Y+178]
Margin = [5,3]
[X_BL,Y_TR,X_TR,Y_BL] = BoxCord
c.setLineWidth(1.0) 

c.line(X_BL,Y_BL,X_BL,Y_TR)
c.line(X_BL,Y_TR,X_TR,Y_TR)
c.line(X_TR,Y_TR,X_TR,Y_BL)
c.line(X_TR,Y_BL,X_BL,Y_BL)

# Text content
text_content = "SUFFERN CUSTOMER DISTRIBUTION CENTER\n30 DUNNIGAN DRIVE\nSUFFERN, NY 10901 US\nUS"


PrintTextInBox(c,BoxCord,Margin,text_content,Font = 'Helvetica-Bold',Alignment = 'L')

# Save the PDF
c.save()