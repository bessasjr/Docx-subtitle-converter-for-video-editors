import docx2txt
import textwrap
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw

leg = docx2txt.process('<insert your subtitle file>.docx').splitlines()
leg = leg[0::2]

cont = 0

for line in leg:
    cont += 1
    MAX_W, MAX_H = 1920, 1080
    img = Image.new('RGBA', (MAX_W, MAX_H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype('enter a text font for the caption', 50) # 50 is the font size
    current_h, pad = 960, 20
    textwrap.wrap(line, width=100)
    w, h = draw.textsize(line, font=font)
    draw.text((((MAX_W - w) / 2), current_h), line, font=font)
    current_h += h + pad
    img.save('<choose the folder to save the files>/{}.png'.format(cont))
