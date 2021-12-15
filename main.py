import PySimpleGUIQt as sg
import docx2txt
import textwrap
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
from time import sleep

progresso= 100.00

sg.theme('DarkGrey6')

# layout
layout=[[sg.Text('Escolha o arquivo',size=(15,0.8))],
        [sg.Input('',size=(30,0.8),key='ARQUIVO'),sg.FileBrowse('Procurar',size=(10,0.9))],
        [sg.Text('Pasta de saída',size=(15,0.8))],
        [sg.Input('',size=(30,0.8),key="SAIDA"),sg.FolderBrowse('Procurar',size=(10,0.9))],
        [sg.Text('',size=(10,0.2))],
        [sg.Button('Converter',size=(10,1)),sg.Button('Sair', size=(6,1))],
        [sg.Text('',size=(10,0.3))],
        [sg.Stretch(),sg.ProgressBar(progresso,orientation='h',size=(39.3, 15),key='PROGRESS'),sg.Stretch(),sg.Stretch(),sg.Stretch(),sg.Stretch()],
        [sg.Text('',size=(10,0.25))],
       ]

# janela
window = sg.Window('CONVERSOR DE LEGENDAS',layout, icon='Logo_PNG.png',keep_on_top=True)

while True:

    button, values = window.read()

    # converter os dados da janela
    if button == ('Converter'):
        count = 0
        sample_num = 0
        leg = docx2txt.process(values['ARQUIVO']).splitlines()
        leg = leg[0:(len(leg)):2]
        
        for line in leg:
            count += 1
            MAX_W, MAX_H = 1920, 1080
            img = Image.new('RGBA',(MAX_W,MAX_H), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype('arcon.regular.otf', 50)
            current_h, pad = 980, 20
            textwrap.wrap(line, width=100)
            w, h = draw.textsize(line, font=font)
            draw.text((((MAX_W - w) / 2), current_h), line, font=font)
            current_h += h + pad
            img.save(values["SAIDA"]+'/' + f'{count}.png')
            sample_num += 100/len(leg)
            window['PROGRESS'].UpdateBar(sample_num % progresso)
        window['PROGRESS'].UpdateBar(progresso)
        sleep(1)
        window['PROGRESS'].UpdateBar(0)
                
    if button == ('Sair', None):
        break
        window.Close()      
