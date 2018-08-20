# Update HTML file ----------------------
import fileinput

fileName='BOTOX ONE Registration Quick Look Guide.html'

texttofind = 'href="#"'
texttoreplace = 'href="{{$4080}}&amp;LinkName=BotoxTX%5FONE%5FRegistration%5FQuick%5FLook%5FGuide%5FEF%2D51043"'

with fileinput.FileInput(fileName, inplace=True, backup='.bak') as file:
    for line in file:
        print(line.replace(texttofind, texttoreplace), end = '')

