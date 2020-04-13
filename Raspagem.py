from docx import *
import re

#carregando o documento
documento = Document(r'1.docx')

def vazio(paragrafo):
    '''Função para verificar se um paragráfo contém texto'''
    if paragrafo == '':
        return False
    else:
        return True

paragrafos=[] #filtrando os paragrafos sem texto
for i in range(0,len(documento.paragraphs)):
    if vazio(documento.paragraphs[i].text):
        paragrafos.append(i)

#passando tudo para letra maiuscula
for i in paragrafos:
    documento.paragraphs[i].text=documento.paragraphs[i].text.lower()

data = ''
perito = ''
tipo_laudo = ''
oficio = ''
ref_oficio = ''
historico = ''
eficiencia = ''
uso = ''
num_laudo = ''
InfosArma = []


for i in paragrafos:
    data = re.search(r'aos \d{2} de \w* de \d{4}', documento.paragraphs[i].text) # buscando o campo de data com expressões regulares
    if data != None:
        data = data.group(0)[4:]
        break

    for i in paragrafos:
        perito = re.search(r'.erit(.*) para', documento.paragraphs[i].text)
        if perito != None:
            perito = perito.group(0)[:-4]
            break

for i in paragrafos:
    tipo_laudo = re.search(r'exame de(.*)', documento.paragraphs[i].text)
    if tipo_laudo != None:
        tipo_laudo = tipo_laudo.group(0)[9:]
        break
        
for i in paragrafos:
    oficio = re.search(r'meio d[ao](.*)protocolado', documento.paragraphs[i].text)
    if oficio != None:
        oficio = oficio.group(0)[7:-12]
        
        aux=str(oficio).split(' ') #Auxilia e separa oficio com data.
        for i in aux:
            if i != 'de':
                data_oficio=[]
                data_oficio.append(i)
                                
        break

for i in paragrafos:
    protocolo = re.search(r'sob o número (.*)', documento.paragraphs[i].text)
    if protocolo != None:
        protocolo = protocolo.group(0)[13:]
        break

for i in paragrafos:
    ref_oficio = re.search(r'ref.(.*)',documento.paragraphs[i].text)
    if ref_oficio != None:
        ref_oficio = ref_oficio.group(0)[3:]
        break

for i in paragrafos:
    historico = re.search(r'i - histórico:', documento.paragraphs[i].text)
    if historico != None:
        historico = documento.paragraphs[i+1].text
        break

for i in paragrafos:
    eficiencia = re.search(r'da eficiência:',documento.paragraphs[i].text)
    if eficiencia != None:
        eficiencia = documento.paragraphs[i+1].text
        break

for i in paragrafos:
    uso = re.search(r'de outros elementos', documento.paragraphs[i].text)
    if uso != None:
        uso = documento.paragraphs[i+1].text
        break

for i in paragrafos:
    num_laudo = re.search(r'laudo nº (.*)', documento.paragraphs[i].text)
    if num_laudo != None:
        num_laudo = num_laudo.group(0)[8:]
        break

def imprimir(texto, dado):
    try:
        print("{}: {}".format(texto, dado))
    except TypeError:
        print("{}: Dado não localizado".format(texto))

for paragrafoss in documento.paragraphs:
    
    if(re.findall(r'(\w*)[:]..', paragrafoss.text)):  # pegano informacoes das armas
        InfosArma.append(re.findall(r'(.*)[:](.*)', paragrafoss.text))
        
        
imprimir("Data do Laudo", data)
imprimir("Referência oficio", ref_oficio)
imprimir("Protocolo com a data", protocolo)
imprimir("Oficio ", oficio[:-14])
imprimir("Data do Oficio", data_oficio)
imprimir("Tipo do Laudo", tipo_laudo)
imprimir("Peritos", perito)
imprimir("Historico", historico)
imprimir("Eficiência", eficiencia)
imprimir("Uso", uso)
imprimir("N° Laudo", num_laudo)
