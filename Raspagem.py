from docx import *
import re

#carregando o documento
documento = Document(r'')

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
aux = None


for i in paragrafos:
    aux = re.search(r'aos \d{2} de \w* de \d{4}', documento.paragraphs[i].text) # buscando o campo de data com expressões regulares
    if aux != None:
        data = aux.group(0)[4:]

    aux = re.search(r'.erit(.*) para', documento.paragraphs[i].text)
    if aux != None:
        perito = aux.group(0)[:-4]
        
    aux = re.search(r'exame de(.*)', documento.paragraphs[i].text)
    if aux != None:
        tipo_laudo = aux.group(0)[9:]
        
    aux = re.search(r'meio d[ao](.*)protocolado', documento.paragraphs[i].text)
    if aux != None:
        oficio = aux.group(0)[7:-12]
        
    aux = re.search(r'sob o número (.*)', documento.paragraphs[i].text)
    if aux != None:
        protocolo = aux.group(0)[13:]
        
    aux = re.search(r'ref.(.*)',documento.paragraphs[i].text)
    if aux != None:
        ref_oficio = aux.group(0)[3:]
        
    aux = re.search(r'i - histórico', documento.paragraphs[i].text)
    if aux != None:
        historico = documento.paragraphs[i+1].text
        
    aux = re.search(r'da eficiência:',documento.paragraphs[i].text)
    if aux != None:
        eficiencia = documento.paragraphs[i+1].text
        
    aux = re.search(r'de outros elementos', documento.paragraphs[i].text)
    if aux != None:
        uso = documento.paragraphs[i+1].text      

    aux = re.search(r'o material:(.*)', documento.paragraphs[i].text)
    if aux != None:
        material = documento.paragraphs[i+2].text
        
    aux = re.search(r'de outros elementos:(.*)', documento.paragraphs[i].text)
    if aux != None:
        outros_elementos = documento.paragraphs[i+1].text
        
    aux = re.search(r'conclus\wo:(.*)', documento.paragraphs[i].text)
    if aux != None:
        conclusao = documento.paragraphs[i+1].text
        
    aux = re.search(r'laudo nº (.*)', documento.paragraphs[i].text)
    if aux != None:
        num_laudo = aux.group(0)[4:]
    
    if(re.findall(r'(\w*)[:]..',  documento.paragraphs[i].text)):  # pegano informacoes das armas
        InfosArma.append(re.findall(r'(.*)[:](.*)',  documento.paragraphs[i].text))
        
        
imprimir("Data do Laudo", data)
imprimir("Referência oficio", ref_oficio)
imprimir("Protocolo com a data", protocolo)
imprimir("Oficio com a data", oficio)
imprimir("Tipo do Laudo", tipo_laudo)
imprimir("Peritos", perito)
imprimir("Historico", historico)
imprimir("Eficiência", eficiencia)
imprimir("Uso", uso)
imprimir("N° Laudo", num_laudo)
