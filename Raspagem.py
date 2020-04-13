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
        documento.paragraphs[i].text=documento.paragraphs[i].text.lower()
        paragrafos.append(i)


data = ''
perito = ''
tipo_laudo = ''
oficio = ''
ref_oficio = ''
historico = ''
eficiencia = ''
uso = ''
material =''
outros_elementos =''
conclusao =''
num_laudo =''
InfosArma = []

for i in paragrafos:
    # buscando o campo de data com expressões regulares

    if (re.search(r'aos \d{2} de \w* de \d{4}', documento.paragraphs[i].text)):
        data = re.search(r'aos \d{2} de \w* de \d{4}', documento.paragraphs[i].text).group(0)[4:]

    if (re.search(r', fo\w* designad\w* \w* perit(.*) para', documento.paragraphs[i].text)):
        perito = re.search(r'.erit(.*) para', documento.paragraphs[i].text).group(0)[:-4]
        
    if (re.search(r'exame de(.*)', documento.paragraphs[i].text)):
        tipo_laudo = re.search(r'exame de(.*)', documento.paragraphs[i].text).group(0)[9:]
        
    if (re.search(r'meio d[ao](.*)protocolado', documento.paragraphs[i].text)):
        oficio = re.search(r'meio d[ao](.*)protocolado', documento.paragraphs[i].text).group(0)[7:-12]
        
    if (re.search(r'sob o número (.*)', documento.paragraphs[i].text)):
        protocolo = re.search(r'sob o número (.*)', documento.paragraphs[i].text).group(0)[13:]
        
    if (re.search(r'ref.(.*)',documento.paragraphs[i].text)):
        ref_oficio = re.search(r'ref.(.*)',documento.paragraphs[i].text).group(0)[3:]
        
    
    if (re.search(r'i - histórico', documento.paragraphs[i].text)):
        historico = documento.paragraphs[i+1].text
        
    if (re.search(r'da eficiência:',documento.paragraphs[i].text)):
        eficiencia = documento.paragraphs[i+1].text
        
    
    if (re.search(r'de outros elementos', documento.paragraphs[i].text)):
        uso = documento.paragraphs[i+1].text
        

    if (re.search(r'o material:(.*)', documento.paragraphs[i].text)):
        material = documento.paragraphs[i+2].text
        
    
    if (re.search(r'de outros elementos:(.*)', documento.paragraphs[i].text)):
        outros_elementos = documento.paragraphs[i+1].text
        
  
    if (re.search(r'conclu\wao:(.*)', documento.paragraphs[i].text)):
        conclusao = documento.paragraphs[i+1].text
        
    if (re.search(r'laudo nº (.*)', documento.paragraphs[i].text)):
        num_laudo = re.search(r'laudo nº (.*)', documento.paragraphs[i].text).group(0)[4:]
 
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
