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
    
def imprimir(texto, dado):
    try:
        print("{}: {}".format(texto, dado))
    except TypeError:
        print("{}: Dado não localizado".format(texto))
 
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


        
#filtrando somente os textos do documetno
for i in range(0,len(documento.paragraphs)):
    if vazio(documento.paragraphs[i].text):
        documento.paragraphs[i].text=documento.paragraphs[i].text.lower()
        paragrafos.append(documento.paragraphs[i].text)  



for num_para in range(0, len(paragrafos)):
    # buscando o campo de data com expressões regulares

    if (re.search(r'aos \d{2} de \w* de \d{4}', paragrafos[num_para])):
        data = re.search(r'aos \d{2} de \w* de \d{4}', paragrafos[num_para]).group(0)[4:]

    if (re.search(r', fo\w* designad\w* \w* perit(.*) para', paragrafos[num_para])):
        perito = re.search(r'.erit(.*) para', paragrafos[num_para]).group(0)[:-4]
        
    if (re.search(r'exame de(.*)', paragrafos[num_para])):
        tipo_laudo = re.search(r'exame de(.*)', paragrafos[num_para]).group(0)[9:]
        
    if (re.search(r'meio d[ao](.*)protocolado', paragrafos[num_para])):
        oficio = re.search(r'meio d[ao](.*)protocolado', paragrafos[num_para]).group(0)[7:-12]
        
    if (re.search(r'sob o número (.*)', paragrafos[num_para])):
        protocolo = re.search(r'sob o número (.*)', paragrafos[num_para]).group(0)[13:]
        
    if (re.search(r'ref.(.*)',paragrafos[num_para])):
        ref_oficio = re.search(r'ref.(.*)',paragrafos[num_para]).group(0)[3:]
        
    
    if (re.search(r'i - hist\wrico', paragrafos[num_para])):
        historico = paragrafos[num_para +1]
        
    if (re.search(r'da efici\wncia:',paragrafos[num_para])):
        eficiencia = paragrafos[num_para +1]
        
    
    if (re.search(r'de outros elementos', paragrafos[num_para])):
        uso = paragrafos[num_para +1] 
        

    if (re.search(r'\w\w material', paragrafos[num_para])):
        material = paragrafos[num_para +2]
        
    
    if (re.search(r'de outros elementos:(.*)', paragrafos[num_para])):
        outros_elementos = paragrafos[num_para +1]
        
  
    if (re.search(r'conclus\wo:(.*)', paragrafos[num_para])):
        conclusao = paragrafos[num_para +1]
        
    if (re.search(r'laudo nº (.*)', paragrafos[num_para])):
        num_laudo = re.search(r'laudo nº (.*)', paragrafos[num_para]).group(0)[4:]
 
    if(re.findall(r'(\w*)[:]..',  paragrafos[num_para])):  # pegano informacoes das armas
        InfosArma.append(re.findall(r'(.*)[:](.*)',  paragrafos[num_para]))
        
        
        
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
