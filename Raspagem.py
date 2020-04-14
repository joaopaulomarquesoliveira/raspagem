from docx import *
import re
import os
import unidecode


#carregando o documento
documento = Document()

    
def imprimir(texto, dado):
    try:
        print("{}: {}\n".format(texto, dado))
    except TypeError:
        print("{}: Dado não localizado\n".format(texto))
        

def PegaTodaTexto(texto, lugar_texto, Marcadores):
    aux = 0
    for posicao in range(0, len(Marcadores)):
        if Marcadores[posicao] == texto[lugar_texto]:
            aux = posicao +1
    armazenar = []
    for posicao in range(lugar_texto +1, len(texto)):
        if Marcadores[aux]== texto[posicao]:
            break
        armazenar.append(texto[posicao])
    return armazenar

def DefineMarcadores(texto):
    marcadores = []
    for texto in paragrafos:
        if re.findall(r'.*:',  texto):
            marcadores.append(texto)
    return marcadores

def FiltraTexto(documento):
    paragrafos = []
    for paragrafo in documento.paragraphs:
        paragrafos.append(paragrafo.text.lower().strip()) 
    paragrafos = [ elem for elem in paragrafos if elem != '']      
    return paragrafos

def FormataTexto( texto):
    paragrafos =[]
    for partes in texto:
        paragrafos.append(unidecode.unidecode(partes))
    return paragrafos

def SelecionarParagrafos( texto, localizacao):
    print(range(0, len(paragrafos)),'\n ??', localizacao )
    salvar = []
    while texto[localizacao].islower():
        salvar.append(texto[localizacao])
        localizacao = localizacao + 1
    return salvar

Dados =  {}
Marcadores = []
Infos = []
paragrafos=[] 

#filtrando somente os textos do documetno
paragrafos = FormataTexto(FiltraTexto(documento))
Marcadores = DefineMarcadores(paragrafos)
        
for num_para in range(0, len(paragrafos)):

### busca os campos dentro do texto
    if (re.search(r'aos \d{2} de \w* de \d{4}', paragrafos[num_para])):
        Dados['data'] = re.search(r'aos \d{2} de \w* de \d{4}', paragrafos[num_para]).group(0)[4:]

    if (re.search(r', fo\w* designad\w* \w* perit(.*) para', paragrafos[num_para])):
        Dados['perito'] = re.search(r'.erit(.*) para', paragrafos[num_para]).group(0)#[:-4]
        
    if (re.search(r'proceder \w* exame \w* (.*)', paragrafos[num_para])):
        Dados['tipo_laudo'] = re.search(r'proceder \w* exame \w* (.*) a fim', paragrafos[num_para]).group(0)
        
    if (re.search(r'meio d[ao](.*)protocolado', paragrafos[num_para])):
        Dados['oficio'] = re.search(r'meio d[ao](.*)protocolado', paragrafos[num_para]).group(0)[7:-12]
        
    if (re.search(r'sob o numero (.*)', paragrafos[num_para])):
        Dados['protocolo'] = re.search(r'sob o numero (.*)', paragrafos[num_para]).group(0)[13:]
        
    if (re.search(r'ref.(.*)',paragrafos[num_para])):
        Dados['ref_oficio'] = re.search(r'ref.(.*)',paragrafos[num_para]).group(0)[3:]
        
    if (re.search(r'laudo no.*', paragrafos[num_para])):
        Dados['num_laudo'] = re.search(r'laudo no(.*)', paragrafos[num_para]).group(0) 

#### utilizam funcao para saber ate onde vai o texto    
    
    if (re.search(r'i - historico', paragrafos[num_para])):
        Dados['historico'] =  PegaTodaTexto(paragrafos, num_para, Marcadores)
        
    if (re.search(r'da efici\wncia',paragrafos[num_para])):
        Dados['eficiencia'] = PegaTodaTexto(paragrafos, num_para, Marcadores)

    if (re.search(r'de outros elementos:', paragrafos[num_para])):
        Dados['uso'] = paragrafos[num_para + 1] 

    if (re.search(r'fo\w* encaminhad\w* \w* .* mater\w*:', paragrafos[num_para])):
        Dados['material'] = PegaTodaTexto(paragrafos, num_para, Marcadores)

    if (re.search(r'\Aiii .* conclusao:(.*)', paragrafos[num_para])):
        Dados['conclusao'] = PegaTodaTexto(paragrafos, num_para, Marcadores)

### pegano informacoes das armas
    if(re.findall(r'(\w*)[:]..',  paragrafos[num_para])):  
        Infos.append(re.findall(r'(.*)[:](.*)',  paragrafos[num_para]))

        


##atribuindo mais dados ao dicionario: bug quando ha repeticao de chaves, perde-se info
for itens in Infos:
    Dados.update(itens)
    
### mostrando os dados:
for key, value in Dados.items():
    print(key, ' : ', value,'\n')
