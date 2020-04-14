from docx import *
import re
import os
import unidecode

os.chdir(r'') #colocar o caminho para a pasta com os arquivos que serão raspados
lista_de_arquivos = os.listdir(os.getcwd())

def EncontraArma(marcadores, texto):
    arma = []
    dicionario_arma = {}
## procura pelas armas no texto e pelo proximo marcador apos a descricao delas
    for indice in range(0, len(marcadores)):
        if(re.findall(r'\Ada.* arma.*:',  marcadores[indice])):
            #evita que pegue a secao sobre as armas como indice pra verificacao de armas
            if(re.findall(r'\Adas.*armas',  marcadores[indice])):
                pass
            else:
                arma.append(indice)
        ### pega o proximo marcador depois das armas e termina a busca
        elif( re.search(r'\Ad\w* .*:', marcadores[indice])):
            arma.append(indice)
            break
            
## verifica se o proximo paragrafo no texto é texto ou chave:valor e guarda a info 
    for  indice in range(0, len(arma) -1):
        dic = {}
        aux = arma[indice] + 1
        posi = 0
    # deixando o texto alinahdo com o marcdor
        while texto[posi]!= marcadores[arma[indice]]:
            posi = posi +1
            
    # verifica se o proximo elemento do texto esta no marcador
    # se o lemento for igual, sinal que descreve em chave:informacao, entao guarda os dados ate a proxima info
        if paragrafos[posi+1] == marcadores[arma[indice]+1]:
            dic['info']= marcadores[aux].split(':')[0]
            aux = aux +1
            while aux != arma[indice+1]:
                lista = marcadores[aux].split(':')
                dic[lista[0]]= lista[1]
                aux = aux + 1
     # se diferente e texto e entao pega ate o proximo marcador
        else:
            dic['info']= PegaTodaTexto(texto, posi, marcadores)
            
        dicionario_arma['Arma'+str(indice)] = dic
    return dicionario_arma 
        

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

for arquivo in lista_de_arquivos:
    documento = Document(arquivo)
    
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



    ## atribuindo as informacoes das armas
    Dados.update(EncontraArma(Marcadores))

    ### mostrando os dados:
    for key, value in Dados.items():
        print(key, ' : ', value,'\n')
