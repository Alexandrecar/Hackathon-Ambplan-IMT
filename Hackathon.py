from asyncio.windows_events import NULL
from datetime import datetime, timedelta, date
import requests
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import os, os.path
import xlsxwriter

#login e senha do inlabs
login = "seu@email.com"
senha = "sua senha"

#terminações dos dois arquivos que serão baixados separadas por espaço
tipo_dou="DO1 DO1E" 

#inlabs
url_login = "https://inlabs.in.gov.br/logar.php"
url_download = "https://inlabs.in.gov.br/index.php?p="

payload = {"email" : login, "password" : senha}
headers = {
    "Content-Type": "application/x-www-form-urlencoded", #como o formulário deve ser codificado e enviado
    "Accept" : "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" #resposta bem sucedida do web server 
    }
s = requests.Session()
#variável que armazenará os nomes dos arquivos que não foram verificados
FALHA= []


def download():
    #caso o programa tenha acessado o inlabs (logado)
    if s.cookies.get('inlabs_session_cookie'):
        #obtem o cookie da sessão
        cookie = s.cookies.get('inlabs_session_cookie')
    #caso o login falhe o pragrama termina
    else:
        print("Falha ao obter cookie. Verifique suas credenciais");
        exit(37)
    
    #obtenção do dia atual:
    ano = date.today().strftime("%Y")
    mes = date.today().strftime("%m")
    dia = date.today().strftime("%d")
    test_date = datetime(int(ano), int(mes), int(dia))

    #cálculo do último dia útil:
    diff = 1
    if test_date.weekday() == 0:
        diff = 3
    elif test_date.weekday() == 6:
        diff = 2
    else :
        diff = 1
    res = test_date - timedelta(days=diff)

    ano = res.strftime("%Y")
    mes = res.strftime("%m")
    dia = res.strftime("%d")
    
    data_completa = ano + "-" + mes + "-" + dia #último dia útil
    
    for dou_secao in tipo_dou.split(' '):
        #montagem das urls para download dos arquivos .zip
        print("Aguarde Download...")
        url_arquivo = url_download + data_completa + "&dl=" + data_completa + "-" + dou_secao + ".zip"
        print(url_arquivo)
        #montagem do headers
        cabecalho_arquivo = {'Cookie': 'inlabs_session_cookie=' + cookie, 'origem': '736372697074'}
        #downloads dos arquivos .zip
        response_arquivo = s.request("GET", url_arquivo, headers = cabecalho_arquivo)
        #se o headers for identificado, o status code retorna 200, indicando o download bem sucedido
        if response_arquivo.status_code == 200:
            #o programa abre os arquivos .zip
            with open(data_completa + "-" + dou_secao + ".zip", "wb") as f:
                #lista os conteúdos
                f.write(response_arquivo.content)
                print("Arquivo %s salvo." % (data_completa + "-" + dou_secao + ".zip"))
            del response_arquivo
            del f
            #extração dos arquivos .zip
            with ZipFile((data_completa + "-" + dou_secao + ".zip"), 'r') as zipObj:
                zipObj.extractall()
        #se o headers não for identificado, o status code retorna 404, 
        # indicando que o arquivo .zip não foi encontrado
        elif response_arquivo.status_code == 404:
            print("Arquivo não encontrado: %s" % (data_completa + "-" + dou_secao + ".zip"))
    
    print("Aplicação encerrada")

#separação dos dados dos arquivos xml e construção do excel
def processar():
    #construção do arquivo excel "Leis.xlsx" vazio
    workbook = xlsxwriter.Workbook('Leis.xlsx')
    #construção da página excel vazia
    worksheet = workbook.add_worksheet()
    i = 0
    bold = workbook.add_format({'bold': True})
    worksheet.set_column("C:L",20)
    #definição das categorias ao colocar em colunas na linha 0
    worksheet.write(0,2, "Nome", bold)
    worksheet.write(0,3, "Categoria", bold)
    worksheet.write(0,4, "Órgão Emissor", bold)
    worksheet.write(0,5, "Âmbito Primário", bold)
    worksheet.write(0,6, "Número", bold)
    worksheet.write(0,7, "Data de publicação", bold)
    worksheet.write(0,8, "Data Início da Vigência", bold)
    worksheet.write(0,9, "Data fim da vigência", bold)
    worksheet.write(0,10, "Texto alternativo", bold)
    worksheet.write(0,11, "Rótulo alternativo", bold)
    #para cada arquivo extraido
    for name in os.listdir('.'):
        #caso o arquivo termine com .xml
        if os.path.isfile(name) and name[-4:] == ".xml":
            print(name)
            try:
                #divisao do arquivo xml em seus elementos
                tree = ET.parse(name)
                #obtenção do elemento root (contém todos os elementos)
                xml = tree.getroot()
                #para cada elemento 'article' dentro do root e sua sub-tree
                for article in xml.iter('article'):
                    #os conteúdos dos atributos do 'article' são armazenados
                    CATEGORIA = article.get("artType")
                    ORGÃO_EMISSOR = article.get("artCategory")
                    DATA_INÍCIO = article.get("pubDate")
                    # para cada elemento filho de article 
                    # (pelo que percebemos sempre é o body e um elemento  vazio)
                    for body in article:
                        #se for um elemento com algum conteúdo 
                        # (necessário para filtrar elementos vazios)
                        if len(body) != 0:
                            #armazena o texto da lei ao acessar o texto do elemento filho 5 do body
                            TEXTO = body[5].text
                            #formatação do texto, remover marcações e mudar de linha
                            TEXTO = TEXTO.replace("</p><p>","\n")
                            TEXTO = TEXTO.replace('''<p class="identifica">''',"")
                            TEXTO = TEXTO.replace('''<p></p>''',"")
                            TEXTO = TEXTO.replace('''<p>''',"")
                            TEXTO = TEXTO.replace('''<p class="subtitulo">''',"")
                            DATA_FINAL = "Não Aplicável"
                            #armazena o nome da lei ao acessar o texto do elemento filho 0 do body
                            NOME = body[0].text

                            #busca pelo número da lei no nome considerando que 
                            #o número aparecerá após "Nº" e terminará em até 15 caracteres
                            string = NOME.upper()
                            IpositionNome = string.find("Nº")
                            stringReduzida = string[IpositionNome:IpositionNome+15]
                            num = ""
                            controle = 0
                            for c in stringReduzida:
                                 if c.isdigit() and controle == 0:
                                    controle = 1
                                 if c.isdigit() and controle == 1:
                                    num = num + c
                                 if not c.isdigit() and controle == 1:
                                    break
                            NÚMERO = num
                            #caso não encontre o número
                            if len(NÚMERO) < 1:
                                NÚMERO = "Não Encontrado"
                            #o nome não contém o número da lei
                            FpositionNome = string.find("Nº") - 1
                            NOME = string[1:FpositionNome]
                            #formatação do nome (remoção de espaços e adição de acentos)
                            NOME = NOME.replace("RETIFICAÇ", "RETIFICAÇÃO")
                            NOME = NOME.replace("            ATO", "ATO")
                            NOME = NOME.replace("ACÓRD", "ACORDÃO")

                #exibição de todos os dados
                print(f"Nome: {NOME}\nCategoria: {CATEGORIA}\nOrgão Emissor: {ORGÃO_EMISSOR}\nÂmbito Primário: Federal\nNúmero: {NÚMERO}\nData de publicação: {DATA_INÍCIO}\nData Início Vigência: {DATA_INÍCIO}\nData Final de Vigência: {DATA_FINAL}\nTexto Alternativo: {TEXTO}\n Rotulo Alternativo: NÃO APLICÁVEL")
                #alocação dos dados no excel em uma nova linha para cada lei
                i = i + 1
                worksheet.write(i,2, NOME)
                worksheet.write(i,3, CATEGORIA)
                worksheet.write(i,4, ORGÃO_EMISSOR)
                worksheet.write(i,5, "Federal")
                worksheet.write(i,6, "Nº "+NÚMERO)
                worksheet.write(i,7, DATA_INÍCIO)
                worksheet.write(i,8, DATA_INÍCIO)
                worksheet.write(i,9, DATA_FINAL)
                worksheet.write(i,10, TEXTO)
                worksheet.write(i,11, "NÃO APLICÁVEL")
            #caso o programa não consiga localizar os elementos necessários
            except:
                #notificação da falha
                print(f"O arquivo {name} não pode ser interpretado.")
                FALHA.append(name)
                #renomeação do arquivo para facilitar a identificação
                os.rename(name, name+".old")
            #exibe todos os arquivos que não foram verificados
            print(f"Os seguintes arquivos não foram indexados. {FALHA}")
    #fecha a tabela do excel
    workbook.close()

def loginn():
    try:
        #logar no inlabs
        response = s.request("POST", url_login, data=payload, headers=headers, verify=False) 
        #download e extração dos arquivos .zip
        download()
    #caso ocorra erro de conexão o programa tentará novamente
    except requests.exceptions.ConnectionError:
        loginn()

loginn()
processar()