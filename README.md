# Hackathon-Ambplan-IMT
Programa elaborado e programado em grupo no legaltech hackathon ambplan da semana smile do IMT

# INTRODUÇÃO DO PROJETO
O projeto tem como objetivo coletar as leis do Diário Oficial, armazenar os arquivos xml, filtrar as informações relevantes e 
exibi-las de forma organizada no excel. Tudo de maneira automática.

# INFORMAÇÕES IMPORTANTES DE PRÉ REQUISITOS
1- Os arquivos xml são individuais por lei, já filtram algumas informações relevantes, são facilmente integrado com python e 
estão compactados em no máximo dois arquivos zip por dia. Por isso são mais práticos que arquivos pdf

2- Os arquivos xml são de fácil acesso por meio da plataforma do governo "inlabs" https://inlabs.in.gov.br, 
que necessita de uma conta para o acesso

3- O projeto foi feito em python e inclui as seguintes bibliotecas:

datetime -- definir o dia do Diário Oficial que será analisado
requests -- acessar e navegar pelo site da inlabs
zipfile -- manipular os arquivos .zip que contém o Diário Oficial
xml.etree.ElementTree -- manipular os arquivos .xml
os, os.path -- gerenciar os arquivos que serão baixados
xlsxwriter -- passar as informações para o excel

# PASSO A PASSO
O projeto tem como objetivo realizar o login necessário para acessar a página do inlabs 
Acessar a pasta do inlabs do dia útil anterior por meio da url padronizada
Baixar os arquivos que contém o diário oficial com os atos legislativos e os atos extras "DO1.zip" e "DO1E.zip"
extrair os arquivos .zip, que contém arquivos xml de todas as leis do dia
organizar as informações já separadas nos arquivos xml e as informações filtradas em um documento excel, preenchendo campos não encontrados com "Não Encontrado"

# DOWNLOADS NECESSÁRIOS PARA O FUNCIONAMENTO DO PROJETO:
python 3.10 -https://www.python.org/downloads/
python requests - https://github.com/psf/requests
python xlswriter - https://github.com/jmcnamara/XlsxWriter
