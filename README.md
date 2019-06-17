# WebScraping Valor Acoes
WebScraping para pegar valor de ações da bolsa de valores e atualizar planilha excel.

## Pré-requisitos
bs4 ( BeautifulSoup ) para o Web Scraping 
openpyxl para as operações com o excel (xlsx)

## Uso

import Acoes

carteira = Acoes.Carteira()
cotacao = Acoes.Carteira()

# BIDI4
x = Acoes.Acao('BIDI4', '30/04/2018', 100, 19.82, 2.49)
carteira.addAcao(x.acao,x.getAcao())

carteira = carteira.getCarteira()

# Abre o excel para atualização
arquivo = Acoes.Arquivo_xlsx('ACOES')

# Navega pelas ações da carteira
for acao in carteira:
    arquivo.updAcao(acao,carteira[acao])

# Salva o arquivo
arquivo.saveArquivo('_ATUAL')

print('Done!')
exit(0)


