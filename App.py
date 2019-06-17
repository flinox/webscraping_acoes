
import Acoes

carteira = Acoes.Carteira()
cotacao = Acoes.Carteira()

# KROT3
x = Acoes.Acao('KROT3', '03/11/2017', 100, 17.46, 2.49)
carteira.addAcao(x.acao,x.getAcao())

# BIDI4
x = Acoes.Acao('BIDI4', '30/04/2018', 100, 19.82, 2.49)
carteira.addAcao(x.acao,x.getAcao())

# OIBR4
x = Acoes.Acao('OIBR4', '20/12/2017', 100, 3.72, 2.49)
carteira.addAcao(x.acao,x.getAcao())

# LAME4
x = Acoes.Acao('LAME4', '24/10/2017', 100, 17.95, 0)
carteira.addAcao(x.acao,x.getAcao())

# GOAU4
x = Acoes.Acao('GOAU4', '11/10/2017', 100, 5.64, 2.49)
carteira.addAcao(x.acao,x.getAcao())

# DMMO3
x = Acoes.Acao('DMMO3', '05/10/2017', 100, 1.67, 0)
carteira.addAcao(x.acao,x.getAcao())

# USIM5
x = Acoes.Acao('USIM5', '03/10/2017', 100, 9.43, 0)
carteira.addAcao(x.acao,x.getAcao())

# VVAR11
x = Acoes.Acao('VVAR11', '18/09/2018', 100, 16.80, 0)
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

