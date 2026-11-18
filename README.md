# Eztools

Automação em Python para acessar a plataforma Eztools, exportar relatórios em Excel e executar scripts de tratamento dos arquivos.

## O que o projeto faz

O script `executoreztools.py`:

1. Acessa a Eztools pelo Google Chrome e realiza o login.
2. Atualiza e exporta os relatórios de **Departamento Pessoal** e **Financeiro**.
3. Aguarda o término do download e verifica se o arquivo Excel contém dados.
4. Move os relatórios baixados para os destinos configurados.
5. Faz backup dos arquivos tratados anteriores e executa os respectivos scripts de tratamento.
6. Registra a execução em arquivos de log.

## Pré-requisitos

- Windows com Python instalado.
- Google Chrome instalado.
- Acesso à Eztools e às pastas de rede configuradas no script.
- Os scripts de tratamento dos relatórios presentes nos caminhos esperados.

## Configuração

Crie um arquivo `.env` com as credenciais e configurações necessárias:

`env`
EZTOOLS_URL=https://endereco-da-eztools
EZTOOLS_USER=seu_usuario
EZTOOLS_PASSWORD=sua_senha
DOWNLOAD_DIR=C:\caminho\para\downloads

`No código atual, o arquivo .env é carregado deste caminho:
text
M:\Business Intelligence\SCRIPTS PYTHON\.env`
Ajuste esse caminho se o arquivo estiver em outro local
Além disso, o script usa caminhos fixos para salvar os relatórios, os backups e localizar os scripts de tratamento. 
Antes de executar em outro computador, revise e ajuste esses caminhos no arquivo executoreztools.py

## Instalação das dependências

No terminal, dentro da pasta do projeto, execute:

`bash
pip install pandas python-dotenv selenium webdriver-manager openpyxl`

## Como executar
Na pasta do projeto, execute:
`python executoreztools.py`

## Arquivos de log
Os logs são gravados na pasta logs, criada junto ao projeto. Eles registram informações da execução e eventuais erros.
