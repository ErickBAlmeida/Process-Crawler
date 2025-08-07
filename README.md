# BOT Polos Status

Este projeto é um bot de automação desenvolvido em Python para consultar e extrair o status de processos judiciais em um determinado sistema, utilizando Selenium WebDriver e manipulação de planilhas Excel.

## Funcionalidades
- Acessa o sistema automaticamente via navegador Chrome.
- Realiza login com certificado digital.
- Navega até a área de consulta de processos.
- Lê uma lista de números de processos a partir de uma planilha Excel.
- Pesquisa cada processo e extrai o status atual (arquivado, baixado, julgado, etc.).
- Exibe o status no console e pode notificar via Toast no Windows.
- Suporta identificação de processos em segredo de justiça.

## Requisitos
- Python 3.8+
- Google Chrome instalado
- ChromeDriver compatível com a versão do Chrome

## Instalação
Instale as dependências:
```bash
pip install -r requirements.txt
```

## Como usar
Execute o script principal:
```bash
python index.py
```
O bot irá abrir o navegador, realizar o login e processar todos os números de processo da planilha, exibindo o status de cada um.