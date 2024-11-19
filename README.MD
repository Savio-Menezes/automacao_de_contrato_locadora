# Automação de Contratos em Massa para Locadora de Veículos

Este projeto visa automatizar a geração de contratos individuais para uma locadora de veículos a partir de uma planilha Excel. Utilizando **Python**, **Selenium** e **python-docx**, a automação preenche automaticamente um modelo de contrato com os dados do cliente e os salva em um arquivo Word personalizado.

## Funcionalidade

- **Entrada**: Uma planilha Excel contendo os dados dos clientes.
- **Saída**: Um contrato gerado e salvo como documento Word para cada cliente, com o nome do cliente como nome do arquivo.
- **Objetivo**: Facilitar a geração de contratos personalizados em massa sem a necessidade de sistemas caros e complexos.

## Tecnologias Utilizadas

- **Python**: Linguagem de programação para desenvolver o script de automação.
- **Selenium**: Biblioteca utilizada para automação de tarefas no navegador (se necessário para interagir com o conteúdo da planilha ou outros dados).
- **python-docx**: Biblioteca utilizada para criar e editar documentos do Word (geração do contrato).
- **openpyxl**: Biblioteca para ler e manipular arquivos Excel (.xlsx).

## Como Funciona

### 1. Preparação dos Dados
O projeto começa com uma planilha Excel contendo os seguintes dados (exemplo):
- Nome
- CPF
- Endereço
- Data de Nascimento
- Telefone de Contato
- E-mail
- Tempo de Locação

### 2. Preenchendo o Modelo de Contrato
O script utiliza os dados da planilha para preencher um modelo de contrato. O modelo de contrato inclui placeholders que são substituídos pelos dados dos clientes.

### 3. Geração dos Contratos
Cada cliente tem um contrato gerado com seus dados e salvo em um arquivo Word com o nome do cliente. Todos os contratos gerados são salvos em uma pasta específica.

## Como Usar

### Requisitos

- Python 3.x
- Bibliotecas necessárias:
  - `openpyxl`
  - `python-docx`
  - `selenium`

Você pode instalar as dependências usando o `pip`:

```bash
pip install openpyxl python-docx selenium


Exemplo de Uso
Planilha de Dados
A planilha dados_clientes_contratos.xlsx deve ter as seguintes colunas (semelhantes a este formato):

Nome	CPF	Endereço	Data de Nascimento	Telefone	E-mail	Tempo de Locação
João Silva	123.456.789-00	Rua Exemplo, 123	01/01/1980	987654321	joao@email.com	30