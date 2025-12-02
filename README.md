# 🛒 Processador de Promoções CRM



Sistema automatizado para processar planilhas de promoções e gerar arquivos formatados para importação no CRM.



## 📋 Descrição



Este projeto é uma aplicação Streamlit que processa planilhas de encartes promocionais, realiza validações, mesclagem de dados de EAN, correção de nomes de produtos e exporta arquivos Excel formatados e prontos para importação no sistema CRM.



## ✨ Funcionalidades



- ✅ Detecção automática de cabeçalhos em planilhas

- ✅ Processamento por perfil de loja (GERAL, PREMIUM, GERAL/PREMIUM)

- ✅ Mesclagem opcional de dados de EAN externos

- ✅ Correção automática de nomes de produtos

- ✅ Integração com repositório de links de imagens

- ✅ Classificação automática de códigos (EAN vs Interno)

- ✅ Validação e destaque visual de campos obrigatórios

- ✅ Exportação para Excel com formatação condicional



## 🚀 Instalação



### Pré-requisitos



- Python 3.8 ou superior

- pip (gerenciador de pacotes Python)



### Passos



1. Clone o repositório:

```bash

git clone <url-do-repositorio>

cd processador-promocoes-crm

```



2. Crie um ambiente virtual (recomendado):

```bash

python -m venv venv

source venv/bin/activate  # Linux/Mac

venv\Scripts\activate     # Windows

```



3. Instale as dependências:

```bash

pip install -r requirements.txt

```



4. Configure os arquivos de dados necessários:

```bash

mkdir -p data

```



## 📁 Estrutura do Projeto



```

encarte_automatizado_online/

│

├── main.py                          # Interface principal Streamlit

├── requirements.txt                 # Dependências do projeto

│

├── data/

│   ├── config.json                  # Configurações do sistema

│   └── default_url.json             # Repositório padrão de links

│

├── src/

│   ├── config/

│   │   ├── __init__.py              # Torna o diretório um pacote Python

│   │   └── config_loader.py         # Carregador de configurações

│   │

│   ├── processors/

│   │   ├── __init__.py              # Torna o diretório um pacote Python

│   │   ├── promotion_processor.py   # Processador principal

│   │   ├── header_detector.py       # Detector de cabeçalhos

│   │   ├── ean_merger.py            # Mesclador de dados EAN

│   │   ├── dataframe_builder.py     # Construtor de DataFrames

│   │   └── excel_exporter.py        # Exportador para Excel

│   │

│   ├── utils/

│   │   ├── __init__.py              # Torna o diretório um pacote Python

│   │   ├── data_utils.py            # Utilitários de dados

│   │   ├── ean_classifier.py        # Classificador de EAN

│   │   ├── file_utils.py            # Utilitários de arquivos

│   │   ├── link_loader.py           # Carregador de links

│   │   └── text_utils.py            # Utilitários de texto

│

├── .gitattributes

├── .gitignore

├── README.md

└── LICENSE

```



## ⚙️ Configuração



### Arquivo `data/config.json`



Estrutura necessária:



```json

{

  "required_columns": [

    "código",

    "ean",

    "descrição do item",

    "preço de:",

    "preço por:",

    "perfil de loja",

    "tipo ação"

  ],

  "buyer_carrossel_map": {

    "compradora de mercearia": "8135 - Mercearia Salgada",

    "comprador de bebidas": "8136 - Bebidas",

    "compradores de higiene": "8137 - Higiene e Beleza",

    "compradora de limpeza": "8138 - Limpeza"

  },

  "product_name_corrections": {

    "\\bfile\\b": "FILÉ",

    "\\bhamb\\b": "HAMBÚRGER",

    "\\bfgo\\b": "FRANGO",

    "\\bespag\\b": "ESPAGUETE",

    "\\blacteo\\b": "LÁCTEO",

    "\\bhig\\b": "HIGIÊNICO",

    "\\bracao\\b": "RAÇÃO",

  }

}

```



### Arquivo `data/default_url.json`



Estrutura para links de imagens:



```json

[

  {

    "url": "https://exemplo.com/imagem1.jpg",

    "eans": ["7891234567890", "7891234567891"]

  }

]

```



## 🎯 Como Usar



1. Inicie a aplicação:

```bash

streamlit run main.py

```



2. Na interface web:

   - Selecione as datas de início e fim do encarte

   - Configure as opções desejadas:

    - ☑️ Aplicar correção de nomes de produtos

    - ☑️ Usar arquivo de EANs

    - ☑️ Usar arquivo JSON de Links

   - Faça upload do arquivo de encarte consolidado

   - Selecione a planilha desejada (se aplicável)

   - Faça upload dos arquivos opcionais (EANs, Links)

   - Clique em "Processar Promoções"



3. Baixe os arquivos gerados para cada perfil



## 📊 Formatos de Entrada



### Arquivo Principal (Encarte Consolidado)

- Formatos aceitos: `.xlsx`, `.xls`, `.csv`

- Deve conter as colunas obrigatórias definidas em `config.json`

- Deve ter uma coluna "tipo ação" contendo "CRM" para as linhas a processar



### Arquivo de EANs (Opcional)

- Formatos aceitos: `.xlsx`, `.xls`, `.csv`

- Deve conter as colunas: `CÓDIGO PRODUTO` e `CÓDIGO EAN`



### Arquivo de Links (Opcional)

- Formato aceito: `.json`

- Estrutura: array de objetos com `url` e `eans`




## 📤 Formato de Saída

Os arquivos Excel gerados por perfil de loja incluem as seguintes colunas:

| Coluna                           | Descrição                      |
| -------------------------------- | -------------------------------- |
| **Nome**                   | Nome do produto                  |
| **Carrossel**              | Categoria do produto             |
| **Check-In**               | Status de check-in               |
| **Preço**                 | Preço original                  |
| **Preço promocional**     | Preço em promoção             |
| **Limite por cliente**     | Limite de compra                 |
| **Dias para Resgate**      | Período de validade             |
| **Unidade**                | Unidade de medida                |
| **Não exigir ativação** | Tipo de ativação               |
| **Ativar em**              | Data/hora de início da oferta   |
| **Inativar em**            | Data/hora de encerramento        |
| **URL da imagem**          | Link da imagem do produto        |
| **Tipo do código**        | Tipo de código (EAN ou interno) |
| **Códigos dos produtos**  | Lista de EANs                    |
| **Tipo Promocional**       | Tipo da promoção               |
| **Sobrescrever lojas**     | Indica se sobrescreve lojas      |
| **Lojas**                  | IDs das lojas                    |

### Formatação Condicional



- 🔴 **Vermelho**: Campos obrigatórios vazios (EAN, Preço, Preço Promocional)

- 🟡 **Amarelo**: Alertas (Unidade = Quilograma, Tipo = Interno)



## 🛠️ Tecnologias Utilizadas



- **Streamlit**: Interface web

- **Pandas**: Manipulação de dados

- **OpenPyXL**: Processamento de arquivos Excel

- **Python 3.8+**: Linguagem base



## 🔍 Regras de Negócio



### Classificação de EAN

- **Interno + Quilograma**:  < 10 dígitos

- **EAN + Unidade**: Caso contrário



### Carrossel Especial

- Produtos com "DESTAQUE CRM" na coluna "tipo ação" recebem automaticamente "8142 - Especial"



### Cópia de Preços

- Se preços estiverem vazios e os 7 primeiros dígitos do EAN coincidirem com a linha anterior, os preços são copiados



### Perfis de Loja

- **GERAL**: Lojas padrão

- **PREMIUM**: Lojas premium

- **GERAL/PREMIUM**: Todas as lojas



## 📝 Dependências



```txt

streamlit>=1.28.0

pandas>=2.0.0

openpyxl>=3.1.0

```



## 🤝 Contribuindo



1. Fork o projeto

2. Crie uma branch para sua feature (`git checkout -b feature/MinhaFeature`)

3. Commit suas mudanças (`git commit -m 'Adiciona MinhaFeature'`)

4. Push para a branch (`git push origin feature/MinhaFeature`)

5. Abra um Pull Request


## 👥 Autores



- **Guilherme Ferreira** – [@Guilherme1ss](https://github.com/Guilherme1ss)



## 📞 Suporte



Para reportar bugs ou solicitar features, abra uma issue no repositório.



---



Desenvolvido com ❤️ para otimizar o processamento de encarte do CRM
