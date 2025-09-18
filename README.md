# Processador de Atos Normativos

Ferramenta de automação para classificar, converter e consolidar arquivos de atos normativos, otimizada para o fluxo de trabalho do TJMG.

## Descrição

Este projeto automatiza o tratamento de grandes volumes de documentos jurídicos (`.doc` e `.docx`). Ele foi desenvolvido para resolver um problema específico: separar atos normativos válidos dos revogados, padronizá-los para o formato `.docx` e consolidar seu conteúdo textual de forma organizada para análises futuras.

A ferramenta classifica os atos com base na formatação do texto (uso de "tachado" para indicar revogação), garantindo que apenas o conteúdo relevante seja processado e arquivado, replicando a estrutura de pastas original para manter a organização por categoria.

## Recursos Principais

  - **Filtragem Automática:** Identifica e descarta arquivos com mais de 90% do texto tachado, considerados revogados.
  - **Conversão em Lote:** Converte arquivos do formato `.doc` para `.docx` de forma automática, utilizando o LibreOffice.
  - **Estrutura de Pastas Espelhada:** Organiza os arquivos de saída (`.docx` e `.txt`) em uma estrutura de diretórios idêntica à de entrada.
  - **Limpeza de Conteúdo:** Extrai o texto dos documentos, removendo qualquer trecho, palavra ou caractere que esteja formatado como tachado.
  - **Consolidação Inteligente:** Agrupa o conteúdo textual por categoria e o fragmenta em arquivos `.txt` com tamanho máximo de 2MB para facilitar a manipulação.
  - **Logging Detalhado:** Gera um arquivo de log completo (`log_processamento.log`) registrando todas as ações, avisos e erros para fácil auditoria.

## Estrutura de Diretórios

Para que o script funcione corretamente, a seguinte estrutura de pastas deve ser criada na raiz `C:\`:

```
C:\ProcessarAtos\
│
├── Entradas\
│   ├── Categoria_A\
│   │   ├── ato_01.doc
│   │   └── ato_02.docx
│   └── Categoria_B\
│       └── ato_03.doc
│
├── Saida_DOCX\      (criado pelo script)
├── Saida_TXT\       (criado pelo script)
└── log_processamento.log  (criado pelo script)
```

## Pré-requisitos

Antes de executar, certifique-se de que você tem os seguintes softwares instalados:

1.  **Python 3.7+:** [Download Python](https://www.python.org/downloads/)
2.  **LibreOffice:** [Download LibreOffice](https://www.google.com/search?q=https://www.libreoffice.org/download/download/)
      - *Importante: O caminho para o executável do LibreOffice (`soffice.exe`) deve ser verificado e, se necessário, ajustado na variável `CAMINHO_SOFFICE` dentro do script.*

## Instalação

1.  **Clone o repositório:**

    ```bash
    git clone https://github.com/seu-usuario/seu-repositorio.git
    cd seu-repositorio
    ```

2.  **Crie um ambiente virtual (recomendado):**

    ```bash
    python -m venv venv
    venv\Scripts\activate  # No Windows
    ```

3.  **Instale as dependências:**
    O script utiliza as seguintes bibliotecas Python. Instale-as usando `pip`:

    ```bash
    pip install python-docx tqdm
    ```

## Como Usar

1.  **Prepare os arquivos:** Coloque seus arquivos `.doc` e `.docx` dentro das respectivas subpastas de categoria em `C:\ProcessarAtos\Entradas\`.
2.  **Ajuste as configurações:** Abra o arquivo `processador_revogados.py` e verifique se os caminhos nas `CONFIGURAÇÕES` (como `PASTA_ENTRADA` e `CAMINHO_SOFFICE`) correspondem ao seu ambiente.
3.  **Execute o script:** Abra o terminal no diretório do projeto e execute o seguinte comando:
    ```bash
    python processador_revogados.py
    ```
4.  **Verifique os resultados:** Após a execução, os arquivos convertidos estarão em `Saida_DOCX`, os textos consolidados em `Saida_TXT` e o log detalhado em `log_processamento.log`.

## Licença

Este projeto está licenciado sob a Licença MIT. Veja o arquivo `LICENSE` para mais detalhes.
