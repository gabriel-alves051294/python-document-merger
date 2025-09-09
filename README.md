# Processador e Unificador de Atos Normativos

Este script em Python foi desenvolvido para automatizar o processo de consolidação de um grande volume de documentos de texto (.doc e .docx), como atos normativos, portarias, leis, etc. A ferramenta extrai o conteúdo textual, ignorando trechos revogados (tachados), e unifica tudo em arquivos de texto (.txt) com um tamanho máximo controlado, facilitando o manuseio e a análise posterior.

## Funcionalidades Principais

-   **Processamento em Lote:** Varre recursivamente uma pasta de entrada e processa todos os arquivos `.doc` e `.docx` que encontrar.
-   **Conversão Automática:** Utiliza o LibreOffice em modo *headless* (sem interface gráfica) para converter arquivos do formato antigo `.doc` para o formato moderno `.docx` de forma transparente.
-   **Extração Inteligente de Conteúdo:**
    -   Lê o conteúdo de parágrafos e tabelas.
    -   Ignora de forma inteligente qualquer texto que esteja formatado como tachado (*strikethrough*), que é comumente usado para indicar trechos revogados.
-   **Divisão de Arquivos por Tamanho:** Consolida o texto extraído em arquivos `.txt`. Quando um arquivo de saída atinge um limite de tamanho configurável (ex: 2MB), o script automaticamente cria um novo arquivo para continuar o processo (ex: `Atos_Unificados_1.txt`, `Atos_Unificados_2.txt`, etc.).
-   **Log de Erros Detalhado:** Cria um arquivo de log (`erros.log`) que registra qualquer falha durante a conversão ou leitura de arquivos, informando qual documento apresentou problema e o motivo, facilitando a depuração.
-   **Barra de Progresso:** Exibe uma barra de progresso (`tqdm`) para que o usuário possa acompanhar o andamento do processamento, especialmente útil para um grande número de arquivos.

## Caso de Uso

Esta ferramenta é ideal para quem precisa:

-   Criar um *corpus* textual a partir de milhares de documentos para projetos de Processamento de Linguagem Natural (NLP).
-   Consolidar uma base de conhecimento dispersa em vários arquivos para facilitar a busca e a consulta.
-   Preparar documentos para importação em sistemas de gestão de conteúdo ou bases de dados.
-   Arquivar de forma organizada o conteúdo de atos normativos, mantendo apenas o texto vigente.

## Pré-requisitos

Para executar o script, você precisará ter o seguinte instalado em seu sistema:

1.  **Python 3.6+**
2.  **LibreOffice:** A suíte de escritório é necessária para a conversão de arquivos `.doc`.
    -   Você pode baixar em [LibreOffice.org](https://www.libreoffice.org/download/download/).
3.  **Bibliotecas Python:** Instale as dependências com o seguinte comando:

    ```bash
    pip install python-docx tqdm
    ```

## Configuração

Antes de executar, você **precisa** ajustar as constantes no início do script `processador_atos.py`:

```python
# --- CONFIGURAÇÕES IMPORTANTES ---

# 1. Pasta onde estão os seus arquivos .doc e .docx
PASTA_DE_ENTRADA = r'C:\ProcessarAtos\Entrada'

# 2. Caminho e nome base para os arquivos de texto que serão gerados
ARQUIVO_DE_SAIDA_TXT_BASE = r'C:\ProcessarAtos\Saida\Atos_Unificados'

# 3. Caminho para o arquivo de log de erros
ARQUIVO_DE_LOG_ERROS = r'C:\ProcessarAtos\erros.log'

# 4. Caminho COMPLETO para o executável do LibreOffice
#    (Verifique onde ele foi instalado no seu sistema)
CAMINHO_SOFFICE = r'C:\Program Files\LibreOffice\program\soffice.exe'

# 5. Limite máximo de tamanho (em Megabytes) para cada arquivo .txt gerado
MAX_TAMANHO_TXT_MB = 2
