# **Unificador de Documentos Word (.doc e .docx)**

**(English Summary)**

A Python tool designed to batch merge a large volume of Word documents (`.doc` and `.docx`) into single, clean text or JSONL files. While this tool is content-agnostic and can process any type of document, it was developed and battle-tested using a large archive of legal documents (normative acts from the TJMG Court of Justice). It's ideal for preparing large document archives for AI training or data analysis.

-----

## 📜 Sobre o Projeto

Este projeto apresenta uma solução em Python para a unificação e o processamento em lote de grandes volumes de documentos Word, nos formatos legados (`.doc`) e modernos (`.docx`).

Embora a ferramenta seja agnóstica ao conteúdo e possa unificar qualquer tipo de documento, ela foi desenvolvida e testada no contexto de um projeto real para consolidar dezenas de milhares de **atos normativos do Tribunal de Justiça de Minas Gerais (TJMG)**.

A solução foi criada para ser robusta e segura (rodando 100% localmente), gerando arquivos de saída limpos e estruturados, ideais para projetos de análise de dados, arquivamento digital ou para a criação de bases de conhecimento para modelos de Inteligência Artificial.

## ✨ Funcionalidades

  * **Suporte a Múltiplos Formatos:** Processa nativamente arquivos `.docx` e `.doc`, garantindo compatibilidade com acervos de documentos mistos.
  * **Conversão Robusta de Arquivos `.doc`:** Utiliza o LibreOffice para realizar a conversão, assegurando a máxima fidelidade na extração de conteúdo, incluindo elementos complexos como tabelas.
  * **Limpeza e Tratamento de Dados:** Identifica e remove automaticamente textos marcados como "tachado" (strikethrough), assegurando que conteúdo revogado não seja incluído na base de dados final.
  * **Extração de Conteúdo Abrangente:** Lê e extrai corretamente o texto do corpo dos documentos e de dentro de tabelas (Anexos).
  * **Múltiplos Formatos de Saída:** Gera dois tipos de arquivo unificado para diferentes finalidades:
      * **`.txt`:** Ideal para leitura humana e buscas textuais simples.
      * **`.jsonl`:** Estruturado para consumo por sistemas, bancos de dados e modelos de IA.
  * **Registro Detalhado de Erros (Logging):** Cria um log (`erros.log`) listando todos os arquivos que falharam durante o processamento e o motivo técnico da falha.
  * **Otimização com Cache de Conversão:** Salva os arquivos `.doc` já convertidos em uma subpasta (`convertidos_docx`), o que pode otimizar significativamente futuras execuções.

## 🛠️ Pré-requisitos

Para a correta execução do script, os seguintes softwares são necessários:

  * **Python 3.x:** A linguagem de programação do script.

      * [Download Oficial do Python](https://www.python.org/)
      * **Importante:** Durante a instalação no Windows, marque a caixa "Add Python to PATH".

  * **LibreOffice:** Pacote de escritório gratuito utilizado para a conversão segura dos arquivos `.doc`.

      * [Download Oficial do LibreOffice](https://pt-br.libreoffice.org/baixe-ja/libreoffice-novo/)

## 🚀 Instalação e Configuração

Siga os passos abaixo para preparar o ambiente e rodar o projeto.

**1. Clone ou Baixe o Repositório**

Utilize o Git para clonar o repositório:

```bash
git clone https://github.com/gabriel-alves051294/unificacao_docx.git
cd unificacao_docx
```

Alternativamente, baixe o projeto como um arquivo ZIP e extraia-o em uma pasta de sua preferência.

**2. Crie um Ambiente Virtual (Recomendado)**

É uma boa prática isolar as dependências do projeto:

```bash
# No Windows
python -m venv venv
.\venv\Scripts\activate

# No macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

**3. Instale as Bibliotecas Necessárias**

Instale as dependências Python via terminal:

```bash
pip install python-docx tqdm
```

**4. Configure o Acesso ao LibreOffice**

O script precisa localizar a instalação do LibreOffice. A maneira mais robusta é adicionar sua pasta de instalação ao PATH do sistema operacional.

  * **Caminho Padrão no Windows:** `C:\Program Files\LibreOffice\program`

## 🏃‍♂️ Como Usar

**1. Prepare a Estrutura de Pastas:**
Crie uma estrutura de pastas para organizar os arquivos. Por exemplo:

```
C:\Meus-Documentos\
├── Entrada\
└── Saida\
```

*(Nota: este é apenas um caminho de exemplo. Você pode criar a estrutura de pastas em qualquer local de sua preferência.)*

**2. Configure o Script:**
Abra o arquivo `UnificarAtos_Versao_2.py` e ajuste as variáveis de configuração no topo do arquivo para que correspondam às pastas que você criou.

**3. Adicione os Arquivos:**
Copie todos os documentos `.doc` e `.docx` a serem unificados para a pasta de `Entrada`.

**4. Execute o Script:**
Navegue até a pasta do projeto via terminal e execute o comando:

```bash
python UnificarAtos_Versao_2.py
```

**5. Aguarde a Conclusão:**
Uma barra de progresso indicará o andamento. O processo pode levar várias horas, dependendo do volume de arquivos `.doc`.

## 📄 Entendendo os Arquivos de Saída

Ao final do processo, você encontrará os seguintes arquivos:

  * **`Atos_Unificados.txt`:** Arquivo de texto puro com o conteúdo limpo de todos os documentos.
  * **`Atos_Unificados.jsonl`:** Arquivo formatado onde cada linha é um objeto JSON contendo a fonte (`fonte`) e o conteúdo (`conteudo`) de um documento.
  * **`erros.log`:** Relatório com a lista de arquivos que apresentaram falhas durante o processo.
  * **Pasta `convertidos_docx`:** Subpasta criada no diretório de entrada, que armazena as versões `.docx` dos arquivos `.doc` processados.
