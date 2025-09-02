# **Unificador de Atos Normativos do TJMG**

---
**(English Summary)**

This project contains a Python script designed to merge thousands of `.doc` and `.docx` files into single, clean text or JSONL files. It leverages LibreOffice for reliable `.doc` conversion and includes features like table extraction and removal of strikethrough text, making it ideal for preparing large document archives for AI training or data analysis.

---

Este projeto contém um script em Python desenvolvido para automatizar a tarefa de unificar dezenas de milhares de arquivos de atos normativos (`.doc` e `.docx`) em arquivos únicos e estruturados, prontos para serem utilizados como base de conhecimento para ferramentas de Inteligência Artificial.

A solução foi criada para ser robusta, segura (rodando 100% localmente) e capaz de lidar com as particularidades de documentos legados, garantindo a máxima qualidade e integridade dos dados extraídos.

## ✨ Funcionalidades Principais

  * **Compatibilidade Dupla:** Processa nativamente arquivos modernos (`.docx`) e antigos (`.doc`).
  * **Conversão Confiável:** Utiliza o LibreOffice para realizar a conversão de `.doc` para `.docx`, garantindo a máxima fidelidade na extração de conteúdo, incluindo tabelas.
  * **Limpeza Inteligente de Dados:** Identifica e remove automaticamente textos marcados como "tachado" (strikethrough), garantindo que atos ou trechos revogados não contaminem a base de conhecimento.
  * **Extração Completa:** É capaz de ler e extrair corretamente o texto do corpo dos documentos e também de dentro de tabelas complexas (Anexos).
  * **Saída Dupla:** Gera dois formatos de arquivo unificado:
      * `.txt`: Ideal para leitura humana e buscas simples.
      * `.jsonl`: Estruturado e perfeito para ser consumido por sistemas e modelos de IA.
  * **Relatório de Erros:** Cria um log detalhado (`erros.log`) listando todos os arquivos que não puderam ser processados e o motivo da falha.
  * **Cache de Conversão:** Salva os arquivos `.doc` já convertidos em uma subpasta (`convertidos_docx`), otimizando futuras execuções.

## 🛠️ Pré-requisitos

Antes de executar o script, garanta que você tenha os seguintes softwares instalados:

1.  **Python 3.x:** A linguagem de programação do script.

      * [Download Oficial do Python](https://www.python.org/)
      * **Importante:** Durante a instalação no Windows, marque a caixa "Add Python to PATH".

2.  **LibreOffice:** O pacote de escritório gratuito utilizado para a conversão segura dos arquivos `.doc`.

      * [Download Oficial do LibreOffice](https://pt-br.libreoffice.org/baixe-ja/libreoffice-novo/)

## 🚀 Instalação e Configuração

Siga os passos abaixo para preparar o ambiente e rodar o projeto.

**1. Clone ou Baixe o Repositório**

Você pode clonar o repositório usando Git:

```bash
git clone https://github.com/gabriel-alves051294/unificacao_docx.git
cd unificacao_docx
```

Ou simplesmente baixe o projeto como um arquivo ZIP e extraia-o em uma pasta no seu computador.

**2. Crie um Ambiente Virtual (Recomendado)**

É uma boa prática criar um ambiente isolado para as dependências do projeto.

```bash
python -m venv venv
# No Windows
.\venv\Scripts\activate
# No macOS/Linux
source venv/bin/activate
```

**3. Instale as Bibliotecas Necessárias**

O script depende de duas bibliotecas Python. Instale-as usando o terminal:

```
pip install python-docx tqdm
```

**4. Configure o Acesso ao LibreOffice**

O script precisa saber onde encontrar o LibreOffice. A maneira mais robusta é adicionar a pasta de instalação ao PATH do seu sistema operacional.

  * **Caminho padrão no Windows:** `C:\Program Files\LibreOffice\program`

Siga as instruções do manual que criamos para adicionar este caminho às suas variáveis de ambiente.

## 🏃‍♂️ Como Usar

1.  **Prepare a Estrutura de Pastas:**

      * Crie uma estrutura de pastas no seu computador, por exemplo:
          * `C:\ProcessarAtos`
          * `C:\ProcessarAtos\Entrada`
          * `C:\ProcessarAtos\Saida`

2.  **Configure o Script:**

      * Abra o arquivo `UnificarAtos_Versao_2.py` em um editor de código (como o PyCharm ou VS Code).
      * Verifique e, se necessário, ajuste as variáveis de configuração no topo do arquivo para que correspondam às pastas que você criou.

3.  **Adicione os Arquivos:**

      * Copie todos os arquivos `.doc` e `.docx` que você deseja unificar para a pasta `C:\ProcessarAtos\Entrada`.

4.  **Execute o Script:**

      * Abra um terminal, navegue até a pasta do projeto (onde o arquivo `.py` está salvo) e execute o comando:

    <!-- end list -->

    ```bash
    python UnificarAtos_Versao_2.py
    ```

5.  **Aguarde a Conclusão:**

      * Uma barra de progresso mostrará o andamento. O processo pode levar várias horas, dependendo da quantidade de arquivos `.doc`.

## 📄 Entendendo os Arquivos de Saída

Ao final do processo, você encontrará os seguintes arquivos:

  * **`Atos_Unificados.txt`**: Um único arquivo de texto contendo o conteúdo limpo de todos os documentos, separados por um marcador. Ótimo para leitura e busca por palavras-chave.
  * **`Atos_Unificados.jsonl`**: Um arquivo de texto onde cada linha é um objeto JSON contendo a fonte (`fonte`) e o conteúdo (`conteudo`) de um ato. Este é o formato ideal para importação em bancos de dados ou para treinamento de modelos de IA.
  * **`erros.log`**: Um relatório que lista todos os arquivos que falharam durante a conversão ou leitura, ajudando a identificar documentos problemáticos.
  * **Pasta `convertidos_docx`**: Criada dentro da sua pasta de entrada, armazena as versões `.docx` dos arquivos `.doc` antigos, servindo como um cache que pode acelerar futuras execuções.

-----
