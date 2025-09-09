# -*- coding: utf-8 -*-

import os
import subprocess
from tqdm import tqdm
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.table import _Cell, Table

# --- CONFIGURAÇÕES IMPORTANTES ---
PASTA_DE_ENTRADA = r'C:\ProcessarAtos\Entrada'
# Nome base para os arquivos de saída de texto. A numeração será adicionada automaticamente.
ARQUIVO_DE_SAIDA_TXT_BASE = r'C:\ProcessarAtos\Saida\Atos_Unificados'
ARQUIVO_DE_LOG_ERROS = r'C:\ProcessarAtos\erros.log'
# Caminho completo para o executável do LibreOffice
CAMINHO_SOFFICE = r'C:\Program Files\LibreOffice\program\soffice.exe'
# NOVO: Limite máximo de tamanho para cada arquivo .txt de saída em Megabytes
MAX_TAMANHO_TXT_MB = 2
MAX_TAMANHO_TXT_BYTES = MAX_TAMANHO_TXT_MB * 1024 * 1024


# --- FIM DAS CONFIGURAÇÕES ---

def converter_doc_para_docx(doc_path, log_erros_file):
    """
    Usa o LibreOffice para converter um arquivo .doc para .docx.
    """
    try:
        output_dir = os.path.join(os.path.dirname(doc_path), 'convertidos_docx')
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        cmd = [
            CAMINHO_SOFFICE,
            '--headless',
            '--convert-to', 'docx',
            '--outdir', output_dir,
            doc_path
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        base_name = os.path.basename(doc_path)
        new_docx_path = os.path.join(output_dir, os.path.splitext(base_name)[0] + '.docx')

        if os.path.exists(new_docx_path):
            return new_docx_path
        else:
            raise FileNotFoundError("Arquivo convertido não foi encontrado após a execução do LibreOffice.")

    except FileNotFoundError:
        msg = f"ERRO DE CONVERSÃO (.doc): O executável 'soffice.exe' não foi encontrado no caminho: {CAMINHO_SOFFICE}."
        tqdm.write(msg)
        log_erros_file.write(f"{doc_path} - FALHA NA CONVERSÃO: {msg}\n")
        return None
    except subprocess.CalledProcessError as e:
        msg = f"ERRO DE CONVERSÃO (.doc) para o arquivo '{os.path.basename(doc_path)}': {e.stderr.decode('utf-8', errors='ignore')}"
        tqdm.write(msg)
        log_erros_file.write(f"{doc_path} - FALHA NA CONVERSÃO: {msg}\n")
        return None
    except Exception as e:
        msg = f"ERRO INESPERADO na conversão de '{os.path.basename(doc_path)}': {e}"
        tqdm.write(msg)
        log_erros_file.write(f"{doc_path} - FALHA NA CONVERSÃO: {msg}\n")
        return None


def obter_texto_sem_tachado(paragrafo):
    """
    Concatena o texto de trechos ('runs') de um parágrafo que não estão tachados.
    """
    texto_valido = []
    for run in paragrafo.runs:
        if not run.font.strike:
            texto_valido.append(run.text)
    return "".join(texto_valido)


def iter_block_items(parent):
    """Itera sobre parágrafos e tabelas na ordem correta dentro de um elemento."""
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Tipo de 'parent' não suportado")
    for child in parent_elm.iterchildren():
        if isinstance(child, docx.oxml.text.paragraph.CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, docx.oxml.table.CT_Tbl):
            yield docx.table.Table(child, parent)


def extrair_texto_de_docx(docx_path, log_erros_file):
    """
    Extrai texto de um arquivo .docx, ignorando trechos tachados e lendo tabelas.
    """
    try:
        documento = docx.Document(docx_path)
        texto_completo = []
        for block in iter_block_items(documento):
            if isinstance(block, Paragraph):
                texto_completo.append(obter_texto_sem_tachado(block))
            elif isinstance(block, Table):
                for row in block.rows:
                    celulas_limpas = ["\n".join([obter_texto_sem_tachado(p) for p in cell.paragraphs]) for cell in
                                      row.cells]
                    texto_completo.append("\t".join(celulas_limpas))
        return "\n".join(filter(None, texto_completo))
    except Exception as e:
        msg = f"ERRO ao ler o arquivo '{os.path.basename(docx_path)}': {e}"
        tqdm.write(msg)
        log_erros_file.write(f"{docx_path} - FALHA NA LEITURA: {msg}\n")
        return None


def main():
    """Função principal que orquestra todo o processo."""
    print("Iniciando o processo de unificação de atos normativos (.doc e .docx).")

    if not os.path.isdir(PASTA_DE_ENTRADA):
        print(f"ERRO CRÍTICO: A pasta de entrada não foi encontrada: '{PASTA_DE_ENTRADA}'")
        return

    arquivos_para_processar = []
    print("Buscando arquivos...")
    for root, _, files in os.walk(PASTA_DE_ENTRADA):
        if os.path.basename(root) == 'convertidos_docx':
            continue
        for file in files:
            if file.lower().endswith(('.doc', '.docx')) and not file.startswith('~'):
                arquivos_para_processar.append(os.path.join(root, file))

    if not arquivos_para_processar:
        print(f"Nenhum arquivo .doc ou .docx encontrado em '{PASTA_DE_ENTRADA}'. Verifique a pasta e as permissões.")
        return

    print(f"Total de arquivos encontrados: {len(arquivos_para_processar)}")

    arquivos_com_erro = 0
    arquivos_txt_gerados = []

    # --- Gerenciamento do arquivo TXT ---
    arquivo_txt_indice = 1
    nome_arquivo_txt_atual = f"{ARQUIVO_DE_SAIDA_TXT_BASE}_{arquivo_txt_indice}.txt"
    arquivos_txt_gerados.append(nome_arquivo_txt_atual)
    f_txt = open(nome_arquivo_txt_atual, 'w', encoding='utf-8')

    with open(ARQUIVO_DE_LOG_ERROS, 'w', encoding='utf-8') as f_log:

        f_log.write("Arquivos que falharam ou foram ignorados durante o processamento:\n")

        for file_path in tqdm(arquivos_para_processar, desc="Processando arquivos"):
            path_para_extrair = file_path

            if file_path.lower().endswith('.doc'):
                path_para_extrair = converter_doc_para_docx(file_path, f_log)

            if path_para_extrair:
                texto_extraido = extrair_texto_de_docx(path_para_extrair, f_log)
                if texto_extraido and texto_extraido.strip():
                    nome_original = os.path.basename(file_path)

                    # Prepara o conteúdo a ser escrito
                    conteudo_para_escrever = (
                        f"--- INÍCIO DO DOCUMENTO: {nome_original} ---\n\n"
                        f"{texto_extraido}\n\n"
                        f"--- FIM DO DOCUMENTO: {nome_original} ---\n\n"
                    )
                    tamanho_conteudo_bytes = len(conteudo_para_escrever.encode('utf-8'))

                    # ALTERAÇÃO: Verifica o tamanho do arquivo antes de escrever
                    if f_txt.tell() + tamanho_conteudo_bytes > MAX_TAMANHO_TXT_BYTES and f_txt.tell() > 0:
                        f_txt.close()
                        arquivo_txt_indice += 1
                        nome_arquivo_txt_atual = f"{ARQUIVO_DE_SAIDA_TXT_BASE}_{arquivo_txt_indice}.txt"
                        f_txt = open(nome_arquivo_txt_atual, 'w', encoding='utf-8')
                        arquivos_txt_gerados.append(nome_arquivo_txt_atual)
                        tqdm.write(
                            f"\nLimite de {MAX_TAMANHO_TXT_MB}MB atingido. Criando novo arquivo: {nome_arquivo_txt_atual}")

                    # Escreve no arquivo TXT
                    f_txt.write(conteudo_para_escrever)
                else:
                    if texto_extraido is not None:
                        f_log.write(f"{file_path} - ARQUIVO IGNORADO: Conteúdo vazio ou apenas com texto revogado.\n")
                    arquivos_com_erro += 1
            else:
                arquivos_com_erro += 1

    f_txt.close()  # Garante que o último arquivo de texto seja fechado

    print(f"\n--- Processo Concluido ---")
    # MENSAGEM AJUSTADA para mostrar todos os arquivos gerados
    print(f"Arquivos de texto salvos em:")
    for nome_arquivo in arquivos_txt_gerados:
        print(f"- {nome_arquivo}")

    if arquivos_com_erro > 0:
        print(f"Atenção: {arquivos_com_erro} arquivo(s) não puderam ser processados (por falha ou por estarem vazios).")
        print(f"Consulte o relatório de detalhes em: {ARQUIVO_DE_LOG_ERROS}")
    else:
        print("Todos os arquivos foram processados com sucesso.")


if __name__ == "__main__":
    main()
