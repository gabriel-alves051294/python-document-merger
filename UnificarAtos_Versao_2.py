# -*- coding: utf-8 -*-

import os
import json
import subprocess
from tqdm import tqdm
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.table import _Cell, Table

# --- CONFIGURAÇÕES IMPORTANTES ---
PASTA_DE_ENTRADA = r'C:\ProcessarAtos\Entrada'
ARQUIVO_DE_SAIDA_TXT = r'C:\ProcessarAtos\Saida\Atos_Unificados.txt'
ARQUIVO_DE_SAIDA_JSONL = r'C:\ProcessarAtos\Saida\Atos_Unificados.jsonl'
ARQUIVO_DE_LOG_ERROS = r'C:\ProcessarAtos\erros.log'
# --- FIM DAS CONFIGURAÇÕES ---

def converter_doc_para_docx(doc_path, log_erros_file):
    """
    Usa o LibreOffice para converter um arquivo .doc para .docx.
    """
    try:
        # Cria uma subpasta para os arquivos convertidos para manter a organização
        output_dir = os.path.join(os.path.dirname(doc_path), 'convertidos_docx')
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Comando para o LibreOffice
        cmd = [
            'soffice',
            '--headless', # Roda sem interface gráfica
            '--convert-to', 'docx',
            '--outdir', output_dir,
            doc_path
        ]
        # Executa o comando e aguarda a finalização
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        # Retorna o caminho do novo arquivo .docx
        base_name = os.path.basename(doc_path)
        new_docx_path = os.path.join(output_dir, os.path.splitext(base_name)[0] + '.docx')
        
        if os.path.exists(new_docx_path):
            return new_docx_path
        else:
            raise FileNotFoundError("Arquivo convertido não foi encontrado.")

    except FileNotFoundError:
        # Erro comum se o LibreOffice não estiver no PATH do sistema
        msg = f"ERRO DE CONVERSÃO (.doc): O comando 'soffice' não foi encontrado. Verifique se o LibreOffice está instalado e se a pasta 'program' está no PATH do sistema."
        tqdm.write(msg)
        log_erros_file.write(f"{doc_path} - FALHA NA CONVERSÃO: {msg}\n")
        return None
    except subprocess.CalledProcessError as e:
        # Outros erros durante a conversão
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
    Processa um objeto de parágrafo, iterando sobre seus trechos ('runs')
    e concatenando apenas o texto que NÃO está tachado (strikethrough).
    """
    texto_valido = []
    for run in paragrafo.runs:
        # A propriedade 'strike' (tachado simples) e 'dstrike' (tachado duplo) são verificadas
        if not run.font.strike and not run.font.dstrike:
            texto_valido.append(run.text)
    return "".join(texto_valido)

def iter_block_items(parent):
    """Função auxiliar para iterar sobre parágrafos e tabelas na ordem correta."""
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
    Abre um arquivo .docx e retorna seu texto, ignorando trechos tachados
    e incluindo o conteúdo de tabelas.
    """
    try:
        documento = docx.Document(docx_path)
        texto_completo = []
        
        for block in iter_block_items(documento):
            if isinstance(block, Paragraph):
                texto_limpo = obter_texto_sem_tachado(block)
                texto_completo.append(texto_limpo)
            elif isinstance(block, Table):
                for row in block.rows:
                    celulas_limpas = []
                    for cell in row.cells:
                        texto_da_celula = "\n".join(
                            [obter_texto_sem_tachado(p) for p in cell.paragraphs]
                        )
                        celulas_limpas.append(texto_da_celula)
                    row_text = "\t".join(celulas_limpas)
                    texto_completo.append(row_text)
        
        # Filtra linhas vazias que podem ter sido geradas por parágrafos totalmente tachados
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
    for root, _, files in os.walk(PASTA_DE_ENTRADA):
        # Ignora a pasta 'convertidos_docx' para não processar arquivos duas vezes
        if 'convertidos_docx' in root:
            continue
        for file in files:
            # Agora busca por .doc e .docx
            if file.lower().endswith(('.doc', '.docx')) and not file.startswith('~'):
                arquivos_para_processar.append(os.path.join(root, file))

    if not arquivos_para_processar:
        print(f"Nenhum arquivo .doc ou .docx encontrado em '{PASTA_DE_ENTRADA}'.")
        return
        
    print(f"Total de arquivos encontrados: {len(arquivos_para_processar)}")

    arquivos_com_erro = 0
    with open(ARQUIVO_DE_SAIDA_TXT, 'w', encoding='utf-8') as f_txt, \
         open(ARQUIVO_DE_SAIDA_JSONL, 'w', encoding='utf-8') as f_jsonl, \
         open(ARQUIVO_DE_LOG_ERROS, 'w', encoding='utf-8') as f_log:

        f_log.write("Arquivos que falharam durante o processamento:\n")

        for file_path in tqdm(arquivos_para_processar, desc="Processando arquivos"):
            path_para_extrair = file_path
            
            # Se for um arquivo .doc, converte primeiro
            if file_path.lower().endswith('.doc'):
                path_para_extrair = converter_doc_para_docx(file_path, f_log)
            
            # Se a conversão foi bem sucedida (ou se já era .docx)
            if path_para_extrair:
                texto_extraido = extrair_texto_de_docx(path_para_extrair, f_log)
                if texto_extraido and texto_extraido.strip(): # Garante que não está vazio
                    nome_original = os.path.basename(file_path)
                    f_txt.write(f"--- INÍCIO DO DOCUMENTO: {nome_original} ---\n\n{texto_extraido}\n\n--- FIM DO DOCUMENTO: {nome_original} ---\n\n")
                    ato_data = {"fonte": nome_original, "conteudo": texto_extraido}
                    f_jsonl.write(json.dumps(ato_data, ensure_ascii=False) + '\n')
                else:
                    arquivos_com_erro += 1
            else:
                arquivos_com_erro += 1
        
    print(f"\n--- Processo Concluído ---")
    print(f"Arquivo de texto simples salvo em: {ARQUIVO_DE_SAIDA_TXT}")
    print(f"Arquivo JSON Lines (.jsonl) salvo em: {ARQUIVO_DE_SAIDA_JSONL}")
    if arquivos_com_erro > 0:
        print(f"Atenção: {arquivos_com_erro} arquivo(s) não puderam ser processados ou estavam vazios.")
        print(f"Consulte o relatório de erros em: {ARQUIVO_DE_LOG_ERROS}")
    else:
        print("Todos os arquivos foram processados com sucesso.")

if __name__ == "__main__":
    main()
