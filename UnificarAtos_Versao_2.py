# -*- coding: utf-8 -*-

import os
import json
from tqdm import tqdm
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.table import _Cell, Table

# --- CONFIGURAÇÕES IMPORTANTES ---
PASTA_DE_ENTRADA = r'C:\Caminho\Para\Sua\Pasta\Com_Os_Atos'
ARQUIVO_DE_SAIDA_TXT = r'C:\Caminho\Para\Salvar\Atos_Unificados.txt'
ARQUIVO_DE_SAIDA_JSONL = r'C:\Caminho\Para\Salvar\Atos_Unificados.jsonl'
ARQUIVO_DE_LOG_ERROS = r'C:\Caminho\Para\Salvar\erros.log'
# --- FIM DAS CONFIGURAÇÕES ---

# NOVA FUNÇÃO PARA PROCESSAR PARÁGRAFOS E REMOVER TEXTO TACHADO
def obter_texto_sem_tachado(paragrafo):
    """
    Processa um objeto de parágrafo, iterando sobre seus trechos ('runs')
    e concatenando apenas o texto que NÃO está tachado.
    """
    texto_valido = []
    for run in paragrafo.runs:
        # A propriedade 'strike' é True se o texto estiver tachado
        if not run.font.strike:
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
                # Usa a nova função para pegar o texto limpo
                texto_limpo = obter_texto_sem_tachado(block)
                texto_completo.append(texto_limpo)
            elif isinstance(block, Table):
                for row in block.rows:
                    celulas_limpas = []
                    for cell in row.cells:
                        # Processa cada parágrafo dentro da célula para remover texto tachado
                        texto_da_celula = "\n".join(
                            [obter_texto_sem_tachado(p) for p in cell.paragraphs]
                        )
                        celulas_limpas.append(texto_da_celula)
                    # Une o texto de cada célula da linha com um TAB ('\t')
                    row_text = "\t".join(celulas_limpas)
                    texto_completo.append(row_text)
        
        # Filtra linhas vazias que podem ter sido geradas por parágrafos totalmente tachados
        return "\n".join(filter(None, texto_completo))
        
    except Exception as e:
        mensagem_erro = f"ERRO ao ler o arquivo '{os.path.basename(docx_path)}': {e}\n"
        tqdm.write(mensagem_erro)
        log_erros_file.write(f"{docx_path}\n")
        return None

def main():
    """Função principal que orquestra todo o processo."""
    print("Iniciando o processo de unificação de atos normativos (.docx).")

    if not os.path.isdir(PASTA_DE_ENTRADA):
        print(f"ERRO CRÍTICO: A pasta de entrada não foi encontrada: '{PASTA_DE_ENTRADA}'")
        return

    arquivos_docx = []
    for root, _, files in os.walk(PASTA_DE_ENTRADA):
        for file in files:
            if file.lower().endswith('.docx') and not file.startswith('~'):
                arquivos_docx.append(os.path.join(root, file))

    if not arquivos_docx:
        print(f"Nenhum arquivo .docx encontrado em '{PASTA_DE_ENTRADA}'.")
        return
        
    print(f"Total de arquivos .docx encontrados: {len(arquivos_docx)}")

    arquivos_com_erro = 0
    with open(ARQUIVO_DE_SAIDA_TXT, 'w', encoding='utf-8') as f_txt, \
         open(ARQUIVO_DE_SAIDA_JSONL, 'w', encoding='utf-8') as f_jsonl, \
         open(ARQUIVO_DE_LOG_ERROS, 'w', encoding='utf-8') as f_log:

        f_log.write("Arquivos que falharam durante o processamento:\n")

        for file_path in tqdm(arquivos_docx, desc="Processando arquivos"):
            texto_extraido = extrair_texto_de_docx(file_path, f_log)
            
            if texto_extraido:
                nome_arquivo = os.path.basename(file_path)
                
                f_txt.write(f"--- INÍCIO DO DOCUMENTO: {nome_arquivo} ---\n\n")
                f_txt.write(texto_extraido)
                f_txt.write(f"\n\n--- FIM DO DOCUMENTO: {nome_arquivo} ---\n\n")
                
                ato_data = { "fonte": nome_arquivo, "conteudo": texto_extraido }
                f_jsonl.write(json.dumps(ato_data, ensure_ascii=False) + '\n')
            else:
                arquivos_com_erro += 1
        
    print(f"\n--- Processo Concluído ---")
    print(f"Arquivo de texto simples salvo em: {ARQUIVO_DE_SAIDA_TXT}")
    print(f"Arquivo JSON Lines (.jsonl) salvo em: {ARQUIVO_DE_SAIDA_JSONL}")
    if arquivos_com_erro > 0:
        print(f"Atenção: {arquivos_com_erro} arquivo(s) não puderam ser processados.")
        print(f"Consulte o relatório de erros em: {ARQUIVO_DE_LOG_ERROS}")
    else:
        print("Todos os arquivos foram processados com sucesso.")

if __name__ == "__main__":
    main()
