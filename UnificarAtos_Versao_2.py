# -*- coding: utf-8 -*-

import os
import json
from tqdm import tqdm
import docx
from docx.document import Document
from docx.table import _Cell, Table

# --- CONFIGURAÇÕES IMPORTANTES ---
PASTA_DE_ENTRADA = r'C:\Caminho\Para\Sua\Pasta\Com_Os_Atos'
ARQUIVO_DE_SAIDA_TXT = r'C:\Caminho\Para\Salvar\Atos_Unificados.txt'
# NOVO: Alterado para .jsonl para mais segurança e eficiência de memória
ARQUIVO_DE_SAIDA_JSONL = r'C:\Caminho\Para\Salvar\Atos_Unificados.jsonl' 
# NOVO: Arquivo de log para registrar arquivos com erro
ARQUIVO_DE_LOG_ERROS = r'C:\Caminho\Para\Salvar\erros.log'
# --- FIM DAS CONFIGURAÇÕES ---

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
            yield docx.text.paragraph.Paragraph(child, parent)
        elif isinstance(child, docx.oxml.table.CT_Tbl):
            yield docx.table.Table(child, parent)

def extrair_texto_de_docx(docx_path, log_erros_file):
    """
    Abre um arquivo .docx e retorna todo o seu texto, incluindo tabelas.
    NOVO: Registra erros em um arquivo de log.
    """
    try:
        documento = docx.Document(docx_path)
        texto_completo = []
        
        for block in iter_block_items(documento):
            if isinstance(block, docx.text.paragraph.Paragraph):
                texto_completo.append(block.text)
            elif isinstance(block, docx.table.Table):
                for row in block.rows:
                    row_text = "\t".join(cell.text for cell in row.cells)
                    texto_completo.append(row_text)
        
        return "\n".join(texto_completo)
        
    except Exception as e:
        # NOVO: Registra o erro no arquivo de log além de avisar na tela.
        mensagem_erro = f"ERRO ao ler o arquivo '{os.path.basename(docx_path)}': {e}\n"
        tqdm.write(mensagem_erro)
        log_erros_file.write(f"{docx_path}\n") # Escreve o caminho do arquivo problemático no log
        return None

def main():
    """Função principal que orquestra todo o processo."""
    print("Iniciando o processo de unificação de atos normativos (.docx).")

    # NOVO: Verificação de segurança para a pasta de entrada
    if not os.path.isdir(PASTA_DE_ENTRADA):
        print(f"ERRO CRÍTICO: A pasta de entrada não foi encontrada no caminho especificado: '{PASTA_DE_ENTRADA}'")
        print("Por favor, verifique o caminho e tente novamente.")
        return # Para a execução

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
    # Abre os três arquivos de saída de uma vez
    with open(ARQUIVO_DE_SAIDA_TXT, 'w', encoding='utf-8') as f_txt, \
         open(ARQUIVO_DE_SAIDA_JSONL, 'w', encoding='utf-8') as f_jsonl, \
         open(ARQUIVO_DE_LOG_ERROS, 'w', encoding='utf-8') as f_log:

        f_log.write("Arquivos que falharam durante o processamento:\n")

        for file_path in tqdm(arquivos_docx, desc="Processando arquivos"):
            texto_extraido = extrair_texto_de_docx(file_path, f_log)
            
            if texto_extraido:
                nome_arquivo = os.path.basename(file_path)
                
                # --- Lógica para o arquivo .TXT ---
                f_txt.write(f"--- INÍCIO DO DOCUMENTO: {nome_arquivo} ---\n\n")
                f_txt.write(texto_extraido)
                f_txt.write(f"\n\n--- FIM DO DOCUMENTO: {nome_arquivo} ---\n\n")
                
                # --- Lógica para o arquivo .JSONL (JSON Lines) ---
                ato_data = { "fonte": nome_arquivo, "conteudo": texto_extraido }
                # Converte o dicionário para uma string JSON e adiciona uma quebra de linha
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
