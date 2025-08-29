import os
import json
import docx  # A biblioteca para ler arquivos .docx

def processar_pasta_docx():
    """
    Lê todos os arquivos .docx de uma pasta especificada e os converte
    para um único arquivo .jsonl, otimizado para IA.
    """
    # --- CONFIGURAÇÃO ---
    # IMPORTANTE: Coloque aqui o caminho exato da pasta onde estão seus arquivos .docx
    pasta_com_docx = r"C:\ProcessarAtos\Entrada"  # <-- MUDE AQUI SE NECESSÁRIO
    
    # Local e nome do arquivo de saída final
    arquivo_saida_jsonl = r"C:\ProcessarAtos\Saida\dados_unificados.jsonl"
    # ---------------------

    print(f">>> Iniciando o processamento da pasta: {pasta_com_docx}")

    # Verifica se a pasta de entrada realmente existe
    if not os.path.isdir(pasta_com_docx):
        print(f"\nERRO FATAL: A pasta de entrada não foi encontrada!")
        print(f"Por favor, verifique se o caminho '{pasta_com_docx}' está correto no script.")
        return

    arquivos_processados = 0
    arquivos_falha = []
    
    # Garante que a pasta de saída exista antes de tentar salvar o arquivo
    os.makedirs(os.path.dirname(arquivo_saida_jsonl), exist_ok=True)

    # Abre o arquivo de saída para escrita
    with open(arquivo_saida_jsonl, 'w', encoding='utf-8') as f_out:
        try:
            # Pega a lista de arquivos e ordena para manter a sequência numérica/alfabética
            lista_arquivos = sorted(os.listdir(pasta_com_docx))
        except FileNotFoundError:
             print(f"\nERRO FATAL: A pasta de entrada não foi encontrada ao tentar listar os arquivos!")
             print(f"Verifique se o caminho '{pasta_com_docx}' está correto.")
             return

        for nome_arquivo in lista_arquivos:
            # Processa apenas arquivos que terminam com .docx
            if nome_arquivo.lower().endswith('.docx'):
                caminho_completo = os.path.join(pasta_com_docx, nome_arquivo)
                
                try:
                    # Abre o documento .docx
                    documento = docx.Document(caminho_completo)
                    
                    # Extrai o texto de todos os parágrafos
                    paragrafos = [p.text for p in documento.paragraphs]
                    texto_completo = "\n".join(paragrafos)
                    
                    # Cria o objeto de dados (dicionário Python) para o JSONL
                    dados = {
                        "source_file": nome_arquivo,
                        "content": texto_completo
                    }
                    
                    # Converte o dicionário para uma string JSON e escreve no arquivo de saída
                    f_out.write(json.dumps(dados, ensure_ascii=False) + '\n')
                    
                    arquivos_processados += 1
                    
                except Exception as e:
                    # Se ocorrer um erro (ex: arquivo corrompido), registra e continua
                    print(f"AVISO: Falha ao processar o arquivo '{nome_arquivo}'. Erro: {e}")
                    arquivos_falha.append(nome_arquivo)

    print("\n--- Processo Concluído ---")
    print(f"✅ Arquivos processados com sucesso: {arquivos_processados}")
    print(f"❌ Arquivos que falharam: {len(arquivos_falha)}")
    
    if arquivos_falha:
        print("\nOs seguintes arquivos não puderam ser lidos (podem estar corrompidos):")
        for arquivo in arquivos_falha:
            print(f"  - {arquivo}")
            
    print(f"\nO arquivo final '{arquivo_saida_jsonl}' foi criado com sucesso!")


# Executa a função principal do script
if __name__ == "__main__":
    processar_pasta_docx()
