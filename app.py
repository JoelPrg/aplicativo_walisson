import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from io import BytesIO
import csv
import re

# === Funções de processamento ===

def gerar_df(uploaded_file):
    try:
        # Tentativa de carregar o arquivo
        wb = load_workbook(filename=uploaded_file, data_only=True)
        ws = wb.active

        tabela = []
        
        # Verifica se a planilha tem dados suficientes
        if ws.max_row < 2:
            raise ValueError("O arquivo não contém dados (menos de 2 linhas)")

        for linha in ws.iter_rows(min_row=2):
            try:
                n_linha = linha[0].row
                
                # Pega a linha inteira como string para mensagem de erro
                linha_completa = " | ".join(str(cell.value) if cell.value is not None else "" for cell in linha)
                
                status = ws.row_dimensions.get(n_linha)
                if not status.hidden:
                    # Verifica se as colunas necessárias existem
                    if len(linha) < 11:
                        raise IndexError(f"A linha não contém colunas suficientes (esperado 11, encontrado {len(linha)})")
                    
                    # Verifica valores nulos nas colunas importantes
                    if linha[3].value is None or linha[10].value is None:
                        raise ValueError("Valores nulos em colunas obrigatórias")
                        
                    tabela.append([linha[3].value, linha[8].value, "", linha[10].value])
                    
            except Exception as linha_error:
                error_msg = f"""
                Erro ao processar linha {n_linha}:
                Conteúdo da linha: {linha_completa}
                Tipo do erro: {type(linha_error).__name__}
                Mensagem: {str(linha_error)}
                """
                print(error_msg)  # Ou logging.error(error_msg)
                
                # Decide se continua ou para conforme a gravidade do erro
                if isinstance(linha_error, (IndexError, AttributeError)):
                    raise  # Erros críticos interrompem a execução
                # Para outros erros, continua para próxima linha
                
        return tabela
        
    except FileNotFoundError:
        raise FileNotFoundError("Arquivo não encontrado ou não pode ser aberto")
    except Exception as e:
        error_msg = f"""
        Erro geral ao processar o arquivo:
        Tipo do erro: {type(e).__name__}
        Mensagem: {str(e)}
        """
        if 'wb' in locals() and 'ws' in locals():
            error_msg += f"\nPlanilha processada: {ws.title}, Total de linhas: {ws.max_row}"
        raise RuntimeError(error_msg) from e

def corrigir_sintaxe_ruas(tabela):
    for linha in tabela:
        endereco = linha[1]
        partes = endereco.split(',', 1)
        if len(partes) < 2:
            continue
        rua = partes[0].strip()
        resto = partes[1].strip()
        match = re.match(r"^(sn|\d+)", resto, re.IGNORECASE)
        if match:
            numero = match.group(1).strip().upper()
            linha[0] = str(linha[0]).replace(".0", "")
            linha[1] = rua
            linha[2] = numero
    return tabela

def corrigir_ruas(tabela):
    caminho_csv = "dicionario_ruas.csv"
    try:
        with open(caminho_csv, encoding="utf-8") as csvfile:
            conteudo_csv = csv.reader(csvfile)
            for linha_conteudo in conteudo_csv:
                for linha_tabela in tabela:
                    if linha_conteudo[0] == linha_tabela[1]:
                        if linha_tabela[2] >= linha_conteudo[1] and linha_tabela[2] <= linha_conteudo[2]:
                            linha_tabela[1] = linha_conteudo[3]
    except FileNotFoundError:
        st.error(f"O arquivo {caminho_csv} não foi encontrado.")
    return tabela

def agrupar_entregas(tabela):
    entregas_agrupadas = []
    lista_ordenada = sorted(tabela, key=lambda x: (x[1], x[2], x[3]))
    grupo_pacotes = []
    endereco_atual = None

    for linha in lista_ordenada:
        numero, rua, num_rua, bairro = linha
        endereco = (rua, num_rua, bairro)
        if endereco == endereco_atual:
            grupo_pacotes.append(numero)
        else:
            if endereco_atual is not None:
                entregas_agrupadas.append(formatar_entrega(grupo_pacotes, *endereco_atual))
            grupo_pacotes = [numero]
            endereco_atual = endereco
    if grupo_pacotes:
        entregas_agrupadas.append(formatar_entrega(grupo_pacotes, *endereco_atual))
    return entregas_agrupadas

def formatar_entrega(lista_pacotes, rua, numero, bairro):
    if len(lista_pacotes) == 1:
        pacotes_str = lista_pacotes[0]
    else:
        pacotes_str = ", ".join(lista_pacotes[:-1]) + " e " + lista_pacotes[-1]
    return [pacotes_str, rua, numero, bairro]

def gerar_planilha(tabela):
    buffer = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Entregas"
    ws.append(["Pacotes", "Rua", "Número", "Bairro"])
    for linha in tabela:
        ws.append(linha)
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# === Streamlit App ===

st.title("Agrupador de Entregas 📦")

uploaded_excel = st.file_uploader("Envie o arquivo de entregas (.xlsx)", type=["xlsx"])

if uploaded_excel:
    with st.spinner("Processando..."):
        tabela = gerar_df(uploaded_excel)
        tabela_corrigida = corrigir_sintaxe_ruas(tabela)
        tabela_corrigida = corrigir_ruas(tabela_corrigida)
        entregas_agrupadas = agrupar_entregas(tabela_corrigida)

        # Mostrar preview
        st.success(f"A rota contém {len(entregas_agrupadas)} paradas.")

        # Arquivo para download
        arquivo_final = gerar_planilha(entregas_agrupadas)
        st.download_button(
            label="📥 Baixar planilha agrupada",
            data=arquivo_final,
            file_name="entregas_agrupadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
