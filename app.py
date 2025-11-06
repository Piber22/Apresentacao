import streamlit as st
import pdfplumber
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, Alignment
import os
import shutil
from datetime import datetime
import tempfile
import io

# Configuração da página
st.set_page_config(page_title="Gerenciador de Longas", layout="centered")
st.title("🛏️ Gerenciador de Longas")
st.markdown("Faça upload do **PDF atual** e, opcionalmente, do **PDF anterior** para atualizar a planilha.")

# === Uploads ===
col1, col2 = st.columns(2)
with col1:
    pdf_novo_file = st.file_uploader("📄 PDF Atual (data.pdf)", type="pdf", key="novo")
with col2:
    pdf_antigo_file = st.file_uploader("📄 PDF Anterior (opcional)", type="pdf", key="antigo")

# === Botão de processamento ===
if st.button("🚀 Atualizar Longas", type="primary"):
    if not pdf_novo_file:
        st.error("Por favor, faça upload do PDF atual.")
    else:
        with st.spinner("Processando..."):
            try:
                # === Criar diretório temporário ===
                temp_dir = tempfile.mkdtemp()
                excel_path = os.path.join(temp_dir, "Longas.xlsx")
                pdf_path_new = os.path.join(temp_dir, "data.pdf")
                pdf_path_old = os.path.join(temp_dir, "data_anterior.pdf")

                # Salvar PDFs
                with open(pdf_path_new, "wb") as f:
                    f.write(pdf_novo_file.getbuffer())

                if pdf_antigo_file:
                    with open(pdf_path_old, "wb") as f:
                        f.write(pdf_antigo_file.getbuffer())

                # === Funções do seu código (adaptadas) ===
                def extrair_data_pdf(caminho_pdf):
                    try:
                        with pdfplumber.open(caminho_pdf) as pdf:
                            texto = pdf.pages[0].extract_text()
                            import re
                            match = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
                            if match:
                                return match.group(1)
                    except:
                        pass
                    return datetime.now().strftime("%d/%m/%Y")

                def extrair_pdf(caminho_pdf):
                    colunas_req = ["Leito", "Atendimento", "Paciente", "Dias de Ocupação"]
                    dfs_filtrados = []
                    try:
                        with pdfplumber.open(caminho_pdf) as pdf:
                            for pagina in pdf.pages:
                                tabelas = pagina.extract_tables(table_settings={
                                    "vertical_strategy": "lines",
                                    "horizontal_strategy": "lines",
                                    "snap_x_tolerance": 3,
                                    "join_y_tolerance": 3,
                                    "join_x_tolerance": 3,
                                })
                                for tab in tabelas:
                                    if not tab or not tab[0]: continue
                                    header = [c.strip() if c else "" for c in tab[0]]
                                    if not set(colunas_req).issubset(header): continue
                                    df = pd.DataFrame(tab[1:], columns=header)
                                    if "Métrica" in df.columns:
                                        df.drop(columns=["Métrica"], inplace=True)
                                    dfs_filtrados.append(df[colunas_req])
                    except Exception as e:
                        st.error(f"Erro ao extrair PDF: {e}")
                        return pd.DataFrame(columns=colunas_req)

                    if not dfs_filtrados:
                        return pd.DataFrame(columns=colunas_req)
                    df_pdf = pd.concat(dfs_filtrados, ignore_index=True)
                    df_pdf["Dias de Ocupação"] = pd.to_numeric(df_pdf["Dias de Ocupação"], errors='coerce').fillna(0).astype(int)
                    df_pdf = df_pdf.drop_duplicates(subset=["Leito"])
                    return df_pdf

                def garantir_planilha():
                    if not os.path.exists(excel_path):
                        wb = Workbook()
                        ws1 = wb.active
                        ws1.title = "Dados"
                        wb.create_sheet("Historico")
                        wb.save(excel_path)

                def atualizar_dados(df_novo, data_pdf):
                    garantir_planilha()
                    wb = load_workbook(excel_path)
                    ws = wb["Dados"]

                    # Preservar observações (colunas E e F)
                    obs_dict = {}
                    for row in ws.iter_rows(min_row=6, values_only=True):
                        leito = row[0]
                        if leito:
                            obs_dict[leito] = (row[4], row[5]) if len(row) > 5 else ("", "")

                    # Limpar a partir da linha 6
                    if ws.max_row >= 6:
                        ws.delete_rows(6, ws.max_row - 5)

                    # Ordenar: BOX primeiro
                    df_novo["BOX_PRIORIDADE"] = df_novo["Leito"].str.startswith("BOX")
                    df_novo = df_novo.sort_values(by=["BOX_PRIORIDADE", "Leito"], ascending=[False, True]).drop(columns=["BOX_PRIORIDADE"])

                    # Estilos
                    borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                    centro = Alignment(horizontal='center', vertical='center')
                    esquerda = Alignment(horizontal='left', vertical='center')

                    # Inserir dados
                    for r_idx, row in enumerate(df_novo.itertuples(index=False), start=6):
                        leito = row[0]
                        for c_idx, value in enumerate(row, start=1):
                            celula = ws.cell(row=r_idx, column=c_idx, value=value)
                            celula.border = borda
                            celula.alignment = centro if c_idx in [1, 2, 4] else esquerda
                        # Restaurar Situação e OBS
                        situacao, obs = obs_dict.get(leito, ("", ""))
                        ws.cell(row=r_idx, column=5, value=situacao).border = borda
                        ws.cell(row=r_idx, column=5, value=situacao).alignment = centro
                        ws.cell(row=r_idx, column=6, value=obs).border = borda
                        ws.cell(row=r_idx, column=6, value=obs).alignment = esquerda

                    ws["F3"] = data_pdf
                    wb.save(excel_path)

                def atualizar_historico(df_antigo, df_novo, data_pdf, obs_dict):
                    if not os.path.exists(pdf_path_old):
                        return 0
                    wb = load_workbook(excel_path)
                    ws = wb["Historico"]
                    leitos_antigos = set(df_antigo["Leito"])
                    leitos_novos = set(df_novo["Leito"])
                    leitos_baixados = leitos_antigos - leitos_novos
                    if not leitos_baixados:
                        return 0

                    df_baixa = df_antigo[df_antigo["Leito"].isin(leitos_baixados)].copy()
                    df_baixa["Situação"] = df_baixa["Leito"].map(lambda x: obs_dict.get(x, ("", ""))[0])
                    df_baixa["OBS"] = df_baixa["Leito"].map(lambda x: obs_dict.get(x, ("", ""))[1])
                    df_baixa["Data de Baixa"] = data_pdf

                    next_row = ws.max_row + 1
                    for r_idx, row in enumerate(df_baixa.itertuples(index=False), start=next_row):
                        for c_idx, value in enumerate(row, start=1):
                            ws.cell(row=r_idx, column=c_idx, value=value)
                    wb.save(excel_path)
                    return len(df_baixa)

                # === EXECUÇÃO ===
                data_pdf = extrair_data_pdf(pdf_path_new)
                df_new = extrair_pdf(pdf_path_new)

                if df_new.empty:
                    st.warning("Nenhuma tabela válida encontrada no PDF.")
                else:
                    # Carregar observações
                    obs_dict = {}
                    if os.path.exists(excel_path):
                        wb_temp = load_workbook(excel_path)
                        ws_temp = wb_temp["Dados"]
                        for row in ws_temp.iter_rows(min_row=6, values_only=True):
                            if row[0]:
                                obs_dict[row[0]] = (row[4] if len(row)>4 else "", row[5] if len(row)>5 else "")

                    atualizar_dados(df_new, data_pdf)

                    baixas = 0
                    if os.path.exists(pdf_path_old):
                        df_old = extrair_pdf(pdf_path_old)
                        baixas = atualizar_historico(df_old, df_new, data_pdf, obs_dict)
                        # Atualizar backup
                        shutil.copy2(pdf_path_new, pdf_path_old)
                    else:
                        st.info("Primeira execução: histórico será criado na próxima.")

                    # === Download ===
                    with open(excel_path, "rb") as f:
                        st.download_button(
                            label=f"📥 Baixar Longas.xlsx ({len(df_new)} leitos, {baixas} baixas)",
                            data=f,
                            file_name=f"Longas_{data_pdf.replace('/', '-')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    st.success("Planilha gerada com sucesso!")

            except Exception as e:
                st.error(f"Erro inesperado: {e}")
