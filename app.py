# app.py
import streamlit as st
import pandas as pd
import os
from datetime import datetime
from io import BytesIO
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from plotly.subplots import make_subplots
import zipfile

# ============================================================================
# CONFIGURAÇÃO INICIAL
# ============================================================================
st.set_page_config(
    page_title="SINISA",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# ESTILOS CSS OTIMIZADOS
# ============================================================================
CSS_STYLES = """
<style>
    :root {
        --primary-color: #1f4788;
        --secondary-color: #2d5aa8;
        --accent-color: #0066cc;
        --light-bg: #f5f7fa;
        --border-color: #d0d8e0;
        --white: #ffffff;
        --dark-text: #333;
        --light-text: #666;
    }
    .appViewContainer, .stAppHeader, .block-container { padding-top: 0 !important; margin-top: 0 !important; }
    header { display: none !important; }
    [data-testid="stTabs"] [role="tablist"] { display: flex; align-items: center; gap: 0 !important; }
    [data-testid="stTabs"] [role="tablist"] button { 
        display: inline-flex; align-items: center; justify-content: center; gap: 0.4rem; 
        padding: 0.6rem 1rem !important; height: 2.5rem; 
    }
    [data-testid="stTabs"] [role="tablist"] button svg { width: 1.2rem; height: 1.2rem; }
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: var(--white); margin: 0 !important; padding: 0 !important; }
    .header-container { 
        background: linear-gradient(135deg, #0B3040 0%, #114157 100%); 
        padding: 2rem; border-radius: 8px; margin-bottom: 1rem; 
        color: white; box-shadow: 0 4px 6px rgba(0,0,0,0.1); 
    }
    .header-title { font-size: 2.5rem; font-weight: 700; margin: 0; color: white; }
    .header-subtitle { font-size: 1rem; margin-top: 0.5rem; opacity: 0.95; color: #e8f0ff; }
    .info-card { 
        background: var(--light-bg); border-left: 4px solid var(--accent-color); 
        padding: 1.5rem; border-radius: 6px; margin: 1rem 0; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.05); 
    }
    .info-card-title { font-size: 1.1rem; font-weight: 600; color: var(--primary-color); margin-bottom: 0.5rem; }
    .info-card-content { font-size: 0.95rem; color: var(--dark-text); line-height: 1.6; }
    .detail-section { 
        background: linear-gradient(135deg, #e8eef5 0%, var(--light-bg) 100%) !important; 
        border: 1px solid var(--border-color) !important; border-radius: 8px !important; 
        padding: 1.5rem !important; margin: 1rem 0 !important; 
        box-shadow: 0 2px 8px rgba(31, 71, 136, 0.08) !important; 
    }
    .detail-label { 
        font-weight: 600; color: var(--primary-color); font-size: 0.9rem; 
        text-transform: uppercase; letter-spacing: 0.5px; margin-top: 1rem; margin-bottom: 0.3rem; 
    }
    .detail-value { 
        color: var(--dark-text); font-size: 0.95rem; padding: 0.5rem 0; 
        border-bottom: 1px solid #c0d9f0; 
    }
    .detail-value-multiline { 
        color: var(--dark-text); font-size: 0.95rem; padding: 0.5rem; 
        border-bottom: 1px solid #c0d9f0; white-space: pre-wrap; word-wrap: break-word; 
        font-family: monospace; margin: 0 !important; line-height: 1.4; display: block; 
    }
    .badge { 
        display: inline-block; padding: 0.3rem 0.8rem; border-radius: 20px; 
        font-size: 0.8rem; font-weight: 600; margin-right: 0.5rem; margin-bottom: 0.5rem; 
    }
    .badge-obrigatorio { background-color: #ffe6e6; color: #c41e3a; }
    .badge-opcional { background-color: #e6f2ff; color: var(--accent-color); }
    .badge-modulo { background-color: #f0e6ff; color: #6b21a8; }
    .table-header { 
        background-color: #0B3040; color: white; padding: 0.3rem 0.75rem !important; 
        font-weight: 600; font-size: 1rem; text-align: left; border-radius: 4px; 
        margin: 0 !important; display: flex; align-items: center; height: var(--table-row-height); line-height: 1.1 !important; 
    }
    .table-row { 
        padding: 0.1rem 0.75rem !important; font-size: 1rem; color: var(--dark-text); 
        margin: 0 !important; line-height: 1.3 !important; height: var(--table-row-height); 
        display: flex; align-items: center; word-wrap: break-word; overflow-wrap: break-word; 
    }
    .table-divider { 
        height: 1px; 
        background-color: #e0e0e0; 
        margin: 0 !important; 
        padding: 0 !important;
        display: flex;
        align-items: center;
    }
    .metric-card { 
        background: linear-gradient(135deg, var(--white) 0%, var(--light-bg) 100%); 
        border: 1px solid var(--border-color); border-radius: 8px; padding: 1rem; 
        text-align: center; box-shadow: 0 2px 8px rgba(31, 71, 136, 0.08); 
    }
    .metric-value { font-size: 2rem; font-weight: 700; color: var(--primary-color); margin: 0.3rem 0; }
    .metric-label { 
        font-size: 0.8rem; color: var(--light-text); text-transform: uppercase; 
        letter-spacing: 0.5px; font-weight: 600; 
    }
    .modal-container { 
        background: #d4e4f7 !important; border-radius: 6px !important; padding: 1.5rem !important; 
        margin: 1rem 0 !important; border: 1px solid #a8c5e8 !important; 
    }
    .modal-container > div { background: transparent !important; }
    /* Módulos - Cores */
    .modulo-agua-header { background-color: var(--primary-color); }
    .modulo-agua-subheader { background-color: var(--secondary-color); }
    .modulo-agua-subform { background-color: #d4e4f7; border-left-color: var(--secondary-color); }
    .modulo-agua-subgrupo { background-color: #e8eef5; border-left-color: var(--secondary-color); }
    .modulo-esgoto-header { background-color: #2d7a4a; }
    .modulo-esgoto-subheader { background-color: #3d9a5a; }
    .modulo-esgoto-subform { background-color: #d4e8d9; border-left-color: #3d9a5a; }
    .modulo-esgoto-subgrupo { background-color: #e8f0e8; border-left-color: #3d9a5a; }
    .modulo-info-complementares-header { background-color: #6b5b95; }
    .modulo-info-complementares-subheader { background-color: #8b7bb8; }
    .modulo-info-complementares-subform { background-color: #e8e4f0; border-left-color: #8b7bb8; }
    .modulo-info-complementares-subgrupo { background-color: #f0eef5; border-left-color: #8b7bb8; }
    [data-testid="stVerticalBlock"], [data-testid="stMarkdownContainer"], 
    [data-testid="stColumn"], .stMarkdown, .main, p, h1, h2, h3, h4, h5, h6, 
    [data-testid="stTabs"], [data-testid="stContainer"] { 
        margin: 0 !important; padding: 0 !important; 
    }
</style>
"""
st.markdown(CSS_STYLES, unsafe_allow_html=True)

# ============================================================================
# FUNÇÕES AUXILIARES - UTILITÁRIOS
# ============================================================================
def hex_to_argb(hex_color):
    """Converte cor hex (#RRGGBB) para aRGB (FFRRGGBB)"""
    hex_color = hex_color.lstrip('#')
    return f"FF{hex_color.upper()}"

def get_color_palette():
    """Retorna paleta de cores centralizada para dashboard e Excel"""
    return {
        'água': {
            'header_color': '#1f4788',
            'subheader_color': '#2d5aa8',
            'subform_color': '#d4e4f7',
            'subgrupo_color': '#e8eef5',
            'border_color': '#2d5aa8',
            'excel_header': '#1F4788',
            'excel_subheader': '#2d5aa8'
        },
        'esgoto': {
            'header_color': '#2d7a4a',
            'subheader_color': '#3d9a5a',
            'subform_color': '#d4e8d9',
            'subgrupo_color': '#e8f0e8',
            'border_color': '#3d9a5a',
            'excel_header': '#2d7a4a',
            'excel_subheader': '#3d9a5a'
        },
        'informações complementares': {
            'header_color': '#6b5b95',
            'subheader_color': '#8b7bb8',
            'subform_color': '#e8e4f0',
            'subgrupo_color': '#f0eef5',
            'border_color': '#8b7bb8',
            'excel_header': '#6b5b95',
            'excel_subheader': '#8b7bb8'
        }
    }

def get_modulo_color(modulo):
    """Retorna cores baseadas no módulo (usa paleta centralizada)"""
    palette = get_color_palette()
    modulo_lower = modulo.lower()
    for chave, cores in palette.items():
        if chave in modulo_lower:
            return cores
    # Padrão se não encontrar
    return palette['água']

def sort_modulos(modulos_list):
    """Ordena módulos: Água, Esgoto, Informações Complementares"""
    modulos_validos = [str(m).strip() for m in modulos_list if pd.notna(m) and str(m).strip() != 'nan']
    agua = [m for m in modulos_validos if 'água' in m.lower()]
    esgoto = [m for m in modulos_validos if 'esgoto' in m.lower()]
    info_complementares = [m for m in modulos_validos if 'água' not in m.lower() and 'esgoto' not in m.lower()]
    return sorted(agua) + sorted(esgoto) + sorted(info_complementares)

def formatar_nome_com_unidade(nome, unidade):
    """Formata nome com unidade"""
    if pd.isna(unidade) or unidade == "" or unidade == "-":
        return nome
    return f"{nome} ({unidade})"

def formatar_brasileiro(valor):
    """Formata número para padrão brasileiro"""
    if pd.isna(valor) or valor == "":
        return ""
    try:
        num = float(valor)
        return f"{num:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return str(valor)

def get_col_letter(col_num):
    """Converte número de coluna para letra (1->A, 27->AA)"""
    result = ""
    while col_num > 0:
        col_num -= 1
        result = chr(65 + col_num % 26) + result
        col_num //= 26
    return result

# ============================================================================
# CARREGAMENTO E CACHE DE DADOS
# ============================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

@st.cache_data
def load_data():
    """Carrega dados do glossário"""
    file_path = os.path.join(DATA_DIR, "Glossario_SINISA.xlsx")
    df = pd.read_excel(file_path, sheet_name=0)
    colunas_remover = ['Coleta', 'Obrigatória?', 'Status', 'Relatório Gerencial - AGEMS']
    df = df.drop(columns=[col for col in colunas_remover if col in df.columns], errors='ignore')
    df.columns = df.columns.str.strip()
    return df

@st.cache_data
def load_dados_agems():
    """Carrega dados AGEMS"""
    file_path = os.path.join(DATA_DIR, "Dados_SINISA.xlsx")
    try:
        df_agems = pd.read_excel(file_path, sheet_name="Dados AGEMS")
        df_agems.columns = df_agems.columns.str.strip()
        mapa_colunas = {
            'Código IBGE - Município': 'codigo_ibge_municipio',
            'Município': 'municipio',
            'Módulo': 'modulo',
            'Formulário': 'formulario',
            'Subformulário/Gruopo': 'subformulario_gruopo',
            'Subgrupo/Palavra-chave': 'subgrupo_palavra_chave',
            'Área Responsável': 'area_responsavel',
            'Código da informação/indicador': 'codigo_da_informacao_indicador',
            'Nome da informação/indicador': 'nome_da_informacao_indicador',
            'Informação/Indicador': 'informacao_indicador',
            'Unidade': 'unidade',
            'Informação Completa': 'informacao_completa'
        }
        for col_original, col_novo in mapa_colunas.items():
            if col_original in df_agems.columns:
                df_agems = df_agems.rename(columns={col_original: col_novo})
        return df_agems
    except Exception as e:
        st.error(f"❌ Erro ao carregar dados AGEMS: {str(e)}")
        return None

@st.cache_data
def load_dados_sinisa():
    """Carrega dados SINISA"""
    file_path = os.path.join(DATA_DIR, "Dados_SINISA.xlsx")
    try:
        df_sinisa = pd.read_excel(file_path, sheet_name="Dados SINISA")
        df_sinisa.columns = df_sinisa.columns.str.strip()
        mapa_colunas = {
            'Código IBGE - Município': 'codigo_ibge_municipio',
            'Município': 'municipio',
            'Módulo': 'modulo',
            'Formulário': 'formulario',
            'Subformulário/Gruopo': 'subformulario_gruopo',
            'Subgrupo/Palavra-chave': 'subgrupo_palavra_chave',
            'Área Responsável': 'area_responsavel',
            'Código da informação/indicador': 'codigo_da_informacao_indicador',
            'Nome da informação/indicador': 'nome_da_informacao_indicador',
            'Informação/Indicador': 'informacao_indicador',
            'Unidade': 'unidade',
            'Informação Completa': 'informacao_completa',
            'Anual': 'anual'
        }
        for col_original, col_novo in mapa_colunas.items():
            if col_original in df_sinisa.columns:
                df_sinisa = df_sinisa.rename(columns={col_original: col_novo})
        return df_sinisa
    except Exception as e:
        st.error(f"❌ Erro ao carregar dados SINISA: {str(e)}")
        return None

try:
    df = load_data()
except FileNotFoundError:
    st.error(f"❌ Arquivo não encontrado em: {os.path.join(DATA_DIR, 'Glossario_SINISA.xlsx')}")
    st.stop()
except Exception as e:
    st.error(f"❌ Erro ao carregar o arquivo: {str(e)}")
    st.stop()

@st.cache_data
def load_correspondencia_sigis_sinisa():
    """Carrega tabela de correspondência SIGIS-SINISA"""
    file_path = os.path.join(DATA_DIR, "Correspondencia_SIGIS_SINISA.xlsx")
    try:
        df = pd.read_excel(file_path, sheet_name=0)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar correspondência: {str(e)}")
        return None

@st.cache_data
def load_dados_sigis():
    """Carrega dados SIGIS"""
    file_path = os.path.join(DATA_DIR, "Dados_SIGIS.xlsx")
    try:
        df = pd.read_excel(file_path, sheet_name=0)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados SIGIS: {str(e)}")
        return None
# ============================================================================
# FUNÇÕES GENÉRICAS DE FILTROS
# ============================================================================

def aplicar_filtros(df, filtros_dict):
    """Aplica múltiplos filtros a um dataframe"""
    df_filtrado = df.copy()
    
    for coluna, valor in filtros_dict.items():
        if valor and valor != "Todos" and coluna in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado[coluna] == valor]
    
    return df_filtrado
# ============================================================================
# FUNÇÕES DE GERAÇÃO DE EXCEL - OTIMIZADAS
# ============================================================================

def criar_tabela_html(df_subgrupo, colunas_meses, largura_cod_desc=350, largura_un=60, largura_mes=60):
    """Cria tabela HTML otimizada para relatórios"""
    html = '<div style="overflow-x: auto; margin: 0 0 0.2rem 0;">'
    html += '<table style="width: 100%; border-collapse: collapse; font-size: 0.85rem; border: none;">'
    html += '<thead><tr style="background-color: #F2F2F2; color: black; position: sticky; top: 0;">'
    html += f'<th style="padding: 0.6rem; text-align: left; border: 1px solid #d0d8e0; width: {largura_cod_desc}px; min-width: {largura_cod_desc}px;">Código - Descrição</th>'
    html += f'<th style="padding: 0.6rem; text-align: center; border: 1px solid #d0d8e0; width: {largura_un}px; min-width: {largura_un}px;">Un.</th>'
    for mes in colunas_meses:
        html += f'<th style="padding: 0.6rem; text-align: center; border: 1px solid #d0d8e0; width: {largura_mes}px; min-width: {largura_mes}px; white-space: nowrap;">{mes}</th>'
    html += '</tr></thead><tbody>'
    for idx, row in df_subgrupo.iterrows():
        html += '<tr style="background-color: #ffffff; border-bottom: 1px solid #d0d8e0;">'
        codigo_descricao = f"{row['codigo_da_informacao_indicador']} - {row['nome_da_informacao_indicador']}"
        html += f'<td style="padding: 0.5rem; text-align: left; border: 1px solid #d0d8e0; width: {largura_cod_desc}px; min-width: {largura_cod_desc}px; font-weight: 500;">{codigo_descricao}</td>'
        html += f'<td style="padding: 0.5rem; text-align: center; border: 1px solid #d0d8e0; width: {largura_un}px; min-width: {largura_un}px;">{row["unidade"]}</td>'
        for mes in colunas_meses:
            valor_formatado = formatar_brasileiro(row[mes])
            html += f'<td style="padding: 0.5rem; text-align: right; border: 1px solid #d0d8e0; width: {largura_mes}px; min-width: {largura_mes}px; font-family: \'Calibri\', monospace; font-weight: 500;">{valor_formatado}</td>'
        html += '</tr>'
    html += '</tbody></table></div>'
    return html

def criar_relatorio_excel_generico(df_para_exportar, municipio_selecionado, colunas_meses):
    """Cria relatório Excel genérico para município - Calibri 11, cores consistentes"""
    wb = Workbook()
    ws = wb.active
    ws.title = municipio_selecionado[:31]
    row_atual = 1
    modulos_unicos = sort_modulos(df_para_exportar['modulo'].unique().tolist())
    
    for modulo in modulos_unicos:
        df_modulo = df_para_exportar[df_para_exportar['modulo'] == modulo]
        cores = get_modulo_color(modulo)
        num_cols = 2 + len(colunas_meses)
        
        # Cabeçalho Módulo
        ws.merge_cells(f'A{row_atual}:{chr(64 + num_cols)}{row_atual}')
        cell = ws[f'A{row_atual}']
        cell.value = modulo
        cell.font = Font(name='Calibri', size=11, bold=True, color=hex_to_argb('#FFFFFF'))
        cell.fill = PatternFill(start_color=hex_to_argb(cores['header_color']).lstrip('FF'), 
                               end_color=hex_to_argb(cores['header_color']).lstrip('FF'), fill_type='solid')
        cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[row_atual].height = 20
        row_atual += 1
        
        for formulario in df_modulo['formulario'].unique():
            df_formulario = df_modulo[df_modulo['formulario'] == formulario]
            
            # Cabeçalho Formulário
            ws.merge_cells(f'A{row_atual}:{chr(64 + num_cols)}{row_atual}')
            cell = ws[f'A{row_atual}']
            cell.value = formulario
            cell.font = Font(name='Calibri', size=11, bold=True, color=hex_to_argb('#FFFFFF'))
            cell.fill = PatternFill(start_color=hex_to_argb(cores['subheader_color']).lstrip('FF'), 
                                   end_color=hex_to_argb(cores['subheader_color']).lstrip('FF'), fill_type='solid')
            cell.alignment = Alignment(horizontal='left', vertical='center')
            ws.row_dimensions[row_atual].height = 18
            row_atual += 1
            
            for subformulario in df_formulario['subformulario_gruopo'].unique():
                df_subformulario = df_formulario[df_formulario['subformulario_gruopo'] == subformulario]
                
                # Cabeçalho Subformulário
                ws.merge_cells(f'A{row_atual}:{chr(64 + num_cols)}{row_atual}')
                cell = ws[f'A{row_atual}']
                cell.value = subformulario
                cell.font = Font(name='Calibri', size=11, bold=True, color=hex_to_argb('#1F4788'))
                cell.fill = PatternFill(start_color=hex_to_argb(cores['subform_color']).lstrip('FF'), 
                                       end_color=hex_to_argb(cores['subform_color']).lstrip('FF'), fill_type='solid')
                cell.alignment = Alignment(horizontal='left', vertical='center')
                ws.row_dimensions[row_atual].height = 16
                row_atual += 1
                
                for subgrupo in df_subformulario['subgrupo_palavra_chave'].unique():
                    df_subgrupo = df_subformulario[df_subformulario['subgrupo_palavra_chave'] == subgrupo]
                    
                    # Cabeçalho Subgrupo
                    ws.merge_cells(f'A{row_atual}:{chr(64 + num_cols)}{row_atual}')
                    cell = ws[f'A{row_atual}']
                    cell.value = subgrupo
                    cell.font = Font(name='Calibri', size=11, bold=True, color=hex_to_argb('#1F4788'))
                    cell.fill = PatternFill(start_color=hex_to_argb(cores['subgrupo_color']).lstrip('FF'), 
                                           end_color=hex_to_argb(cores['subgrupo_color']).lstrip('FF'), fill_type='solid')
                    cell.border = Border(left=Side(style='medium', color=hex_to_argb(cores['border_color']).lstrip('FF')))
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                    ws.row_dimensions[row_atual].height = 16
                    row_atual += 1
                    
                    # Cabeçalhos tabela
                    headers = ['Código - Descrição', 'Un.'] + colunas_meses
                    for col_num, header in enumerate(headers, 1):
                        cell = ws.cell(row=row_atual, column=col_num)
                        cell.value = header
                        cell.font = Font(name='Calibri', size=11, bold=True, color=hex_to_argb("#000000"))
                        cell.fill = PatternFill(start_color=hex_to_argb("#E1E4E9").lstrip('FF'), 
                                    end_color=hex_to_argb('#1F4788').lstrip('FF'), fill_type='solid')
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        cell.border = Border(
                            left=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                            right=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                            top=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                            bottom=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF'))
                        )
                    ws.row_dimensions[row_atual].height = 20
                    row_atual += 1
                    
                    # Dados
                    for idx, row in df_subgrupo.iterrows():
                        ws.cell(row=row_atual, column=1).value = f"{row['codigo_da_informacao_indicador']} - {row['nome_da_informacao_indicador']}"
                        ws.cell(row=row_atual, column=1).font = Font(name='Calibri', size=11)
                        ws.cell(row=row_atual, column=2).value = row['unidade']
                        ws.cell(row=row_atual, column=2).font = Font(name='Calibri', size=11)
                        
                        for col_num, mes in enumerate(colunas_meses, 3):
                            valor = row[mes]
                            if pd.notna(valor) and valor != "":
                                try:
                                    ws.cell(row=row_atual, column=col_num).value = float(valor)
                                    ws.cell(row=row_atual, column=col_num).number_format = '#,##0.00'
                                except:
                                    ws.cell(row=row_atual, column=col_num).value = valor
                            ws.cell(row=row_atual, column=col_num).font = Font(name='Calibri', size=11)
                        
                        for col_num in range(1, len(headers) + 1):
                            cell = ws.cell(row=row_atual, column=col_num)
                            cell.border = Border(
                                left=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                                right=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                                top=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                                bottom=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF'))
                            )
                            cell.alignment = Alignment(horizontal='left' if col_num == 1 else 'right', vertical='center', wrap_text=True)
                        ws.row_dimensions[row_atual].height = 18
                        row_atual += 1
                    row_atual += 1
    
    ws.column_dimensions['A'].width = 131
    ws.column_dimensions['B'].width = 11
    for col_num in range(3, 3 + len(colunas_meses)):
        ws.column_dimensions[chr(64 + col_num)].width = 14
    
    return wb

@st.cache_data
def criar_relatorio_consolidado_excel(df_para_exportar, indicadores_filtrados_tuple):
    """Cria relatório consolidado em Excel - Calibri 11, cores consistentes - COM CACHE"""
    # Converter tuple de volta para lista
    indicadores_filtrados = list(indicadores_filtrados_tuple) if indicadores_filtrados_tuple else []
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado"
    df_pivot_data = df_para_exportar[['codigo_ibge_municipio', 'municipio', 'codigo_da_informacao_indicador', 
                                       'nome_da_informacao_indicador', 'unidade', 'anual']].copy()
    df_pivot_data = df_pivot_data.drop_duplicates(subset=['codigo_ibge_municipio', 'municipio', 'codigo_da_informacao_indicador'])
    if len(indicadores_filtrados) > 0:
        df_pivot_data = df_pivot_data[df_pivot_data['codigo_da_informacao_indicador'].isin(indicadores_filtrados)]
    codigo_nome_unidade_excel = {}
    for _, row in df_pivot_data.iterrows():
        codigo = row['codigo_da_informacao_indicador']
        codigo_nome_unidade_excel[codigo] = formatar_nome_com_unidade(row['nome_da_informacao_indicador'], row['unidade'])
    df_pivot_data = df_pivot_data.pivot_table(
        index=['codigo_ibge_municipio', 'municipio'],
        columns='codigo_da_informacao_indicador',
        values='anual',
        aggfunc='first'
    ).reset_index()
    headers = list(df_pivot_data.columns)
    
    # Cabeçalho linha 1
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = "" if col_num <= 2 else codigo_nome_unidade_excel.get(header, "")
        cell.font = Font(name='Calibri', size=11, bold=True, color=hex_to_argb('#FFFFFF'))
        cell.fill = PatternFill(start_color=hex_to_argb('#0B3040').lstrip('FF'), 
                               end_color=hex_to_argb('#0B3040').lstrip('FF'), fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(
            left=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
            right=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
            top=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
            bottom=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF'))
        )
    ws.row_dimensions[1].height = 25
    
    # Cabeçalho linha 2
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_num)
        cell.value = header
        cell.font = Font(name='Calibri', size=11, bold=True, color=hex_to_argb('#FFFFFF'))
        cell.fill = PatternFill(start_color=hex_to_argb('#0B3040').lstrip('FF'), 
                               end_color=hex_to_argb('#0B3040').lstrip('FF'), fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(
            left=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
            right=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
            top=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
            bottom=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF'))
        )
    ws.row_dimensions[2].height = 20
    ws.freeze_panes = 'C3'
    
    # Dados
    for row_num, (idx, row) in enumerate(df_pivot_data.iterrows(), 3):
        for col_num, col_name in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=col_num)
            value = row[col_name]
            if col_num == 1:
                cell.value = value
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif col_num == 2:
                cell.value = value
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                if pd.notna(value) and value != "":
                    try:
                        cell.value = float(value)
                        cell.number_format = '#,##0.00'
                    except:
                        cell.value = value
                cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.font = Font(name='Calibri', size=11)
            cell.border = Border(
                left=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                right=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                top=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                bottom=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF'))
            )
    
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 35
    for col_num in range(3, len(headers) + 1):
        ws.column_dimensions[get_col_letter(col_num)].width = 18
    
    # RETORNAR O WORKBOOK, NÃO O BYTESIO
    return wb

def criar_tabela_consolidada_html(df_para_tabela, indicadores_filtrados):
    """Cria tabela HTML consolidada com Município e Indicadores"""
    df_pivot_html = df_para_tabela[['codigo_ibge_municipio', 'municipio', 'codigo_da_informacao_indicador', 
                                     'nome_da_informacao_indicador', 'unidade', 'anual']].copy()
    df_pivot_html = df_pivot_html.drop_duplicates(subset=['codigo_ibge_municipio', 'municipio', 'codigo_da_informacao_indicador'])
    
    if len(indicadores_filtrados) > 0:
        df_pivot_html = df_pivot_html[df_pivot_html['codigo_da_informacao_indicador'].isin(indicadores_filtrados)]
    
    codigo_nome_unidade_html = {}
    for _, row in df_pivot_html.iterrows():
        codigo = row['codigo_da_informacao_indicador']
        codigo_nome_unidade_html[codigo] = formatar_nome_com_unidade(row['nome_da_informacao_indicador'], row['unidade'])
    
    df_pivot_html = df_pivot_html.pivot_table(
        index=['codigo_ibge_municipio', 'municipio'],
        columns='codigo_da_informacao_indicador',
        values='anual',
        aggfunc='first'
    ).reset_index()
    
    indicadores = [col for col in df_pivot_html.columns if col not in ['codigo_ibge_municipio', 'municipio']]
    
    html = '<div style="overflow-x: auto; margin: 0.5rem 0; border-radius: 4px; border: 1px solid #d0d8e0;">'
    html += '<table style="width: 100%; border-collapse: collapse; font-size: 0.8rem;">'
    
    # Cabeçalho linha 1 - Descrição
    html += '<thead><tr style="background-color: #104861; color: white; position: sticky; top: 0; z-index: 2;">'
    html += '<th style="padding: 0.25rem 0.3rem; text-align: center; border: 1px solid #d0d8e0; min-width: 280px; font-weight: 700; font-size: 0.8rem; white-space: normal; position: sticky; left: 0; z-index: 3; background-color: #104861;">Município</th>'
    
    for indicador in indicadores:
        nome = codigo_nome_unidade_html.get(indicador, "")
        html += f'<th style="padding: 0.25rem 0.3rem; text-align: center; border: 1px solid #d0d8e0; min-width: 110px; font-weight: 600; font-size: 0.85rem; white-space: normal;">{nome}</th>'
    
    html += '</tr></thead>'
    
    # Cabeçalho linha 2 - Código
    html += '<thead><tr style="background-color: #0B3040; color: white; position: sticky; top: 28px; z-index: 2;">'
    html += '<th style="padding: 0.25rem 0.3rem; text-align: center; border: 1px solid #d0d8e0; min-width: 280px; font-weight: 700; font-size: 0.8rem; white-space: nowrap; position: sticky; left: 0; z-index: 3; background-color: #0B3040;"></th>'
    
    for indicador in indicadores:
        html += f'<th style="padding: 0.25rem 0.3rem; text-align: center; border: 1px solid #d0d8e0; min-width: 110px; font-weight: 600; font-size: 0.85rem; white-space: nowrap;">{indicador}</th>'
    
    html += '</tr></thead>'
    
    # Corpo da tabela
    html += '<tbody>'
    
    for idx, row in df_pivot_html.iterrows():
        html += '<tr style="background-color: #ffffff; border-bottom: 1px solid #d0d8e0; height: 16px;">'
        
        html += f'<td style="padding: 0.2rem 0.3rem; text-align: left; border: 1px solid #d0d8e0; font-weight: 700; font-size: 0.85rem; white-space: nowrap; position: sticky; left: 0; z-index: 1; background-color: #f0f0f0;">{row["municipio"]}</td>'
        
        for indicador in indicadores:
            valor = row[indicador]
            valor_formatado = formatar_brasileiro(valor)
            html += f'<td style="padding: 0.2rem 0.3rem; text-align: right; border: 1px solid #d0d8e0; font-family: \'Calibri\', monospace; font-weight: 500; font-size: 0.85rem; white-space: nowrap;">{valor_formatado}</td>'
        
        html += '</tr>'
    
    html += '</tbody></table></div>'
    return html

# ============================================================================
# FUNÇÕES DE EXPORTAÇÃO ZIP
# ============================================================================

@st.cache_data
def criar_exportacao_zip_por_modulo(df_sinisa_tuple):
    """Cria ZIP com dados por módulo - Calibri 11, cores consistentes - COM CACHE"""
    # Converter tuple de volta para DataFrame (necessário para cache)
    df_sinisa = pd.DataFrame(df_sinisa_tuple)
    
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for modulo in ['Água', 'Esgoto', 'Informações Complementares']:
            df_modulo = df_sinisa[df_sinisa['modulo'] == modulo].copy()
            if len(df_modulo) == 0:
                continue
            for info_tipo in ['Informação', 'Indicador']:
                df_tipo = df_modulo[df_modulo['informacao_indicador'] == info_tipo].copy()
                if len(df_tipo) == 0:
                    continue
                wb = Workbook()
                ws_primeiro = wb.active
                ws_primeiro.title = "Índice"
                for grupo in sorted(df_tipo['formulario'].dropna().unique().tolist()):
                    df_grupo = df_tipo[df_tipo['formulario'] == grupo]
                    codigo_nome_unidade = {}
                    for _, row in df_grupo.iterrows():
                        codigo = row['codigo_da_informacao_indicador']
                        codigo_nome_unidade[codigo] = formatar_nome_com_unidade(row['nome_da_informacao_indicador'], row['unidade'])
                    df_pivot = df_grupo[['codigo_ibge_municipio', 'municipio', 'codigo_da_informacao_indicador', 'anual']].copy()
                    df_pivot = df_pivot.drop_duplicates(subset=['codigo_ibge_municipio', 'municipio', 'codigo_da_informacao_indicador'])
                    df_pivot = df_pivot.pivot_table(
                        index=['codigo_ibge_municipio', 'municipio'],
                        columns='codigo_da_informacao_indicador',
                        values='anual',
                        aggfunc='first'
                    ).reset_index()
                    ws = wb.create_sheet(title=grupo[:31])
                    headers = list(df_pivot.columns)
                    # Cabeçalhos
                    for col_num, header in enumerate(headers, 1):
                        for row_num in [1, 2]:
                            cell = ws.cell(row=row_num, column=col_num)
                            if row_num == 1:
                                cell.value = "" if col_num <= 2 else codigo_nome_unidade.get(header, "")
                            else:
                                cell.value = header
                            cell.font = Font(name='Calibri', size=11, bold=True, color=hex_to_argb('#FFFFFF'))
                            cell.fill = PatternFill(start_color=hex_to_argb('#2d5aa8' if row_num == 1 else '#1F4788').lstrip('FF'), 
                                                   end_color=hex_to_argb('#2d5aa8' if row_num == 1 else '#1F4788').lstrip('FF'), fill_type='solid')
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            cell.border = Border(
                                left=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                                right=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                                top=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                                bottom=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF'))
                            )
                    ws.row_dimensions[1].height = 25
                    ws.row_dimensions[2].height = 20
                    ws.freeze_panes = 'C3'
                    # Dados
                    for row_num, (idx, row) in enumerate(df_pivot.iterrows(), 3):
                        for col_num, col_name in enumerate(headers, 1):
                            cell = ws.cell(row=row_num, column=col_num)
                            value = row[col_name]
                            if col_num == 1:
                                cell.value = value
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                            elif col_num == 2:
                                cell.value = value
                                cell.alignment = Alignment(horizontal='left', vertical='center')
                            else:
                                if pd.notna(value) and value != "":
                                    try:
                                        cell.value = float(value)
                                        cell.number_format = '#,##0.00'
                                    except:
                                        cell.value = value
                                cell.alignment = Alignment(horizontal='right', vertical='center')
                            cell.font = Font(name='Calibri', size=11)
                            cell.border = Border(
                                left=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                                right=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                                top=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                                bottom=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF'))
                            )
                    ws.column_dimensions['A'].width = 18
                    ws.column_dimensions['B'].width = 35
                    for col_num in range(3, len(headers) + 1):
                        ws.column_dimensions[get_col_letter(col_num)].width = 18
                if len(wb.sheetnames) > 1:
                    wb.remove(ws_primeiro)
                modulo_nome = modulo.lower().replace(' ', '_')
                tipo_nome = info_tipo.lower().replace(' ', '_')
                buffer = BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                zip_file.writestr(f'{modulo_nome}_{tipo_nome}.xlsx', buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

@st.cache_data
def criar_exportacao_zip_por_area(df_sinisa_tuple):
    """Cria ZIP com dados por área responsável - Calibri 11, cores consistentes - COM CACHE"""
    # Converter tuple de volta para DataFrame (necessário para cache)
    df_sinisa = pd.DataFrame(df_sinisa_tuple)
    
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        df_informacoes_area = df_sinisa[df_sinisa['informacao_indicador'] == 'Informação'].copy()
        for area in sorted(df_informacoes_area['area_responsavel'].dropna().unique().tolist()):
            df_area = df_informacoes_area[df_informacoes_area['area_responsavel'] == area].copy()
            if len(df_area) == 0:
                continue
            wb = Workbook()
            ws_primeiro = wb.active
            ws_primeiro.title = "Índice"
            for subform in sorted(df_area['subformulario_gruopo'].dropna().unique().tolist()):
                df_subform = df_area[df_area['subformulario_gruopo'] == subform]
                codigo_nome_unidade = {}
                for _, row in df_subform.iterrows():
                    codigo = row['codigo_da_informacao_indicador']
                    codigo_nome_unidade[codigo] = formatar_nome_com_unidade(row['nome_da_informacao_indicador'], row['unidade'])
                df_pivot = df_subform[['codigo_ibge_municipio', 'municipio', 'codigo_da_informacao_indicador', 'anual']].copy()
                df_pivot = df_pivot.drop_duplicates(subset=['codigo_ibge_municipio', 'municipio', 'codigo_da_informacao_indicador'])
                df_pivot = df_pivot.pivot_table(
                    index=['codigo_ibge_municipio', 'municipio'],
                    columns='codigo_da_informacao_indicador',
                    values='anual',
                    aggfunc='first'
                ).reset_index()
                ws = wb.create_sheet(title=subform[:31])
                headers = list(df_pivot.columns)
                # Cabeçalhos
                for col_num, header in enumerate(headers, 1):
                    for row_num in [1, 2]:
                        cell = ws.cell(row=row_num, column=col_num)
                        if row_num == 1:
                            cell.value = "" if col_num <= 2 else codigo_nome_unidade.get(header, "")
                        else:
                            cell.value = header
                        cell.font = Font(name='Calibri', size=11, bold=True, color=hex_to_argb('#FFFFFF'))
                        cell.fill = PatternFill(start_color=hex_to_argb('#2d5aa8' if row_num == 1 else '#1F4788').lstrip('FF'), 
                                               end_color=hex_to_argb('#2d5aa8' if row_num == 1 else '#1F4788').lstrip('FF'), fill_type='solid')
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        cell.border = Border(
                            left=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                            right=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                            top=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                            bottom=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF'))
                        )
                ws.row_dimensions[1].height = 25
                ws.row_dimensions[2].height = 20
                ws.freeze_panes = 'C3'
                # Dados
                for row_num, (idx, row) in enumerate(df_pivot.iterrows(), 3):
                    for col_num, col_name in enumerate(headers, 1):
                        cell = ws.cell(row=row_num, column=col_num)
                        value = row[col_name]
                        if col_num == 1:
                            cell.value = value
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        elif col_num == 2:
                            cell.value = value
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                        else:
                            if pd.notna(value) and value != "":
                                try:
                                    cell.value = float(value)
                                    cell.number_format = '#,##0.00'
                                except:
                                    cell.value = value
                            cell.alignment = Alignment(horizontal='right', vertical='center')
                        cell.font = Font(name='Calibri', size=11)
                        cell.border = Border(
                            left=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                            right=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                            top=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF')),
                            bottom=Side(style='thin', color=hex_to_argb('#D0D8E0').lstrip('FF'))
                        )
                ws.column_dimensions['A'].width = 18
                ws.column_dimensions['B'].width = 35
                for col_num in range(3, len(headers) + 1):
                    ws.column_dimensions[get_col_letter(col_num)].width = 18
            if len(wb.sheetnames) > 1:
                wb.remove(ws_primeiro)
            area_nome = area.lower().replace(' ', '_').replace('/', '_')
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            zip_file.writestr(f'{area_nome}.xlsx', buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

# ============================================================================
# FUNÇÕES DE PÁGINAS
# ============================================================================

# ============================================================================
# FUNÇÃO DE EXIBIÇÃO DE DETALHES EM POPUP MODAL (st.dialog)
# ============================================================================


@st.dialog("📋 Detalhes", width="large")
def exibir_detalhes_popup(row):
    """Exibe detalhes em popup modal com verificações robustas de campos vazios"""
    
    # Função auxiliar para verificar se o valor é realmente válido e não vazio
    def get_valor_valido(valor):
        if pd.isna(valor):
            return ""
        texto = str(valor).strip()
        if texto.lower() in ['nan', 'none', '']:
            return ""
        return texto

    # Título compacto
    st.markdown(f"""
    <div style="background: #d4e4f7; border-radius: 6px; padding: 0.8rem; margin-bottom: 0.8rem; border-left: 4px solid #1f4788;">
        <h4 style="color: #1f4788; margin: 0; font-size: 1.1rem;">
            <strong>{row['Código da informação']}</strong> — {row['Nome da informação']}
        </h4>
    </div>
    """, unsafe_allow_html=True)
    
    # Linha única com 5 colunas
    col1, col2, col3, col4, col5 = st.columns(5, gap="small")
    
    with col1:
        st.markdown(f"""
        <div style="font-size: 0.8rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Módulo</div>
        <div style="font-size: 0.95rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.3rem; line-height: 1.3;">{get_valor_valido(row['Módulo']) or 'N/A'}</div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div style="font-size: 0.8rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Formulário</div>
        <div style="font-size: 0.95rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.3rem; line-height: 1.3;">{get_valor_valido(row['Formulário']) or 'N/A'}</div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div style="font-size: 0.8rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Subformulário</div>
        <div style="font-size: 0.95rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.3rem; line-height: 1.3;">{get_valor_valido(row['Subformulário']) or 'N/A'}</div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div style="font-size: 0.8rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Subgrupo</div>
        <div style="font-size: 0.95rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.3rem; line-height: 1.3;">{get_valor_valido(row['Subgrupo']) or 'N/A'}</div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
        <div style="font-size: 0.8rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Área Responsável</div>
        <div style="font-size: 0.95rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.3rem; line-height: 1.3;">{get_valor_valido(row['Área Responsável']) or 'N/A'}</div>
        """, unsafe_allow_html=True)
    

    
    # Descrição e Unidade
    val_descricao = get_valor_valido(row['Descrição SINISA'])
    descricao_texto = val_descricao if val_descricao else 'Não informada'
    
    val_unidade = get_valor_valido(row['Unidade'])
    unidade_texto = f"<br><span style='color: #1f4788;'><strong>[Unidade: {val_unidade}]</strong></span>" if val_unidade else ""
    
    st.markdown(f"""
    <div>
        <div style="font-size: 0.85rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Descrição</div>
        <div style="font-size: 1rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.4rem; line-height: 1.5;">
            {descricao_texto}{unidade_texto}
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Fórmula
    val_formula = get_valor_valido(row['Fórmula'])
    formula_texto = val_formula if val_formula else 'Não informada'
    st.markdown(f"""
    <div style="margin-top: 0.8rem;">
        <div style="font-size: 0.85rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Fórmula</div>
        <div style="background: #c0d9f0; padding: 0.6rem; border-radius: 4px; font-family: monospace; font-size: 0.95rem; color: #333; line-height: 1.4; overflow-x: auto;">
            {formula_texto}
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Referência
    val_referencia = get_valor_valido(row['Referência'])
    if val_referencia:
        if '\n' in val_referencia:
            referencia_texto = val_referencia.replace('\n', '<br>')
        elif ', ' in val_referencia:
            refs = [r.strip() for r in val_referencia.split(',')]
            referencia_texto = "<br>".join([f"• {r}" for r in refs if r])
        else:
            referencia_texto = val_referencia
    else:
        referencia_texto = 'Não informada'
        
    st.markdown(f"""
    <div style="margin-top: 0.8rem;">
        <div style="font-size: 0.85rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Informações/Referência</div>
        <div style="background: #c0d9f0; padding: 0.6rem; border-radius: 4px; font-size: 0.95rem; color: #333; line-height: 1.6;">
            {referencia_texto}
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Fonte
    val_fonte = get_valor_valido(row['Fonte'])
    fonte_texto = val_fonte if val_fonte else 'Não informada'
    st.markdown(f"""
    <div style="margin-top: 0.8rem;">
        <div style="font-size: 0.85rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Fonte</div>
        <div style="font-size: 1rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.4rem;">
            {fonte_texto}
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Observação (Oculta se vazia e com tratamento para evitar erro de HTML)
    val_observacao = get_valor_valido(row['Observação'])
    if val_observacao:
        obs_texto = val_observacao.replace('\n', '<br>')
        st.markdown(f"""
        <div style="margin-top: 0.8rem;">
            <div style="font-size: 0.85rem; color: #c41e3a; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">⚠️ Observação</div>
            <div style="background: #ffe6e6; padding: 0.6rem; border-radius: 4px; border-left: 3px solid #c41e3a; font-size: 0.95rem; color: #333; line-height: 1.5;">{obs_texto}</div>
        </div>
        """, unsafe_allow_html=True)


def exibir_detalhes_modal(row, idx):
    """Exibe detalhes em popup modal com verificações robustas de campos vazios"""
    
    # Função auxiliar para verificar se o valor é realmente válido e não vazio
    def get_valor_valido(valor):
        if pd.isna(valor):
            return ""
        texto = str(valor).strip()
        if texto.lower() in ['nan', 'none', '']:
            return ""
        return texto

    # Título compacto
    st.markdown(f"""
    <div style="background: #d4e4f7; border-radius: 6px; padding: 0.8rem; margin-bottom: 0.8rem; border-left: 4px solid #1f4788;">
        <h4 style="color: #1f4788; margin: 0; font-size: 1.1rem;">
            <strong>{row['Código da informação']}</strong> — {row['Nome da informação']}
        </h4>
    </div>
    """, unsafe_allow_html=True)
    
    # Linha única com 5 colunas
    col1, col2, col3, col4, col5 = st.columns(5, gap="small")
    
    with col1:
        st.markdown(f"""
        <div style="font-size: 0.8rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Módulo</div>
        <div style="font-size: 0.95rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.3rem; line-height: 1.3;">{get_valor_valido(row['Módulo']) or 'N/A'}</div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div style="font-size: 0.8rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Formulário</div>
        <div style="font-size: 0.95rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.3rem; line-height: 1.3;">{get_valor_valido(row['Formulário']) or 'N/A'}</div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div style="font-size: 0.8rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Subformulário</div>
        <div style="font-size: 0.95rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.3rem; line-height: 1.3;">{get_valor_valido(row['Subformulário']) or 'N/A'}</div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div style="font-size: 0.8rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Subgrupo</div>
        <div style="font-size: 0.95rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.3rem; line-height: 1.3;">{get_valor_valido(row['Subgrupo']) or 'N/A'}</div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
        <div style="font-size: 0.8rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Área Responsável</div>
        <div style="font-size: 0.95rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.3rem; line-height: 1.3;">{get_valor_valido(row['Área Responsável']) or 'N/A'}</div>
        """, unsafe_allow_html=True)
    

    
    # Descrição e Unidade
    val_descricao = get_valor_valido(row['Descrição SINISA'])
    descricao_texto = val_descricao if val_descricao else 'Não informada'
    
    val_unidade = get_valor_valido(row['Unidade'])
    unidade_texto = f"<br><span style='color: #1f4788;'><strong>[Unidade: {val_unidade}]</strong></span>" if val_unidade else ""
    
    st.markdown(f"""
    <div>
        <div style="font-size: 0.85rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Descrição</div>
        <div style="font-size: 1rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.4rem; line-height: 1.5;">
            {descricao_texto}{unidade_texto}
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Fórmula
    val_formula = get_valor_valido(row['Fórmula'])
    formula_texto = val_formula if val_formula else 'Não informada'
    st.markdown(f"""
    <div style="margin-top: 0.8rem;">
        <div style="font-size: 0.85rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Fórmula</div>
        <div style="background: #c0d9f0; padding: 0.6rem; border-radius: 4px; font-family: monospace; font-size: 0.95rem; color: #333; line-height: 1.4; overflow-x: auto;">
            {formula_texto}
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Referência
    val_referencia = get_valor_valido(row['Referência'])
    if val_referencia:
        if '\n' in val_referencia:
            referencia_texto = val_referencia.replace('\n', '<br>')
        elif ', ' in val_referencia:
            refs = [r.strip() for r in val_referencia.split(',')]
            referencia_texto = "<br>".join([f"• {r}" for r in refs if r])
        else:
            referencia_texto = val_referencia
    else:
        referencia_texto = 'Não informada'
        
    st.markdown(f"""
    <div style="margin-top: 0.8rem;">
        <div style="font-size: 0.85rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Informações/Referência</div>
        <div style="background: #c0d9f0; padding: 0.6rem; border-radius: 4px; font-size: 0.95rem; color: #333; line-height: 1.6;">
            {referencia_texto}
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Fonte
    val_fonte = get_valor_valido(row['Fonte'])
    fonte_texto = val_fonte if val_fonte else 'Não informada'
    st.markdown(f"""
    <div style="margin-top: 0.8rem;">
        <div style="font-size: 0.85rem; color: #1f4788; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">Fonte</div>
        <div style="font-size: 1rem; color: #333; border-bottom: 1px solid #c0d9f0; padding-bottom: 0.4rem;">
            {fonte_texto}
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Observação (Oculta se vazia e com tratamento para evitar erro de HTML)
    val_observacao = get_valor_valido(row['Observação'])
    if val_observacao:
        obs_texto = val_observacao.replace('\n', '<br>')
        st.markdown(f"""
        <div style="margin-top: 0.8rem;">
            <div style="font-size: 0.85rem; color: #c41e3a; font-weight: 700; text-transform: uppercase; margin-bottom: 0.2rem; letter-spacing: 0.5px;">⚠️ Observação</div>
            <div style="background: #ffe6e6; padding: 0.6rem; border-radius: 4px; border-left: 3px solid #c41e3a; font-size: 0.95rem; color: #333; line-height: 1.5;">{obs_texto}</div>
        </div>
        """, unsafe_allow_html=True)


def pagina_sobre():
    """Página Sobre"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📖 Sobre o Sistema")
    
    st.info("""
    ⏳ **Esta seção está em construção.**
    
    Em breve, você poderá acessar informações detalhadas sobre o SINISA, sua estrutura, objetivos e funcionalidades.
    """)
    
    st.markdown("---")
    
    st.markdown("""
    ### 🔗 Links Úteis
    - [Portal SINISA](https://www.gov.br/cidades/pt-br/acesso-a-informacao/acoes-e-programas/saneamento/sinisa)
    - [Documentação](https://www.gov.br/cidades/pt-br/acesso-a-informacao/acoes-e-programas/saneamento/sinisa/sinisa-1)
    """)

def pagina_glossario_informacoes():
    """Página Glossário - Informações"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📋 Glossário - Informações")
    
    # Inicializar session state
    if "modulo_filter" not in st.session_state:
        st.session_state.modulo_filter = "Todos"
    if "formulario_filter" not in st.session_state:
        st.session_state.formulario_filter = "Todos"
    if "subformulario_filter" not in st.session_state:
        st.session_state.subformulario_filter = "Todos"
    if "subgrupo_filter" not in st.session_state:
        st.session_state.subgrupo_filter = "Todos"
    if "area_filter" not in st.session_state:
        st.session_state.area_filter = "Todos"
    if "search_codigo" not in st.session_state:
        st.session_state.search_codigo = ""
    if "search_descricao" not in st.session_state:
        st.session_state.search_descricao = ""
    
    # Funções de filtro
    @st.cache_data
    def get_modulos():
        modulos = ["Todos"] + sort_modulos(df['Módulo'].dropna().unique().tolist())
        return modulos
    
    def get_filtro_glossario(coluna, modulo="Todos", formulario="Todos", subformulario="Todos", subgrupo="Todos"):
        """Obtém valores únicos para filtros do glossário"""
        df_temp = df.copy()
        if modulo != "Todos":
            df_temp = df_temp[df_temp['Módulo'] == modulo]
        if formulario != "Todos":
            df_temp = df_temp[df_temp['Formulário'] == formulario]
        if subformulario != "Todos":
            df_temp = df_temp[df_temp['Subformulário'] == subformulario]
        if subgrupo != "Todos":
            df_temp = df_temp[df_temp['Subgrupo'] == subgrupo]
        return ["Todos"] + sorted(df_temp[coluna].dropna().unique().tolist())
    
    # Linha 1 - Filtros
    col1, col2, col3, col4, col5, col6 = st.columns([1.5, 1.5, 1.5, 1.5, 1.5, 1])
    
    with col6:
        st.markdown("<div style='padding-top: 12px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Limpar Filtros", key="btn_clear_filters", use_container_width=True):
            st.session_state.modulo_filter = "Todos"
            st.session_state.formulario_filter = "Todos"
            st.session_state.subformulario_filter = "Todos"
            st.session_state.subgrupo_filter = "Todos"
            st.session_state.area_filter = "Todos"
            st.session_state.search_codigo = ""
            st.session_state.search_descricao = ""
            st.rerun()
    
    with col1:
        modulo_selecionado = st.selectbox("Módulo", get_modulos(), key="modulo_filter", index=0)
    
    with col2:
        formulario_selecionado = st.selectbox("Formulário", 
            get_filtro_glossario('Formulário', modulo_selecionado), key="formulario_filter", index=0)
    
    with col3:
        subformulario_selecionado = st.selectbox("Subformulário", 
            get_filtro_glossario('Subformulário', modulo_selecionado, formulario_selecionado), 
            key="subformulario_filter", index=0)
    
    with col4:
        subgrupo_selecionado = st.selectbox("Subgrupo", 
            get_filtro_glossario('Subgrupo', modulo_selecionado, formulario_selecionado, subformulario_selecionado), 
            key="subgrupo_filter", index=0)
    
    with col5:
        area_selecionada = st.selectbox("Área Responsável", 
            get_filtro_glossario('Área Responsável', modulo_selecionado, formulario_selecionado, 
                                subformulario_selecionado, subgrupo_selecionado), 
            key="area_filter", index=0)
    
    # Linha 2 - Buscas
    col_codigo, col_descricao = st.columns([1.5, 4.5])
    with col_codigo:
        busca_codigo = st.text_input("🔎 Buscar por Código", key="search_codigo")
    with col_descricao:
        busca_descricao = st.text_input("🔎 Buscar por Descrição", key="search_descricao")
    
    # Aplicar filtros
    df_filtrado = aplicar_filtros(df, {
        'Módulo': modulo_selecionado,
        'Formulário': formulario_selecionado,
        'Subformulário': subformulario_selecionado,
        'Subgrupo': subgrupo_selecionado,
        'Área Responsável': area_selecionada
    })
    
    if busca_codigo:
        df_filtrado = df_filtrado[df_filtrado['Código da informação'].str.contains(busca_codigo, case=False, na=False)]
    if busca_descricao:
        df_filtrado = df_filtrado[df_filtrado['Nome da informação'].str.contains(busca_descricao, case=False, na=False)]
    
    # Função de exibição de detalhes
    def exibir_detalhes_modal(row, idx):
        """Exibe detalhes em modal"""
        st.markdown(f"""
        <div class="modal-container">
            <h3 style="color: #1f4788; margin-top: 0; margin-bottom: 1rem;">
                {row['Código da informação']} - {row['Nome da informação']}
            </h3>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"""
            <div class="detail-label">Módulo</div>
            <div class="detail-value">{row['Módulo']}</div>
            <div class="detail-label">Área Responsável</div>
            <div class="detail-value">{row['Área Responsável'] if pd.notna(row['Área Responsável']) else 'N/A'}</div>
            <div class="detail-label">Unidade</div>
            <div class="detail-value">{row['Unidade'] if pd.notna(row['Unidade']) else 'N/A'}</div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="detail-label">Formulário</div>
            <div class="detail-value">{row['Formulário']}</div>
            <div class="detail-label">Subformulário</div>
            <div class="detail-value">{row['Subformulário'] if pd.notna(row['Subformulário']) else 'N/A'}</div>
            <div class="detail-label">Subgrupo</div>
            <div class="detail-value">{row['Subgrupo'] if pd.notna(row['Subgrupo']) else 'N/A'}</div>
            """, unsafe_allow_html=True)
        
        if pd.notna(row['Descrição SINISA']):
            st.markdown(f"""
            <div class="detail-label">Descrição</div>
            <div class="detail-value">{row['Descrição SINISA']}</div>
            """, unsafe_allow_html=True)
        
        if pd.notna(row['Fórmula']):
            st.markdown(f"""
            <div class="detail-label">Fórmula</div>
            <div class="detail-value" style="background: #c0d9f0; padding: 0.5rem; border-radius: 4px; font-family: monospace;">
                {row['Fórmula']}
            </div>
            """, unsafe_allow_html=True)
        
        if pd.notna(row['Referência']):
            referencia_formatada = row['Referência'].replace(', ', '\n').lstrip()
            st.markdown(f"""
            <div class="detail-label">Informações</div>
            <div style="background: #c0d9f0; padding: 0.5rem; border-radius: 4px; border-bottom: 1px solid #c0d9f0; color: #333; font-size: 0.95rem; line-height: 1.5; white-space: pre-line; margin: 0;">
            {referencia_formatada}</div>
            """, unsafe_allow_html=True)
        
        if pd.notna(row['Fonte']):
            st.markdown(f"""
            <div class="detail-label">Fonte</div>
            <div class="detail-value">{row['Fonte']}</div>
            """, unsafe_allow_html=True)
        
        if pd.notna(row['Observação']):
            st.markdown(f"""
            <div class="detail-label">Observação</div>
            <div class="detail-value">{row['Observação']}</div>
            """, unsafe_allow_html=True)
        
        st.markdown("<div style='margin-top: 1rem;'></div>", unsafe_allow_html=True)
        col_close, col_space = st.columns([1, 9])
        with col_close:
            if st.button("✕ Fechar", key=f"close_detail_{idx}", help="Fechar detalhes", use_container_width=True):
                st.session_state[f"show_detail_{idx}"] = False
                st.rerun()
    
    st.markdown("---")
    
    if len(df_filtrado) == 0:
        st.warning("Nenhum resultado encontrado. Tente ajustar os filtros.")
    else:
        col_results, col_export = st.columns([3, 1])
        
        with col_results:
            st.markdown(f"#### Resultados: {len(df_filtrado)} registros")
        
        with col_export:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Glossário')
            buffer.seek(0)
            st.download_button(
                label="📥 Baixar dados em Excel",
                data=buffer,
                file_name=f"glossario_sinisa_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_glossario_{datetime.now().timestamp()}"
            )
        
        st.markdown("")
        
        # Tabela de resultados
        col1, col2, col3, col4, col5, col6, col7, col8 = st.columns([0.6, 3.0, 0.6, 0.6, 1.5, 0.6, 0.8, 0.4], vertical_alignment="center")
        
        with col1:
            st.markdown('<div class="table-header">Código</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="table-header">Nome da Informação</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="table-header">Módulo</div>', unsafe_allow_html=True)
        with col4:
            st.markdown('<div class="table-header">Formulário</div>', unsafe_allow_html=True)
        with col5:
            st.markdown('<div class="table-header">Subformulário</div>', unsafe_allow_html=True)
        with col6:
            st.markdown('<div class="table-header">Unidade</div>', unsafe_allow_html=True)
        with col7:
            st.markdown('<div class="table-header">Área Resp.</div>', unsafe_allow_html=True)
        with col8:
            st.markdown('<div class="table-header">Ação</div>', unsafe_allow_html=True)
        
        for idx, row in df_filtrado.iterrows():
            col1, col2, col3, col4, col5, col6, col7, col8 = st.columns([0.6, 3.0, 0.6, 0.6, 1.5, 0.6, 0.8, 0.4], vertical_alignment="center")
            
            with col1:
                st.markdown(f'<div class="table-row">{row["Código da informação"]}</div>', unsafe_allow_html=True)
            
            with col2:
                st.markdown(f'<div class="table-row">{row["Nome da informação"]}</div>', unsafe_allow_html=True)
            
            with col3:
                st.markdown(f'<div class="table-row">{row["Módulo"]}</div>', unsafe_allow_html=True)
            
            with col4:
                st.markdown(f'<div class="table-row">{row["Formulário"]}</div>', unsafe_allow_html=True)
            
            with col5:
                subform = row['Subformulário'] if pd.notna(row['Subformulário']) else '-'
                st.markdown(f'<div class="table-row">{subform}</div>', unsafe_allow_html=True)
            
            with col6:
                unidade = row['Unidade'] if pd.notna(row['Unidade']) else '-'
                st.markdown(f'<div class="table-row">{unidade}</div>', unsafe_allow_html=True)
            
            with col7:
                area = row['Área Responsável'] if pd.notna(row['Área Responsável']) else '-'
                st.markdown(f'<div class="table-row">{area}</div>', unsafe_allow_html=True)
            
            with col8:                    
                if st.button("🔍", key=f"detail_{idx}", help="Ver detalhes"):
                    st.session_state[f"show_detail_{idx}"] = True
            
            if st.session_state.get(f"show_detail_{idx}", False):
                exibir_detalhes_modal(row, idx)
            
            st.markdown('<div class="table-divider"></div>', unsafe_allow_html=True)

def pagina_glossario_indicadores():
    """Página Glossário - Indicadores"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📈 Indicadores")
    
    st.info("""
    ⏳ **Esta seção está em construção.**
    
    Em breve, você poderá acessar análises e indicadores detalhados do sistema SINISA.
    """)

def pagina_glossario_avisos():
    """Página Glossário - Avisos e Erros"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### ⚠️ Avisos e Erros")
    
    st.info("""
    ⏳ **Esta seção está em construção.**
    
    Em breve, você poderá acessar avisos e erros do sistema SINISA.
    """)

def pagina_relatorios_municipio():
    """Página Relatórios - Gerencial por Município"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 🏘️ Relatório Gerencial por Município")
    
    LARGURA_COD_DESC = 350
    LARGURA_UN = 60
    LARGURA_MES = 60
    
    df_agems = load_dados_agems()
    
    if df_agems is not None and len(df_agems) > 0:
        # Inicializar session state
        if "municipio_filter_rel" not in st.session_state:
            st.session_state.municipio_filter_rel = sorted(df_agems['municipio'].dropna().unique().tolist())[0]
        if "modulo_filter_rel" not in st.session_state:
            st.session_state.modulo_filter_rel = "Todos"
        if "formulario_filter_rel" not in st.session_state:
            st.session_state.formulario_filter_rel = "Todos"
        if "subformulario_filter_rel" not in st.session_state:
            st.session_state.subformulario_filter_rel = "Todos"
        if "subgrupo_filter_rel" not in st.session_state:
            st.session_state.subgrupo_filter_rel = "Todos"
        if "search_unificado_rel" not in st.session_state:
            st.session_state.search_unificado_rel = ""
        
        # Callback para limpar filtros
        def limpar_filtros_callback():
            """Callback para limpar todos os filtros"""
            st.session_state.municipio_filter_rel = sorted(df_agems['municipio'].dropna().unique().tolist())[0]
            st.session_state.modulo_filter_rel = "Todos"
            st.session_state.formulario_filter_rel = "Todos"
            st.session_state.subformulario_filter_rel = "Todos"
            st.session_state.subgrupo_filter_rel = "Todos"
            st.session_state.search_unificado_rel = ""
        
        # Funções para obter valores únicos dos filtros
        @st.cache_data
        def get_municipios_rel():
            return sorted(df_agems['municipio'].dropna().unique().tolist())
        
        def get_filtro_relatorio(coluna, municipio, modulo="Todos", formulario="Todos", subformulario="Todos"):
            """Obtém valores únicos para filtros do relatório"""
            df_temp = df_agems[df_agems['municipio'] == municipio]
            if modulo != "Todos":
                df_temp = df_temp[df_temp['modulo'] == modulo]
            if formulario != "Todos":
                df_temp = df_temp[df_temp['formulario'] == formulario]
            if subformulario != "Todos":
                df_temp = df_temp[df_temp['subformulario_gruopo'] == subformulario]
            return ["Todos"] + sorted(df_temp[coluna].dropna().unique().tolist())
        
        # Linha 1 - Município e botões
        col1, col2, col3 = st.columns([3, 0.5, 0.7])
        
        with col1:
            municipio_selecionado = st.selectbox(
                "🏘️ Município",
                get_municipios_rel(),
                key="municipio_filter_rel"
            )
        
        with col2:
            st.markdown("<div style='padding-top: 10px;'></div>", unsafe_allow_html=True)
            st.button(
                "🔄 Limpar Filtros",
                use_container_width=True,
                key="btn_clear_rel",
                on_click=limpar_filtros_callback
            )
        
        with col3:
            st.markdown("<div style='padding-top: 10px;'></div>", unsafe_allow_html=True)
            export_placeholder = st.empty()
        
        # Linha 2 - Filtros
        col1, col2, col3, col4, col5 = st.columns([1, 1, 1, 1, 1.5])
        
        with col1:
            modulo_selecionado = st.selectbox(
                "Módulo",
                ["Todos"] + sort_modulos(get_filtro_relatorio('modulo', municipio_selecionado)[1:]),
                key="modulo_filter_rel"
            )
        
        with col2:
            formulario_selecionado = st.selectbox(
                "Formulário",
                get_filtro_relatorio('formulario', municipio_selecionado, modulo_selecionado),
                key="formulario_filter_rel"
            )
        
        with col3:
            subformulario_selecionado = st.selectbox(
                "Subformulário",
                get_filtro_relatorio('subformulario_gruopo', municipio_selecionado, modulo_selecionado, formulario_selecionado),
                key="subformulario_filter_rel"
            )
        
        with col4:
            subgrupo_selecionado = st.selectbox(
                "Subgrupo",
                get_filtro_relatorio('subgrupo_palavra_chave', municipio_selecionado, modulo_selecionado, 
                                    formulario_selecionado, subformulario_selecionado),
                key="subgrupo_filter_rel"
            )
        
        with col5:
            busca_unificada = st.text_input(
                "🔎 Buscar (Código ou Descrição)",
                key="search_unificado_rel",
                placeholder="Digite código ou descrição"
            )
        
        # Filtrar dados
        df_municipio = df_agems[df_agems['municipio'] == municipio_selecionado].copy()

        df_municipio = aplicar_filtros(df_municipio, {
            'modulo': modulo_selecionado,
            'formulario': formulario_selecionado,
            'subformulario_gruopo': subformulario_selecionado,
            'subgrupo_palavra_chave': subgrupo_selecionado
        })

        if busca_unificada:
            df_municipio = df_municipio[
                (df_municipio['codigo_da_informacao_indicador'].str.contains(busca_unificada, case=False, na=False)) |
                (df_municipio['nome_da_informacao_indicador'].str.contains(busca_unificada, case=False, na=False))
            ]

        # Botão de exportação
        with export_placeholder.container():
            try:
                todas_colunas_meses = sorted([col for col in df_municipio.columns if '/' in col and len(col) == 7])
                colunas_meses = [mes for mes in todas_colunas_meses 
                                if len(df_municipio[mes].dropna()) > 0]
                
                wb = criar_relatorio_excel_generico(df_municipio, municipio_selecionado, colunas_meses)
                buffer = BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Baixar Relatório em Excel",
                    data=buffer,
                    file_name=f"relatorio_{municipio_selecionado.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_relatorio_municipio_{datetime.now().timestamp()}"
                )
            except Exception as e:
                st.error(f"❌ Erro ao gerar relatório: {str(e)}")

        if len(df_municipio) > 0:
            # Obter colunas de meses
            todas_colunas_meses = sorted([col for col in df_municipio.columns if '/' in col and len(col) == 7])
            colunas_meses = [mes for mes in todas_colunas_meses 
                            if len(df_municipio[mes].dropna()) > 0]
            
            # Cabeçalho do Relatório
            st.markdown("---")
            st.markdown(f"""
            <div class="header-container">
                <h2 style="color: white; margin: 0;">📊 {municipio_selecionado}</h2>
                <p style="color: #e8f0ff; margin: 0.5rem 0 0 0;">Dados AGEMS - {datetime.now().strftime('%d/%m/%Y')}</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Ordenar dados
            df_municipio = df_municipio.sort_values(['modulo', 'codigo_da_informacao_indicador'])
            modulos_unicos = sort_modulos(df_municipio['modulo'].unique().tolist())
            
            # Exibir dados por módulo
            for modulo in modulos_unicos:
                df_modulo = df_municipio[df_municipio['modulo'] == modulo]
                cores = get_modulo_color(modulo)
                
                st.markdown(f"""
                <div style="background-color: {cores['header_color']}; color: white; padding: 0.8rem; border-radius: 4px; margin-top: 1rem; margin-bottom: 0rem;">
                    <h3 style="margin: 0; font-size: 1.2rem;">{modulo}</h3>
                </div>
                """, unsafe_allow_html=True)
                
                # Se for "Informações Complementares", exibir sem agrupar por formulário
                if 'informações complementares' in modulo.lower():
                    html_tabela = criar_tabela_html(df_modulo, colunas_meses, LARGURA_COD_DESC, LARGURA_UN, LARGURA_MES)
                    st.markdown(html_tabela, unsafe_allow_html=True)
                    st.markdown("")
                else:
                    # Para Água e Esgoto, manter a estrutura original com formulário/subformulário
                    for formulario in df_modulo['formulario'].unique():
                        df_formulario = df_modulo[df_modulo['formulario'] == formulario]
                        
                        st.markdown(f"""
                        <div style="background-color: {cores['subheader_color']}; color: white; padding: 0.6rem 0.8rem; margin-top: 0rem; margin-bottom: 0rem; border-radius: 3px;">
                            <h4 style="margin: 0; font-size: 1rem;">{formulario}</h4>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        for subformulario in df_formulario['subformulario_gruopo'].unique():
                            df_subformulario = df_formulario[df_formulario['subformulario_gruopo'] == subformulario]
                            
                            st.markdown(f"""
                            <div style="background-color: {cores['subform_color']}; color: #1f4788; padding: 0.5rem 0.8rem; margin-top: 0rem; margin-bottom: 0rem; border-left: 3px solid {cores['border_color']}; font-weight: 600; font-size: 0.95rem;">
                                {subformulario}
                            </div>
                            """, unsafe_allow_html=True)
                            
                            for subgrupo in df_subformulario['subgrupo_palavra_chave'].unique():
                                df_subgrupo = df_subformulario[df_subformulario['subgrupo_palavra_chave'] == subgrupo]
                                
                                st.markdown(f"""
                                <div style="background-color: {cores['subgrupo_color']}; color: #1f4788; padding: 0.4rem 0.8rem; margin-top: 0rem; margin-bottom: 0rem; border-left: 4px solid {cores['border_color']}; font-weight: 600; font-size: 0.9rem;">
                                    {subgrupo}
                                </div>
                                """, unsafe_allow_html=True)
                                
                                html_tabela = criar_tabela_html(df_subgrupo, colunas_meses, LARGURA_COD_DESC, LARGURA_UN, LARGURA_MES)
                                st.markdown(html_tabela, unsafe_allow_html=True)
                                st.markdown("")
            
            st.markdown("---")
        else:
            st.warning(f"⚠️ Nenhum dado encontrado com os filtros selecionados")
    else:
        st.error("❌ Não foi possível carregar os dados AGEMS")

def pagina_relatorios_consolidado():
    """Página Relatórios - Consolidado"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📊 Relatório Consolidado")
    
    df_sinisa = load_dados_sinisa()
    
    if df_sinisa is not None and len(df_sinisa) > 0:
        # Inicializar session state
        if "modulo_filter_cons" not in st.session_state:
            st.session_state.modulo_filter_cons = "Todos"
        if "informacao_filter_cons" not in st.session_state:
            st.session_state.informacao_filter_cons = "Todos"
        if "formulario_filter_cons" not in st.session_state:
            st.session_state.formulario_filter_cons = "Todos"
        if "subformulario_filter_cons" not in st.session_state:
            st.session_state.subformulario_filter_cons = "Todos"
        if "subgrupo_filter_cons" not in st.session_state:
            st.session_state.subgrupo_filter_cons = "Todos"
        if "area_filter_cons" not in st.session_state:
            st.session_state.area_filter_cons = "Todos"
        if "indicadores_selecionados" not in st.session_state:
            st.session_state.indicadores_selecionados = []
        if "filtros_anteriores" not in st.session_state:
            st.session_state.filtros_anteriores = {
                "modulo": "Todos",
                "informacao": "Todos",
                "formulario": "Todos",
                "subformulario": "Todos",
                "subgrupo": "Todos",
                "area": "Todos"
            }
        
        # Callback para limpar filtros
        def limpar_filtros_consolidado_callback():
            """Callback para limpar todos os filtros do consolidado"""
            st.session_state.modulo_filter_cons = "Todos"
            st.session_state.informacao_filter_cons = "Todos"
            st.session_state.formulario_filter_cons = "Todos"
            st.session_state.subformulario_filter_cons = "Todos"
            st.session_state.subgrupo_filter_cons = "Todos"
            st.session_state.area_filter_cons = "Todos"
        
        # Funções para obter valores únicos dos filtros
        def get_filtro_consolidado(coluna, modulo="Todos", informacao="Todos", formulario="Todos", 
                                  subformulario="Todos", subgrupo="Todos"):
            """Obtém valores únicos para filtros do consolidado"""
            df_temp = df_sinisa.copy()
            if modulo != "Todos":
                df_temp = df_temp[df_temp['modulo'] == modulo]
            if informacao != "Todos":
                df_temp = df_temp[df_temp['informacao_indicador'] == informacao]
            if formulario != "Todos":
                df_temp = df_temp[df_temp['formulario'] == formulario]
            if subformulario != "Todos":
                df_temp = df_temp[df_temp['subformulario_gruopo'] == subformulario]
            if subgrupo != "Todos":
                df_temp = df_temp[df_temp['subgrupo_palavra_chave'] == subgrupo]
            return ["Todos"] + sorted(df_temp[coluna].dropna().unique().tolist())
        
        # Linha 1 - Filtros
        col1, col2, col3, col4, col5, col6 = st.columns([1, 1, 1, 1, 1, 1])
        
        with col1:
            modulo_selecionado_cons = st.selectbox(
                "Módulo",
                ["Todos"] + sort_modulos(get_filtro_consolidado('modulo')[1:]),
                key="modulo_filter_cons"
            )
        
        with col2:
            informacao_selecionada_cons = st.selectbox(
                "Informação/Indicador",
                get_filtro_consolidado('informacao_indicador', modulo_selecionado_cons),
                key="informacao_filter_cons"
            )
        
        with col3:
            formulario_selecionado_cons = st.selectbox(
                "Formulário",
                get_filtro_consolidado('formulario', modulo_selecionado_cons, informacao_selecionada_cons),
                key="formulario_filter_cons"
            )
        
        with col4:
            subformulario_selecionado_cons = st.selectbox(
                "Subformulário",
                get_filtro_consolidado('subformulario_gruopo', modulo_selecionado_cons, informacao_selecionada_cons, formulario_selecionado_cons),
                key="subformulario_filter_cons"
            )
        
        with col5:
            subgrupo_selecionado_cons = st.selectbox(
                "Subgrupo",
                get_filtro_consolidado('subgrupo_palavra_chave', modulo_selecionado_cons, informacao_selecionada_cons, 
                                      formulario_selecionado_cons, subformulario_selecionado_cons),
                key="subgrupo_filter_cons"
            )
        
        with col6:
            area_selecionada_cons = st.selectbox(
                "Área Responsável",
                get_filtro_consolidado('area_responsavel', modulo_selecionado_cons, informacao_selecionada_cons, 
                                      formulario_selecionado_cons, subformulario_selecionado_cons, subgrupo_selecionado_cons),
                key="area_filter_cons"
            )
        
        # Filtrar dados
        df_consolidado = aplicar_filtros(df_sinisa, {
            'modulo': modulo_selecionado_cons,
            'informacao_indicador': informacao_selecionada_cons,
            'formulario': formulario_selecionado_cons,
            'subformulario_gruopo': subformulario_selecionado_cons,
            'subgrupo_palavra_chave': subgrupo_selecionado_cons,
            'area_responsavel': area_selecionada_cons
        })
        
        # Remover linha do Estado
        df_consolidado = df_consolidado[df_consolidado['municipio'] != 'Estado']
        
        # Verificar se os filtros mudaram
        filtros_atuais = {
            "modulo": modulo_selecionado_cons,
            "informacao": informacao_selecionada_cons,
            "formulario": formulario_selecionado_cons,
            "subformulario": subformulario_selecionado_cons,
            "subgrupo": subgrupo_selecionado_cons,
            "area": area_selecionada_cons
        }
        
        if filtros_atuais != st.session_state.filtros_anteriores:
            st.session_state.filtros_anteriores = filtros_atuais.copy()
            st.session_state.indicadores_selecionados = sorted(df_consolidado['codigo_da_informacao_indicador'].unique().tolist())
        
        if len(df_consolidado) == 0:
            st.warning("⚠️ Nenhum dado encontrado com os filtros selecionados")
        else:               
            # Seção de seleção de indicadores
            st.markdown("""
            <div class="detail-section">
                <h3 style="color: #1f4788; margin-top: 0; margin-bottom: 1rem;">📊 Seleção de Informações/Indicadores</h3>
            </div>
            """, unsafe_allow_html=True)
            
            indicadores_disponiveis = sorted(df_consolidado['codigo_da_informacao_indicador'].unique().tolist())
            
            codigo_nome_unidade_map = {}
            for codigo in indicadores_disponiveis:
                row_data = df_consolidado[df_consolidado['codigo_da_informacao_indicador'] == codigo].iloc[0]
                codigo_nome_unidade_map[codigo] = formatar_nome_com_unidade(row_data['nome_da_informacao_indicador'], row_data['unidade'])
            
            indicadores_selecionados = st.multiselect(
                "Selecione as informações/indicadores para incluir no relatório",
                options=indicadores_disponiveis,
                default=st.session_state.indicadores_selecionados,
                format_func=lambda x: f"{x} - {codigo_nome_unidade_map.get(x, '')}",
                key="multiselect_indicadores"
            )
            
            st.session_state.indicadores_selecionados = indicadores_selecionados
            
            if len(indicadores_selecionados) > 0:
                df_consolidado_final = df_consolidado[df_consolidado['codigo_da_informacao_indicador'].isin(indicadores_selecionados)].copy()
            else:
                df_consolidado_final = df_consolidado.copy()           

            st.markdown("<div style='padding-top: 5px;'></div>", unsafe_allow_html=True)

            # Linha de botões - Separada em container para melhor performance
            col1, col2, col3, col4 = st.columns([1, 1, 1, 1], gap="small")


            with col1:
                st.button(
                    "🔄 Limpar Filtros",
                    key="btn_clear_cons",
                    on_click=limpar_filtros_consolidado_callback,
                    use_container_width=True
                )
            
            with col2:
                btn_excel_placeholder = st.empty()
            
            with col3:
                btn_modulo_placeholder = st.empty()
            
            with col4:
                btn_area_placeholder = st.empty()
            
            # Gerar botões sob demanda
            with btn_excel_placeholder.container():
                try:
                    # Converter para tuple para cache funcionar
                    wb = criar_relatorio_consolidado_excel(df_consolidado_final, tuple(indicadores_selecionados))
                    buffer = BytesIO()
                    wb.save(buffer)
                    buffer.seek(0)
                    st.download_button(
                        label="📥 Exportar em Excel",
                        data=buffer,
                        file_name=f"relatorio_consolidado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_relatorio_consolidado_{datetime.now().timestamp()}",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Erro ao gerar relatório: {str(e)}")
            
            with btn_modulo_placeholder.container():
                try:
                    zip_buffer = criar_exportacao_zip_por_modulo(df_sinisa)
                    st.download_button(
                        label="📦 Exportar por Módulo",
                        data=zip_buffer,
                        file_name=f"dados_por_modulo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip",
                        key=f"download_modulo_{datetime.now().timestamp()}",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Erro ao gerar exportação por módulo: {str(e)}")
            
            with btn_area_placeholder.container():
                try:
                    zip_buffer = criar_exportacao_zip_por_area(df_sinisa)
                    st.download_button(
                        label="📦 Exportar por Área",
                        data=zip_buffer,
                        file_name=f"dados_por_area_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip",
                        key=f"download_area_{datetime.now().timestamp()}",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Erro ao gerar exportação por área: {str(e)}")
            
            st.markdown("---")
            
            # Exibir tabela consolidada
            html_tabela_consolidada = criar_tabela_consolidada_html(df_consolidado_final, indicadores_selecionados)
            st.markdown(html_tabela_consolidada, unsafe_allow_html=True)
            
            st.markdown("---")
    else:
        st.error("❌ Não foi possível carregar os dados SINISA")

def pagina_relatorios_verificacao():
    """Página Relatórios - Verificação"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("### ✅ Verificação")
    
    # Carregar dados
    df_glossario = load_data()
    df_correspondencia = load_correspondencia_sigis_sinisa()
    df_dados_sigis = load_dados_sigis()
    
    if df_glossario is None or df_correspondencia is None or df_dados_sigis is None:
        st.error("❌ Não foi possível carregar os dados necessários")
        return
    
    # TRATAMENTO ROBUSTO DE DADOS
    df_correspondencia['codigo_sinisa'] = df_correspondencia['codigo_sinisa'].astype(str).str.strip()
    df_correspondencia['termos_sinisa'] = df_correspondencia['termos_sinisa'].astype(str).str.strip()
    
    def limpar_codigo_sigis(val):
        if pd.isna(val) or val == "" or str(val).strip() == "nan":
            return ""
        val_str = str(val).strip()
        if val_str.endswith('.0'):
            return val_str[:-2]
        return val_str
        
    df_correspondencia['termos_sigis'] = df_correspondencia['termos_sigis'].apply(limpar_codigo_sigis)
    df_dados_sigis['cod_sigis'] = df_dados_sigis['cod_sigis'].apply(limpar_codigo_sigis)
    
    # Funções de filtro
    @st.cache_data
    def get_informacoes_verificacao():
        """Obtém lista de informações únicas do glossário com código + descrição"""
        informacoes_dict = {}
        for _, row in df_glossario.iterrows():
            if pd.notna(row['Código da informação']):
                codigo = str(row['Código da informação']).strip()
                nome = str(row['Nome da informação']).strip()
                informacoes_dict[f"{codigo} - {nome}"] = codigo
        
        opcoes = ["Selecione uma informação"] + sorted(informacoes_dict.keys())
        return opcoes, informacoes_dict
    
    @st.cache_data
    def get_municipios_verificacao():
        """Obtém lista de municípios"""
        municipios = sorted(df_dados_sigis['nome_localidade'].dropna().unique().tolist())
        return municipios
    
    def get_descritivo_informacao(codigo_informacao):
        """Obtém descritivo completo da informação"""
        df_info = df_glossario[df_glossario['Código da informação'].astype(str).str.strip() == codigo_informacao]
        if len(df_info) == 0:
            return None
        return df_info.iloc[0]
    
    def get_componentes_informacao(codigo_selecionado):
        """Obtém os componentes (termos) de uma informação SINISA"""
        # 1. Tentar buscar como código principal (ex: GFI1003)
        df_comp_principal = df_correspondencia[df_correspondencia['codigo_sinisa'] == codigo_selecionado]
        
        if len(df_comp_principal) > 0:
            return df_comp_principal[['termos_sinisa', 'termos_sigis']].drop_duplicates().values.tolist()
            
        # 2. Se não encontrou, tentar buscar como termo (ex: GTA1209)
        df_comp_termo = df_correspondencia[df_correspondencia['termos_sinisa'] == codigo_selecionado]
        
        if len(df_comp_termo) > 0:
            return df_comp_termo[['termos_sinisa', 'termos_sigis']].drop_duplicates().values.tolist()
            
        return []
    
    def get_descricao_termo(termo_sinisa):
        """Obtém descrição de um termo SINISA"""
        df_termo = df_glossario[df_glossario['Código da informação'].astype(str).str.strip() == termo_sinisa]
        if len(df_termo) == 0:
            return ""
        return str(df_termo.iloc[0]['Nome da informação']).strip()
    
    def criar_tabela_verificacao(codigo_sinisa, municipio_selecionado):
        """Cria tabela com componentes e valores dos últimos 12 meses"""
        # Obter componentes
        componentes = get_componentes_informacao(codigo_sinisa)
        
        if not componentes:
            st.warning(f"Nenhum componente encontrado para {codigo_sinisa} na tabela de correspondência.")
            return None, None
        
        # Filtrar dados SIGIS para o município
        df_municipio = df_dados_sigis[df_dados_sigis['nome_localidade'] == municipio_selecionado].copy()
        
        if len(df_municipio) == 0:
            st.warning(f"Nenhum dado encontrado para o município {municipio_selecionado}.")
            return None, None
        
        # Converter data para datetime
        df_municipio['data'] = pd.to_datetime(df_municipio['data'])
        
        # Obter últimos 12 meses
        data_maxima = df_municipio['data'].max()
        data_minima = data_maxima - pd.DateOffset(months=11)
        df_municipio = df_municipio[(df_municipio['data'] >= data_minima) & (df_municipio['data'] <= data_maxima)]
        
        # Criar lista de meses (MM/AAAA)
        meses_unicos = sorted(df_municipio['data'].dt.to_period('M').unique())
        meses_formatados = [f"{mes.month:02d}/{mes.year}" for mes in meses_unicos]
        
        # Informação principal
        info_principal = get_descritivo_informacao(codigo_sinisa)
        descricao_principal = info_principal['Nome da informação'] if info_principal is not None else ""
        
        # Construir tabela HTML
        html = '<div style="overflow-x: auto; margin: 1rem 0; border-radius: 4px; border: 1px solid #d0d8e0;">'
        html += '<table style="width: 100%; border-collapse: collapse; font-size: 0.9rem;">'
        
        # Cabeçalho da tabela
        html += '<thead><tr style="background-color: #0B3040; color: white;">'
        html += '<th style="padding: 0.5rem 0.75rem; text-align: left; border: 1px solid #d0d8e0; font-weight: 600; min-width: 500px; height: 40px;">Informação</th>'
        for mes in meses_formatados:
            html += f'<th style="padding: 0.5rem 0.75rem; text-align: center; border: 1px solid #d0d8e0; font-weight: 600; white-space: nowrap; height: 40px;">{mes}</th>'
        html += '</tr></thead>'
        
        # Dados
        html += '<tbody>'
        
        # Armazenar dados para o gráfico e download
        dados_grafico = {
            'informacao': [],
            'descricao': [],
            'mes': [],
            'mes_formatado': [],
            'valor': [],
            'nivel': [],
            'cor': []
        }
        
        # Dicionário de cores por nível
        cores_por_nivel = {
            'principal': '#1f4788',
            'nivel1': '#5a8ac4',
            'sigis': '#a8c5e8'
        }
        
        # Linha 1: Informação Principal (GFI1003 ou GTA1209) - AZUL CLARO
        html += f'<tr style="background-color: #ADD8E6;">'
        html += f'<td style="padding: 0.75rem; border: 1px solid #d0d8e0; font-weight: 600; word-wrap: break-word; white-space: normal;">{codigo_sinisa} - {descricao_principal}</td>'
        
        # Valores para a informação principal (soma dos componentes)
        for mes, mes_fmt in zip(meses_unicos, meses_formatados):
            valor_total = 0
            for termo_sinisa, termo_sigis in componentes:
                if termo_sigis:
                    df_termo = df_municipio[df_municipio['cod_sigis'] == termo_sigis]
                    df_mes = df_termo[df_termo['data'].dt.to_period('M') == mes]
                    if len(df_mes) > 0:
                        try:
                            valor_total += float(df_mes['valor'].values[0])
                        except:
                            pass
            
            valor_formatado = f"{valor_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if valor_total > 0 else "0,00"
            html += f'<td style="padding: 0.75rem; border: 1px solid #d0d8e0; text-align: right; font-weight: 600;">{valor_formatado}</td>'
            
            # Adicionar ao gráfico
            dados_grafico['informacao'].append(codigo_sinisa)
            dados_grafico['descricao'].append(f"{codigo_sinisa} - {descricao_principal}")
            dados_grafico['mes'].append(str(mes))
            dados_grafico['mes_formatado'].append(mes_fmt)
            dados_grafico['valor'].append(valor_total)
            dados_grafico['nivel'].append('principal')
            dados_grafico['cor'].append(cores_por_nivel['principal'])
        
        html += '</tr>'
        
        # Verificar se a informação selecionada é ela mesma um termo (como GTA1209)
        is_termo_direto = all(t_sinisa == codigo_sinisa for t_sinisa, _ in componentes)
        
        # Se for um termo direto (GTA1209), não precisamos repetir a linha do SINISA, vamos direto para os SIGIS
        if is_termo_direto:
            for idx, (termo_sinisa, termo_sigis) in enumerate(componentes):
                
                if termo_sigis:
                    df_sigis_desc = df_municipio[df_municipio['cod_sigis'] == termo_sigis]
                    desc_sigis = df_sigis_desc['desc_sigis'].values[0] if len(df_sigis_desc) > 0 else f"{termo_sigis}"
                    
                    # BRANCO para SIGIS
                    html += f'<tr style="background-color: #FFFFFF;">'
                    html += f'<td style="padding: 0.75rem; border: 1px solid #d0d8e0; font-weight: 400; padding-left: 2rem; color: #333; font-size: 0.85rem; word-wrap: break-word; white-space: normal;">SIGIS: {desc_sigis}</td>'
                    
                    for mes, mes_fmt in zip(meses_unicos, meses_formatados):
                        df_mes = df_sigis_desc[df_sigis_desc['data'].dt.to_period('M') == mes]
                        valor = df_mes['valor'].values[0] if len(df_mes) > 0 else "-"
                        
                        if valor != "-":
                            try:
                                valor_float = float(valor)
                                valor_formatado = f"{valor_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                            except:
                                valor_formatado = str(valor)
                        else:
                            valor_formatado = "-"
                        
                        html += f'<td style="padding: 0.75rem; border: 1px solid #d0d8e0; text-align: right;">{valor_formatado}</td>'
                        
                        # Adicionar ao gráfico
                        if valor != "-":
                            dados_grafico['informacao'].append(f"SIGIS: {termo_sigis}")
                            dados_grafico['descricao'].append(f"SIGIS: {desc_sigis}")
                            dados_grafico['mes'].append(str(mes))
                            dados_grafico['mes_formatado'].append(mes_fmt)
                            dados_grafico['valor'].append(float(valor))
                            dados_grafico['nivel'].append('sigis')
                            dados_grafico['cor'].append(cores_por_nivel['sigis'])
                    
                    html += '</tr>'
        
        # Se for uma informação pai (GFI1003), mostramos a hierarquia completa (SINISA -> SIGIS)
        else:
            # Agrupar por termo_sinisa para não repetir a linha do SINISA se houver múltiplos SIGIS
            termos_agrupados = {}
            for t_sinisa, t_sigis in componentes:
                if t_sinisa not in termos_agrupados:
                    termos_agrupados[t_sinisa] = []
                if t_sigis:
                    termos_agrupados[t_sinisa].append(t_sigis)
            
            for termo_sinisa, lista_sigis in termos_agrupados.items():
                descricao_termo = get_descricao_termo(termo_sinisa)
                
                # CINZA para componentes de nível 1 (SINISA)
                html += f'<tr style="background-color: #D3D3D3;">'
                html += f'<td style="padding: 0.75rem; border: 1px solid #d0d8e0; font-weight: 500; padding-left: 2rem; word-wrap: break-word; white-space: normal;">{termo_sinisa} - {descricao_termo}</td>'
                
                # Valores por mês (soma dos SIGIS filhos)
                for mes, mes_fmt in zip(meses_unicos, meses_formatados):
                    valor_total_termo = 0
                    for t_sigis in lista_sigis:
                        df_termo = df_municipio[df_municipio['cod_sigis'] == t_sigis]
                        df_mes = df_termo[df_termo['data'].dt.to_period('M') == mes]
                        if len(df_mes) > 0:
                            try:
                                valor_total_termo += float(df_mes['valor'].values[0])
                            except:
                                pass
                    
                    valor_formatado = f"{valor_total_termo:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if valor_total_termo > 0 else "0,00"
                    html += f'<td style="padding: 0.75rem; border: 1px solid #d0d8e0; text-align: right;">{valor_formatado}</td>'
                    
                    # Adicionar ao gráfico
                    dados_grafico['informacao'].append(termo_sinisa)
                    dados_grafico['descricao'].append(f"{termo_sinisa} - {descricao_termo}")
                    dados_grafico['mes'].append(str(mes))
                    dados_grafico['mes_formatado'].append(mes_fmt)
                    dados_grafico['valor'].append(valor_total_termo)
                    dados_grafico['nivel'].append('nivel1')
                    dados_grafico['cor'].append(cores_por_nivel['nivel1'])
                
                html += '</tr>'
                
                # Linhas dos filhos SIGIS (ex: 8747) - BRANCO
                for t_sigis in lista_sigis:
                    df_sigis_desc = df_municipio[df_municipio['cod_sigis'] == t_sigis]
                    desc_sigis = df_sigis_desc['desc_sigis'].values[0] if len(df_sigis_desc) > 0 else f"{t_sigis}"
                    
                    html += f'<tr style="background-color: #FFFFFF;">'
                    html += f'<td style="padding: 0.75rem; border: 1px solid #d0d8e0; font-weight: 400; padding-left: 4rem; color: #333; font-size: 0.85rem; word-wrap: break-word; white-space: normal;">SIGIS: {desc_sigis}</td>'
                    
                    for mes, mes_fmt in zip(meses_unicos, meses_formatados):
                        df_mes = df_sigis_desc[df_sigis_desc['data'].dt.to_period('M') == mes]
                        valor = df_mes['valor'].values[0] if len(df_mes) > 0 else "-"
                        
                        if valor != "-":
                            try:
                                valor_float = float(valor)
                                valor_formatado = f"{valor_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                            except:
                                valor_formatado = str(valor)
                        else:
                            valor_formatado = "-"
                        
                        html += f'<td style="padding: 0.75rem; border: 1px solid #d0d8e0; text-align: right;">{valor_formatado}</td>'
                        
                        # Adicionar ao gráfico
                        if valor != "-":
                            dados_grafico['informacao'].append(f"SIGIS: {t_sigis}")
                            dados_grafico['descricao'].append(f"SIGIS: {desc_sigis}")
                            dados_grafico['mes'].append(str(mes))
                            dados_grafico['mes_formatado'].append(mes_fmt)
                            dados_grafico['valor'].append(float(valor))
                            dados_grafico['nivel'].append('sigis')
                            dados_grafico['cor'].append(cores_por_nivel['sigis'])
                    
                    html += '</tr>'
        
        html += '</tbody></table></div>'
        st.markdown(html, unsafe_allow_html=True)
        
        return dados_grafico, html
    
    def gerar_excel_formatado(dados_grafico, html_tabela, codigo_sinisa, municipio_selecionado):
        """Gera Excel com a mesma formatação e estilo da tabela"""
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Verificação"
        
        # Obter dados únicos
        df_dados = pd.DataFrame(dados_grafico)
        
        # Cabeçalho - USAR mes_formatado (MM/AAAA)
        meses_unicos = sorted(df_dados['mes_formatado'].unique())
        
        # Linha 1: Cabeçalho
        ws['A1'] = "Informação"
        for idx, mes in enumerate(meses_unicos, start=2):
            ws.cell(row=1, column=idx).value = mes
        
        # Estilos
        header_fill = PatternFill(start_color="0B3040", end_color="0B3040", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        
        azul_claro_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        cinza_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        branco_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        border = Border(
            left=Side(style='thin', color='d0d8e0'),
            right=Side(style='thin', color='d0d8e0'),
            top=Side(style='thin', color='d0d8e0'),
            bottom=Side(style='thin', color='d0d8e0')
        )
        
        # Aplicar estilo ao cabeçalho
        for col in range(1, len(meses_unicos) + 2):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # Dados
        row_num = 2
        informacoes_unicas = df_dados['descricao'].unique()
        
        for desc in informacoes_unicas:
            df_info = df_dados[df_dados['descricao'] == desc]
            nivel = df_info['nivel'].values[0]
            
            # Determinar cor baseado no nível
            if nivel == 'principal':
                fill = azul_claro_fill
                font = Font(bold=True, size=11)
            elif nivel == 'nivel1':
                fill = cinza_fill
                font = Font(bold=True, size=10)
            else:  # sigis
                fill = branco_fill
                font = Font(size=10)
            
            # Coluna de informação - DESCRIÇÃO COMPLETA
            ws.cell(row=row_num, column=1).value = desc
            ws.cell(row=row_num, column=1).fill = fill
            ws.cell(row=row_num, column=1).font = font
            ws.cell(row=row_num, column=1).border = border
            ws.cell(row=row_num, column=1).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # Valores por mês - USAR mes_formatado
            for idx, mes in enumerate(meses_unicos, start=2):
                valor = df_info[df_info['mes_formatado'] == mes]['valor'].values
                if len(valor) > 0:
                    cell_value = valor[0]
                    ws.cell(row=row_num, column=idx).value = cell_value
                    ws.cell(row=row_num, column=idx).number_format = '#,##0.00'
                else:
                    ws.cell(row=row_num, column=idx).value = 0
                
                cell = ws.cell(row=row_num, column=idx)
                cell.fill = fill
                cell.font = font
                cell.border = border
                cell.alignment = Alignment(horizontal='right', vertical='center')
            
            row_num += 1
        
        # Ajustar largura das colunas
        ws.column_dimensions['A'].width = 80
        for idx in range(2, len(meses_unicos) + 2):
            ws.column_dimensions[get_column_letter(idx)].width = 15
        
        # Salvar em buffer
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        return buffer
    
    # Interface
    col1, col2 = st.columns([2, 3])
    
    # Obter opções de informações
    opcoes_info, info_dict = get_informacoes_verificacao()
    
    with col1:
        informacao_selecionada_display = st.selectbox(
            "Selecione a Informação/Indicador",
            opcoes_info,
            key="informacao_verif_select",
            index=0
        )
        # Extrair código da seleção
        if informacao_selecionada_display != "Selecione uma informação":
            informacao_selecionada = info_dict[informacao_selecionada_display]
        else:
            informacao_selecionada = None
    
    with col2:
        municipio_selecionado = st.selectbox(
            "Selecione o Município",
            get_municipios_verificacao(),
            key="municipio_verif_select"
        )
    
    st.markdown("---")
    
    # Exibir descritivo da informação selecionada
    if informacao_selecionada is not None:
        row = get_descritivo_informacao(informacao_selecionada)
        
        if row is not None:
            st.markdown("#### 📖 Detalhes da Informação")
            exibir_detalhes_modal(row, "verif_0")
            st.markdown("---")
            
            # Tabela de verificação
            col_titulo, col_botao = st.columns([1, 0.3], vertical_alignment="center")
            
            with col_titulo:
                st.markdown("#### 📊 Composição da Informação")
            
            with col_botao:
                pass
            
            dados_grafico, html_tabela = criar_tabela_verificacao(informacao_selecionada, municipio_selecionado)
            
            if dados_grafico is not None:
                # Botão de download alinhado à direita
                with col_botao:
                    # Gerar Excel formatado
                    buffer_excel = gerar_excel_formatado(dados_grafico, html_tabela, informacao_selecionada, municipio_selecionado)
                    
                    st.download_button(
                        label="📥 Baixar",
                        data=buffer_excel,
                        file_name=f"verificacao_{informacao_selecionada}_{municipio_selecionado}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                st.markdown("---")
                
                # ============================================================
                # GRÁFICO OPÇÃO 1: COLUNAS SIGIS + LINHAS HIERARQUIA
                # ============================================================
                st.markdown("#### 📈 Gráfico de Evolução")
                
                df_grafico = pd.DataFrame(dados_grafico)
                
                # Separar dados
                df_sigis = df_grafico[df_grafico['nivel'] == 'sigis'].copy()
                df_nivel1 = df_grafico[df_grafico['nivel'] == 'nivel1'].copy()
                df_principal = df_grafico[df_grafico['nivel'] == 'principal'].copy()
                
                # Obter valores mín e máx para escala uniforme
                todos_valores = pd.concat([df_sigis['valor'], df_nivel1['valor'], df_principal['valor']])
                valor_min = todos_valores.min()
                valor_max = todos_valores.max()
                
                # Criar figura com eixos secundários
                fig1 = make_subplots(specs=[[{"secondary_y": True}]])
                
                # Paleta de cores para SIGIS com transparência
                paleta_cores = [
                    'rgba(255, 107, 107, 0.6)',      # Vermelho coral
                    'rgba(78, 205, 196, 0.6)',       # Turquesa
                    'rgba(69, 183, 209, 0.6)',       # Azul céu
                    'rgba(255, 160, 122, 0.6)',      # Salmão claro
                    'rgba(152, 216, 200, 0.6)',      # Menta
                    'rgba(247, 220, 111, 0.6)',      # Amarelo suave
                    'rgba(187, 143, 206, 0.6)',      # Roxo suave
                    'rgba(133, 193, 226, 0.6)',      # Azul claro
                    'rgba(248, 184, 139, 0.6)',      # Pêssego
                    'rgba(169, 223, 191, 0.6)',      # Verde suave
                    'rgba(241, 148, 138, 0.6)',      # Rosa coral
                    'rgba(215, 189, 226, 0.6)'       # Lavanda
                ]
                
                # Adicionar colunas empilhadas SIGIS
                sigis_unicos = df_sigis['descricao'].unique()
                for idx, sigis in enumerate(sigis_unicos):
                    df_sigis_item = df_sigis[df_sigis['descricao'] == sigis]
                    cor = paleta_cores[idx % len(paleta_cores)]
                    
                    fig1.add_trace(
                        go.Bar(
                            x=df_sigis_item['mes_formatado'],
                            y=df_sigis_item['valor'],
                            name=sigis,
                            marker=dict(color=cor),
                            hovertemplate='<b>%{fullData.name}</b><br>Valor: %{y:,.2f}<extra></extra>',
                            showlegend=True
                        ),
                        secondary_y=False
                    )
                
                # Adicionar linhas para Nível 1 (SINISA intermediários)
                nivel1_unicos = df_nivel1['descricao'].unique()
                for nivel1 in nivel1_unicos:
                    df_nivel1_item = df_nivel1[df_nivel1['descricao'] == nivel1]
                    
                    fig1.add_trace(
                        go.Scatter(
                            x=df_nivel1_item['mes_formatado'],
                            y=df_nivel1_item['valor'],
                            mode='lines',
                            name=nivel1,
                            line=dict(color='rgba(90, 138, 196, 0.8)', width=2, dash='dash', shape='spline'),
                            hovertemplate='<b>%{fullData.name}</b><br>Valor: %{y:,.2f}<extra></extra>',
                            showlegend=True
                        ),
                        secondary_y=True
                    )
                
                # Adicionar linha para Principal
                fig1.add_trace(
                    go.Scatter(
                        x=df_principal['mes_formatado'],
                        y=df_principal['valor'],
                        mode='lines',
                        name=f"{informacao_selecionada} (Total)",
                        line=dict(color='rgba(31, 71, 136, 1)', width=3, shape='spline'),
                        hovertemplate='<b>%{fullData.name}</b><br>Valor: %{y:,.2f}<extra></extra>',
                        showlegend=True
                    ),
                    secondary_y=True
                )
                
                fig1.update_layout(
                    title=f"Composição de {informacao_selecionada} - {municipio_selecionado}",
                    xaxis_title="",
                    barmode='stack',
                    hovermode='x unified',
                    height=650,
                    template='plotly_white',
                    font=dict(size=11, family="Arial"),
                    showlegend=True,
                    legend=dict(
                        orientation="v",
                        yanchor="top",
                        y=0.99,
                        xanchor="right",
                        x=0.99,
                        bgcolor="rgba(255, 255, 255, 0.9)",
                        bordercolor="rgba(0, 0, 0, 0.2)",
                        borderwidth=1,
                        font=dict(size=10)
                    ),
                    margin=dict(l=60, r=60, t=100, b=60),
                    plot_bgcolor='white',
                    paper_bgcolor='white'
                )
                
                # Configurar eixos Y com mesma escala e SEM VALORES
                fig1.update_yaxes(
                    range=[valor_min * 0.95, valor_max * 1.05],
                    secondary_y=False,
                    gridcolor='rgba(200, 200, 200, 0.2)',
                    showticklabels=False
                )
                fig1.update_yaxes(
                    range=[valor_min * 0.95, valor_max * 1.05],
                    secondary_y=True,
                    gridcolor='rgba(200, 200, 200, 0.2)',
                    showticklabels=False
                )
                
                fig1.update_xaxes(gridcolor='rgba(200, 200, 200, 0.2)')
                
                st.plotly_chart(fig1, use_container_width=True)
    
    # Seção de Avisos e Erros
    st.markdown("---")
    st.markdown("#### ⚠️ Avisos e Erros")
    st.info("⏳ **Esta seção está em construção.**")

def pagina_infraestrutura():
    """Página Infraestrutura"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 🏗️ Infraestrutura")
    
    st.info("""
    ⏳ **Esta seção está em construção.**
    
    Em breve, você poderá acessar informações sobre a infraestrutura e configurações do sistema SINISA.
    """)

def pagina_relatorios_indicadores_ana():
    """Página Relatórios - Indicadores ANA"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("### 💧 Indicadores ANA")
    st.info("""
    ⏳ **Esta seção está em fase de construção.**
    Em breve, você poderá acessar os relatórios de Indicadores da ANA.
    """)

def pagina_relatorios_rad_agems():
    """Página Relatórios - RAD (AGEMS)"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("### 📈 RAD (AGEMS)")
    st.info("""
    ⏳ **Esta seção está em fase de construção.**
    Em breve, você poderá acessar os relatórios RAD da AGEMS.
    """)

def pagina_contrato_programa():
    """Página Contrato Programa"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("### 📄 Contrato Programa")
    st.info("""
    ⏳ **Esta seção está em fase de construção.**
    Em breve, você poderá acessar as informações sobre Contratos de Programa.
    """)

def pagina_investimento():
    """Página Investimento"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("### 💰 Investimento")
    st.info("""
    ⏳ **Esta seção está em fase de construção.**
    Em breve, você poderá acessar os dados e relatórios de Investimentos.
    """)    

# ============================================================================
# ESTILOS GLOBAIS PARA SIDEBAR - VERSÃO FINAL OTIMIZADA
# ============================================================================

SIDEBAR_STYLES = """
<style>
    /* ========== SIDEBAR CONTAINER ========== */
    [data-testid="stSidebar"] {
        background-color: #f5f7fa !important;
        padding: 0 !important;
    }
    
    /* ========== SIDEBAR CONTENT ========== */
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
        padding: 1rem !important;
        gap: 0.5rem !important;
    }
    
    /* ========== TÍTULOS (H3) ========== */
    [data-testid="stSidebar"] h3 {
        color: #1f4788 !important;
        font-size: 0.9rem !important;
        font-weight: 600 !important;
        margin: 0.75rem 0 0.5rem 0 !important;
        padding: 0 !important;
        text-align: left !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    
    /* ========== EXPANDERS (TÓPICOS) ========== */
    [data-testid="stSidebar"] .streamlit-expanderHeader {
        background-color: #e8eef5 !important;
        border: 1px solid #d0d8e0 !important;
        border-radius: 6px !important;
        padding: 0.5rem 0.8rem !important;
        margin-bottom: 0.5rem !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
        color: #1f4788 !important;
        text-align: left !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    
    [data-testid="stSidebar"] .streamlit-expanderHeader:hover {
        background-color: #d4e4f7 !important;
        border-color: #2d5aa8 !important;
    }
    
    [data-testid="stSidebar"] .streamlit-expanderHeader p {
        margin: 0 !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        font-size: 0.85rem !important;
        text-align: left !important;
    }
    
    /* ========== CONTEÚDO DENTRO DO EXPANDER ========== */
    [data-testid="stSidebar"] .streamlit-expanderContent {
        padding: 0.5rem 0 !important;
        border-left: 3px solid #2d5aa8 !important;
        margin-left: 0.5rem !important;
    }
    
    /* ========== BOTÕES NO SIDEBAR ========== */
    [data-testid="stSidebar"] button {
        width: 100% !important;
        text-align: left !important;
        padding: 0.5rem 0.8rem !important;
        margin: 0.25rem 0 !important;
        border: 1px solid #d0d8e0 !important;
        border-radius: 4px !important;
        background-color: #ffffff !important;
        color: #333 !important;
        font-size: 0.85rem !important;
        font-weight: 500 !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        transition: all 0.2s ease !important;
        display: flex !important;
        align-items: center !important;
        justify-content: flex-start !important;
    }
    
    [data-testid="stSidebar"] button:hover {
        background-color: #e8eef5 !important;
        border-color: #2d5aa8 !important;
        color: #1f4788 !important;
    }
    
    [data-testid="stSidebar"] button:active {
        background-color: #d4e4f7 !important;
        border-color: #1f4788 !important;
    }
    
    /* ========== TEXTO DENTRO DOS BOTÕES ========== */
    [data-testid="stSidebar"] button p {
        margin: 0 !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        font-size: 0.85rem !important;
        text-align: left !important;
    }
    
    /* ========== ÍCONES DENTRO DOS BOTÕES ========== */
    [data-testid="stSidebar"] button svg {
        margin-right: 0.5rem !important;
        flex-shrink: 0 !important;
    }
    
    /* ========== MARKDOWN NO SIDEBAR ========== */
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        text-align: left !important;
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        margin: 0 !important;
        font-size: 0.85rem !important;
        text-align: left !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    
    /* ========== ESPAÇAMENTO GERAL ========== */
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div {
        margin: 0 !important;
    }
    
    /* ========== REMOVER ESPAÇOS EXTRAS ========== */
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div > div {
        margin: 0 !important;
        padding: 0 !important;
    }
</style>
"""

st.markdown(SIDEBAR_STYLES, unsafe_allow_html=True)

# ============================================================================
# ESTRUTURA DE NAVEGAÇÃO COM SIDEBAR - VERSÃO CORRIGIDA
# ============================================================================

def pagina_glossario_informacoes_com_stats():
    """Página Glossário - Informações com Estatísticas"""
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">📚 SINISA</h1>
        <p class="header-subtitle">Sistema Nacional de Informações sobre Saneamento</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📋 Glossário - Informações")
    
    # Inicializar session state
    if "modulo_filter" not in st.session_state:
        st.session_state.modulo_filter = "Todos"
    if "formulario_filter" not in st.session_state:
        st.session_state.formulario_filter = "Todos"
    if "subformulario_filter" not in st.session_state:
        st.session_state.subformulario_filter = "Todos"
    if "subgrupo_filter" not in st.session_state:
        st.session_state.subgrupo_filter = "Todos"
    if "area_filter" not in st.session_state:
        st.session_state.area_filter = "Todos"
    if "search_codigo" not in st.session_state:
        st.session_state.search_codigo = ""
    if "search_descricao" not in st.session_state:
        st.session_state.search_descricao = ""
    if "show_stats" not in st.session_state:
        st.session_state.show_stats = False
    
    # Funções de filtro
    @st.cache_data
    def get_modulos():
        modulos = ["Todos"] + sort_modulos(df['Módulo'].dropna().unique().tolist())
        return modulos
    
    def get_filtro_glossario(coluna, modulo="Todos", formulario="Todos", subformulario="Todos", subgrupo="Todos"):
        """Obtém valores únicos para filtros do glossário"""
        df_temp = df.copy()
        if modulo != "Todos":
            df_temp = df_temp[df_temp['Módulo'] == modulo]
        if formulario != "Todos":
            df_temp = df_temp[df_temp['Formulário'] == formulario]
        if subformulario != "Todos":
            df_temp = df_temp[df_temp['Subformulário'] == subformulario]
        if subgrupo != "Todos":
            df_temp = df_temp[df_temp['Subgrupo'] == subgrupo]
        return ["Todos"] + sorted(df_temp[coluna].dropna().unique().tolist())
    
    # Linha 1 - Filtros
    col1, col2, col3, col4, col5, col6 = st.columns([1.5, 1.5, 1.5, 1.5, 1.5, 1], gap="small")
    
    with col6:
        st.markdown("<div style='padding-top: 12px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Limpar Filtros", key="btn_clear_filters", use_container_width=True):
            st.session_state.modulo_filter = "Todos"
            st.session_state.formulario_filter = "Todos"
            st.session_state.subformulario_filter = "Todos"
            st.session_state.subgrupo_filter = "Todos"
            st.session_state.area_filter = "Todos"
            st.session_state.search_codigo = ""
            st.session_state.search_descricao = ""
            st.rerun()
    
    with col1:
        modulo_selecionado = st.selectbox("Módulo", get_modulos(), key="modulo_filter", index=0)
    
    with col2:
        formulario_selecionado = st.selectbox("Formulário", 
            get_filtro_glossario('Formulário', modulo_selecionado), key="formulario_filter", index=0)
    
    with col3:
        subformulario_selecionado = st.selectbox("Subformulário", 
            get_filtro_glossario('Subformulário', modulo_selecionado, formulario_selecionado), 
            key="subformulario_filter", index=0)
    
    with col4:
        subgrupo_selecionado = st.selectbox("Subgrupo", 
            get_filtro_glossario('Subgrupo', modulo_selecionado, formulario_selecionado, subformulario_selecionado), 
            key="subgrupo_filter", index=0)
    
    with col5:
        area_selecionada = st.selectbox("Área Responsável", 
            get_filtro_glossario('Área Responsável', modulo_selecionado, formulario_selecionado, 
                                subformulario_selecionado, subgrupo_selecionado), 
            key="area_filter", index=0)
    
    # Linha 2 - Buscas
    col_codigo, col_descricao = st.columns([1.5, 4.5], gap="small")
    with col_codigo:
        busca_codigo = st.text_input("🔎 Buscar por Código", key="search_codigo")
    with col_descricao:
        busca_descricao = st.text_input("🔎 Buscar por Descrição", key="search_descricao")
    
    # Aplicar filtros
    df_filtrado = aplicar_filtros(df, {
        'Módulo': modulo_selecionado,
        'Formulário': formulario_selecionado,
        'Subformulário': subformulario_selecionado,
        'Subgrupo': subgrupo_selecionado,
        'Área Responsável': area_selecionada
    })
    
    if busca_codigo:
        df_filtrado = df_filtrado[df_filtrado['Código da informação'].str.contains(busca_codigo, case=False, na=False)]
    if busca_descricao:
        df_filtrado = df_filtrado[df_filtrado['Nome da informação'].str.contains(busca_descricao, case=False, na=False)]   
    
    # Função de exibição de estatísticas modal
    def exibir_estatisticas_modal(df_filtrado):
        """Exibe estatísticas dos dados filtrados"""
        st.markdown(f"""
        <div class="detail-section">
            <h3 style="color: #1f4788; margin-top: 0; margin-bottom: 1rem;">📊 Análise dos Dados</h3>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3, col4 = st.columns(4, gap="small")
        
        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">📋 Total</div>
                <div class="metric-value">{len(df_filtrado)}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">📦 Módulos</div>
                <div class="metric-value">{df_filtrado['Módulo'].nunique()}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">📝 Formulários</div>
                <div class="metric-value">{df_filtrado['Formulário'].nunique()}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">🏷️ Subgrupos</div>
                <div class="metric-value">{df_filtrado['Subgrupo'].nunique() if 'Subgrupo' in df_filtrado.columns else 0}</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Gráficos
        col1, col2, col3, col4 = st.columns(4, gap="small")
        
        with col1:
            if len(df_filtrado) > 0:
                modulo_count = df_filtrado['Módulo'].value_counts().head(6)
                fig = go.Figure(data=[go.Bar(x=modulo_count.values, y=modulo_count.index, orientation='h', 
                    marker=dict(color=modulo_count.values, colorscale='Blues', line=dict(color='#1f4788', width=1)), 
                    text=modulo_count.values, textposition='auto', hovertemplate='<b>%{y}</b><br>%{x}<extra></extra>')])
                fig.update_layout(height=250, margin=dict(l=100, r=10, t=10, b=10), plot_bgcolor='rgba(245, 247, 250, 0.5)', 
                    paper_bgcolor='rgba(255, 255, 255, 0)', xaxis=dict(showgrid=True, gridwidth=1, gridcolor='#e0e0e0', showticklabels=False), 
                    yaxis=dict(showgrid=False), font=dict(family='Segoe UI', size=9, color='#333'), hovermode='closest', showlegend=False)
                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        
        with col2:
            if len(df_filtrado) > 0:
                formulario_count = df_filtrado['Formulário'].value_counts().head(6)
                fig = go.Figure(data=[go.Bar(x=formulario_count.values, y=formulario_count.index, orientation='h', 
                    marker=dict(color=formulario_count.values, colorscale='Greens', line=dict(color='#2d5aa8', width=1)), 
                    text=formulario_count.values, textposition='auto', hovertemplate='<b>%{y}</b><br>%{x}<extra></extra>')])
                fig.update_layout(height=250, margin=dict(l=100, r=10, t=10, b=10), plot_bgcolor='rgba(245, 247, 250, 0.5)', 
                    paper_bgcolor='rgba(255, 255, 255, 0)', xaxis=dict(showgrid=True, gridwidth=1, gridcolor='#e0e0e0', showticklabels=False), 
                    yaxis=dict(showgrid=False), font=dict(family='Segoe UI', size=9, color='#333'), hovermode='closest', showlegend=False)
                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        
        with col3:
            if len(df_filtrado) > 0 and 'Subgrupo' in df_filtrado.columns:
                subgrupo_count = df_filtrado['Subgrupo'].value_counts().head(6)
                fig = go.Figure(data=[go.Bar(x=subgrupo_count.values, y=subgrupo_count.index, orientation='h', 
                    marker=dict(color=subgrupo_count.values, colorscale='Purples', line=dict(color='#6b21a8', width=1)), 
                    text=subgrupo_count.values, textposition='auto', hovertemplate='<b>%{y}</b><br>%{x}<extra></extra>')])
                fig.update_layout(height=250, margin=dict(l=100, r=10, t=10, b=10), plot_bgcolor='rgba(245, 247, 250, 0.5)', 
                    paper_bgcolor='rgba(255, 255, 255, 0)', xaxis=dict(showgrid=True, gridwidth=1, gridcolor='#e0e0e0', showticklabels=False), 
                    yaxis=dict(showgrid=False), font=dict(family='Segoe UI', size=9, color='#333'), hovermode='closest', showlegend=False)
                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        
        with col4:
            if len(df_filtrado) > 0 and 'Área Responsável' in df_filtrado.columns:
                area_count = df_filtrado['Área Responsável'].value_counts().head(6)
                fig = go.Figure(data=[go.Bar(x=area_count.values, y=area_count.index, orientation='h', 
                    marker=dict(color=area_count.values, colorscale='Oranges', line=dict(color='#d97706', width=1)), 
                    text=area_count.values, textposition='auto', hovertemplate='<b>%{y}</b><br>%{x}<extra></extra>')])
                fig.update_layout(height=250, margin=dict(l=100, r=10, t=10, b=10), plot_bgcolor='rgba(245, 247, 250, 0.5)', 
                    paper_bgcolor='rgba(255, 255, 255, 0)', xaxis=dict(showgrid=True, gridwidth=1, gridcolor='#e0e0e0', showticklabels=False), 
                    yaxis=dict(showgrid=False), font=dict(family='Segoe UI', size=9, color='#333'), hovermode='closest', showlegend=False)
                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        
        st.markdown("<div style='margin-top: 1rem;'></div>", unsafe_allow_html=True)
        if st.button("✕ Fechar", key="close_stats", help="Fechar estatísticas"):
            st.session_state.show_stats = False
            st.rerun()
    
    st.markdown("---")
    
    if len(df_filtrado) == 0:
        st.warning("Nenhum resultado encontrado. Tente ajustar os filtros.")
    else:
        # Linha com resultados e botões - USANDO st.columns COM gap="small"
        col_results, col_stats, col_export = st.columns([2.5, 0.7, 0.7], gap="small")
        
        with col_results:
            st.markdown(f"#### Resultados: {len(df_filtrado)} registros")
        
        with col_stats:
            if st.button("📊 Estatísticas", key="btn_stats_main", use_container_width=True):
                st.session_state.show_stats = True
        
        with col_export:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Glossário')
            buffer.seek(0)
            st.download_button(
                label="📥 Baixar dados em Excel",
                data=buffer,
                file_name=f"glossario_sinisa_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_glossario_{datetime.now().timestamp()}",
                use_container_width=True
            )
        
        st.markdown("")

        if st.session_state.get("show_stats", False):
            exibir_estatisticas_modal(df_filtrado)
            st.markdown("---")
        
        # Tabela de resultados
        col1, col2, col3, col4, col5, col6, col7, col8 = st.columns([0.6, 3.0, 0.6, 0.6, 1.5, 0.6, 0.8, 0.4], gap="small", vertical_alignment="center")
        
        with col1:
            st.markdown('<div class="table-header">Código</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="table-header">Nome da Informação</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="table-header">Módulo</div>', unsafe_allow_html=True)
        with col4:
            st.markdown('<div class="table-header">Formulário</div>', unsafe_allow_html=True)
        with col5:
            st.markdown('<div class="table-header">Subformulário</div>', unsafe_allow_html=True)
        with col6:
            st.markdown('<div class="table-header">Unidade</div>', unsafe_allow_html=True)
        with col7:
            st.markdown('<div class="table-header">Área Resp.</div>', unsafe_allow_html=True)
        with col8:
            st.markdown('<div class="table-header">Ação</div>', unsafe_allow_html=True)
        
        for idx, row in df_filtrado.iterrows():
            col1, col2, col3, col4, col5, col6, col7, col8 = st.columns([0.6, 3.0, 0.6, 0.6, 1.5, 0.6, 0.8, 0.4], gap="small", vertical_alignment="center")
            
            with col1:
                st.markdown(f'<div class="table-row">{row["Código da informação"]}</div>', unsafe_allow_html=True)
            
            with col2:
                st.markdown(f'<div class="table-row">{row["Nome da informação"]}</div>', unsafe_allow_html=True)
            
            with col3:
                st.markdown(f'<div class="table-row">{row["Módulo"]}</div>', unsafe_allow_html=True)
            
            with col4:
                st.markdown(f'<div class="table-row">{row["Formulário"]}</div>', unsafe_allow_html=True)
            
            with col5:
                subform = row['Subformulário'] if pd.notna(row['Subformulário']) else '-'
                st.markdown(f'<div class="table-row">{subform}</div>', unsafe_allow_html=True)
            
            with col6:
                unidade = row['Unidade'] if pd.notna(row['Unidade']) else '-'
                st.markdown(f'<div class="table-row">{unidade}</div>', unsafe_allow_html=True)
            
            with col7:
                area = row['Área Responsável'] if pd.notna(row['Área Responsável']) else '-'
                st.markdown(f'<div class="table-row">{area}</div>', unsafe_allow_html=True)
            
            with col8:                    
                if st.button("🔍", key=f"detail_{idx}", help="Ver detalhes", use_container_width=True):
                    exibir_detalhes_popup(row)
            
            st.markdown('<div class="table-divider"></div>', unsafe_allow_html=True)

# ============================================================================
# DICIONÁRIO DE PÁGINAS
# ============================================================================
PAGES = {
    "Sobre": pagina_sobre,
    "Glossário": {
        "Informações": pagina_glossario_informacoes_com_stats,
        "Indicadores": pagina_glossario_indicadores,
        "Avisos e Erros": pagina_glossario_avisos,

    },
    "Relatórios": {
        "Gerencial por Município": pagina_relatorios_municipio,
        "Consolidado": pagina_relatorios_consolidado,
        "Verificação": pagina_relatorios_verificacao,
        "Indicadores ANA": pagina_relatorios_indicadores_ana,
        "RAD (AGEMS)": pagina_relatorios_rad_agems,

    },
    "Contrato Programa": pagina_contrato_programa,
    "Investimento": pagina_investimento,
    "Infraestrutura": pagina_infraestrutura,
}

# ============================================================================
# NAVEGAÇÃO NO SIDEBAR
# ============================================================================

def render_sidebar_navigation():
    """Renderiza navegação no sidebar com melhor organização"""
    with st.sidebar:
        st.markdown("### 📌 Navegação")
        
        # INICIALIZAÇÃO SEGURA DO SESSION STATE (Evita o KeyError)
        if "current_page" not in st.session_state:
            st.session_state["current_page"] = "Sobre"
            
        # Página Sobre
        if st.button("📖 Sobre", use_container_width=True, key="btn_sobre"):
            st.session_state["current_page"] = "Sobre"
            st.rerun()
            
        # Glossário (com subtópicos)
        with st.expander("📚 Glossário", expanded=False):
            if st.button("📋 Informações", use_container_width=True, key="btn_info"):
                st.session_state["current_page"] = "Informações"
                st.rerun()
            if st.button("📈 Indicadores", use_container_width=True, key="btn_ind"):
                st.session_state["current_page"] = "Indicadores"
                st.rerun()
            if st.button("⚠️ Avisos e Erros", use_container_width=True, key="btn_avisos"):
                st.session_state["current_page"] = "Avisos e Erros"
                st.rerun()
                
        # Relatórios (com subtópicos)
        with st.expander("📊 Relatórios", expanded=False):
            if st.button("🏘️ Gerencial por Município", use_container_width=True, key="btn_mun"):
                st.session_state["current_page"] = "Gerencial por Município"
                st.rerun()
            if st.button("📊 Consolidado", use_container_width=True, key="btn_cons"):
                st.session_state["current_page"] = "Consolidado"
                st.rerun()

            if st.button("✅ Verificação", use_container_width=True, key="btn_verif"):
                st.session_state["current_page"] = "Verificação"
                st.rerun()

            if st.button("💧 Indicadores ANA", use_container_width=True, key="btn_ana"):
                st.session_state["current_page"] = "Indicadores ANA"
                st.rerun()
            if st.button("📈 RAD (AGEMS)", use_container_width=True, key="btn_rad"):
                st.session_state["current_page"] = "RAD (AGEMS)"
                st.rerun()
                
        # Infraestrutura
        if st.button("🏗️ Infraestrutura", use_container_width=True, key="btn_infra"):
            st.session_state["current_page"] = "Infraestrutura"
            st.rerun()

        # Contrato Programa
        if st.button("📄 Contrato Programa", use_container_width=True, key="btn_contrato"):
            st.session_state["current_page"] = "Contrato Programa"
            st.rerun()

        # Investimento
        if st.button("💰 Investimento", use_container_width=True, key="btn_investimento"):
            st.session_state["current_page"] = "Investimento"
            st.rerun()
                


def get_current_page_function():
    """Obtém a função da página atual"""
    current = st.session_state.get("current_page", "Sobre")
    
    # Procurar em páginas únicas
    if current in PAGES and not isinstance(PAGES[current], dict):
        return PAGES[current]
    
    # Procurar em subtópicos
    for section_name, section_content in PAGES.items():
        if isinstance(section_content, dict):
            if current in section_content:
                return section_content[current]
    
    # Padrão: retornar página "Sobre"
    return PAGES["Sobre"]

# ============================================================================
# EXECUÇÃO PRINCIPAL
# ============================================================================

def main():
    """Função principal"""
    # Renderizar sidebar
    render_sidebar_navigation()
    
    # Executar página atual
    current_func = get_current_page_function()
    current_func()

if __name__ == "__main__":
    main()