import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os

# Configuração da página
st.set_page_config(
    page_title="Alfa Sharks Flow Dashboard",
    page_icon="🦈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        font-weight: bold;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 1.5rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
        margin-bottom: 1rem;
    }
    .positive-value {
        color: #28a745;
        font-weight: bold;
    }
    .negative-value {
        color: #dc3545;
        font-weight: bold;
    }
    .section-header {
        font-size: 1.5rem;
        color: #2c3e50;
        margin-top: 2rem;
        margin-bottom: 1rem;
        border-bottom: 2px solid #1f77b4;
        padding-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

class FlowDashboard:
    def __init__(self, file_path):
        self.file_path = file_path
        self.saldo_data = None
        self.fluxo_data = None
        self.load_data()
        
    def load_data(self):
        """Carrega os dados das planilhas"""
        try:
            # Carregar dados da aba 'saldo'
            self.saldo_data = pd.read_excel(self.file_path, sheet_name='saldo')
            
            # Carregar dados da aba 'fluxortd'
            self.fluxo_data = pd.read_excel(self.file_path, sheet_name='fluxortd')
            
            st.success("Dados carregados com sucesso!")
            
        except Exception as e:
            st.error(f"Erro ao carregar dados: {str(e)}")
    
    def get_saldo_final(self):
        """Obtém os saldos finais diretamente da coluna T - CORRIGIDO"""
        saldo_df = self.saldo_data.copy()
        
        # Procurar pelos valores na coluna T (índice 19) onde a coluna S tem os nomes
        saldos = {}
        
        for idx, row in saldo_df.iterrows():
            if pd.notna(row.iloc[18]):  # Coluna S (nomes: Estrangeiros, Bancos, etc.)
                categoria = str(row.iloc[18]).strip()
                valor = row.iloc[19]  # Coluna T (valores)
                
                if pd.notna(valor) and categoria in ['Estrangeiros', 'Bancos', 'Pessoas Físicas', 'Saldo']:
                    # Converter para número, removendo "R$" e pontos
                    if isinstance(valor, str):
                        valor_limpo = valor.replace('R$', '').replace('.', '').replace(',', '.').strip()
                        try:
                            valor_num = float(valor_limpo)
                        except:
                            valor_num = 0
                    else:
                        valor_num = float(valor)
                    
                    saldos[categoria] = valor_num
        
        return saldos
    
    def process_fluxo_data(self):
        """Processa os dados de fluxo"""
        fluxo_df = self.fluxo_data.copy()
        
        # Verificar se as colunas existem
        if 'Compradora' in fluxo_df.columns and 'Quantidade' in fluxo_df.columns:
            # Filtrar apenas linhas com dados válidos
            valid_fluxo = fluxo_df[(fluxo_df['Compradora'].notna()) & 
                                 (fluxo_df['Compradora'] != '-') &
                                 (fluxo_df['Quantidade'] > 0)]
            return valid_fluxo
        else:
            st.warning("Colunas 'Compradora' ou 'Quantidade' não encontradas na aba fluxortd")
            return pd.DataFrame()
    
    def create_dashboard(self):
        """Cria o dashboard completo - SIMPLIFICADO"""
        # Header
        st.markdown('<h1 class="main-header">🦈 Alfa Sharks Flow Dashboard</h1>', unsafe_allow_html=True)
        
        # Obter saldos finais
        saldos = self.get_saldo_final()
        
        # Layout principal
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown("### 🌍 Estrangeiros")
            if 'Estrangeiros' in saldos:
                total = saldos['Estrangeiros']
                color_class = "positive-value" if total >= 0 else "negative-value"
                st.markdown(f'<div class="metric-card">Saldo: <span class="{color_class}">R$ {total:,.0f}</span></div>', unsafe_allow_html=True)
                st.write(f"**{'Compradores' if total >= 0 else 'Vendedores'}**")
            else:
                st.warning("Dados de Estrangeiros não encontrados")
        
        with col2:
            st.markdown("### 🏦 Bancos")
            if 'Bancos' in saldos:
                total = saldos['Bancos']
                color_class = "positive-value" if total >= 0 else "negative-value"
                st.markdown(f'<div class="metric-card">Saldo: <span class="{color_class}">R$ {total:,.0f}</span></div>', unsafe_allow_html=True)
                st.write(f"**{'Compradores' if total >= 0 else 'Vendedores'}**")
            else:
                st.warning("Dados de Bancos não encontrados")
        
        with col3:
            st.markdown("### 👥 Pessoas Físicas")
            if 'Pessoas Físicas' in saldos:
                total = saldos['Pessoas Físicas']
                color_class = "positive-value" if total >= 0 else "negative-value"
                st.markdown(f'<div class="metric-card">Saldo: <span class="{color_class}">R$ {total:,.0f}</span></div>', unsafe_allow_html=True)
                st.write(f"**{'Compradores' if total >= 0 else 'Vendedores'}**")
            else:
                st.warning("Dados de Pessoas Físicas não encontrados")
        
        with col4:
            st.markdown("### ⚖️ Saldo Total")
            if 'Saldo' in saldos:
                total = saldos['Saldo']
                color_class = "positive-value" if total >= 0 else "negative-value"
                st.markdown(f'<div class="metric-card">Total: <span class="{color_class}">R$ {total:,.0f}</span></div>', unsafe_allow_html=True)
                st.write(f"**{'Comprador Líquido' if total >= 0 else 'Vendedor Líquido'}**")
            else:
                st.warning("Saldo total não encontrado")
        
        # Gráfico de barras dos saldos
        st.markdown('<div class="section-header">📈 Análise de Saldos por Categoria</div>', unsafe_allow_html=True)
        
        if saldos:
            # Preparar dados para o gráfico
            categorias = []
            valores = []
            cores = []
            
            for cat, val in saldos.items():
                if cat != 'Saldo':  # Excluir o saldo total do gráfico de categorias
                    categorias.append(cat)
                    valores.append(val)
                    cores.append('#28a745' if val >= 0 else '#dc3545')
            
            # Gráfico de barras
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=categorias,
                y=valores,
                marker_color=cores,
                text=[f'R$ {v:,.0f}' for v in valores],
                textposition='auto',
            ))
            
            fig.update_layout(
                title='Saldos por Categoria',
                xaxis_title='Categoria',
                yaxis_title='Valor (R$)',
                showlegend=False,
                height=400
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Gráfico de pizza da distribuição
            st.markdown('<div class="section-header">🥧 Distribuição dos Fluxos</div>', unsafe_allow_html=True)
            
            fig_pie = px.pie(
                values=[abs(v) for v in valores],
                names=categorias,
                title='Distribuição do Volume por Categoria',
                color=categorias,
                color_discrete_sequence=['#1f77b4', '#ff7f0e', '#2ca02c']
            )
            
            st.plotly_chart(fig_pie, use_container_width=True)
        
        # Força relativa
        st.markdown('<div class="section-header">💪 Força Relativa</div>', unsafe_allow_html=True)
        
        if len(saldos) >= 3:
            col1, col2, col3 = st.columns(3)
            
            with col1:
                estr = abs(saldos.get('Estrangeiros', 0))
                total_abs = sum(abs(v) for k, v in saldos.items() if k != 'Saldo')
                if total_abs > 0:
                    perc_estr = (estr / total_abs) * 100
                    st.metric("Força Estrangeiros", f"{perc_estr:.1f}%")
            
            with col2:
                bancos = abs(saldos.get('Bancos', 0))
                if total_abs > 0:
                    perc_bancos = (bancos / total_abs) * 100
                    st.metric("Força Bancos", f"{perc_bancos:.1f}%")
            
            with col3:
                pf = abs(saldos.get('Pessoas Físicas', 0))
                if total_abs > 0:
                    perc_pf = (pf / total_abs) * 100
                    st.metric("Força Pessoas Físicas", f"{perc_pf:.1f}%")
        
        # Informações do arquivo
        st.sidebar.markdown("### 📁 Informações do Arquivo")
        st.sidebar.write(f"Local: {self.file_path}")
        st.sidebar.write(f"Última atualização: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Mostrar saldos brutos na sidebar
        st.sidebar.markdown("### 💰 Saldos Brutos")
        for categoria, valor in saldos.items():
            st.sidebar.write(f"**{categoria}:** R$ {valor:,.0f}")

def main():
    # Configuração do caminho do arquivo
    file_path = "planilhafluxo_PROCESSADO_COMPLETO.xlsx"
    
    # Verificar se o arquivo existe
    if not os.path.exists(file_path):
        st.error(f"Arquivo não encontrado: {file_path}")
        st.info("Por favor, verifique o caminho do arquivo e tente novamente.")
        return
    
    # Inicializar e exibir o dashboard
    dashboard = FlowDashboard(file_path)
    dashboard.create_dashboard()

if __name__ == "__main__":
    main()