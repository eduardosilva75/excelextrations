import os
import pandas as pd
import numpy as np

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QProgressBar, QTableWidget,
    QTableWidgetItem, QHeaderView, QComboBox, QLineEdit,
    QCheckBox, QMainWindow, QWidget, QApplication
)

from PyQt5.QtPrintSupport import QPrinter

from PyQt5.QtGui import (
    QFont,
    QColor,
    QTextDocument,
    QTextCursor,
    QTextTableFormat,
    QTextTableCellFormat,
    QTextCharFormat,
    QTextBlockFormat,
    QTextLength,
    QPageSize,
    QPageLayout
)

from PyQt5.QtCore import Qt, QMarginsF
from PyQt5.QtGui import QTextFrameFormat

class ArtigosSemPSDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Artigos sem Presentation Stock")
        self.setGeometry(100, 100, 1600, 800)  # Largura maior para mais colunas
        self.df = None
        self.df_filtered = None
        self.df_com_ps = None  # DataFrame com artigos que têm Presentation Stock > 0
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        # Título
        title = QLabel("Artigos sem Presentation Stock (Presentation Stock = 0)")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("margin: 20px;")
        layout.addWidget(title)
        
        # Área de upload
        upload_layout = QHBoxLayout()
        self.btn_file = QPushButton("📁 Carregar Ficheiro Excel")
        self.btn_file.setFont(QFont("Arial", 12))
        self.btn_file.setMinimumHeight(40)
        self.btn_file.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        self.btn_file.clicked.connect(self.carregar_ficheiro)
        upload_layout.addWidget(self.btn_file)
        
        self.label_file = QLabel("Nenhum ficheiro carregado")
        self.label_file.setStyleSheet("color: #666; padding: 10px;")
        upload_layout.addWidget(self.label_file)
        upload_layout.addStretch()
        layout.addLayout(upload_layout)
        
        # Filtros
        filters_layout = QHBoxLayout()

        filters_layout.addWidget(QLabel("Filtrar por Secção:"))

        self.combo_seccao = QComboBox()
        self.combo_seccao.setMinimumWidth(150)
        self.combo_seccao.addItem("Todas as Secções")
        self.combo_seccao.currentTextChanged.connect(self.filtrar_por_seccao)
        filters_layout.addWidget(self.combo_seccao)

        filters_layout.addStretch()

        self.label_contador = QLabel("Total de artigos sem Presentation Stock: 0")
        self.label_contador.setStyleSheet("font-weight: bold;")
        filters_layout.addWidget(self.label_contador)

        # NOVO: Filtro por Status
        filters_layout.addWidget(QLabel("Status:"))

        self.combo_status = QComboBox()
        self.combo_status.setMinimumWidth(150)
        self.combo_status.addItem("Todos os Status")
        self.combo_status.currentTextChanged.connect(self.aplicar_filtros)
        filters_layout.addWidget(self.combo_status)

        filters_layout.addStretch()

        self.label_contador = QLabel("Total de artigos sem Presentation Stock: 0")
        self.label_contador.setStyleSheet("font-weight: bold;")
        filters_layout.addWidget(self.label_contador)        

        layout.addLayout(filters_layout)
        
        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Tabela
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget {
                gridline-color: #d0d0d0;
                background-color: white;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 5px;
                border: 1px solid #d0d0d0;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.table)
        
        # Botões de ação
        buttons_layout = QHBoxLayout()

        self.btn_exportar_excel = QPushButton("💾 Exportar para Excel")
        self.btn_exportar_excel.setFont(QFont("Arial", 12))
        self.btn_exportar_excel.setMinimumHeight(40)
        self.btn_exportar_excel.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.btn_exportar_excel.clicked.connect(self.exportar_excel)
        self.btn_exportar_excel.setEnabled(False)
        buttons_layout.addWidget(self.btn_exportar_excel)

        self.btn_exportar_pdf = QPushButton("📄 Exportar para PDF")
        self.btn_exportar_pdf.setFont(QFont("Arial", 12))
        self.btn_exportar_pdf.setMinimumHeight(40)
        self.btn_exportar_pdf.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.btn_exportar_pdf.clicked.connect(self.exportar_pdf)
        self.btn_exportar_pdf.setEnabled(False)
        buttons_layout.addWidget(self.btn_exportar_pdf)

        self.btn_limpar = QPushButton("🗑️ Limpar")
        self.btn_limpar.setFont(QFont("Arial", 12))
        self.btn_limpar.setMinimumHeight(40)
        self.btn_limpar.setStyleSheet("""
            QPushButton {
                background-color: #ff9800;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #e68900;
            }
        """)
        self.btn_limpar.clicked.connect(self.limpar_tudo)
        buttons_layout.addWidget(self.btn_limpar)

        buttons_layout.addStretch()

        self.btn_fechar = QPushButton("Fechar")
        self.btn_fechar.setFont(QFont("Arial", 12))
        self.btn_fechar.setMinimumHeight(40)
        self.btn_fechar.setStyleSheet("""
            QPushButton {
                background-color: #607D8B;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #546E7A;
            }
        """)
        self.btn_fechar.clicked.connect(self.close)
        buttons_layout.addWidget(self.btn_fechar)

        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)

    def aplicar_filtros(self):
        if self.df_filtered is None:
            return
        
        try:
            seccao_selecionada = self.combo_seccao.currentText()
            status_selecionado = self.combo_status.currentText()
            
            # Aplicar filtro por secção
            if seccao_selecionada == "Todas as Secções":
                df_temp = self.df_filtered.copy()
            else:
                df_temp = self.df_filtered[self.df_filtered['Secção'] == seccao_selecionada].copy()
            
            # Aplicar filtro por status (se a coluna existir)
            if 'Status' in df_temp.columns and status_selecionado != "Todos os Status":
                df_temp = df_temp[df_temp['Status'] == status_selecionado]
            
            self.atualizar_tabela(df_temp)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao aplicar filtros: {str(e)}")

    def calcular_sugestao_ps(self):
        """Calcula sugestão de Presentation Stock baseada em múltiplos fatores - Versão Conservadora"""
        try:
            # Primeiro, identificar artigos com Presentation Stock > 0
            if 'Presentation Stock' not in self.df.columns:
                QMessageBox.warning(self, "Aviso", "Coluna 'Presentation Stock' não encontrada no ficheiro.")
                self.df['Sugestão Presentation Stock'] = 0
                return
            
            # Separar artigos com Presentation Stock > 0
            self.df_com_ps = self.df[self.df['Presentation Stock'] > 0].copy()
            
            if self.df_com_ps.empty:
                QMessageBox.warning(self, "Aviso", "Não foram encontrados artigos com Presentation Stock > 0 para análise.")
                self.df['Sugestão Presentation Stock'] = 0
                return
            
            # Calcular estatísticas por secção E flow-type para artigos com Presentation Stock
            stats_por_grupo = self.df_com_ps.groupby(['Secção', 'Flow-type']).agg({
                'Presentation Stock': ['count', 'mean', 'median', 'min', 'max'],
                'Unit Sales': ['mean', 'median'],
                'Sales Value': ['mean', 'median'],
                'PVP Em Vigor': 'median',
                'Sup.Pack Size': 'median'
            }).round(2)
            
            # Para fallback, calcular só por secção
            stats_por_seccao = self.df_com_ps.groupby('Secção').agg({
                'Presentation Stock': ['count', 'mean', 'median', 'min', 'max'],
                'Unit Sales': ['mean', 'median'],
                'Sales Value': ['mean', 'median'],
                'PVP Em Vigor': 'median',
                'Sup.Pack Size': 'median'
            }).round(2)
            
            # Para cada artigo sem Presentation Stock, calcular sugestão
            artigos_sem_ps = self.df[self.df['Presentation Stock'] == 0].copy()
            artigos_sem_ps['Sugestão Presentation Stock'] = 0
            
            for idx, artigo in artigos_sem_ps.iterrows():
                seccao = artigo['Secção']
                flow_type = artigo.get('Flow-type', 'N/A')
                unit_sales = artigo['Unit Sales'] if pd.notna(artigo['Unit Sales']) else 0
                sales_value = artigo.get('Sales Value', 0) if pd.notna(artigo.get('Sales Value', 0)) else 0
                pvp = artigo.get('PVP Em Vigor', 0) if pd.notna(artigo.get('PVP Em Vigor', 0)) else 0
                pack_size = artigo.get('Sup.Pack Size', 1) if pd.notna(artigo.get('Sup.Pack Size', 1)) else 1
                
                # Tentar encontrar grupo mais específico primeiro (secção + flow-type)
                grupo_chave = (seccao, flow_type)
                
                if grupo_chave in stats_por_grupo.index:
                    # Usar estatísticas do grupo específico
                    ps_median = stats_por_grupo.loc[grupo_chave, ('Presentation Stock', 'median')]
                    ps_min = stats_por_grupo.loc[grupo_chave, ('Presentation Stock', 'min')]
                    unit_sales_median = stats_por_grupo.loc[grupo_chave, ('Unit Sales', 'median')]
                    grupo_count = stats_por_grupo.loc[grupo_chave, ('Presentation Stock', 'count')]
                    
                    # Base: usar o MÍNIMO do grupo como referência mais conservadora
                    sugestao_base = max(ps_min, ps_median * 0.5)  # Usar mínimo ou metade da mediana
                    
                elif seccao in stats_por_seccao.index:
                    # Fallback: usar só estatísticas da secção
                    ps_median = stats_por_seccao.loc[seccao, ('Presentation Stock', 'median')]
                    ps_min = stats_por_seccao.loc[seccao, ('Presentation Stock', 'min')]
                    unit_sales_median = stats_por_seccao.loc[seccao, ('Unit Sales', 'median')]
                    grupo_count = stats_por_seccao.loc[seccao, ('Presentation Stock', 'count')]
                    
                    sugestao_base = max(ps_min, ps_median * 0.5)  # Mais conservador
                else:
                    # Se não há dados da secção, usar sugestão mínima de 3 unidades
                    artigos_sem_ps.at[idx, 'Sugestão Presentation Stock'] = 3
                    continue
                
                # --- FACTOR 1: Vendas do artigo (28 DIAS - mais conservador) ---
                fator_vendas = 1.0
                
                if unit_sales_median > 0:
                    if unit_sales == 0:
                        # Artigo sem vendas em 28 dias - sugestão muito conservadora
                        fator_vendas = 0.1  # Apenas 10% da base
                    elif unit_sales <= 2:  # 2 ou menos vendas em 28 dias
                        fator_vendas = 0.3
                    elif unit_sales < (unit_sales_median * 0.1):
                        # Vendas muito baixas (<10% da mediana)
                        fator_vendas = 0.4
                    elif unit_sales < (unit_sales_median * 0.3):
                        # Vendas baixas (<30% da mediana)
                        fator_vendas = 0.6
                    elif unit_sales > (unit_sales_median * 3):
                        # Vendas altas (>300% da mediana)
                        fator_vendas = 1.3  # Reduzido de 1.5
                    elif unit_sales > (unit_sales_median * 2):
                        # Vendas acima da média (>200% da mediana)
                        fator_vendas = 1.1  # Reduzido de 1.3
                    else:
                        # Vendas normais
                        fator_vendas = 1.0
                
                # --- FACTOR 2: Valor das vendas (mais agressivo para preços altos) ---
                fator_valor = 1.0
                if pvp > 0:
                    # Reduzir drasticamente para artigos de alto valor
                    if pvp > 100:  # Muito alto valor
                        fator_valor = 0.3
                    elif pvp > 50:  # Alto valor
                        fator_valor = 0.5
                    elif pvp > 20:  # Valor médio-alto
                        fator_valor = 0.7
                    elif pvp > 10:  # Valor médio
                        fator_valor = 0.9
                
                # --- FACTOR 3: Tipo de fluxo ---
                fator_flow = 1.0
                if flow_type in ['PBSPC', 'PBsPC']:  # Produtos básicos/promocionais
                    fator_flow = 1.1  # Ligeiro aumento
                elif flow_type in ['NOV', 'NEW']:  # Novidades
                    fator_flow = 1.2  # Stock inicial moderado
                elif flow_type in ['SLOW', 'LENTO']:  # Vendas lentas
                    fator_flow = 0.5  # Redução significativa
                
                # --- FACTOR 4: Pack Size (mais inteligente) ---
                fator_pack = 1.0
                sugestao_antes_pack = sugestao_base * fator_vendas * fator_valor * fator_flow
                
                # Só ajustar ao pack se fizer sentido
                if pack_size > 1:
                    if pack_size <= 24:
                        # Packs pequenos: ajustar ao pack completo mas mínimo de 3 unidades
                        packs_sugeridos = max(1, round(sugestao_antes_pack / pack_size))
                        sugestao_ajustada = packs_sugeridos * pack_size
                        # Garantir mínimo de 3 unidades mesmo para packs
                        if packs_sugeridos == 1 and pack_size < 3:
                            sugestao_ajustada = 3
                    elif pack_size <= 36:
                        # Packs médios: mais flexibilidade
                        if sugestao_antes_pack >= pack_size * 1.5:
                            packs_sugeridos = max(1, round(sugestao_antes_pack / pack_size))
                            sugestao_ajustada = packs_sugeridos * pack_size
                        else:
                            # Não justifica pack completo, manter unidades com mínimo de 3
                            sugestao_ajustada = max(3, sugestao_antes_pack)
                    else:
                        # Packs grandes (>36): só ajustar se vendas justificarem
                        if sugestao_antes_pack >= pack_size * 2:
                            packs_sugeridos = max(1, round(sugestao_antes_pack / pack_size))
                            sugestao_ajustada = packs_sugeridos * pack_size
                        else:
                            # Manter em unidades para packs grandes com mínimo de 3
                            sugestao_ajustada = max(3, sugestao_antes_pack)
                else:
                    sugestao_ajustada = max(3, sugestao_antes_pack)  # Mínimo de 3 unidades
                
                # --- LIMITES FINAIS ---
                # Mínimo: 3 unidades (alterado de 1)
                sugestao_minima = 3
                
                # Máximo: mais conservador
                if seccao in stats_por_seccao.index:
                    ps_max_seccao = stats_por_seccao.loc[seccao, ('Presentation Stock', 'max')]
                    sugestao_maxima = min(ps_max_seccao, ps_median * 2)  # Máximo 2x a mediana
                else:
                    sugestao_maxima = ps_median * 2
                
                # Aplicar limites
                sugestao_final = max(sugestao_minima, min(sugestao_ajustada, sugestao_maxima))
                
                # Arredondar para inteiro
                sugestao_final = int(round(sugestao_final))
                
                artigos_sem_ps.at[idx, 'Sugestão Presentation Stock'] = sugestao_final
            
            # Atualizar o DataFrame principal com as sugestões
            self.df.loc[artigos_sem_ps.index, 'Sugestão Presentation Stock'] = artigos_sem_ps['Sugestão Presentation Stock']
            
            # Para artigos com Presentation Stock > 0, definir sugestão como 0
            self.df.loc[self.df['Presentation Stock'] > 0, 'Sugestão Presentation Stock'] = 0
            
            # Log de estatísticas
            sugestoes = artigos_sem_ps['Sugestão Presentation Stock']
            QMessageBox.information(
                self, 
                "Sugestões Calculadas", 
                f"Sugestões de Presentation Stock calculadas para {len(artigos_sem_ps)} artigos.\n\n"
                f"Estatísticas das sugestões:\n"
                f"• Média: {sugestoes.mean():.1f} unidades\n"
                f"• Mediana: {sugestoes.median():.1f} unidades\n"
                f"• Mínimo: {sugestoes.min()} unidades\n"
                f"• Máximo: {sugestoes.max()} unidades\n\n"
                f"Baseado em análise conservadora de {len(self.df_com_ps)} artigos com PS > 0."
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao calcular sugestões de Presentation Stock: {str(e)}")
            import traceback
            print(traceback.format_exc())
            self.df['Sugestão Presentation Stock'] = 0

    def carregar_ficheiro(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, 
                "Selecionar Ficheiro", 
                "", 
                "Ficheiros Suportados (*.xlsx *.xls *.csv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv)"
            )
            
            if file_path:
                self.progress_bar.setVisible(True)
                self.progress_bar.setValue(0)
                
                # Determinar o tipo de ficheiro pela extensão
                file_extension = file_path.lower().split('.')[-1]
                
                # Carregar o ficheiro conforme a extensão
                if file_extension in ['xlsx', 'xls']:
                    self.df = pd.read_excel(file_path)
                elif file_extension == 'csv':
                    self.df = self.carregar_csv(file_path)
                else:
                    QMessageBox.critical(self, "Erro", "Formato de ficheiro não suportado.")
                    self.progress_bar.setVisible(False)
                    return
                
                self.progress_bar.setValue(50)
                
                # Verificar se as colunas necessárias existem
                colunas_necessarias = ['Sku', 'Description', 'Unit Sales', 'Stock', 'Merc.Struct Code', 'Presentation Stock']
                colunas_faltantes = [col for col in colunas_necessarias if col not in self.df.columns]
                
                if colunas_faltantes:
                    QMessageBox.critical(
                        self, 
                        "Erro", 
                        f"Colunas faltantes no ficheiro: {', '.join(colunas_faltantes)}\n\nColunas encontradas: {', '.join(self.df.columns)}"
                    )
                    self.progress_bar.setVisible(False)
                    return
                
                # Extrair secção do Merc.Struct Code
                self.df['Secção'] = self.df['Merc.Struct Code'].astype(str).str[2:4]
                
                # Adicionar coluna de sugestão Presentation Stock
                self.df['Sugestão Presentation Stock'] = 0
                
                # Calcular sugestões de Presentation Stock
                self.calcular_sugestao_ps()
                
                # Filtrar apenas artigos com Presentation Stock = 0
                self.df_filtered = self.df[self.df['Presentation Stock'] == 0].copy()
                
                # Ordenar por secção e depois por Unit Sales (decrescente)
                self.df_filtered = self.df_filtered.sort_values(['Secção', 'Unit Sales'], ascending=[True, False])
                
                # Preencher combobox com secções únicas dos artigos sem Presentation Stock
                seccoes = sorted(self.df_filtered['Secção'].unique())
                self.combo_seccao.clear()
                self.combo_seccao.addItem("Todas as Secções")
                self.combo_seccao.addItems([str(sec) for sec in seccoes])
                
                # NOVO: Preencher combobox com status únicos
                if 'Status' in self.df_filtered.columns:
                    status_unicos = sorted(self.df_filtered['Status'].dropna().unique())
                    self.combo_status.clear()
                    self.combo_status.addItem("Todos os Status")
                    self.combo_status.addItems([str(status) for status in status_unicos])
                else:
                    self.combo_status.clear()
                    self.combo_status.addItem("Todos os Status")
                    self.combo_status.setEnabled(False)
                
                self.progress_bar.setValue(100)
                
                # Atualizar interface
                self.label_file.setText(os.path.basename(file_path))
                self.btn_exportar_excel.setEnabled(True)
                self.btn_exportar_pdf.setEnabled(True)
                self.aplicar_filtros()  # Mudado de filtrar_por_seccao para aplicar_filtros
                
                QMessageBox.information(
                    self, 
                    "Sucesso", 
                    f"Ficheiro carregado com sucesso!\n"
                    f"{len(self.df_filtered)} artigos sem Presentation Stock encontrados.\n"
                    f"{len(self.df_com_ps) if self.df_com_ps is not None else 0} artigos com Presentation Stock analisados."
                )
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar ficheiro: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def carregar_csv(self, file_path):
        """Carrega ficheiro CSV com deteção automática de delimitador e encoding"""
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    first_lines = [f.readline() for _ in range(5)]
                
                delimiters = [',', ';', '\t', '|']
                delimiter_scores = {}
                
                for delimiter in delimiters:
                    score = 0
                    for line in first_lines:
                        if line:
                            score += line.count(delimiter)
                    delimiter_scores[delimiter] = score
                
                best_delimiter = max(delimiter_scores, key=delimiter_scores.get)
                
                if delimiter_scores[best_delimiter] == 0:
                    best_delimiter = ','
                
                df = pd.read_csv(file_path, delimiter=best_delimiter, encoding=encoding)
                df.columns = df.columns.str.strip()
                
                print(f"CSV carregado com encoding: {encoding}, delimitador: '{best_delimiter}'")
                return df
                
            except UnicodeDecodeError:
                continue
            except Exception as e:
                print(f"Tentativa com encoding {encoding} falhou: {e}")
                continue
        
        return self.carregar_csv_manual(file_path)

    def carregar_csv_manual(self, file_path):
        """Fallback para carregamento manual de CSV"""
        QMessageBox.warning(
            self, 
            "Deteção Automática Falhou", 
            "Não foi possível detetar automaticamente o formato do CSV.\n"
            "Por favor, selecione manualmente o delimitador e encoding."
        )
        
        # Implementar diálogo para seleção manual se necessário
        # Por enquanto, tentar com valores padrão
        try:
            df = pd.read_csv(file_path, delimiter=',', encoding='latin-1')
            df.columns = df.columns.str.strip()
            return df
        except:
            try:
                df = pd.read_csv(file_path, delimiter=';', encoding='latin-1')
                df.columns = df.columns.str.strip()
                return df
            except Exception as e:
                raise Exception(f"Não foi possível ler o ficheiro CSV: {str(e)}")

    def filtrar_por_seccao(self):
        if self.df_filtered is None:
            return
        
        try:
            seccao_selecionada = self.combo_seccao.currentText()
            
            if seccao_selecionada == "Todas as Secções":
                df_temp = self.df_filtered.copy()
            else:
                df_temp = self.df_filtered[self.df_filtered['Secção'] == seccao_selecionada].copy()
            
            self.atualizar_tabela(df_temp)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao filtrar: {str(e)}")

    def atualizar_tabela(self, df):
        try:
            # Configurar tabela com 9 colunas
            self.table.setRowCount(len(df))
            self.table.setColumnCount(9)
            self.table.setHorizontalHeaderLabels([
                'Sku', 'Description', 'Sup.Pack Size', 'PVP Em Vigor', 'Stock', 
                'Unit Sales', 'Flow-type', 'Secção', 'Sugestão Presentation Stock'
            ])
            
            # Preencher tabela
            for row_idx, (_, row) in enumerate(df.iterrows()):
                # Sku
                item_sku = QTableWidgetItem(str(row['Sku']))
                item_sku.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(row_idx, 0, item_sku)
                
                # Description
                item_desc = QTableWidgetItem(str(row['Description']))
                item_desc.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(row_idx, 1, item_desc)
                
                # Sup.Pack Size
                pack_size = row.get('Sup.Pack Size', 0) if pd.notna(row.get('Sup.Pack Size')) else 0
                item_pack = QTableWidgetItem(f"{pack_size:,.0f}")
                item_pack.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 2, item_pack)
                
                # PVP Em Vigor
                pvp = row.get('PVP Em Vigor', 0) if pd.notna(row.get('PVP Em Vigor')) else 0
                item_pvp = QTableWidgetItem(f"€ {pvp:,.2f}")
                item_pvp.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 3, item_pvp)
                
                # Stock
                stock_value = row['Stock'] if pd.notna(row['Stock']) else 0
                item_stock = QTableWidgetItem(f"{stock_value:,.0f}")
                item_stock.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                # Colorir stock baixo
                if stock_value == 0:
                    item_stock.setBackground(QColor(255, 200, 200))  # Vermelho claro para stock 0
                elif stock_value < (row.get('Sugestão Presentation Stock', 0) or 0):
                    item_stock.setBackground(QColor(255, 255, 200))  # Amarelo para stock abaixo da sugestão
                
                self.table.setItem(row_idx, 4, item_stock)
                
                # Unit Sales
                unit_sales_value = row['Unit Sales'] if pd.notna(row['Unit Sales']) else 0
                item_unit_sales = QTableWidgetItem(f"{unit_sales_value:,.0f}")
                item_unit_sales.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                # Colorir vendas altas
                if unit_sales_value > 100:
                    item_unit_sales.setBackground(QColor(200, 255, 200))  # Verde claro para vendas altas
                elif unit_sales_value > 50:
                    item_unit_sales.setBackground(QColor(255, 255, 200))  # Amarelo para vendas médias
                
                self.table.setItem(row_idx, 5, item_unit_sales)
                
                # Flow-type
                flow_type = str(row.get('Flow-type', 'N/A')) if 'Flow-type' in row and pd.notna(row.get('Flow-type')) else "N/A"
                item_flow = QTableWidgetItem(flow_type)
                item_flow.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 6, item_flow)
                
                # Secção
                item_seccao = QTableWidgetItem(str(row['Secção']))
                item_seccao.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 7, item_seccao)
                
                # Sugestão Presentation Stock
                sugestao = row.get('Sugestão Presentation Stock', 0)
                item_sugestao = QTableWidgetItem(f"{sugestao:,.0f}")
                item_sugestao.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                # Colorir sugestões altas
                if sugestao > 10:
                    item_sugestao.setBackground(QColor(200, 230, 255))  # Azul claro para sugestões altas
                
                self.table.setItem(row_idx, 8, item_sugestao)
            
            # Ajustar tamanho das colunas
            header = self.table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Sku
            header.setSectionResizeMode(1, QHeaderView.Stretch)          # Description
            header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Sup.Pack Size
            header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # PVP Em Vigor
            header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Stock
            header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Unit Sales
            header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # Flow-type
            header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # Secção
            header.setSectionResizeMode(8, QHeaderView.ResizeToContents)  # Sugestão Presentation Stock
            
            # Atualizar contador
            self.label_contador.setText(f"Total de artigos sem Presentation Stock: {len(df):,}")
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar tabela: {str(e)}")

    def exportar_pdf(self):
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Aviso", "Não existem dados para exportar.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Exportar para PDF", "Artigos_Sem_PS.pdf", "PDF (*.pdf)"
        )
        if not file_path:
            return

        try:
            seccao_selecionada = self.combo_seccao.currentText()
            status_selecionado = self.combo_status.currentText()

            # Aplicar os mesmos filtros da visualização atual
            if seccao_selecionada == "Todas as Secções":
                df_export = self.df_filtered.copy()
            else:
                df_export = self.df_filtered[self.df_filtered['Secção'] == seccao_selecionada].copy()

            if 'Status' in df_export.columns and status_selecionado != "Todos os Status":
                df_export = df_export[df_export['Status'] == status_selecionado]

            # Configuração PDF
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(file_path)

            layout = QPageLayout(
                QPageSize(QPageSize.A4),
                QPageLayout.Landscape,
                QMarginsF(10, 10, 10, 10),
                QPageLayout.Millimeter
            )
            printer.setPageLayout(layout)

            doc = QTextDocument()
            cursor = QTextCursor(doc)
            doc.setDefaultFont(QFont("Arial", 8))

            # Título e info
            title_fmt = QTextCharFormat()
            title_fmt.setFont(QFont("Arial", 16, QFont.Bold))
            block_fmt = QTextBlockFormat()
            block_fmt.setAlignment(Qt.AlignCenter)
            cursor.insertBlock(block_fmt)
            cursor.setCharFormat(title_fmt)
            cursor.insertText("ARTIGOS SEM PRESENTATION STOCK\n\n")

            info = f"Secção: {seccao_selecionada} | " \
                f"Total artigos: {len(df_export):,} | " \
                f"Gerado em: {pd.Timestamp.now():%d/%m/%Y %H:%M}\n\n"
            cursor.insertText(info)

            # Cabeçalhos
            headers = [
                'Sku', 'Description', 'Pack', 'PVP', 'Stock', 
                'Unit Sales', 'Flow', 'Sec', 'Sug. Presentation Stock'
            ]

            # Larguras ajustadas
            larguras_percentagem = [
                10,   # Sku
                30,   # Description
                6,    # Pack
                8,    # PVP
                8,    # Stock
                9,    # Unit Sales
                8,    # Flow
                6,    # Sec
                8     # Sug. Presentation Stock
            ]  # soma = 93% (deixa margem)

            # Formato da tabela
            table_fmt = QTextTableFormat()
            table_fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100))
            table_fmt.setCellPadding(4)
            table_fmt.setCellSpacing(0)
            table_fmt.setBorder(0.5)
            table_fmt.setBorderStyle(QTextFrameFormat.BorderStyle_Solid)

            constraints = [QTextLength(QTextLength.PercentageLength, w) for w in larguras_percentagem]
            table_fmt.setColumnWidthConstraints(constraints)

            table = cursor.insertTable(len(df_export) + 1, len(headers), table_fmt)

            # Cabeçalho
            header_cell_fmt = QTextTableCellFormat()
            header_cell_fmt.setBackground(QColor("#d0d0d0"))

            header_char_fmt = QTextCharFormat()
            header_char_fmt.setFontWeight(QFont.Bold)
            header_char_fmt.setFontPointSize(9)

            for col, texto in enumerate(headers):
                cell = table.cellAt(0, col)
                cell.setFormat(header_cell_fmt)
                cur = cell.firstCursorPosition()
                cur.insertText(texto, header_char_fmt)

            # Dados
            normal_fmt = QTextCharFormat()
            normal_fmt.setFontPointSize(8)

            for row_idx, (_, row) in enumerate(df_export.iterrows(), start=1):
                for col_idx, col_name in enumerate(headers):
                    cell = table.cellAt(row_idx, col_idx)
                    cur = cell.firstCursorPosition()

                    # Mapear cabeçalhos curtos para colunas reais
                    col_mapping = {
                        'Sku': 'Sku',
                        'Description': 'Description', 
                        'Pack': 'Sup.Pack Size',
                        'PVP': 'PVP Em Vigor',
                        'Stock': 'Stock',
                        'Unit Sales': 'Unit Sales',
                        'Flow': 'Flow-type',
                        'Sec': 'Secção',
                        'Sug. Presentation Stock': 'Sugestão Presentation Stock'
                    }
                    
                    real_col = col_mapping[col_name]
                    value = row.get(real_col, '')

                    if pd.isna(value):
                        text = "N/A"
                    else:
                        if real_col == "Description":
                            desc = str(value)
                            text = desc if len(desc) <= 40 else desc[:37] + "..."
                        elif real_col in ["Unit Sales", "Stock", "Sup.Pack Size", "Sugestão Presentation Stock"]:
                            text = f"{int(value):,}" if value else "0"
                        elif real_col == "PVP Em Vigor":
                            text = f"€{float(value):,.2f}" if value else "€0"
                        elif real_col == "Secção":
                            text = str(value)
                        else:
                            text = str(value)

                    cur.insertText(text, normal_fmt)

            # Rodapé
            cursor.movePosition(QTextCursor.End)
            cursor.insertBlock()
            footer = QTextCharFormat()
            footer.setFontPointSize(7)
            footer.setFontItalic(True)
            footer.setForeground(QColor("gray"))
            cursor.setCharFormat(footer)
            cursor.insertText(f"Documento gerado automaticamente • {len(df_export):,} artigos sem Presentation Stock")

            # Exportar
            doc.print_(printer)

            QMessageBox.information(
                self, "Sucesso",
                f"PDF exportado com sucesso!\n\n"
                f"→ {len(df_export):,} artigos exportados\n"
                f"→ Guardado em: {os.path.basename(file_path)}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar PDF:\n{str(e)}")

    def exportar_excel(self):
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados para exportar.")
            return
        
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Exportar para Excel",
                "artigos_sem_ps.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if file_path:
                self.progress_bar.setVisible(True)
                self.progress_bar.setValue(50)
                
                # Obter dados filtrados atuais
                seccao_selecionada = self.combo_seccao.currentText()
                status_selecionado = self.combo_status.currentText()

                # Aplicar os mesmos filtros da visualização atual
                if seccao_selecionada == "Todas as Secções":
                    df_export = self.df_filtered.copy()
                else:
                    df_export = self.df_filtered[self.df_filtered['Secção'] == seccao_selecionada].copy()

                if 'Status' in df_export.columns and status_selecionado != "Todos os Status":
                    df_export = df_export[df_export['Status'] == status_selecionado]
                
                # Colunas para exportação
                colunas_export = [
                    'Sku', 'Description', 'Sup.Pack Size', 'PVP Em Vigor', 'Stock', 
                    'Unit Sales', 'Flow-type', 'Secção', 'Sugestão Presentation Stock'
                ]
                
                # Filtrar apenas colunas que existem
                colunas_disponiveis = [col for col in colunas_export if col in df_export.columns]
                df_export = df_export[colunas_disponiveis].copy()
                
                # Exportar para Excel
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Artigos Sem Presentation Stock')
                    
                    # Acessar a worksheet para ajustar as colunas
                    worksheet = writer.sheets['Artigos Sem Presentation Stock']
                    
                    # Ajustar largura das colunas
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        
                        for cell in column:
                            try:
                                if cell.value:
                                    cell_length = len(str(cell.value))
                                    max_length = max(max_length, cell_length)
                            except:
                                pass
                        
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                    
                self.progress_bar.setValue(100)
                
                QMessageBox.information(
                    self, 
                    "Sucesso", 
                    f"Dados exportados com sucesso!\n{len(df_export)} artigos exportados."
                )
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def limpar_tudo(self):
        self.df = None
        self.df_filtered = None
        self.df_com_ps = None
        self.table.setRowCount(0)
        self.label_file.setText("Nenhum ficheiro carregado")
        self.combo_seccao.clear()
        self.combo_seccao.addItem("Todas as Secções")
        self.combo_status.clear()  # NOVO
        self.combo_status.addItem("Todos os Status")  # NOVO
        self.combo_status.setEnabled(True)  # NOVO
        self.label_contador.setText("Total de artigos sem Presentation Stock: 0")
        self.btn_exportar_excel.setEnabled(False)
        self.btn_exportar_pdf.setEnabled(False)

def mostrar_artigos_sem_ps():
    dialog = ArtigosSemPSDialog()
    dialog.exec_()