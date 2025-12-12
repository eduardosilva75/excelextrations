import os
import pandas as pd

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QProgressBar, QTableWidget,
    QTableWidgetItem, QHeaderView, QComboBox, QLineEdit,
    QCheckBox, QMainWindow, QWidget, QApplication  # (adiciona os que já usas na tua app)
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
    QTextBlockFormat,      # ← necessário para alinhar título
    QTextLength,           # ← necessário para largura das colunas em %
    QPageSize,             # ← necessário para A4
    QPageLayout            # ← necessário para Landscape + margens
)

from PyQt5.QtCore import Qt, QMarginsF
from PyQt5.QtGui import QTextFrameFormat

class HitParadeDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Hit Parade por Secção")
        self.setGeometry(100, 100, 1400, 800)  # Aumentar largura para mais colunas
        self.df = None
        self.df_filtered = None
        self.ordenacao_atual = 'Unit Sales'  # Ordenação padrão
        self.ordem_decrescente = True  # Ordem padrão decrescente
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        # Título
        title = QLabel("Hit Parade por Secção")
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
        
        # Filtros e ordenação
        filters_layout = QHBoxLayout()

        filters_layout.addWidget(QLabel("Filtrar por Secção:"))

        self.combo_seccao = QComboBox()
        self.combo_seccao.setMinimumWidth(150)
        self.combo_seccao.addItem("Todas as Secções")
        self.combo_seccao.currentTextChanged.connect(self.filtrar_por_seccao)
        filters_layout.addWidget(self.combo_seccao)

        self.check_mostrar_todos = QCheckBox("Mostrar todos os artigos")
        self.check_mostrar_todos.stateChanged.connect(self.filtrar_por_seccao)
        filters_layout.addWidget(self.check_mostrar_todos)

        filters_layout.addStretch()

        # Controles de ordenação
        filters_layout.addWidget(QLabel("Ordenar por:"))
        self.combo_ordenacao = QComboBox()
        self.combo_ordenacao.addItems(["Unit Sales", "Sales Value", "%Vendas"])
        self.combo_ordenacao.currentTextChanged.connect(self.alterar_ordenacao)
        filters_layout.addWidget(self.combo_ordenacao)

        self.btn_ordem = QPushButton("🔽")
        self.btn_ordem.setToolTip("Alternar entre ordem crescente/decrescente")
        self.btn_ordem.clicked.connect(self.alternar_ordem)
        self.btn_ordem.setFixedSize(30, 30)
        filters_layout.addWidget(self.btn_ordem)

        filters_layout.addStretch()

        self.label_contador = QLabel("Total de artigos: 0")
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

        # Novo botão para PDF
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

    def calcular_percentual_vendas(self):
        """Calcula a percentagem de vendas (Stock Total em relação a Unit Sales) e a previsão de caixas"""
        try:
            import numpy as np
            
            # Inicializar a coluna %Vendas como float
            self.df['%Vendas'] = 0.0
            
            # Calcular Stock Total como soma de todas as colunas de stock
            self.df['Stock Total'] = self.df['Stock'].fillna(0).astype(float)
            
            # Adicionar outras colunas de stock se existirem
            stock_columns = ['Stock In Transit', 'Stock Expected', 'Stock On Order']
            for col in stock_columns:
                if col in self.df.columns:
                    self.df['Stock Total'] += self.df[col].fillna(0).astype(float)
            
            # Converter Unit Sales para float
            self.df['Unit Sales'] = pd.to_numeric(self.df['Unit Sales'], errors='coerce').fillna(0).astype(float)
            
            # %Vendas = (Stock Total / Unit Sales) * 100
            mask = (self.df['Unit Sales'] > 0) & (self.df['Stock Total'].notna())
            
            # Usar .loc com dtype explícito
            pct_vendas = (self.df.loc[mask, 'Stock Total'] / self.df.loc[mask, 'Unit Sales']) * 100
            self.df.loc[mask, '%Vendas'] = pct_vendas.astype(float)
            
            self.df.loc[self.df['Unit Sales'] == 0, '%Vendas'] = 99999.0
            self.df.loc[(self.df['Stock Total'] == 0) & (self.df['Unit Sales'] > 0), '%Vendas'] = 0.0
            self.df['%Vendas'] = self.df['%Vendas'].round(2)
            
            print(f"\n=== VERIFICAÇÃO %Vendas ===")
            print(f"%Vendas tipo: {self.df['%Vendas'].dtype}")
            print(f"%Vendas calculado: {self.df['%Vendas'].min():.1f}% a {self.df['%Vendas'].max():.1f}%")
            
            # CALCULAR PREVISÃO DE CAIXAS (Pr.Cx)
            colunas_necessarias = ['Unit Sales', 'Stock Total']
            if all(col in self.df.columns for col in colunas_necessarias):
                
                if 'Sup.Pack Size' not in self.df.columns:
                    self.df['Sup.Pack Size'] = 1.0
                
                self.df['Sup.Pack Size'] = pd.to_numeric(self.df['Sup.Pack Size'], errors='coerce').fillna(1.0).astype(float)
                self.df['Sup.Pack Size'] = self.df['Sup.Pack Size'].replace(0, 1.0)
                
                print(f"\n=== TESTE DE VALORES PARA Pr.Cx ===")
                print(f"Unit Sales: {self.df['Unit Sales'].min():.1f} a {self.df['Unit Sales'].max():.1f}")
                print(f"Stock Total: {self.df['Stock Total'].min():.1f} a {self.df['Stock Total'].max():.1f}")
                print(f"Sup.Pack Size: {self.df['Sup.Pack Size'].min():.1f} a {self.df['Sup.Pack Size'].max():.1f}")
                
                # Cálculo passo a passo
                print(f"\n=== CÁLCULO Pr.Cx PASSO A PASSO ===")
                
                # 1. Meta mínima = 25% das Unit Sales
                meta_minima = 0.25 * self.df['Unit Sales']
                print(f"1. Meta (25% Unit Sales): {meta_minima.min():.1f} a {meta_minima.max():.1f}")
                
                # 2. Necessidade = Meta - Stock Total
                necessidade = meta_minima - self.df['Stock Total']
                print(f"2. Necessidade (Meta - Stock): {necessidade.min():.1f} a {necessidade.max():.1f}")
                
                # 3. Garantir valores não negativos
                necessidade_total = necessidade.clip(lower=0)
                print(f"3. Necessidade (após clip): {necessidade_total.min():.1f} a {necessidade_total.max():.1f}")
                
                # 4. Calcular caixas em float
                # Evitar divisão por zero
                sup_pack_nonzero = self.df['Sup.Pack Size'].copy()
                sup_pack_nonzero[sup_pack_nonzero == 0] = 1.0
                caixas_float = necessidade_total / sup_pack_nonzero
                print(f"4. Caixas (float): {caixas_float.min():.3f} a {caixas_float.max():.3f}")
                
                # 5. Arredondar para CIMA e converter para inteiro
                self.df['Pr.Cx'] = np.ceil(caixas_float).astype(int)
                print(f"5. Pr.Cx (ceil->int): {self.df['Pr.Cx'].min()} a {self.df['Pr.Cx'].max()}")
                
                # 6. Para Unit Sales = 0, Pr.Cx = 0
                self.df.loc[self.df['Unit Sales'] == 0, 'Pr.Cx'] = 0
                print(f"6. Artigos com Unit Sales=0: {(self.df['Unit Sales'] == 0).sum()}")
                
                # EXEMPLOS DETALHADOS
                print(f"\n=== EXEMPLOS DETALHADOS (Primeiros 10) ===")
                exemplos_indices = list(range(min(10, len(self.df))))
                for idx in exemplos_indices:
                    unit = self.df.iloc[idx]['Unit Sales']
                    stock = self.df.iloc[idx]['Stock Total']
                    sup = self.df.iloc[idx]['Sup.Pack Size']
                    meta = 0.5 * unit
                    necessidade = max(0, meta - stock)
                    caixas = np.ceil(necessidade / sup).astype(int) if sup > 0 else 0
                    pct_vendas = self.df.iloc[idx]['%Vendas']
                    print(f"[{idx}] U={unit}, S={stock}, M={meta:.1f}, N={necessidade:.1f}, Sup={sup}, Pr.Cx={caixas}, %V={pct_vendas}%")
                
                print(f"\n=== EXEMPLOS ONDE DEVERIA TER Pr.Cx > 0 ===")
                # Procurar artigos onde necessidade > 0 mas Sup.Pack Size é grande
                mask_necessidade = necessidade_total > 0
                exemplos_necessidade = self.df[mask_necessidade].head(5)
                
                if not exemplos_necessidade.empty:
                    for idx, row in exemplos_necessidade.iterrows():
                        unit = row['Unit Sales']
                        stock = row['Stock Total']
                        sup = row['Sup.Pack Size']
                        meta = 0.5 * unit
                        necessidade = max(0, meta - stock)
                        caixas = np.ceil(necessidade / sup).astype(int) if sup > 0 else 0
                        print(f"[Idx {idx}] U={unit}, S={stock}, M={meta:.1f}, N={necessidade:.1f}, Sup={sup}, Pr.Cx={caixas}")
                else:
                    print("Nenhum artigo com necessidade > 0 encontrado")
                
                print(f"\n=== RESULTADO FINAL Pr.Cx ===")
                print(f"Total artigos: {len(self.df)}")
                print(f"Artigos com Pr.Cx > 0: {(self.df['Pr.Cx'] > 0).sum()} ({100*(self.df['Pr.Cx'] > 0).sum()/len(self.df):.1f}%)")
                
                # Análise detalhada
                if (self.df['Pr.Cx'] > 0).sum() == 0:
                    print(f"\n⚠️ Pr.Cx é zero para todos os artigos! Análise:")
                    
                    # Verificar por que razão
                    print(f"1. Artigos com Unit Sales > 0: {(self.df['Unit Sales'] > 0).sum()}")
                    print(f"2. Destes, com Stock Total < Meta: {((self.df['Unit Sales'] > 0) & (self.df['Stock Total'] < 0.5 * self.df['Unit Sales'])).sum()}")
                    
                    # Verificar o maior exemplo de necessidade
                    mask_analise = (self.df['Unit Sales'] > 0) & (self.df['Stock Total'] < 0.5 * self.df['Unit Sales'])
                    if mask_analise.any():
                        print(f"\nMaior exemplo de necessidade:")
                        # Encontrar a maior necessidade percentual
                        maior_idx = (self.df['Unit Sales'] * 0.5 - self.df['Stock Total']).idxmax()
                        exemplo = self.df.loc[maior_idx]
                        necessidade_exemplo = max(0, 0.5 * exemplo['Unit Sales'] - exemplo['Stock Total'])
                        caixas_exemplo = np.ceil(necessidade_exemplo / exemplo['Sup.Pack Size']).astype(int) if exemplo['Sup.Pack Size'] > 0 else 0
                        print(f"  Unit Sales: {exemplo['Unit Sales']}")
                        print(f"  Stock Total: {exemplo['Stock Total']}")
                        print(f"  Meta (25%): {0.25 * exemplo['Unit Sales']:.1f}")
                        print(f"  Sup.Pack Size: {exemplo['Sup.Pack Size']}")
                        print(f"  Necessidade: {necessidade_exemplo:.2f}")
                        print(f"  Pr.Cx calculado: {caixas_exemplo}")
                        print(f"  Razão necessidade/sup: {necessidade_exemplo/exemplo['Sup.Pack Size']:.4f}")
            
            print(f"\n=== RESUMO FINAL ===")
            print(f"Unit Sales: {self.df['Unit Sales'].min():.0f} a {self.df['Unit Sales'].max():.0f}")
            print(f"Stock Total: {self.df['Stock Total'].min():.0f} a {self.df['Stock Total'].max():.0f}")
            print(f"Sup.Pack Size: {self.df['Sup.Pack Size'].min():.0f} a {self.df['Sup.Pack Size'].max():.0f}")
            print(f"%Vendas: {self.df['%Vendas'].min():.1f}% a {self.df['%Vendas'].max():.1f}%")
            if 'Pr.Cx' in self.df.columns:
                print(f"Pr.Cx: {self.df['Pr.Cx'].min()} a {self.df['Pr.Cx'].max()}")
                print(f"Artigos com Pr.Cx=1: {(self.df['Pr.Cx'] == 1).sum()}")
                print(f"Artigos com Pr.Cx=2: {(self.df['Pr.Cx'] == 2).sum()}")
                print(f"Artigos com Pr.Cx>=3: {(self.df['Pr.Cx'] >= 3).sum()}")
            
        except Exception as e:
            print(f"Erro ao calcular %Vendas e Pr.Cx: {e}")
            import traceback
            traceback.print_exc()
            self.df['%Vendas'] = 0
            self.df['Pr.Cx'] = 0
    
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
                    # Tentar detetar automaticamente o delimitador e encoding
                    self.df = self.carregar_csv(file_path)
                else:
                    QMessageBox.critical(self, "Erro", "Formato de ficheiro não suportado.")
                    self.progress_bar.setVisible(False)
                    return
                
                self.progress_bar.setValue(50)
                
                # Verificar se as colunas necessárias existem
                colunas_necessarias = ['Sku', 'Description', 'Unit Sales', 'Sales Value', 'Stock', 'Merc.Struct Code']
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
                
                # Calcular %Vendas
                self.calcular_percentual_vendas()
                
                # Ordenar por Unit Sales (decrescente) por padrão
                self.df = self.df.sort_values('Unit Sales', ascending=False)
                
                # Preencher combobox com secções únicas
                seccoes = sorted(self.df['Secção'].unique())
                self.combo_seccao.clear()
                self.combo_seccao.addItem("Todas as Secções")
                self.combo_seccao.addItems([str(sec) for sec in seccoes])
                
                self.progress_bar.setValue(100)
                
                # Atualizar interface
                self.label_file.setText(os.path.basename(file_path))
                self.btn_exportar_excel.setEnabled(True)
                self.btn_exportar_pdf.setEnabled(True)  # Ativar botão PDF também
                self.filtrar_por_seccao()
                
                QMessageBox.information(
                    self, 
                    "Sucesso", 
                    f"Ficheiro carregado com sucesso!\n{len(self.df)} artigos encontrados.\nTipo: {file_extension.upper()}"
                )
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar ficheiro: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def carregar_csv(self, file_path):
        """Carrega ficheiro CSV com deteção automática de delimitador e encoding"""
        
        # Tentar diferentes encodings
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                # Ler as primeiras linhas para detetar o delimitador
                with open(file_path, 'r', encoding=encoding) as f:
                    first_lines = [f.readline() for _ in range(5)]
                
                # Detetar delimitador mais comum
                delimiters = [',', ';', '\t', '|']
                delimiter_scores = {}
                
                for delimiter in delimiters:
                    score = 0
                    for line in first_lines:
                        if line:
                            score += line.count(delimiter)
                    delimiter_scores[delimiter] = score
                
                # Usar o delimitador com maior score
                best_delimiter = max(delimiter_scores, key=delimiter_scores.get)
                
                # Se o melhor delimitador tiver score 0, tentar com vírgula
                if delimiter_scores[best_delimiter] == 0:
                    best_delimiter = ','
                
                # Carregar o CSV completo
                df = pd.read_csv(file_path, delimiter=best_delimiter, encoding=encoding)
                
                # Limpar espaços em branco nos nomes das colunas
                df.columns = df.columns.str.strip()
                
                print(f"CSV carregado com encoding: {encoding}, delimitador: '{best_delimiter}'")
                return df
                
            except UnicodeDecodeError:
                continue
            except Exception as e:
                print(f"Tentativa com encoding {encoding} falhou: {e}")
                continue
        
        # Se todos os encodings falharem, pedir ao utilizador
        return self.carregar_csv_manual(file_path)

    
    
    def filtrar_por_seccao(self):
        if self.df is None:
            return
        
        try:
            seccao_selecionada = self.combo_seccao.currentText()
            mostrar_todos = self.check_mostrar_todos.isChecked()
            
            if seccao_selecionada == "Todas as Secções":
                self.df_filtered = self.df.copy()
            else:
                self.df_filtered = self.df[self.df['Secção'] == seccao_selecionada].copy()
            
            # Se não mostrar todos, limitar aos top artigos
            if not mostrar_todos:
                self.df_filtered = self.df_filtered.head(100)  # Top 100 por Unit Sales
            
            self.atualizar_tabela()
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao filtrar: {str(e)}")

    def alterar_ordenacao(self, coluna):
        """Altera a coluna de ordenação"""
        self.ordenacao_atual = coluna
        self.aplicar_ordenacao()

    def alternar_ordem(self):
        """Alterna entre ordem crescente e decrescente"""
        self.ordem_decrescente = not self.ordem_decrescente
        self.btn_ordem.setText("🔽" if self.ordem_decrescente else "🔼")
        self.aplicar_ordenacao()

    def aplicar_ordenacao(self):
        """Aplica a ordenação atual aos dados"""
        if self.df is None:
            return
        
        try:
            # Calcular %Vendas se necessário
            if self.ordenacao_atual == '%Vendas' and '%Vendas' not in self.df.columns:
                self.calcular_percentual_vendas()
            
            # Ordenar o DataFrame
            coluna_ordenacao = self.ordenacao_atual
            
            # Para %Vendas, usar sempre ordenação crescente
            if coluna_ordenacao == '%Vendas':
                # Ordenar por %Vendas de forma crescente (menor % primeiro)
                self.df = self.df.sort_values(
                    coluna_ordenacao, 
                    ascending=True,  # Sempre crescente para %Vendas
                    na_position='last'
                )
            else:
                # Para outras colunas, usar a ordem selecionada
                self.df = self.df.sort_values(
                    coluna_ordenacao, 
                    ascending=not self.ordem_decrescente
                )
            
            # Reaplicar filtros se houver dados filtrados
            if self.df_filtered is not None:
                self.filtrar_por_seccao()
            else:
                self.atualizar_tabela()
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ordenar: {str(e)}")
        
    def atualizar_tabela(self):
        if self.df_filtered is None:
            return
        
        try:
            # Limpar tabela antes de reconfigurar
            self.table.clear()
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            
            # Configurar tabela com 13 colunas (agora incluindo Pr.Cx)
            self.table.setRowCount(len(self.df_filtered))
            self.table.setColumnCount(13)
            self.table.setHorizontalHeaderLabels([
                'Sku', 'Description', 'Unit Sales', 'Sales Value', 'Stock', 
                'Sup.Pack Size', 'Pr.Cx', 'Presentation Stock', '%Vendas', 
                'Ultima Recepcao', 'Flow-type', 'GLP', 'Status'
            ])
            
            # Calcular valores para o gradiente de cores da Unit Sales
            if not self.df_filtered.empty:
                max_unit_sales = self.df_filtered['Unit Sales'].max()
                min_unit_sales = self.df_filtered['Unit Sales'].min()
                range_unit_sales = max_unit_sales - min_unit_sales if max_unit_sales != min_unit_sales else 1
            
            # Preencher tabela
            for row_idx, (_, row) in enumerate(self.df_filtered.iterrows()):
                # Sku
                item_sku = QTableWidgetItem(str(row.get('Sku', '')))
                item_sku.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(row_idx, 0, item_sku)
                
                # Description
                item_desc = QTableWidgetItem(str(row.get('Description', '')))
                item_desc.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(row_idx, 1, item_desc)
                
                # Unit Sales com gradiente de cores
                unit_sales_value = float(row.get('Unit Sales', 0)) if pd.notna(row.get('Unit Sales')) else 0
                item_unit_sales = QTableWidgetItem(f"{int(unit_sales_value):,}")
                item_unit_sales.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                # Aplicar gradiente de cor (verde para alto, vermelho para baixo)
                if not self.df_filtered.empty and range_unit_sales > 0:
                    normalized_value = (unit_sales_value - min_unit_sales) / range_unit_sales
                    # Verde (alto) -> Amarelo (médio) -> Vermelho (baixo)
                    if normalized_value > 0.5:
                        green = 255
                        red = int(255 * (1 - (normalized_value - 0.5) * 2))
                    else:
                        red = 255
                        green = int(255 * (normalized_value * 2))
                    
                    blue = 50
                    item_unit_sales.setBackground(QColor(red, green, blue))
                    item_unit_sales.setForeground(QColor(0, 0, 0))
                
                self.table.setItem(row_idx, 2, item_unit_sales)
                
                # Sales Value
                sales_value = float(row.get('Sales Value', 0)) if pd.notna(row.get('Sales Value')) else 0
                item_sales_value = QTableWidgetItem(f"€ {sales_value:,.2f}")
                item_sales_value.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 3, item_sales_value)
                
                # Stock (soma total) - usar Stock Total se existir
                if 'Stock Total' in row:
                    stock_total = float(row.get('Stock Total', 0)) if pd.notna(row.get('Stock Total')) else 0
                else:
                    stock = float(row.get('Stock', 0)) if pd.notna(row.get('Stock')) else 0
                    in_transit = float(row.get('Stock In Transit', 0)) if pd.notna(row.get('Stock In Transit')) else 0
                    expected = float(row.get('Stock Expected', 0)) if pd.notna(row.get('Stock Expected')) else 0
                    on_order = float(row.get('Stock On Order', 0)) if pd.notna(row.get('Stock On Order')) else 0
                    stock_total = stock + in_transit + expected + on_order
                
                item_stock = QTableWidgetItem(f"{int(stock_total):,}")
                item_stock.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 4, item_stock)
                
                # Sup.Pack Size
                sup_pack = float(row.get('Sup.Pack Size', 0)) if pd.notna(row.get('Sup.Pack Size')) else 0
                item_sup_pack = QTableWidgetItem(f"{int(sup_pack):,}")
                item_sup_pack.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 5, item_sup_pack)
                
                # Pr.Cx (Previsão de Caixas) - NOVA COLUNA
                pr_cx = float(row.get('Pr.Cx', 0)) if pd.notna(row.get('Pr.Cx')) else 0
                item_pr_cx = QTableWidgetItem(f"{int(pr_cx):,}")
                item_pr_cx.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                # Aplicar cor diferente para valores > 0
                if pr_cx > 0:
                    item_pr_cx.setBackground(QColor(255, 255, 200))  # Amarelo claro
                    item_pr_cx.setForeground(QColor(0, 0, 0))
                elif pr_cx == 0:
                    item_pr_cx.setBackground(QColor(240, 240, 240))  # Cinza claro
                    item_pr_cx.setForeground(QColor(100, 100, 100))
                
                self.table.setItem(row_idx, 6, item_pr_cx)
                
                # Presentation Stock
                pres_stock = float(row.get('Presentation Stock', 0)) if pd.notna(row.get('Presentation Stock')) else 0
                item_pres_stock = QTableWidgetItem(f"{int(pres_stock):,}")
                item_pres_stock.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 7, item_pres_stock)
                
                # %Vendas
                percentual = float(row.get('%Vendas', 0)) if pd.notna(row.get('%Vendas')) else 0
                if percentual >= 99999:  # Valor que usamos para Unit Sales = 0
                    percent_text = "N/A"
                else:
                    percent_text = f"{percentual:.1f}%"
                item_percent = QTableWidgetItem(percent_text)
                item_percent.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 8, item_percent)
                
                # Ultima Recepcao
                ultima_recepcao = str(row.get('Ultima Recepcao', 'N/A'))[:10] if pd.notna(row.get('Ultima Recepcao')) else "N/A"
                item_recepcao = QTableWidgetItem(ultima_recepcao)
                item_recepcao.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 9, item_recepcao)
                
                # Flow-type
                flow_type = str(row.get('Flow-type', 'N/A')) if pd.notna(row.get('Flow-type')) else "N/A"
                item_flow = QTableWidgetItem(flow_type)
                item_flow.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 10, item_flow)
                
                # GLP
                glp = str(row.get('GLP', 'N/A')) if pd.notna(row.get('GLP')) else "N/A"
                item_glp = QTableWidgetItem(glp)
                item_glp.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 11, item_glp)
                
                # Status
                status = str(row.get('Status', 'N/A')) if pd.notna(row.get('Status')) else "N/A"
                item_status = QTableWidgetItem(status)
                item_status.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 12, item_status)
            
            # Ajustar tamanho das colunas
            header = self.table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Sku
            header.setSectionResizeMode(1, QHeaderView.Stretch)          # Description
            header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Unit Sales
            header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Sales Value
            header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Stock
            header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Sup.Pack Size
            header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # Pr.Cx (NOVA)
            header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # Presentation Stock
            header.setSectionResizeMode(8, QHeaderView.ResizeToContents)  # %Vendas
            header.setSectionResizeMode(9, QHeaderView.ResizeToContents)  # Ultima Recepcao
            header.setSectionResizeMode(10, QHeaderView.ResizeToContents) # Flow-type
            header.setSectionResizeMode(11, QHeaderView.ResizeToContents) # GLP
            header.setSectionResizeMode(12, QHeaderView.ResizeToContents) # Status
            
            # Atualizar contador
            self.label_contador.setText(f"Total de artigos: {len(self.df_filtered):,}")
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar tabela: {str(e)}")
            import traceback
            traceback.print_exc()
    

    def alterar_ordenacao(self, coluna):
        """Altera a coluna de ordenação"""
        self.ordenacao_atual = coluna
        
        # Desativar o botão de ordem quando %Vendas estiver selecionado
        if coluna == '%Vendas':
            self.btn_ordem.setEnabled(False)
            self.btn_ordem.setToolTip("Ordenação fixa para %Vendas (crescente)")
        else:
            self.btn_ordem.setEnabled(True)
            self.btn_ordem.setToolTip("Alternar entre ordem crescente/decrescente")
        
        self.aplicar_ordenacao()

    def exportar_pdf(self):
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Aviso", "Não existem dados para exportar.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Exportar para PDF", "HitParade.pdf", "PDF (*.pdf)"
        )
        if not file_path:
            return

        try:
            # Usar o df_filtered que já tem todas as colunas calculadas
            df_pdf = self.df_filtered.copy()
            
            # ------------------- Configuração PDF (A4 Landscape) -------------------
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

            # ------------------- Título e info -------------------
            title_fmt = QTextCharFormat()
            title_fmt.setFont(QFont("Arial", 16, QFont.Bold))
            block_fmt = QTextBlockFormat()
            block_fmt.setAlignment(Qt.AlignCenter)
            cursor.insertBlock(block_fmt)
            cursor.setCharFormat(title_fmt)
            cursor.insertText("HIT PARADE POR SECÇÃO\n\n")

            info = f"Secção: {self.combo_seccao.currentText()} | " \
                f"Total artigos: {len(df_pdf):,} | " \
                f"Gerado em: {pd.Timestamp.now():%d/%m/%Y %H:%M}\n\n"
            cursor.insertText(info)

            # ------------------- Cabeçalhos (com Pr.Cx adicionada) -------------------
            headers = [
                'Sku', 'Description', 'Unit Sales', 'Sales Value',
                'Stock', 'Sup.Pack Size', 'Presentation Stock', 'Pr.Cx', '%Vendas', 
                'Ultima Recepcao', 'Flow-type', 'GLP', 'S'
            ]

            # ------------------- Larguras ajustadas -------------------
            larguras_percentagem = [
                7,   # Sku
                24,  # Description
                6,   # Unit Sales
                6,   # Sales Value
                6,   # Stock (total)
                6,   # Sup.Pack Size
                7,   # Presentation Stock
                5,   # Pr.Cx (nova coluna)
                6,   # %Vendas
                9,   # Ultima Recepcao
                7,   # Flow-type
                2,   # GLP
                2    # S
            ]  # soma = 93%

            # ------------------- Formato da tabela -------------------
            table_fmt = QTextTableFormat()
            table_fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100))
            table_fmt.setCellPadding(5)
            table_fmt.setCellSpacing(0)
            table_fmt.setBorder(0.5)
            table_fmt.setBorderStyle(QTextFrameFormat.BorderStyle_Solid)

            constraints = [QTextLength(QTextLength.PercentageLength, w) for w in larguras_percentagem]
            table_fmt.setColumnWidthConstraints(constraints)

            table = cursor.insertTable(len(df_pdf) + 1, len(headers), table_fmt)

            # ------------------- Cabeçalho -------------------
            header_cell_fmt = QTextTableCellFormat()
            header_cell_fmt.setBackground(QColor("#d0d0d0"))

            header_char_fmt = QTextCharFormat()
            header_char_fmt.setFontWeight(QFont.Bold)
            header_char_fmt.setFontPointSize(9)

            unit_sales_header_fmt = QTextCharFormat(header_char_fmt)
            unit_sales_header_fmt.setFontPointSize(10)

            for col, texto in enumerate(headers):
                cell = table.cellAt(0, col)
                cell.setFormat(header_cell_fmt)
                cur = cell.firstCursorPosition()
                if texto == "Unit Sales":
                    cur.insertText(texto, unit_sales_header_fmt)
                else:
                    cur.insertText(texto, header_char_fmt)

            # ------------------- Dados -------------------
            normal_fmt = QTextCharFormat()
            normal_fmt.setFontPointSize(8)

            bold_fmt = QTextCharFormat(normal_fmt)
            bold_fmt.setFontWeight(QFont.Bold)
            bold_fmt.setFontPointSize(9)

            for row_idx, (_, row) in enumerate(df_pdf.iterrows(), start=1):
                for col_idx, col_name in enumerate(headers):
                    cell = table.cellAt(row_idx, col_idx)
                    cur = cell.firstCursorPosition()

                    # Mapear nomes de colunas
                    if col_name == 'S':
                        value = row.get('Status', '')
                    elif col_name == 'Stock':
                        # Usar Stock Total (que já foi calculado)
                        value = row.get('Stock Total', row.get('Stock', 0))
                    elif col_name == 'Pr.Cx':
                        value = row.get('Pr.Cx', 0)
                    else:
                        value = row.get(col_name, '')

                    if pd.isna(value):
                        text = "N/A"
                    else:
                        if col_name == "Description":
                            desc = str(value)
                            text = desc if len(desc) <= 45 else desc[:42] + "..."
                        elif col_name in ["Unit Sales", "Stock", "Sup.Pack Size", "Presentation Stock", "Pr.Cx"]:
                            text = f"{int(value):,}" if value else "0"
                        elif col_name == "Sales Value":
                            text = f"€{float(value):,.0f}" if value else "€0"
                        elif col_name == "%Vendas":
                            text = "N/A" if value == 99999 else f"{value:.1f}%"
                        elif col_name == "Ultima Recepcao":
                            text = str(value)[:10] if pd.notna(value) else "N/A"
                        else:
                            text = str(value)

                    if col_name == "Unit Sales":
                        cur.insertText(text, bold_fmt)
                    else:
                        cur.insertText(text, normal_fmt)

            # ------------------- Rodapé -------------------
            cursor.movePosition(QTextCursor.End)
            cursor.insertBlock()
            footer = QTextCharFormat()
            footer.setFontPointSize(7)
            footer.setFontItalic(True)
            footer.setForeground(QColor("gray"))
            cursor.setCharFormat(footer)
            cursor.insertText(f"Documento gerado automaticamente • {len(df_pdf):,} artigos • Pr.Cx = Previsão de Caixas (meta: 25% do stock atual)")

            # ------------------- Exportar -------------------
            doc.print_(printer)

            QMessageBox.information(
                self, "Sucesso",
                f"PDF exportado com sucesso!\n\n"
                f"→ {len(df_pdf):,} artigos exportados\n"
                f"→ Guardado em: {os.path.basename(file_path)}\n"
                f"→ Pr.Cx: Previsão de caixas para atingir 25% do stock atual"
            )

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar PDF:\n{str(e)}")
            import traceback; traceback.print_exc()

    def exportar_excel(self):
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados para exportar.")
            return
        
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Exportar para Excel",
                "hit_parade_export.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if file_path:
                self.progress_bar.setVisible(True)
                self.progress_bar.setValue(50)
                
                # Usar Stock Total que já foi calculado na função calcular_percentual_vendas
                # Renomear 'Stock Total' para 'Stock' na exportação
                df_export = self.df_filtered.copy()
                
                # Renomear Stock Total para Stock (já que é a soma de todas as colunas de stock)
                if 'Stock Total' in df_export.columns:
                    df_export['Stock'] = df_export['Stock Total']
                
                # Definir colunas para exportação
                colunas_export = ['Sku', 'Description', 'EAN', 'Unit Sales', 'Sales Value', 'Stock',
                                '%Vendas', 'Pr.Cx', 'Ultima Recepcao', 'Flow-type', 'Status', 
                                'Secção', 'GLP', 'Sup.Pack Size', 'Presentation Stock']
                
                # Filtrar apenas colunas que existem no DataFrame
                colunas_disponiveis = [col for col in colunas_export if col in df_export.columns]
                df_export = df_export[colunas_disponiveis]
                
                # Exportar para Excel
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Hit Parade')
                    
                    # Acessar a worksheet para ajustar as colunas
                    worksheet = writer.sheets['Hit Parade']
                    
                    # Ajustar largura das colunas baseado no conteúdo
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        
                        # Encontrar o comprimento máximo na coluna
                        for cell in column:
                            try:
                                # Calcular comprimento do conteúdo
                                if cell.value:
                                    cell_length = len(str(cell.value))
                                    max_length = max(max_length, cell_length)
                            except:
                                pass
                        
                        # Ajustar largura (com margem de segurança)
                        adjusted_width = min(max_length + 2, 50)  # Máximo de 50 caracteres
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                    
                self.progress_bar.setValue(100)
                
                QMessageBox.information(
                    self, 
                    "Sucesso", 
                    f"Dados exportados com sucesso!\n{len(df_export)} artigos exportados."
                )
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar: {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            self.progress_bar.setVisible(False)

    def limpar_tudo(self):
        self.df = None
        self.df_filtered = None
        self.table.setRowCount(0)
        self.label_file.setText("Nenhum ficheiro carregado")
        self.combo_seccao.clear()
        self.combo_seccao.addItem("Todas as Secções")
        self.check_mostrar_todos.setChecked(False)
        self.label_contador.setText("Total de artigos: 0")
        self.btn_exportar_excel.setEnabled(False)
        self.btn_exportar_pdf.setEnabled(False)  # Desativar botão PDF também

def mostrar_hit_parade():
    dialog = HitParadeDialog()
    dialog.exec_()