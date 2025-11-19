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

class TendenciasDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Análise de Tendências")
        self.setGeometry(100, 100, 1400, 800)
        self.df_tendencias = None
        self.df_filtered = None
        self.ordenacao_atual = '% Crescimento'  # Ordenação padrão
        self.ordem_decrescente = True  # Ordem padrão decrescente
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        # Título
        title = QLabel("Análise de Tendências - Comparação de Vendas")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("margin: 20px;")
        layout.addWidget(title)
        
        # Área de upload dos 2 ficheiros
        upload_layout1 = QHBoxLayout()
        self.btn_file1 = QPushButton("📁 Carregar Ficheiro Período 1")
        self.btn_file1.setFont(QFont("Arial", 12))
        self.btn_file1.setMinimumHeight(40)
        self.btn_file1.setStyleSheet("""
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
        self.btn_file1.clicked.connect(lambda: self.carregar_ficheiro(1))
        upload_layout1.addWidget(self.btn_file1)
        
        self.label_file1 = QLabel("Nenhum ficheiro carregado")
        self.label_file1.setStyleSheet("color: #666; padding: 10px;")
        upload_layout1.addWidget(self.label_file1)
        upload_layout1.addStretch()
        layout.addLayout(upload_layout1)
        
        upload_layout2 = QHBoxLayout()
        self.btn_file2 = QPushButton("📁 Carregar Ficheiro Período 2")
        self.btn_file2.setFont(QFont("Arial", 12))
        self.btn_file2.setMinimumHeight(40)
        self.btn_file2.setStyleSheet("""
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
        """)
        self.btn_file2.clicked.connect(lambda: self.carregar_ficheiro(2))
        upload_layout2.addWidget(self.btn_file2)
        
        self.label_file2 = QLabel("Nenhum ficheiro carregado")
        self.label_file2.setStyleSheet("color: #666; padding: 10px;")
        upload_layout2.addWidget(self.label_file2)
        upload_layout2.addStretch()
        layout.addLayout(upload_layout2)
        
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
        self.combo_ordenacao.addItems(["% Crescimento", "Unit Sales P1", "Unit Sales P2", "Sales Value P1", "Sales Value P2"])
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
        
        # Variáveis para armazenar os DataFrames
        self.df_periodo1 = None
        self.df_periodo2 = None

    def carregar_ficheiro(self, periodo):
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, 
                f"Selecionar Ficheiro - Período {periodo}", 
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
                    df = pd.read_excel(file_path)
                elif file_extension == 'csv':
                    df = self.carregar_csv(file_path)
                else:
                    QMessageBox.critical(self, "Erro", "Formato de ficheiro não suportado.")
                    self.progress_bar.setVisible(False)
                    return
                
                self.progress_bar.setValue(50)
                
                # Verificar se as colunas necessárias existem
                colunas_necessarias = ['Sku', 'Description', 'Unit Sales', 'Sales Value', 'Merc.Struct Code']
                colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
                
                if colunas_faltantes:
                    QMessageBox.critical(
                        self, 
                        "Erro", 
                        f"Colunas faltantes no ficheiro do Período {periodo}: {', '.join(colunas_faltantes)}"
                    )
                    self.progress_bar.setVisible(False)
                    return
                
                # Extrair secção do Merc.Struct Code
                df['Secção'] = df['Merc.Struct Code'].astype(str).str[2:4]
                
                # Armazenar o DataFrame conforme o período
                if periodo == 1:
                    self.df_periodo1 = df
                    self.label_file1.setText(os.path.basename(file_path))
                else:
                    self.df_periodo2 = df
                    self.label_file2.setText(os.path.basename(file_path))
                
                self.progress_bar.setValue(100)
                
                # Verificar se ambos os ficheiros foram carregados
                if self.df_periodo1 is not None and self.df_periodo2 is not None:
                    self.processar_tendencias()
                
                QMessageBox.information(
                    self, 
                    "Sucesso", 
                    f"Ficheiro do Período {periodo} carregado com sucesso!\n{len(df)} artigos encontrados."
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
            "Atenção", 
            "Não foi possível detetar automaticamente o formato do CSV.\n"
            "Por favor, selecione as opções manualmente."
        )
        
        # Implementar diálogo manual se necessário
        # Por enquanto, tentar com encoding padrão e vírgula
        try:
            return pd.read_csv(file_path, delimiter=',', encoding='utf-8')
        except:
            try:
                return pd.read_csv(file_path, delimiter=';', encoding='latin-1')
            except Exception as e:
                raise Exception(f"Não foi possível ler o ficheiro CSV: {str(e)}")

    def processar_tendencias(self):
        """Processa a comparação entre os dois períodos"""
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # Fazer merge dos dois DataFrames pelo SKU
            df_merge = pd.merge(
                self.df_periodo1[['Sku', 'Description', 'Unit Sales', 'Sales Value', 'Secção']],
                self.df_periodo2[['Sku', 'Description', 'Unit Sales', 'Sales Value', 'Secção']],
                on='Sku',
                suffixes=('_P1', '_P2'),
                how='outer'
            )
            
            self.progress_bar.setValue(30)
            
            # Preencher valores NaN com 0
            df_merge['Unit Sales_P1'] = df_merge['Unit Sales_P1'].fillna(0)
            df_merge['Unit Sales_P2'] = df_merge['Unit Sales_P2'].fillna(0)
            df_merge['Sales Value_P1'] = df_merge['Sales Value_P1'].fillna(0)
            df_merge['Sales Value_P2'] = df_merge['Sales Value_P2'].fillna(0)
            
            # Usar Description do P2 se P1 estiver vazio, e vice-versa
            df_merge['Description_P1'] = df_merge['Description_P1'].fillna(df_merge['Description_P2'])
            df_merge['Description_P2'] = df_merge['Description_P2'].fillna(df_merge['Description_P1'])
            
            # Usar a Description do P2 como principal (ou P1 se P2 não existir)
            df_merge['Description'] = df_merge['Description_P2'].fillna(df_merge['Description_P1'])
            
            # Usar Secção do P2 se P1 estiver vazio, e vice-versa
            df_merge['Secção_P1'] = df_merge['Secção_P1'].fillna(df_merge['Secção_P2'])
            df_merge['Secção_P2'] = df_merge['Secção_P2'].fillna(df_merge['Secção_P1'])
            df_merge['Secção'] = df_merge['Secção_P2']
            
            self.progress_bar.setValue(60)
            
            # Calcular crescimento percentual
            def calcular_crescimento(p1, p2):
                if p1 == 0:
                    if p2 == 0:
                        return 0
                    else:
                        return 99999  # Crescimento infinito (novo produto)
                else:
                    return ((p2 - p1) / p1) * 100
            
            df_merge['% Crescimento'] = df_merge.apply(
                lambda row: calcular_crescimento(row['Unit Sales_P1'], row['Unit Sales_P2']), 
                axis=1
            )
            
            # Arredondar para 2 casas decimais
            df_merge['% Crescimento'] = df_merge['% Crescimento'].round(2)
            
            self.progress_bar.setValue(80)
            
            # Ordenar por % Crescimento (decrescente)
            df_merge = df_merge.sort_values('% Crescimento', ascending=False)
            
            # Atualizar DataFrame principal
            self.df_tendencias = df_merge
            
            # Preencher combobox com secções únicas
            seccoes = sorted(self.df_tendencias['Secção'].unique())
            self.combo_seccao.clear()
            self.combo_seccao.addItem("Todas as Secções")
            self.combo_seccao.addItems([str(sec) for sec in seccoes if pd.notna(sec)])
            
            self.progress_bar.setValue(100)
            
            # Atualizar interface
            self.btn_exportar_excel.setEnabled(True)
            self.btn_exportar_pdf.setEnabled(True)
            self.filtrar_por_seccao()
            
            QMessageBox.information(
                self, 
                "Sucesso", 
                f"Análise de tendências concluída!\n{len(self.df_tendencias)} artigos processados."
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao processar tendências: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def filtrar_por_seccao(self):
        if self.df_tendencias is None:
            return
        
        try:
            seccao_selecionada = self.combo_seccao.currentText()
            mostrar_todos = self.check_mostrar_todos.isChecked()
            
            if seccao_selecionada == "Todas as Secções":
                self.df_filtered = self.df_tendencias.copy()
            else:
                self.df_filtered = self.df_tendencias[self.df_tendencias['Secção'] == seccao_selecionada].copy()
            
            # Se não mostrar todos, limitar aos top artigos
            if not mostrar_todos:
                self.df_filtered = self.df_filtered.head(100)
            
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
        if self.df_tendencias is None:
            return
        
        try:
            # Mapear nomes das colunas para os nomes reais no DataFrame
            coluna_map = {
                '% Crescimento': '% Crescimento',
                'Unit Sales P1': 'Unit Sales_P1',
                'Unit Sales P2': 'Unit Sales_P2',
                'Sales Value P1': 'Sales Value_P1',
                'Sales Value P2': 'Sales Value_P2'
            }
            
            coluna_ordenacao = coluna_map.get(self.ordenacao_atual, '% Crescimento')
            
            # Ordenar o DataFrame
            self.df_tendencias = self.df_tendencias.sort_values(
                coluna_ordenacao, 
                ascending=not self.ordem_decrescente
            )
            
            # Reaplicar filtros
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
            # Configurar tabela com 9 colunas
            self.table.setRowCount(len(self.df_filtered))
            self.table.setColumnCount(9)
            self.table.setHorizontalHeaderLabels([
                'Sku', 'Description', 'Unit Sales P1', 'Unit Sales P2', 
                'Sales Value P1', 'Sales Value P2', '% Crescimento', 
                'Secção', 'Tendência'
            ])
            
            # Calcular valores para o gradiente de cores do % Crescimento
            if not self.df_filtered.empty:
                crescimento_values = self.df_filtered['% Crescimento'].replace([99999, -99999], np.nan)
                max_crescimento = crescimento_values.max()
                min_crescimento = crescimento_values.min()
                range_crescimento = max_crescimento - min_crescimento if max_crescimento != min_crescimento else 1
            
            # Preencher tabela
            for row_idx, (_, row) in enumerate(self.df_filtered.iterrows()):
                # Sku
                item_sku = QTableWidgetItem(str(row['Sku']))
                item_sku.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(row_idx, 0, item_sku)
                
                # Description
                item_desc = QTableWidgetItem(str(row['Description']))
                item_desc.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(row_idx, 1, item_desc)
                
                # Unit Sales P1
                unit_sales_p1 = row['Unit Sales_P1'] if pd.notna(row['Unit Sales_P1']) else 0
                item_unit_sales_p1 = QTableWidgetItem(f"{unit_sales_p1:,.0f}")
                item_unit_sales_p1.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 2, item_unit_sales_p1)
                
                # Unit Sales P2
                unit_sales_p2 = row['Unit Sales_P2'] if pd.notna(row['Unit Sales_P2']) else 0
                item_unit_sales_p2 = QTableWidgetItem(f"{unit_sales_p2:,.0f}")
                item_unit_sales_p2.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 3, item_unit_sales_p2)
                
                # Sales Value P1
                sales_value_p1 = row['Sales Value_P1'] if pd.notna(row['Sales Value_P1']) else 0
                item_sales_value_p1 = QTableWidgetItem(f"€ {sales_value_p1:,.2f}")
                item_sales_value_p1.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 4, item_sales_value_p1)
                
                # Sales Value P2
                sales_value_p2 = row['Sales Value_P2'] if pd.notna(row['Sales Value_P2']) else 0
                item_sales_value_p2 = QTableWidgetItem(f"€ {sales_value_p2:,.2f}")
                item_sales_value_p2.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 5, item_sales_value_p2)
                
                # % Crescimento com gradiente de cores
                percent_crescimento = row['% Crescimento'] if pd.notna(row['% Crescimento']) else 0
                
                if percent_crescimento == 99999:
                    percent_text = "Novo"
                elif percent_crescimento == -100:
                    percent_text = "Descontinuado"
                else:
                    percent_text = f"{percent_crescimento:+.1f}%"
                
                item_percent = QTableWidgetItem(percent_text)
                item_percent.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                # Aplicar gradiente de cor (verde para crescimento positivo, vermelho para negativo)
                if percent_crescimento != 99999 and percent_crescimento != -100:
                    if not self.df_filtered.empty and range_crescimento > 0:
                        # Normalizar valor entre -1 e 1 para o gradiente
                        if percent_crescimento >= 0:
                            normalized_value = min(percent_crescimento / max(1, max_crescimento), 1)
                            # Verde escuro para alto crescimento, verde claro para baixo crescimento positivo
                            green = 255
                            red = int(255 * (1 - normalized_value))
                        else:
                            normalized_value = min(abs(percent_crescimento) / max(1, abs(min_crescimento)), 1)
                            # Vermelho escuro para grande queda, vermelho claro para pequena queda
                            red = 255
                            green = int(255 * (1 - normalized_value))
                        
                        blue = 50
                        item_percent.setBackground(QColor(red, green, blue))
                        item_percent.setForeground(QColor(0, 0, 0))
                elif percent_crescimento == 99999:
                    item_percent.setBackground(QColor(0, 255, 0))  # Verde forte para novos
                    item_percent.setForeground(QColor(0, 0, 0))
                elif percent_crescimento == -100:
                    item_percent.setBackground(QColor(255, 0, 0))  # Vermelho forte para descontinuados
                    item_percent.setForeground(QColor(255, 255, 255))
                
                self.table.setItem(row_idx, 6, item_percent)
                
                # Secção
                seccao = str(row['Secção']) if pd.notna(row['Secção']) else "N/A"
                item_seccao = QTableWidgetItem(seccao)
                item_seccao.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 7, item_seccao)
                
                # Tendência (indicador visual)
                if percent_crescimento == 99999:
                    tendencia_text = "📈 NOVO"
                elif percent_crescimento == -100:
                    tendencia_text = "📉 DESCONT."
                elif percent_crescimento > 20:
                    tendencia_text = "📈 ALTA"
                elif percent_crescimento > 0:
                    tendencia_text = "↗️ SUBIU"
                elif percent_crescimento > -20:
                    tendencia_text = "↘️ BAIXOU"
                else:
                    tendencia_text = "📉 QUEDA"
                
                item_tendencia = QTableWidgetItem(tendencia_text)
                item_tendencia.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 8, item_tendencia)
            
            # Ajustar tamanho das colunas
            header = self.table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Sku
            header.setSectionResizeMode(1, QHeaderView.Stretch)          # Description
            header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Unit Sales P1
            header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Unit Sales P2
            header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Sales Value P1
            header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Sales Value P2
            header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # % Crescimento
            header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # Secção
            header.setSectionResizeMode(8, QHeaderView.ResizeToContents)  # Tendência
            
            # Atualizar contador
            self.label_contador.setText(f"Total de artigos: {len(self.df_filtered):,}")
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar tabela: {str(e)}")

    def exportar_pdf(self):
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Aviso", "Não existem dados para exportar.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Exportar para PDF", "Tendencias.pdf", "PDF (*.pdf)"
        )
        if not file_path:
            return

        try:
            # Configuração PDF (A4 Landscape)
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
            cursor.insertText("ANÁLISE DE TENDÊNCIAS - COMPARAÇÃO DE PERÍODOS\n\n")

            info = f"Secção: {self.combo_seccao.currentText()} | " \
                f"Total artigos: {len(self.df_filtered):,} | " \
                f"Gerado em: {pd.Timestamp.now():%d/%m/%Y %H:%M}\n\n"
            cursor.insertText(info)

            # Cabeçalhos
            headers = [
                'Sku', 'Description', 'Unit P1', 'Unit P2', 'Value P1',
                'Value P2', '% Cresc.', 'Secção', 'Tendência'
            ]

            # Larguras ajustadas
            larguras_percentagem = [
                8,   # Sku
                30,  # Description
                7,   # Unit P1
                7,   # Unit P2
                8,   # Value P1
                8,   # Value P2
                8,   # % Cresc.
                6,   # Secção
                8    # Tendência
            ]  # soma = 100%

            # Formato da tabela
            table_fmt = QTextTableFormat()
            table_fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100))
            table_fmt.setCellPadding(5)
            table_fmt.setCellSpacing(0)
            table_fmt.setBorder(0.5)
            table_fmt.setBorderStyle(QTextFrameFormat.BorderStyle_Solid)

            constraints = [QTextLength(QTextLength.PercentageLength, w) for w in larguras_percentagem]
            table_fmt.setColumnWidthConstraints(constraints)

            table = cursor.insertTable(len(self.df_filtered) + 1, len(headers), table_fmt)

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

            for row_idx, (_, row) in enumerate(self.df_filtered.iterrows(), start=1):
                for col_idx, col_name in enumerate(headers):
                    cell = table.cellAt(row_idx, col_idx)
                    cur = cell.firstCursorPosition()

                    if col_name == 'Sku':
                        text = str(row['Sku'])
                    elif col_name == 'Description':
                        desc = str(row['Description'])
                        text = desc if len(desc) <= 35 else desc[:32] + "..."
                    elif col_name == 'Unit P1':
                        text = f"{int(row['Unit Sales_P1']):,}" if row['Unit Sales_P1'] else "0"
                    elif col_name == 'Unit P2':
                        text = f"{int(row['Unit Sales_P2']):,}" if row['Unit Sales_P2'] else "0"
                    elif col_name == 'Value P1':
                        text = f"€{float(row['Sales Value_P1']):,.0f}" if row['Sales Value_P1'] else "€0"
                    elif col_name == 'Value P2':
                        text = f"€{float(row['Sales Value_P2']):,.0f}" if row['Sales Value_P2'] else "€0"
                    elif col_name == '% Cresc.':
                        percent = row['% Crescimento']
                        if percent == 99999:
                            text = "Novo"
                        elif percent == -100:
                            text = "Descont."
                        else:
                            text = f"{percent:+.1f}%"
                    elif col_name == 'Secção':
                        text = str(row['Secção']) if pd.notna(row['Secção']) else "N/A"
                    elif col_name == 'Tendência':
                        percent = row['% Crescimento']
                        if percent == 99999:
                            text = "NOVO"
                        elif percent == -100:
                            text = "DESCONT."
                        elif percent > 20:
                            text = "ALTA"
                        elif percent > 0:
                            text = "SUBIU"
                        elif percent > -20:
                            text = "BAIXOU"
                        else:
                            text = "QUEDA"
                    else:
                        text = str(row[col_name]) if pd.notna(row[col_name]) else "N/A"

                    cur.insertText(text, normal_fmt)

            # Rodapé
            cursor.movePosition(QTextCursor.End)
            cursor.insertBlock()
            footer = QTextCharFormat()
            footer.setFontPointSize(7)
            footer.setFontItalic(True)
            footer.setForeground(QColor("gray"))
            cursor.setCharFormat(footer)
            cursor.insertText(f"Análise de tendências • {len(self.df_filtered):,} artigos comparados")

            # Exportar
            doc.print_(printer)

            QMessageBox.information(
                self, "Sucesso",
                f"PDF exportado com sucesso!\n\n"
                f"→ {len(self.df_filtered):,} artigos exportados\n"
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
                "tendencias_export.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if file_path:
                self.progress_bar.setVisible(True)
                self.progress_bar.setValue(50)
                
                # Criar DataFrame para exportação
                df_export = self.df_filtered[[
                    'Sku', 'Description', 'Unit Sales_P1', 'Unit Sales_P2', 
                    'Sales Value_P1', 'Sales Value_P2', '% Crescimento', 'Secção'
                ]].copy()
                
                # Renomear colunas para melhor legibilidade
                df_export.columns = [
                    'Sku', 'Description', 'Unit Sales Período 1', 'Unit Sales Período 2',
                    'Sales Value Período 1', 'Sales Value Período 2', '% Crescimento', 'Secção'
                ]
                
                # Adicionar coluna de tendência
                def classificar_tendencia(percent):
                    if percent == 99999:
                        return "NOVO PRODUTO"
                    elif percent == -100:
                        return "DESCONTINUADO"
                    elif percent > 20:
                        return "ALTA SIGNIFICATIVA"
                    elif percent > 0:
                        return "CRESCIMENTO"
                    elif percent > -20:
                        return "LEVE QUEDA"
                    else:
                        return "QUEDA SIGNIFICATIVA"
                
                df_export['Tendência'] = df_export['% Crescimento'].apply(classificar_tendencia)
                
                # Exportar para Excel
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Tendências')
                    
                    # Acessar a worksheet para ajustar as colunas
                    worksheet = writer.sheets['Tendências']
                    
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
        self.df_tendencias = None
        self.df_filtered = None
        self.df_periodo1 = None
        self.df_periodo2 = None
        self.table.setRowCount(0)
        self.label_file1.setText("Nenhum ficheiro carregado")
        self.label_file2.setText("Nenhum ficheiro carregado")
        self.combo_seccao.clear()
        self.combo_seccao.addItem("Todas as Secções")
        self.check_mostrar_todos.setChecked(False)
        self.label_contador.setText("Total de artigos: 0")
        self.btn_exportar_excel.setEnabled(False)
        self.btn_exportar_pdf.setEnabled(False)

def mostrar_tendencias():
    dialog = TendenciasDialog()
    dialog.exec_()