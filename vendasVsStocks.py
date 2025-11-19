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

class VendasVsStocksDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Análise Vendas vs Stocks")
        self.setGeometry(100, 100, 1400, 800)
        self.df = None
        self.df_filtered = None
        self.ordenacao_atual = 'Qty'
        self.ordem_decrescente = True
        self.coluna_uneg = None  # NOVO: armazenar o nome real da coluna
        self.initUI()
        self.btn_ordem.setText("Down Arrow")

    def initUI(self):
        layout = QVBoxLayout()
        
        # Título
        title = QLabel("Análise Vendas vs Stocks")
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

        # Filtro por Secção (U.Neg.)
        filters_layout.addWidget(QLabel("Filtrar por Secção (U.Neg.):"))
        self.combo_seccao = QComboBox()
        self.combo_seccao.setMinimumWidth(220)
        self.combo_seccao.addItem("Todas as Secções")
        self.combo_seccao.currentTextChanged.connect(self.filtrar_dados)
        filters_layout.addWidget(self.combo_seccao)

        # Checkbox mostrar todos
        self.check_mostrar_todos = QCheckBox("Mostrar todos os artigos")
        self.check_mostrar_todos.stateChanged.connect(self.filtrar_dados)
        filters_layout.addWidget(self.check_mostrar_todos)

        filters_layout.addStretch()

        # Ordenação
        filters_layout.addWidget(QLabel("Ordenar por:"))
        self.combo_ordenacao = QComboBox()
        self.combo_ordenacao.addItems([
            "Qty", "Stock", "%Stock/Vendas", "Margem %", "Vendas -Descontos"
        ])
        self.combo_ordenacao.setCurrentText("Qty")
        self.combo_ordenacao.currentTextChanged.connect(self.alterar_ordenacao)
        filters_layout.addWidget(self.combo_ordenacao)

        self.btn_ordem = QPushButton("Down Arrow")
        self.btn_ordem.setToolTip("Alternar crescente/decrescente")
        self.btn_ordem.setFixedSize(32, 32)
        self.btn_ordem.clicked.connect(self.alternar_ordem)
        filters_layout.addWidget(self.btn_ordem)

        filters_layout.addStretch()

        # Contador
        self.label_contador = QLabel("Total de artigos: 0")
        self.label_contador.setStyleSheet("font-weight: bold; color: #333;")
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

    def normalizar_nome_coluna(self, nome):
        """Normaliza o nome da coluna removendo espaços extras"""
        if pd.isna(nome):
            return ""
        return str(nome).strip()

    def encontrar_coluna_uneg(self):
        """Encontra a coluna U.Neg. mesmo com variações no nome"""
        if self.df is None:
            return None
        
        # Procurar por variações do nome
        possiveis_nomes = ['U.Neg.', 'U.Neg', 'UNeg', 'U Neg', 'Uneg', 'UNEG']
        
        for col in self.df.columns:
            col_normalizado = self.normalizar_nome_coluna(col)
            if col_normalizado in possiveis_nomes:
                print(f"✓ Coluna U.Neg. encontrada: '{col}'")
                return col
        
        # Se não encontrar exato, procurar por contém
        for col in self.df.columns:
            col_lower = str(col).lower().replace(' ', '').replace('.', '')
            if 'uneg' in col_lower:
                print(f"✓ Coluna U.Neg. encontrada (aproximada): '{col}'")
                return col
        
        print("✗ Coluna U.Neg. NÃO encontrada")
        print(f"Colunas disponíveis: {list(self.df.columns)}")
        return None

    def converter_numero_portugues(self, valor):
        """Converte números no formato português (1.234,56) para float"""
        if pd.isna(valor) or valor == '':
            return 0.0
        
        try:
            if isinstance(valor, (int, float)):
                return float(valor)
            
            valor_str = str(valor).strip().replace(' ', '')
            
            if not valor_str:
                return 0.0
            
            if '.' in valor_str and ',' in valor_str:
                partes = valor_str.split(',')
                parte_inteira = partes[0].replace('.', '')
                parte_decimal = partes[1] if len(partes) > 1 else '0'
                return float(f"{parte_inteira}.{parte_decimal}")
            elif ',' in valor_str and '.' not in valor_str:
                return float(valor_str.replace(',', '.'))
            else:
                return float(valor_str)
                
        except (ValueError, AttributeError):
            return 0.0

    def calcular_indicadores(self):
        """Calcula os indicadores de stock vs vendas"""
        try:
            colunas_numericas = ['Qty', 'Stock', 'Total P/Venda', 'Margem %', 'Vendas -Descontos']
            
            for coluna in colunas_numericas:
                if coluna in self.df.columns:
                    self.df[coluna] = self.df[coluna].apply(self.converter_numero_portugues)
            
            self.df['%Stock/Vendas'] = 0
            mask = (self.df['Qty'] > 0) & (self.df['Stock'].notna())
            self.df.loc[mask, '%Stock/Vendas'] = (self.df.loc[mask, 'Stock'] / self.df.loc[mask, 'Qty']) * 100
            self.df.loc[self.df['Qty'] == 0, '%Stock/Vendas'] = 99999
            self.df.loc[(self.df['Stock'] == 0) & (self.df['Qty'] > 0), '%Stock/Vendas'] = 0
            self.df['%Stock/Vendas'] = self.df['%Stock/Vendas'].round(2)
            
            self.df['Dias Stock'] = 0
            mask_dias = (self.df['Qty'] > 0) & (self.df['Stock'].notna())
            self.df.loc[mask_dias, 'Dias Stock'] = (self.df.loc[mask_dias, 'Stock'] / self.df.loc[mask_dias, 'Qty']) * 30
            self.df['Dias Stock'] = self.df['Dias Stock'].round(1)
            
        except Exception as e:
            print(f"Erro ao calcular indicadores: {e}")
            self.df['%Stock/Vendas'] = 0
            self.df['Dias Stock'] = 0
    
    def carregar_ficheiro(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Selecionar Ficheiro",
                "",
                "Ficheiros Suportados (*.xlsx *.xls *.csv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv)"
            )
            if not file_path:
                return

            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(10)

            file_extension = file_path.lower().split('.')[-1]

            if file_extension in ['xlsx', 'xls']:
                temp_df = pd.read_excel(file_path, header=None, nrows=30)

                header_row = 0
                for idx, row in temp_df.iterrows():
                    row_str = " ".join(row.astype(str).str.lower())
                    if "artigo" in row_str and ("descrição" in row_str or "qty" in row_str or "stock" in row_str):
                        header_row = idx
                        print(f"Cabeçalho encontrado na linha {idx + 1} (índice {idx})")
                        break

                self.df = pd.read_excel(file_path, skiprows=header_row)

            elif file_extension == 'csv':
                self.df = self.carregar_csv(file_path)
            else:
                QMessageBox.critical(self, "Erro", "Formato não suportado.")
                self.progress_bar.setVisible(False)
                return

            self.progress_bar.setValue(50)

            # Normalizar nomes das colunas
            self.df.columns = [self.normalizar_nome_coluna(col) for col in self.df.columns]
            self.df = self.df.loc[:, ~self.df.columns.str.contains('^Unnamed', na=False)]

            # Verificar colunas obrigatórias
            colunas_necessarias = ['Artigo', 'Descrição', 'Qty', 'Stock']
            colunas_faltantes = [c for c in colunas_necessarias if c not in self.df.columns]

            if colunas_faltantes:
                QMessageBox.critical(
                    self, "Colunas em falta",
                    f"Não foram encontradas as colunas obrigatórias:\n{', '.join(colunas_faltantes)}\n\n"
                    f"Colunas detetadas: {list(self.df.columns)[:15]}\n\n"
                    f"Total de linhas lidas: {len(self.df)}"
                )
                self.df = None
                self.progress_bar.setVisible(False)
                return

            # NOVO: Encontrar a coluna U.Neg.
            self.coluna_uneg = self.encontrar_coluna_uneg()

            self.calcular_indicadores()
            self.df = self.df.sort_values('Qty', ascending=False)
            self.preencher_filtros()

            self.label_file.setText(os.path.basename(file_path))
            self.btn_exportar_excel.setEnabled(True)
            self.btn_exportar_pdf.setEnabled(True)
            self.filtrar_dados()

            self.progress_bar.setValue(100)
            QMessageBox.information(
                self, "Sucesso",
                f"Ficheiro carregado!\n{len(self.df):,} artigos encontrados."
            )

        except Exception as e:
            import traceback
            print(traceback.format_exc())
            QMessageBox.critical(self, "Erro", f"Erro ao carregar o ficheiro:\n{str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def preencher_filtros(self):
        """Preenche o combo de Secção (U.Neg.) - CORRIGIDO"""
        self.combo_seccao.clear()
        self.combo_seccao.addItem("Todas as Secções")
        
        if self.coluna_uneg and self.coluna_uneg in self.df.columns:
            # Filtrar valores não nulos e únicos
            seccoes = self.df[self.coluna_uneg].dropna().unique()
            seccoes = sorted([str(s) for s in seccoes if str(s).strip() != ''])
            
            print(f"Secções encontradas: {seccoes}")
            self.combo_seccao.addItems(seccoes)
        else:
            self.combo_seccao.addItem("(coluna U.Neg. não encontrada)")
            print("Aviso: Coluna U.Neg. não está disponível para filtragem")

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
        
        return pd.read_csv(file_path)

    def filtrar_dados(self):
        """Filtra os dados por secção e limite - CORRIGIDO"""
        if self.df is None:
            return

        self.df_filtered = self.df.copy()

        # Filtrar por secção se disponível
        seccao = self.combo_seccao.currentText()
        if seccao != "Todas as Secções" and seccao != "(coluna U.Neg. não encontrada)":
            if self.coluna_uneg and self.coluna_uneg in self.df_filtered.columns:
                # CORREÇÃO: Filtro direto pela coluna
                self.df_filtered = self.df_filtered[
                    self.df_filtered[self.coluna_uneg].astype(str) == str(seccao)
                ]
                print(f"Filtrando por {self.coluna_uneg} = {seccao}: {len(self.df_filtered)} artigos")
            else:
                print(f"Aviso: Não foi possível filtrar por {seccao}")

        # Limitar a 100 se checkbox não estiver marcado
        if not self.check_mostrar_todos.isChecked():
            self.df_filtered = self.df_filtered.head(100)

        self.aplicar_ordenacao()

    def alterar_ordenacao(self, coluna):
        self.ordenacao_atual = coluna
        if coluna == "%Stock/Vendas":
            self.btn_ordem.setEnabled(False)
            self.btn_ordem.setText("Up Arrow")
        else:
            self.btn_ordem.setEnabled(True)
            self.btn_ordem.setText("Down Arrow" if self.ordem_decrescente else "Up Arrow")
        self.aplicar_ordenacao()

    def alternar_ordem(self):
        """Alterna entre ordem crescente e decrescente"""
        self.ordem_decrescente = not self.ordem_decrescente
        self.btn_ordem.setText("🔽" if self.ordem_decrescente else "🔼")
        self.aplicar_ordenacao()

    def aplicar_ordenacao(self):
        if self.df_filtered is None:
            return

        coluna = self.ordenacao_atual

        if coluna == "%Stock/Vendas":
            ascending = True
            self.btn_ordem.setEnabled(False)
            self.btn_ordem.setText("Up Arrow")
        else:
            ascending = not self.ordem_decrescente
            self.btn_ordem.setEnabled(True)
            self.btn_ordem.setText("Down Arrow" if self.ordem_decrescente else "Up Arrow")

        try:
            self.df_filtered = self.df_filtered.sort_values(
                by=coluna,
                ascending=ascending,
                na_position='last'
            )
            self.atualizar_tabela()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro na ordenação: {e}")
            
    def atualizar_tabela(self):
        if self.df_filtered is None:
            return
        
        try:
            self.table.setRowCount(len(self.df_filtered))
            self.table.setColumnCount(12)
            self.table.setHorizontalHeaderLabels([
                'Artigo', 'Descrição', 'Status', 'Cat.', 'Sub-C.', 'Qty', 
                'Stock', '%Stock/Vendas', 'Dias Stock', 'Total P/Venda', 
                'Margem %', 'Vendas -Descontos'
            ])
            
            if not self.df_filtered.empty:
                max_qty = self.df_filtered['Qty'].max()
                min_qty = self.df_filtered['Qty'].min()
                range_qty = max_qty - min_qty if max_qty != min_qty else 1
            
            for row_idx, (_, row) in enumerate(self.df_filtered.iterrows()):
                item_artigo = QTableWidgetItem(str(row['Artigo']))
                item_artigo.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(row_idx, 0, item_artigo)
                
                descricao = str(row['Descrição']) if pd.notna(row['Descrição']) else "N/A"
                item_desc = QTableWidgetItem(descricao)
                item_desc.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(row_idx, 1, item_desc)
                
                status = str(row['Status']) if 'Status' in row and pd.notna(row['Status']) else "N/A"
                item_status = QTableWidgetItem(status)
                item_status.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 2, item_status)
                
                categoria = str(row['Cat.']) if 'Cat.' in row and pd.notna(row['Cat.']) else "N/A"
                item_cat = QTableWidgetItem(categoria)
                item_cat.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 3, item_cat)
                
                subcat = str(row['Sub-C.']) if 'Sub-C.' in row and pd.notna(row['Sub-C.']) else "N/A"
                item_subcat = QTableWidgetItem(subcat)
                item_subcat.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 4, item_subcat)
                
                qty_value = row['Qty'] if pd.notna(row['Qty']) else 0
                item_qty = QTableWidgetItem(f"{qty_value:,.0f}")
                item_qty.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                if not self.df_filtered.empty and range_qty > 0:
                    normalized_value = (qty_value - min_qty) / range_qty
                    if normalized_value > 0.5:
                        green = 255
                        red = int(255 * (1 - (normalized_value - 0.5) * 2))
                    else:
                        red = 255
                        green = int(255 * (normalized_value * 2))
                    
                    blue = 50
                    item_qty.setBackground(QColor(red, green, blue))
                    item_qty.setForeground(QColor(0, 0, 0))
                
                self.table.setItem(row_idx, 5, item_qty)
                
                stock_value = row['Stock'] if pd.notna(row['Stock']) else 0
                item_stock = QTableWidgetItem(f"{stock_value:,.0f}")
                item_stock.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 6, item_stock)
                
                percentual = row.get('%Stock/Vendas', 0) if pd.notna(row.get('%Stock/Vendas', 0)) else 0
                if percentual == 99999:
                    percent_text = "N/A"
                else:
                    percent_text = f"{percentual:.1f}%"
                item_percent = QTableWidgetItem(percent_text)
                item_percent.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                if percentual != 99999:
                    if percentual > 200:
                        item_percent.setBackground(QColor(255, 200, 200))
                    elif percentual < 50:
                        item_percent.setBackground(QColor(255, 255, 200))
                    else:
                        item_percent.setBackground(QColor(200, 255, 200))
                
                self.table.setItem(row_idx, 7, item_percent)
                
                dias_stock = row.get('Dias Stock', 0) if pd.notna(row.get('Dias Stock', 0)) else 0
                item_dias = QTableWidgetItem(f"{dias_stock:.1f}")
                item_dias.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 8, item_dias)
                
                total_venda = row['Total P/Venda'] if 'Total P/Venda' in row and pd.notna(row['Total P/Venda']) else 0
                item_venda = QTableWidgetItem(f"€ {total_venda:,.2f}")
                item_venda.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 9, item_venda)
                
                margem = row['Margem %'] if 'Margem %' in row and pd.notna(row['Margem %']) else 0
                item_margem = QTableWidgetItem(f"{margem:.1f}%")
                item_margem.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 10, item_margem)
                
                vendas_descontos = row['Vendas -Descontos'] if 'Vendas -Descontos' in row and pd.notna(row['Vendas -Descontos']) else 0
                item_vendas_desc = QTableWidgetItem(f"€ {vendas_descontos:,.2f}")
                item_vendas_desc.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 11, item_vendas_desc)
            
            header = self.table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.Stretch)
            header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(8, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(9, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(10, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(11, QHeaderView.ResizeToContents)
            
            self.label_contador.setText(f"Total de artigos: {len(self.df_filtered):,}")
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar tabela: {str(e)}")

    def exportar_pdf(self):
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Aviso", "Não existem dados para exportar.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Exportar para PDF", "VendasVsStocks.pdf", "PDF (*.pdf)"
        )
        if not file_path:
            return

        try:
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

            title_fmt = QTextCharFormat()
            title_fmt.setFont(QFont("Arial", 16, QFont.Bold))
            block_fmt = QTextBlockFormat()
            block_fmt.setAlignment(Qt.AlignCenter)
            cursor.insertBlock(block_fmt)
            cursor.setCharFormat(title_fmt)
            cursor.insertText("ANÁLISE VENDAS VS STOCKS\n\n")

            info = f"Secção: {self.combo_seccao.currentText()} | " \
                   f"Total artigos: {len(self.df_filtered):,} | " \
                   f"Gerado em: {pd.Timestamp.now():%d/%m/%Y %H:%M}\n\n"
            
            cursor.insertText(info)

            headers = [
                'Artigo', 'Descrição', 'Status', 'Cat.', 'Sub-C.', 'Qty', 
                'Stock', '%S/V', 'Dias', 'Total Venda', 'Margem%', 'Vendas-Desc'
            ]

            larguras_percentagem = [
                8, 25, 6, 6, 7, 7, 7, 6, 5, 8, 6, 9
            ]

            table_fmt = QTextTableFormat()
            table_fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100))
            table_fmt.setCellPadding(4)
            table_fmt.setCellSpacing(0)
            table_fmt.setBorder(0.5)
            table_fmt.setBorderStyle(QTextFrameFormat.BorderStyle_Solid)

            constraints = [QTextLength(QTextLength.PercentageLength, w) for w in larguras_percentagem]
            table_fmt.setColumnWidthConstraints(constraints)

            table = cursor.insertTable(len(self.df_filtered) + 1, len(headers), table_fmt)

            header_cell_fmt = QTextTableCellFormat()
            header_cell_fmt.setBackground(QColor("#d0d0d0"))

            header_char_fmt = QTextCharFormat()
            header_char_fmt.setFontWeight(QFont.Bold)
            header_char_fmt.setFontPointSize(8)

            for col, texto in enumerate(headers):
                cell = table.cellAt(0, col)
                cell.setFormat(header_cell_fmt)
                cur = cell.firstCursorPosition()
                cur.insertText(texto, header_char_fmt)

            normal_fmt = QTextCharFormat()
            normal_fmt.setFontPointSize(7)

            for row_idx, (_, row) in enumerate(self.df_filtered.iterrows(), start=1):
                for col_idx, col_name in enumerate(headers):
                    cell = table.cellAt(row_idx, col_idx)
                    cur = cell.firstCursorPosition()

                    col_mapping = {
                        'Artigo': 'Artigo',
                        'Descrição': 'Descrição', 
                        'Status': 'Status',
                        'Cat.': 'Cat.',
                        'Sub-C.': 'Sub-C.',
                        'Qty': 'Qty',
                        'Stock': 'Stock',
                        '%S/V': '%Stock/Vendas',
                        'Dias': 'Dias Stock',
                        'Total Venda': 'Total P/Venda',
                        'Margem%': 'Margem %',
                        'Vendas-Desc': 'Vendas -Descontos'
                    }
                    
                    coluna_real = col_mapping[col_name]
                    value = row.get(coluna_real, '')
                    
                    if pd.isna(value):
                        text = "N/A"
                    else:
                        if coluna_real == "Descrição":
                            desc = str(value)
                            text = desc if len(desc) <= 35 else desc[:32] + "..."
                        elif coluna_real in ["Qty", "Stock"]:
                            text = f"{int(value):,}" if value else "0"
                        elif coluna_real in ["Total P/Venda", "Vendas -Descontos"]:
                            text = f"€{float(value):,.0f}" if value else "€0"
                        elif coluna_real == "%Stock/Vendas":
                            text = "N/A" if value == 99999 else f"{value:.1f}%"
                        elif coluna_real == "Dias Stock":
                            text = f"{value:.1f}" if value else "0.0"
                        elif coluna_real == "Margem %":
                            text = f"{value:.1f}%" if value else "0.0%"
                        else:
                            text = str(value)

                    cur.insertText(text, normal_fmt)

            cursor.movePosition(QTextCursor.End)
            cursor.insertBlock()
            footer = QTextCharFormat()
            footer.setFontPointSize(6)
            footer.setFontItalic(True)
            footer.setForeground(QColor("gray"))
            cursor.setCharFormat(footer)
            cursor.insertText(f"Documento gerado automaticamente • {len(self.df_filtered):,} artigos")

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
                "vendas_vs_stocks_export.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if file_path:
                self.progress_bar.setVisible(True)
                self.progress_bar.setValue(50)
                
                colunas_export = [
                    'Artigo', 'Descrição', 'Status', 'Cat.', 'Sub-C.', 'Qty', 
                    'Stock', '%Stock/Vendas', 'Dias Stock', 'Total P/Venda', 
                    'Margem %', 'Vendas -Descontos'
                ]
                
                colunas_disponiveis = [col for col in colunas_export if col in self.df_filtered.columns]
                df_export = self.df_filtered[colunas_disponiveis].copy()
                
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Vendas vs Stocks')
                    
                    worksheet = writer.sheets['Vendas vs Stocks']
                    
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
        self.coluna_uneg = None
        self.table.setRowCount(0)
        self.label_file.setText("Nenhum ficheiro carregado")

        self.combo_seccao.clear()
        self.combo_seccao.addItem("Todas as Secções")

        self.check_mostrar_todos.setChecked(False)
        self.label_contador.setText("Total de artigos: 0")
        self.btn_exportar_excel.setEnabled(False)
        self.btn_exportar_pdf.setEnabled(False)

def mostrar_vendas_stocks():
    dialog = VendasVsStocksDialog()
    dialog.exec_()
            