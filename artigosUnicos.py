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
from datetime import datetime, timedelta
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog

class ArtigosUnicosDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Artigos Únicos - Comparação de Ficheiros")
        self.setGeometry(100, 100, 1400, 800)
        self.df_unicos = None
        self.df_filtered = None
        self.df_principal = None  # Já tens esta
        self.df_comparacao = None  # Já tens esta
        self.ordenacao_atual = 'Unit Sales'  # Ordenação padrão
        self.ordem_decrescente = True  # Ordem padrão decrescente
        
        # VARIÁVEIS NOVAS PARA ARMAZENAR NOMES DOS FICHEIROS
        self.ficheiro_principal_nome = ""
        self.ficheiro_comparacao_nome = ""
        
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        # Título
        title = QLabel("Artigos com Não Vendas")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("margin: 20px;")
        layout.addWidget(title)
        
        # Descrição
        descricao = QLabel("Compara 2 ficheiros pelo SKU e mostra apenas os artigos únicos do primeiro ficheiro")
        descricao.setFont(QFont("Arial", 10))
        descricao.setAlignment(Qt.AlignCenter)
        descricao.setStyleSheet("color: #666; margin-bottom: 20px;")
        layout.addWidget(descricao)
        
        # Área de upload dos 2 ficheiros
        upload_layout1 = QHBoxLayout()
        self.btn_file1 = QPushButton("📁 Carregar Ficheiro Principal")
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
        self.btn_file2 = QPushButton("📁 Carregar Ficheiro de Comparação")
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
        self.combo_ordenacao.addItems(["Unit Sales", "Sales Value", "Stock", "%Vendas"])
        self.combo_ordenacao.currentTextChanged.connect(self.alterar_ordenacao)
        filters_layout.addWidget(self.combo_ordenacao)

        self.btn_ordem = QPushButton("🔽")
        self.btn_ordem.setToolTip("Alternar entre ordem crescente/decrescente")
        self.btn_ordem.clicked.connect(self.alternar_ordem)
        self.btn_ordem.setFixedSize(30, 30)
        filters_layout.addWidget(self.btn_ordem)

        filters_layout.addStretch()

        self.label_contador = QLabel("Total de artigos únicos: 0")
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

        # Grupo de botões PDF (Exportar e Imprimir)
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

        # BOTÃO NOVO - Imprimir PDF
        self.btn_imprimir_pdf = QPushButton("🖨️ Imprimir PDF")
        self.btn_imprimir_pdf.setFont(QFont("Arial", 12))
        self.btn_imprimir_pdf.setMinimumHeight(40)
        self.btn_imprimir_pdf.setStyleSheet("""
            QPushButton {
                background-color: #9C27B0;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #7B1FA2;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.btn_imprimir_pdf.clicked.connect(self.imprimir_pdf_direto)
        self.btn_imprimir_pdf.setEnabled(False)
        buttons_layout.addWidget(self.btn_imprimir_pdf)

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
        self.df_principal = None
        self.df_comparacao = None

    def carregar_ficheiro(self, tipo):
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, 
                f"Selecionar Ficheiro - {'Principal' if tipo == 1 else 'Comparação'}", 
                "", 
                "Ficheiros Suportados (*.xlsx *.xls *.csv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv)"
            )
            
            if file_path:
                # ARMAZENAR O NOME DO FICHEIRO
                file_name = os.path.basename(file_path)
                
                if tipo == 1:
                    self.ficheiro_principal_nome = file_name
                else:
                    self.ficheiro_comparacao_nome = file_name
                
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
                
                # Validação diferente para cada tipo de ficheiro
                if tipo == 1:
                    # Ficheiro Principal: precisa de todas as colunas
                    colunas_necessarias = ['Sku', 'Description', 'Unit Sales', 'Sales Value', 'Stock', 'Merc.Struct Code']
                    colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
                    
                    if colunas_faltantes:
                        QMessageBox.critical(
                            self, 
                            "Erro", 
                            f"Colunas faltantes no ficheiro Principal: {', '.join(colunas_faltantes)}\n\nColunas encontradas: {', '.join(df.columns)}"
                        )
                        self.progress_bar.setVisible(False)
                        return
                    
                    # Extrair secção do Merc.Struct Code
                    df['Secção'] = df['Merc.Struct Code'].astype(str).str[2:4]
                    
                    # Calcular %Vendas
                    self.calcular_percentual_vendas(df)
                    
                    # Armazenar como df_principal (que já usas no código)
                    self.df_principal = df
                    self.label_file1.setText(file_name)
                    
                else:
                    # Ficheiro de Comparação: apenas precisa do Artigo
                    if 'Artigo' not in df.columns:
                        QMessageBox.critical(
                            self, 
                            "Erro", 
                            f"Coluna 'Artigo' não encontrada no ficheiro de Comparação.\n\nColunas encontradas: {', '.join(df.columns)}"
                        )
                        self.progress_bar.setVisible(False)
                        return
                    
                    # Armazenar como df_comparacao (que já usas no código)
                    self.df_comparacao = df
                    self.label_file2.setText(file_name)
                
                self.progress_bar.setValue(100)
                
                # Verificar se ambos os ficheiros foram carregados
                if self.df_principal is not None and self.df_comparacao is not None:
                    self.encontrar_artigos_unicos()
                
                QMessageBox.information(
                    self, 
                    "Sucesso", 
                    f"Ficheiro {'Principal' if tipo == 1 else 'de Comparação'} carregado com sucesso!\n"
                    f"Nome: {file_name}\n"
                    f"{len(df)} artigos encontrados."
                )
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar ficheiro: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def calcular_percentual_vendas(self, df):
        """Calcula a percentagem de vendas (Stock em relação a Unit Sales)"""
        try:
            # Inicializar a coluna %Vendas
            df['%Vendas'] = 0
            
            # %Vendas = (Stock / Unit Sales) * 100
            # Proteger contra divisão por zero
            mask = (df['Unit Sales'] > 0) & (df['Stock'].notna())
            
            # Calcular %Vendas apenas onde Unit Sales > 0
            df.loc[mask, '%Vendas'] = (df.loc[mask, 'Stock'] / df.loc[mask, 'Unit Sales']) * 100
            
            # Para casos onde Unit Sales é 0, definir %Vendas como um valor muito alto (99999)
            df.loc[df['Unit Sales'] == 0, '%Vendas'] = 99999
            
            # Para casos onde Stock é 0 mas Unit Sales > 0, %Vendas = 0
            df.loc[(df['Stock'] == 0) & (df['Unit Sales'] > 0), '%Vendas'] = 0
            
            # Arredondar para 2 casas decimais
            df['%Vendas'] = df['%Vendas'].round(2)
            
        except Exception as e:
            print(f"Erro ao calcular %Vendas: {e}")
            df['%Vendas'] = 0

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
        
        try:
            return pd.read_csv(file_path, delimiter=',', encoding='utf-8')
        except:
            try:
                return pd.read_csv(file_path, delimiter=';', encoding='latin-1')
            except Exception as e:
                raise Exception(f"Não foi possível ler o ficheiro CSV: {str(e)}")

    def encontrar_artigos_unicos(self):
        """Encontra os artigos únicos do primeiro ficheiro com Stock > 0"""
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # Obter SKUs de ambos os ficheiros
            skus_principal = set(self.df_principal['Sku'].unique())
            skus_comparacao = set(self.df_comparacao['Artigo'].unique())
            
            self.progress_bar.setValue(30)
            
            # Encontrar SKUs únicos no primeiro ficheiro
            skus_unicos = skus_principal - skus_comparacao
            
            self.progress_bar.setValue(60)
            
            # Filtrar o DataFrame principal para mostrar apenas os SKUs únicos COM STOCK > 0
            self.df_unicos = self.df_principal[
                (self.df_principal['Sku'].isin(skus_unicos)) & 
                (self.df_principal['Stock'] > 0)
            ].copy()
            
            # Formatar datas para remover a hora (se a coluna existir)
            if 'Ultima Recepcao' in self.df_unicos.columns:
                # Converter para string e extrair apenas a parte da data
                self.df_unicos['Ultima Recepcao'] = self.df_unicos['Ultima Recepcao'].apply(
                    lambda x: str(x).split()[0] if pd.notna(x) and ' ' in str(x) else str(x) if pd.notna(x) else 'N/A'
                )
            
            # Ordenar por Secção (crescente) e depois por Unit Sales (decrescente)
            self.df_unicos = self.df_unicos.sort_values(['Secção', 'Unit Sales'], ascending=[True, False])
            
            self.progress_bar.setValue(80)
            
            # Preencher combobox com secções únicas
            seccoes = sorted(self.df_unicos['Secção'].unique())
            self.combo_seccao.clear()
            self.combo_seccao.addItem("Todas as Secções")
            self.combo_seccao.addItems([str(sec) for sec in seccoes if pd.notna(sec)])
            
            self.progress_bar.setValue(100)
            
            # Atualizar interface
            self.btn_exportar_excel.setEnabled(True)
            self.btn_exportar_pdf.setEnabled(True)
            self.btn_imprimir_pdf.setEnabled(True)  # Adiciona esta linha
            self.filtrar_por_seccao()
            
            QMessageBox.information(
                self, 
                "Sucesso", 
                f"Análise de artigos únicos concluída!\n"
                f"→ {len(skus_unicos)} artigos únicos encontrados no ficheiro principal\n"
                f"→ {len(self.df_unicos)} artigos únicos com stock > 0\n"
                f"→ {len(skus_principal)} artigos no ficheiro principal\n"
                f"→ {len(skus_comparacao)} artigos no ficheiro de comparação"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao encontrar artigos únicos: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def filtrar_por_seccao(self):
        if self.df_unicos is None:
            return
        
        try:
            seccao_selecionada = self.combo_seccao.currentText()
            mostrar_todos = self.check_mostrar_todos.isChecked()
            
            if seccao_selecionada == "Todas as Secções":
                self.df_filtered = self.df_unicos.copy()
            else:
                self.df_filtered = self.df_unicos[self.df_unicos['Secção'] == seccao_selecionada].copy()
            
            # O filtro Stock > 0 já foi aplicado no df_unicos, então não precisa repetir aqui
            # Se não mostrar todos, limitar aos top artigos
            if not mostrar_todos:
                self.df_filtered = self.df_filtered.head(100)
            
            self.atualizar_tabela()
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao filtrar: {str(e)}")

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

    def alternar_ordem(self):
        """Alterna entre ordem crescente e decrescente"""
        self.ordem_decrescente = not self.ordem_decrescente
        self.btn_ordem.setText("🔽" if self.ordem_decrescente else "🔼")
        self.aplicar_ordenacao()

    def aplicar_ordenacao(self):
        """Aplica a ordenação atual aos dados"""
        if self.df_unicos is None:
            return
        
        try:
            # Para %Vendas, usar sempre ordenação crescente
            if self.ordenacao_atual == '%Vendas':
                self.df_unicos = self.df_unicos.sort_values(
                    [self.ordenacao_atual],  # Apenas pela coluna selecionada
                    ascending=True,  # Sempre crescente para %Vendas
                    na_position='last'
                )
            else:
                # Para outras colunas, ordenar por Secção primeiro e depois pela coluna selecionada
                self.df_unicos = self.df_unicos.sort_values(
                    ['Secção', self.ordenacao_atual],  # Secção sempre primeiro
                    ascending=[True, not self.ordem_decrescente]  # Secção crescente, coluna selecionada conforme botão
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
            # Configurar tabela com 10 colunas (mantendo 10 colunas no total)
            self.table.setRowCount(len(self.df_filtered))
            self.table.setColumnCount(10)
            self.table.setHorizontalHeaderLabels([
                'Sku', 'Description', 'PVP Permanente', 'PVPR', 'Unit Sales', 'Sales Value', 'Stock', 
                '%Vendas', 'Ultima Recepcao', 'Flow-type'
            ])
            
            # Calcular valores para o gradiente de cores da Unit Sales
            if not self.df_filtered.empty:
                max_unit_sales = self.df_filtered['Unit Sales'].max()
                min_unit_sales = self.df_filtered['Unit Sales'].min()
                range_unit_sales = max_unit_sales - min_unit_sales if max_unit_sales != min_unit_sales else 1
            
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
                
                # PVP Permanente (substitui o EAN)
                # Assumindo que a coluna existe no DataFrame, ajuste conforme necessário
                pvp_permanente = str(row['PVP Permanente']) if 'PVP Permanente' in row and pd.notna(row['PVP Permanente']) else "N/A"
                item_pvp_permanente = QTableWidgetItem(pvp_permanente)
                item_pvp_permanente.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 2, item_pvp_permanente)
                
                # PVPR
                # Assumindo que a coluna existe no DataFrame, ajuste conforme necessário
                pvpr = str(row['PVPR']) if 'PVPR' in row and pd.notna(row['PVPR']) else "N/A"
                item_pvpr = QTableWidgetItem(pvpr)
                item_pvpr.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 3, item_pvpr)
                
                # Unit Sales com gradiente de cores
                unit_sales_value = row['Unit Sales'] if pd.notna(row['Unit Sales']) else 0
                item_unit_sales = QTableWidgetItem(f"{unit_sales_value:,.0f}")
                item_unit_sales.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                # Aplicar gradiente de cor (verde para alto, vermelho para baixo)
                if not self.df_filtered.empty and range_unit_sales > 0:
                    normalized_value = (unit_sales_value - min_unit_sales) / range_unit_sales
                    # Verde (alto) -> Amarelo (médio) -> Vermelho (baixo)
                    if normalized_value > 0.5:
                        # Verde para amarelo
                        green = 255
                        red = int(255 * (1 - (normalized_value - 0.5) * 2))
                    else:
                        # Amarelo para vermelho
                        red = 255
                        green = int(255 * (normalized_value * 2))
                    
                    blue = 50  # Azul baixo para melhor contraste
                    item_unit_sales.setBackground(QColor(red, green, blue))
                    item_unit_sales.setForeground(QColor(0, 0, 0))  # Texto preto para contraste
                
                self.table.setItem(row_idx, 4, item_unit_sales)  # Agora na coluna 4
                
                # Sales Value
                sales_value = row['Sales Value'] if pd.notna(row['Sales Value']) else 0
                item_sales_value = QTableWidgetItem(f"€ {sales_value:,.2f}")
                item_sales_value.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 5, item_sales_value)  # Agora na coluna 5
                
                # Stock
                stock_value = row['Stock'] if pd.notna(row['Stock']) else 0
                item_stock = QTableWidgetItem(f"{stock_value:,.0f}")
                item_stock.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 6, item_stock)  # Agora na coluna 6
                
                # %Vendas
                percentual = row.get('%Vendas', 0) if pd.notna(row.get('%Vendas', 0)) else 0
                if percentual == 99999:  # Valor que usamos para Unit Sales = 0
                    percent_text = "N/A"
                else:
                    percent_text = f"{percentual:.1f}%"
                item_percent = QTableWidgetItem(percent_text)
                item_percent.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, 7, item_percent)  # Agora na coluna 7
                
                # Ultima Recepcao - já formatada sem hora
                ultima_recepcao = str(row.get('Ultima Recepcao', 'N/A')) if 'Ultima Recepcao' in row and pd.notna(row.get('Ultima Recepcao')) else "N/A"
                item_recepcao = QTableWidgetItem(ultima_recepcao)
                item_recepcao.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 8, item_recepcao)  # Agora na coluna 8
                
                # Flow-type
                flow_type = str(row.get('Flow-type', 'N/A')) if 'Flow-type' in row and pd.notna(row.get('Flow-type')) else "N/A"
                item_flow = QTableWidgetItem(flow_type)
                item_flow.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self.table.setItem(row_idx, 9, item_flow)  # Agora na coluna 9
                
                # Status foi removido para manter 10 colunas
            
            # Ajustar tamanho das colunas
            header = self.table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Sku
            header.setSectionResizeMode(1, QHeaderView.Stretch)          # Description
            header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # PVP Permanente
            header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # PVPR
            header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Unit Sales
            header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Sales Value
            header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # Stock
            header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # %Vendas
            header.setSectionResizeMode(8, QHeaderView.ResizeToContents)  # Ultima Recepcao
            header.setSectionResizeMode(9, QHeaderView.ResizeToContents)  # Flow-type
            
            # Atualizar contador
            self.label_contador.setText(f"Total de artigos únicos: {len(self.df_filtered):,}")
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar tabela: {str(e)}")

    def exportar_pdf(self):
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Aviso", "Não existem dados para exportar.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Exportar para PDF", "ArtigosUnicos.pdf", "PDF (*.pdf)"
        )
        if not file_path:
            return

        try:
            # Configuração PDF (A4 Landscape)
            printer = QPrinter(QPrinter.HighResolution)            
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

            # Calcular a data de ontem
            data_ontem = datetime.now() - timedelta(days=1)
            data_ontem_formatada = data_ontem.strftime("%d-%m-%Y")

            cursor.insertText(f"Não Vendas - {data_ontem_formatada}\n\n")                        

            info = f"Secção: {self.combo_seccao.currentText()} | " \
                f"Total artigos únicos: {len(self.df_filtered):,} | " \
                f"Gerado em: {pd.Timestamp.now():%d/%m/%Y %H:%M}\n\n"
            cursor.insertText(info)

            # Cabeçalhos atualizados (sem EAN, com PVP Permanente e PVPR)
            headers = [
                'Sku', 'Description', 'PVP Perm.', 'PVPR', 'Unit Sales', 
                'Sales Value', 'Stock', '%Vendas', 'Ultima Recepcao', 'Flow-type'
            ]

            # Larguras ajustadas para as novas colunas
            larguras_percentagem = [
                8,   # Sku
                24,  # Description (reduzida de 26 para 24)
                9,   # PVP Perm. (nova coluna)
                8,   # PVPR (nova coluna)
                8,   # Unit Sales (reduzida de 9 para 8)
                9,   # Sales Value
                6,   # Stock
                7,   # %Vendas (reduzida de 8 para 7)
                10,  # Ultima Recepcao
                9    # Flow-type (reduzida de 9 para 9)
            ]  # soma = 100% (8+24+9+8+8+9+6+7+10+9 = 98, ajustar se necessário)

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

            # Dados
            normal_fmt = QTextCharFormat()
            normal_fmt.setFontPointSize(8)

            bold_fmt = QTextCharFormat(normal_fmt)
            bold_fmt.setFontWeight(QFont.Bold)
            bold_fmt.setFontPointSize(9)

            for row_idx, (_, row) in enumerate(self.df_filtered.iterrows(), start=1):
                for col_idx, col_name in enumerate(headers):
                    cell = table.cellAt(row_idx, col_idx)
                    cur = cell.firstCursorPosition()

                    # Mapeamento dos nomes das colunas para os nomes no DataFrame
                    if col_name == 'PVP Perm.':
                        value = row.get('PVP Permanente', '')
                    elif col_name == 'PVPR':
                        value = row.get('PVPR', '')
                    else:
                        value = row.get(col_name, '')

                    if pd.isna(value):
                        text = "N/A"
                    else:
                        if col_name == "Description":
                            desc = str(value)
                            text = desc if len(desc) <= 45 else desc[:42] + "..."
                        elif col_name in ["Unit Sales", "Stock"]:
                            text = f"{int(value):,}" if value else "0"
                        elif col_name in ["Sales Value", "PVP Perm.", "PVPR"]:
                            # Formatar como moeda para colunas de preço
                            try:
                                if value:
                                    text = f"€{float(value):,.2f}"
                                else:
                                    text = "€0.00"
                            except:
                                text = str(value)
                        elif col_name == "%Vendas":
                            text = "N/A" if value == 99999 else f"{value:.1f}%"
                        elif col_name == "Ultima Recepcao":
                            text = str(value)[:10] if pd.notna(value) else "N/A"
                        else:
                            text = str(value)

                    if col_name == "Unit Sales" or col_name == "Stock":
                        cur.insertText(text, bold_fmt)
                    else:
                        cur.insertText(text, normal_fmt)

            # Rodapé
            cursor.movePosition(QTextCursor.End)
            cursor.insertBlock()
            footer = QTextCharFormat()
            footer.setFontPointSize(7)
            footer.setFontItalic(True)
            footer.setForeground(QColor("gray"))
            cursor.setCharFormat(footer)
            cursor.insertText(f"Artigos únicos do ficheiro principal com stock > 0 • {len(self.df_filtered):,} artigos")

            # Exportar
            doc.print_(printer)

            QMessageBox.information(
                self, "Sucesso",
                f"PDF exportado com sucesso!\n\n"
                f"→ {len(self.df_filtered):,} artigos únicos exportados\n"
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
                "artigos_unicos_export.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if file_path:
                self.progress_bar.setVisible(True)
                self.progress_bar.setValue(50)
                
                # Criar DataFrame para exportação com todas as colunas
                colunas_export = ['Sku', 'Description', 'EAN', 'Unit Sales', 'Sales Value', 'Stock', 
                                '%Vendas', 'Ultima Recepcao', 'Flow-type', 'Status', 'Secção']
                
                # Filtrar apenas colunas que existem no DataFrame
                colunas_disponiveis = [col for col in colunas_export if col in self.df_filtered.columns]
                df_export = self.df_filtered[colunas_disponiveis].copy()
                
                # Exportar para Excel
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Artigos Únicos')
                    
                    # Acessar a worksheet para ajustar as colunas
                    worksheet = writer.sheets['Artigos Únicos']
                    
                    # Ajustar largura das colunas baseado no conteúdo
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        
                        # Encontrar o comprimento máximo na coluna
                        for cell in column:
                            try:
                                if cell.value:
                                    cell_length = len(str(cell.value))
                                    max_length = max(max_length, cell_length)
                            except:
                                pass
                        
                        # Ajustar largura (com margem de segurança)
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                    
                self.progress_bar.setValue(100)
                
                QMessageBox.information(
                    self, 
                    "Sucesso", 
                    f"Dados exportados com sucesso!\n{len(df_export)} artigos únicos exportados."
                )
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def imprimir_pdf_direto(self):
        """
        Imprime diretamente o PDF sem pedir para salvar primeiro
        """
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Aviso", "Não existem dados para imprimir.")
            return

        try:
            # Diálogo de impressora
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            
            # Configurar impressora (formato vertical padrão)
            printer_dialog = QPrintDialog(printer, self)
            if printer_dialog.exec() != QPrintDialog.Accepted:
                return  # Usuário cancelou

            # Criar documento PDF
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

            # Calcular a data de ontem
            from datetime import datetime, timedelta
            data_ontem = datetime.now() - timedelta(days=1)
            data_ontem_formatada = data_ontem.strftime("%d-%m-%Y")

            # Título com nome do ficheiro principal (se disponível)
            titulo = f"Não Vendas - {data_ontem_formatada}"
            if hasattr(self, 'ficheiro_principal_nome') and self.ficheiro_principal_nome:
                # Remove a extensão do arquivo para ficar mais limpo
                nome_sem_extensao = os.path.splitext(self.ficheiro_principal_nome)[0]
                titulo = f"Não Vendas - {nome_sem_extensao} - {data_ontem_formatada}"
            
            cursor.insertText(f"{titulo}\n\n")                        

            # Informações adicionais
            info_parts = []
            info_parts.append(f"Secção: {self.combo_seccao.currentText()}")
            info_parts.append(f"Total artigos únicos: {len(self.df_filtered):,}")
            
            # Adicionar nome do ficheiro principal se disponível
            #if hasattr(self, 'ficheiro_principal_nome') and self.ficheiro_principal_nome:
             #   info_parts.append(f"Ficheiro: {self.ficheiro_principal_nome}")
            
            info_parts.append(f"Gerado em: {pd.Timestamp.now():%d/%m/%Y %H:%M}")
            
            info = " | ".join(info_parts) + "\n\n"
            cursor.insertText(info)

            # Cabeçalhos atualizados
            headers = [
                'Sku', 'Description', 'PVP Perm.', 'PVPR', 'Unit Sales', 
                'Sales Value', 'Stock', '%Vendas', 'Ultima Recepcao', 'Flow-type'
            ]

            # Larguras ajustadas para formato vertical (mais largas)
            larguras_percentagem = [
                8,   # Sku
                26,  # Description
                9,   # PVP Perm.
                8,   # PVPR
                8,   # Unit Sales
                9,   # Sales Value
                6,   # Stock
                7,   # %Vendas
                10,  # Ultima Recepcao
                9    # Flow-type
            ]

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

            # Dados
            normal_fmt = QTextCharFormat()
            normal_fmt.setFontPointSize(8)

            bold_fmt = QTextCharFormat(normal_fmt)
            bold_fmt.setFontWeight(QFont.Bold)
            bold_fmt.setFontPointSize(9)

            for row_idx, (_, row) in enumerate(self.df_filtered.iterrows(), start=1):
                for col_idx, col_name in enumerate(headers):
                    cell = table.cellAt(row_idx, col_idx)
                    cur = cell.firstCursorPosition()

                    # Mapeamento dos nomes das colunas
                    if col_name == 'PVP Perm.':
                        value = row.get('PVP Permanente', '')
                    elif col_name == 'PVPR':
                        value = row.get('PVPR', '')
                    else:
                        value = row.get(col_name, '')

                    if pd.isna(value):
                        text = "N/A"
                    else:
                        if col_name == "Description":
                            desc = str(value)
                            text = desc if len(desc) <= 45 else desc[:42] + "..."
                        elif col_name in ["Unit Sales", "Stock"]:
                            text = f"{int(value):,}" if value else "0"
                        elif col_name in ["Sales Value", "PVP Perm.", "PVPR"]:
                            try:
                                if value:
                                    text = f"€{float(value):,.2f}"
                                else:
                                    text = "€0.00"
                            except:
                                text = str(value)
                        elif col_name == "%Vendas":
                            text = "N/A" if value == 99999 else f"{value:.1f}%"
                        elif col_name == "Ultima Recepcao":
                            text = str(value)[:10] if pd.notna(value) else "N/A"
                        else:
                            text = str(value)

                    # Aplicar bold tanto para Unit Sales como para Stock
                    if col_name in ["Unit Sales", "Stock"]:
                        cur.insertText(text, bold_fmt)
                    else:
                        cur.insertText(text, normal_fmt)

            # Rodapé
            cursor.movePosition(QTextCursor.End)
            cursor.insertBlock()
            footer = QTextCharFormat()
            footer.setFontPointSize(7)
            footer.setFontItalic(True)
            footer.setForeground(QColor("gray"))
            cursor.setCharFormat(footer)
            
            # Rodapé com mais informações
            rodape_parts = [f"Artigos únicos do ficheiro principal com stock > 0"]
            rodape_parts.append(f"{len(self.df_filtered):,} artigos")
            
            if hasattr(self, 'ficheiro_principal_nome') and self.ficheiro_principal_nome:
                rodape_parts.append(f"Fonte: {self.ficheiro_principal_nome}")
            
            if hasattr(self, 'ficheiro_comparacao_nome') and self.ficheiro_comparacao_nome:
                rodape_parts.append(f"vs {self.ficheiro_comparacao_nome}")
            
            rodape = " • ".join(rodape_parts)
            cursor.insertText(rodape)

            # Imprimir
            doc.print_(printer)
            
            # Mensagem de sucesso
            msg_sucesso = f"PDF enviado para impressão com sucesso!\n\n"
            msg_sucesso += f"→ {len(self.df_filtered):,} artigos únicos impressos\n"
            msg_sucesso += f"→ Formato: Vertical (A4 Portrait)"
            
            if hasattr(self, 'ficheiro_principal_nome') and self.ficheiro_principal_nome:
                msg_sucesso += f"\n→ Ficheiro: {self.ficheiro_principal_nome}"
            
            QMessageBox.information(self, "Sucesso", msg_sucesso)

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao imprimir PDF:\n{str(e)}")

    def limpar_tudo(self):
        self.df_unicos = None
        self.df_filtered = None
        self.df_principal = None
        self.df_comparacao = None
        
        # LIMPAR NOMES DOS FICHEIROS
        self.ficheiro_principal_nome = ""
        self.ficheiro_comparacao_nome = ""
        
        self.table.setRowCount(0)
        self.label_file1.setText("Nenhum ficheiro carregado")
        self.label_file2.setText("Nenhum ficheiro carregado")
        self.combo_seccao.clear()
        self.combo_seccao.addItem("Todas as Secções")
        self.check_mostrar_todos.setChecked(False)
        self.label_contador.setText("Total de artigos únicos: 0")
        self.btn_exportar_excel.setEnabled(False)
        self.btn_exportar_pdf.setEnabled(False)
        # Se tiveres o botão de imprimir, adiciona também:
        if hasattr(self, 'btn_imprimir_pdf'):
            self.btn_imprimir_pdf.setEnabled(False)


def mostrar_artigos_unicos():
    dialog = ArtigosUnicosDialog()
    dialog.exec_()