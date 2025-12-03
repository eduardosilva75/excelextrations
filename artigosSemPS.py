import os
import sys
import pandas as pd
import numpy as np
import tempfile
from datetime import datetime

# IMPORTS PARA ML
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor, AdaBoostRegressor
from sklearn.tree import DecisionTreeRegressor
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.model_selection import cross_val_score
from sklearn.metrics import mean_absolute_error, r2_score
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QProgressBar, QTableWidget,
    QTableWidgetItem, QHeaderView, QComboBox, QLineEdit,
    QCheckBox, QMainWindow, QWidget, QApplication, QSplitter,
    QTextEdit, QScrollArea
)

from PyQt5.QtPrintSupport import QPrinter, QPrintDialog

from PyQt5.QtGui import (
    QFont, QColor, QTextDocument, QTextCursor, QTextTableFormat,
    QTextTableCellFormat, QTextCharFormat, QTextBlockFormat,
    QTextLength, QPageSize, QPageLayout, QPainter, QDesktopServices
)

from PyQt5.QtCore import Qt, QMarginsF, QUrl, QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QTextFrameFormat


# ============================================
# THREAD PARA PROCESSAMENTO DE ML (evita congelamento)
# ============================================
class MLWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(object, object, object, object, object)
    error = pyqtSignal(str)

    def __init__(self, df_treino, df_prever, numeric_features, categorical_features):
        super().__init__()
        self.df_treino = df_treino
        self.df_prever = df_prever
        self.numeric_features = numeric_features
        self.categorical_features = categorical_features

    def run(self):
        try:
            self.progress.emit(10)
            
            X_train, y_train, X_pred, preprocessor, feature_names = self.preparar_dados_ml()
            
            self.progress.emit(40)
            
            modelo, mae_score, nome_modelo = self.treinar_modelo_ml(X_train, y_train)
            
            self.progress.emit(70)
            
            previsoes = modelo.predict(X_pred)
            
            self.progress.emit(90)
            
            self.finished.emit(modelo, previsoes, mae_score, nome_modelo, feature_names)
            
        except Exception as e:
            self.error.emit(str(e))

    def preparar_dados_ml(self):
        all_features = self.numeric_features + self.categorical_features
        
        X_train = self.df_treino[all_features].copy()
        y_train = self.df_treino['Presentation Stock'].values
        X_pred = self.df_prever[all_features].copy()
        
        numeric_transformer = Pipeline(steps=[
            ('imputer', SimpleImputer(strategy='median')),
            ('scaler', StandardScaler())
        ])
        
        categorical_transformer = Pipeline(steps=[
            ('imputer', SimpleImputer(strategy='constant', fill_value='missing')),
            ('encoder', OneHotEncoder(handle_unknown='ignore', sparse_output=False))
        ])
        
        preprocessor = ColumnTransformer(
            transformers=[
                ('num', numeric_transformer, self.numeric_features),
                ('cat', categorical_transformer, self.categorical_features)
            ],
            remainder='drop'
        )
        
        X_train_processed = preprocessor.fit_transform(X_train)
        X_pred_processed = preprocessor.transform(X_pred)
        
        feature_names = self._get_feature_names(preprocessor)
        
        return X_train_processed, y_train, X_pred_processed, preprocessor, feature_names

    def _get_feature_names(self, preprocessor):
        feature_names = []
        feature_names.extend(self.numeric_features)
        
        if self.categorical_features:
            try:
                cat_encoder = preprocessor.named_transformers_['cat'].named_steps['encoder']
                for i, cat_feature in enumerate(self.categorical_features):
                    categories = cat_encoder.categories_[i]
                    feature_names.extend([f"{cat_feature}_{cat}" for cat in categories])
            except:
                pass
        
        return feature_names

    def treinar_modelo_ml(self, X_train, y_train):
        modelos = {
            'RandomForest': RandomForestRegressor(
                n_estimators=100, 
                max_depth=10,
                min_samples_split=5,
                min_samples_leaf=2,
                random_state=42,
                n_jobs=-1
            ),
            'GradientBoosting': GradientBoostingRegressor(
                n_estimators=100,
                max_depth=6,
                learning_rate=0.1,
                subsample=0.8,
                random_state=42
            ),
            'AdaBoost': AdaBoostRegressor(
                estimator=DecisionTreeRegressor(max_depth=6, min_samples_split=15),
                n_estimators=100,
                learning_rate=0.1,
                random_state=42
            )
        }
        
        melhor_modelo = None
        melhor_score = float('inf')
        melhor_nome = ''
        
        for nome, modelo in modelos.items():
            try:
                cv_folds = min(5, len(X_train) // 2)
                
                scores = cross_val_score(
                    modelo, X_train, y_train, 
                    cv=cv_folds, 
                    scoring='neg_mean_absolute_error',
                    n_jobs=1  # Reduzido para Windows
                )
                
                mae_score = -scores.mean()
                
                if mae_score < melhor_score:
                    melhor_score = mae_score
                    melhor_modelo = modelo
                    melhor_nome = nome
                    
            except Exception:
                continue
        
        if melhor_modelo is not None:
            melhor_modelo.fit(X_train, y_train)
            return melhor_modelo, melhor_score, melhor_nome
        else:
            modelo = RandomForestRegressor(n_estimators=50, random_state=42, n_jobs=1)
            modelo.fit(X_train, y_train)
            return modelo, 0, "RandomForest (Fallback)"


# ============================================
# DIALOG PRINCIPAL
# ============================================
class ArtigosSemPSDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Artigos sem Presentation Stock")
        self.setGeometry(100, 100, 1600, 800)
        self.df = None
        self.df_filtered = None
        self.df_com_ps = None
        self.ml_worker = None  # Referência para o worker
        self.initUI()
        
        # Flag para evitar múltiplos processamentos
        self.processing = False

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
        self.combo_seccao.currentTextChanged.connect(self.aplicar_filtros)
        filters_layout.addWidget(self.combo_seccao)

        filters_layout.addWidget(QLabel("Status:"))

        self.combo_status = QComboBox()
        self.combo_status.setMinimumWidth(150)
        self.combo_status.addItem("Todos os Status")
        self.combo_status.currentTextChanged.connect(self.aplicar_filtros)
        filters_layout.addWidget(self.combo_status)

        # Filtro para Stock
        filters_layout.addWidget(QLabel("Stock:"))

        self.combo_stock = QComboBox()
        self.combo_stock.setMinimumWidth(150)
        self.combo_stock.addItem("Todos")
        self.combo_stock.addItem("Stock > 0")
        self.combo_stock.addItem("Stock = 0")
        self.combo_stock.currentTextChanged.connect(self.aplicar_filtros)
        filters_layout.addWidget(self.combo_stock)

        # Filtro de busca rápida
        filters_layout.addWidget(QLabel("Buscar:"))
        self.search_input = QLineEdit()
        self.search_input.setMinimumWidth(150)
        self.search_input.setPlaceholderText("SKU ou Descrição...")
        self.search_input.textChanged.connect(self.aplicar_filtros)
        filters_layout.addWidget(self.search_input)

        filters_layout.addStretch()

        self.label_contador = QLabel("Total de artigos sem Presentation Stock: 0")
        self.label_contador.setStyleSheet("font-weight: bold;")
        filters_layout.addWidget(self.label_contador)

        layout.addLayout(filters_layout)
        
        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Tabela dentro de um scroll area
        self.table_scroll = QScrollArea()
        self.table_scroll.setWidgetResizable(True)
        self.table_widget = QWidget()
        self.table_layout = QVBoxLayout(self.table_widget)
        
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget {
                gridline-color: #d0d0d0;
                background-color: white;
                border: 1px solid #d0d0d0;
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
            QTableWidget::item:selected {
                background-color: #4CAF50;
                color: white;
            }
        """)
        self.table_layout.addWidget(self.table)
        self.table_scroll.setWidget(self.table_widget)
        layout.addWidget(self.table_scroll)
        
        # Botões de ação
        buttons_layout = QHBoxLayout()

        # Botão ML
        self.btn_ml = QPushButton("🤖 Calcular com ML")
        self.btn_ml.setFont(QFont("Arial", 12))
        self.btn_ml.setMinimumHeight(40)
        self.btn_ml.setStyleSheet("""
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
        self.btn_ml.clicked.connect(self.calcular_sugestao_ps_ml)
        self.btn_ml.setEnabled(False)
        buttons_layout.addWidget(self.btn_ml)

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
        self.btn_exportar_pdf.clicked.connect(self.exportar_pdf_fixed)
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
            stock_selecionado = self.combo_stock.currentText()
            search_text = self.search_input.text().lower()
            
            # Aplicar filtro por secção
            if seccao_selecionada == "Todas as Secções":
                df_temp = self.df_filtered.copy()
            else:
                df_temp = self.df_filtered[self.df_filtered['Secção'] == seccao_selecionada].copy()
            
            # Aplicar filtro por status (se a coluna existir)
            if 'Status' in df_temp.columns and status_selecionado != "Todos os Status":
                df_temp = df_temp[df_temp['Status'] == status_selecionado]
            
            # Aplicar filtro por stock
            if stock_selecionado == "Stock > 0":
                df_temp = df_temp[df_temp['Stock'] > 0]
            elif stock_selecionado == "Stock = 0":
                df_temp = df_temp[df_temp['Stock'] == 0]
            
            # Aplicar filtro de busca
            if search_text:
                mask = (
                    df_temp['Sku'].astype(str).str.lower().str.contains(search_text) |
                    df_temp['Description'].astype(str).str.lower().str.contains(search_text)
                )
                df_temp = df_temp[mask]
            
            self.atualizar_tabela(df_temp)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao aplicar filtros: {str(e)}")

    def calcular_sugestao_ps(self):
        """Método principal - escolhe entre ML ou regras"""
        try:
            if 'Presentation Stock' not in self.df.columns:
                QMessageBox.warning(self, "Aviso", "Coluna 'Presentation Stock' não encontrada.")
                self.df['Sugestão Presentation Stock'] = 0
                return

            df_com_ps = self.df[self.df['Presentation Stock'] > 0].copy()
            
            if len(df_com_ps) >= 20:
                resposta = QMessageBox.question(
                    self,
                    "Escolher Método",
                    f"Encontrados {len(df_com_ps)} artigos com PS > 0.\n\n"
                    "Deseja usar Machine Learning (mais preciso) ou regras manuais?\n\n"
                    "• 🤖 ML: Recomendado para dados suficientes\n"
                    "• 📊 Regras: Mais conservador",
                    QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                    QMessageBox.Yes
                )
                
                if resposta == QMessageBox.Yes:
                    self.calcular_sugestao_ps_ml()
                    return
                elif resposta == QMessageBox.No:
                    self.calcular_sugestao_ps_regras()
                    return
                else:
                    return
            else:
                self.calcular_sugestao_ps_regras()
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao calcular sugestões: {str(e)}")
            self.df['Sugestão Presentation Stock'] = 0

    def calcular_sugestao_ps_ml(self):
        """Calcula sugestão de Presentation Stock usando Machine Learning em thread separada"""
        if self.processing:
            QMessageBox.warning(self, "Aviso", "Já existe um processamento em andamento.")
            return
            
        try:
            self.processing = True
            self.btn_ml.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(5)
            
            df_ml = self.df.copy()
            
            # Features numéricas
            numeric_features = [
                'Sup.Code', 'Sup.Pack Size', 'PVP Permanente', 'Stock', 
                'Stock In Transit', 'Stock Expected', 'Last Order Point', 
                'Lead Time WH to Location', 'Unit Sales', 'Sales Value'
            ]
            
            # Features categóricas
            categorical_features = ['Flow-type', 'Dta activacao', 'GLP']
            
            # Verifica quais features existem no DataFrame
            available_numeric = [f for f in numeric_features if f in df_ml.columns]
            available_categorical = [f for f in categorical_features if f in df_ml.columns]
            
            all_features = available_numeric + available_categorical
            
            if len(all_features) < 3:
                QMessageBox.warning(self, "Aviso", 
                    f"Poucas features disponíveis ({len(all_features)}). Mínimo recomendado: 3")
                self.calcular_sugestao_ps_regras()
                return
            
            # Filtra dados para treino (artigos com PS > 0)
            df_treino = df_ml[df_ml['Presentation Stock'] > 0].copy()
            
            if len(df_treino) < 10:
                QMessageBox.warning(self, "Aviso", 
                    "Poucos artigos com PS > 0 para treinar modelo (mínimo: 10).")
                self.calcular_sugestao_ps_regras()
                return
            
            # Filtra dados para previsão (artigos com PS = 0)
            df_prever = df_ml[df_ml['Presentation Stock'] == 0].copy()
            
            if df_prever.empty:
                QMessageBox.information(self, "Info", "Todos os artigos já têm PS definido.")
                return
            
            self.progress_bar.setValue(15)
            
            # Criar e iniciar worker
            self.ml_worker = MLWorker(df_treino, df_prever, available_numeric, available_categorical)
            self.ml_worker.progress.connect(self.progress_bar.setValue)
            self.ml_worker.finished.connect(self.on_ml_finished)
            self.ml_worker.error.connect(self.on_ml_error)
            self.ml_worker.start()
            
        except Exception as e:
            self.processing = False
            self.btn_ml.setEnabled(True)
            self.progress_bar.setVisible(False)
            QMessageBox.critical(self, "Erro", f"Erro ao iniciar ML: {str(e)}")
            import traceback
            print(traceback.format_exc())

    def on_ml_finished(self, modelo, previsoes, mae_score, nome_modelo, feature_names):
        """Callback quando ML termina"""
        try:
            # Aplica previsões com constraints
            df_prever = self.df[self.df['Presentation Stock'] == 0].copy()
            
            if len(previsoes) != len(df_prever):
                raise ValueError("Número de previsões não corresponde ao número de artigos")
            
            df_prever['Sugestão Presentation Stock'] = np.round(previsoes).astype(int)
            df_prever['Sugestão Presentation Stock'] = df_prever['Sugestão Presentation Stock'].clip(
                lower=1, upper=200
            )
            
            # Atualiza DataFrame principal
            self.df.loc[df_prever.index, 'Sugestão Presentation Stock'] = df_prever['Sugestão Presentation Stock']
            self.df.loc[self.df['Presentation Stock'] > 0, 'Sugestão Presentation Stock'] = 0
            
            # Mostra métricas
            sugestoes = df_prever['Sugestão Presentation Stock']
            msg = f"=== Resultados do Machine Learning ===\n\n"
            msg += f"Modelo utilizado: {nome_modelo}\n"
            msg += f"MAE (Validação Cruzada): {mae_score:.2f}\n\n"
            msg += f"Dados de previsão: {len(df_prever)} artigos\n\n"
            msg += f"Sugestões geradas:\n"
            msg += f"  - Mínimo: {sugestoes.min()}\n"
            msg += f"  - Média: {sugestoes.mean():.1f}\n"
            msg += f"  - Mediana: {sugestoes.median():.1f}\n"
            msg += f"  - Máximo: {sugestoes.max()}\n"
            
            QMessageBox.information(self, "Métricas ML", msg)
            
            # Atualiza visualização
            self.df_filtered = self.df[self.df['Presentation Stock'] == 0].copy()
            self.aplicar_filtros()
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao aplicar resultados ML: {str(e)}")
        finally:
            self.processing = False
            self.btn_ml.setEnabled(True)
            self.progress_bar.setVisible(False)
            self.ml_worker = None

    def on_ml_error(self, error_msg):
        """Callback para erro no ML"""
        self.processing = False
        self.btn_ml.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        QMessageBox.critical(self, "Erro no ML", f"{error_msg}\n\nUsando regras manuais como fallback.")
        self.calcular_sugestao_ps_regras()
        
        self.ml_worker = None

    def calcular_sugestao_ps_regras(self):
        """Método baseado em regras (fallback)"""
        try:
            if 'Presentation Stock' not in self.df.columns:
                return

            self.df_com_ps = self.df[self.df['Presentation Stock'] > 0].copy()
            
            if self.df_com_ps.empty:
                self.df['Sugestão Presentation Stock'] = 3
                return

            stats_por_seccao = self.df_com_ps.groupby('Secção').agg({
                'Presentation Stock': ['median', 'min'],
                'Unit Sales': 'median'
            }).round(2)

            artigos_sem_ps = self.df[self.df['Presentation Stock'] == 0].copy()
            artigos_sem_ps['Sugestão Presentation Stock'] = 0

            for idx, artigo in artigos_sem_ps.iterrows():
                seccao = artigo['Secção']
                unit_sales = float(artigo['Unit Sales']) if pd.notna(artigo['Unit Sales']) else 0.0
                pvp = float(artigo.get('PVP Em Vigor', 0)) if pd.notna(artigo.get('PVP Em Vigor')) else 0.0
                pack_size = int(artigo.get('Sup.Pack Size', 1)) if pd.notna(artigo.get('Sup.Pack Size')) else 1

                if seccao in stats_por_seccao.index:
                    ps_median = stats_por_seccao.loc[seccao, ('Presentation Stock', 'median')]
                    ps_min = stats_por_seccao.loc[seccao, ('Presentation Stock', 'min')]
                    unit_sales_median = stats_por_seccao.loc[seccao, ('Unit Sales', 'median')]
                    
                    sugestao_base = max(ps_min, ps_median * 0.5)
                    
                    # Fator vendas simplificado
                    if unit_sales_median > 0:
                        ratio = unit_sales / unit_sales_median
                        if unit_sales == 0:
                            fator_vendas = 0.1
                        elif ratio < 0.3: fator_vendas = 0.5
                        elif ratio < 1.0: fator_vendas = 0.8
                        elif ratio < 2.0: fator_vendas = 1.5
                        elif ratio < 4.0: fator_vendas = 2.5
                        else: fator_vendas = 3.5
                    else:
                        fator_vendas = 1.0
                    
                    # Fator preço
                    fator_valor = 1.0
                    if pvp > 50: fator_valor = 0.6
                    elif pvp > 20: fator_valor = 0.8
                    
                    sugestao = sugestao_base * fator_vendas * fator_valor
                    
                    # Ajustar pack size
                    if pack_size > 1 and pack_size <= 12:
                        packs = max(1, round(sugestao / pack_size))
                        sugestao = packs * pack_size
                    else:
                        sugestao = max(3, sugestao)
                    
                    sugestao_final = int(round(min(sugestao, 100)))  # Limite máximo
                    artigos_sem_ps.at[idx, 'Sugestão Presentation Stock'] = max(3, sugestao_final)
                else:
                    artigos_sem_ps.at[idx, 'Sugestão Presentation Stock'] = 3

            self.df.loc[artigos_sem_ps.index, 'Sugestão Presentation Stock'] = artigos_sem_ps['Sugestão Presentation Stock']
            self.df.loc[self.df['Presentation Stock'] > 0, 'Sugestão Presentation Stock'] = 0
            
            QMessageBox.information(self, "Concluído", "Sugestões calculadas usando regras manuais.")
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro nas regras manuais: {str(e)}")
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
                
                # Processar diretamente (sem timer)
                self.processar_ficheiro(file_path)
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar ficheiro: {str(e)}")
            self.progress_bar.setVisible(False)

    def processar_ficheiro(self, file_path):
        try:
            file_extension = file_path.lower().split('.')[-1]
            
            if file_extension in ['xlsx', 'xls']:
                self.df = pd.read_excel(file_path)
            elif file_extension == 'csv':
                self.df = self.carregar_csv(file_path)
            else:
                QMessageBox.critical(self, "Erro", "Formato de ficheiro não suportado.")
                self.progress_bar.setVisible(False)
                return
            
            self.progress_bar.setValue(50)
            
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
            
            self.df['Secção'] = self.df['Merc.Struct Code'].astype(str).str[2:4]
            self.df['Sugestão Presentation Stock'] = 0
            
            # Calcular sugestões
            self.calcular_sugestao_ps()
            
            self.df_filtered = self.df[self.df['Presentation Stock'] == 0].copy()
            self.df_filtered = self.df_filtered.sort_values(['Secção', 'Unit Sales'], ascending=[True, False])
            
            seccoes = sorted(self.df_filtered['Secção'].unique())
            self.combo_seccao.clear()
            self.combo_seccao.addItem("Todas as Secções")
            self.combo_seccao.addItems([str(sec) for sec in seccoes])
            
            if 'Status' in self.df_filtered.columns:
                status_unicos = sorted(self.df_filtered['Status'].dropna().unique())
                self.combo_status.clear()
                self.combo_status.addItem("Todos os Status")
                self.combo_status.addItems([str(status) for status in status_unicos])
                self.combo_status.setEnabled(True)
            else:
                self.combo_status.clear()
                self.combo_status.addItem("Todos os Status")
                self.combo_status.setEnabled(False)
            
            self.progress_bar.setValue(100)
            
            self.label_file.setText(os.path.basename(file_path))
            self.btn_exportar_excel.setEnabled(True)
            self.btn_exportar_pdf.setEnabled(True)
            self.btn_ml.setEnabled(True)
            self.aplicar_filtros()
            
            QMessageBox.information(
                self, 
                "Sucesso", 
                f"Ficheiro carregado com sucesso!\n"
                f"{len(self.df_filtered)} artigos sem Presentation Stock encontrados."
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao processar ficheiro: {str(e)}")
            import traceback
            print(traceback.format_exc())
        finally:
            self.progress_bar.setVisible(False)

    def atualizar_tabela(self, df):
        try:
            if df is None or df.empty:
                self.table.setRowCount(0)
                self.label_contador.setText("Total de artigos sem Presentation Stock: 0")
                return
            
            # Limitar número de linhas para performance
            max_rows = 1000
            if len(df) > max_rows:
                QMessageBox.warning(self, "Muitos Resultados", 
                                  f"Mostrando apenas os primeiros {max_rows} resultados de {len(df)}.\nUse filtros para refinar a busca.")
                df_display = df.head(max_rows)
            else:
                df_display = df.copy()
            
            self.table.setRowCount(len(df_display))
            
            # Definir colunas baseadas nos dados disponíveis
            colunas_disponiveis = []
            colunas_possiveis = [
                'Sku', 'Description', 'Sup.Pack Size', 'PVP Em Vigor', 'Stock', 
                'Unit Sales', 'Flow-type', 'Secção', 'Sugestão Presentation Stock'
            ]
            
            for col in colunas_possiveis:
                if col in df_display.columns:
                    colunas_disponiveis.append(col)
            
            self.table.setColumnCount(len(colunas_disponiveis))
            self.table.setHorizontalHeaderLabels(colunas_disponiveis)
            
            # Configurar largura das colunas
            for i, col in enumerate(colunas_disponiveis):
                if col == 'Description':
                    self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
                else:
                    self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeToContents)
            
            # Preencher tabela
            for row_idx, (_, row) in enumerate(df_display.iterrows()):
                for col_idx, col_name in enumerate(colunas_disponiveis):
                    value = row[col_name] if pd.notna(row[col_name]) else ""
                    
                    if col_name in ['Stock', 'Unit Sales', 'Sup.Pack Size', 'Sugestão Presentation Stock']:
                        try:
                            text = f"{int(value):,}" if value != "" else "0"
                        except:
                            text = str(value)
                        alignment = Qt.AlignRight | Qt.AlignVCenter
                    elif col_name == 'PVP Em Vigor':
                        try:
                            text = f"€{float(value):,.2f}" if value != "" else "€0.00"
                        except:
                            text = str(value)
                        alignment = Qt.AlignRight | Qt.AlignVCenter
                    elif col_name == 'Secção':
                        text = str(value)
                        alignment = Qt.AlignCenter | Qt.AlignVCenter
                    else:
                        text = str(value)
                        alignment = Qt.AlignLeft | Qt.AlignVCenter
                    
                    item = QTableWidgetItem(text)
                    item.setTextAlignment(alignment)
                    
                    # Cores condicionais
                    if col_name == 'Stock':
                        try:
                            stock_val = float(value) if value != "" else 0
                            if stock_val == 0:
                                item.setBackground(QColor(255, 200, 200))
                            elif 'Sugestão Presentation Stock' in colunas_disponiveis:
                                sugestao_idx = colunas_disponiveis.index('Sugestão Presentation Stock')
                                sugestao_val = float(row[colunas_disponiveis[sugestao_idx]]) if pd.notna(row[colunas_disponiveis[sugestao_idx]]) else 0
                                if stock_val < sugestao_val:
                                    item.setBackground(QColor(255, 255, 200))
                        except:
                            pass
                    
                    self.table.setItem(row_idx, col_idx, item)
            
            self.label_contador.setText(f"Total de artigos sem Presentation Stock: {len(df):,}")
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar tabela: {str(e)}")

    def exportar_pdf_fixed(self):
        """Versão corrigida para exportar PDF no Windows"""
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Aviso", "Não existem dados para exportar.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Exportar para PDF", "Artigos_Sem_PS.pdf", "PDF (*.pdf)"
        )
        
        if not file_path:
            return

        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # Preparar dados
            seccao_selecionada = self.combo_seccao.currentText()
            status_selecionado = self.combo_status.currentText()

            if seccao_selecionada == "Todas as Secções":
                df_export = self.df_filtered.copy()
            else:
                df_export = self.df_filtered[self.df_filtered['Secção'] == seccao_selecionada].copy()

            if 'Status' in df_export.columns and status_selecionado != "Todos os Status":
                df_export = df_export[df_export['Status'] == status_selecionado]

            self.progress_bar.setValue(30)
            
            # Criar documento PDF
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(file_path)
            printer.setPageSize(QPageSize(QPageSize.A4))
            printer.setPageOrientation(QPageLayout.Landscape)
            
            # Configurar margens
            printer.setPageMargins(10, 15, 10, 15, QPrinter.Millimeter)
            
            doc = QTextDocument()
            doc.setPageSize(printer.pageRect().size())
            
            # Construir HTML
            html = self.gerar_html_pdf(df_export, seccao_selecionada, status_selecionado)
            doc.setHtml(html)
            
            self.progress_bar.setValue(70)
            
            # Imprimir
            doc.print_(printer)
            
            self.progress_bar.setValue(100)
            
            # Abrir o PDF automaticamente
            resposta = QMessageBox.question(
                self, "Sucesso",
                f"PDF exportado com sucesso!\n\n"
                f"→ {len(df_export):,} artigos exportados\n"
                f"→ Guardado em: {os.path.basename(file_path)}\n\n"
                f"Deseja abrir o ficheiro agora?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            
            if resposta == QMessageBox.Yes:
                QDesktopServices.openUrl(QUrl.fromLocalFile(file_path))
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar PDF:\n{str(e)}")
            import traceback
            print(traceback.format_exc())
        finally:
            self.progress_bar.setVisible(False)

    def gerar_html_pdf(self, df, seccao, status):
        """Gera HTML para PDF de forma mais simples e compatível"""
        
        # Cabeçalho
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    margin: 0;
                    padding: 10mm;
                }}
                .header {{
                    text-align: center;
                    margin-bottom: 20px;
                    border-bottom: 2px solid #333;
                    padding-bottom: 10px;
                }}
                .title {{
                    font-size: 18pt;
                    font-weight: bold;
                }}
                .info {{
                    font-size: 10pt;
                    color: #666;
                    margin: 10px 0;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    font-size: 8pt;
                }}
                th {{
                    background-color: #f0f0f0;
                    padding: 5px;
                    text-align: left;
                    border: 1px solid #ccc;
                    font-weight: bold;
                }}
                td {{
                    padding: 4px;
                    border: 1px solid #ccc;
                }}
                .footer {{
                    margin-top: 20px;
                    font-size: 7pt;
                    color: #999;
                    text-align: center;
                }}
                .numeric {{
                    text-align: right;
                }}
                .center {{
                    text-align: center;
                }}
            </style>
        </head>
        <body>
            <div class="header">
                <div class="title">ARTIGOS SEM PRESENTATION STOCK</div>
                <div class="info">
                    Secção: {seccao} | Status: {status} | 
                    Total de Artigos: {len(df):,} | 
                    Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}
                </div>
            </div>
        """
        
        # Tabela
        html += "<table>"
        html += "<tr>"
        headers = ['SKU', 'Descrição', 'Pack', 'PVP', 'Stock', 'Vendas', 'Flow', 'Secção', 'Sug. PS']
        colunas = ['Sku', 'Description', 'Sup.Pack Size', 'PVP Em Vigor', 
                  'Stock', 'Unit Sales', 'Flow-type', 'Secção', 'Sugestão Presentation Stock']
        
        # Verificar quais colunas existem
        for header, coluna in zip(headers, colunas):
            if coluna in df.columns:
                html += f"<th>{header}</th>"
        
        html += "</tr>"
        
        # Dados
        for _, row in df.iterrows():
            html += "<tr>"
            for header, coluna in zip(headers, colunas):
                if coluna in df.columns:
                    value = row[coluna] if pd.notna(row[coluna]) else ""
                    
                    # Formatação
                    if coluna in ['Stock', 'Unit Sales', 'Sup.Pack Size', 'Sugestão Presentation Stock']:
                        try:
                            text = f"{int(value):,}" if value != "" else "0"
                            html += f"<td class='numeric'>{text}</td>"
                        except:
                            html += f"<td>{value}</td>"
                    elif coluna == 'PVP Em Vigor':
                        try:
                            text = f"€{float(value):,.2f}" if value != "" else "€0.00"
                            html += f"<td class='numeric'>{text}</td>"
                        except:
                            html += f"<td>{value}</td>"
                    elif coluna == 'Secção':
                        html += f"<td class='center'>{value}</td>"
                    elif coluna == 'Description' and len(str(value)) > 50:
                        html += f"<td>{str(value)[:47]}...</td>"
                    else:
                        html += f"<td>{value}</td>"
            
            html += "</tr>"
        
        html += "</table>"
        
        # Rodapé
        html += f"""
            <div class="footer">
                Exportado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} | Total: {len(df):,} artigos
            </div>
        </body>
        </html>
        """
        
        return html

    def exportar_excel(self):
        """Exportar para Excel com tratamento de memória"""
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados para exportar.")
            return
        
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Exportar para Excel",
                f"artigos_sem_ps_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if not file_path:
                return
            
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(30)
            
            # Preparar dados
            seccao_selecionada = self.combo_seccao.currentText()
            status_selecionado = self.combo_status.currentText()

            if seccao_selecionada == "Todas as Secções":
                df_export = self.df_filtered.copy()
            else:
                df_export = self.df_filtered[self.df_filtered['Secção'] == seccao_selecionada].copy()

            if 'Status' in df_export.columns and status_selecionado != "Todos os Status":
                df_export = df_export[df_export['Status'] == status_selecionado]
            
            self.progress_bar.setValue(60)
            
            # Exportar para Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Artigos Sem PS')
                
                # Ajustar largura das colunas
                worksheet = writer.sheets['Artigos Sem PS']
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
                f"Dados exportados com sucesso!\n"
                f"{len(df_export)} artigos exportados para:\n{file_path}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def limpar_tudo(self):
        # Cancelar processamento ML se estiver ativo
        if self.ml_worker and self.ml_worker.isRunning():
            self.ml_worker.terminate()
            self.ml_worker.wait()
        
        self.df = None
        self.df_filtered = None
        self.df_com_ps = None
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.label_file.setText("Nenhum ficheiro carregado")
        self.combo_seccao.clear()
        self.combo_seccao.addItem("Todas as Secções")
        self.combo_status.clear()
        self.combo_status.addItem("Todos os Status")
        self.combo_stock.clear()
        self.combo_stock.addItem("Todos")
        self.combo_stock.addItem("Stock > 0")
        self.combo_stock.addItem("Stock = 0")
        self.search_input.clear()
        self.label_contador.setText("Total de artigos sem Presentation Stock: 0")
        self.btn_exportar_excel.setEnabled(False)
        self.btn_exportar_pdf.setEnabled(False)
        self.btn_ml.setEnabled(False)
        self.processing = False

    def carregar_csv(self, file_path):
        """Método para carregar CSV"""
        try:
            # Tentar diferentes encodings
            encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding, sep=None, engine='python')
                    return df
                except UnicodeDecodeError:
                    continue
                except Exception:
                    continue
            
            # Se nenhum encoding funcionar, tentar com delimitador específico
            return pd.read_csv(file_path, encoding='latin-1', delimiter=';')
            
        except Exception as e:
            raise Exception(f"Erro ao ler ficheiro CSV: {str(e)}")
    
    def closeEvent(self, event):
        """Garantir que threads são terminadas ao fechar"""
        if self.ml_worker and self.ml_worker.isRunning():
            self.ml_worker.terminate()
            self.ml_worker.wait()
        event.accept()

def mostrar_artigos_sem_ps():
    """Função para ser chamada por outros módulos - versão segura"""
    try:
        import sys
        from PyQt5.QtWidgets import QApplication
        
        # Verificar se já existe uma QApplication
        app = QApplication.instance()
        if app is None:
            # Não existe, criar nova
            app = QApplication(sys.argv)
            app.setStyle('Fusion')
            
            # Criar diálogo
            dialog = ArtigosSemPSDialog()
            dialog.show()
            
            # Executar
            sys.exit(app.exec_())
        else:
            # Já existe uma QApplication, apenas criar e mostrar o diálogo
            dialog = ArtigosSemPSDialog()
            dialog.show()
            dialog.raise_()
            dialog.activateWindow()
            
    except Exception as e:
        print(f"Erro ao abrir Artigos Sem PS: {e}")


def main():
    """Função principal - para executar diretamente este ficheiro"""
    import sys
    from PyQt5.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    dialog = ArtigosSemPSDialog()
    dialog.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()