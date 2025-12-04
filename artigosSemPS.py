import os
import pandas as pd
import numpy as np

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

import multiprocessing
import sys

class ArtigosSemPSDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Artigos sem Presentation Stock")
        self.setGeometry(100, 100, 1600, 800)
        self.df = None
        self.df_filtered = None
        self.df_com_ps = None
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
        self.combo_seccao.currentTextChanged.connect(self.aplicar_filtros)
        filters_layout.addWidget(self.combo_seccao)

        filters_layout.addWidget(QLabel("Status:"))

        self.combo_status = QComboBox()
        self.combo_status.setMinimumWidth(150)
        self.combo_status.addItem("Todos os Status")
        self.combo_status.currentTextChanged.connect(self.aplicar_filtros)
        filters_layout.addWidget(self.combo_status)

        # NOVO: Filtro para Stock
        filters_layout.addWidget(QLabel("Stock:"))

        self.combo_stock = QComboBox()
        self.combo_stock.setMinimumWidth(150)
        self.combo_stock.addItem("Todos")
        self.combo_stock.addItem("Stock > 0")
        self.combo_stock.addItem("Stock = 0")
        self.combo_stock.currentTextChanged.connect(self.aplicar_filtros)
        filters_layout.addWidget(self.combo_stock)

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
            stock_selecionado = self.combo_stock.currentText()  # NOVO
            
            # Aplicar filtro por secção
            if seccao_selecionada == "Todas as Secções":
                df_temp = self.df_filtered.copy()
            else:
                df_temp = self.df_filtered[self.df_filtered['Secção'] == seccao_selecionada].copy()
            
            # Aplicar filtro por status (se a coluna existir)
            if 'Status' in df_temp.columns and status_selecionado != "Todos os Status":
                df_temp = df_temp[df_temp['Status'] == status_selecionado]
            
            # NOVO: Aplicar filtro por stock
            if stock_selecionado == "Stock > 0":
                df_temp = df_temp[df_temp['Stock'] > 0]
            elif stock_selecionado == "Stock = 0":
                df_temp = df_temp[df_temp['Stock'] == 0]
            # "Todos" não aplica filtro
            
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
        """Calcula sugestão de Presentation Stock usando Machine Learning"""
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(10)
            
            df_ml = self.df.copy()
            
            # Features numéricas especificadas
            numeric_features = [
                'Sup.Code', 'Sup.Pack Size', 'PVP Permanente', 'Stock', 
                'Stock In Transit', 'Stock Expected', 'Last Order Point', 
                'Lead Time WH to Location', 'Unit Sales', 'Sales Value'
            ]
            
            # Features categóricas especificadas
            categorical_features = ['Flow-type', 'Dta activacao', 'GLP']
            
            # Verifica quais features existem no DataFrame
            available_numeric = [f for f in numeric_features if f in df_ml.columns]
            available_categorical = [f for f in categorical_features if f in df_ml.columns]
            
            all_features = available_numeric + available_categorical
            
            if len(all_features) < 3:
                QMessageBox.warning(self, "Aviso", 
                    f"Poucas features disponíveis ({len(all_features)}). Mínimo recomendado: 3")
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
            
            self.progress_bar.setValue(30)
            
            # Prepara dados
            X_train, y_train, X_pred, preprocessor, feature_names = self.preparar_dados_ml(
                df_treino, df_prever, available_numeric, available_categorical
            )
            
            self.progress_bar.setValue(60)
            
            # Treina modelo
            modelo, mae_score, nome_modelo = self.treinar_modelo_ml(X_train, y_train)
            
            self.progress_bar.setValue(80)
            
            # Faz previsões
            previsoes = modelo.predict(X_pred)
            
            # Aplica previsões com constraints
            df_prever['Sugestão Presentation Stock'] = np.round(previsoes).astype(int)
            df_prever['Sugestão Presentation Stock'] = df_prever['Sugestão Presentation Stock'].clip(
                lower=1, upper=200
            )
            
            # Aplica lógica de pack size (se existir)
            if hasattr(self, 'aplicar_logica_pack_size'):
                df_prever = self.aplicar_logica_pack_size(df_prever)
            
            # Atualiza DataFrame principal
            self.df.loc[df_prever.index, 'Sugestão Presentation Stock'] = df_prever['Sugestão Presentation Stock']
            self.df.loc[self.df['Presentation Stock'] > 0, 'Sugestão Presentation Stock'] = 0
            
            self.progress_bar.setValue(100)
            
            # Mostra métricas
            self.mostrar_metricas_ml(modelo, X_train, y_train, df_prever, mae_score, 
                                    nome_modelo, feature_names)
            
            # Atualiza visualização
            self.df_filtered = self.df[self.df['Presentation Stock'] == 0].copy()
            self.aplicar_filtros()
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro no ML: {str(e)}")
            import traceback
            print(traceback.format_exc())
            if hasattr(self, 'calcular_sugestao_ps_regras'):
                self.calcular_sugestao_ps_regras()
        finally:
            self.progress_bar.setVisible(False)


    def preparar_dados_ml(self, df_treino, df_prever, numeric_features, categorical_features):
        """Prepara dados para Machine Learning com tratamento de valores ausentes"""
        
        all_features = numeric_features + categorical_features
        
        # Cria cópias dos dados
        X_train = df_treino[all_features].copy()
        y_train = df_treino['Presentation Stock'].values
        X_pred = df_prever[all_features].copy()
        
        # Pipeline para features numéricas (imputer + scaler)
        numeric_transformer = Pipeline(steps=[
            ('imputer', SimpleImputer(strategy='median')),
            ('scaler', StandardScaler())
        ])
        
        # Pipeline para features categóricas (imputer + encoder)
        categorical_transformer = Pipeline(steps=[
            ('imputer', SimpleImputer(strategy='constant', fill_value='missing')),
            ('encoder', OneHotEncoder(handle_unknown='ignore', sparse_output=False))
        ])
        
        # Combina transformers
        preprocessor = ColumnTransformer(
            transformers=[
                ('num', numeric_transformer, numeric_features),
                ('cat', categorical_transformer, categorical_features)
            ],
            remainder='drop'
        )
        
        # Transforma dados
        X_train_processed = preprocessor.fit_transform(X_train)
        X_pred_processed = preprocessor.transform(X_pred)
        
        # Gera nomes de features para análise posterior
        feature_names = self._get_feature_names(preprocessor, numeric_features, categorical_features)
        
        print(f"Features utilizadas: {len(all_features)}")
        print(f"- Numéricas: {len(numeric_features)}")
        print(f"- Categóricas: {len(categorical_features)}")
        print(f"Amostras treino: {len(X_train_processed)}, Amostras previsão: {len(X_pred_processed)}")
        
        return X_train_processed, y_train, X_pred_processed, preprocessor, feature_names


    def _get_feature_names(self, preprocessor, numeric_features, categorical_features):
        """Obtém nomes das features após transformação"""
        try:
            feature_names = []
            
            # Features numéricas mantêm o nome
            feature_names.extend(numeric_features)
            
            # Features categóricas são expandidas pelo OneHotEncoder
            if categorical_features:
                cat_encoder = preprocessor.named_transformers_['cat'].named_steps['encoder']
                for i, cat_feature in enumerate(categorical_features):
                    categories = cat_encoder.categories_[i]
                    feature_names.extend([f"{cat_feature}_{cat}" for cat in categories])
            
            return feature_names
        except:
            return [f"feature_{i}" for i in range(len(numeric_features) + len(categorical_features))]


    def treinar_modelo_ml(self, X_train, y_train):
        """Treina modelo com validação cruzada e seleção do melhor"""        
        
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
        
        print("\n=== Avaliação dos Modelos ===")
        
        for nome, modelo in modelos.items():
            try:
                # Validação cruzada
                cv_folds = min(5, len(X_train) // 2)  # Ajusta número de folds
                
                scores = cross_val_score(
                    modelo, X_train, y_train, 
                    cv=cv_folds, 
                    scoring='neg_mean_absolute_error',
                    n_jobs=-1
                )
                
                mae_score = -scores.mean()
                mae_std = scores.std()
                
                print(f"{nome}: MAE = {mae_score:.2f} (±{mae_std:.2f})")
                
                if mae_score < melhor_score:
                    melhor_score = mae_score
                    melhor_modelo = modelo
                    melhor_nome = nome
                    
            except Exception as e:
                print(f"Erro ao treinar {nome}: {str(e)}")
                continue
        
        # Treina o melhor modelo com todos os dados
        if melhor_modelo is not None:
            melhor_modelo.fit(X_train, y_train)
            print(f"\n✓ Melhor modelo selecionado: {melhor_nome} (MAE: {melhor_score:.2f})")
            return melhor_modelo, melhor_score, melhor_nome
        else:
            # Fallback para modelo simples
            print("\n⚠ Usando modelo fallback (RandomForest simplificado)")
            modelo = RandomForestRegressor(n_estimators=50, random_state=42)
            modelo.fit(X_train, y_train)
            return modelo, 0, "RandomForest (Fallback)"


    def mostrar_metricas_ml(self, modelo, X_train, y_train, df_prever, mae_score, 
                            nome_modelo, feature_names):
        """Mostra métricas do modelo treinado"""
        
        msg = f"=== Resultados do Machine Learning ===\n\n"
        msg += f"Modelo utilizado: {nome_modelo}\n"
        msg += f"MAE (Validação Cruzada): {mae_score:.2f}\n\n"
        
        msg += f"Dados de treino: {len(X_train)} artigos\n"
        msg += f"Previsões geradas: {len(df_prever)} artigos\n\n"
        
        # Estatísticas das previsões
        sugestoes = df_prever['Sugestão Presentation Stock']
        msg += f"Sugestões geradas:\n"
        msg += f"  - Mínimo: {sugestoes.min()}\n"
        msg += f"  - Média: {sugestoes.mean():.1f}\n"
        msg += f"  - Mediana: {sugestoes.median():.1f}\n"
        msg += f"  - Máximo: {sugestoes.max()}\n\n"
        
        # Feature importance (se disponível)
        if hasattr(modelo, 'feature_importances_'):
            importances = modelo.feature_importances_
            top_n = min(5, len(importances))
            top_indices = np.argsort(importances)[-top_n:][::-1]
            
            msg += "Top 5 Features mais importantes:\n"
            for idx in top_indices:
                if idx < len(feature_names):
                    msg += f"  - {feature_names[idx]}: {importances[idx]:.3f}\n"
        
        QMessageBox.information(self, "Métricas ML", msg)
        print(msg)

    def calcular_sugestao_ps_regras(self):
        """Método baseado em regras (fallback) - VERSÃO SIMPLIFICADA"""
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

    def atualizar_tabela(self, df):
        try:
            self.table.setRowCount(len(df))
            self.table.setColumnCount(9)
            self.table.setHorizontalHeaderLabels([
                'Sku', 'Description', 'Sup.Pack Size', 'PVP Em Vigor', 'Stock', 
                'Unit Sales', 'Flow-type', 'Secção', 'Sugestão Presentation Stock'
            ])
            
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
                
                if stock_value == 0:
                    item_stock.setBackground(QColor(255, 200, 200))
                elif stock_value < (row.get('Sugestão Presentation Stock', 0) or 0):
                    item_stock.setBackground(QColor(255, 255, 200))
                
                self.table.setItem(row_idx, 4, item_stock)
                
                # Unit Sales
                unit_sales_value = row['Unit Sales'] if pd.notna(row['Unit Sales']) else 0
                item_unit_sales = QTableWidgetItem(f"{unit_sales_value:,.0f}")
                item_unit_sales.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                if unit_sales_value > 100:
                    item_unit_sales.setBackground(QColor(200, 255, 200))
                elif unit_sales_value > 50:
                    item_unit_sales.setBackground(QColor(255, 255, 200))
                
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
                
                if sugestao > 10:
                    item_sugestao.setBackground(QColor(200, 230, 255))
                
                self.table.setItem(row_idx, 8, item_sugestao)
            
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

            if seccao_selecionada == "Todas as Secções":
                df_export = self.df_filtered.copy()
            else:
                df_export = self.df_filtered[self.df_filtered['Secção'] == seccao_selecionada].copy()

            if 'Status' in df_export.columns and status_selecionado != "Todos os Status":
                df_export = df_export[df_export['Status'] == status_selecionado]

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
            cursor.insertText("ARTIGOS SEM PRESENTATION STOCK\n\n")

            info = f"Secção: {seccao_selecionada} | " \
                f"Total artigos: {len(df_export):,} | " \
                f"Gerado em: {pd.Timestamp.now():%d/%m/%Y %H:%M}\n\n"
            cursor.insertText(info)

            headers = [
                'Sku', 'Description', 'Pack', 'PVP', 'Stock', 
                'Unit Sales', 'Flow', 'Sec', 'Sug. Presentation Stock'
            ]

            larguras_percentagem = [10, 30, 6, 8, 8, 9, 8, 6, 8]

            table_fmt = QTextTableFormat()
            table_fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100))
            table_fmt.setCellPadding(4)
            table_fmt.setCellSpacing(0)
            table_fmt.setBorder(0.5)
            table_fmt.setBorderStyle(QTextFrameFormat.BorderStyle_Solid)

            constraints = [QTextLength(QTextLength.PercentageLength, w) for w in larguras_percentagem]
            table_fmt.setColumnWidthConstraints(constraints)

            table = cursor.insertTable(len(df_export) + 1, len(headers), table_fmt)

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

            normal_fmt = QTextCharFormat()
            normal_fmt.setFontPointSize(8)

            for row_idx, (_, row) in enumerate(df_export.iterrows(), start=1):
                for col_idx, col_name in enumerate(headers):
                    cell = table.cellAt(row_idx, col_idx)
                    cur = cell.firstCursorPosition()

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

            cursor.movePosition(QTextCursor.End)
            cursor.insertBlock()
            footer = QTextCharFormat()
            footer.setFontPointSize(7)
            footer.setFontItalic(True)
            footer.setForeground(QColor("gray"))
            cursor.setCharFormat(footer)
            cursor.insertText(f"Documento gerado automaticamente • {len(df_export):,} artigos sem Presentation Stock")

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
                
                seccao_selecionada = self.combo_seccao.currentText()
                status_selecionado = self.combo_status.currentText()

                if seccao_selecionada == "Todas as Secções":
                    df_export = self.df_filtered.copy()
                else:
                    df_export = self.df_filtered[self.df_filtered['Secção'] == seccao_selecionada].copy()

                if 'Status' in df_export.columns and status_selecionado != "Todos os Status":
                    df_export = df_export[df_export['Status'] == status_selecionado]
                
                colunas_export = [
                    'Sku', 'Description', 'Sup.Pack Size', 'PVP Em Vigor', 'Stock', 
                    'Unit Sales', 'Flow-type', 'Secção', 'Sugestão Presentation Stock'
                ]
                
                colunas_disponiveis = [col for col in colunas_export if col in df_export.columns]
                df_export = df_export[colunas_disponiveis].copy()
                
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Artigos Sem Presentation Stock')
                    
                    worksheet = writer.sheets['Artigos Sem Presentation Stock']
                    
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
        self.combo_status.clear()
        self.combo_status.addItem("Todos os Status")
        self.combo_stock.clear()  # NOVO
        self.combo_stock.addItem("Todos")  # NOVO
        self.combo_stock.addItem("Stock > 0")  # NOVO
        self.combo_stock.addItem("Stock = 0")  # NOVO
        self.combo_status.setEnabled(True)
        self.label_contador.setText("Total de artigos sem Presentation Stock: 0")
        self.btn_exportar_excel.setEnabled(False)
        self.btn_exportar_pdf.setEnabled(False)
        self.btn_ml.setEnabled(False)

def mostrar_artigos_sem_ps():
    """Função auxiliar para abrir o dialog - segura para chamada externa"""
    if QApplication.instance() is None:
        # Se não houver app, cria uma (não deve acontecer normalmente)
        app = QApplication(sys.argv)
        dialog = ArtigosSemPSDialog()
        dialog.exec_()
    else:
        # Usa a app existente
        dialog = ArtigosSemPSDialog()
        dialog.exec_()


def main():
    """Função principal - apenas para execução direta deste arquivo"""
    app = QApplication(sys.argv)
    dialog = ArtigosSemPSDialog()
    dialog.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    # Configuração multiprocessing ANTES de tudo
    if sys.platform.startswith('win'):
        multiprocessing.freeze_support()
        multiprocessing.set_start_method('spawn', force=True)
    
    main()
