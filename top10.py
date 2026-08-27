# top10.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QProgressBar, QTableWidget,
    QTableWidgetItem, QHeaderView, QComboBox, QSpinBox,
    QCheckBox, QApplication
)
from PyQt5.QtGui import (
    QFont, QColor, QPageLayout, QPageSize, QTextDocument, QTextCursor,
    QTextTableFormat, QTextTableCellFormat, QTextCharFormat,
    QTextBlockFormat, QTextLength, QTextFrameFormat
)
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtCore import Qt, QMarginsF


class TopNDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Top N Artigos - ExcelExtractions")
        self.setGeometry(100, 100, 1400, 800)
        self.df = None
        self.df_top = None
        self.df_filtrado = None
        self.modo_atual = 'top'          # 'top' ou 'pareto'
        self.metrica_pareto_atual = None
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        layout.setSpacing(15)

        title = QLabel("🏆 Gerador de Top N Artigos")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("margin: 10px;")
        layout.addWidget(title)

        upload_layout = QHBoxLayout()
        self.btn_file = QPushButton("📁 Carregar Ficheiro (CSV/Excel)")
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
            QPushButton:hover { background-color: #45a049; }
        """)
        self.btn_file.clicked.connect(self.carregar_ficheiro)
        upload_layout.addWidget(self.btn_file)

        self.label_file = QLabel("Nenhum ficheiro carregado")
        self.label_file.setStyleSheet("color: #666; padding: 10px;")
        upload_layout.addWidget(self.label_file)
        upload_layout.addStretch()
        layout.addLayout(upload_layout)

        filtros_layout = QHBoxLayout()
        filtros_layout.setSpacing(10)

        filtros_layout.addWidget(QLabel("Categoria:"))
        self.combo_categoria = QComboBox()
        self.combo_categoria.setMinimumWidth(150)
        self.combo_categoria.addItem("Todas")
        self.combo_categoria.currentTextChanged.connect(self.aplicar_filtros)
        filtros_layout.addWidget(self.combo_categoria)

        filtros_layout.addWidget(QLabel("Sup.Name:"))
        self.combo_supname = QComboBox()
        self.combo_supname.setMinimumWidth(150)
        self.combo_supname.addItem("Todos")
        self.combo_supname.currentTextChanged.connect(self.aplicar_filtros)
        filtros_layout.addWidget(self.combo_supname)

        filtros_layout.addWidget(QLabel("Brand:"))
        self.combo_brand = QComboBox()
        self.combo_brand.setMinimumWidth(150)
        self.combo_brand.addItem("Todos")
        self.combo_brand.currentTextChanged.connect(self.aplicar_filtros)
        filtros_layout.addWidget(self.combo_brand)

        filtros_layout.addStretch()
        layout.addLayout(filtros_layout)

        config_layout = QHBoxLayout()
        config_layout.addWidget(QLabel("Número de artigos:"))

        self.spin_n = QSpinBox()
        self.spin_n.setMinimum(5)
        self.spin_n.setMaximum(500)
        self.spin_n.setValue(10)
        self.spin_n.setSuffix(" itens")
        self.spin_n.setFixedWidth(120)
        config_layout.addWidget(self.spin_n)

        config_layout.addSpacing(20)

        self.chk_novos = QCheckBox("Incluir 5% de artigos novos com vendas (>0)")
        self.chk_novos.setChecked(True)
        config_layout.addWidget(self.chk_novos)

        config_layout.addSpacing(20)

        self.chk_apenas_com_vendas = QCheckBox("Apenas artigos com vendas (>0)")
        self.chk_apenas_com_vendas.setChecked(True)
        config_layout.addWidget(self.chk_apenas_com_vendas)

        config_layout.addStretch()

        self.btn_calcular = QPushButton("🔍 Calcular Top")
        self.btn_calcular.setFont(QFont("Arial", 12))
        self.btn_calcular.setMinimumHeight(40)
        self.btn_calcular.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover { background-color: #0b7dda; }
            QPushButton:disabled { background-color: #cccccc; color: #666; }
        """)
        self.btn_calcular.setEnabled(False)
        self.btn_calcular.clicked.connect(self.calcular_top)
        config_layout.addWidget(self.btn_calcular)

        layout.addLayout(config_layout)

        # --- Linha do Pareto 80/20 ---
        pareto_layout = QHBoxLayout()
        pareto_layout.addWidget(QLabel("Regra 80/20 (Pareto) por:"))

        self.combo_pareto_metrica = QComboBox()
        self.combo_pareto_metrica.addItems(["Sales Value", "Unit Sales"])
        self.combo_pareto_metrica.setMinimumWidth(120)
        pareto_layout.addWidget(self.combo_pareto_metrica)

        self.btn_pareto = QPushButton("📊 Top 80/20 (Pareto)")
        self.btn_pareto.setFont(QFont("Arial", 12))
        self.btn_pareto.setMinimumHeight(40)
        self.btn_pareto.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover { background-color: #e68900; }
            QPushButton:disabled { background-color: #cccccc; color: #666; }
        """)
        self.btn_pareto.setEnabled(False)
        self.btn_pareto.clicked.connect(self.calcular_pareto)
        pareto_layout.addWidget(self.btn_pareto)

        pareto_layout.addStretch()
        layout.addLayout(pareto_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget {
                gridline-color: #d0d0d0;
                background-color: white;
            }
            QTableWidget::item { padding: 5px; }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 5px;
                border: 1px solid #d0d0d0;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.table)

        buttons_layout = QHBoxLayout()

        self.btn_exportar_excel = QPushButton("💾 Exportar para Excel")
        self.btn_exportar_excel.setFont(QFont("Arial", 12))
        self.btn_exportar_excel.setMinimumHeight(40)
        self.btn_exportar_excel.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover { background-color: #45a049; }
            QPushButton:disabled { background-color: #cccccc; color: #666; }
        """)
        self.btn_exportar_excel.setEnabled(False)
        self.btn_exportar_excel.clicked.connect(self.exportar_excel)
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
            QPushButton:hover { background-color: #d32f2f; }
            QPushButton:disabled { background-color: #cccccc; color: #666; }
        """)
        self.btn_exportar_pdf.setEnabled(False)
        self.btn_exportar_pdf.clicked.connect(self.exportar_pdf)
        buttons_layout.addWidget(self.btn_exportar_pdf)

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
            QPushButton:hover { background-color: #546E7A; }
        """)
        self.btn_fechar.clicked.connect(self.close)
        buttons_layout.addWidget(self.btn_fechar)

        buttons_layout.addStretch()
        layout.addLayout(buttons_layout)

        self.setLayout(layout)

    def carregar_ficheiro(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Selecionar Ficheiro",
                "",
                "Ficheiros Suportados (*.xlsx *.xls *.csv);;Excel (*.xlsx *.xls);;CSV (*.csv)"
            )
            if not file_path:
                return

            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(10)

            if file_path.lower().endswith(('.xlsx', '.xls')):
                self.df = pd.read_excel(file_path)
            else:
                self.df = self._carregar_csv(file_path)

            self.df.columns = self.df.columns.str.strip()

            self.progress_bar.setValue(50)

            colunas_necessarias = ['Sku', 'Unit Sales', 'Sales Value', 'PVP Em Vigor', 'Merc.Struct Code']
            faltam = [c for c in colunas_necessarias if c not in self.df.columns]
            if faltam:
                primeiras_linhas = self.df.head(2).to_string()
                QMessageBox.critical(
                    self, "Erro",
                    f"Colunas obrigatórias não encontradas: {', '.join(faltam)}\n\n"
                    f"Colunas disponíveis: {', '.join(self.df.columns)}\n\n"
                    f"Primeiras linhas do ficheiro:\n{primeiras_linhas}"
                )
                self.df = None
                self.progress_bar.setVisible(False)
                return

            self.df['Categoria'] = self.df['Merc.Struct Code'].astype(str).str[:8]

            self.df['Unit Sales'] = pd.to_numeric(self.df['Unit Sales'], errors='coerce').fillna(0)
            self.df['Sales Value'] = pd.to_numeric(self.df['Sales Value'], errors='coerce').fillna(0)
            self.df['PVP Em Vigor'] = pd.to_numeric(self.df['PVP Em Vigor'], errors='coerce').fillna(0)

            try:
                self.df['Sku_num'] = pd.to_numeric(self.df['Sku'], errors='coerce')
            except:
                self.df['Sku_num'] = self.df['Sku'].astype(str).apply(lambda x: int(x) if x.isdigit() else hash(x))

            stock_cols = ['Stock', 'Stock In Transit', 'Stock Expected', 'Stock On Order']
            existentes = [c for c in stock_cols if c in self.df.columns]
            for c in existentes:
                self.df[c] = pd.to_numeric(self.df[c], errors='coerce').fillna(0)
            self.df['Stock Total'] = self.df[existentes].sum(axis=1) if existentes else 0

            if 'Warehouse' not in self.df.columns:
                self.df['Warehouse'] = ''

            self._preencher_filtros()

            self.progress_bar.setValue(100)

            self.label_file.setText(os.path.basename(file_path))
            self.btn_calcular.setEnabled(True)
            self.btn_pareto.setEnabled(True)
            self.btn_exportar_excel.setEnabled(False)
            self.btn_exportar_pdf.setEnabled(False)

            self.df_filtrado = self.df.copy()
            self.aplicar_filtros()

            QMessageBox.information(
                self, "Sucesso",
                f"Ficheiro carregado com {len(self.df):,} artigos válidos."
            )

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar ficheiro: {str(e)}")
            import traceback
            print(traceback.format_exc())
        finally:
            self.progress_bar.setVisible(False)

    def _carregar_csv(self, file_path):
        encodings = ['utf-8', 'latin-1', 'cp1252']
        delimiters = [';', ',', '\t', '|']
        
        for enc in encodings:
            for delim in delimiters:
                try:
                    df = pd.read_csv(file_path, encoding=enc, sep=delim, engine='python')
                    df.columns = df.columns.str.strip()
                    if len(df.columns) > 2 and 'Sku' in df.columns:
                        print(f"CSV lido com encoding={enc}, delimitador='{delim}'")
                        return df
                except:
                    continue
        
        try:
            with open(file_path, 'r', encoding='latin-1') as f:
                sample = f.read(1024)
            import csv
            sniffer = csv.Sniffer()
            dialect = sniffer.sniff(sample)
            df = pd.read_csv(file_path, encoding='latin-1', sep=dialect.delimiter, engine='python')
            df.columns = df.columns.str.strip()
            print(f"CSV lido com Sniffer: delimitador='{dialect.delimiter}'")
            return df
        except Exception as e:
            raise Exception(f"Não foi possível ler o CSV: {str(e)}")

    def _preencher_filtros(self):
        categorias = sorted(self.df['Categoria'].dropna().unique())
        self.combo_categoria.clear()
        self.combo_categoria.addItem("Todas")
        self.combo_categoria.addItems([str(c) for c in categorias])

        if 'Sup.Name' in self.df.columns:
            sup_names = sorted(self.df['Sup.Name'].dropna().unique())
            self.combo_supname.clear()
            self.combo_supname.addItem("Todos")
            self.combo_supname.addItems([str(s) for s in sup_names])
            self.combo_supname.setEnabled(True)
        else:
            self.combo_supname.clear()
            self.combo_supname.addItem("Todos")
            self.combo_supname.setEnabled(False)

        if 'Brand' in self.df.columns:
            brands = sorted(self.df['Brand'].dropna().unique())
            self.combo_brand.clear()
            self.combo_brand.addItem("Todos")
            self.combo_brand.addItems([str(b) for b in brands])
            self.combo_brand.setEnabled(True)
        else:
            self.combo_brand.clear()
            self.combo_brand.addItem("Todos")
            self.combo_brand.setEnabled(False)

    def aplicar_filtros(self):
        if self.df is None:
            return

        df_filtrado = self.df.copy()

        cat = self.combo_categoria.currentText()
        if cat != "Todas":
            df_filtrado = df_filtrado[df_filtrado['Categoria'] == cat]

        sup = self.combo_supname.currentText()
        if sup != "Todos" and 'Sup.Name' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['Sup.Name'] == sup]

        brand = self.combo_brand.currentText()
        if brand != "Todos" and 'Brand' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['Brand'] == brand]

        self.df_filtrado = df_filtrado

    def calcular_top(self):
        if self.df_filtrado is None or self.df_filtrado.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados após aplicar os filtros.")
            return

        self.modo_atual = 'top'
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(10)

            N = self.spin_n.value()
            incluir_novos = self.chk_novos.isChecked()
            apenas_com_vendas = self.chk_apenas_com_vendas.isChecked()

            df_trab = self.df_filtrado.copy()

            if apenas_com_vendas:
                df_trab = df_trab[df_trab['Unit Sales'] > 0]
                if df_trab.empty:
                    QMessageBox.warning(self, "Aviso", "Nenhum artigo com vendas > 0.")
                    self.progress_bar.setVisible(False)
                    return

            min_sales = df_trab['Unit Sales'].min()
            max_sales = df_trab['Unit Sales'].max()
            df_trab['norm_sales'] = (df_trab['Unit Sales'] - min_sales) / (max_sales - min_sales) if max_sales > min_sales else 0

            min_val = df_trab['Sales Value'].min()
            max_val = df_trab['Sales Value'].max()
            df_trab['norm_value'] = (df_trab['Sales Value'] - min_val) / (max_val - min_val) if max_val > min_val else 0

            min_pvp = df_trab['PVP Em Vigor'].min()
            max_pvp = df_trab['PVP Em Vigor'].max()
            df_trab['norm_price'] = (df_trab['PVP Em Vigor'] - min_pvp) / (max_pvp - min_pvp) if max_pvp > min_pvp else 0

            min_sku = df_trab['Sku_num'].min()
            max_sku = df_trab['Sku_num'].max()
            df_trab['norm_sku'] = (df_trab['Sku_num'] - min_sku) / (max_sku - min_sku) if max_sku > min_sku else 0

            self.progress_bar.setValue(40)

            w_sales, w_value, w_price, w_sku = 0.4, 0.3, 0.2, 0.1
            df_trab['score'] = (w_sales * df_trab['norm_sales'] +
                                w_value * df_trab['norm_value'] +
                                w_price * df_trab['norm_price'] +
                                w_sku * df_trab['norm_sku'])

            df_trab = df_trab.sort_values('score', ascending=False)

            self.progress_bar.setValue(60)

            if incluir_novos:
                limite_novos = max(1, int(0.1 * len(df_trab)))
                df_novos = df_trab.nlargest(limite_novos, 'Sku_num')
                df_novos_com_vendas = df_novos[df_novos['Unit Sales'] > 0]

                qtd_novos = max(1, int(0.05 * N))
                if len(df_novos_com_vendas) >= qtd_novos:
                    selecionados_novos = df_novos_com_vendas.nlargest(qtd_novos, 'score')
                else:
                    selecionados_novos = df_novos_com_vendas

                restantes = df_trab[~df_trab.index.isin(selecionados_novos.index)]
                qtd_restantes = N - len(selecionados_novos)
                selecionados_restantes = restantes.head(qtd_restantes)

                df_top = pd.concat([selecionados_novos, selecionados_restantes])
                df_top = df_top.sort_values('score', ascending=False)
            else:
                df_top = df_trab.head(N)

            self.df_top = df_top.reset_index(drop=True)

            self.progress_bar.setValue(80)
            self.atualizar_tabela(df_top)
            self.progress_bar.setValue(100)

            self.btn_exportar_excel.setEnabled(True)
            self.btn_exportar_pdf.setEnabled(True)

            QMessageBox.information(self, "Concluído", f"Top {len(df_top)} artigos calculados!")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao calcular Top: {str(e)}")
            import traceback
            print(traceback.format_exc())
        finally:
            self.progress_bar.setVisible(False)

    def calcular_pareto(self):
        """Calcula quantos artigos (e que %) perfazem 80% das vendas — Unit Sales e Sales Value."""
        if self.df_filtrado is None or self.df_filtrado.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados após aplicar os filtros.")
            return

        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(20)

            metrica_col = self.combo_pareto_metrica.currentText()
            df_base = self.df_filtrado.copy()
            total_artigos = len(df_base)

            resumo = []
            for col in ['Unit Sales', 'Sales Value']:
                df_tmp = df_base[df_base[col] > 0].sort_values(col, ascending=False).reset_index(drop=True)
                total_col = df_tmp[col].sum()
                if total_col <= 0 or df_tmp.empty:
                    resumo.append(f"• {col}: sem vendas > 0 nos dados filtrados.")
                    continue
                cum_pct = df_tmp[col].cumsum() / total_col * 100
                corte = int((cum_pct >= 80).idxmax())
                n_artigos = corte + 1
                pct_artigos = n_artigos / total_artigos * 100
                resumo.append(
                    f"• {col}: {n_artigos:,} artigos ({pct_artigos:.1f}% do total filtrado) geram 80%"
                )

            self.progress_bar.setValue(50)

            df_pareto = df_base[df_base[metrica_col] > 0].sort_values(metrica_col, ascending=False).reset_index(drop=True)
            if df_pareto.empty:
                QMessageBox.warning(self, "Aviso", f"Nenhum artigo com {metrica_col} > 0 após os filtros.")
                self.progress_bar.setVisible(False)
                return

            total_metrica = df_pareto[metrica_col].sum()
            df_pareto['cum_pct'] = df_pareto[metrica_col].cumsum() / total_metrica * 100
            corte = int((df_pareto['cum_pct'] >= 80).idxmax())
            df_pareto_top = df_pareto.iloc[:corte + 1].copy()

            self.df_top = df_pareto_top.reset_index(drop=True)
            self.modo_atual = 'pareto'
            self.metrica_pareto_atual = metrica_col

            self.progress_bar.setValue(80)
            self.atualizar_tabela(self.df_top)
            self.progress_bar.setValue(100)

            self.btn_exportar_excel.setEnabled(True)
            self.btn_exportar_pdf.setEnabled(True)

            n_pareto = len(df_pareto_top)
            pct_pareto = n_pareto / total_artigos * 100

            QMessageBox.information(
                self, "Análise 80/20 (Pareto)",
                "Regra 80/20 nos dados filtrados:\n\n" + "\n".join(resumo) +
                f"\n\nTabela e exportação carregadas com os {n_pareto:,} artigos "
                f"({pct_pareto:.1f}% do total filtrado) que geram 80% de {metrica_col}."
            )

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao calcular Pareto: {str(e)}")
            import traceback
            print(traceback.format_exc())
        finally:
            self.progress_bar.setVisible(False)

    def atualizar_tabela(self, df):
        if df is None or df.empty:
            self.table.setRowCount(0)
            return

        if self.modo_atual == 'pareto':
            ultima_label = '% Acumulada'
        else:
            ultima_label = 'Score'

        colunas = ['Posição', 'Sku', 'Description', 'Categoria', 'Warehouse',
                   'PVP Em Vigor', 'Stock', 'Unit Sales', 'Sales Value', ultima_label]
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(colunas))
        self.table.setHorizontalHeaderLabels(colunas)

        for i, (_, row) in enumerate(df.iterrows()):
            item_pos = QTableWidgetItem(str(i+1))
            item_pos.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 0, item_pos)

            item_sku = QTableWidgetItem(str(row['Sku']))
            item_sku.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.table.setItem(i, 1, item_sku)

            desc = str(row['Description']) if pd.notna(row['Description']) else ''
            item_desc = QTableWidgetItem(desc[:35])
            item_desc.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.table.setItem(i, 2, item_desc)

            cat = str(row['Categoria']) if pd.notna(row['Categoria']) else ''
            item_cat = QTableWidgetItem(cat)
            item_cat.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 3, item_cat)

            wh = str(row['Warehouse']) if pd.notna(row['Warehouse']) else ''
            item_wh = QTableWidgetItem(wh)
            item_wh.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(i, 4, item_wh)

            pvp = row['PVP Em Vigor'] if pd.notna(row['PVP Em Vigor']) else 0
            item_pvp = QTableWidgetItem(f"€ {pvp:,.2f}")
            item_pvp.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(i, 5, item_pvp)

            stock = row['Stock Total'] if pd.notna(row['Stock Total']) else 0
            item_stock = QTableWidgetItem(f"{stock:,.0f}")
            item_stock.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(i, 6, item_stock)

            sales = row['Unit Sales'] if pd.notna(row['Unit Sales']) else 0
            item_sales = QTableWidgetItem(f"{sales:,.0f}")
            item_sales.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(i, 7, item_sales)

            valor = row['Sales Value'] if pd.notna(row['Sales Value']) else 0
            item_valor = QTableWidgetItem(f"€ {valor:,.2f}")
            item_valor.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(i, 8, item_valor)

            if self.modo_atual == 'pareto':
                ultimo_valor = row['cum_pct'] if pd.notna(row.get('cum_pct')) else 0
                texto_ultimo = f"{ultimo_valor:.1f}%"
            else:
                ultimo_valor = row['score'] if pd.notna(row.get('score')) else 0
                texto_ultimo = f"{ultimo_valor:.4f}"
            item_ultimo = QTableWidgetItem(texto_ultimo)
            item_ultimo.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(i, 9, item_ultimo)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(8, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(9, QHeaderView.ResizeToContents)

    def exportar_excel(self):
        if self.df_top is None or self.df_top.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados para exportar.")
            return

        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Exportar para Excel", "Top_Artigos.xlsx", "Excel Files (*.xlsx)"
            )
            if not file_path:
                return

            df_export = self.df_top.copy()
            df_export.insert(0, 'Posição', range(1, len(df_export)+1))
            if 'Stock Total' in df_export.columns:
                df_export.rename(columns={'Stock Total': 'Stock'}, inplace=True)

            if self.modo_atual == 'pareto' and 'cum_pct' in df_export.columns:
                df_export.rename(columns={'cum_pct': '% Acumulada'}, inplace=True)
            elif 'score' in df_export.columns:
                df_export.rename(columns={'score': 'Score'}, inplace=True)

            for col in ['norm_sales', 'norm_value', 'norm_price', 'norm_sku', 'Sku_num',
                        'Stock In Transit', 'Stock Expected', 'Stock On Order']:
                if col in df_export.columns:
                    df_export.drop(columns=[col], inplace=True)

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Top Artigos')
                worksheet = writer.sheets['Top Artigos']
                for col in worksheet.columns:
                    max_len = max(len(str(cell.value)) for cell in col)
                    worksheet.column_dimensions[col[0].column_letter].width = min(max_len+2, 40)

            QMessageBox.information(self, "Sucesso", f"Excel exportado: {os.path.basename(file_path)}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar: {str(e)}")

    def exportar_pdf(self):
        if self.df_top is None or self.df_top.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados para exportar.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Exportar para PDF", "Top_Artigos.pdf", "PDF (*.pdf)"
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
            cursor.insertText("TOP N ARTIGOS\n\n")

            info = f"Total: {len(self.df_top):,} | Gerado em: {pd.Timestamp.now():%d/%m/%Y %H:%M}\n\n"
            cursor.insertText(info)

            headers = ['Pos.', 'Sku', 'Description', 'Cat.', 'Wh.', 'PVP', 'Stock', 'Vendas', 'Valor',
                       '% Acum.' if self.modo_atual == 'pareto' else 'Score']
            larguras = [4, 8, 24, 7, 6, 7, 7, 8, 10, 7]

            table_fmt = QTextTableFormat()
            table_fmt.setWidth(QTextLength(QTextLength.PercentageLength, 100))
            table_fmt.setCellPadding(4)
            table_fmt.setCellSpacing(0)
            table_fmt.setBorder(0.5)
            constraints = [QTextLength(QTextLength.PercentageLength, w) for w in larguras]
            table_fmt.setColumnWidthConstraints(constraints)

            table = cursor.insertTable(len(self.df_top)+1, len(headers), table_fmt)

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

            for row_idx, (_, row) in enumerate(self.df_top.iterrows(), start=1):
                if self.modo_atual == 'pareto':
                    ultima_val = f"{row['cum_pct']:.1f}%" if pd.notna(row.get('cum_pct')) else '0%'
                else:
                    ultima_val = f"{row['score']:.4f}" if pd.notna(row.get('score')) else '0'

                valores = [
                    str(row_idx),
                    str(row['Sku']),
                    str(row['Description'])[:35] if pd.notna(row['Description']) else '',
                    str(row['Categoria']) if pd.notna(row['Categoria']) else '',
                    str(row['Warehouse']) if pd.notna(row.get('Warehouse')) else '',
                    f"€{row['PVP Em Vigor']:,.2f}" if pd.notna(row['PVP Em Vigor']) else '€0',
                    f"{row['Stock Total']:,.0f}" if pd.notna(row.get('Stock Total')) else '0',
                    f"{int(row['Unit Sales']):,}" if pd.notna(row['Unit Sales']) else '0',
                    f"€{row['Sales Value']:,.2f}" if pd.notna(row['Sales Value']) else '€0',
                    ultima_val
                ]
                for col_idx, text in enumerate(valores):
                    cell = table.cellAt(row_idx, col_idx)
                    cur = cell.firstCursorPosition()
                    cur.insertText(text, normal_fmt)

            doc.print_(printer)
            QMessageBox.information(self, "Sucesso", f"PDF exportado: {os.path.basename(file_path)}")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar PDF: {str(e)}")


def mostrar_top10():
    if QApplication.instance() is None:
        app = QApplication(sys.argv)
        dialog = TopNDialog()
        dialog.exec_()
    else:
        dialog = TopNDialog()
        dialog.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    dialog = TopNDialog()
    dialog.show()
    sys.exit(app.exec_())