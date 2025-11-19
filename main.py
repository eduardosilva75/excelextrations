#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QFrame, QHBoxLayout, QMessageBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Comparação de Excel")
        self.setGeometry(100, 100, 800, 700)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(40, 40, 40, 40)
        
        # Título
        title = QLabel("Sistema de Comparação de Excel")
        title.setFont(QFont("Arial", 24, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Subtítulo
        subtitle = QLabel("Selecione o módulo desejado")
        subtitle.setFont(QFont("Arial", 12))
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #666; margin-bottom: 20px;")
        layout.addWidget(subtitle)
        
        # Frame para os botões
        buttons_frame = QFrame()
        buttons_layout = QVBoxLayout()
        buttons_layout.setSpacing(15)
        
        # Botão Hit Parade por Secção
        btn_hit_parade = QPushButton("🏆 Hit Parade - Merchorg")
        btn_hit_parade.setFont(QFont("Arial", 14))
        btn_hit_parade.setMinimumHeight(80)
        btn_hit_parade.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        btn_hit_parade.clicked.connect(self.abrir_hit_parade)
        buttons_layout.addWidget(btn_hit_parade)
        
        # Botão Tendências
        btn_tendencias = QPushButton("📈 Tendências - Merchorg")
        btn_tendencias.setFont(QFont("Arial", 14))
        btn_tendencias.setMinimumHeight(80)
        btn_tendencias.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
            QPushButton:pressed {
                background-color: #0960a8;
            }
        """)
        btn_tendencias.clicked.connect(self.abrir_tendencias)
        buttons_layout.addWidget(btn_tendencias)
        
        # Botão Artigos Únicos
        btn_artigos_unicos = QPushButton("🔍 Artigos Únicos - Merchorg vs Daily Sales")
        btn_artigos_unicos.setFont(QFont("Arial", 14))
        btn_artigos_unicos.setMinimumHeight(80)
        btn_artigos_unicos.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #e68900;
            }
            QPushButton:pressed {
                background-color: #cc7a00;
            }
        """)
        btn_artigos_unicos.clicked.connect(self.abrir_artigos_unicos)
        buttons_layout.addWidget(btn_artigos_unicos)
        
        # Botão Artigos sem PS
        btn_artigos_sem_ps = QPushButton("📊 Artigos sem PS - Merchorg")
        btn_artigos_sem_ps.setFont(QFont("Arial", 14))
        btn_artigos_sem_ps.setMinimumHeight(80)
        btn_artigos_sem_ps.setStyleSheet("""
            QPushButton {
                background-color: #9C27B0;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #7B1FA2;
            }
            QPushButton:pressed {
                background-color: #6A1B9A;
            }
        """)
        btn_artigos_sem_ps.clicked.connect(self.abrir_artigos_sem_ps)
        buttons_layout.addWidget(btn_artigos_sem_ps)
        
        # Botão Vendas vs Stocks
        btn_vendas_stocks = QPushButton("📦 Vendas vs Stocks - Daily Sales")
        btn_vendas_stocks.setFont(QFont("Arial", 14))
        btn_vendas_stocks.setMinimumHeight(80)
        btn_vendas_stocks.setStyleSheet("""
            QPushButton {
                background-color: #607D8B;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #546E7A;
            }
            QPushButton:pressed {
                background-color: #455A64;
            }
        """)
        btn_vendas_stocks.clicked.connect(self.abrir_vendas_stocks)
        buttons_layout.addWidget(btn_vendas_stocks)
        
        buttons_frame.setLayout(buttons_layout)
        layout.addWidget(buttons_frame)
        
        # Espaçador
        layout.addStretch()
        
        # Botão Fechar Aplicação
        btn_fechar_layout = QHBoxLayout()
        btn_fechar = QPushButton("🚪 Fechar Aplicação")
        btn_fechar.setFont(QFont("Arial", 12))
        btn_fechar.setMinimumHeight(50)
        btn_fechar.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:pressed {
                background-color: #b71c1c;
            }
        """)
        btn_fechar.clicked.connect(self.fechar_aplicacao)
        btn_fechar_layout.addStretch()
        btn_fechar_layout.addWidget(btn_fechar)
        btn_fechar_layout.addStretch()
        layout.addLayout(btn_fechar_layout)
        
        # Rodapé
        footer = QLabel("© 2025 Sistema de Comparação de Excel")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("color: #999; font-size: 10px;")
        layout.addWidget(footer)
        
        central_widget.setLayout(layout)
        
        # Estilo geral da janela
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
        """)
    
    def abrir_hit_parade(self):
        try:
            from hitParade import mostrar_hit_parade
            mostrar_hit_parade()
        except Exception as e:
            self.mostrar_erro(f"Erro ao abrir Hit Parade: {e}")
    
    def abrir_tendencias(self):
        try:
            from tendencias import mostrar_tendencias
            mostrar_tendencias()
        except Exception as e:
            self.mostrar_erro(f"Erro ao abrir Tendências: {e}")
    
    def abrir_artigos_unicos(self):
        try:
            from artigosUnicos import mostrar_artigos_unicos
            mostrar_artigos_unicos()
        except Exception as e:
            self.mostrar_erro(f"Erro ao abrir Artigos Únicos: {e}")
    
    def abrir_artigos_sem_ps(self):
        try:
            from artigosSemPS import mostrar_artigos_sem_ps
            mostrar_artigos_sem_ps()
        except Exception as e:
            self.mostrar_erro(f"Erro ao abrir Artigos sem PS: {e}")
    
    def abrir_vendas_stocks(self):
        try:
            from vendasVsStocks import mostrar_vendas_stocks
            mostrar_vendas_stocks()
        except Exception as e:
            self.mostrar_erro(f"Erro ao abrir Vendas vs Stocks: {e}")
    
    def mostrar_erro(self, mensagem):
        QMessageBox.critical(self, "Erro", mensagem)
    
    def fechar_aplicacao(self):
        """Fecha a aplicação completamente"""
        reply = QMessageBox.question(self, 'Fechar Aplicação', 
                                   'Tem a certeza que deseja fechar a aplicação?',
                                   QMessageBox.Yes | QMessageBox.No, 
                                   QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            self.close()
            QApplication.quit()

def main():
    app = QApplication(sys.argv)
    # Estilo moderno
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()