#!/usr/bin/env python3
"""
Navegação de Suportes AutoCAD

Aplicação para navegação e controle de suportes de tubulação no AutoCAD.
Funciona apenas em Windows com AutoCAD instalado.

Usage:
    python suporte_navegacao.py
"""

import sys
import os

# Adiciona diretório atual ao path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Verifica plataforma
if sys.platform != 'win32':
    print("AVISO: Esta aplicação requer Windows e AutoCAD instalados.")
    print("A interface será carregada, mas a conexão com AutoCAD não funcionará.")
    input("Pressione Enter para continuar mesmo assim...")

# Verifica PySide6
try:
    from PySide6.QtWidgets import QApplication
    from PySide6.QtCore import Qt
except ImportError:
    print("ERRO: PySide6 não está instalado.")
    print("Execute: pip install PySide6")
    sys.exit(1)

# Verifica pywin32 no Windows
if sys.platform == 'win32':
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("ERRO: pywin32 não está instalado.")
        print("Execute: pip install pywin32")
        sys.exit(1)

from gui.main_window import MainWindow


def main():
    """Função principal."""
    # Configuração de alta DPI
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    # Cria aplicação
    app = QApplication(sys.argv)
    app.setApplicationName("Navegação de Suportes AutoCAD")
    app.setOrganizationName("SuportesAutoCAD")
    app.setStyle('Fusion')

    # Cria e mostra janela principal
    window = MainWindow()
    window.show()

    # Executa
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
