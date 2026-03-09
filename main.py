#!/usr/bin/env python3
"""
main.py — ponto de entrada da aplicação.
Execute: python main.py
"""

import sys
import os

# Garante que o diretório do projeto está no path,
# independente de onde o script for chamado
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import customtkinter as ctk
from config.settings import CTK_APPEARANCE, CTK_COLOR_THEME


def main():
    try:
        import pandas  # noqa: F401
        import openpyxl  # noqa: F401
    except ImportError as e:
        print(f"Dependência ausente: {e}")
        print("Instale com:  pip install pandas openpyxl customtkinter")
        input("Pressione Enter para sair…")
        sys.exit(1)

    ctk.set_appearance_mode(CTK_APPEARANCE)
    ctk.set_default_color_theme(CTK_COLOR_THEME)

    try:
        from ui.main_window import AplicacaoLeitorXML
        app = AplicacaoLeitorXML()
        app.run()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input(f"\nErro inesperado: {e}\nPressione Enter para sair…")


if __name__ == "__main__":
    main()
    