import os
import sys

# Garante que a raiz do repositório esteja no sys.path para importação do main.py
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import app
