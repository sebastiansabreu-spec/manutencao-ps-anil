import os
import sys

# Garante que a raiz do repositório esteja no sys.path para importação do main.py
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import app as fastapi_app

async def app(scope, receive, send):
    if scope.get("type") == "http":
        headers = dict(scope.get("headers", []))
        # Vercel envia o caminho original requisitado no cabeçalho x-matched-path
        matched = headers.get(b"x-matched-path", b"").decode("utf-8")
        if matched:
            scope["path"] = matched
            scope["raw_path"] = matched.encode("utf-8")
        elif scope.get("path", "").startswith("/api/index.py"):
            p = scope["path"][len("/api/index.py"):] or "/"
            scope["path"] = p
            scope["raw_path"] = p.encode("utf-8")
        elif scope.get("path", "").startswith("/api/index"):
            p = scope["path"][len("/api/index"):] or "/"
            scope["path"] = p
            scope["raw_path"] = p.encode("utf-8")
    await fastapi_app(scope, receive, send)
