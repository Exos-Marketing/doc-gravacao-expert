"""clickup_api.py — REST do ClickUp com o token do Keychain (serviço clickup-exos-api). Nunca imprime o token.

GET repete em 429/5xx/rede. Escrita (PUT) só repete em 429, que o ClickUp recusa sem executar: em 5xx ou queda
de rede quem chamou relê e confere pela chave antes de tentar de novo (append repetido duplicaria o rodapé).
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request

BASE = "https://api.clickup.com"
_TOKEN = None


def token() -> str:
    global _TOKEN
    if not _TOKEN:
        t = os.environ.get("CLICKUP_TOKEN", "").strip() or subprocess.run(
            ["security", "find-generic-password", "-s", "clickup-exos-api", "-w"],
            capture_output=True, text=True).stdout.strip()
        if not t:
            sys.exit("ERRO: token do ClickUp ausente (CLICKUP_TOKEN ou Keychain clickup-exos-api)")
        _TOKEN = t
    return _TOKEN


def api(method: str, path: str, body: dict | None = None, tries: int = 6) -> tuple[int, object]:
    data = json.dumps(body).encode() if body is not None else None
    for a in range(tries):
        req = urllib.request.Request(BASE + path, data=data, method=method)
        req.add_header("Authorization", token())
        req.add_header("Content-Type", "application/json")
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                t = r.read().decode()
                return r.status, (json.loads(t) if t.strip() else {})
        except urllib.error.HTTPError as e:
            if e.code == 429 and a < tries - 1:
                reset = e.headers.get("X-RateLimit-Reset") or ""
                time.sleep(max(2, min(65, int(reset) - int(time.time()) + 1)) if reset.isdigit() else 5 * (a + 1))
                continue
            if e.code in (500, 502, 503, 504) and method == "GET" and a < tries - 1:
                time.sleep(3 * (a + 1))
                continue
            txt = e.read().decode()
            try:
                return e.code, json.loads(txt)
            except ValueError:
                return e.code, txt
        except OSError:
            if method == "GET" and a < tries - 1:
                time.sleep(3)
                continue
            return 0, "sem resposta (rede)"
    return 0, "esgotou as tentativas"


def _ok(c: int, b: object, o_que: str) -> dict:
    if c != 200 or not isinstance(b, dict):
        raise RuntimeError(f"ClickUp {o_que} -> {c}: {str(b)[:300]}")
    return b


def pagina(ws: str, doc: str, page: str) -> dict:
    return _ok(*api("GET", f"/api/v3/workspaces/{ws}/docs/{doc}/pages/{page}?content_format=text/md"), "GET página")


def pagina_acrescentar(ws: str, doc: str, page: str, md: str) -> None:
    _ok(*api("PUT", f"/api/v3/workspaces/{ws}/docs/{doc}/pages/{page}",
             {"content": md, "content_format": "text/md", "content_edit_mode": "append"}), "PUT página (append)")


def tarefa(tid: str) -> dict:
    return _ok(*api("GET", f"/api/v2/task/{tid}?include_markdown_description=true"), "GET task")


def tarefa_trocar_descricao(tid: str, md: str) -> None:
    _ok(*api("PUT", f"/api/v2/task/{tid}", {"markdown_content": md}), "PUT task (descrição)")
