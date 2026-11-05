"""drive_api.py — Drive API v3 com o token OAuth do rclone (remote `gdrive`, conta drive@exosmkt.com).

Vendorizado de Exos/drive_docs_collab_dm.py e drive_docs_gravacao_et_fls.py para a skill ser autocontida.
A Google Docs API dá 403 nesse projeto OAuth: tudo aqui é Drive API, e o conteúdo do doc entra por DOCX.
Nunca imprime o token.
"""
from __future__ import annotations

import configparser
import json
import random
import subprocess
import sys
import time
import unicodedata
import uuid
from pathlib import Path

import requests

API = "https://www.googleapis.com/drive/v3"
UPLOAD_API = "https://www.googleapis.com/upload/drive/v3"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
GDOC_MIME = "application/vnd.google-apps.document"
FOLDER_MIME = "application/vnd.google-apps.folder"


def rclone_token() -> str:
    """Renova o access token via rclone e lê o conf."""
    subprocess.run(["rclone", "lsd", "gdrive:", "-q", "--max-depth", "1"], capture_output=True, timeout=90)
    cp = configparser.RawConfigParser()
    cp.read(Path.home() / ".config" / "rclone" / "rclone.conf")
    try:
        return json.loads(cp["gdrive"]["token"])["access_token"]
    except Exception as e:  # noqa: BLE001
        sys.exit(f"ERRO: sem token do rclone no remote gdrive ({type(e).__name__}). "
                 "Reconecte com `rclone config reconnect gdrive:` e rode de novo.")


def norm_nome(s: str) -> str:
    """Compara nomes de pasta/doc sem cair em NFD do macOS, caixa ou espaço duplicado."""
    return " ".join(unicodedata.normalize("NFC", s or "").casefold().split())


def folder_link(fid: str) -> str:
    return f"https://drive.google.com/drive/folders/{fid}?usp=sharing"


def doc_link(fid: str) -> str:
    return f"https://docs.google.com/document/d/{fid}/edit?usp=sharing"


class Drive:
    def __init__(self):
        self.s = requests.Session()
        self.drive_id: str | None = None
        self._auth()

    def _auth(self):
        self.s.headers["Authorization"] = "Bearer " + rclone_token()

    def call(self, method: str, url: str, retries: int = 6, criar: bool = False, **kw) -> requests.Response:
        """criar=True: 5xx não repete às cegas (o POST pode ter dado certo); quem chamou reconfere pelo nome."""
        params = kw.pop("params", {}) or {}
        params.setdefault("supportsAllDrives", "true")
        for attempt in range(retries):
            r = self.s.request(method, url, params=params, timeout=120, **kw)
            if r.ok:
                return r
            if r.status_code == 401 and attempt == 0:
                self._auth()
                continue
            cota = r.status_code == 429 or (r.status_code == 403 and ("ateLimit" in r.text or "quota" in r.text.lower()))
            servidor = r.status_code in (500, 502, 503) and not criar
            if (cota or servidor) and attempt < retries - 1:
                wait = 2 ** attempt + random.random()
                print(f"    ⚠ {r.status_code} transitório, tentando de novo em {wait:.1f}s", file=sys.stderr)
                time.sleep(wait)
                continue
            raise RuntimeError(f"{method} {url.split('?')[0]} -> {r.status_code}: {r.text[:400]}")
        raise RuntimeError("unreachable")

    # ---- leitura
    def meta(self, file_id: str, fields: str = "id,name,mimeType,parents,trashed,driveId") -> dict:
        return self.call("GET", f"{API}/files/{file_id}", params={"fields": fields}).json()

    def filhos(self, parent_id: str, mime: str | None = None) -> list[dict]:
        q = f"'{parent_id}' in parents and trashed = false" + (f" and mimeType = '{mime}'" if mime else "")
        out, token = [], None
        while True:
            params = {"q": q, "includeItemsFromAllDrives": "true", "pageSize": 1000,
                      "fields": "nextPageToken,incompleteSearch,files(id,name,mimeType,modifiedTime)"}
            if self.drive_id:
                params.update(corpora="drive", driveId=self.drive_id)
            else:
                params.update(corpora="allDrives")
            if token:
                params["pageToken"] = token
            j = self.call("GET", f"{API}/files", params=params).json()
            if j.get("incompleteSearch"):
                raise RuntimeError(f"busca incompleta nos filhos de {parent_id}")
            out += j.get("files", [])
            token = j.get("nextPageToken")
            if not token:
                return out

    def achar(self, parent_id: str, nome: str, mime: str) -> dict | None:
        """Filho com o nome (normalizado). Mais de um = erro: resolver à mão, nunca escolher calado."""
        achados = [f for f in self.filhos(parent_id, mime) if norm_nome(f["name"]) == norm_nome(nome)]
        if len(achados) > 1:
            raise RuntimeError(f"'{nome}' aparece {len(achados)}× em {parent_id}: {[f['id'] for f in achados]} — resolver à mão")
        return achados[0] if achados else None

    def exportar(self, file_id: str, mime: str) -> bytes:
        return self.call("GET", f"{API}/files/{file_id}/export", params={"mimeType": mime}).content

    # ---- escrita
    def criar_pasta(self, parent_id: str, nome: str) -> dict:
        return self.call("POST", f"{API}/files", criar=True, params={"fields": "id,name"},
                         json={"name": nome, "mimeType": FOLDER_MIME, "parents": [parent_id]}).json()

    def subir_docx_como_gdoc(self, docx: bytes, nome: str, parent_id: str) -> dict:
        """multipart/related montado à mão (o `files=` do requests vira form-data e o Drive recusa)."""
        meta = {"name": nome, "mimeType": GDOC_MIME, "parents": [parent_id]}
        b = "exos_" + uuid.uuid4().hex
        body = (f"--{b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{json.dumps(meta, ensure_ascii=False)}\r\n"
                f"--{b}\r\nContent-Type: {DOCX_MIME}\r\n\r\n").encode("utf-8") + docx + f"\r\n--{b}--".encode("utf-8")
        return self.call("POST", f"{UPLOAD_API}/files", criar=True, params={"uploadType": "multipart", "fields": "id,name"},
                         data=body, headers={"Content-Type": f"multipart/related; boundary={b}"}).json()

    def trocar_conteudo(self, file_id: str, docx: bytes) -> dict:
        """Troca o conteúdo do Google Doc pelo DOCX: mesmo ID, mesmo link, mesmas permissões."""
        return self.call("PATCH", f"{UPLOAD_API}/files/{file_id}", params={"uploadType": "media", "fields": "id,name"},
                         data=docx, headers={"Content-Type": DOCX_MIME}).json()

    # ---- permissões
    def permissoes(self, file_id: str) -> list[dict]:
        return self.call("GET", f"{API}/files/{file_id}/permissions",
                         params={"fields": "permissions(id,type,role,domain)"}).json().get("permissions", [])

    def publicos(self, file_id: str) -> list[tuple[str, str]]:
        """Acessos por link (anyone) ou por domínio, herdados ou diretos."""
        return [(p["type"], p["role"]) for p in self.permissoes(file_id) if p["type"] in ("anyone", "domain")]

    def papel_publico(self, file_id: str) -> str | None:
        return next((r for t, r in self.publicos(file_id) if t == "anyone"), None)

    def dar_publico(self, file_id: str, role: str) -> None:
        self.call("POST", f"{API}/files/{file_id}/permissions",
                  json={"type": "anyone", "role": role, "allowFileDiscovery": False})


def acesso_anonimo(url: str) -> tuple[int, str, bytes]:
    """GET sem sessão nem token: (status, url final, início do corpo). Login no meio = não é público."""
    r = requests.get(url, allow_redirects=True, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
    return r.status_code, r.url, r.content[:8]
