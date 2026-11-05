#!/usr/bin/env python3
# /// script
# requires-python = ">=3.12"
# dependencies = ["python-docx==1.2.0", "requests", "pymupdf"]
# ///
"""
doc_gravacao.py — motor da skill doc-gravacao-expert.

Cria o Documento de Gravação da expert de um projeto Exos a partir da copy no ClickUp (lida ao vivo):
  - Google Doc com o logo Exos no cabeçalho, título, "DRIVE para subir arquivos:" + link da pasta de upload,
    ORIENTAÇÕES DE GRAVAÇÃO e os roteiros (ADS NN), em `1.Copy` (qualquer pessoa com o link = editor);
  - pasta de upload em `14.Arquivo Bruto › <categoria> › <pasta_upload>` (só a subpasta abre por link);
  - registro no ClickUp: rodapé "Documento Gravação Expert" na página da copy e 2 linhas na descrição da task.
Configuração por projeto: <Expert>/06-Copy/gravacao-<SIGLA>/config.json (estado.json e links.md ao lado).

Uso (dry-run por padrão; nada é escrito sem --apply):
  SK=~/.claude/skills/doc-gravacao-expert/scripts
  uv run --python 3.12 $SK/doc_gravacao.py gerar --config CFG [--preview]
  uv run --python 3.12 $SK/doc_gravacao.py gerar --config CFG --docx-only --preview   # só local
  uv run --python 3.12 $SK/doc_gravacao.py gerar --config CFG --apply [--sem-clickup]
  uv run --python 3.12 $SK/doc_gravacao.py gerar --config CFG --apply --update [--forcar]
  uv run --python 3.12 $SK/doc_gravacao.py verificar --config CFG
  uv run --python 3.12 $SK/doc_gravacao.py modelo [--doc-id ID]                        # recongela o modelo
  uv run --python 3.12 $SK/doc_gravacao.py listar

Saída: 0 ok · 1 abortou/verificação falhou · 2 a copy mudou e o doc não foi atualizado · 3 doc editado à mão.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import io
import json
import re
import shutil
import subprocess
import sys
import time
import unicodedata
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Pt

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import clickup_api as cu  # noqa: E402
import drive_api as dapi  # noqa: E402
from drive_api import doc_link, folder_link, norm_nome  # noqa: E402,F401

ASSET = HERE.parent / "assets" / "modelo_gravacao.docx"
MANIFESTO = HERE.parent / "assets" / "modelo_gravacao.json"
MODELO_PADRAO = "143AxLubVTWMZPjfOLeuc1-1DtYqJNcebeql42RCFW4A"     # doc feito à mão na Black da Elen (logo Exos)
EXOS = Path.home() / "Documents" / "Exos"
SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
DRIVE_ROTULO = "DRIVE para subir arquivos:"
RODAPE_ROTULO = "Documento Gravação Expert:"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
VML, O = "{urn:schemas-microsoft-com:vml}", "{urn:schemas-microsoft-com:office:office}"
PARA_ID = "{http://schemas.microsoft.com/office/word/2010/wordml}paraId"


# ----------------------------------------------------------------------------- markdown do ClickUp
UNESC = re.compile(r"\\([\\`*_{}\[\]()#+\-.!=|>~\"'])")
BOLD = re.compile(r"(?<!\\)\*\*(.+?)(?<!\\)\*\*")
SOBRA_MD = re.compile(r"(?<!\\)[*_`]")
ESTRUTURA = re.compile(r"^(#|>|\||!\[|[*+-]\s|\d+\.\s)")
BASE_PROIBIDO = r"https?://|clickup|\]\(|\*\*|\\[\[\]_*]|\[inserir"


def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s or "")


def unesc(s: str) -> str:
    return UNESC.sub(r"\1", s)


def _solto(linha: str) -> str:
    """Texto de uma linha qualquer da página, sem negrito nem escapes (para âncoras; não valida marcação)."""
    return nfc(unesc(BOLD.sub(r"\1", linha))).strip()


def runs(linha: str) -> list[tuple[str, bool]]:
    """Negrito do markdown → [(texto, negrito)]. Qualquer outra marcação aborta."""
    out, pos = [], 0
    for m in BOLD.finditer(linha):
        if m.start() > pos:
            out.append((linha[pos:m.start()], False))
        out.append((m.group(1), True))
        pos = m.end()
    if pos < len(linha):
        out.append((linha[pos:], False))
    if any(SOBRA_MD.search(t) for t, _ in out):
        raise ValueError(f"marcação não prevista na linha: {linha[:90]!r}")
    return [(nfc(unesc(t)), b) for t, b in out]


def plain(rs) -> str:
    return "".join(t for t, _ in rs)


def contem(md: str, chave: str) -> bool:
    """O ClickUp escapa `_` no texto do link (1\\_oSNf…): procurar com e sem escape."""
    return chave in md or chave in unesc(md)


def _trim(rs: list) -> list:
    rs = [[t, b] for t, b in rs if t]
    if rs:
        rs[0][0], rs[-1][0] = rs[0][0].lstrip(), rs[-1][0].rstrip()
    return [(t, b) for t, b in rs if t]


# ----------------------------------------------------------------------------- leitura da copy
def secao(md: str, inicio: str, fim: str) -> list[str]:
    """Linhas entre a âncora de início (1 linha, casamento total) e a 1ª linha de fim depois dela."""
    linhas = [l.strip() for l in md.replace("\r\n", "\n").split("\n")]
    ini, f = re.compile(inicio), re.compile(fim)
    achou = [i for i, l in enumerate(linhas) if l and ini.fullmatch(_solto(l))]
    if len(achou) != 1:
        raise ValueError(f"âncora de início {inicio!r}: achei {len(achou)}× na página (precisa ser 1)")
    for j in range(achou[0] + 1, len(linhas)):
        if linhas[j] and f.fullmatch(_solto(linhas[j])):
            return linhas[achou[0] + 1:j]
    raise ValueError(f"âncora de fim {fim!r} não encontrada depois do início")


def adaptar(md: str, regras: list[dict]) -> tuple[str, list[str]]:
    log = []
    for r in regras:
        md, k = re.subn(r["padrao"], r.get("troca", ""), md, flags=re.M)
        if k != r["vezes"]:
            raise ValueError(f"ajuste '{r['rotulo']}': casou {k}×, esperava {r['vezes']}× — a copy mudou, revise o config")
        log.append(f"ajuste '{r['rotulo']}': {k}×")
    return md, log


def _dividir(rs: list, rotulos: list[str], principal: str) -> list[list]:
    """Quebra a linha nos rótulos de `dividir_em` (todos, na ordem). Outro rótulo em negrito aborta."""
    partes, atual, faltam = [], [], list(rotulos)
    for i, (t, b) in enumerate(rs):
        tt = t.strip()
        if b and i > 0 and tt.endswith(":"):
            if faltam and tt == faltam[0]:
                partes.append(_trim(atual))
                atual = []
                faltam.pop(0)
            else:
                raise ValueError(f"rótulo {tt!r} dentro de '{principal}' não está no config (dividir_em)")
        atual.append((t, b))
    if faltam:
        raise ValueError(f"'{principal}': faltou o rótulo {faltam} para dividir")
    return partes + [_trim(atual)]


def rotulo_do_texto(rot: str, texto_rot: str) -> bool:
    """'Texto para gravação:' ou com complemento entre parênteses ('Texto para gravação (guia):')."""
    return rot == texto_rot or bool(re.fullmatch(re.escape(texto_rot[:-1]) + r" \(.+\):", rot))


def ler_copia(md: str, cc: dict) -> dict:
    """md da página → modelo do doc. Allow-list: qualquer linha fora do previsto aborta (nada vai calado)."""
    sec, fora = adaptar("\n".join(secao(md, cc["inicio"], cc["fim"])), cc.get("adaptar", []))
    it = cc["itens"]
    titulo_re = re.compile(it["titulo"])
    descartar = [nfc(x) for x in cc.get("descartar", [])]
    campos_item = [nfc(x) for x in it.get("manter", [])]
    texto_rot = nfc(it.get("texto", ""))
    paragrafos, itens, vistos, cta = [], [], {}, None
    atual, em_texto = None, False

    for linha in filter(None, (l.strip() for l in sec.split("\n"))):
        if ESTRUTURA.match(linha):
            raise ValueError(f"estrutura não prevista (lista, título, tabela ou imagem): {linha[:90]!r}")
        rs = runs(linha)
        txt = plain(rs).strip()
        so_negrito = len(rs) == 1 and rs[0][1]
        rot = rs[0][0].strip() if rs[0][1] and rs[0][0].strip().endswith(":") else None

        if so_negrito and titulo_re.fullmatch(txt):                         # item novo (ADS NN [...])
            atual, em_texto = {"num": titulo_re.fullmatch(txt).group(1), "titulo": txt,
                               "campos": [], "texto_rotulo": None, "falas": []}, False
            itens.append(atual)
            continue
        if rot and any(rot.startswith(x) for x in descartar):               # interno: fica de fora
            fora.append(f"descartado: {txt[:80]}…")
            continue
        if atual is not None:
            if rot and so_negrito and texto_rot and rotulo_do_texto(rot, texto_rot):
                atual["texto_rotulo"], em_texto = rot, True
            elif rot and not em_texto and any(rot.startswith(x) for x in campos_item):
                atual["campos"].append(_trim(rs))
            elif em_texto and not rot:
                atual["falas"].append(_trim(rs))
            else:
                raise ValueError(f"{atual['titulo']}: linha não prevista: {txt[:90]!r}")
            continue
        if rot is None:
            raise ValueError(f"linha sem rótulo antes dos itens: {txt[:90]!r} — decida no config (manter/descartar)")
        regras = [e for e in cc["manter"] if rot.startswith(nfc(e["rotulo"]))]
        if len(regras) != 1:
            raise ValueError(f"rótulo {rot!r} " + ("desconhecido — decida no config se é manter ou descartar"
                                                    if not regras else "casa com mais de uma regra do config"))
        e = regras[0]
        vistos[e["rotulo"]] = vistos.get(e["rotulo"], 0) + 1
        paragrafos += _dividir(rs, [nfc(x) for x in e.get("dividir_em", [])], rot)
        if e.get("cta"):
            cta = "".join(t for t, b in rs if not b).strip().strip('"“”').strip()

    for e in cc["manter"]:
        n = vistos.get(e["rotulo"], 0)
        if (n == 0 and not e.get("opcional")) or n > 1:
            raise ValueError(f"'{e['rotulo']}' apareceu {n}× na copy (esperava 1)")
    nums = [i["num"] for i in itens]
    if nums != list(it["numeros"]):
        raise ValueError(f"itens na copy: {nums}; o config espera {list(it['numeros'])}")
    for i in itens:
        for c in campos_item:
            k = sum(1 for p in i["campos"] if p[0][0].strip().startswith(c))
            if k != 1:
                raise ValueError(f"{i['titulo']}: '{c}' apareceu {k}×")
        if texto_rot and not i["falas"]:
            raise ValueError(f"{i['titulo']}: sem texto para gravação")
        if it.get("termina_no_cta") and cta and plain(i["falas"][-1]).strip() != cta:
            raise ValueError(f"{i['titulo']}: a última fala não é o CTA ({cta!r})")

    m = {"paragrafos": paragrafos, "cta_texto": cta, "itens": itens}
    texto = "\n".join(linhas_esperadas(m, ""))
    achado = re.search(BASE_PROIBIDO + (f"|{cc['proibido']}" if cc.get("proibido") else ""), texto, re.I)
    if achado:
        a, b = max(0, achado.start() - 40), achado.end() + 40
        raise ValueError(f"termo interno no texto que vai para a expert: {achado.group(0)!r} em …{texto[a:b]!r}…")
    m["hash"] = hashlib.sha256(json.dumps(m, ensure_ascii=False, sort_keys=True).encode()).hexdigest()[:12]
    m["fora"] = fora
    return m


def fontes(cfg: dict) -> list[dict]:
    """Copies que formam o doc, na ordem. `fontes` junta várias páginas num doc só (ex.: trailer + vídeos gravados
    na mesma sessão); sem `fontes`, vale o par `clickup.copia` + `copia` de sempre."""
    if cfg.get("fontes"):
        return cfg["fontes"]
    cp = cfg["clickup"]["copia"]
    return [{"nome": "copia", "doc_id": cp["doc_id"], "page_id": cp["page_id"], "copia": cfg["copia"]}]


def ler_fontes(mds: list[str], fs: list[dict]) -> dict:
    """md de cada fonte → um modelo só: as orientações de todas antes dos roteiros, os roteiros na ordem das fontes.
    Uma fonte devolve o modelo de ler_copia intacto (mesmo hash dos estados já gravados)."""
    ms = [ler_copia(md, f["copia"]) for md, f in zip(mds, fs, strict=True)]
    if len(ms) == 1:
        return ms[0]
    nums = [i["num"] for x in ms for i in x["itens"]]
    if len(set(nums)) != len(nums):
        raise ValueError(f"roteiro repetido entre as copies: {nums}")
    m = {"paragrafos": [p for x in ms for p in x["paragrafos"]], "cta_texto": [x["cta_texto"] for x in ms],
         "itens": [i for x in ms for i in x["itens"]]}
    m["hash"] = hashlib.sha256(json.dumps(m, ensure_ascii=False, sort_keys=True).encode()).hexdigest()[:12]
    m["fora"] = [f"{f.get('nome', f['page_id'])} · {x}" for f, xm in zip(fs, ms) for x in xm["fora"]]
    return m


def _alternativas(rx: str) -> list[str]:
    """Separa uma regex nas alternativas do nível de cima (o `|` dentro de grupo ou classe não conta)."""
    out, atual, nivel, esc = [], "", 0, False
    for ch in rx:
        if esc:
            atual, esc = atual + ch, False
            continue
        if ch == "\\":
            atual, esc = atual + ch, True
            continue
        if ch in "([":
            nivel += 1
        elif ch in ")]":
            nivel -= 1
        if ch == "|" and nivel == 0:
            out, atual = out + [atual], ""
            continue
        atual += ch
    return [a for a in out + [atual] if a]


def proibido_re(cfg: dict) -> str:
    """Termos barrados no doc inteiro: a base + o que TODAS as fontes barram. Cada fonte já foi conferida com a sua
    lista na leitura, e uma pode liberar o que outra barra (o "Clube da PPK" falado nos áudios da reta final)."""
    listas = [_alternativas(f["copia"].get("proibido") or "") for f in fontes(cfg)]
    comuns = [t for t in listas[0] if all(t in x for x in listas[1:])]
    return BASE_PROIBIDO + "".join(f"|{t}" for t in comuns)


def ler_paginas(cfg: dict, mds_locais: list[str] | None = None) -> dict:
    """Lê as copies (ClickUp ao vivo ou, com --md, um arquivo local por fonte) e devolve o modelo do doc."""
    fs = fontes(cfg)
    if mds_locais:
        if len(mds_locais) != len(fs):
            raise ValueError(f"--md: {len(mds_locais)} arquivo(s) para {len(fs)} fonte(s)")
        return ler_fontes([Path(p).read_text(encoding="utf-8") for p in mds_locais], fs)
    mds = []
    for f in fs:
        pg = cu.pagina(cfg["clickup"]["workspace_id"], f["doc_id"], f["page_id"])
        if cfg["projeto"] not in (pg.get("name") or ""):
            print(f"   ⚠ a página da copy se chama {pg.get('name')!r} (sem {cfg['projeto']})")
        mds.append(pg.get("content") or "")
    return ler_fontes(mds, fs)


def linhas_esperadas(m: dict, titulo_orientacoes: str) -> list[str]:
    """Textos do corpo do doc, na ordem."""
    out = [titulo_orientacoes] + [plain(p) for p in m["paragrafos"]]
    for i in m["itens"]:
        out += [i["titulo"], *(plain(c) for c in i["campos"])]
        if i.get("texto_rotulo"):
            out.append(i["texto_rotulo"])
        out += [plain(f) for f in i["falas"]]
    return out


# ----------------------------------------------------------------------------- DOCX
def ptext(el) -> str:
    return "".join(t.text or "" for t in el.iter(W + "t"))


def pstyle(el) -> str:
    s = el.find(f"{W}pPr/{W}pStyle")
    return s.get(W + "val") if s is not None else ""


def is_hr(p) -> bool:
    return any(r.get(O + "hr") == "t" for r in p.iter(VML + "rect"))


def localizar_cabecalho(body) -> dict:
    """HR · título · HR · 'DRIVE para subir arquivos:' · link · HR, só parágrafos vazios entre eles."""
    seq = []
    for e in body:
        if e.tag == W + "sectPr":
            break
        if e.tag != W + "p":
            seq.append(("outro", e))
        elif is_hr(e):
            seq.append(("hr", e))
        elif ptext(e).strip():
            seq.append(("txt", e))
        if sum(1 for k, _ in seq if k == "hr") == 3:
            break
    kinds = [k for k, _ in seq]
    if kinds != ["hr", "txt", "hr", "txt", "txt", "hr"]:
        raise RuntimeError(f"modelo fora do padrão no cabeçalho (HR, título, HR, DRIVE, link, HR): {kinds}")
    els = [e for _, e in seq]
    if ptext(els[3]).strip() != DRIVE_ROTULO:
        raise RuntimeError(f"modelo: esperava {DRIVE_ROTULO!r}, achei {ptext(els[3]).strip()!r}")
    if len(els[4].findall(W + "hyperlink")) != 1:
        raise RuntimeError("modelo: o parágrafo do link precisa ter exatamente 1 hyperlink")
    return {"hr1": els[0], "titulo": els[1], "hr2": els[2], "drive": els[3], "link": els[4], "hr3": els[5]}


def _trocar_texto(p, texto: str) -> None:
    ts = list(p.iter(W + "t"))
    ts[0].text = texto
    for t in ts[1:]:
        t.text = ""


def _cortar_depois(body, hr3) -> None:
    """Remove tudo depois do 3º HR, menos o 1º parágrafo vazio (o respiro antes do corpo)."""
    resto = [e for e in hr3.itersiblings() if e.tag != W + "sectPr"]
    fica = resto[0] if resto and resto[0].tag == W + "p" and not ptext(resto[0]).strip() and not is_hr(resto[0]) else None
    for e in resto:
        if e is not fica:
            body.remove(e)


def montar_docx(asset: Path, titulo: str, url_pasta: str, m: dict, titulo_orientacoes: str, out: Path) -> dict:
    doc = Document(str(asset))
    body = doc.element.body
    cab = localizar_cabecalho(body)
    _trocar_texto(cab["titulo"], titulo)
    hl = cab["link"].find(W + "hyperlink")
    doc.part.rels[hl.get(R + "id")]._target = url_pasta
    _trocar_texto(hl, url_pasta)
    _cortar_depois(body, cab["hr3"])

    def para(rs=(), style=None, size=None, after=6):
        p = doc.add_paragraph(style=style)               # entra antes do sectPr, em ordem
        for text, bold in rs:
            r = p.add_run(text)
            if bold:
                r.bold = True                            # os headings do Google não vêm em negrito
            if size:
                r.font.size = Pt(size)
        p.paragraph_format.space_after = Pt(after)
        return p

    para([(titulo_orientacoes, True)], style="Heading 2")
    for p in m["paragrafos"]:
        para(p)
    for i in m["itens"]:
        para([(i["titulo"], True)], style="Heading 3")
        for c in i["campos"]:
            para(c)
        if i.get("texto_rotulo"):
            para([(i["texto_rotulo"], True)])
        for f in i["falas"]:
            para(f, size=12, after=8)
    para()
    hr = copy.deepcopy(cab["hr3"])
    hr.attrib.pop(PARA_ID, None)
    body.find(W + "sectPr").addprevious(hr)
    doc.save(str(out))
    st = docx_stats(Path(out).read_bytes())
    if (st["hr"], st["h2"], st["h3"]) != (4, 1, len(m["itens"])) or not st["logo"]:
        raise RuntimeError(f"docx montado fora do esperado: {st} (esperava HR=4, H2=1, H3={len(m['itens'])}, logo)")
    return st


def docx_stats(blob: bytes) -> dict:
    z = zipfile.ZipFile(io.BytesIO(blob))
    nomes = z.namelist()
    ps = [e for e in Document(io.BytesIO(blob)).element.body if e.tag == W + "p"]
    estilos = [pstyle(p) for p in ps]
    logo_sha = None
    for h in (n for n in nomes if re.fullmatch(r"word/header\d*\.xml", n)):
        rels = f"word/_rels/{h.split('/')[-1]}.rels"
        if rels in nomes:
            for alvo in re.findall(rb'Type="[^"]*/image"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Type="[^"]*/image"', z.read(rels)):
                alvo = next(a for a in alvo if a).decode()
                caminho = "word/" + alvo.lstrip("./") if not alvo.startswith("/") else alvo.lstrip("/")
                if caminho in nomes:
                    logo_sha = hashlib.sha256(z.read(caminho)).hexdigest()[:16]
    return {"hr": sum(map(is_hr, ps)), "h2": estilos.count("Heading2"), "h3": estilos.count("Heading3"),
            "logo": logo_sha is not None, "logo_sha": logo_sha,
            "comentarios": any(n.startswith("word/comments") for n in nomes)}


# ----------------------------------------------------------------------------- ClickUp (texto)
def rodape_copia(url: str) -> str:
    return f"\n\n**{RODAPE_ROTULO}** [{url}]({url})"


def tarefas(cfg: dict) -> list[dict]:
    """Tasks que recebem as linhas: `clickup.tarefas` (várias) ou a `clickup.tarefa` de sempre."""
    cc = cfg["clickup"]
    return cc.get("tarefas") or ([cc["tarefa"]] if cc.get("tarefa") else [])


def linhas_tarefa(cfg: dict, tarefa: dict, doc_id: str, upload_id: str) -> list[tuple[str, str]]:
    """[(chave, texto)] da task: pasta de upload e doc, ou só o que `linhas` pedir. O doc vai com o nome, porque
    uma task pode receber mais de um doc de gravação (ex.: a copy da captação de uma Black)."""
    doc_url, up_url = doc_link(doc_id), folder_link(upload_id)
    nome = re.sub(r"([\[\]])", r"\\\1", cfg["doc"]["nome"])
    todas = {"upload": (upload_id, f"Vídeos brutos da {cfg['expert']['curto']} (upload): [{up_url}]({up_url})"),
             "doc": (doc_id, f"{RODAPE_ROTULO} {nome} — [{doc_url}]({doc_url})")}
    return [todas[k] for k in tarefa.get("linhas", ["upload", "doc"])]


def inserir_linhas(md: str, apos: str, linhas: list[tuple[str, str]]) -> tuple[str, int, str | None]:
    """linhas = [(chave, texto)]: insere, depois da linha que casa `apos`, só as que ainda não têm a chave."""
    novas = [t for k, t in linhas if not contem(md, k)]
    if not novas:
        return md, 0, None
    bloco = [x for t in novas for x in ("", t)]
    ls = md.split("\n")
    idx = [i for i, l in enumerate(ls) if re.search(apos, _solto(l))]
    if len(idx) == 1:
        return "\n".join(ls[:idx[0] + 1] + bloco + ls[idx[0] + 1:]), len(novas), None
    return (md.rstrip() + "\n" + "\n".join(bloco), len(novas),
            f"linha de referência {apos!r} achada {len(idx)}× — as linhas foram para o fim da descrição")


def descricao_preservada(antes: str, depois: str, linhas: list[tuple[str, str]]) -> bool:
    """A descrição relida é a de antes + as linhas novas (nada mais mudou)? Compara sem os escapes do ClickUp."""
    sem_novas = unesc(depois)
    for _, txt in linhas:
        sem_novas = sem_novas.replace(unesc(txt), "")
    return re.sub(r"\s+", " ", unesc(sem_novas)).strip() == re.sub(r"\s+", " ", unesc(antes)).strip()


# ----------------------------------------------------------------------------- config / estado
def carregar_config(path: Path) -> dict:
    cfg = json.loads(path.read_text(encoding="utf-8"))
    p = cfg["projeto"]
    cfg["drive"]["pasta_upload"] = cfg["drive"]["pasta_upload"].replace("{projeto}", p)
    cfg["doc"]["nome"] = cfg["doc"]["nome"].replace("{projeto}", p)
    cfg["doc"].setdefault("titulo_orientacoes", "ORIENTAÇÕES DE GRAVAÇÃO")
    return cfg


def ler_json(path: Path, padrao):
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else padrao


def gravar_json(path: Path, obj) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")


def agora() -> str:
    return time.strftime("%Y-%m-%d_%H%M")


def texto_sha(txt: str) -> str:
    return hashlib.sha256(re.sub(r"\s+", " ", txt.replace("﻿", "")).strip().encode()).hexdigest()[:16]


def conferir_asset() -> dict:
    if not ASSET.exists() or not MANIFESTO.exists():
        sys.exit("ERRO: modelo congelado ausente. Rode `doc_gravacao.py modelo` uma vez.")
    man = ler_json(MANIFESTO, {})
    if hashlib.sha256(ASSET.read_bytes()).hexdigest() != man.get("sha256"):
        sys.exit("ERRO: assets/modelo_gravacao.docx não bate com o manifesto. Rode `doc_gravacao.py modelo` de novo.")
    return man


# ----------------------------------------------------------------------------- prévia local
def pdf_para_png(pdf: Path, saida: Path, dpi: int = 70) -> list[Path]:
    import pymupdf
    outs = []
    with pymupdf.open(pdf) as doc:
        for i, page in enumerate(doc):
            p = saida / f"{pdf.stem}_p{i + 1}.png"
            page.get_pixmap(dpi=dpi).save(p)
            outs.append(p)
    return outs


def perfil_lo(perfil: Path) -> str:
    """Perfil próprio do LibreOffice (não briga com uma instância aberta). URI codificada: espaço cru derruba o soffice."""
    return f"-env:UserInstallation={perfil.as_uri()}"


def previa(docx: Path, saida: Path) -> list[Path]:
    saida.mkdir(parents=True, exist_ok=True)
    perfil = saida / ".lo-perfil"
    try:
        subprocess.run([SOFFICE, perfil_lo(perfil), "--headless", "--convert-to", "pdf",
                        "--outdir", str(saida), str(docx)], capture_output=True, timeout=240)
    finally:
        shutil.rmtree(perfil, ignore_errors=True)
    pdf = saida / (docx.stem + ".pdf")
    if not pdf.exists():
        raise RuntimeError("LibreOffice não gerou o PDF da prévia")
    return pdf_para_png(pdf, saida)


# ----------------------------------------------------------------------------- Drive (orquestração)
def _pasta_do_projeto(d: dapi.Drive, cfg: dict, chave: str) -> dict:
    """1.Copy / 14.Arquivo Bruto: pelo ID do config (conferido) ou, sem ID, pelo prefixo do nome."""
    alvo, proj = cfg["drive"][chave], cfg["drive"]["projeto_id"]
    if alvo.get("id"):
        f = d.meta(alvo["id"])
        if f["mimeType"] != dapi.FOLDER_MIME or f.get("trashed") or proj not in f.get("parents", []) \
                or not nfc(f["name"]).startswith(alvo["prefixo"]):
            raise RuntimeError(f"pasta {chave} ({alvo['id']}) não confere: {f.get('name')!r}, pais {f.get('parents')}")
        return f
    achados = [f for f in d.filhos(proj, dapi.FOLDER_MIME) if nfc(f["name"]).startswith(alvo["prefixo"])]
    if len(achados) != 1:
        raise RuntimeError(f"pasta com prefixo {alvo['prefixo']!r} no projeto: {len(achados)} (precisa ser 1)")
    return achados[0]


def _pelo_estado(d: dapi.Drive, fid: str | None, pai: str, rotulo: str) -> dict | None:
    if not fid:
        return None
    f = d.meta(fid)
    if f.get("trashed") or pai not in f.get("parents", []):
        raise RuntimeError(f"{rotulo} do estado.json ({fid}) está na lixeira ou fora do lugar — resolva à mão")
    return f


def garantir_pasta(d: dapi.Drive, pai: str, nome: str, fid: str | None, apply: bool, rotulo: str) -> dict | None:
    f = _pelo_estado(d, fid, pai, rotulo) or d.achar(pai, nome, dapi.FOLDER_MIME)
    if f:
        aviso = "" if nfc(f["name"]) == nfc(nome) else f" (nome no Drive: {f['name']!r})"
        print(f"   {rotulo}: existe{aviso} · {f['id']}")
        return f
    if not apply:
        print(f"   [dry] {rotulo}: criaria '{nome}'")
        return None
    try:
        f = d.criar_pasta(pai, nome)
    except Exception as e:  # noqa: BLE001 — 5xx/timeout: pode ter criado; reconfere antes de repetir
        print(f"   ⚠ criar '{nome}' falhou ({str(e)[:80]}); reconferindo pelo nome")
        f = d.achar(pai, nome, dapi.FOLDER_MIME) or d.criar_pasta(pai, nome)
    print(f"   {rotulo}: criada · {f['id']}")
    return f


def garantir_publico(d: dapi.Drive, fid: str, papel: str, rotulo: str, apply: bool) -> bool:
    atual = d.papel_publico(fid)
    if atual == papel:
        print(f"   {rotulo}: qualquer pessoa com o link = {papel} (já)")
        return True
    if not apply:
        print(f"   [dry] {rotulo}: daria qualquer pessoa com o link = {papel} (hoje: {atual})")
        return True
    d.dar_publico(fid, papel)
    ok = d.papel_publico(fid) == papel
    print(f"   {rotulo}: qualquer pessoa com o link = {papel} {'✔' if ok else '✘ (não pegou)'}")
    return ok


# ----------------------------------------------------------------------------- verificação
def verificar_doc(d: dapi.Drive, doc_id: str, cfg: dict, m: dict, url_pasta: str) -> tuple[list[str], str]:
    errs = []
    txt = d.exportar(doc_id, "text/plain").decode("utf-8-sig")
    norm = re.sub(r"\s+", " ", txt)
    pos = 0
    for ln in [cfg["doc"]["nome"], DRIVE_ROTULO, url_pasta, *linhas_esperadas(m, cfg["doc"]["titulo_orientacoes"])]:
        i = norm.find(re.sub(r"\s+", " ", ln).strip(), pos)
        if i < 0:
            errs.append(f"texto ausente ou fora de ordem no doc: {ln[:70]!r}")
        else:
            pos = i + len(ln)
    resto = norm.replace(url_pasta, "")
    achado = re.search(proibido_re(cfg), resto, re.I)
    if achado:
        errs.append(f"termo interno no doc: {achado.group(0)!r}")
    st = docx_stats(d.exportar(doc_id, dapi.DOCX_MIME))
    if (st["hr"], st["h2"], st["h3"]) != (4, 1, len(m["itens"])) or not st["logo"]:
        errs.append(f"estrutura do doc no Google: {st}")
    print(f"   texto {'ok' if not any('texto' in e for e in errs) else 'FALHOU'} · estrutura HR={st['hr']} "
          f"H2={st['h2']} H3={st['h3']} logo={'sim' if st['logo'] else 'NÃO'}")
    return errs, texto_sha(txt)


def verificar_acessos(d: dapi.Drive, cfg: dict, doc_id: str, upload_id: str, categoria_id: str, saida: Path) -> list[str]:
    errs = []
    esperado = {doc_id: cfg["drive"]["papel_doc"], upload_id: cfg["drive"]["papel_pasta_upload"]}
    for fid, papel in esperado.items():
        if d.papel_publico(fid) != papel:
            errs.append(f"{fid}: acesso por link {d.papel_publico(fid)} (esperava {papel})")
    if d.publicos(categoria_id):
        errs.append(f"categoria {categoria_id} tem acesso público {d.publicos(categoria_id)} — vaza para as subpastas")
    for url in (doc_link(doc_id), folder_link(upload_id)):
        st, final, _ = dapi.acesso_anonimo(url)
        if st != 200 or "accounts.google.com" in final or "ServiceLogin" in final:
            errs.append(f"sem acesso anônimo: {url} -> {st} {final[:80]}")
    st, _, ini = dapi.acesso_anonimo(f"https://docs.google.com/document/d/{doc_id}/export?format=pdf")
    if st == 200 and ini.startswith(b"%PDF"):
        import requests
        pdf = saida / "doc_google.pdf"
        saida.mkdir(parents=True, exist_ok=True)
        pdf.write_bytes(requests.get(f"https://docs.google.com/document/d/{doc_id}/export?format=pdf", timeout=60).content)
        print(f"   prévia do doc real: {', '.join(p.name for p in pdf_para_png(pdf, saida))} em {saida}")
    else:
        errs.append(f"export PDF anônimo do doc falhou ({st})")
    print(f"   acessos {'ok' if not errs else 'FALHARAM'}")
    return errs


# ----------------------------------------------------------------------------- ClickUp (escrita)
def linkar_clickup(cfg: dict, doc_id: str, upload_id: str, pasta: Path, apply: bool) -> list[str]:
    errs, cc, fs = [], cfg["clickup"], fontes(cfg)
    doc_url = doc_link(doc_id)
    # rodapé na página de cada copy
    for f in fs:
        rot = "copia" if len(fs) == 1 else f"copia {f.get('nome', f['page_id'])}"
        md = cu.pagina(cc["workspace_id"], f["doc_id"], f["page_id"]).get("content") or ""
        if contem(md, doc_id):
            print(f"   {rot}: rodapé já aponta para este doc (já)")
        elif RODAPE_ROTULO in unesc(md):
            errs.append(f"{rot}: já existe um rodapé 'Documento Gravação Expert' de OUTRO doc — não acrescentei; resolva à mão")
        elif not apply:
            print(f"   [dry] {rot}: acrescentaria no fim → {rodape_copia(doc_url).strip()}")
        else:
            try:
                cu.pagina_acrescentar(cc["workspace_id"], f["doc_id"], f["page_id"], rodape_copia(doc_url))
            except RuntimeError as e:
                errs.append(f"{rot}: {e}")
            volta = cu.pagina(cc["workspace_id"], f["doc_id"], f["page_id"]).get("content") or ""
            n = unesc(volta).count(RODAPE_ROTULO)
            antes_ok = re.sub(r"\s+", " ", unesc(volta)).startswith(re.sub(r"\s+", " ", unesc(md)).strip())
            if not contem(volta, doc_id) or n != 1 or not antes_ok:
                errs.append(f"{rot}: releitura não confere (doc no rodapé={contem(volta, doc_id)}, rodapés={n}, "
                            f"conteúdo anterior preservado={antes_ok})")
            else:
                print(f"   {rot}: rodapé acrescentado ✔")
    # linhas na descrição de cada task
    for tarefa in tarefas(cfg):
        errs += _linhas_na_tarefa(cfg, tarefa, doc_id, upload_id, pasta, apply)
    return errs


def _linhas_na_tarefa(cfg: dict, tarefa: dict, doc_id: str, upload_id: str, pasta: Path, apply: bool) -> list[str]:
    tid = tarefa["id"]
    t = cu.tarefa(tid)
    if cfg["projeto"] not in (t.get("name") or ""):
        return [f"task {tid}: nome {t.get('name')!r} não tem {cfg['projeto']} — não mexi"]
    rot, md = f"task {tid}", t.get("markdown_description") or ""
    linhas = linhas_tarefa(cfg, tarefa, doc_id, upload_id)
    novas = [(k, x) for k, x in linhas if not contem(md, k)]
    novo, n, aviso = inserir_linhas(md, tarefa.get("inserir_apos", "$^"), linhas)
    if aviso:
        print(f"   ⚠ {rot}: {aviso}")
    if not n:
        print(f"   {rot}: linhas já estão na descrição (já)")
        return []
    if not apply:
        print(f"   [dry] {rot} {t.get('name')}: inseriria {n} linha(s) depois de {tarefa.get('inserir_apos')!r}")
        return []
    gravar_json(pasta / "historico" / f"tarefa_{tid}_{agora()}.json", {"markdown_description": md, "text_content": t.get("text_content")})
    if (cu.tarefa(tid).get("markdown_description") or "") != md:              # outra sessão mexeu no meio
        return [f"{rot}: a descrição mudou durante a execução — não sobrescrevi; rode de novo"]
    try:
        cu.tarefa_trocar_descricao(tid, novo)
    except RuntimeError as e:
        return [f"{rot}: {e}"]
    volta = cu.tarefa(tid).get("markdown_description") or ""
    if not all(contem(volta, k) for k, _ in linhas):
        return [f"{rot}: releitura sem os links novos"]
    if not descricao_preservada(md, volta, novas):
        print(f"   ⚠ {rot}: o ClickUp reformatou parte da descrição ao gravar (o backup está em historico/)")
    print(f"   {rot}: {n} linha(s) inserida(s) ✔")
    return []


# ----------------------------------------------------------------------------- comandos
def cmd_gerar(args) -> int:
    cfg_path = Path(args.config).expanduser().resolve()
    cfg, pasta = carregar_config(cfg_path), cfg_path.parent
    estado = ler_json(pasta / "estado.json", {})
    dv, fs = cfg["drive"], fontes(cfg)
    modo = "DOCX-ONLY" if args.docx_only else "APPLY" if args.apply else "DRY-RUN"
    print(f"== doc_gravacao · {cfg['projeto']} · {modo}")
    man = conferir_asset()

    # 1. copy (uma página ou várias, em `fontes`)
    m = ler_paginas(cfg, args.md)
    corpo = linhas_esperadas(m, cfg["doc"]["titulo_orientacoes"])
    print(f"\n-- VAI PARA A EXPERT ({len(m['itens'])} roteiros · hash {m['hash']})")
    for ln in [cfg["doc"]["nome"], DRIVE_ROTULO + " <link da pasta de upload>", *corpo]:
        print(f"   | {ln}")
    print("\n-- FICA DE FORA")
    for f in m["fora"]:
        print(f"   - {f}")

    saida = pasta / "preview"
    nome_docx = re.sub(r"[^\w.-]+", "_", cfg["doc"]["nome"]).strip("_") + ".docx"
    if args.docx_only:
        url = estado.get("upload_url") or "https://drive.google.com/drive/folders/PASTA_DE_UPLOAD"
        saida.mkdir(parents=True, exist_ok=True)
        st = montar_docx(ASSET, cfg["doc"]["nome"], url, m, cfg["doc"]["titulo_orientacoes"], saida / nome_docx)
        print(f"\ndocx: {saida / nome_docx} · {st}")
        if args.preview:
            print(f"prévia: {', '.join(str(p) for p in previa(saida / nome_docx, saida))}")
        return 0

    # 2. Drive: alvos
    d = dapi.Drive()
    d.drive_id = d.meta(dv["projeto_id"], "id,name,driveId").get("driveId")
    print("\n-- Drive")
    copia_f, bruto_f = _pasta_do_projeto(d, cfg, "copy"), _pasta_do_projeto(d, cfg, "bruto")
    print(f"   1.Copy: {copia_f['name']} · 14: {bruto_f['name']}")
    if d.publicos(bruto_f["id"]):
        print(f"!! {bruto_f['name']} tem acesso público {d.publicos(bruto_f['id'])}: a categoria nunca ficaria privada")
        return 1
    num = dv["categoria"].split(" ")[0]
    for f in d.filhos(bruto_f["id"], dapi.FOLDER_MIME):
        if f["name"].startswith(num + " ") and norm_nome(f["name"]) != norm_nome(dv["categoria"]):
            print(f"   ⚠ já existe outra pasta com o número {num} em {bruto_f['name']}: {f['name']!r}")
    cat = garantir_pasta(d, bruto_f["id"], dv["categoria"], estado.get("categoria_id"), args.apply, "categoria")
    if cat and d.publicos(cat["id"]):
        print(f"!! a categoria tem acesso público {d.publicos(cat['id'])}: vazaria para as subpastas — tire à mão")
        return 1
    up = garantir_pasta(d, cat["id"], dv["pasta_upload"], estado.get("upload_id"), args.apply, "pasta de upload") if cat else None
    if not cat:
        print(f"   [dry] pasta de upload: criaria '{dv['pasta_upload']}' dentro da categoria")
    ok = garantir_publico(d, up["id"], dv["papel_pasta_upload"], "pasta de upload", args.apply) if up else True
    url_pasta = folder_link(up["id"]) if up else "https://drive.google.com/drive/folders/PASTA_DE_UPLOAD"

    # 3. DOCX
    saida.mkdir(parents=True, exist_ok=True)
    docx_path = saida / nome_docx
    st = montar_docx(ASSET, cfg["doc"]["nome"], url_pasta, m, cfg["doc"]["titulo_orientacoes"], docx_path)
    print(f"\n-- docx montado: {docx_path.name} · HR={st['hr']} H2={st['h2']} H3={st['h3']} logo={st['logo']}")
    if args.preview:
        print(f"   prévia: {', '.join(p.name for p in previa(docx_path, saida))} em {saida}")

    # 4. doc no Drive
    print("\n-- documento")
    gdoc = _pelo_estado(d, estado.get("doc_id"), copia_f["id"], "doc") or d.achar(copia_f["id"], cfg["doc"]["nome"], dapi.GDOC_MIME)
    saida_cod = 0
    if gdoc and args.update:
        atual = texto_sha(d.exportar(gdoc["id"], "text/plain").decode("utf-8-sig"))
        if estado.get("doc_texto_sha") and atual != estado["doc_texto_sha"] and not args.forcar:
            meta = d.meta(gdoc["id"], "modifiedTime,lastModifyingUser(displayName)")
            print(f"!! o doc foi editado depois da última geração ({meta.get('lastModifyingUser', {}).get('displayName')}, "
                  f"{meta.get('modifiedTime')}). Confira com o Cleber; para sobrescrever: --update --forcar")
            return 3
        if args.apply:
            (pasta / "historico").mkdir(parents=True, exist_ok=True)
            (pasta / "historico" / f"doc_antes_{agora()}.docx").write_bytes(d.exportar(gdoc["id"], dapi.DOCX_MIME))
            d.trocar_conteudo(gdoc["id"], docx_path.read_bytes())
            print(f"   doc atualizado no mesmo link · {gdoc['id']}")
        else:
            print(f"   [dry] trocaria o conteúdo do doc {gdoc['id']} (o link não muda)")
    elif gdoc:
        print(f"   doc existe · {gdoc['id']}")
        if estado.get("copy_hash") and estado["copy_hash"] != m["hash"]:
            print("!! a copy mudou desde a geração e o doc NÃO foi atualizado: rode --apply --update")
            saida_cod = 2
    elif args.apply:
        try:
            gdoc = d.subir_docx_como_gdoc(docx_path.read_bytes(), cfg["doc"]["nome"], copia_f["id"])
        except Exception as e:  # noqa: BLE001
            print(f"   ⚠ upload falhou ({str(e)[:80]}); reconferindo pelo nome")
            gdoc = d.achar(copia_f["id"], cfg["doc"]["nome"], dapi.GDOC_MIME) or \
                d.subir_docx_como_gdoc(docx_path.read_bytes(), cfg["doc"]["nome"], copia_f["id"])
        print(f"   doc criado · {gdoc['id']}")
    else:
        print(f"   [dry] criaria o Google Doc '{cfg['doc']['nome']}' em {copia_f['name']}")
    if gdoc:
        ok &= garantir_publico(d, gdoc["id"], dv["papel_doc"], "doc", args.apply)

    if not args.apply:
        print("\n-- ClickUp")
        if gdoc and up:
            linkar_clickup(cfg, gdoc["id"], up["id"], pasta, apply=False)
        else:
            print(f"   [dry] rodapé '{RODAPE_ROTULO}' na(s) página(s) da copy e linhas nas tasks {[x['id'] for x in tarefas(cfg)]}")
        print("\n[DRY-RUN] nada foi escrito. Rode com --apply.")
        return saida_cod

    # 5. verificação (o doc só vale se o Google devolver o que mandamos)
    print("\n-- verificação")
    if saida_cod == 2:
        errs, doc_sha = [], estado.get("doc_texto_sha")
    else:
        errs, doc_sha = verificar_doc(d, gdoc["id"], cfg, m, url_pasta)
    errs += verificar_acessos(d, cfg, gdoc["id"], up["id"], cat["id"], saida)
    if not ok:
        errs.append("permissão não aplicada")
    if errs:
        for e in errs:
            print(f"   !! {e}")
        print("\n!! verificação falhou — estado e ClickUp NÃO foram tocados")
        return 1

    # 6. ClickUp
    print("\n-- ClickUp")
    ck = []
    if args.sem_clickup:
        print("   (pulado: --sem-clickup)")
    else:
        ck = linkar_clickup(cfg, gdoc["id"], up["id"], pasta, apply=True)
        for e in ck:
            print(f"   !! {e}")

    # 7. estado + links
    novo = {"projeto": cfg["projeto"], "doc_nome": cfg["doc"]["nome"], "doc_id": gdoc["id"], "doc_url": doc_link(gdoc["id"]),
            "upload_id": up["id"], "upload_url": url_pasta, "categoria_id": cat["id"],
            "upload_caminho": f"{bruto_f['name']} › {dv['categoria']} › {dv['pasta_upload']}",
            "copy_hash": m["hash"] if saida_cod != 2 else estado.get("copy_hash"), "doc_texto_sha": doc_sha,
            "modelo_sha": man.get("sha256", "")[:16], "gerado_em": time.strftime("%d/%m/%Y %H:%M"),
            "clickup": {"pendencias": ck, "sem_clickup": bool(args.sem_clickup)}}
    gravar_json(pasta / "estado.json", novo)
    (pasta / "links.md").write_text("\n".join([
        f"# Documento de Gravação Expert — {cfg['projeto']} ({cfg['expert']['nome']})", "",
        f"Gerado pela skill `doc-gravacao-expert` em {novo['gerado_em']} a partir "
        f"{'da copy' if len(fs) == 1 else 'das copies'} do ClickUp ("
        + "; ".join(("" if len(fs) == 1 else f"{f.get('nome', f['page_id'])}: ") + f"doc `{f['doc_id']}`, página `{f['page_id']}`"
                    for f in fs) + f", hash `{novo['copy_hash']}`).", "",
        "| Documento | Doc de gravação (qualquer pessoa com o link edita) | Pasta de upload (qualquer pessoa com o link sobe arquivo) |",
        "|---|---|---|", f"| {cfg['doc']['nome']} | {novo['doc_url']} | {url_pasta} |", "",
        f"Upload: `{novo['upload_caminho']}` (a categoria fica sem acesso público).", "",
        "Fora do doc (interno): " + "; ".join(m["fora"]) + ".", "",
        "Se a copy mudar no ClickUp: `gerar --config <este config> --apply --update` troca o conteúdo e mantém o link.", ""]),
        encoding="utf-8")
    print(f"\n== RESUMO\n  doc:    {novo['doc_url']}\n  upload: {url_pasta}\n          {novo['upload_caminho']}"
          f"\n  estado: {pasta / 'estado.json'}")
    return 1 if ck else saida_cod


def cmd_verificar(args) -> int:
    cfg_path = Path(args.config).expanduser().resolve()
    cfg, pasta = carregar_config(cfg_path), cfg_path.parent
    est = ler_json(pasta / "estado.json", {})
    if not est.get("doc_id"):
        print("sem estado.json: rode `gerar --apply` primeiro")
        return 1
    m = ler_paginas(cfg)
    d = dapi.Drive()
    errs, sha = verificar_doc(d, est["doc_id"], cfg, m, est["upload_url"])
    errs += verificar_acessos(d, cfg, est["doc_id"], est["upload_id"], est["categoria_id"], pasta / "preview")
    if m["hash"] != est.get("copy_hash"):
        errs.append("a copy mudou no ClickUp desde a geração: rode `gerar --apply --update`")
    if est.get("doc_texto_sha") and sha != est["doc_texto_sha"]:
        print("   ⚠ o doc foi editado depois da geração (normal se a expert anotou nele)")
    for e in errs:
        print(f"   !! {e}")
    print("OK" if not errs else "COM PENDÊNCIAS")
    return 1 if errs else 0


def cmd_modelo(args) -> int:
    """Exporta o doc-modelo ao vivo, guarda só o cabeçalho (logo, título, DRIVE + link) e congela em assets/."""
    d = dapi.Drive()
    meta = d.meta(args.doc_id, "id,name,modifiedTime")
    doc = Document(io.BytesIO(d.exportar(args.doc_id, dapi.DOCX_MIME)))
    body = doc.element.body
    cab = localizar_cabecalho(body)
    _cortar_depois(body, cab["hr3"])
    for tag in ("commentRangeStart", "commentRangeEnd", "del"):
        for e in list(body.iter(W + tag)):
            e.getparent().remove(e)
    for e in [r for r in body.iter(W + "r") if r.find(W + "commentReference") is not None]:
        e.getparent().remove(e)
    for e in list(body.iter(W + "ins")):
        pai, i = e.getparent(), e.getparent().index(e)
        for filho in list(e):
            pai.insert(i, filho)
            i += 1
        pai.remove(e)
    for rid, rel in list(doc.part.rels.items()):
        if rel.reltype.split("/")[-1].startswith("comments") or rel.reltype.endswith("/people"):
            doc.part.rels.pop(rid)
    _trocar_texto(cab["titulo"], "TÍTULO DO DOCUMENTO")
    hl = cab["link"].find(W + "hyperlink")
    doc.part.rels[hl.get(R + "id")]._target = "https://drive.google.com/"
    _trocar_texto(hl, "https://drive.google.com/")
    ASSET.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(ASSET))
    st = docx_stats(ASSET.read_bytes())
    if (st["hr"], st["h2"], st["h3"]) != (3, 0, 0) or not st["logo"] or st["comentarios"]:
        raise RuntimeError(f"modelo congelado fora do esperado: {st}")
    gravar_json(MANIFESTO, {"fonte_id": args.doc_id, "fonte_nome": meta.get("name"), "fonte_modificado": meta.get("modifiedTime"),
                            "congelado_em": time.strftime("%Y-%m-%d %H:%M"), "sha256": hashlib.sha256(ASSET.read_bytes()).hexdigest(),
                            "logo_sha": st["logo_sha"]})
    print(f"modelo congelado: {ASSET} · {st} · fonte {meta.get('name')!r} ({meta.get('modifiedTime')})")
    return 0


def cmd_listar(_args) -> int:
    for cfg_path in sorted(EXOS.glob("*/06-Copy/gravacao-*/config.json")):
        cfg, est = json.loads(cfg_path.read_text(encoding="utf-8")), ler_json(cfg_path.parent / "estado.json", {})
        print(f"{cfg['projeto']:24} {cfg['expert']['nome']:24} {est.get('doc_url', '(ainda não gerado)')}\n{'':24} {cfg_path}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    g = sub.add_parser("gerar", help="cria/atualiza pasta de upload, doc e links no ClickUp (dry-run por padrão)")
    g.add_argument("--config", required=True)
    g.add_argument("--apply", action="store_true", help="escreve no Drive e no ClickUp")
    g.add_argument("--update", action="store_true", help="doc já existe: troca o conteúdo e mantém o link")
    g.add_argument("--forcar", action="store_true", help="com --update: sobrescreve mesmo se o doc foi editado à mão")
    g.add_argument("--docx-only", action="store_true", help="só monta o DOCX local (nada no Drive/ClickUp)")
    g.add_argument("--preview", action="store_true", help="gera PDF/PNG local do DOCX (LibreOffice + PyMuPDF)")
    g.add_argument("--md", action="append", default=None,
                   help="markdown local da copy, um por fonte e na ordem delas (padrão: lê o ClickUp ao vivo)")
    g.add_argument("--sem-clickup", action="store_true", help="não escreve no ClickUp")
    v = sub.add_parser("verificar", help="confere doc, acessos e copy (só leitura)")
    v.add_argument("--config", required=True)
    mo = sub.add_parser("modelo", help="recongela o cabeçalho do doc-modelo em assets/")
    mo.add_argument("--doc-id", default=MODELO_PADRAO)
    sub.add_parser("listar", help="projetos configurados")
    args = ap.parse_args()
    return {"gerar": cmd_gerar, "verificar": cmd_verificar, "modelo": cmd_modelo, "listar": cmd_listar}[args.cmd](args)


if __name__ == "__main__":
    sys.exit(main())
