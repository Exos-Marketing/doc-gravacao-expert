"""Testes offline do motor doc_gravacao.py: leitura da copy, montagem do DOCX e linhas do ClickUp.

Rodar:
  uv run --python 3.12 --with python-docx==1.2.0 --with requests --with pytest \
    python -m pytest ~/.claude/skills/doc-gravacao-expert/scripts/test_doc_gravacao.py -q

Fixtures = copies reais lidas do ClickUp em 15/10/26 (Daniela, Parte 2; Elen, com o rodapé que entrou depois)
e a descrição da task Criativos Vídeo [DM][PEP][FLS][26]. O gabarito da Elen saiu do parser antigo
(drive_docs_gravacao_et_fls.py) com o rodapé removido.
"""
import copy
import json
import sys
import unicodedata
from pathlib import Path

import pytest
from docx import Document

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import doc_gravacao as dg  # noqa: E402

FX = HERE / "fixtures"
ASSET = HERE.parent / "assets" / "modelo_gravacao.docx"
CFG_DM = json.loads((HERE / "config_exemplo.json").read_text(encoding="utf-8"))
CFG_ET = json.loads((FX / "et_pep_fls_26.config.json").read_text(encoding="utf-8"))
MD_DM = (FX / "dm_pep_fls_26.md").read_text(encoding="utf-8")
MD_ET = (FX / "et_pep_fls_26.md").read_text(encoding="utf-8")
TASK_DM = (FX / "dm_task_criativos_video.md").read_text(encoding="utf-8")
CFG_BF = json.loads((FX / "dm_bf_antecipacao.config.json").read_text(encoding="utf-8"))
MDS_BF = [(FX / f"dm_bf_antecipacao_{n}.md").read_text(encoding="utf-8") for n in ("trailer", "video")]
CFG_VENDAS = json.loads((FX / "dm_bf_vendas.config.json").read_text(encoding="utf-8"))
MD_VENDAS = (FX / "dm_bf_vendas.md").read_text(encoding="utf-8")
CTA = "Clica em Saiba Mais e entra no grupo do Clube."
URL = "https://drive.google.com/drive/folders/PASTA?usp=sharing"


def ler_dm(md=MD_DM, cc=None):
    return dg.ler_copia(md, cc or CFG_DM["copia"])


def ler_bf(mds=None):
    return dg.ler_fontes(mds or MDS_BF, dg.fontes(CFG_BF))


# ----------------------------------------------------------------------------- copy da Daniela
def test_dm_recorta_so_a_parte_2():
    texto = "\n".join(dg.linhas_esperadas(ler_dm(), "ORIENTAÇÕES DE GRAVAÇÃO"))
    assert "PARTE 1" not in texto and "ADS 01" not in texto and "Lote 1" not in texto


def test_dm_orientacoes_na_ordem_da_copy_com_cta_no_fim():
    m = ler_dm()
    rotulos = [dg.plain(p).split(":")[0] for p in m["paragrafos"]]
    assert rotulos == ["Quem grava", "Orientações de gravação", "Nome falado", "Fala neutralizada",
                       "Cenário e figurino", "Identificação na tela (texto, não é fala), nos primeiros segundos",
                       "Tomada livre", "CTA (fecha os três)"]
    assert m["cta_texto"] == CTA


def test_dm_descarta_o_interno_e_tira_a_remissao_a_parte_1():
    m = ler_dm()
    texto = " ".join(dg.linhas_esperadas(m, "X"))
    assert "Configuração" not in texto and "Texto do anúncio" not in texto and "CPL" not in texto
    assert "(arte na Parte 1)" not in texto and "Fechamento com o banner do Clube." in texto
    assert any("Configuração:" in f for f in m["fora"]) and any("Texto do anúncio:" in f for f in m["fora"])


def test_dm_itens_12_13_14_com_headline_e_fim_no_cta():
    m = ler_dm()
    assert [i["num"] for i in m["itens"]] == ["12", "13", "14"]
    assert [len(i["falas"]) for i in m["itens"]] == [7, 5, 6]
    assert all(dg.plain(i["falas"][-1]) == CTA for i in m["itens"])
    assert dg.plain(m["itens"][0]["campos"][0]) == "Headline: PROFISSIONAL DA SAÚDE, TODA TERÇA, 20H, AO VIVO E DE GRAÇA"


def test_dm_negrito_interno_fica_negrito():
    nome = next(p for p in ler_dm()["paragrafos"] if dg.plain(p).startswith("Nome falado:"))
    assert ("Clube da Dra. Daniela", True) in nome


def test_hash_estavel_e_cego_ao_que_fica_fora_da_secao():
    a = ler_dm()["hash"]
    rodape = "\n\n**Documento Gravação Expert:** [https://docs.google.com/document/d/abc/edit](https://docs.google.com/document/d/abc/edit)\n"
    assert a == ler_dm(MD_DM + rodape)["hash"] and len(a) == 12


# ----------------------------------------------------------------------------- copy da Elen (schema genérico)
def test_et_reproduz_o_gabarito_do_parser_antigo_e_ignora_o_rodape():
    m = dg.ler_copia(MD_ET, CFG_ET["copia"])
    gabarito = (FX / "et_pep_fls_26.expected.txt").read_text(encoding="utf-8").splitlines()
    assert "Documento Gravação Expert" in MD_ET
    assert dg.linhas_esperadas(m, "ORIENTAÇÕES DE GRAVAÇÃO") == gabarito


# ----------------------------------------------------------------------------- várias copies num doc (Black da Daniela)
def test_sem_fontes_e_a_copia_de_sempre_com_o_mesmo_hash():
    fs = dg.fontes(CFG_DM)
    assert len(fs) == 1 and fs[0]["copia"] is CFG_DM["copia"]
    assert dg.ler_fontes([MD_DM], fs) == ler_dm()


def test_fontes_juntam_trailer_e_videos_na_ordem():
    m = ler_bf()
    assert [i["num"] for i in m["itens"]] == ["TRAILER", "02", "03", "04", "05"]
    assert [dg.plain(p).split(":")[0] for p in m["paragrafos"]] == ["Orientações de gravação"]
    trailer = m["itens"][0]
    assert [dg.plain(c).split(":")[0] for c in trailer["campos"]] == ["Orientações de gravação"]
    assert "Exemplo do formato: instagram.com/reel/DO_OCNgjY4a. Sem CTA" in dg.plain(trailer["campos"][0])
    assert dg.plain(trailer["falas"][-1]).endswith("especialista na Estética Íntima está chegando.")
    assert dg.plain(m["itens"][2]["falas"][0]) == "[Início, lendo a pergunta da caixinha]"
    assert dg.plain(m["itens"][4]["falas"][0]).startswith("(roteiro-guia; ela pode falar")
    assert dg.plain(m["itens"][4]["falas"][-1]) == "Fica de olho!"


def test_fontes_deixam_o_interno_de_fora():
    m = ler_bf()
    texto = "\n".join(dg.linhas_esperadas(m, "ORIENTAÇÕES DE GRAVAÇÃO")).lower()
    for interno in ("configuração", "tráfego", "orgânico", "variação livre", "texto na tela", "regras desta fase",
                    "para o documento de gravação", "anúncio:", "* * *", "https", "finados"):
        assert interno not in texto, interno
    assert any(f.startswith("Trailer · descartado: Texto na tela") for f in m["fora"])
    assert any(f.startswith("Vídeos · descartado: Regras desta fase") for f in m["fora"])


def test_fontes_hash_depende_da_ordem_e_roteiro_repetido_aborta():
    fs = dg.fontes(CFG_BF)
    assert dg.ler_fontes(MDS_BF[::-1], fs[::-1])["hash"] != ler_bf()["hash"]
    with pytest.raises(ValueError, match="repetido"):
        dg.ler_fontes([MDS_BF[1], MDS_BF[1]], [fs[1], fs[1]])


def test_proibido_do_doc_inteiro_e_o_que_todas_as_fontes_barram():
    assert dg.proibido_re(CFG_DM) == dg.BASE_PROIBIDO + "|" + CFG_DM["copia"]["proibido"]
    assert dg.proibido_re(CFG_BF).count("|tr[áa]fego") == 1
    rx = dg.proibido_re(CFG_VENDAS)                      # os áudios da reta final falam "Clube da PPK"
    assert r"\bppk\b" not in rx and "|349" in rx and "|banner" in rx
    assert dg._alternativas(r"a|(b|c)|[|d]|e\|f") == ["a", "(b|c)", "[|d]", r"e\|f"]


def test_vendas_duas_secoes_da_mesma_pagina_e_valor_fora_da_fala():
    m = dg.ler_fontes([MD_VENDAS, MD_VENDAS], dg.fontes(CFG_VENDAS))
    assert [i["num"] for i in m["itens"]] == [f"{n:02d}" for n in range(1, 9)] + ["RV 01", "RV 02", "RV 03", "RV 04"]
    texto = "\n".join(dg.linhas_esperadas(m, "ORIENTAÇÕES DE GRAVAÇÃO"))
    for fora in ("349", "EDITOR", "RV 05", "Orientações para a edição", "Cadência", "Headline"):
        assert fora not in texto, fora
    assert texto.count("[aponta para o lado ou para baixo da tela]") == 8
    assert "[Cada frase de presente é gravada como frase solta, para o editor cortar sem regravar.]\nAs três primeiras" in texto
    assert m["itens"][7]["texto_rotulo"] == "Texto para gravação (guia):"
    assert m["itens"][8]["texto_rotulo"] == "Texto para gravação (áudio da Dra. Daniela Moleiro):"


def test_docx_com_varias_fontes(tmp_path):
    m = ler_bf()
    st = dg.montar_docx(ASSET, "Copy de Criativo Antecipação Vídeo [DM][BF][POS][NOV][26]", URL, m,
                        "ORIENTAÇÕES DE GRAVAÇÃO", tmp_path / "bf.docx")
    assert (st["hr"], st["h2"], st["h3"]) == (4, 1, 5) and st["logo"]


# ----------------------------------------------------------------------------- abortos
def _trocar(antigo, novo):
    assert antigo in MD_DM
    return MD_DM.replace(antigo, novo, 1)


def test_aborta_rotulo_desconhecido():
    with pytest.raises(ValueError, match="Novo campo"):
        ler_dm(_trocar("**Tomada livre:**", "**Novo campo:** algo.\n\n**Tomada livre:**"))


def test_aborta_quando_ajuste_nao_casa_o_numero_esperado():
    with pytest.raises(ValueError, match="arte da Parte 1"):
        ler_dm(MD_DM.replace(" (arte na Parte 1)", ""))


def test_aborta_termo_proibido_no_texto_mantido():
    with pytest.raises(ValueError, match="(?i)cpl"):
        ler_dm(_trocar("Não tem replay. Quem está ao vivo, aprende.", "Não tem replay. O CPL desse vídeo é ótimo."))


def test_aborta_quando_falta_item():
    cc = copy.deepcopy(CFG_DM["copia"])
    cc["itens"]["numeros"] = ["12", "13", "14", "15"]
    with pytest.raises(ValueError, match="15"):
        ler_dm(cc=cc)


def test_aborta_quando_item_nao_termina_no_cta():
    corte = MD_DM.index("**ADS 14")
    antes, depois = MD_DM[:corte], MD_DM[corte:]
    i = antes.rindex(CTA)
    md = antes[:i] + "Clica no link e entra no grupo." + antes[i + len(CTA):] + depois
    with pytest.raises(ValueError, match="ADS 13"):
        ler_dm(md)


def test_ancora_obrigatoria_e_unica():
    cc = copy.deepcopy(CFG_DM["copia"])
    cc["inicio"] = "PARTE 9 · NÃO EXISTE"
    with pytest.raises(ValueError, match="âncora"):
        ler_dm(cc=cc)
    with pytest.raises(ValueError, match="âncora"):
        ler_dm(MD_DM + "\n\n**PARTE 2 · ROTEIROS DE VÍDEO (só teste para público frio)**\n")


# ----------------------------------------------------------------------------- DOCX
def test_cabecalho_do_modelo_validado_por_posicao():
    doc = Document(str(ASSET))
    cab = dg.localizar_cabecalho(doc.element.body)
    assert dg.ptext(cab["drive"]).strip() == "DRIVE para subir arquivos:"
    cab["drive"].getparent().remove(cab["drive"])
    with pytest.raises(RuntimeError, match="modelo"):
        dg.localizar_cabecalho(doc.element.body)


@pytest.fixture(scope="module")
def docx_dm(tmp_path_factory):
    m = ler_dm()
    out = tmp_path_factory.mktemp("docx") / "dm.docx"
    dg.montar_docx(ASSET, "Copy Criativos Vídeo [DM][PEP][FLS][26]", URL, m, "ORIENTAÇÕES DE GRAVAÇÃO", out)
    return m, out


def test_docx_estrutura_e_logo(docx_dm):
    _, out = docx_dm
    st = dg.docx_stats(out.read_bytes())
    assert (st["hr"], st["h2"], st["h3"]) == (4, 1, 3)
    assert st["logo"]


def test_docx_titulo_link_e_texto_em_ordem(docx_dm):
    m, out = docx_dm
    doc = Document(str(out))
    textos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    assert textos == ["Copy Criativos Vídeo [DM][PEP][FLS][26]", "DRIVE para subir arquivos:", URL,
                      *dg.linhas_esperadas(m, "ORIENTAÇÕES DE GRAVAÇÃO")]
    assert URL in [r._target for r in doc.part.rels.values() if r.reltype.endswith("/hyperlink")]


# ----------------------------------------------------------------------------- ClickUp
def test_inserir_linhas_depois_da_linha_da_copy_e_idempotente():
    linhas = [("PASTA123", "Vídeos brutos da Daniela (upload): [https://x/PASTA123](https://x/PASTA123)"),
              ("DOC456", "Documento Gravação Expert: [https://y/DOC456](https://y/DOC456)")]
    novo, n, aviso = dg.inserir_linhas(TASK_DM, r"^Documento de entrega da copy:", linhas)
    assert n == 2 and aviso is None
    ls = novo.split("\n")
    i = next(k for k, l in enumerate(ls) if l.startswith("Documento de entrega da copy:"))
    assert [l for l in ls[i + 1:i + 5] if l] == [t for _, t in linhas]
    assert novo.replace("".join(f"\n\n{t}" for _, t in linhas), "", 1) == TASK_DM
    de_novo, n2, _ = dg.inserir_linhas(novo, r"^Documento de entrega da copy:", linhas)
    assert n2 == 0 and de_novo == novo


def test_descricao_preservada_mesmo_com_o_clickup_escapando_underscore_no_link():
    # o ClickUp devolve o texto do link com `\_` (1kiYIslbU\_kof…): a conferência não pode acusar mudança
    u = "https://drive.google.com/drive/folders/1kiYIslbU_kofMQ?usp=sharing"
    linhas = [("1kiYIslbU_kofMQ", f"Vídeos brutos da Daniela (upload): [{u}]({u})")]
    novo, _, _ = dg.inserir_linhas(TASK_DM, r"^Documento de entrega da copy:", linhas)
    como_o_clickup_devolve = novo.replace(f"[{u}]", "[" + u.replace("_", "\\_") + "]")
    assert dg.descricao_preservada(TASK_DM, como_o_clickup_devolve, linhas)
    assert not dg.descricao_preservada(TASK_DM, como_o_clickup_devolve.replace("Briefing", "Resumo"), linhas)


def test_inserir_sem_a_linha_de_referencia_vai_para_o_fim_com_aviso():
    novo, n, aviso = dg.inserir_linhas("Linha qualquer.", r"^Documento de entrega da copy:", [("K1", "Documento Gravação Expert: K1")])
    assert n == 1 and aviso and novo.endswith("Documento Gravação Expert: K1")


def test_rotulo_do_texto_aceita_complemento_entre_parenteses():
    t = "Texto para gravação:"
    assert dg.rotulo_do_texto("Texto para gravação:", t)
    assert dg.rotulo_do_texto("Texto para gravação (guia):", t)
    assert dg.rotulo_do_texto("Texto para gravação (áudio da Dra. Daniela Moleiro):", t)
    assert not dg.rotulo_do_texto("Texto para gravar:", t) and not dg.rotulo_do_texto("Texto para gravação (guia)", t)


def test_tarefas_aceita_a_forma_antiga_e_a_lista():
    assert dg.tarefas(CFG_DM) == [CFG_DM["clickup"]["tarefa"]]
    lista = [{"id": "A", "linhas": ["doc"]}, {"id": "B"}]
    assert dg.tarefas({"clickup": {"tarefas": lista}}) == lista
    assert dg.tarefas({"clickup": {}}) == []


def test_linhas_da_task_levam_o_nome_do_doc_e_respeitam_a_escolha():
    cfg = {"expert": {"curto": "Daniela"}, "doc": {"nome": "Copy Vídeo [DM][BF][POS][NOV][26]"}}
    duas = dg.linhas_tarefa(cfg, {"id": "T"}, "DOC1", "PASTA1")
    assert [k for k, _ in duas] == ["PASTA1", "DOC1"]
    assert duas[0][1].startswith("Vídeos brutos da Daniela (upload): [https://drive.google.com/drive/folders/PASTA1")
    assert duas[1][1].startswith("Documento Gravação Expert: Copy Vídeo \\[DM\\]\\[BF\\]\\[POS\\]\\[NOV\\]\\[26\\] — [https://")
    assert [k for k, _ in dg.linhas_tarefa(cfg, {"id": "T", "linhas": ["doc"]}, "DOC1", "PASTA1")] == ["DOC1"]


def test_contem_reconhece_id_com_underscore_escapado():
    assert dg.contem("[https://docs.google.com/document/d/1\\_oSNf/edit]", "1_oSNf")


def test_rodape_da_copy_no_formato_da_elen():
    u = "https://docs.google.com/document/d/X/edit?usp=sharing"
    assert dg.rodape_copia(u) == f"\n\n**Documento Gravação Expert:** [{u}]({u})"


def test_perfil_do_libreoffice_codifica_espaco_no_caminho():
    # "Daniela Moleiro" tem espaço: file:// cru derruba o soffice (uno RuntimeException) e não sai PDF
    arg = dg.perfil_lo(Path("/Users/x/Daniela Moleiro/preview/.lo-perfil"))
    assert arg == "-env:UserInstallation=file:///Users/x/Daniela%20Moleiro/preview/.lo-perfil"


def test_nome_de_pasta_normalizado():
    nfd = unicodedata.normalize("NFD", "6. Criativos ADS de Captação")
    assert dg.norm_nome(nfd) == dg.norm_nome("6.  Criativos ADS de captação ")
