# doc-gravacao-expert

Skill do Claude Code que cria o **Documento de Gravação** da expert de um projeto da Exos: um Google Doc com o logo
da Exos no cabeçalho, o link da pasta de upload, as orientações de gravação e os roteiros dos vídeos (ADS NN) lidos
da copy no ClickUp. Cria também a pasta de upload dos vídeos brutos em `14.Arquivo Bruto` e registra os links no
ClickUp (rodapé da copy e descrição da task de criativos).

```
copy no ClickUp  →  dry-run (VAI PARA A EXPERT × FICA DE FORA)  →  --apply: pasta de upload + Google Doc  →  links no ClickUp
```

## Instalação

```bash
git clone https://github.com/Exos-Marketing/doc-gravacao-expert.git \
  ~/.claude/skills/doc-gravacao-expert
```

Requisitos:

- [uv](https://docs.astral.sh/uv/): os comandos rodam com `uv run --python 3.12`, que instala as dependências do
  script (`python-docx`, `requests`, `pymupdf`).
- rclone com o remote `gdrive` autenticado na conta `drive@exosmkt.com`: o script usa o token OAuth dele para a
  API do Drive.
- O **seu** token da API do ClickUp no Keychain, no serviço `clickup-exos-api`.
- A pasta de trabalho da Exos em `~/Documents/Exos`: cada projeto tem o config em
  `<Expert>/06-Copy/gravacao-<SIGLA>/config.json`.

**Nenhuma credencial mora neste repo.**

## Uso

```bash
SK=~/.claude/skills/doc-gravacao-expert/scripts
CFG="$HOME/Documents/Exos/<Expert>/06-Copy/gravacao-<SIGLA>/config.json"
uv run -q --python 3.12 $SK/doc_gravacao.py gerar --config "$CFG"                    # dry-run
uv run -q --python 3.12 $SK/doc_gravacao.py gerar --config "$CFG" --docx-only --preview
uv run -q --python 3.12 $SK/doc_gravacao.py gerar --config "$CFG" --apply
uv run -q --python 3.12 $SK/doc_gravacao.py gerar --config "$CFG" --apply --update   # a copy mudou: troca o conteúdo e mantém o link
uv run -q --python 3.12 $SK/doc_gravacao.py verificar --config "$CFG"                # só leitura
```

## Regras

- Dry-run antes de tudo: nada vai para o Drive ou o ClickUp sem `--apply`.
- Acesso por link só na pasta de upload (gestor de conteúdo) e no doc (editor), nunca na pasta do projeto.
- Nunca renomeia, move ou apaga; nunca mexe em status, responsável ou prazo da task.
- Nada é mandado para a expert: os links vão para quem pediu.
- Linha fora do previsto na copy aborta de propósito: ajuste o config, não o motor.

## O que tem aqui

| arquivo | o que faz |
|---|---|
| `SKILL.md` | projeto novo, copy que mudou, regras e erros comuns |
| `REFERENCE.md` | schema do config, regras do parser e armadilhas |
| `assets/modelo_gravacao.docx` | modelo do doc com o logo no cabeçalho (`modelo_gravacao.json` registra de onde ele veio) |
| `scripts/doc_gravacao.py` | o motor: lê a copy, monta o DOCX, cria a pasta e o doc no Drive e linka no ClickUp |
| `scripts/drive_api.py` | API do Drive com o token do rclone |
| `scripts/clickup_api.py` | REST do ClickUp com o token do Keychain |
| `scripts/config_exemplo.json` | modelo de config de projeto |
| `scripts/test_doc_gravacao.py` e `scripts/fixtures/` | testes offline, com copies reais de projetos da Daniela e da Elen |

---

Interno · Exos Marketing Digital e Internacionalização Ltda.
