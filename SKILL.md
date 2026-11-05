---
name: doc-gravacao-expert
description: Cria o Documento de Gravação da expert de um projeto da Exos — Google Doc com o logo Exos no cabeçalho, "DRIVE para subir arquivos:" com o link da pasta de upload, as orientações de gravação e os roteiros dos vídeos (ADS NN) lidos ao vivo da copy no ClickUp — e a pasta de upload dos vídeos brutos em 14.Arquivo Bruto, com acesso por link (doc = editor, pasta = gestor de conteúdo), registrando os links no ClickUp (rodapé da copy e descrição da task de criativos). Config por projeto na pasta da expert, dry-run por padrão, idempotente. Use quando o usuário pedir "documento de gravação da expert", "doc para a expert gravar", "roteiros para a expert gravar", "doc de gravação dos criativos de vídeo", "pasta de upload dos vídeos", "pasta para a expert subir os vídeos", "atualiza o doc de gravação, a copy mudou" ou "confere o doc de gravação". Não serve para roteiro de anfitriã/lives nem para criativos estáticos.
---

# Documento de Gravação da expert

O motor é `scripts/doc_gravacao.py` (testes em `scripts/test_doc_gravacao.py`). Cada projeto tem uma pasta de
dados na pasta da expert: `<Expert>/06-Copy/gravacao-<SIGLA sem colchetes>/` com `config.json`, `estado.json`,
`links.md`, `historico/` e `preview/`. Schema do config, regras do parser e gotchas: [REFERENCE.md](REFERENCE.md).

```bash
SK=~/.claude/skills/doc-gravacao-expert/scripts
CFG="$HOME/Documents/Exos/<Expert>/06-Copy/gravacao-<SIGLA>/config.json"
uv run -q --python 3.12 $SK/doc_gravacao.py listar                                  # projetos configurados
uv run -q --python 3.12 $SK/doc_gravacao.py gerar --config "$CFG"                   # dry-run (não escreve)
uv run -q --python 3.12 $SK/doc_gravacao.py gerar --config "$CFG" --docx-only --preview
uv run -q --python 3.12 $SK/doc_gravacao.py gerar --config "$CFG" --apply
uv run -q --python 3.12 $SK/doc_gravacao.py verificar --config "$CFG"               # só leitura
```

Exemplo real: `Daniela Moleiro/06-Copy/gravacao-DM-PEP-FLS-26/` (Clube da Dra. Daniela, ADS 12 a 14). Duas copies
num doc só (trailer + ADS 02 a 05 da Black, gravados na mesma sessão): `gravacao-DM-BF-POS-NOV-26-antecipacao/`,
com `fontes` no config.

## Projeto novo

1. **Copy.** Ache no ClickUp o Doc de copy dos criativos do projeto (a task `Criativos Vídeo [SIGLA]` costuma
   linkar). Leia a página em md e anote: a linha-âncora da seção de vídeo, os rótulos de orientação (Quem grava,
   Orientações de gravação, Nome falado…), o que é interno (Configuração, Texto do anúncio, CPL, verba…), o padrão
   dos itens (`ADS NN [...]`) e o CTA. Se a sessão de gravação junta peças de mais de uma copy (ex.: trailer +
   vídeos da Black), liste as páginas em `fontes`, uma por copy, na ordem do doc.
2. **Drive.** Ache a pasta do projeto (`<Cliente> › 2. Projetos › [SIGLA]`, descobrindo pelos filhos da pasta do
   cliente) e as subpastas `1.Copy…` e `14.Arquivo Bruto…`.
3. **Config.** Copie `scripts/config_exemplo.json` para `<Expert>/06-Copy/gravacao-<SIGLA>/config.json` e ajuste
   (schema no REFERENCE). Rótulo novo na copy = decisão do usuário: `manter` ou `descartar`.
4. **Dry-run.** Mostre ao usuário os blocos **VAI PARA A EXPERT** e **FICA DE FORA** antes do primeiro `--apply`.
5. **Prévia.** `--docx-only --preview` e leia o PNG da 1ª página em `preview/` (logo, título, link, títulos).
6. **Apply.** Cria `6. Criativos ADS de Captação` (privada) › `Criativos Vídeo [SIGLA]` (link = gestor de
   conteúdo), o Google Doc em `1.Copy` (link = editor), verifica tudo e linka no ClickUp. Leia
   `preview/doc_google_p1.png`: é o doc real, exportado do Google sem login. Para ver o doc antes de linkar no
   ClickUp: `--apply --sem-clickup` e depois `--apply` (idempotente).
7. **Entrega.** Link do doc e da pasta + mensagem curta para o usuário mandar à expert.

## A copy mudou

`gerar --config "$CFG" --apply --update` troca o conteúdo e **mantém o link**. Se o doc foi editado à mão depois
da geração, sai com código 3: pergunte ao usuário antes de `--forcar` (o doc antigo vai para `historico/`).

## Regras

- Dry-run antes de tudo; nada vai para o Drive ou o ClickUp sem `--apply`.
- Acesso público só na subpasta de upload e no doc. Nunca na categoria nem na pasta do projeto (vaza para as filhas).
- Nunca renomear, mover ou mandar para a lixeira; nunca mexer em status, responsável ou prazo da task.
- Nunca mandar nada para a expert: os links vão para o usuário.
- Linha fora do previsto aborta de propósito. Ajuste o config (`manter`, `descartar`, `adaptar`), não o motor.
- Mexeu no motor: rode os testes (`uv run -q --python 3.12 --with python-docx==1.2.0 --with requests --with pytest python -m pytest $SK/test_doc_gravacao.py -q`).

## Erros comuns

| Saída | O que fazer |
|---|---|
| `âncora de início … 0×` ou `2×` | ajustar `copia.inicio` (casamento total, uma linha só) |
| `rótulo desconhecido 'X:'` | perguntar ao usuário se é `manter` ou `descartar` e pôr no config |
| `ajuste … casou N×` | a copy mudou: revisar `copia.adaptar` |
| `termo interno no texto que vai para a expert` | falar com quem escreveu a copy, ou `adaptar` |
| `modelo fora do padrão` / manifesto não bate | `doc_gravacao.py modelo` recongela o cabeçalho com o logo |
| `403 transitório` em sequência | cota do client_id do rclone; esperar e rodar de novo |
| saída 2 | a copy mudou e o doc não: `--apply --update` |
| saída 3 | doc editado à mão: perguntar antes de `--forcar` |
