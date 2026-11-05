# doc-gravacao-expert — referência

## Config (`<Expert>/06-Copy/gravacao-<SIGLA>/config.json`)

Modelo comentado: o da Daniela, em `scripts/config_exemplo.json`. Outro exemplo (copy só de vídeo, sem "Parte 2"):
`scripts/fixtures/et_pep_fls_26.config.json` (Clube da Estomato da Elen). Com duas copies num doc (`fontes`):
`scripts/fixtures/dm_bf_antecipacao.config.json` (Black da Daniela: trailer + ADS 02 a 05, gravados juntos).

Projeto com mais de um doc de gravação (Black: antecipação, captação…): uma pasta de dados por doc, com a fase no
fim (`gravacao-DM-BF-POS-NOV-26-antecipacao`), e `pasta_upload`/`doc.nome` com a fase no nome.

| Chave | O que é |
|---|---|
| `projeto` | sigla `[XX][TIPO][PROD][26]`; substitui `{projeto}` em `drive.pasta_upload` e `doc.nome` |
| `expert.nome` / `expert.curto` | nome completo (links.md) e curto ("Vídeos brutos da Daniela (upload)") |
| `clickup.workspace_id` | 9013012202 |
| `clickup.copia.doc_id` / `page_id` | Doc e página da copy (`/docs/<doc>/<page>` na URL) |
| `fontes` | opcional, no lugar de `clickup.copia` + `copia`: várias copies num doc só, na ordem do doc, cada uma `{nome, doc_id, page_id, copia}`. Orientações (`manter`) de todas antes dos roteiros; roteiros na ordem das fontes; número repetido entre fontes aborta; cada página ganha o rodapé; `--md` vira um arquivo por fonte. Uma fonte só = mesmo modelo e mesmo hash de antes |
| `clickup.tarefa.id` / `inserir_apos` | task que recebe as 2 linhas; regex da linha depois da qual elas entram (sem achar ou achando 2: vão para o fim, com aviso). Sem `tarefa`: só o rodapé da copy |
| `clickup.tarefas` | no lugar de `tarefa`: várias tasks, cada uma `{id, inserir_apos, linhas?}`. `linhas` = `["upload", "doc"]` (padrão) ou só `["doc"]` (task de copy). Padrão Exos: task de copy só com o doc; task de vídeo/gravação/edição com pasta + doc |
| `copia.inicio` | regex que casa a linha-âncora **inteira** (texto sem `**` nem escapes). Obrigatória e única |
| `copia.fim` | regex da 1ª linha depois da âncora que fecha a seção (`=+` para o `\====` do ClickUp). O que vem depois (inclusive o rodapé que a skill acrescenta) não é lido |
| `copia.adaptar` | `[{rotulo, padrao, troca, vezes}]` sobre o md da seção; cada regra tem de casar exatamente `vezes` vezes |
| `copia.descartar` | rótulos (prefixo) de linhas internas que ficam de fora, em qualquer ponto da seção |
| `copia.manter` | `[{rotulo, dividir_em?, cta?, opcional?}]` rótulos (prefixo) das orientações. `dividir_em` quebra a linha em parágrafos nos rótulos dados (na ordem); `cta: true` marca a linha do CTA. Cada um aparece 1× (salvo `opcional`). O doc segue a ordem da copy |
| `copia.itens` | `titulo` (regex com o número no grupo 1, ex. `ADS (\d{2}) \[.+\]`), `numeros` (exatos, na ordem), `manter` (rótulos dentro do item, ex. `Headline:`), `texto` (rótulo que abre as falas; aceita complemento entre parênteses, ex. "Texto para gravação (guia):", que vai para o doc como está), `termina_no_cta` |
| `copia.proibido` | regex (sem caixa) de termos internos que não podem ir para a expert. Somada à base: `https?://`, `clickup`, `](`, `**`, escape sobrando, `[inserir`. Não pôr "meta": a copy fala em "revisão da Meta" |
| `drive.projeto_id` | pasta `[SIGLA]` do projeto |
| `drive.copy` / `drive.bruto` | `{id, prefixo}` de `1.Copy…` e `14.Arquivo Bruto…`; o ID é conferido (pasta, viva, filha do projeto, prefixo). Sem ID: acha pelo prefixo (1 só) |
| `drive.categoria` | `6. Criativos ADS de Captação` (sem acesso público) |
| `drive.pasta_upload` / `papel_pasta_upload` | `Criativos Vídeo {projeto}` / `fileOrganizer` (gestor de conteúdo: a expert sobe e apaga) |
| `drive.papel_doc` | `writer` (editor) |
| `doc.nome` / `doc.titulo_orientacoes` | `Copy Criativos Vídeo {projeto}` / `ORIENTAÇÕES DE GRAVAÇÃO` |

## Parser (allow-list)

Linha a linha na seção: título de item (só negrito, casa `itens.titulo`) → item novo; rótulo em `descartar` → fora;
dentro de item: `itens.manter` antes do `itens.texto`, depois só falas (negrito interno pode, rótulo `X:` não);
antes dos itens: só rótulos de `manter`. Qualquer outra coisa aborta com a linha no erro, assim como lista, título
markdown, tabela, imagem, `*`/`_`/crase sem escape. Depois: itens = `numeros`, cada item com seus rótulos 1× e
falas, última fala = CTA, nenhum termo proibido. O hash (`estado.copy_hash`) é do modelo lido, não do md bruto.

## Modelo do cabeçalho (logo)

`assets/modelo_gravacao.docx` = só o cabeçalho do doc feito à mão na Black da Elen (`143AxLub…`): logo Exos no
header, HR, título, HR, "DRIVE para subir arquivos:", link, HR. Congelado porque o doc de origem é uma copy viva
(mudou em 25/09/26 e quebrou a checagem rígida do script antigo). `assets/modelo_gravacao.json` guarda fonte,
data e sha. `doc_gravacao.py modelo [--doc-id ID]` recongela: valida o cabeçalho por posição, corta o resto e
tira comentários e revisões. O doc final tem HR = 4, H2 = 1, H3 = nº de itens.

## Drive

- Token: OAuth do rclone (remote `gdrive`, conta `drive@exosmkt.com`, organizer dos drives compartilhados). A
  **Google Docs API dá 403** nesse projeto OAuth: o conteúdo entra por DOCX (upload multipart como
  `application/vnd.google-apps.document`); o Google preserva header com logo, H2/H3, HR e negrito.
- **Atualizar sem mudar o link:** `PATCH upload/drive/v3/files/{id}?uploadType=media` com o DOCX (`--update`).
- **Herança:** `anyone` numa pasta-mãe de drive compartilhado vaza para as filhas e não dá para rebaixar no
  filho. Por isso a categoria fica privada e só a subpasta abre. O motor aborta se `14.Arquivo Bruto` ou a
  categoria tiverem acesso público.
- Criação não é repetida às cegas: em 5xx/timeout, reconfere pelo nome antes de tentar de novo. Busca por nome
  normalizado (NFC, caixa, espaços); 2 achados = erro.
- `403 rateLimitExceeded` vem da cota do client_id compartilhado do rclone; o motor espera e repete.
- Expert costuma ser membro do drive compartilhado do cliente (a Daniela é `fileOrganizer`). O link aberto
  serve para quem não está logado nessa conta (assistente, outro e-mail).

## ClickUp

- Leitura e escrita pela REST com o token do Keychain `clickup-exos-api`; GET repete em 429/5xx, PUT só em 429.
- **Copy:** `PUT /api/v3/workspaces/{ws}/docs/{doc}/pages/{page}` com `content_edit_mode: append` e
  `\n\n**Documento Gravação Expert:** [url](url)` (formato da Elen), em cada página de `fontes`. Pula se o ID do
  doc já estiver lá; se já houver rodapé de outro doc, não acrescenta e avisa. Relê: 1 rodapé, conteúdo anterior
  preservado.
- **Link na copy:** o parser aborta em link markdown e a base de proibidos barra `https?://`. Referência útil para
  a expert (ex.: reel de exemplo) entra por `adaptar` como texto sem esquema, com o `_` escapado na troca
  (`instagram.com/reel/DO\\_OCN…`), senão o `runs()` acusa marcação.
- **Task:** `GET /api/v2/task/{id}?include_markdown_description=true` → insere `Vídeos brutos da <curto>
  (upload): [url](url)` e `Documento Gravação Expert: <nome do doc> — [url](url)` depois de `inserir_apos` → relê a
  descrição logo antes de gravar (aborta se mudou) → `PUT /api/v2/task/{id}` com `markdown_content` → relê. Backup em
  `historico/tarefa_<id>_<data>.json`. O nome do doc entrou em 19/10/26 (uma task pode receber vários docs); as linhas
  antigas, sem nome (Clube da Daniela, Edição da antecipação da Black), ficam como estão.
- **Vários docs na mesma task** (ex.: os 3 da captação de uma Black, com pasta de upload única): cada linha nova
  entra logo depois de `inserir_apos`, então rode os configs na ordem inversa da copy. Na task de vídeo, use como
  âncora a própria linha "Vídeos brutos …": na 1ª rodada ela não existe e as linhas vão para o fim (com aviso); nas
  seguintes, o doc entra logo abaixo dela.
- O ClickUp **escapa `_` no texto do link** (`1kiYIslbU\_kof…`): toda busca por ID é feita com e sem escape.
- Regravar a descrição **achata menção de Doc** (chip) em texto + link. O conteúdo fica; só o visual muda.

## Verificação (`--apply` e `verificar`)

Export `text/plain` do doc: título, "DRIVE…", URL da pasta e todas as linhas, na ordem; nenhum termo proibido.
Export DOCX: HR/H2/H3 e imagem no header. Permissões: doc `writer`, subpasta `fileOrganizer`, categoria sem
público. Acesso anônimo (sem sessão) ao doc e à pasta, e PDF do doc em `preview/doc_google_p*.png`. Falhou:
saída 1 e nada de estado nem ClickUp. Saídas: 0 ok · 1 abortou · 2 copy mudou sem `--update` · 3 doc editado à mão.

## Prévia local

LibreOffice headless (`/Applications/LibreOffice.app/Contents/MacOS/soffice`) com perfil próprio passado como
**URI codificada** (`Path.as_uri()`): caminho com espaço cru ("Daniela Moleiro") derruba o soffice sem PDF.
PNG pelo PyMuPDF.

## Extensão futura

"Doc-casca por peça" da Black (vários docs sem corpo, um por FASE, cada um com a sua pasta de upload; precedente
`Exos/drive_docs_gravacao_et_bf.py`): exigiria uma lista de docs por config. Não implementado.
