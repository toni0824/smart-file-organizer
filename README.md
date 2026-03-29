# Smart File Organizer

`smart-file-organizer` organiza ficheiros e pastas da `Downloads` para destinos configurados no teu computador.

Neste projeto, o destino principal está configurado para:

- `~/Documents/Tecnico/2 ano/...`
- `~/Documents/Tecnico/3 ano/...`
- `~/Documents/Tecnico/Por_Organizar`

O script combina:

- regras locais por palavras-chave
- leitura direta do conteúdo de ficheiros
- classificação com IA local via `Ollama`

## O que faz

- organiza ficheiros da `Downloads`
- processa também pastas dentro da `Downloads`
- agrupa imagens soltas (`jpg`, `jpeg`, `png`) em `Downloads/Imagens`
- lê conteúdo de `pdf`, `docx`, `pptx`, `xlsx`, `txt`, `py`, `html`, `csv` e `zip`
- usa IA local para ficheiros ambíguos
- tem um modo agressivo que usa IA em quase tudo
- pode deixar os itens não classificados em `Downloads` ou movê-los para `Por_Organizar`

## Requisitos

- Python 3.10+
- `Ollama` instalado e aberto
- modelo local, por exemplo:

```bash
ollama pull qwen2.5:1.5b
```

Pacotes Python usados para leitura de ficheiros:

```bash
python3 -m pip install --break-system-packages --user pypdf python-docx python-pptx openpyxl
```

## Como usar

Teste sem mover nada:

```bash
python3 path.py --dry-run
```

Usar regras locais + IA quando necessário:

```bash
python3 path.py --dry-run --use-ai
python3 path.py --use-ai
```

Usar IA em modo mais agressivo:

```bash
python3 path.py --dry-run --ai-aggressive
python3 path.py --ai-aggressive
```

Mover também os não classificados para `Por_Organizar`:

```bash
python3 path.py --use-ai --move-unsorted
```

Limitar o teste aos primeiros itens:

```bash
python3 path.py --dry-run --ai-aggressive --limit 20
```

Escolher outro modelo do Ollama:

```bash
python3 path.py --use-ai --model qwen2.5:1.5b
```

## Flags principais

- `--dry-run`: mostra o que seria movido
- `--use-ai`: usa IA apenas nos casos não resolvidos pelas regras
- `--ai-aggressive`: usa IA em quase todos os ficheiros e pastas válidos
- `--move-unsorted`: move o que não foi classificado para `Por_Organizar`
- `--no-directories`: ignora pastas dentro de `Downloads`
- `--limit N`: processa apenas os primeiros `N` itens
- `--model`: escolhe o modelo local do Ollama
- `--min-confidence`: confiança mínima para aceitar a decisão da IA
- `--max-text-chars`: quantidade máxima de texto enviada ao modelo

## Como classifica

1. tenta classificar por regras locais
2. extrai texto e metadados do ficheiro ou da pasta
3. se necessário, envia um resumo ao Ollama
4. move para a pasta destino se houver confiança suficiente
5. caso contrário, deixa em `Downloads` ou move para `Por_Organizar`

## Configuração

As regras e destinos estão definidos em [`path.py`](./path.py):

- `SUBJECT_RULES`
- `SUBJECT_DESTINATIONS`
- `BASE_FOLDER`
- `DEST_ROOT`

Se quiseres usar outras pastas, basta alterar esses valores.
