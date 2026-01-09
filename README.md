# CasaCred - DADOS_BI

Automacao de agentes para extracao de dados de credito e geracao de planilhas de saida.
O fluxo principal roda todos os agentes em sequencia e grava os resultados em `03_OUTPUT`
com logs em `05_LOGS`.

## Requisitos

- Python 3.10+ (recomendado)
- Dependencias do `requirements.txt`
- Poppler (para PDF) e Tesseract (OCR), se necessario

## Instalacao

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Como rodar

Opcao 1 (Python):

```bash
python 02_SCRIPTS/00_RUN_ALL.py
```

Opcao 2 (PowerShell):

```powershell
02_SCRIPTS\00_RUN_ALL.ps1
```

Antes de rodar, ajuste no `02_SCRIPTS/00_RUN_ALL.py`:
- `POPPLER_BIN` (caminho do poppler)
- `TESSERACT_EXE` (se usar OCR)
- flags como `DEBUG_OCR`, `STOP_ON_ERROR` e `PROCESS_ONLY_CHANGED`

## Estrutura de pastas

- `01_INPUT/` entradas por empresa (nao versionado)
- `02_SCRIPTS/` scripts Python
- `03_OUTPUT/` saidas geradas (nao versionado)
- `05_LOGS/` logs de execucao (nao versionado)

## SharePoint (opcional)

Se `UPLOAD_TO_SHAREPOINT = True`, configure as variaveis de ambiente:

- `SP_TENANT_ID`
- `SP_CLIENT_ID`
- `SP_CLIENT_SECRET`
- `SP_SITE_URL`
- `SP_LIBRARY_NAME`
- `SP_FOLDER_PATH`

## Boas praticas

- Nao versionar dados sensiveis nem saidas.
- Use as pastas `01_INPUT/`, `03_OUTPUT/` e `05_LOGS/` apenas localmente.
