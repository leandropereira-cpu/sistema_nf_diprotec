"""
Verifica se há arquivos novos/atualizados no Google Drive desde a última execução.
Sai com código 0 se há atualizações (rodar pipeline), 1 se não há (pular).

Persiste o timestamp da última verificação em .last_drive_check (gerenciado pelo
cache do GitHub Actions entre execuções).
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path

TIMESTAMP_FILE  = Path(__file__).parent / '.last_drive_check'
DEFAULT_FOLDER  = '1WSI7dA669fWgLQ4XsosGs4ei6_Csg2Ic'


def main():
    creds_src = os.environ.get('GOOGLE_CREDENTIALS', '')
    if not creds_src:
        print('[check_drive] GOOGLE_CREDENTIALS nao definido — rodando pipeline por precaucao.')
        sys.exit(0)

    folder_id = os.environ.get('DRIVE_FOLDER_ID', DEFAULT_FOLDER)

    # Timestamp da última verificação (padrão: 24h atrás se arquivo não existe)
    if TIMESTAMP_FILE.exists():
        last_check = datetime.fromisoformat(TIMESTAMP_FILE.read_text().strip())
    else:
        last_check = datetime.now(timezone.utc) - timedelta(hours=24)

    print(f'[check_drive] Ultima verificacao: {last_check.strftime("%d/%m/%Y %H:%M:%S UTC")}')

    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    info  = json.loads(creds_src)
    creds = Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    svc = build('drive', 'v3', credentials=creds, cache_discovery=False)

    # Lista arquivos da pasta raiz e subpastas (1 nível)
    all_files = []
    result = svc.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields='files(id,name,mimeType,modifiedTime)',
        pageSize=200,
    ).execute()
    for f in result.get('files', []):
        all_files.append(f)
        if f['mimeType'] == 'application/vnd.google-apps.folder':
            sub = svc.files().list(
                q=f"'{f['id']}' in parents and trashed=false",
                fields='files(id,name,modifiedTime)',
                pageSize=200,
            ).execute()
            all_files.extend(sub.get('files', []))

    has_updates = False
    for f in all_files:
        modified = datetime.fromisoformat(f['modifiedTime'].replace('Z', '+00:00'))
        if modified > last_check:
            print(f'  ATUALIZADO: {f["name"]}  ({modified.strftime("%d/%m %H:%M UTC")})')
            has_updates = True

    # Salva novo timestamp antes de sair
    TIMESTAMP_FILE.write_text(datetime.now(timezone.utc).isoformat())

    if has_updates:
        print('[check_drive] Atualizacoes encontradas — iniciando pipeline.')
        sys.exit(0)
    else:
        print('[check_drive] Sem atualizacoes no Drive — aguardando proxima rotina.')
        sys.exit(1)


if __name__ == '__main__':
    main()
