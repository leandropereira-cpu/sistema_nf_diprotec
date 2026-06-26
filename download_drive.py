"""
Baixa arquivos do Google Drive para uso no pipeline.

Estrutura esperada no Drive:
  <DRIVE_FOLDER_ID>/        ← XLS do Domínio Web (notas_excel/)
  └── flexotom/             ← XLSX da Flexotom (downloads_cfop/flexotom/)

Credenciais: variável de ambiente GOOGLE_CREDENTIALS com o JSON
da conta de serviço (ou --credentials com o caminho do arquivo).

Uso:
    python download_drive.py
    python download_drive.py --credentials credenciais.json
    python download_drive.py --folder 1WSI7dA669fWgLQ4XsosGs4ei6_Csg2Ic
"""

import io
import json
import os
import argparse
from pathlib import Path

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES            = ['https://www.googleapis.com/auth/drive.readonly']
DEFAULT_FOLDER_ID = '1WSI7dA669fWgLQ4XsosGs4ei6_Csg2Ic'
BASE_DIR          = Path(__file__).parent

XLS_MIMES = [
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
]
FOLDER_MIME = 'application/vnd.google-apps.folder'


def _service(creds_src: str):
    if Path(creds_src).exists():
        info = json.loads(Path(creds_src).read_text())
    else:
        info = json.loads(creds_src)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds, cache_discovery=False)


def _list(svc, folder_id, mimes=None):
    q = f"'{folder_id}' in parents and trashed=false"
    if mimes:
        q += ' and (' + ' or '.join(f"mimeType='{m}'" for m in mimes) + ')'
    return svc.files().list(q=q, fields='files(id,name,mimeType)', pageSize=200).execute().get('files', [])


def _download(svc, file_id, dest: Path):
    req = svc.files().get_media(fileId=file_id)
    with io.FileIO(str(dest), 'wb') as fh:
        dl = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = dl.next_chunk()


def main(args):
    creds_src = args.credentials or os.environ.get('GOOGLE_CREDENTIALS', '')
    if not creds_src:
        raise RuntimeError(
            'Credenciais não encontradas.\n'
            'Defina a variável GOOGLE_CREDENTIALS (JSON da conta de serviço) '
            'ou use --credentials <arquivo.json>.'
        )

    folder_id = args.folder or os.environ.get('DRIVE_FOLDER_ID', DEFAULT_FOLDER_ID)
    svc = _service(creds_src)

    # ── XLS do Domínio Web → notas_excel/ ──────────────────────────────────────
    xls_dir = BASE_DIR / 'notas_excel'
    xls_dir.mkdir(exist_ok=True)

    xls_files = _list(svc, folder_id, mimes=XLS_MIMES)
    print(f'\n[Drive] {len(xls_files)} arquivo(s) XLS na pasta raiz → notas_excel/')
    for f in xls_files:
        dest = xls_dir / f['name']
        _download(svc, f['id'], dest)
        print(f'  ✓ {f["name"]}')

    # ── Subpasta flexotom/ → downloads_cfop/flexotom/ ──────────────────────────
    subfolders = _list(svc, folder_id, mimes=[FOLDER_MIME])
    flex = next((f for f in subfolders if f['name'].lower() == 'flexotom'), None)
    if flex:
        flex_dir = BASE_DIR / 'downloads_cfop' / 'flexotom'
        flex_dir.mkdir(parents=True, exist_ok=True)
        flex_files = _list(svc, flex['id'], mimes=XLS_MIMES)
        print(f'\n[Drive/flexotom] {len(flex_files)} XLSX → downloads_cfop/flexotom/')
        for f in flex_files:
            dest = flex_dir / f['name']
            _download(svc, f['id'], dest)
            print(f'  ✓ {f["name"]}')
    else:
        print('\n[Drive] Pasta flexotom/ não encontrada — sem XLSX da Flexotom.')

    print('\nDownload concluído.')


if __name__ == '__main__':
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--credentials', help='JSON da conta de serviço (string ou caminho do arquivo)')
    p.add_argument('--folder',      help='ID da pasta no Google Drive')
    main(p.parse_args())
