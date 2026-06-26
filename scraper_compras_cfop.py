"""
Scraper para download do Relatorio de Compras (CFOP) - sistemas Diprotec.
Usa requests puro (sem browser). Muito mais rapido que a versao com Playwright.

Uso:
    py scraper_compras_cfop.py                          # todos os sistemas, todas as lojas
    py scraper_compras_cfop.py --sistema diprotec       # so um sistema
    py scraper_compras_cfop.py --sistema diprotec --loja BH
    py scraper_compras_cfop.py --inicio 01/05/2026 --fim 31/05/2026
    py scraper_compras_cfop.py --destino C:\\Planilhas\\CFOP

Sistemas:
    diprotec, mafise, tecnisul, revestech, geo, spazio, cda, flexotom
"""

import argparse
import re
import sys
from datetime import date
from pathlib import Path

import requests

USUARIO = "thiago.santos"
SENHA   = "santos081"

SISTEMAS = {
    "diprotec":  "http://sistema.diprotec.com.br:44451/diprotec",
    "mafise":    "http://sistema.diprotec.com.br:44451/mafise",
    "tecnisul":  "http://sistema.diprotec.com.br:44450/tecnisul",
    "revestech": "http://sistema.diprotec.com.br:44451/revestech",
    "geo":       "http://sistema.diprotec.com.br:44451/diprotecGeo",
    "spazio":    "http://sistema.diprotec.com.br:44451/spazio",
    "cda":       "http://sistema.diprotec.com.br:44450/cda",
    "flexotom":  "https://novoflex.lcdc.net.br:43440/novotom",
}

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120"}


# ---------------------------------------------------------------------------
# Login e inicializacao da sessao
# ---------------------------------------------------------------------------

def criar_sessao(base_url):
    """Faz login e retorna (requests.Session, nomeSessao)."""
    sess = requests.Session()
    sess.headers.update(HEADERS)
    today = date.today().strftime("%d/%m/%Y")

    # Passo 1 — pega PHPSESSID inicial
    sess.get(f"{base_url}/autentica.php", timeout=15)

    # Passo 2 — envia credenciais
    sess.post(f"{base_url}/autentica.php", timeout=15, data={
        "dataAtual": today, "nomeSessao": "",
        "nomeUsuario": USUARIO, "senha": SENHA,
        "resolucao": "1920", "browserWidth": "1920",
        "browserHeight": "1080", "Login": "",
    })

    # Passo 3 — segue o redirect JS para index.php e extrai nomeSessao
    r = sess.get(f"{base_url}/index.php", timeout=15)
    m = re.search(r"sessionStorage\.tabID\s*=\s*[\"'](s[a-z0-9]+)", r.text)
    if not m:
        raise RuntimeError(f"Login falhou em {base_url} (nomeSessao nao encontrado)")
    ns = m.group(1)

    # Passo 4 — registra a sessao no servidor
    sess.get(f"{base_url}/index.php", params={"ns": "1", "nomeSessao": ns}, timeout=15)

    return sess, ns


def inicializar_pagina_relatorio(sess, base_url, ns):
    """Carrega o frameset e cor.php para inicializar o objeto PHP da sessao."""
    sess.get(f"{base_url}/rel_emnf.php", params={"nomeSessao": ns}, timeout=15)
    sess.get(f"{base_url}/cor.php",      params={"nomeSessao": ns}, timeout=15)
    sess.get(f"{base_url}/rel_emnf.php", params={"nomeSessao": ns, "LerTela": "1"}, timeout=15)


# ---------------------------------------------------------------------------
# Detecta lojas disponíveis
# ---------------------------------------------------------------------------

def detectar_lojas(sess, base_url, ns):
    r = sess.get(f"{base_url}/rel_emnf.php",
                 params={"nomeSessao": ns, "LerTela": "1"}, timeout=15)
    # Extrai especificamente o <select name="loja">
    m = re.search(
        r'<select[^>]+name=["\']?loja["\']?[^>]*>(.*?)</select>',
        r.text, re.DOTALL | re.IGNORECASE
    )
    if not m:
        return {}
    select_html = m.group(1)
    # Options podem nao ter fechamento (HTML mal-formado no sistema)
    lojas = {}
    for value, text in re.findall(r'<option[^>]+value="([^"]+)"[^>]*>([^<]+)', select_html):
        v = value.strip()
        t = text.strip()
        if v:
            lojas[v] = t
    return lojas


# ---------------------------------------------------------------------------
# Download de um relatorio
# ---------------------------------------------------------------------------

def download_relatorio(sess, base_url, ns, sistema, loja_codigo, loja_nome,
                        data_inicio, data_fim, pasta_destino):
    today = date.today().strftime("%d/%m/%Y")

    # POST para gerar o relatorio
    r_post = sess.post(f"{base_url}/rel_emnf.php", timeout=60, data={
        "dataAtual": today, "nomeSessao": ns,
        "tipoRelatorio": "cfop",
        "dataInicial": data_inicio,
        "dataFinal":   data_fim,
        "loja": loja_codigo,
        "natureza": "", "cfop": "", "codPessoa": "", "nomePessoa": "",
        "ordem": "dataPedido", "quebra1": "", "quebra2": "",
        "sumario": "N", "Procurar": "Procurar",
    })

    # Extrai a URL do relatorio gerado (chama_relEmnf.php?...&pgt=...)
    m = re.search(r'(chama_relEmnf\.php\?nomeSessao=[^"\']+&pgt=\d+)', r_post.text)
    if not m:
        raise RuntimeError(f"URL do relatorio nao encontrada. Response: {r_post.text[:200]}")

    rel_path = m.group(1)
    rel_url   = f"{base_url}/{rel_path}"
    excel_url = f"{rel_url}&excel=1"

    # Carrega o relatorio HTML primeiro (o servidor renderiza e armazena internamente)
    r_html = sess.get(rel_url, timeout=120)

    # Download do XLSX
    r_excel = sess.get(excel_url, timeout=120)
    ct = r_excel.headers.get("content-type", "")
    if r_excel.status_code != 200 or "spreadsheet" not in ct:
        raise RuntimeError(f"Resposta inesperada ao baixar Excel: {r_excel.status_code} {ct}")

    tag_loja  = (loja_codigo or "TODAS").replace("/", "_")
    nome_arq  = f"{sistema}_{tag_loja}_{data_inicio.replace('/','_')}_{data_fim.replace('/','_')}.xlsx"
    caminho   = pasta_destino / nome_arq

    with open(caminho, "wb") as f:
        f.write(r_excel.content)

    return caminho


# ---------------------------------------------------------------------------
# Processa um sistema completo
# ---------------------------------------------------------------------------

def processar_sistema(nome_sistema, base_url, lojas_filtro, data_inicio, data_fim, pasta_base):
    print(f"\n{'='*55}")
    print(f"  {nome_sistema.upper()}  ({base_url})")
    print(f"{'='*55}")

    arquivos = []
    try:
        sess, ns = criar_sessao(base_url)
        print(f"  Logado. ns={ns}")

        inicializar_pagina_relatorio(sess, base_url, ns)

        lojas = detectar_lojas(sess, base_url, ns)
        if not lojas:
            print("  Aviso: nenhuma loja detectada — baixando sem filtro.")
            lojas = {"": "TODAS"}
        else:
            print(f"  Lojas: {', '.join(f'{k}={v}' for k,v in lojas.items())}")

        if lojas_filtro:
            lojas = {k: v for k, v in lojas.items() if k.upper() in lojas_filtro}
            if not lojas:
                print(f"  Aviso: filtro de lojas nao encontrado neste sistema.")
                return []

        pasta = pasta_base / nome_sistema
        pasta.mkdir(parents=True, exist_ok=True)

        for codigo, nome in lojas.items():
            print(f"  Loja: {nome} ({codigo or 'todas'})... ", end="", flush=True)
            try:
                # Re-inicializa o contexto do relatorio a cada download
                inicializar_pagina_relatorio(sess, base_url, ns)
                caminho = download_relatorio(
                    sess, base_url, ns, nome_sistema,
                    codigo, nome, data_inicio, data_fim, pasta
                )
                print(f"OK  ({caminho.stat().st_size//1024} KB)")
                arquivos.append(caminho)
            except Exception as e:
                print(f"ERRO: {e}")

    except Exception as e:
        print(f"  ERRO no sistema {nome_sistema}: {e}")

    return arquivos


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(args):
    pasta_base  = Path(args.destino).expanduser().resolve()
    pasta_base.mkdir(parents=True, exist_ok=True)

    data_inicio = args.inicio or date.today().replace(day=1).strftime("%d/%m/%Y")
    data_fim    = args.fim    or date.today().strftime("%d/%m/%Y")

    if args.sistema:
        nomes = [s.strip().lower() for s in args.sistema.split(",")]
        invalidos = [s for s in nomes if s not in SISTEMAS]
        if invalidos:
            print(f"Sistema(s) desconhecido(s): {', '.join(invalidos)}")
            print(f"Disponiveis: {', '.join(SISTEMAS)}")
            sys.exit(1)
        sistemas = {k: SISTEMAS[k] for k in nomes}
    else:
        sistemas = SISTEMAS

    lojas_filtro = {l.strip().upper() for l in args.loja.split(",")} if args.loja else set()

    print(f"Periodo : {data_inicio} ate {data_fim}")
    print(f"Sistemas: {', '.join(sistemas)}")
    print(f"Destino : {pasta_base}")

    total = []
    for nome, url in sistemas.items():
        total.extend(processar_sistema(nome, url, lojas_filtro, data_inicio, data_fim, pasta_base))

    print(f"\n{'='*55}")
    print(f"Concluido. {len(total)} arquivo(s) baixado(s):")
    for f in total:
        print(f"  {f}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Download Relatorio de Compras CFOP — sistemas Diprotec (sem browser)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--sistema", help="Sistema(s) separados por virgula. Padrao: todos")
    parser.add_argument("--loja",    help="Codigo(s) de loja separados por virgula. Padrao: todas")
    parser.add_argument("--inicio",  help="Data inicial DD/MM/AAAA (padrao: 1o do mes)")
    parser.add_argument("--fim",     help="Data final DD/MM/AAAA (padrao: hoje)")
    parser.add_argument("--destino", default="downloads_cfop", help="Pasta destino (padrao: downloads_cfop)")
    main(parser.parse_args())
