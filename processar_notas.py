#!/usr/bin/env python3
"""
Processamento de Notas Fiscais - Dominio.Web vs ERP (multi-empresa)
Gera arquivos Parquet em ./data/ para deploy no Firebase.

Uso:
    py processar_notas.py
    py -m http.server 8000   -> http://localhost:8000
"""

import re, os, json, difflib
from pathlib import Path

import msal, requests
import pandas as pd
import pdfplumber
import pyarrow as pa
import pyarrow.parquet as pq
from python_calamine import CalamineWorkbook

BASE_DIR      = Path(__file__).parent
XLS_DIR       = BASE_DIR / "notas_excel"
DOWNLOADS_DIR = BASE_DIR / "downloads_cfop"
DATA_DIR      = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

MES_INICIO = (2026, 6, 1)   # primeiro dia do mês a reconciliar
MES_FIM    = (2026, 6, 30)  # último  dia do mês a reconciliar

SISTEMA_EMPRESA = {
    "diprotec":  "DIPROTEC",
    "mafise":    "MAFISE",
    "tecnisul":  "TECNISUL",
    "revestech": "REVESTECH",
    "geo":       "GEO",
    "spazio":    "SPAZIO",
    "cda":       "CASADAGUA",
    "flexotom":  "FLEXOTOM",
}

EXCEL_PATH = Path.home() / "Downloads" / "CODIFICAÇÃO EMPRESAS.xlsx"

# ─── Azure / Power BI ────────────────────────────────────────────────────────
TENANT_ID    = "171d6e99-c6df-422f-9edf-dd2003e591ad"
CLIENT_ID    = "32d59780-3020-461e-bc70-f12bb2091966"
WORKSPACE_ID = "1903160a-9801-4f96-af8a-30d4db9abc4f"
DATASET_ID   = "a560d7d4-b9bb-4f25-b505-060475bceabd"
TABELA_ERP   = "querys_apoio relatorio_natureza"
AUTHORITY    = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE        = ["https://analysis.windows.net/powerbi/api/Dataset.Read.All"]
CACHE_FILE   = BASE_DIR / ".pbi_token_cache.json"
API_BASE     = "https://api.powerbi.com/v1.0/myorg"


def _pbi_headers() -> dict:
    cache = msal.SerializableTokenCache()
    if CACHE_FILE.exists():
        cache.deserialize(CACHE_FILE.read_text())
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPE, account=accounts[0]) if accounts else None
    if not result:
        print("        Abrindo browser para login no Azure...")
        result = app.acquire_token_interactive(scopes=SCOPE)
    if "access_token" not in result:
        raise RuntimeError(f"Erro de autenticação: {result.get('error_description')}")
    if cache.has_state_changed:
        CACHE_FILE.write_text(cache.serialize())
    return {"Authorization": f"Bearer {result['access_token']}"}


def _dax(headers: dict, query: str) -> list:
    r = requests.post(
        f"{API_BASE}/groups/{WORKSPACE_ID}/datasets/{DATASET_ID}/executeQueries",
        headers=headers,
        json={"queries": [{"query": query}], "serializerSettings": {"includeNulls": True}},
        timeout=120,
    )
    r.raise_for_status()
    return r.json()["results"][0]["tables"][0].get("rows", [])

# ─── Helpers ─────────────────────────────────────────────────────────────────

def br_float(value):
    if pd.isna(value) or str(value).strip() == "":
        return 0.0
    try:
        return float(str(value).replace(".", "").replace(",", "."))
    except ValueError:
        return 0.0


def normalize_cfop(v) -> str:
    """Normaliza CFOP para string 4 digitos sem hifen: '2-353' ou 2353 -> '2353'"""
    return re.sub(r"[^0-9]", "", str(v))[:4]


def _nota_candidates(nota_int: int) -> list:
    """Retorna possíveis numeros de NF (strippando 1-3 dígitos de série do final)."""
    result = [nota_int]
    n = nota_int
    for _ in range(3):
        if n < 10:
            break
        n = n // 10
        result.append(n)
    return result


def _load_empresa_apelidos() -> dict:
    """Carrega CODIFICAÇÃO EMPRESAS.xlsx e retorna dict {empresa_id_str: apelido}."""
    if not EXCEL_PATH.exists():
        return {}
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=0)
        df.columns = [c.strip() for c in df.columns]
        codigo_col  = next((c for c in df.columns if c.upper().startswith("C") and "D" in c.upper()), df.columns[0])
        apelido_col = next((c for c in df.columns if "APELIDO" in c.upper()), df.columns[1])
        result = {}
        for _, row in df.iterrows():
            try:
                cod = str(int(float(row[codigo_col])))
                ap  = str(row[apelido_col]).strip()
                if ap and ap.lower() not in ("nan", ""):
                    result[cod] = ap
            except (ValueError, TypeError):
                pass
        return result
    except Exception as e:
        print(f"    WARNING: Excel empresas nao carregado: {e}")
        return {}


# ─── 1. Dominio.Web — Azure (querys_apoio relatorio_natureza) ────────────────

def parse_dominio_web() -> pd.DataFrame:
    """Busca todas as entradas do ERP (Acesse) via modelo semântico do Azure."""
    headers = _pbi_headers()
    print(f"        Consultando Azure: '{TABELA_ERP}'...")
    rows = _dax(headers, f"""
EVALUATE
FILTER(
    '{TABELA_ERP}',
    VAR dEntrada = '{TABELA_ERP}'[DATA]
    VAR dEmissao = '{TABELA_ERP}'[EMISSAO]
    VAR entradaEmJunho = NOT(ISBLANK(dEntrada)) && dEntrada >= DATE(2026, 6, 1) && dEntrada <= DATE(2026, 6, 30)
    VAR semEntradaEmissaoMaioJunho = ISBLANK(dEntrada) && NOT(ISBLANK(dEmissao)) && dEmissao >= DATE(2026, 5, 1) && dEmissao <= DATE(2026, 6, 30)
    RETURN entradaEmJunho || semEntradaEmissaoMaioJunho
)
""")
    print(f"        {len(rows)} linhas recebidas")

    # Normaliza nomes de colunas: remove prefixo da tabela
    prefix = f"{TABELA_ERP}["
    clean  = [{k.replace(prefix, "").rstrip("]"): v for k, v in r.items()} for r in rows]
    df = pd.DataFrame(clean)

    df["data"]       = pd.to_datetime(df["DATA"], errors="coerce")
    _emissao         = pd.to_datetime(df["EMISSAO"], errors="coerce")
    df["data"]       = df["data"].where(df["data"].notna(), _emissao)
    df = df[df["data"].notna() & (df["data"] >= "2000-01-01")]
    _total   = pd.to_numeric(df["TOTAL"],   errors="coerce")
    _totalnf = pd.to_numeric(df.get("TOTALNF", pd.Series(dtype=float)), errors="coerce").reindex(df.index)
    df["total"]   = _total.fillna(_totalnf).fillna(0.0)            # TOTAL por item; TOTALNF fallback quando nulo
    df["totalnf"] = _totalnf.fillna(_total).fillna(0.0)            # TOTALNF da NF (valor total declarado)
    df["numero_int"] = pd.to_numeric(df["NUMERO"], errors="coerce").astype("Int64")
    _cfop_raw  = df["CFOP"].fillna("").astype(str)
    _cfopnf_raw = df.get("CFOPNF", pd.Series("", index=df.index)).fillna("").astype(str)
    df["cfop_norm"]  = _cfop_raw.where(_cfop_raw != "", _cfopnf_raw).apply(normalize_cfop)
    df["natureza"]      = df["NATUREZA"].fillna("").astype(str)
    df["estado"]        = df["ESTADO"].fillna("").astype(str)
    df["loja"]          = df["LOJA"].fillna("").astype(str)
    df["fornecedor"]    = df["FORNECEDOR"].fillna("").astype(str)
    df["empresa"]       = df["EMPRESA"].fillna("").astype(str) if "EMPRESA" in df.columns else ""
    df["tipo_entrada"]  = df["TIPO_ENTRADA_OC"].fillna("").astype(str) if "TIPO_ENTRADA_OC" in df.columns else ""
    df["fonte"]         = "Acesse"

    # Exclui entradas canceladas do Acesse: natureza contém "CANCEL" ou começa com "XCANCEL"
    # Não exclui "XDEVOLUCAO" e similares — devolução de compras é entrada válida
    _nat = df["NATUREZA"].fillna("").astype(str).str.upper()
    df = df[~(_nat.str.contains("CANCEL"))]

    # Agrupa por nota+CFOP: soma TOTAL dos itens (cada item tem seu TOTAL individual)
    grp_keys = ["numero_int", "loja", "cfop_norm"]
    agg = (
        df.groupby(grp_keys, dropna=False)
        .agg(
            data=("data",          "first"),
            fornecedor=("fornecedor", "first"),
            total=("total",        "sum"),
            totalnf=("totalnf",    "first"),   # TOTALNF da NF (igual em todos os itens)
            natureza=("natureza",  "first"),
            estado=("estado",      "first"),
            empresa=("empresa",    "first"),
            tipo_entrada=("tipo_entrada", "first"),
            fonte=("fonte",        "first"),
        )
        .reset_index()
    )
    return agg[["empresa", "numero_int", "data", "loja", "fornecedor", "total", "totalnf",
                "cfop_norm", "natureza", "estado", "fonte", "tipo_entrada"]]


# ─── 1b. Acesse — XLSX baixados pelo scraper ─────────────────────────────────

def parse_acesse_xlsx() -> pd.DataFrame:
    """Lê os XLSX de downloads_cfop/ (scraper do Acesse) no mesmo schema de parse_dominio_web."""
    xlsx_files = sorted(DOWNLOADS_DIR.glob("**/*.xlsx"))
    print(f"      {len(xlsx_files)} XLSX encontrados em downloads_cfop/")

    _COLS = ["Pedido","Loja","Data","Numero","Emissao","Fornecedor","Total",
             "Base Icms","Aliq Icms","Valor Icms","Base Ipi","Ipi","Base Subs",
             "Valor Subs","Frete","Seguro","Base Pis","Pis","Cofins",
             "Outros Creditos","Partilha Icms","Frete Custo","Cfop",
             "Referente Nf","Natureza","Mes","Cst","Estado","Cnpj"]

    d_ini = pd.Timestamp(*MES_INICIO)
    d_fim = pd.Timestamp(*MES_FIM)

    all_records = []
    for xlsx_path in xlsx_files:
        sistema = xlsx_path.parent.name
        if sistema not in SISTEMA_EMPRESA:
            continue   # ignora pastas desconhecidas (ex: diprotec_req, runs antigos)
        empresa = SISTEMA_EMPRESA[sistema]
        try:
            wb   = CalamineWorkbook.from_path(str(xlsx_path))
            rows = list(wb.get_sheet_by_index(0).to_python())
            if len(rows) < 2:
                continue
            header = [str(h).strip() if h else "" for h in rows[0]]
            def idx(name):
                try: return header.index(name)
                except ValueError: return None
            def get(row, name, default=None):
                i = idx(name)
                if i is None or i >= len(row): return default
                v = row[i]
                return v if v is not None else default

            count = 0
            for row in rows[1:]:
                numero_raw = get(row, "Numero")
                if not numero_raw:
                    continue
                try:
                    numero_int = int(float(str(numero_raw).strip()))
                except (ValueError, TypeError):
                    continue

                # Data de entrada (Data) — sempre preenchida no XLSX do Acesse
                data_raw = get(row, "Data")
                if isinstance(data_raw, str):
                    data_val = pd.to_datetime(data_raw, dayfirst=True, errors="coerce")
                else:
                    data_val = pd.Timestamp(data_raw) if data_raw else pd.NaT

                # Filtra pelo mês de reconciliação pela data de entrada
                if pd.isna(data_val) or not (d_ini <= data_val <= d_fim):
                    continue

                cfop_raw  = str(get(row, "Cfop", "") or "").strip()
                cfop_norm = normalize_cfop(cfop_raw) if cfop_raw else ""

                try:
                    total_val = float(get(row, "Total", 0.0) or 0.0)
                except (ValueError, TypeError):
                    total_val = 0.0

                all_records.append({
                    "empresa":      empresa,
                    "numero_int":   numero_int,
                    "data":         data_val,
                    "loja":         str(get(row, "Loja", "") or "").strip(),
                    "fornecedor":   str(get(row, "Fornecedor", "") or "").strip(),
                    "total":        total_val,
                    "cfop_norm":    cfop_norm,
                    "natureza":     str(get(row, "Natureza", "") or "").strip(),
                    "estado":       str(get(row, "Estado", "") or "").strip(),
                    "fonte":        "Acesse",
                    "tipo_entrada": "",
                })
                count += 1

            print(f"        {xlsx_path.name:52s}  ->  {count:4d} linhas  ({empresa})")
        except Exception as e:
            print(f"        {xlsx_path.name:52s}  ->  ERRO: {e}")

    if not all_records:
        return pd.DataFrame()

    df = pd.DataFrame(all_records)
    df["numero_int"] = pd.to_numeric(df["numero_int"], errors="coerce").astype("Int64")

    # Agrupa por (empresa, numero_int, loja, cfop_norm) — mesmo schema do parse_dominio_web
    grp_keys = ["empresa", "numero_int", "loja", "cfop_norm"]
    agg = (
        df.groupby(grp_keys, dropna=False)
        .agg(
            data=("data",          "first"),
            fornecedor=("fornecedor", "first"),
            total=("total",        "sum"),
            natureza=("natureza",  "first"),
            estado=("estado",      "first"),
            fonte=("fonte",        "first"),
        )
        .reset_index()
    )
    # totalnf = soma de todas as linhas da nota (para matching multi-cfop)
    totalnf_map = (
        agg[agg["numero_int"].notna()]
        .groupby(["empresa", "numero_int"])["total"]
        .sum()
        .to_dict()
    )
    agg["totalnf"]       = agg.apply(
        lambda r: totalnf_map.get((r["empresa"], r["numero_int"]), r["total"])
        if pd.notna(r["numero_int"]) else r["total"], axis=1
    )
    agg["tipo_entrada"] = ""

    return agg[["empresa", "numero_int", "data", "loja", "fornecedor", "total", "totalnf",
                "cfop_norm", "natureza", "estado", "fonte", "tipo_entrada"]]


# ─── 2. ERP — todos os PDFs ──────────────────────────────────────────────────

_CODE_DATE  = re.compile(r"^(\d+)(\d{2}/\d{2}/\d{4})$")   # código de tamanho variável
_CODNAME    = re.compile(r"^(\d+)([A-Z].*)$")
# CFOP: 1-152, 1.152, 1-152PR, 1.152SP (hífen ou ponto, com ou sem UF colada)
_CFOP_RE    = re.compile(r"^(\d)[.\-](\d{3})([A-Z]{2})?$")
# CFOP sem separador: 1152, 1152PR
_CFOP_BARE  = re.compile(r"^(\d)(\d{3})([A-Z]{2})$")   # exige UF para diferenciar de número comum
_VAL_TYPE   = re.compile(r"^([\d.,]+)(ICMS|IPI|SUBTRI|DIFALI?|ST)?$")
_EMPRESA_ID = re.compile(r"Entradas_(\d+)_")
_CNPJ_RE   = re.compile(r"CNPJ[:\s]*([\d./-]+)")


def _extract_empresa_meta(pdf_path: Path, apelidos: dict) -> tuple[str, str, str]:
    """Retorna (empresa_id, empresa_label, cnpj) a partir do arquivo PDF."""
    m = _EMPRESA_ID.search(pdf_path.name)
    empresa_id = m.group(1) if m else "?"

    with pdfplumber.open(pdf_path) as pdf:
        words = pdf.pages[0].extract_words()

    # CNPJ: palavra com "/" que contenha dígitos
    cnpj = next(
        (w["text"] for w in words if "/" in w["text"] and re.search(r"\d{2}\.\d{3}", w["text"])),
        ""
    )

    # Apelido do Excel (preferido), senão nome do PDF, senão empresa_id
    if empresa_id in apelidos:
        empresa_label = apelidos[empresa_id]
    else:
        pagina_x = next((w["x0"] for w in words if re.search(r"gina", w["text"], re.I)), 999)
        nome = " ".join(w["text"] for w in words if w["top"] < 10 and w["x0"] < pagina_x).strip()
        empresa_label = f"{nome} ({cnpj})" if cnpj else nome or empresa_id

    return empresa_id, empresa_label, cnpj


def _words_to_record(words: list) -> dict | None:
    rec = {
        "codigo": None, "data": None, "nota": None, "especie": None,
        "cod_fornecedor": None, "fornecedor": None, "cfop": None, "cfop_norm": None,
        "uf": None, "valor": 0.0, "tipo": None,
        "base_calculo": 0.0, "aliq": 0.0, "valor_icms": 0.0,
        "isentas": 0.0, "outras": 0.0,
    }
    words = sorted(words, key=lambda w: w["x0"])
    nome_parts = []
    found_data = False
    val_prefix = ""  # dígito(s) inteiros antes do ponto decimal, quando o PDF divide o número

    for w in words:
        x, text = w["x0"], w["text"]

        if x < 9:
            continue
        elif x < 95:
            m = _CODE_DATE.match(text)
            if m:
                rec["codigo"] = m.group(1)
                rec["data"]   = m.group(2)
                found_data    = True
            elif x >= 80 and rec["nota"] is None and text.isdigit():
                # Nota+série longa começa antes de x=95 (ex: 1260563521 para série 21)
                rec["nota"] = text
        elif x < 152:
            if rec["nota"] is None and text.isdigit():
                rec["nota"] = text
        elif x < 190:
            if rec["especie"] is None and text.isdigit():
                rec["especie"] = text
        elif x < 270:
            m = _CODNAME.match(text)
            if m and rec["cod_fornecedor"] is None:
                rec["cod_fornecedor"] = m.group(1)
                nome_parts.append(m.group(2))
            elif text and text[0].isalpha():
                nome_parts.append(text)
        elif x < 295:
            m = _CFOP_RE.match(text)
            if m:
                rec["cfop"]      = f"{m.group(1)}-{m.group(2)}"
                rec["cfop_norm"] = normalize_cfop(f"{m.group(1)}{m.group(2)}")
                if m.group(3) and rec["uf"] is None:
                    rec["uf"] = m.group(3)
            elif rec["cfop"] is None:
                m2 = _CFOP_BARE.match(text)
                if m2:
                    rec["cfop"]      = f"{m2.group(1)}-{m2.group(2)}"
                    rec["cfop_norm"] = normalize_cfop(f"{m2.group(1)}{m2.group(2)}")
                    if rec["uf"] is None:
                        rec["uf"] = m2.group(3)
                elif rec["uf"] is None:
                    uf_m = re.search(r"([A-Z]{2})$", text)
                    if uf_m:
                        rec["uf"] = uf_m.group(1)
        elif x < 342:
            # Valor contábil com sufixo de regime (ex: "36.063,58ICMS")
            # ou UF puro (ex: "PR", "SP")
            # ou valor intercalado com letras do nome (ex: "2L1T.D41A4,80ICMS" = LTDA + 21.414,80)
            # Alguns PDFs dividem o número em tokens adjacentes (ex: "PARA1" + "6C.O8N42S,T5R8UICCMS"
            # = "1" + "6.842,58ICMS" = 16.842,58). val_prefix acumula o prefixo inteiro.
            m = _VAL_TYPE.match(text)
            digits_only = re.sub(r"[^0-9.,]", "", text)
            if m and re.search(r"\d,\d{2}", m.group(1)):
                # Valor decimal completo — prepend prefixo se houver
                rec["valor"] = br_float(val_prefix + m.group(1))
                val_prefix = ""
                if m.group(2):
                    rec["tipo"] = m.group(2)
            elif m and re.fullmatch(r"[\d.]+", m.group(1)):
                # Token puramente numérico sem decimal — prefixo de número dividido
                val_prefix = m.group(1).replace(".", "")
            elif digits_only and re.search(r"\d,\d{2}$", digits_only):
                # Valor com letras intercaladas (ex: "6C.O8N42S,T5R8UICCMS" → "6.842,58")
                rec["valor"] = br_float(val_prefix + digits_only)
                val_prefix = ""
            elif digits_only and re.search(r",\d$", digits_only):
                # Decimal parcial: apenas 1 dígito após vírgula — 2º dígito vem no próximo token
                val_prefix = digits_only
            elif digits_only and re.fullmatch(r"\d+", digits_only) and text[-1].isdigit():
                # Texto misto que termina em dígito(s) inteiros (ex: "PARA1" → "1")
                # Pode ser o prefixo do valor que vem no próximo token
                val_prefix = digits_only
            else:
                val_prefix = ""
                if rec["uf"] is None:
                    uf_m = re.search(r"([A-Z]{2})$", text)
                    if uf_m:
                        rec["uf"] = uf_m.group(1)
        elif x < 415:
            m = _VAL_TYPE.match(text)
            if rec["valor"] == 0.0 and val_prefix and re.search(r",\d$", val_prefix):
                # Completa decimal parcial: val_prefix tem 1 decimal, este token fornece o 2º
                cont = re.sub(r"[^0-9]", "", text)
                if re.fullmatch(r"\d", cont):
                    rec["valor"] = br_float(val_prefix + cont)
                    val_prefix = ""
                elif m:
                    num = m.group(1)
                    if re.search(r",\d{2}$", num):
                        rec["valor"] = br_float(num)
                    else:
                        rec["valor"] = br_float(val_prefix + num)
                    val_prefix = ""
                    if m.group(2):
                        rec["tipo"] = m.group(2)
            elif m:
                if rec["valor"] == 0.0:
                    # Fallback: valor contábil sem sufixo (ex: "36.063,58")
                    # Valida 2 decimais para rejeitar tokens garbled (ex: "70,18440ICMS")
                    num = m.group(1)
                    if re.search(r",\d{2}$", num):
                        rec["valor"] = br_float(val_prefix + num)
                        val_prefix = ""
                        if m.group(2):
                            rec["tipo"] = m.group(2)
                else:
                    # Valor contábil já capturado em x<342 — este é a base de cálculo
                    rec["base_calculo"] = br_float(m.group(1))
        elif x < 435:
            rec["base_calculo"] = br_float(text)
        elif x < 478:
            rec["aliq"]         = br_float(text)
        elif x < 524:
            rec["valor_icms"]   = br_float(text)
        elif x < 563:
            rec["isentas"]      = br_float(text)
        else:
            rec["outras"]       = br_float(text)

    if not found_data:
        return None
    rec["fornecedor"] = " ".join(nome_parts).strip() or None

    # Fallback: CFOP interleaved with fornecedor name (PDF columns overlap).
    # Pattern constraints: first digit 1-3 (valid entry CFOPs); 3-digit code starts 1-9;
    # not followed by comma (prevents matching BR value format like 1.990,00 or 4.000,00).
    _CFOP_PAT = re.compile(r"([1-3])[.\-]+([1-9]\d{2})(?!,)")
    if rec["cfop"] is None:
        zone = sorted([w for w in words if 210 <= w["x0"] < 340], key=lambda w: w["x0"])
        def _clean(s: str) -> str:
            """Strip letters/parens/slashes, then remove CNPJ-style dots between digits."""
            s = re.sub(r"[A-Za-z()/]", "", s)
            s = re.sub(r"(\d)\.(\d)", r"\1\2", s)  # e.g. 2-1.02 → 2-102
            return s

        # Pass 1: single tokens (common case — CFOP fused with one name word)
        for w in zone:
            m = _CFOP_PAT.search(_clean(w["text"]))
            if m:
                rec["cfop"]      = f"{m.group(1)}-{m.group(2)}"
                rec["cfop_norm"] = normalize_cfop(f"{m.group(1)}{m.group(2)}")
                break
        # Pass 2: concatenate zone tokens (CFOP split across 2-3 adjacent tokens)
        if rec["cfop"] is None and zone:
            concat = _clean("".join(w["text"] for w in zone))
            m = _CFOP_PAT.search(concat)
            if m:
                rec["cfop"]      = f"{m.group(1)}-{m.group(2)}"
                rec["cfop_norm"] = normalize_cfop(f"{m.group(1)}{m.group(2)}")

    return rec


def _is_data_row(words: list) -> bool:
    return any(9 <= w["x0"] < 95 and _CODE_DATE.match(w["text"]) for w in words)


def _parse_single_pdf(pdf_path: Path, empresa_id: str,
                      empresa_nome: str, cnpj: str) -> list[dict]:
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            raw_words = page.extract_words(keep_blank_chars=False)
            row_groups: dict[int, list] = {}
            for w in raw_words:
                key = round(w["top"] / 2) * 2
                row_groups.setdefault(key, []).append(w)
            for key in sorted(row_groups.keys()):
                rw = row_groups[key]
                if not _is_data_row(rw):
                    continue
                rec = _words_to_record(rw)
                if rec:
                    rec["empresa_id"]   = empresa_id
                    rec["empresa_nome"] = empresa_nome
                    rec["cnpj_empresa"] = cnpj
                    records.append(rec)
    return records


def parse_erp_all_pdfs() -> pd.DataFrame:
    all_records = []
    pdf_files = sorted(BASE_DIR.glob("Entradas_*.pdf"))
    print(f"      {len(pdf_files)} PDFs encontrados")
    apelidos = _load_empresa_apelidos()
    print(f"      {len(apelidos)} empresas carregadas do Excel de codificacao")

    for pdf_path in pdf_files:
        try:
            empresa_id, empresa_nome, cnpj = _extract_empresa_meta(pdf_path, apelidos)
            recs = _parse_single_pdf(pdf_path, empresa_id, empresa_nome, cnpj)
            all_records.extend(recs)
            print(f"        {pdf_path.name:35s}  ->  {len(recs):4d} linhas  ({empresa_nome[:30]})")

        except Exception as e:
            print(f"        {pdf_path.name:35s}  ->  ERRO (arquivo ignorado): {e}")

    df = pd.DataFrame(all_records)
    if df.empty:
        return df

    df["data"]     = pd.to_datetime(df["data"], format="%d/%m/%Y", errors="coerce")
    df["nota_int"] = pd.to_numeric(df["nota"], errors="coerce").astype("Int64")
    df["fonte"]    = "Dominio.Web"
    return df


# ─── 2b. ERP — XLS (alternativa ao PDF, mais confiável) ─────────────────────

def _detect_xls_cols(header_row: list) -> dict:
    """Detecta posições de colunas dinamicamente a partir do header do XLS."""
    cols = {}
    valor_cont_idx = None
    for i, v in enumerate(header_row):
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        sl = s.lower()
        if s == "Nota":
            cols["nota"] = i
        elif s == "CFOP":
            cols["cfop"] = i
        elif s == "UF":
            cols["uf"] = i
        elif s == "Tipo":
            cols["tipo"] = i
        elif s == "Isentas":
            cols["isentas"] = i
        elif s == "Outras":
            cols["outras"] = i
        elif s == "Fornecedor":
            cols["fornecedor"] = i
        elif s.startswith("Data") and "Entrada" not in s and "data" not in cols:
            cols["data"] = i
        elif s.startswith("Valor Cont") or sl.startswith("valor cont"):
            cols["valor"] = i
            valor_cont_idx = i
        elif s.startswith("Base C") or sl.startswith("base c"):
            cols["base_calculo"] = i
        elif (s.startswith("Al") or sl.startswith("al")) and ("q." in sl or "q" in sl) and "base" not in sl:
            cols["aliq"] = i
        elif s == "Valor" and valor_cont_idx is not None:
            cols["valor_icms"] = i
        elif s == "C�igo" or s.startswith("C") and "digo" in s:
            if "codigo" not in cols:
                cols["codigo"] = i
    return cols

def _xls_get(row: list, cols: dict, key: str, default=None):
    idx = cols.get(key)
    if idx is None or idx >= len(row):
        return default
    v = row[idx]
    return v if v is not None else default

def _xls_flt(row: list, cols: dict, key: str) -> float:
    v = _xls_get(row, cols, key)
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0

def parse_erp_xls() -> pd.DataFrame:
    """Lê todos os XLS em notas_excel/ e retorna DataFrame no mesmo schema do PDF."""
    all_records = []
    xls_files = sorted(XLS_DIR.glob("*.xls"))
    print(f"      {len(xls_files)} XLS encontrados")
    apelidos = _load_empresa_apelidos()
    print(f"      {len(apelidos)} empresas carregadas do Excel de codificacao")

    for xls_path in xls_files:
        m = re.search(r"(\d+)\.xls$", xls_path.name, re.IGNORECASE)
        if not m:
            continue
        empresa_id   = m.group(1)
        empresa_nome = apelidos.get(empresa_id, xls_path.stem)

        try:
            wb    = CalamineWorkbook.from_path(str(xls_path))
            sheet = wb.get_sheet_by_index(0)
            rows  = list(sheet.to_python())

            # CNPJ: procura nas 5 primeiras linhas
            cnpj = ""
            for row in rows[:5]:
                for v in row:
                    if v and re.search(r"\d{14}", re.sub(r"[^0-9]", "", str(v))):
                        cnpj = re.sub(r"[^0-9]", "", str(v))
                        break

            # Localiza linha do header (contém 'Nota') e detecta posições
            header_idx = next(
                (i for i, row in enumerate(rows) if any(str(v) == "Nota" for v in row if v)),
                5
            )
            cols = _detect_xls_cols(rows[header_idx])

            records = []
            for row in rows[header_idx + 1:]:
                cod = _xls_get(row, cols, "codigo", row[0] if row else None)
                if not cod or str(cod).strip() in ("", "nan") or str(cod).startswith("Total"):
                    continue
                nota_raw = _xls_get(row, cols, "nota")
                if not nota_raw or str(nota_raw).strip() == "":
                    continue

                try:
                    nota_num = int(float(str(nota_raw).strip()))
                    cfop_raw = _xls_get(row, cols, "cfop")
                    cfop_int = int(float(str(cfop_raw))) if cfop_raw else 0
                    cfop_str  = f"{cfop_int // 1000}-{cfop_int % 1000:03d}" if cfop_int else ""
                    cfop_norm = normalize_cfop(str(cfop_int)) if cfop_int else ""

                    data_raw = _xls_get(row, cols, "data")
                    if isinstance(data_raw, str):
                        data_val = pd.to_datetime(data_raw, dayfirst=True, errors="coerce")
                    else:
                        data_val = pd.Timestamp(data_raw) if data_raw else pd.NaT

                    records.append({
                        "empresa_id":   empresa_id,
                        "empresa_nome": empresa_nome,
                        "cnpj_empresa": cnpj,
                        "codigo":       str(int(float(str(cod)))),
                        "data":         data_val,
                        "nota":         str(nota_num),
                        "nota_int":     nota_num,
                        "cfop":         cfop_str,
                        "cfop_norm":    cfop_norm,
                        "uf":           str(_xls_get(row, cols, "uf") or "").strip(),
                        "valor":        _xls_flt(row, cols, "valor"),
                        "tipo":         str(_xls_get(row, cols, "tipo") or "").strip() or None,
                        "base_calculo": _xls_flt(row, cols, "base_calculo"),
                        "aliq":         _xls_flt(row, cols, "aliq"),
                        "valor_icms":   _xls_flt(row, cols, "valor_icms"),
                        "isentas":      _xls_flt(row, cols, "isentas"),
                        "outras":       _xls_flt(row, cols, "outras"),
                        "fornecedor":   str(_xls_get(row, cols, "fornecedor") or "").strip(),
                    })
                except (ValueError, TypeError):
                    continue

            all_records.extend(records)
            print(f"        {xls_path.name:38s}  ->  {len(records):4d} linhas  ({empresa_nome[:28]})")

        except Exception as e:
            print(f"        {xls_path.name:38s}  ->  ERRO: {e}")

    df = pd.DataFrame(all_records)
    if df.empty:
        return df
    df["nota_int"] = pd.to_numeric(df["nota_int"], errors="coerce").astype("Int64")
    df["fonte"]    = "Dominio.Web"
    return df


# ─── 2c. Deduplicação DW ─────────────────────────────────────────────────────

def _dedup_dw(df: pd.DataFrame) -> pd.DataFrame:
    """Remove linhas duplicadas do DW antes da comparação:
    1. Linha cfop='' cujo total coincide (±0,01) com alguma linha cfop≠'' da mesma nota
       → é linha de sumário do Acesse, redundante frente à linha de CFOP específico
    2. Linhas com (empresa, numero_int, cfop_norm, total) idênticos → mantém a com natureza
    """
    if df.empty:
        return df
    df = df.copy()

    # Passo 1: linha cfop="" com mesmo total que linha cfop-específica → remover
    specific = df[df["cfop_norm"].str.strip().ne("") & df["numero_int"].notna()]
    specific_totals = (
        specific.groupby(["empresa", "numero_int"])["total"]
        .apply(lambda s: set(round(float(v), 2) for v in s))
        .to_dict()
    )

    def _is_summary(row):
        if row["cfop_norm"].strip() != "":
            return False
        if pd.isna(row["numero_int"]):
            return False
        key = (row["empresa"], row["numero_int"])
        return round(float(row["total"]), 2) in specific_totals.get(key, set())

    df = df[~df.apply(_is_summary, axis=1)]

    # Passo 2: duplicatas exatas (empresa, numero, cfop, total) → preferir a com natureza
    df["_nat_ok"] = df["natureza"].str.strip().ne("")
    df = df.sort_values("_nat_ok", ascending=False)
    df = df.drop_duplicates(subset=["empresa", "numero_int", "cfop_norm", "total"], keep="first")
    df = df.drop(columns=["_nat_ok"])

    return df


# ─── 3. Comparacao ────────────────────────────────────────────────────────────

def _similar_supplier(a: str, b: str) -> bool:
    """Retorna True se os nomes de fornecedor são similares (≥60% ou um contém o outro)."""
    if not a or not b:
        return False
    def norm(s):
        return re.sub(r'[^A-Z0-9 ]', '', s.upper().strip())
    a, b = norm(a), norm(b)
    if not a or not b:
        return False
    if a in b or b in a:
        return True
    return difflib.SequenceMatcher(None, a, b).ratio() >= 0.60

VALOR_TOLERANCE = 0.50

def build_comparison(df_dw: pd.DataFrame, df_erp: pd.DataFrame):
    # DW lookup: nota_num -> lista de valores por CFOP (já agregados em parse_dominio_web)
    # Exclui linhas com cfop vazio — são agregados (soma de todos os itens) que duplicam
    # o valor e confundem o matching. O totalnf dessas linhas ainda é usado via dw_nota_totalnf_vals.
    _dw_valid = df_dw[df_dw["numero_int"].notna()].copy()
    _dw_valid["_n"] = _dw_valid["numero_int"].astype(float).astype(int)
    _dw_cfop = _dw_valid[_dw_valid["cfop_norm"].notna() & (_dw_valid["cfop_norm"] != "")]
    dw_nota_vals: dict[int, list[float]] = (
        _dw_cfop.groupby("_n")["total"]
        .apply(lambda s: [float(v) for v in s.dropna()])
        .to_dict()
    )
    # Lookup: nota_num -> lista de todos os TOTALNF individuais (não só o máximo),
    # para casar ERP linha individual contra o totalnf de cada grupo do Acesse.
    dw_nota_totalnf_vals: dict[int, list[float]] = (
        _dw_valid[_dw_valid.get("totalnf", pd.Series(0.0, index=_dw_valid.index)) > 0]
        .groupby("_n")["totalnf"]
        .apply(lambda s: list({float(v) for v in s.dropna()}))
        .to_dict()
    ) if "totalnf" in _dw_valid.columns else {}
    # Mantém também o max para estratégias de soma
    dw_nota_totalnf: dict[int, float] = {
        n: max(vs) for n, vs in dw_nota_totalnf_vals.items()
    }

    # Pré-calcula soma das linhas ERP por (empresa_id, nota_int) — evita somar notas
    # homônimas de empresas diferentes (ex: nota 145 na Flexotom e na Spazio).
    _erp_valid = df_erp[df_erp["nota_int"].notna()].copy()
    _erp_valid["_n"] = _erp_valid["nota_int"].astype(float).astype(int)
    erp_nota_sum: dict = (
        _erp_valid
        .groupby(["empresa_id", "_n"])["valor"]
        .sum()
        .to_dict()  # chave: (empresa_id, nota_int_int)
    )

    def find_dw_nota(nota_int, erp_valor, empresa_id):
        """Retorna nota_num do DW que bate com nota_int (±série) E valor ±0,50.
        Verifica: TOTAL por CFOP individual, soma dos CFOPs, TOTALNF da NF,
        e soma das linhas ERP contra TOTALNF."""
        if pd.isna(nota_int):
            return None
        nota_int_int = int(nota_int)
        erp_val  = float(erp_valor) if pd.notna(erp_valor) else None
        erp_sum  = float(erp_nota_sum.get((empresa_id, nota_int_int), erp_val or 0))
        for candidate in _nota_candidates(nota_int_int):
            if candidate not in dw_nota_vals and candidate not in dw_nota_totalnf_vals:
                continue
            dw_vals         = dw_nota_vals.get(candidate, [])
            dw_totalnf      = dw_nota_totalnf.get(candidate)
            dw_totalnf_list = dw_nota_totalnf_vals.get(candidate, [])
            if erp_val is None:
                return candidate
            # 1. ERP linha vs TOTAL por CFOP
            if dw_vals and any(abs(erp_val - dv) <= VALOR_TOLERANCE for dv in dw_vals):
                return candidate
            # 2. ERP linha vs soma de todos os CFOPs do Acesse
            if dw_vals and abs(erp_val - sum(dw_vals)) <= VALOR_TOLERANCE:
                return candidate
            # 3. ERP linha vs TOTALNF individual de cada grupo do Acesse
            if any(abs(erp_val - tv) <= VALOR_TOLERANCE for tv in dw_totalnf_list):
                return candidate
            # 4. Soma das linhas ERP vs TOTALNF máximo do Acesse
            if dw_totalnf is not None and abs(erp_sum - dw_totalnf) <= VALOR_TOLERANCE:
                return candidate
            # 5. Soma das linhas ERP vs soma dos CFOPs do Acesse
            if dw_vals and abs(erp_sum - sum(dw_vals)) <= VALOR_TOLERANCE:
                return candidate
        return None

    df_erp = df_erp.copy()
    df_erp["nota_num"] = df_erp.apply(
        lambda r: find_dw_nota(r["nota_int"], r["valor"], r.get("empresa_id", "")), axis=1
    )

    # Conjunto de números DW que têm match no ERP (nota + valor)
    ambos = set(df_erp["nota_num"].dropna().astype(int))

    # CFOP lookup: dw_num -> set[cfop] no DW
    dw_cfop_map = (
        df_dw[df_dw["cfop_norm"].notna()]
        .groupby(df_dw["numero_int"].astype("Int64").astype(float).astype("Int64"))["cfop_norm"]
        .apply(set)
        .to_dict()
    )
    # CFOP lookup: dw_num -> set[cfop] no ERP (usando nota_num)
    erp_valid = df_erp[df_erp["nota_num"].notna() & df_erp["cfop_norm"].notna()].copy()
    erp_cfop_map = (
        erp_valid.groupby(erp_valid["nota_num"].astype(int))["cfop_norm"]
        .apply(set)
        .to_dict()
    )

    def tag_dw(row):
        n = row["numero_int"]
        if pd.isna(n):
            return "Sem numero"
        n = int(n)
        if n not in ambos:
            return "So no Acesse"
        erp_cfops = erp_cfop_map.get(n, set())
        if erp_cfops and row["cfop_norm"] and row["cfop_norm"] not in erp_cfops:
            return "Ambos - CFOP divergente"
        return "Ambos - OK"

    def tag_erp(row):
        dw_match = row["nota_num"]
        if dw_match is None or (isinstance(dw_match, float) and pd.isna(dw_match)):
            if pd.isna(row["nota_int"]):
                return "Sem numero"
            if float(row.get("valor") or 0) == 0.0:
                return "Cancelado - Dominio.Web"
            return "So no Dominio.Web"
        dw_match = int(dw_match)
        dw_cfops = dw_cfop_map.get(dw_match, set())
        if dw_cfops and pd.notna(row["cfop_norm"]) and row["cfop_norm"] not in dw_cfops:
            return "Ambos - CFOP divergente"
        return "Ambos - OK"

    df_dw  = df_dw.copy()
    df_dw["status"]  = df_dw.apply(tag_dw,  axis=1)
    df_erp["status"] = df_erp.apply(tag_erp, axis=1)

    # ── Detecta "Valor Divergente": número casado + fornecedor similar + valor diferente ──
    # Lookups: numero → lista de {valor, fornecedor}
    _dw_num_info: dict[int, list] = {}
    for _, r in df_dw[df_dw["numero_int"].notna()].iterrows():
        _dw_num_info.setdefault(int(r["numero_int"]), []).append({
            "total": float(r["total"] or 0),
            "fornecedor": str(r.get("fornecedor") or ""),
        })

    _erp_num_info: dict[int, list] = {}
    for _, r in df_erp[df_erp["nota_int"].notna()].iterrows():
        for cand in _nota_candidates(int(r["nota_int"])):
            _erp_num_info.setdefault(cand, []).append({
                "valor": float(r["valor"] or 0),
                "fornecedor": str(r.get("fornecedor") or ""),
            })

    def _check_valor_div_dw(row):
        if row["status"] != "So no Acesse" or pd.isna(row["numero_int"]):
            return row["status"]
        n = int(row["numero_int"])
        erp_list = _erp_num_info.get(n, [])
        forn = str(row.get("fornecedor") or "")
        if any(_similar_supplier(forn, e["fornecedor"]) for e in erp_list):
            return "Valor Divergente"
        return row["status"]

    def _check_valor_div_erp(row):
        if row["status"] != "So no Dominio.Web" or pd.isna(row["nota_int"]):
            return row["status"]
        forn = str(row.get("fornecedor") or "")
        for cand in _nota_candidates(int(row["nota_int"])):
            dw_list = _dw_num_info.get(cand, [])
            if any(_similar_supplier(forn, d["fornecedor"]) for d in dw_list):
                return "Valor Divergente"
        return row["status"]

    df_dw["status"]  = df_dw.apply(_check_valor_div_dw,  axis=1)
    df_erp["status"] = df_erp.apply(_check_valor_div_erp, axis=1)

    # Contagens refinadas
    n_ok         = int((df_dw["status"] == "Ambos - OK").sum())
    n_cfop_div   = int((df_dw["status"] == "Ambos - CFOP divergente").sum())
    n_so_dw      = int((df_dw["status"] == "So no Acesse").sum())
    n_so_erp_col = int((df_erp["status"] == "So no Dominio.Web").sum())
    n_valor_div  = int((df_dw["status"] == "Valor Divergente").sum())

    # DataFrame de coincidencias (join por nota_num + filtro de valor)
    if ambos:
        cols_dw  = ["numero_int", "data", "loja", "fornecedor",
                    "total", "cfop_norm", "natureza", "status"]
        cols_erp = ["nota_num", "nota_int", "data", "empresa_nome", "fornecedor",
                    "valor", "cfop_norm", "uf", "status"]

        merge_dw  = df_dw[df_dw["numero_int"].isin(ambos)][cols_dw].rename(columns={
            "numero_int": "nota_num", "data": "data_dw", "fornecedor": "forn_dw",
            "total": "valor_dw", "cfop_norm": "cfop_dw", "status": "status_dw",
        })
        merge_erp = df_erp[df_erp["nota_num"].notna() & df_erp["nota_num"].isin(ambos)][cols_erp].rename(columns={
            "nota_int": "nota_serie", "data": "data_erp", "empresa_nome": "empresa",
            "fornecedor": "forn_erp", "valor": "valor_erp",
            "cfop_norm": "cfop_erp", "status": "status_erp",
        })
        merge_erp["nota_num"] = merge_erp["nota_num"].astype(int)
        merge_dw["nota_num"]  = merge_dw["nota_num"].astype(int)

        df_joined = merge_dw.merge(merge_erp, on="nota_num", how="outer")
        # Mantém só pares onde valor bate dentro da tolerância (ou um dos lados sem valor)
        df_match = df_joined[
            df_joined["valor_dw"].isna() |
            df_joined["valor_erp"].isna() |
            ((df_joined["valor_dw"] - df_joined["valor_erp"]).abs() <= VALOR_TOLERANCE)
        ].copy()

        df_match["cfop_ok"] = df_match["cfop_erp"].isna() | (df_match["cfop_dw"] == df_match["cfop_erp"])
        df_match["status"]  = df_match.apply(
            lambda r: "Ambos - CFOP divergente" if not r["cfop_ok"] else "Ambos - OK", axis=1
        )
    else:
        df_match = pd.DataFrame(columns=["nota_num", "status"])

    def fmt_period(col, df):
        valid = df[col].dropna()
        if valid.empty:
            return "—"
        return f"{valid.min().strftime('%d/%m/%Y')} a {valid.max().strftime('%d/%m/%Y')}"

    stats = {
        "total_dw":    int(df_dw["numero_int"].nunique()),
        "total_erp":   int(df_erp["nota_int"].nunique()),
        "ambos_ok":    n_ok,
        "cfop_div":    n_cfop_div,
        "so_dw":       n_so_dw,
        "so_erp":      n_so_erp_col,
        "valor_div":   n_valor_div,
        "periodo_dw":  fmt_period("data", df_dw),
        "periodo_erp": fmt_period("data", df_erp),
        "linhas_dw":   len(df_dw),
        "linhas_erp":  len(df_erp),
        "num_empresas": int(df_erp["empresa_id"].nunique()) if "empresa_id" in df_erp.columns else 0,
        "arquivo_dw":  TABELA_ERP,
    }
    return df_dw, df_erp, df_match, stats


# ─── 4. Exportar Parquet ──────────────────────────────────────────────────────

def _to_parquet(df: pd.DataFrame, name: str):
    path = DATA_DIR / f"{name}.parquet"
    df = df.copy()
    for col in df.select_dtypes(include=["datetime64[ns]"]).columns:
        df[col] = df[col].dt.strftime("%Y-%m-%d")
    for col in df.select_dtypes(include=["Int64"]).columns:
        df[col] = df[col].astype("float64")
    table = pa.Table.from_pandas(df, preserve_index=False)
    pq.write_table(table, path, compression="snappy")
    size_kb = path.stat().st_size // 1024
    print(f"  OK {name}.parquet  --  {len(df)} linhas  ({size_kb} KB)")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  Processamento de Notas - Acesse vs Dominio.Web")
    print("=" * 60)

    CFOP_EXCLUIR = {"1933", "2933"}

    _acesse_xlsx = DOWNLOADS_DIR.exists() and any(DOWNLOADS_DIR.glob("**/*.xlsx"))
    if _acesse_xlsx:
        print("\n[1/4] Lendo Acesse (XLSX do scraper)...")
        df_dw = parse_acesse_xlsx()
    else:
        print(f"\n[1/4] Lendo Acesse (Azure: '{TABELA_ERP}')...")
        df_dw = parse_dominio_web()

    df_dw = _dedup_dw(df_dw)
    df_dw = df_dw[~df_dw["cfop_norm"].isin(CFOP_EXCLUIR)]
    # Exclui canceladas
    _nat = df_dw["natureza"].fillna("").astype(str).str.upper()
    df_dw = df_dw[~_nat.str.contains("CANCEL")]
    print(f"      {len(df_dw)} linhas  |  {df_dw['numero_int'].nunique()} NFs unicas")

    if XLS_DIR.exists() and any(XLS_DIR.glob("*.xls")):
        print("\n[2/4] Lendo Dominio.Web (XLS)...")
        df_erp = parse_erp_xls()
    else:
        print("\n[2/4] Lendo Dominio.Web (PDFs)...")
        df_erp = parse_erp_all_pdfs()
    df_erp = df_erp[~df_erp["cfop_norm"].isin(CFOP_EXCLUIR)]
    print(f"      Total: {len(df_erp)} linhas  |  {df_erp['nota_int'].nunique()} NFs unicas")

    print("\n[3/4] Comparando...")
    df_dw, df_erp, df_match, stats = build_comparison(df_dw, df_erp)

    # Com XLSX do scraper a data de entrada já filtra corretamente — sem ajuste necessário.
    # Com Azure (fallback): remove "Só no Acesse" de meses anteriores que não têm ERP de junho
    if not _acesse_xlsx:
        _data_col = pd.to_datetime(df_dw["data"], errors="coerce")
        df_dw = df_dw[
            (df_dw["status"] != "So no Acesse") |
            (_data_col >= "2026-06-01")
        ].copy()

    # Recalcula stats a partir do df_dw final
    stats["total_dw"]  = int(df_dw["numero_int"].nunique())
    stats["ambos_ok"]  = int((df_dw["status"] == "Ambos - OK").sum())
    stats["cfop_div"]  = int((df_dw["status"] == "Ambos - CFOP divergente").sum())
    stats["so_dw"]     = int((df_dw["status"] == "So no Acesse").sum())
    stats["valor_div"] = int((df_dw["status"] == "Valor Divergente").sum())

    print(f"\n  Periodo DW  : {stats['periodo_dw']}")
    print(f"  Periodo ERP : {stats['periodo_erp']}")
    print(f"  Empresas ERP: {stats['num_empresas']}")
    print(f"  NFs unicas DW       : {stats['total_dw']}")
    print(f"  NFs unicas ERP      : {stats['total_erp']}")
    print(f"  Ambos - OK          : {stats['ambos_ok']}")
    print(f"  Ambos - CFOP diverg.: {stats['cfop_div']}")
    print(f"  So no Acesse        : {stats['so_dw']}")
    print(f"  So no Dominio.Web   : {stats['so_erp']}")

    print("\n[4/4] Gerando Parquet em ./data/ ...")
    _to_parquet(df_dw,   "dominio_web")
    _to_parquet(df_erp,  "erp_entradas")
    _to_parquet(df_match, "coincidencias")
    _to_parquet(pd.DataFrame([stats]), "resumo")

    print("\n" + "=" * 60)
    print("  Concluido!")
    print("  Abra o painel com:")
    print("    py -m http.server 8000")
    print("    -> http://localhost:8000")
    print("=" * 60)



def debug_pdf_columns(pdf_path: str, max_rows: int = 5):
    """Imprime posições x e valores das primeiras linhas de dados de um PDF."""
    from pathlib import Path
    import pdfplumber
    path = Path(pdf_path)
    print(f"\nDiagnóstico: {path.name}")
    print("-" * 80)
    rows_found = 0
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            raw_words = page.extract_words(keep_blank_chars=False)
            row_groups: dict[int, list] = {}
            for w in raw_words:
                key = round(w["top"] / 2) * 2
                row_groups.setdefault(key, []).append(w)
            for key in sorted(row_groups.keys()):
                rw = row_groups[key]
                if not _is_data_row(rw):
                    continue
                print(f"\n  Linha {rows_found + 1}:")
                for w in sorted(rw, key=lambda w: w["x0"]):
                    print(f"    x={w['x0']:6.1f}  text={w['text']}")
                rows_found += 1
                if rows_found >= max_rows:
                    return


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--debug-pdf":
        pdf = sys.argv[2] if len(sys.argv) > 2 else next(Path(".").glob("Entradas_*.pdf"), None)
        if pdf:
            debug_pdf_columns(str(pdf))
        else:
            print("Nenhum PDF encontrado.")
    else:
        main()
