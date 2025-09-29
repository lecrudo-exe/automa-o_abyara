import ast
import os
import re
import time
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import pandas as pd
import requests

BASE_URL = "https://abyara.sigavi360.com.br"
TOKEN_URL = f"{BASE_URL}/Sigavi/api/Acesso/Token"
LEAD_URL = f"{BASE_URL}/Sigavi/api/Leads/NovaLead"

SIGAVI_USER = os.getenv("SIGAVI_USER", "integracao")
SIGAVI_PASS = os.getenv("SIGAVI_PASS", "USdZzIwAZcONM6y")

SIGAVI_PLANILHA_ENV = os.getenv("SIGAVI_PLANILHA")
SHEET_NAME = os.getenv("SIGAVI_ABA") or None
SIGAVI_TIPO_ENV = os.getenv("SIGAVI_TIPO")
INTERVALO_REQUESTS = float(os.getenv("SIGAVI_INTERVALO", "0.5"))
DRY_RUN = os.getenv("SIGAVI_DRY_RUN", "").lower() in {"1", "true", "sim"}

_TOKEN_CACHE: Dict[str, Any] = {"token": None, "exp": 0.0}
_SESSION = requests.Session()



def resolve_planilha() -> Path:
    candidates = []
    if SIGAVI_PLANILHA_ENV:
        candidates.append(Path(SIGAVI_PLANILHA_ENV).expanduser())
    candidates.append(Path('Clientes_Abyara_moema.xlsx'))
    candidates.append(Path('Leads ZAIT.xlsx'))
    for candidate in candidates:
        if candidate and candidate.exists():
            return candidate
    raise FileNotFoundError(
        'Nenhuma planilha encontrada. Defina SIGAVI_PLANILHA ou mantenha Clientes_Abyara_moema.xlsx / Leads ZAIT.xlsx na pasta.'
    )


PLANILHA_PATH = resolve_planilha()
TIPO_CADASTRO = SIGAVI_TIPO_ENV or 'Leads'

_tag_env = os.getenv('SIGAVI_TAG')
if _tag_env:
    TAG_ORIGEM = _tag_env
else:
    TAG_ORIGEM = PLANILHA_PATH.stem.replace('_', ' ')

_ref_env = os.getenv('SIGAVI_REFERENCIA')
if _ref_env:
    REFERENCIA_PADRAO = _ref_env
else:
    REFERENCIA_PADRAO = TAG_ORIGEM

_result_env = os.getenv('SIGAVI_RESULTADO')
if _result_env:
    OUTPUT_PATH = Path(_result_env)
else:
    OUTPUT_PATH = PLANILHA_PATH.with_name(f'resultado_{PLANILHA_PATH.stem}.csv')


def _read_file_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="latin-1")


def load_corretores_gerentes() -> Dict[str, str]:
    confio_path = Path(__file__).with_name("confio.py")
    if not confio_path.exists():
        print("Aviso: confio.py nao encontrado. Sem mapa de corretores.")
        return {}
    try:
        tree = ast.parse(_read_file_text(confio_path))
    except Exception as exc:
        print(f"Aviso: nao foi possivel interpretar confio.py ({exc}).")
        return {}
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "corretores_gerentes":
                    try:
                        data = ast.literal_eval(node.value)
                        if isinstance(data, dict):
                            return data
                    except Exception as exc:
                        print(f"Aviso: falha ao carregar corretores_gerentes ({exc}).")
                    return {}
    print("Aviso: corretores_gerentes nao encontrado em confio.py.")
    return {}


def normalize_corretor(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).upper()


def normalize_label(label: str) -> str:
    base = unicodedata.normalize('NFKD', str(label or ''))
    base = ''.join(ch for ch in base if not unicodedata.combining(ch))
    return re.sub(r'\s+', ' ', base).strip().upper()

def build_corretor_lookup(raw_map: Dict[str, str]) -> Dict[str, str]:
    lookup: Dict[str, str] = {}
    for key, val in raw_map.items():
        lookup[normalize_corretor(key)] = str(val).strip()
    return lookup


def only_digits(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    if text.lower() in {"nan", "none"}:
        return ""
    return re.sub(r"\D+", "", text)


def fmt_phone(value: Any) -> str:
    digits = only_digits(value)
    if len(digits) == 11:
        return f"({digits[0:2]}) {digits[2:7]}-{digits[7:11]}"
    if len(digits) == 10:
        return f"({digits[0:2]}) {digits[2:6]}-{digits[6:10]}"
    return str(value or "").strip()


def load_planilha(path: Path, sheet_name: Optional[str]) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f'Planilha nao encontrada: {path}')
    df = pd.read_excel(path, sheet_name=sheet_name or 0)
    if isinstance(df, dict):
        df = next(iter(df.values()))
    df.columns = [str(col).strip() for col in df.columns]

    column_lookup: Dict[str, str] = {}
    for col in df.columns:
        column_lookup.setdefault(col.upper(), col)
        column_lookup.setdefault(normalize_label(col), col)

    def pick(*candidates: str) -> str:
        for candidate in candidates:
            key_norm = normalize_label(candidate)
            if key_norm in column_lookup:
                return column_lookup[key_norm]
            key_upper = candidate.upper()
            if key_upper in column_lookup:
                return column_lookup[key_upper]
        raise KeyError(f'Coluna nao encontrada. Procurado: {candidates}')

    nome_col = pick('NOME', 'NOME COMPLETO', 'NOME DO CONTATO')
    telefone_col = pick('FONE2', 'TELEFONE', 'CELULAR')
    corretor_col = pick('CORRETOR DE ORIGEM', 'CORRETOR ORIGEM', 'INDICACAO CORRETOR', 'VENDEDOR')

    df = df[[nome_col, telefone_col, corretor_col]].copy()
    df.columns = ['nome', 'telefone', 'corretor_origem']
    df = df.dropna(how='all', subset=['nome', 'telefone'])
    df = df.reset_index(drop=True)
    return df


def build_payload(
    nome: str,
    telefone_fmt: str,
    telefone_digits: str,
    corretor_origem: str,
    gerente: Optional[str],
) -> Dict[str, Any]:
    parts = [
        "Cadastro gerado via API",
        f"Fonte: {TAG_ORIGEM}",
    ]
    if corretor_origem:
        parts.append(f"Corretor origem: {corretor_origem}")
    if gerente:
        parts.append(f"Gerente: {gerente}")
    mensagem = " | ".join(parts)

    payload: Dict[str, Any] = {
        "Tipo": TIPO_CADASTRO,
        "Nome": nome,
        "Telefone": telefone_fmt,
        "TelefoneNumero": telefone_digits,
        "Mensagem": mensagem,
        "Empreendimento": True,
        "OrigemDescricao": TAG_ORIGEM,
        "Referencia": REFERENCIA_PADRAO,
    }
    return {k: v for k, v in payload.items() if v}


def get_token() -> str:
    now = time.time()
    cached = _TOKEN_CACHE["token"]
    expira = _TOKEN_CACHE["exp"]
    if cached and now < expira - 60:
        return cached

    resp = _SESSION.post(
        TOKEN_URL,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={
            "username": SIGAVI_USER,
            "password": SIGAVI_PASS,
            "grant_type": "password",
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    token = data["access_token"]
    expires_in = float(data.get("expires_in", 1800))
    _TOKEN_CACHE["token"] = token
    _TOKEN_CACHE["exp"] = now + expires_in
    return token


def post_nova_lead(payload: Dict[str, Any]) -> Tuple[bool, int, Optional[Dict[str, Any]], str]:
    token = get_token()
    resp = _SESSION.post(
        LEAD_URL,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )
    try:
        data = resp.json()
    except ValueError:
        data = None
    sucesso = resp.status_code == 200 and isinstance(data, dict) and data.get("Sucesso") is True
    return sucesso, resp.status_code, data, resp.text


def main() -> None:
    print(f"Planilha selecionada: {PLANILHA_PATH} (aba: {SHEET_NAME or 0})")
    try:
        df = load_planilha(PLANILHA_PATH, SHEET_NAME)
    except Exception as exc:
        raise SystemExit(f"Erro ao carregar planilha: {exc}")

    corretores_raw = load_corretores_gerentes()
    corretores_lookup = build_corretor_lookup(corretores_raw)
    print(f"Corretores carregados: {len(corretores_lookup)}")
    print(f"Leads lidos: {len(df)}")

    resultados: List[Dict[str, Any]] = []
    telefones_vistos: Set[str] = set()

    for idx, row in df.iterrows():
        nome_raw = row["nome"]
        telefone_raw = row["telefone"]
        corretor_raw = row["corretor_origem"]

        nome = "" if pd.isna(nome_raw) else str(nome_raw).strip()
        telefone_digits = only_digits(telefone_raw)
        telefone_fmt = fmt_phone(telefone_raw)
        corretor = "" if pd.isna(corretor_raw) else str(corretor_raw).strip()
        gerente = corretores_lookup.get(normalize_corretor(corretor)) if corretor else None

        if not nome:
            resultados.append(
                {"linha": idx + 1, "nome": nome, "telefone": telefone_fmt, "status": "IGNORADO", "motivo": "nome vazio"}
            )
            continue
        if not telefone_digits:
            resultados.append(
                {"linha": idx + 1, "nome": nome, "telefone": telefone_fmt, "status": "IGNORADO", "motivo": "telefone vazio"}
            )
            continue
        if telefone_digits in telefones_vistos:
            resultados.append(
                {
                    "linha": idx + 1,
                    "nome": nome,
                    "telefone": telefone_fmt,
                    "status": "IGNORADO",
                    "motivo": "telefone duplicado no arquivo",
                }
            )
            continue

        telefones_vistos.add(telefone_digits)
        payload = build_payload(nome, telefone_fmt, telefone_digits, corretor, gerente)
        print(f"[{idx + 1}/{len(df)}] enviando {nome} - {telefone_fmt} (corretor: {corretor or 'sem info'})")

        if DRY_RUN:
            resultados.append(
                {"linha": idx + 1, "nome": nome, "telefone": telefone_fmt, "status": "DRY-RUN", "lead_id": "nao enviado"}
            )
            continue

        try:
            ok, status, data, raw_text = post_nova_lead(payload)
        except requests.RequestException as exc:
            resultados.append(
                {"linha": idx + 1, "nome": nome, "telefone": telefone_fmt, "status": "ERRO REDE", "motivo": str(exc)}
            )
            print(f"  erro de rede: {exc}")
            continue

        if ok:
            lead_id = data.get("Id") if data else None
            resultados.append(
                {"linha": idx + 1, "nome": nome, "telefone": telefone_fmt, "status": "OK", "lead_id": lead_id}
            )
            print(f"  ok - lead_id: {lead_id}")
        else:
            erros = None
            if data and isinstance(data.get("Erros"), list):
                erros = "; ".join(str(item) for item in data["Erros"])
            resultados.append(
                {
                    "linha": idx + 1,
                    "nome": nome,
                    "telefone": telefone_fmt,
                    "status": f"FALHA {status}",
                    "motivo": erros or raw_text[:300],
                }
            )
            print(f"  falha {status}: {erros or raw_text[:200]}")

        time.sleep(INTERVALO_REQUESTS)

    if resultados:
        out_df = pd.DataFrame(resultados)
        out_df.to_csv(OUTPUT_PATH, index=False, encoding="utf-8-sig")
        print(f"Relatorio salvo em: {OUTPUT_PATH.resolve()}")
    else:
        print("Nenhum registro processado.")


if __name__ == "__main__":
    main()
