import os
import re
import time
import requests
import pandas as pd

# ===== Config =====
BASE = "https://abyara.sigavi360.com.br"  # host confirmado
USER = os.getenv("SIGAVI_USER", "integracao")
PASS = os.getenv("SIGAVI_PASS", "")
ARQUIVO_EXCEL = "./ReservaSaoCaetano_0209a08092025_aby[1].xlsx"
SHEET_NAME = None  # deixe None p/ usar a 1ª aba

# IMPORTANTE: reutilize o MESMO dicionário que você já tem
from seu_modulo_ou_arquivo import corretores_gerentes  # ou copie seu dict aqui

# ===== Token cache simples =====
_TOKEN = {"val": None, "exp": 0}

def get_token():
    """Obtém e cacheia o token do Sigavi (expira em ~12h / 30min dependendo do tenant)."""
    global _TOKEN
    now = time.time()
    if _TOKEN["val"] and now < _TOKEN["exp"] - 60:
        return _TOKEN["val"]

    resp = requests.post(
        f"{BASE}/Sigavi/api/Acesso/Token",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={
            "username": USER,
            "password": PASS,
            "grant_type": "password",
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    _TOKEN["val"] = data["access_token"]
    _TOKEN["exp"] = now + int(data.get("expires_in", 1800))
    return _TOKEN["val"]

def fmt_phone(s):
    """Formata para (00) 00000-0000 / (00) 0000-0000 quando possível; caso contrário retorna original."""
    digits = re.sub(r"\D", "", str(s or ""))
    if len(digits) == 11:
        return f"({digits[0:2]}) {digits[2:7]}-{digits[7:11]}"
    if len(digits) == 10:
        return f"({digits[0:2]}) {digits[2:6]}-{digits[6:10]}"
    return str(s or "").strip()

def post_nova_lead(payload):
    """Chama a API de criação de lead."""
    token = get_token()
    resp = requests.post(
        f"{BASE}/Sigavi/api/Leads/NovaLead",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )
    # A API normalmente retorna 200 com {Id, Sucesso, Erros: []}
    ok_status = (resp.status_code == 200)
    data = None
    try:
        data = resp.json()
    except Exception:
        pass
    return ok_status, resp.status_code, data, resp.text

def main():
    # Leitura da planilha (mantendo seus renomes)
    xls = pd.ExcelFile(ARQUIVO_EXCEL)
    sheet = SHEET_NAME or xls.sheet_names[0]
    df = pd.read_excel(ARQUIVO_EXCEL, sheet_name=sheet)
    df = df.rename(columns={
        'CORRETOR ORIGEM': 'CORRETOR DE ORIGEM',
        'TELEFONE': 'FONE2',
        'NOME COMPLETO': 'NOME'
    })

    resultados = []
    for _, row in df.iterrows():
        nome = str(row.get('NOME', '')).strip()
        telefone = fmt_phone(row.get('FONE2'))
        corretor = str(row.get('CORRETOR DE ORIGEM', '')).strip()
        gerente = corretores_gerentes.get(corretor) or corretores_gerentes.get(corretor.upper())

        if not nome or not telefone:
            resultados.append({
                "nome": nome, "telefone": telefone,
                "status": "FALHA 422", "erros": "Nome/Telefone ausente(s)"
            })
            continue

        # Monte a mensagem com metadados úteis (já que não mapeamos campos avançados)
        msg_parts = ["Importação automática via API", f"Corretor origem: {corretor or 'N/D'}"]
        if gerente:
            msg_parts.append(f"Gerente: {gerente}")
        mensagem = " | ".join(msg_parts)

        # Payload mínimo (adicione campos extras conforme “Layout dos Campos”)
        payload = {
            "Nome": nome,
            "Telefone": telefone,
            "Mensagem": mensagem,
            "Empreendimento": True,
        }

        # Email se existir na planilha
        email_raw = row.get('EMAIL') or row.get('Email') or row.get('E-MAIL')
        if isinstance(email_raw, str) and email_raw.strip():
            payload["Email"] = email_raw.strip()

        # Referência (coluna, se houver) — fallback para a tag do lote
        ref = row.get('REFERENCIA') or row.get('Referencia') or "Reserva São Caetano"
        if isinstance(ref, str) and ref.strip():
            payload["Referencia"] = ref.strip()

        ok, status, data, raw = post_nova_lead(payload)

        if ok and data and data.get("Sucesso") is True:
            resultados.append({
                "nome": nome, "telefone": telefone, "status": "OK", "lead_id": data.get("Id")
            })
        else:
            erros = None
            if data and isinstance(data.get("Erros"), list):
                erros = "; ".join(str(e) for e in data["Erros"])
            resultados.append({
                "nome": nome, "telefone": telefone,
                "status": f"FALHA {status}", "erros": erros or (raw[:300] if isinstance(raw, str) else "erro desconhecido")
            })

        # pequena pausa para não sobrecarregar
        time.sleep(0.3)

    out = pd.DataFrame(resultados)
    out.to_csv("resultado_importacao.csv", index=False, encoding="utf-8")
    print(out)

if __name__ == "__main__":
    main()
