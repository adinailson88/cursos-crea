"""
gerar_json.py

Le planilha/Base_IES_Cursos_CREA_BA.xlsx (abas BASE_DASHBOARD e PENDENCIAS) e
gera dados/dados.json para consumo do dashboard estatico em dashboard/.

Uso:
    python scripts/gerar_json.py
"""
import json
import os
from datetime import datetime

import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX_PATH = os.path.join(BASE_DIR, "planilha", "Base_IES_Cursos_CREA_BA.xlsx")
OUT_JSON = os.path.join(BASE_DIR, "dados", "dados.json")


def ler_aba(ws):
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    out = []
    for r in rows[1:]:
        d = {header[i]: ("" if r[i] is None else r[i]) for i in range(len(header))}
        out.append(d)
    return out


def to_float_or_none(v):
    try:
        if v in ("", None):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def main():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    registros = ler_aba(wb["BASE_DASHBOARD"])
    pendencias = ler_aba(wb["PENDENCIAS"])
    catalogo_titulos = ler_aba(wb["TITULOS_PROFISSIONAIS"])

    for r in registros:
        r["latitude"] = to_float_or_none(r.get("latitude"))
        r["longitude"] = to_float_or_none(r.get("longitude"))

    ies_ids = {r["id_ies"] for r in registros if r["id_ies"]}
    campus_ids = {r["id_campus"] for r in registros if r["id_campus"]}
    municipios = {r["municipio"] for r in registros if r["municipio"]}
    curso_ids = {r["id_curso"] for r in registros if r["id_curso"]}
    titulo_codigos = {r["codigo_titulo"] for r in registros if r["codigo_titulo"]}
    cursos_com_base_legal = {
        r["id_curso"] for r in registros
        if r["id_curso"] and r.get("norma") and r.get("status_validacao_legal") != "PENDENTE DE VALIDAÇÃO"
    }
    publicaveis = [r for r in registros if r.get("publicar_dashboard") == "SIM"]
    campos_completude = ["municipio", "latitude", "codigo_titulo", "norma"]
    total_campos = len(publicaveis) * len(campos_completude) or 1
    preenchidos = sum(
        1 for r in publicaveis for c in campos_completude
        if r.get(c) not in (None, "", "PENDENTE DE VALIDAÇÃO")
    )
    completude = round(100 * preenchidos / total_campos, 1)

    meta = {
        "gerado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "fonte_institucional_principal": "Relatório Instituições (SITAC/CREA-BA)",
        "total_instituicoes": len(ies_ids),
        "total_campi": len(campus_ids),
        "total_municipios": len(municipios),
        "total_cursos": len(curso_ids),
        "total_titulos": len(catalogo_titulos),
        "total_titulos_vinculados_a_cursos": len(titulo_codigos),
        "total_cursos_com_base_legal_validada": len(cursos_com_base_legal),
        "total_registros_dashboard": len(registros),
        "total_registros_publicaveis": len(publicaveis),
        "total_pendencias": len(pendencias),
        "percentual_completude_cadastral": completude,
        "aviso": "Consulta pública de instituições de ensino, cursos, títulos profissionais e referências normativas cadastradas na base analisada do CREA-BA. Este painel não substitui certidão, decisão de Câmara Especializada ou análise formal do CREA-BA.",
    }

    payload = {"meta": meta, "registros": registros, "pendencias": pendencias}

    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=None, separators=(",", ":"))

    print("dados.json gravado em %s" % OUT_JSON)
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
