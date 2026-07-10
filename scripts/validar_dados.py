"""
validar_dados.py

Roda verificacoes automaticas sobre planilha/Base_IES_Cursos_CREA_BA.xlsx e sobre
dados/dados.json, gera dados/validacao.json e imprime um resumo. O
RELATORIO_VALIDACAO.md e escrito manualmente a partir dos numeros deste script
(ver README/INSTRUCOES_ATUALIZACAO).

Uso:
    python scripts/validar_dados.py
"""
import json
import os
import re
from collections import Counter

import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX_PATH = os.path.join(BASE_DIR, "planilha", "Base_IES_Cursos_CREA_BA.xlsx")
JSON_PATH = os.path.join(BASE_DIR, "dados", "dados.json")
OUT_JSON = os.path.join(BASE_DIR, "dados", "validacao.json")

BAHIA_LAT_RANGE = (-19.0, -8.0)
BAHIA_LON_RANGE = (-46.7, -37.0)


def ler_aba(wb, nome):
    ws = wb[nome]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    return [{header[i]: r[i] for i in range(len(header))} for r in rows[1:]]


def main():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ies = ler_aba(wb, "IES")
    campi = ler_aba(wb, "CAMPI")
    cursos = ler_aba(wb, "CURSOS")
    titulos = ler_aba(wb, "TITULOS_PROFISSIONAIS")
    curso_titulo = ler_aba(wb, "CURSO_TITULO")
    base_legal = ler_aba(wb, "BASE_LEGAL")
    titulo_atrib_legal = ler_aba(wb, "TITULO_ATRIBUICAO_LEGAL")

    achados = []  # cada item: {nivel, area, mensagem, registro}

    def achado(nivel, area, mensagem, registro=""):
        achados.append({"nivel": nivel, "area": area, "mensagem": mensagem, "registro": registro})

    # --- IDs duplicados ---
    for nome_tab, linhas, campo_id in [
        ("IES", ies, "id_ies"), ("CAMPI", campi, "id_campus"), ("CURSOS", cursos, "id_curso"),
        ("TITULOS_PROFISSIONAIS", titulos, "codigo_titulo"), ("CURSO_TITULO", curso_titulo, "id_curso_titulo"),
        ("BASE_LEGAL", base_legal, "id_norma"),
    ]:
        contagem = Counter(l[campo_id] for l in linhas)
        dups = [k for k, v in contagem.items() if v > 1]
        if dups:
            achado("erro", nome_tab, f"{len(dups)} valor(es) de {campo_id} duplicados", ", ".join(map(str, dups[:10])))
        else:
            achado("ok", nome_tab, f"Nenhum {campo_id} duplicado ({len(linhas)} registros)")

    # --- CNPJ invalido/ficticio ---
    invalidos = [i for i in ies if i["cnpj_valido"] in ("NÃO", "FICTÍCIO")]
    achado("aviso" if invalidos else "ok", "IES", f"{len(invalidos)} instituição(ões) com CNPJ inválido ou fictício de {len(ies)}")

    # --- Codigo MEC ausente ---
    sem_mec = [i for i in ies if i["codigo_mec"] == "PENDENTE DE VALIDAÇÃO"]
    achado("aviso" if sem_mec else "ok", "IES", f"{len(sem_mec)} instituição(ões) sem Código MEC de {len(ies)}")

    # --- Campus sem municipio ---
    sem_municipio = [c for c in campi if not c["municipio"] and c["situacao_campus"] != "TRANSFERIDO"]
    achado("aviso" if sem_municipio else "ok", "CAMPI", f"{len(sem_municipio)} campus(i) sem município identificado de {len(campi)}")

    # --- Coordenadas fora da faixa da Bahia ---
    fora_faixa = []
    for c in campi:
        lat, lon = c["latitude"], c["longitude"]
        if lat in ("", None) or lon in ("", None):
            continue
        lat, lon = float(lat), float(lon)
        if not (BAHIA_LAT_RANGE[0] <= lat <= BAHIA_LAT_RANGE[1] and BAHIA_LON_RANGE[0] <= lon <= BAHIA_LON_RANGE[1]):
            fora_faixa.append(c["id_campus"])
    achado("erro" if fora_faixa else "ok", "CAMPI", f"{len(fora_faixa)} coordenada(s) fora da faixa geográfica da Bahia", ", ".join(fora_faixa[:10]))

    # --- Curso sem instituicao/campus ---
    campus_validos = {c["id_campus"] for c in campi}
    curso_sem_campus = [c for c in cursos if c["id_campus"] not in campus_validos]
    achado("erro" if curso_sem_campus else "ok", "CURSOS", f"{len(curso_sem_campus)} curso(s) sem campus válido de {len(cursos)}")

    # --- Codigo de titulo inexistente referenciado em CURSO_TITULO ---
    codigos_validos = {t["codigo_titulo"] for t in titulos}
    ct_invalidos = [c for c in curso_titulo if c["codigo_titulo"] not in codigos_validos]
    achado("erro" if ct_invalidos else "ok", "CURSO_TITULO", f"{len(ct_invalidos)} vínculo(s) referenciando código de título inexistente")

    # --- Titulo duplicado (mesmo codigo mais de uma vez na aba principal) ---
    contagem_titulo = Counter(t["codigo_titulo"] for t in titulos)
    tit_dup = [k for k, v in contagem_titulo.items() if v > 1]
    achado("erro" if tit_dup else "ok", "TITULOS_PROFISSIONAIS", f"{len(tit_dup)} código(s) de título duplicado(s) na aba principal (deveriam estar só em TITULOS_EQUIVALENCIAS)")

    # --- Base legal sem link ---
    sem_link = [n for n in base_legal if not n["url_fonte_oficial"]]
    achado("erro" if sem_link else "ok", "BASE_LEGAL", f"{len(sem_link)} norma(s) sem URL de fonte oficial de {len(base_legal)}")

    # --- Artigo sem norma (vinculo legal sem id_norma) ---
    sem_norma = [v for v in titulo_atrib_legal if not v["id_norma"]]
    achado("erro" if sem_norma else "ok", "TITULO_ATRIBUICAO_LEGAL", f"{len(sem_norma)} vínculo(s) de atribuição legal sem norma associada")

    # --- Registros marcados para publicacao mas ainda pendentes de validacao legal ---
    if os.path.exists(JSON_PATH):
        with open(JSON_PATH, encoding="utf-8") as f:
            dados = json.load(f)
        registros = dados["registros"]
        pub_pendente_legal = [
            r for r in registros
            if r.get("publicar_dashboard") == "SIM" and r.get("status_validacao_legal") == "PENDENTE DE VALIDAÇÃO" and r.get("id_curso")
        ]
        achado("aviso" if pub_pendente_legal else "ok", "BASE_DASHBOARD",
               f"{len(pub_pendente_legal)} registro(s) publicáveis com curso vinculado mas sem base legal validada")

        # --- divergencia planilha x json (contagem) ---
        if len(registros) != len(ler_aba(wb, "BASE_DASHBOARD")):
            achado("erro", "BASE_DASHBOARD", "Divergência de contagem entre planilha (BASE_DASHBOARD) e dados.json")
        else:
            achado("ok", "BASE_DASHBOARD", f"Planilha e dados.json com {len(registros)} registros — contagens batem")
    else:
        achado("erro", "dados.json", "Arquivo dados/dados.json não encontrado — rode scripts/gerar_json.py antes de validar")

    resumo = Counter(a["nivel"] for a in achados)
    resultado = {
        "resumo": dict(resumo),
        "achados": achados,
    }
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print("Validação concluída: %d ok | %d avisos | %d erros" % (
        resumo.get("ok", 0), resumo.get("aviso", 0), resumo.get("erro", 0)))
    for a in achados:
        if a["nivel"] != "ok":
            print(f"  [{a['nivel'].upper()}] {a['area']}: {a['mensagem']}")
    print("\nRelatório salvo em: %s" % OUT_JSON)


if __name__ == "__main__":
    main()
