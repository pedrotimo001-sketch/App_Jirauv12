import json
import os
import shutil
import sqlite3
import tempfile
import uuid
import re
from decimal import Decimal, InvalidOperation
from copy import deepcopy
from pathlib import Path

import pandas as pd

DEFAULT_DB = {"materiais": [], "indices": [], "obras": []}
CAMPOS_INDICE = [
    ("vigas_h", "ESCORAMENTO DE VIGAS — HORIZONTAL"),
    ("vigas_v", "ESCORAMENTO DE VIGAS — VERTICAL"),
    ("lajes_h", "ESCORAMENTO DE LAJES — HORIZONTAL"),
    ("lajes_v", "ESCORAMENTO DE LAJES — VERTICAL"),
    ("reesc", "REESCORAMENTO DE VIGAS E LAJES"),
    ("trav_pil", "TRAVAMENTO DE PILARES"),
    ("trav_vig", "TRAVAMENTO DE VIGAS"),
]
CAMPOS_AREA = {"vigas_h", "lajes_h", "trav_vig"}
CAMPOS_VOLUME = {"vigas_v", "lajes_v", "reesc", "trav_pil"}

SERVICOS_PROPOSTA = [
    {"chave": "escoramento", "descricao": "ESCORAMENTO DE VIGAS E LAJES", "campos": ["vigas_h", "vigas_v", "lajes_h", "lajes_v"]},
    {"chave": "reescoramento", "descricao": "REESCORAMENTO DE VIGAS E LAJES", "campos": ["reesc"]},
    {"chave": "trav_pilar", "descricao": "TRAVAMENTO DE PILAR", "campos": ["trav_pil"]},
    {"chave": "trav_viga", "descricao": "TRAVAMENTO DE VIGA", "campos": ["trav_vig"]},
]

def configuracao_servicos_padrao():
    return {s["chave"]: {"percentual": 100.0, "fator": 1.0, "quantidade_jogos": 1.0} for s in SERVICOS_PROPOSTA}

def normalizar_configuracao_servicos(config):
    base = configuracao_servicos_padrao()
    if not isinstance(config, dict):
        return base
    for chave, padrao in base.items():
        atual = config.get(chave, {}) if isinstance(config.get(chave), dict) else {}
        base[chave] = {
            "percentual": max(0.0, _float(atual.get("percentual"), padrao["percentual"])),
            # Na planilha oficial, FATOR é um indicador calculado e não um multiplicador comercial.
            # Mantemos o campo apenas por compatibilidade com bancos antigos, sempre neutralizado em 1,0.
            "fator": 1.0,
            "quantidade_jogos": max(0.0, _float(atual.get("quantidade_jogos"), padrao["quantidade_jogos"])),
        }
    return base


def uid():
    return str(uuid.uuid4())


def _float(valor, padrao=0.0):
    """Converte números, inclusive no padrão brasileiro (1.234,56 e R$ 1.234,56)."""
    if valor is None:
        return padrao
    if isinstance(valor, bool):
        return float(valor)
    if isinstance(valor, (int, float, Decimal)):
        try:
            n = float(valor)
            return n if n == n and n not in (float("inf"), float("-inf")) else padrao
        except (TypeError, ValueError, OverflowError):
            return padrao
    texto = str(valor).strip()
    if not texto:
        return padrao
    texto = re.sub(r"[^0-9, .+\-]", "", texto).replace(" ", "")
    if not texto:
        return padrao
    # Quando há ponto e vírgula, o último separador é tratado como decimal.
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    else:
        # Pontos múltiplos sem vírgula são separadores de milhar, salvo o último.
        if texto.count(".") > 1:
            partes = texto.split(".")
            texto = "".join(partes[:-1]) + "." + partes[-1]
    try:
        n = float(Decimal(texto))
        return n if n == n and n not in (float("inf"), float("-inf")) else padrao
    except (InvalidOperation, ValueError, OverflowError):
        return padrao


def numero_float(valor, padrao=0.0):
    """Conversor público usado pela interface e pelos importadores."""
    return _float(valor, padrao)


def _int(valor, padrao=0):
    try:
        return int(valor)
    except (TypeError, ValueError):
        return padrao


def moeda(valor):
    return f"R$ {_float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def numero_br(valor):
    return f"{_float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def consumo_total_item(item):
    valores = [_float(item.get(chave, 0)) for chave, _ in CAMPOS_INDICE]
    if not any(valores) and "consumo" in item:
        return _float(item.get("consumo", 0))
    return sum(valores)


def normalizar_db(db):
    if not isinstance(db, dict):
        db = deepcopy(DEFAULT_DB)
    for chave in DEFAULT_DB:
        if not isinstance(db.get(chave), list):
            db[chave] = []

    codigos = set()
    materiais = []
    for material in db["materiais"]:
        if not isinstance(material, dict):
            continue
        codigo = str(material.get("codigo", "")).strip()
        descricao = str(material.get("descricao", "")).strip()
        if not codigo or not descricao or codigo.casefold() in codigos:
            continue
        codigos.add(codigo.casefold())
        materiais.append({
            "id": str(material.get("id") or uid()), "codigo": codigo,
            "descricao": descricao, "peso": max(0.0, _float(material.get("peso"))),
            "valor": max(0.0, _float(material.get("valor"))),
        })
    db["materiais"] = materiais
    materiais_ids = {m["id"] for m in materiais}

    indices = []
    indices_ids = set()
    for indice in db["indices"]:
        if not isinstance(indice, dict):
            continue
        iid = str(indice.get("id") or uid())
        if iid in indices_ids:
            iid = uid()
        indices_ids.add(iid)
        itens = []
        vistos = set()
        for item in indice.get("itens", []) if isinstance(indice.get("itens", []), list) else []:
            if not isinstance(item, dict) or item.get("material_id") not in materiais_ids:
                continue
            mid = item["material_id"]
            if mid in vistos:
                continue
            vistos.add(mid)
            if not any(chave in item for chave, _ in CAMPOS_INDICE):
                item = {**item, "lajes_h": _float(item.get("consumo"))}
            valores = {chave: max(0.0, _float(item.get(chave))) for chave, _ in CAMPOS_INDICE}
            if any(valores.values()):
                itens.append({"material_id": mid, **valores})
        indices.append({
            "id": iid, "nome": str(indice.get("nome", "Índice sem nome")).strip() or "Índice sem nome",
            "descricao": str(indice.get("descricao", "")).strip(),
            "area_base": max(0.01, _float(indice.get("area_base"), 1.0)),
            "pe_direito_base": max(0.01, _float(indice.get("pe_direito_base"), 1.0)),
            "itens": itens,
        })
    db["indices"] = indices
    indices_ids = {i["id"] for i in indices}

    obras = []
    for obra in db["obras"]:
        if not isinstance(obra, dict):
            continue
        tetos = []
        for teto in obra.get("tetos", []) if isinstance(obra.get("tetos", []), list) else []:
            if not isinstance(teto, dict) or teto.get("indice_id") not in indices_ids:
                continue
            area = max(0.0, _float(teto.get("area")))
            pe = max(0.0, _float(teto.get("pe_direito")))
            if area <= 0 or pe <= 0:
                continue
            tetos.append({
                "id": str(teto.get("id") or uid()), "nome": str(teto.get("nome", "Teto")).strip() or "Teto",
                "area": area, "pe_direito": pe, "dias": max(1, _int(teto.get("dias"), 30)),
                "indice_id": teto["indice_id"],
                "servicos_config": normalizar_configuracao_servicos(teto.get("servicos_config")),
            })
        obras.append({
            "id": str(obra.get("id") or uid()), "nome": str(obra.get("nome", "Obra sem nome")).strip() or "Obra sem nome",
            "cliente": str(obra.get("cliente", "")).strip(), "referencia": str(obra.get("referencia", "")).strip(),
            "endereco": str(obra.get("endereco", "")).strip(), "tetos": tetos,
        })
    db["obras"] = obras
    recalcular_todos(db, persistir=False)
    return db


def _sqlite_conectar(db_file):
    db_file = Path(db_file)
    db_file.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(db_file))
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    con.execute("PRAGMA journal_mode = WAL")
    con.execute("PRAGMA synchronous = FULL")
    con.executescript("""
    CREATE TABLE IF NOT EXISTS materiais (
        id TEXT PRIMARY KEY, codigo TEXT NOT NULL UNIQUE COLLATE NOCASE,
        descricao TEXT NOT NULL, peso REAL NOT NULL DEFAULT 0, valor REAL NOT NULL DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS indices (
        id TEXT PRIMARY KEY, nome TEXT NOT NULL, descricao TEXT NOT NULL DEFAULT '',
        area_base REAL NOT NULL, pe_direito_base REAL NOT NULL
    );
    CREATE TABLE IF NOT EXISTS indice_itens (
        indice_id TEXT NOT NULL REFERENCES indices(id) ON DELETE CASCADE,
        material_id TEXT NOT NULL REFERENCES materiais(id) ON DELETE RESTRICT,
        vigas_h REAL NOT NULL DEFAULT 0, vigas_v REAL NOT NULL DEFAULT 0,
        lajes_h REAL NOT NULL DEFAULT 0, lajes_v REAL NOT NULL DEFAULT 0,
        reesc REAL NOT NULL DEFAULT 0, trav_pil REAL NOT NULL DEFAULT 0, trav_vig REAL NOT NULL DEFAULT 0,
        PRIMARY KEY (indice_id, material_id)
    );
    CREATE TABLE IF NOT EXISTS obras (
        id TEXT PRIMARY KEY, nome TEXT NOT NULL, cliente TEXT NOT NULL DEFAULT '',
        referencia TEXT NOT NULL DEFAULT '', endereco TEXT NOT NULL DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS tetos (
        id TEXT PRIMARY KEY, obra_id TEXT NOT NULL REFERENCES obras(id) ON DELETE CASCADE,
        nome TEXT NOT NULL, area REAL NOT NULL, pe_direito REAL NOT NULL,
        dias INTEGER NOT NULL, indice_id TEXT NOT NULL REFERENCES indices(id) ON DELETE RESTRICT
    );
    CREATE TABLE IF NOT EXISTS servicos_config (
        teto_id TEXT NOT NULL REFERENCES tetos(id) ON DELETE CASCADE,
        chave TEXT NOT NULL, percentual REAL NOT NULL, fator REAL NOT NULL, quantidade_jogos REAL NOT NULL,
        PRIMARY KEY (teto_id, chave)
    );
    CREATE TABLE IF NOT EXISTS metadados (chave TEXT PRIMARY KEY, valor TEXT NOT NULL);
    """)
    return con


def _carregar_sqlite(db_file):
    with _sqlite_conectar(db_file) as con:
        materiais = [dict(r) for r in con.execute("SELECT id,codigo,descricao,peso,valor FROM materiais ORDER BY codigo")]
        indices = []
        for r in con.execute("SELECT id,nome,descricao,area_base,pe_direito_base FROM indices ORDER BY nome"):
            i = dict(r)
            i["itens"] = [dict(x) for x in con.execute(
                "SELECT material_id,vigas_h,vigas_v,lajes_h,lajes_v,reesc,trav_pil,trav_vig FROM indice_itens WHERE indice_id=?",
                (i["id"],),
            )]
            indices.append(i)
        obras = []
        for r in con.execute("SELECT id,nome,cliente,referencia,endereco FROM obras ORDER BY nome"):
            o = dict(r); o["tetos"] = []
            for trow in con.execute("SELECT id,nome,area,pe_direito,dias,indice_id FROM tetos WHERE obra_id=? ORDER BY rowid", (o["id"],)):
                t = dict(trow)
                cfg = {}
                for c in con.execute("SELECT chave,percentual,fator,quantidade_jogos FROM servicos_config WHERE teto_id=?", (t["id"],)):
                    cfg[c["chave"]] = {"percentual": c["percentual"], "fator": c["fator"], "quantidade_jogos": c["quantidade_jogos"]}
                t["servicos_config"] = cfg
                o["tetos"].append(t)
            obras.append(o)
    return normalizar_db({"materiais": materiais, "indices": indices, "obras": obras})


def _salvar_sqlite(db, db_file):
    db = normalizar_db(deepcopy(db))
    db_file = Path(db_file)
    backup = db_file.with_suffix(db_file.suffix + ".bak")
    if db_file.exists():
        try: shutil.copy2(db_file, backup)
        except OSError: pass
    with _sqlite_conectar(db_file) as con:
        con.execute("BEGIN IMMEDIATE")
        con.execute("DELETE FROM servicos_config")
        con.execute("DELETE FROM tetos")
        con.execute("DELETE FROM obras")
        con.execute("DELETE FROM indice_itens")
        con.execute("DELETE FROM indices")
        con.execute("DELETE FROM materiais")
        con.executemany("INSERT INTO materiais(id,codigo,descricao,peso,valor) VALUES (?,?,?,?,?)", [
            (m["id"],m["codigo"],m["descricao"],m["peso"],m["valor"]) for m in db["materiais"]])
        con.executemany("INSERT INTO indices(id,nome,descricao,area_base,pe_direito_base) VALUES (?,?,?,?,?)", [
            (i["id"],i["nome"],i["descricao"],i["area_base"],i["pe_direito_base"]) for i in db["indices"]])
        for i in db["indices"]:
            con.executemany("INSERT INTO indice_itens(indice_id,material_id,vigas_h,vigas_v,lajes_h,lajes_v,reesc,trav_pil,trav_vig) VALUES (?,?,?,?,?,?,?,?,?)", [
                (i["id"],x["material_id"],x["vigas_h"],x["vigas_v"],x["lajes_h"],x["lajes_v"],x["reesc"],x["trav_pil"],x["trav_vig"]) for x in i["itens"]])
        con.executemany("INSERT INTO obras(id,nome,cliente,referencia,endereco) VALUES (?,?,?,?,?)", [
            (o["id"],o["nome"],o["cliente"],o["referencia"],o["endereco"]) for o in db["obras"]])
        for o in db["obras"]:
            for t in o["tetos"]:
                con.execute("INSERT INTO tetos(id,obra_id,nome,area,pe_direito,dias,indice_id) VALUES (?,?,?,?,?,?,?)",
                            (t["id"],o["id"],t["nome"],t["area"],t["pe_direito"],t["dias"],t["indice_id"]))
                cfg = normalizar_configuracao_servicos(t.get("servicos_config"))
                con.executemany("INSERT INTO servicos_config(teto_id,chave,percentual,fator,quantidade_jogos) VALUES (?,?,?,?,?)", [
                    (t["id"],chave,v["percentual"],v["fator"],v["quantidade_jogos"]) for chave,v in cfg.items()])
        con.execute("INSERT OR REPLACE INTO metadados(chave,valor) VALUES ('schema_version','11')")
    return db


def migrar_json_para_sqlite(json_file, sqlite_file):
    json_file, sqlite_file = Path(json_file), Path(sqlite_file)
    if sqlite_file.exists() or not json_file.exists():
        return False
    db = carregar_db(json_file)
    _salvar_sqlite(db, sqlite_file)
    return True


def carregar_db(db_file):
    db_file = Path(db_file)
    if db_file.suffix.lower() in {".db", ".sqlite", ".sqlite3"}:
        if not db_file.exists():
            _salvar_sqlite(deepcopy(DEFAULT_DB), db_file)
        try:
            return _carregar_sqlite(db_file)
        except sqlite3.DatabaseError:
            backup = db_file.with_suffix(db_file.suffix + ".bak")
            if backup.exists():
                shutil.copy2(backup, db_file)
                return _carregar_sqlite(db_file)
            raise
    if not db_file.exists():
        salvar_db(deepcopy(DEFAULT_DB), db_file)
    try:
        db = json.loads(db_file.read_text(encoding="utf-8"))
    except Exception:
        backup = db_file.with_suffix(db_file.suffix + ".bak")
        try: db = json.loads(backup.read_text(encoding="utf-8"))
        except Exception: db = deepcopy(DEFAULT_DB)
    db = normalizar_db(db)
    salvar_db(db, db_file)
    return db


def salvar_db(db, db_file):
    db_file = Path(db_file)
    if db_file.suffix.lower() in {".db", ".sqlite", ".sqlite3"}:
        return _salvar_sqlite(db, db_file)
    db_file.parent.mkdir(parents=True, exist_ok=True)
    conteudo = json.dumps(db, ensure_ascii=False, indent=2, allow_nan=False)
    fd, temporario = tempfile.mkstemp(prefix=db_file.name + ".", suffix=".tmp", dir=db_file.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as arq:
            arq.write(conteudo); arq.flush(); os.fsync(arq.fileno())
        if db_file.exists(): shutil.copy2(db_file, db_file.with_suffix(db_file.suffix + ".bak"))
        os.replace(temporario, db_file)
    finally:
        if os.path.exists(temporario): os.unlink(temporario)

def encontrar_material(db, material_id):
    return next((m for m in db["materiais"] if m.get("id") == material_id), None)


def encontrar_indice(db, indice_id):
    return next((i for i in db["indices"] if i.get("id") == indice_id), None)


def recalcular_teto(db, teto):
    """Recalcula um teto a partir dos índices unitários do projeto-base.

    Conforme a planilha oficial JIRAU/TIPO:
    - Vigas H, Lajes H e Travamento de Vigas: baseados em área (m²).
    - Vigas V, Lajes V, Reescoramento e Travamento de Pilares: baseados em volume (m³).
    """
    indice = encontrar_indice(db, teto.get("indice_id"))
    if not indice:
        teto.update({"indice_nome": "Índice não encontrado", "materiais": [], "servicos": [], "valor_mensal": 0.0, "valor_total": 0.0})
        return teto

    area = max(0.0, _float(teto.get("area")))
    pe = max(0.0, _float(teto.get("pe_direito")))
    volume = area * pe
    dias = max(1, _int(teto.get("dias"), 30))
    area_base = max(0.01, _float(indice.get("area_base"), 1.0))
    pe_base = max(0.01, _float(indice.get("pe_direito_base"), 1.0))
    volume_base = area_base * pe_base
    config = normalizar_configuracao_servicos(teto.get("servicos_config"))

    acumulado = {}
    for item in indice.get("itens", []):
        mat = encontrar_material(db, item.get("material_id"))
        if not mat:
            continue
        detalhe = acumulado.setdefault(mat["id"], {
            "codigo": mat["codigo"], "descricao": mat["descricao"],
            "peso_unitario": _float(mat["peso"]), "valor_unitario": _float(mat["valor"]),
            **{chave: 0.0 for chave, _ in CAMPOS_INDICE},
            "indices_unitarios": {},
        })
        for chave, _ in CAMPOS_INDICE:
            qtd_base = max(0.0, _float(item.get(chave)))
            divisor = area_base if chave in CAMPOS_AREA else volume_base
            grandeza = area if chave in CAMPOS_AREA else volume
            indice_unitario = qtd_base / divisor if divisor else 0.0
            detalhe["indices_unitarios"][chave] = detalhe["indices_unitarios"].get(chave, 0.0) + indice_unitario
            detalhe[chave] += indice_unitario * grandeza

    servicos = []
    valor_mensal = 0.0
    valor_periodo = 0.0
    for servico in SERVICOS_PROPOSTA:
        cfg = config[servico["chave"]]
        percentual = cfg["percentual"] / 100.0
        fator_preco = 1.0
        jogos = cfg["quantidade_jogos"]
        horizontal_base = vertical_base = peso_base = 0.0
        for detalhe in acumulado.values():
            for campo in servico["campos"]:
                qtd = detalhe.get(campo, 0.0)
                valor = qtd * detalhe["valor_unitario"]
                peso = qtd * detalhe["peso_unitario"]
                if campo in CAMPOS_AREA:
                    horizontal_base += valor
                else:
                    vertical_base += valor
                peso_base += peso
        multiplicador_quantidade = percentual * jogos
        horizontal = horizontal_base * multiplicador_quantidade
        vertical = vertical_base * multiplicador_quantidade
        mensal = horizontal + vertical
        periodo = mensal * dias / 30.0
        servicos.append({
            "chave": servico["chave"], "descricao": servico["descricao"],
            "quantidade_jogos": jogos, "percentual": cfg["percentual"], "fator": fator_preco,
            "horizontal_mensal": horizontal, "vertical_mensal": vertical,
            "horizontal_rs_m2": horizontal / area if area else 0.0,
            "vertical_rs_m3": vertical / volume if volume else 0.0,
            "unitario_dia": mensal / 30.0, "total_mensal": mensal, "total_periodo": periodo,
            "peso_total": peso_base * multiplicador_quantidade,
        })
        valor_mensal += mensal
        valor_periodo += periodo

    materiais = []
    peso_total_geral = 0.0
    campo_para_servico = {campo: sv["chave"] for sv in SERVICOS_PROPOSTA for campo in sv["campos"]}
    for detalhe in acumulado.values():
        qtd_total = valor_total_item = peso_total_item = 0.0
        saida = {k: v for k, v in detalhe.items()}
        for chave, _ in CAMPOS_INDICE:
            cfg = config[campo_para_servico[chave]]
            qtd_ajustada = detalhe[chave] * (cfg["percentual"] / 100.0) * cfg["quantidade_jogos"]
            saida[chave] = qtd_ajustada
            qtd_total += qtd_ajustada
            peso_total_item += qtd_ajustada * detalhe["peso_unitario"]
            valor_total_item += qtd_ajustada * detalhe["valor_unitario"]
        saida.update({"quantidade": qtd_total, "peso_total": peso_total_item, "valor_total": valor_total_item})
        materiais.append(saida)
        peso_total_geral += peso_total_item

    teto.update({
        "area": area, "pe_direito": pe, "dias": dias, "volume": volume,
        "area_base_indice": area_base, "pe_direito_base_indice": pe_base, "volume_base_indice": volume_base,
        "indice_nome": indice["nome"], "servicos_config": config, "servicos": servicos,
        "custo_m2": valor_mensal / area if area else 0.0,
        "peso_m2": peso_total_geral / area if area else 0.0,
        "valor_mensal": valor_mensal, "valor_total": valor_periodo, "materiais": materiais,
    })
    return teto


def recalcular_todos(db, db_file=None, persistir=True):
    for obra in db.get("obras", []):
        for teto in obra.get("tetos", []):
            recalcular_teto(db, teto)
    if persistir and db_file is not None:
        salvar_db(db, db_file)
    return db


def substituir_materiais_preservando_vinculos(db, novos):
    antigos_por_codigo = {m["codigo"].casefold(): m for m in db["materiais"]}
    ids_validos = set()
    normalizados = []
    vistos = set()
    for novo in novos:
        codigo = str(novo.get("codigo", "")).strip()
        descricao = str(novo.get("descricao", "")).strip()
        if not codigo or not descricao or codigo.casefold() in vistos:
            continue
        vistos.add(codigo.casefold())
        antigo = antigos_por_codigo.get(codigo.casefold())
        mid = antigo["id"] if antigo else str(novo.get("id") or uid())
        ids_validos.add(mid)
        normalizados.append({"id": mid, "codigo": codigo, "descricao": descricao, "peso": max(0.0, _float(novo.get("peso"))), "valor": max(0.0, _float(novo.get("valor")))})
    db["materiais"] = normalizados
    for indice in db["indices"]:
        indice["itens"] = [item for item in indice.get("itens", []) if item.get("material_id") in ids_validos]
    recalcular_todos(db, persistir=False)


def normalizar_coluna(texto):
    tabela = str.maketrans("ÇÃÁÀÂÉÊÍÓÔÕÚÜ", "CAAAAEEIOOOUU")
    return str(texto).strip().upper().translate(tabela)


def carregar_indice_tipo_oficial(db, arquivo_template=None, substituir_existente=True):
    """Carrega materiais e o índice TIPO oficial da planilha Jirau 2026.

    O template contém quantidades brutas do pavimento-base de 403,23 m² e P.D. 3,30 m.
    Materiais existentes são preservados por código e têm peso/locação atualizados.
    """
    caminho = Path(arquivo_template) if arquivo_template else Path(__file__).with_name("indice_tipo_oficial_2026.json")
    with open(caminho, "r", encoding="utf-8") as arq:
        template = json.load(arq)

    por_codigo = {str(m.get("codigo", "")).strip().casefold(): m for m in db.get("materiais", [])}
    itens_indice = []
    for item in template.get("itens", []):
        codigo = str(item.get("codigo", "")).strip()
        chave_codigo = codigo.casefold()
        mat = por_codigo.get(chave_codigo)
        if mat is None:
            mat = {
                "id": uid(), "codigo": codigo,
                "descricao": str(item.get("descricao", "")).strip(),
                "peso": max(0.0, _float(item.get("peso"))),
                "valor": max(0.0, _float(item.get("valor"))),
            }
            db.setdefault("materiais", []).append(mat)
            por_codigo[chave_codigo] = mat
        else:
            mat["descricao"] = str(item.get("descricao", mat.get("descricao", ""))).strip() or mat.get("descricao", "")
            mat["peso"] = max(0.0, _float(item.get("peso"), mat.get("peso", 0)))
            mat["valor"] = max(0.0, _float(item.get("valor"), mat.get("valor", 0)))
        itens_indice.append({"material_id": mat["id"], **{chave: max(0.0, _float(item.get(chave))) for chave, _ in CAMPOS_INDICE}})

    nome = str(template.get("nome", "TIPO OFICIAL JIRAU 2026"))
    existente = next((i for i in db.get("indices", []) if i.get("nome", "").strip().casefold() == nome.casefold()), None)
    dados = {
        "nome": nome,
        "descricao": str(template.get("descricao", "")),
        "area_base": max(0.01, _float(template.get("area_base"), 403.23)),
        "pe_direito_base": max(0.01, _float(template.get("pe_direito_base"), 3.30)),
        "itens": itens_indice,
    }
    if existente and substituir_existente:
        existente.update(dados)
        indice = existente
    else:
        indice = {"id": uid(), **dados}
        db.setdefault("indices", []).append(indice)
    recalcular_todos(db, persistir=False)
    return indice


def auditar_indice(indice):
    """Retorna alertas para bases que provavelmente produzirão escala duplicada."""
    alertas = []
    area = _float(indice.get("area_base"), 0)
    pe = _float(indice.get("pe_direito_base"), 0)
    total = sum(consumo_total_item(i) for i in indice.get("itens", []))
    if total > 10 and area <= 1.01:
        alertas.append("Área-base igual a 1,00 m² com quantidades brutas elevadas: provável escala duplicada.")
    if total > 10 and pe <= 1.01:
        alertas.append("Pé-direito-base igual a 1,00 m com quantidades brutas elevadas: provável escala vertical duplicada.")
    return alertas

def detectar_materiais_excel(arquivo):
    xls = pd.ExcelFile(arquivo)
    abas = ["JIRAU"] if "JIRAU" in xls.sheet_names else xls.sheet_names
    for aba in abas:
        df = pd.read_excel(arquivo, sheet_name=aba, header=None)
        for idx, row in df.iterrows():
            textos = [normalizar_coluna(v) for v in row.tolist()]
            linha = " | ".join(textos)
            if "CODIGO" not in linha or not ("DESCRICAO" in linha or "MATERIAL" in linha):
                continue
            mapa = {}
            for pos, nome in enumerate(textos):
                if "CODIGO" in nome and "codigo" not in mapa: mapa["codigo"] = pos
                elif ("DESCRICAO" in nome or "MATERIAL" in nome) and "descricao" not in mapa: mapa["descricao"] = pos
                elif "PESO" in nome and "peso" not in mapa: mapa["peso"] = pos
                elif ("VALOR" in nome or "LOCACAO" in nome or "PRECO" in nome) and "valor" not in mapa: mapa["valor"] = pos
            if not {"codigo", "descricao"}.issubset(mapa):
                continue
            novos, vistos = [], set()
            for _, dados in df.iloc[idx + 1:].iterrows():
                codigo = dados.iloc[mapa["codigo"]] if mapa["codigo"] < len(dados) else None
                descricao = dados.iloc[mapa["descricao"]] if mapa["descricao"] < len(dados) else None
                if pd.isna(codigo) or pd.isna(descricao): continue
                codigo = str(codigo).strip()
                if not codigo or codigo.casefold() in vistos: continue
                vistos.add(codigo.casefold())
                peso = _float(dados.iloc[mapa["peso"]]) if "peso" in mapa and mapa["peso"] < len(dados) and not pd.isna(dados.iloc[mapa["peso"]]) else 0.0
                valor = _float(dados.iloc[mapa["valor"]]) if "valor" in mapa and mapa["valor"] < len(dados) and not pd.isna(dados.iloc[mapa["valor"]]) else 0.0
                novos.append({"id": uid(), "codigo": codigo, "descricao": str(descricao).strip(), "peso": max(0, peso), "valor": max(0, valor)})
            if novos: return novos, aba
    raise ValueError("Não foi possível localizar uma tabela com Código e Descrição no arquivo.")
