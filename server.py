from flask import Flask, render_template, request, jsonify, send_file, abort, session, redirect, url_for
from functools import wraps
import subprocess, threading, json, os, sys, uuid, io, zipfile, re, time, socket
from datetime import datetime
import unicodedata
import pandas as pd

app = Flask(__name__)
app.secret_key = "nortao_robo_certidoes_chave_secreta_2026"
app.config["SESSION_PERMANENT"] = False  # sessão expira ao fechar o navegador

SENHA = "nortaocontabilidade2026"

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("autenticado"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

@app.after_request
def sem_cache(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"]        = "no-cache"
    response.headers["Expires"]       = "0"
    return response

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
PASTA_CERT = os.path.join(os.path.expanduser("~"), "Documents", "certidões")
MEI_DIR    = os.path.join(BASE_DIR, "Emissão de Guias MEI")

ARQUIVO_EXCEL = os.path.join(BASE_DIR, "cnpjs.xlsx.xlsx")
COLUNA_CNPJ   = "CNPJ (MF) N.º"
COLUNA_CIDADE = "Cidade / UF"
COLUNA_NOME   = "Razão Social"
_excel_lock   = threading.Lock()

CIDADES_CONFIG_FILE = os.path.join(BASE_DIR, "cidades_config.json")
ESTADOS_BETHA       = {"MT", "MS"}
_cidades_lock       = threading.Lock()

def _norm_str(s):
    """Remove acentos e coloca em minúsculo — igual ao normalizar() do robo.py."""
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower().strip()

def _ler_excel():
    if not os.path.exists(ARQUIVO_EXCEL):
        return pd.DataFrame(columns=[COLUNA_CNPJ, COLUNA_NOME, COLUNA_CIDADE])
    try:
        df = pd.read_excel(ARQUIVO_EXCEL, dtype=str).fillna("")
        # Renomeia colunas para os nomes esperados, ignorando acentos e capitalização
        esperadas = [COLUNA_CNPJ, COLUNA_NOME, COLUNA_CIDADE]
        rename = {}
        for col in df.columns:
            for esp in esperadas:
                if _norm_str(col) == _norm_str(esp) and col != esp:
                    rename[col] = esp
        if rename:
            df = df.rename(columns=rename)
        return df
    except Exception:
        return pd.DataFrame(columns=[COLUNA_CNPJ, COLUNA_NOME, COLUNA_CIDADE])

def _salvar_excel(df):
    df.to_excel(ARQUIVO_EXCEL, index=False)

def _ler_cidades():
    if os.path.exists(CIDADES_CONFIG_FILE):
        try:
            with open(CIDADES_CONFIG_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def _salvar_cidades(cfg):
    with open(CIDADES_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
HIST_FILE  = os.path.join(BASE_DIR, "historico.json")

_lock      = threading.Lock()
fila       = []
job_atual  = None
historico  = []
_logs      = {}
_proc_atual = None

# ── persistência ─────────────────────────────────────────────────────────────

def _carregar():
    if os.path.exists(HIST_FILE):
        try:
            with open(HIST_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def _salvar():
    with open(HIST_FILE, "w", encoding="utf-8") as f:
        json.dump(historico[:50], f, ensure_ascii=False, indent=2)

historico = _carregar()

# ── execução do robô ─────────────────────────────────────────────────────────

def _rodar_job(job):
    global job_atual, _proc_atual
    jid       = job["id"]
    _logs[jid] = []
    inicio_ts  = time.time()

    is_mei = job.get("script") == "mei"

    if is_mei:
        env  = os.environ.copy()
        env["PYTHONUNBUFFERED"]  = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        args = [sys.executable, "main.py"]
        run_dir = MEI_DIR
    else:
        env = os.environ.copy()
        env["ROBO_HEADLESS"]    = "1"
        env["ROBO_WEB"]         = "1"
        env["PYTHONUNBUFFERED"] = "1"
        env["ROBO_JOB_ID"]      = jid
        run_dir = BASE_DIR

        script_name = job.get("script", "robo.py")
        args = [sys.executable, os.path.join(BASE_DIR, script_name)]
        if script_name == "robo_geral.py":
            args.append(job["modo"])
        elif job["modo"] == "cnpj":
            args += ["cnpj", job.get("cnpj", "")]
        else:
            args.append(job["modo"])

    job["inicio"] = datetime.now().strftime("%d/%m %H:%M")
    job["status"] = "rodando"

    cnpj_atual       = ""
    erros_detalhe    = []
    em_relatorio     = False
    relatorio_linhas = []

    try:
        proc = subprocess.Popen(
            args,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL if is_mei else subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            text=True, encoding="utf-8", errors="replace",
            env=env, cwd=run_dir,
        )
        _proc_atual = proc
        _em_traceback = False
        for linha in proc.stdout:
            linha = linha.rstrip()
            if not linha:
                _em_traceback = False
                continue
            stripped = linha.strip()
            # Detecta início de traceback Python
            if stripped == 'Traceback (most recent call last):' or stripped.startswith('Traceback ('):
                _em_traceback = True
                continue
            if _em_traceback:
                # Linha de frame: File "...", linha de código indentada, ou continuação
                if (stripped.startswith('File "') or
                        linha.startswith('    ') or
                        linha.startswith('  File') or
                        stripped.startswith('During handling')):
                    continue
                # Linha final do traceback (ex: "RuntimeError: ..." ou "EOFError")
                if re.match(r'^[A-Z][a-zA-Z]+Error', stripped) or re.match(r'^[A-Z][a-zA-Z]+Exception', stripped):
                    _em_traceback = False
                    continue
                _em_traceback = False
            _logs[jid].append(linha)

            if is_mei:
                # Formato MEI: "HH:MM:SS | INFO     | [1/50] Nome  |  CNPJ: ..."
                m = re.search(r'\[(\d+)/(\d+)\]', linha)
                if m:
                    try:
                        job["total"] = int(m.group(2))
                    except Exception:
                        pass
                if "CNPJ:" in linha:
                    try:
                        cnpj_atual = linha.split("CNPJ:")[-1].strip().split()[0]
                    except Exception:
                        pass
                # Linhas de resultado têm formato: "| INFO | [✓/✗] STATUS — detalhes"
                if "| INFO" in linha and "SUCESSO" in linha and "—" in linha:
                    job["sucesso"] = job.get("sucesso", 0) + 1
                if "| INFO" in linha and "ERRO" in linha and "—" in linha:
                    job["erros"] = job.get("erros", 0) + 1
                    partes = linha.split("—", 1)
                    motivo = partes[1].strip()[:120] if len(partes) > 1 else linha.split("|")[-1].strip()[:120]
                    if cnpj_atual:
                        erros_detalhe.append({"cnpj": cnpj_atual, "motivo": motivo})
            else:
                # Rastreia CNPJ atual: "[1/141] CNPJ: 12345678000195 | ..."
                if linha.startswith("[") and "CNPJ:" in linha:
                    try:
                        cnpj_atual = linha.split("CNPJ:")[1].split("|")[0].strip()
                    except Exception:
                        pass

                if "--- Processamento finalizado ---" in linha:
                    em_relatorio = True
                if em_relatorio:
                    relatorio_linhas.append(linha)

                if "CNPJs selecionados" in linha:
                    try:
                        job["total"] = int(linha.split("|")[1].strip().split()[0])
                    except Exception:
                        pass
                if ("Certidão" in linha or "Certidao" in linha) and "salva:" in linha:
                    job["sucesso"] = job.get("sucesso", 0) + 1
                if "TIMEOUT:" in linha:
                    job["erros"] = job.get("erros", 0) + 1
                    if cnpj_atual:
                        erros_detalhe.append({"cnpj": cnpj_atual, "motivo": "Timeout — portal demorou demais"})
                elif "ERRO:" in linha:
                    job["erros"] = job.get("erros", 0) + 1
                    motivo = linha.split("ERRO:", 1)[-1].strip()[:120]
                    if cnpj_atual:
                        erros_detalhe.append({"cnpj": cnpj_atual, "motivo": f"Erro: {motivo}"})
                elif "não encontrado na base" in linha.lower() or "nao encontrado" in linha.lower():
                    if cnpj_atual:
                        erros_detalhe.append({"cnpj": cnpj_atual, "motivo": "CNPJ não encontrado no portal"})
                elif "com débitos" in linha.lower() or "com debitos" in linha.lower():
                    if cnpj_atual:
                        erros_detalhe.append({"cnpj": cnpj_atual, "motivo": "CNPJ com débitos"})
                elif "sem portal" in linha.lower() or "não suportado" in linha.lower() or "nao suportado" in linha.lower():
                    if cnpj_atual:
                        erros_detalhe.append({"cnpj": cnpj_atual, "motivo": "Cidade sem portal cadastrado"})
                elif "captcha falhou" in linha.lower():
                    if cnpj_atual:
                        erros_detalhe.append({"cnpj": cnpj_atual, "motivo": "Captcha não resolvido"})

        proc.wait()
        if job["status"] != "cancelado":
            job["status"] = "concluido"
    except Exception as e:
        _logs[jid].append(f"  ERRO INTERNO: {e}")
        job["status"] = "erro"
    finally:
        _proc_atual = None

    job["erros_detalhe"]  = erros_detalhe
    job["relatorio_txt"]  = "\n".join(relatorio_linhas)
    job["fim"] = datetime.now().strftime("%d/%m %H:%M")

    if is_mei:
        downloads_dir = os.path.join(MEI_DIR, "downloads")
        pdfs = []
        if os.path.exists(downloads_dir):
            for root, dirs, files in os.walk(downloads_dir):
                for arq in files:
                    if arq.endswith(".pdf"):
                        rel = os.path.relpath(os.path.join(root, arq), downloads_dir).replace(os.sep, "/")
                        pdfs.append(rel)
        job["pdfs"]      = pdfs
        job["pasta_job"] = downloads_dir
    else:
        pasta_job = os.path.join(PASTA_CERT, jid)
        pdfs = []
        for sub in ("negativas", "positivas"):
            pasta_sub = os.path.join(pasta_job, sub)
            if os.path.exists(pasta_sub):
                for cidade in os.listdir(pasta_sub):
                    pasta_cidade = os.path.join(pasta_sub, cidade)
                    if os.path.isdir(pasta_cidade):
                        for arq in os.listdir(pasta_cidade):
                            if arq.endswith(".pdf"):
                                pdfs.append(f"{sub}/{cidade}/{arq}")
        job["pdfs"]      = pdfs
        job["pasta_job"] = pasta_job

    with _lock:
        historico.insert(0, dict(job))
        if len(historico) > 50:
            historico[:] = historico[:50]
        _salvar()
        job_atual = None
        _verificar_fila()

def _verificar_fila():
    global job_atual
    if job_atual is None and fila:
        job_atual = fila.pop(0)
        threading.Thread(target=_rodar_job, args=(job_atual,), daemon=True).start()

# ── rotas ─────────────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    erro = None
    if request.method == "POST":
        if request.form.get("senha") == SENHA:
            session["autenticado"] = True
            return redirect(url_for("index"))
        erro = "Senha incorreta. Tente novamente."
    return render_template("login.html", erro=erro)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
@login_required
def index():
    return render_template("index.html")

@app.route("/api/status")
@login_required
def api_status():
    with _lock:
        return jsonify({
            "job_atual": job_atual,
            "fila":      list(fila),
            "historico": historico[:20],
        })

@app.route("/api/logs/<jid>")
@login_required
def api_logs(jid):
    offset    = int(request.args.get("offset", 0))
    linhas    = _logs.get(jid, [])
    concluido = (
        job_atual is None or job_atual.get("id") != jid
    ) and any(j["id"] == jid for j in historico)
    return jsonify({
        "linhas":    linhas[offset:],
        "total":     len(linhas),
        "concluido": concluido,
    })

@app.route("/api/iniciar", methods=["POST"])
@login_required
def api_iniciar():
    data    = request.get_json() or {}
    usuario = (data.get("usuario") or "Usuário").strip()[:40]
    modo    = data.get("modo", "todas")
    cnpj    = re.sub(r"[.\-/]", "", data.get("cnpj") or "").strip()

    if modo not in ("todas", "apiacas", "outras", "cnpj"):
        return jsonify({"erro": "Modo inválido"}), 400
    if modo == "cnpj" and not cnpj:
        return jsonify({"erro": "Informe o CNPJ"}), 400

    job = {
        "id":      uuid.uuid4().hex[:8],
        "usuario": usuario,
        "modo":    modo,
        "cnpj":    cnpj,
        "status":  "aguardando",
        "inicio":  None, "fim":    None,
        "total":   0,    "sucesso": 0,  "erros": 0,
        "pdfs":    [],
    }
    with _lock:
        fila.append(job)
        _verificar_fila()
    return jsonify({"ok": True, "job_id": job["id"]})

@app.route("/api/parar", methods=["POST"])
@login_required
def api_parar():
    global _proc_atual, job_atual
    with _lock:
        if _proc_atual:
            _proc_atual.kill()
            _proc_atual = None
        if job_atual:
            job_atual["status"] = "cancelado"
            _logs.get(job_atual["id"], []).append("  *** Processamento cancelado pelo usuário. ***")
    return jsonify({"ok": True})

@app.route("/api/download/<jid>")
@login_required
def api_download(jid):
    job = next((j for j in historico if j["id"] == jid), None)
    if not job:
        abort(404)
    pasta_job = job.get("pasta_job", PASTA_CERT)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for rel in job.get("pdfs", []):
            caminho = os.path.join(pasta_job, rel.replace("/", os.sep))
            if os.path.exists(caminho):
                zf.write(caminho, rel)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"certidoes_{jid}.zip",
                     mimetype="application/zip")

# ── cidades / portais ────────────────────────────────────────────────────────

def _nome_cidade(cidade_uf):
    """Extrai, remove acentos e normaliza o nome da cidade de 'Sinop/MT' → 'Sinop'."""
    partes = str(cidade_uf).strip().replace("-", "/").split("/")
    nome = partes[0].strip()
    nome = unicodedata.normalize("NFD", nome)
    nome = "".join(c for c in nome if unicodedata.category(c) != "Mn")
    return nome.title()

def _estado_cidade(cidade_uf):
    partes = str(cidade_uf).strip().replace("-", "/").split("/")
    return partes[-1].strip().upper()[:2] if len(partes) > 1 else ""

@app.route("/api/cidades", methods=["GET"])
@login_required
def api_cidades():
    with _cidades_lock:
        cfg = _ler_cidades()
    return jsonify({
        "configs":       cfg,
        "estados_betha": list(ESTADOS_BETHA),
    })

# ── gerenciamento de empresas ─────────────────────────────────────────────────

@app.route("/api/empresas", methods=["GET"])
@login_required
def api_empresas_listar():
    with _excel_lock:
        df = _ler_excel()
    registros = []
    for _, row in df.iterrows():
        cnpj = str(row.get(COLUNA_CNPJ, "") or "").strip()
        if cnpj:
            registros.append({
                "cnpj":   cnpj,
                "nome":   str(row.get(COLUNA_NOME, "") or "").strip(),
                "cidade": str(row.get(COLUNA_CIDADE, "") or "").strip(),
            })
    return jsonify({"empresas": registros})

@app.route("/api/empresas", methods=["POST"])
@login_required
def api_empresas_adicionar():
    data   = request.get_json() or {}
    cnpj   = re.sub(r"[.\-/\s]", "", data.get("cnpj") or "").strip()
    nome   = (data.get("nome") or "").strip()[:120]
    cidade = (data.get("cidade") or "").strip()[:80]

    if not cnpj or len(cnpj) != 14 or not cnpj.isdigit():
        return jsonify({"erro": "CNPJ inválido — informe os 14 dígitos"}), 400
    if not cidade:
        return jsonify({"erro": "Informe a Cidade/UF (ex: Sinop/MT)"}), 400

    # Configuração de portal (opcional — só quando a cidade é nova)
    sistema    = (data.get("sistema") or "").strip().lower()
    url_portal = (data.get("url_portal") or "").strip()
    perfil_gpsrv = str(data.get("perfil_gpsrv") or "A").upper()

    cnpj_fmt   = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    nome_cidade = _nome_cidade(cidade)
    estado      = _estado_cidade(cidade)

    # Se enviou configuração de portal, salva no JSON
    if sistema in ("agili", "i7sgp", "gpsrv") and url_portal:
        nova_cfg = {"sistema": sistema, "url": url_portal}
        if sistema == "gpsrv":
            if perfil_gpsrv == "B":
                nova_cfg.update({"tipo_certidao": "2", "clicar_emitir": False, "clicar_table_finalidade": True})
            else:
                nova_cfg.update({"tipo_certidao": "1", "clicar_emitir": True,  "clicar_table_finalidade": False})
        with _cidades_lock:
            cfg = _ler_cidades()
            cfg[nome_cidade] = nova_cfg
            _salvar_cidades(cfg)

    # Salva empresa no Excel
    with _excel_lock:
        df = _ler_excel()
        if COLUNA_CNPJ in df.columns:
            ja_existe = df[COLUNA_CNPJ].str.replace(r"[.\-/\s]", "", regex=True).str.strip().eq(cnpj)
            if ja_existe.any():
                return jsonify({"erro": "CNPJ já cadastrado na planilha"}), 409
        nova = {COLUNA_CNPJ: cnpj_fmt, COLUNA_NOME: nome, COLUNA_CIDADE: cidade}
        df = pd.concat([df, pd.DataFrame([nova])], ignore_index=True)
        _salvar_excel(df)

    # Avisa se a cidade ficou sem portal (não é MT/MS e não foi configurada)
    with _cidades_lock:
        cfg = _ler_cidades()
    tem_portal = nome_cidade in cfg or estado in ESTADOS_BETHA
    return jsonify({"ok": True, "sem_portal": not tem_portal})

@app.route("/api/iniciar_geral", methods=["POST"])
@login_required
def api_iniciar_geral():
    data    = request.get_json() or {}
    usuario = (data.get("usuario") or "Usuário").strip()[:40]
    tipos   = [t for t in (data.get("tipos") or [])
               if t in ("fgts", "trabalhista", "estadual_mt")]
    if not tipos:
        return jsonify({"erro": "Selecione ao menos um tipo de certidão"}), 400
    job = {
        "id":      uuid.uuid4().hex[:8],
        "usuario": usuario,
        "script":  "robo_geral.py",
        "modo":    ",".join(tipos),
        "status":  "aguardando",
        "inicio":  None, "fim": None,
        "total":   0,    "sucesso": 0, "erros": 0,
        "pdfs":    [],
    }
    with _lock:
        fila.append(job)
        _verificar_fila()
    return jsonify({"ok": True, "job_id": job["id"]})

@app.route("/api/iniciar_mei", methods=["POST"])
@login_required
def api_iniciar_mei():
    data    = request.get_json() or {}
    usuario = (data.get("usuario") or "Usuário").strip()[:40]
    job = {
        "id":      uuid.uuid4().hex[:8],
        "usuario": usuario,
        "script":  "mei",
        "modo":    "mei",
        "status":  "aguardando",
        "inicio":  None, "fim": None,
        "total":   0,    "sucesso": 0, "erros": 0,
        "pdfs":    [],
    }
    with _lock:
        fila.append(job)
        _verificar_fila()
    return jsonify({"ok": True, "job_id": job["id"]})

@app.route("/api/empresas/<cnpj>", methods=["DELETE"])
@login_required
def api_empresas_remover(cnpj):
    cnpj = re.sub(r"[.\-/\s]", "", cnpj).strip()
    if not cnpj or not cnpj.isdigit():
        return jsonify({"erro": "CNPJ inválido"}), 400
    with _excel_lock:
        df = _ler_excel()
        if COLUNA_CNPJ not in df.columns:
            return jsonify({"erro": "Planilha sem coluna de CNPJ"}), 400
        cnpjs_norm = df[COLUNA_CNPJ].str.replace(r"[.\-/\s]", "", regex=True).str.strip()
        antes = len(df)
        df = df[cnpjs_norm != cnpj].reset_index(drop=True)
        if len(df) == antes:
            return jsonify({"erro": "CNPJ não encontrado"}), 404
        _salvar_excel(df)
    return jsonify({"ok": True})

# ── inicialização ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    try:
        ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        ip = "localhost"
    print("=" * 52)
    print("  Robô de Certidões — Servidor Web")
    print(f"  Acesse neste computador:  http://localhost:{port}")
    print(f"  Acesse na rede:           http://{ip}:{port}")
    print("=" * 52)
    app.run(host="0.0.0.0", port=port, debug=False)
