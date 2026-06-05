from flask import Flask, render_template, request, jsonify, send_file, abort
import subprocess, threading, json, os, sys, uuid, io, zipfile, re, time, socket
from datetime import datetime

app = Flask(__name__)

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
PASTA_CERT = os.path.join(os.path.expanduser("~"), "Documents", "certidões")
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

    env = os.environ.copy()
    env["ROBO_HEADLESS"]    = "1"
    env["ROBO_WEB"]         = "1"
    env["PYTHONUNBUFFERED"] = "1"

    args = [sys.executable, os.path.join(BASE_DIR, "robo.py")]
    if job["modo"] == "cnpj":
        args += ["cnpj", job.get("cnpj", "")]
    else:
        args.append(job["modo"])

    job["inicio"] = datetime.now().strftime("%d/%m %H:%M")
    job["status"] = "rodando"

    cnpj_atual     = ""
    erros_detalhe  = []
    em_relatorio   = False
    relatorio_linhas = []

    try:
        proc = subprocess.Popen(
            args,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace",
            env=env, cwd=BASE_DIR,
        )
        _proc_atual = proc
        for linha in proc.stdout:
            linha = linha.rstrip()
            if not linha:
                continue
            _logs[jid].append(linha)

            # Rastreia CNPJ atual: "[1/141] CNPJ: 12345678000195 | ..."
            if linha.startswith("[") and "CNPJ:" in linha:
                try:
                    cnpj_atual = linha.split("CNPJ:")[1].split("|")[0].strip()
                except Exception:
                    pass

            # Captura bloco do relatório final
            if "--- Processamento finalizado ---" in linha:
                em_relatorio = True
            if em_relatorio:
                relatorio_linhas.append(linha)

            if "CNPJs selecionados" in linha:
                try:
                    job["total"] = int(linha.split("|")[1].strip().split()[0])
                except Exception:
                    pass
            if "Certidão" in linha and "salva:" in linha:
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
            elif "com débitos" in linha.lower():
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

    pdfs = []
    for sub in ("negativas", "positivas"):
        pasta = os.path.join(PASTA_CERT, sub)
        if os.path.exists(pasta):
            for arq in os.listdir(pasta):
                if arq.endswith(".pdf"):
                    caminho = os.path.join(pasta, arq)
                    if os.path.getmtime(caminho) >= inicio_ts:
                        pdfs.append(f"{sub}/{arq}")
    job["pdfs"] = pdfs

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

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/status")
def api_status():
    with _lock:
        return jsonify({
            "job_atual": job_atual,
            "fila":      list(fila),
            "historico": historico[:20],
        })

@app.route("/api/logs/<jid>")
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
def api_download(jid):
    job = next((j for j in historico if j["id"] == jid), None)
    if not job:
        abort(404)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for rel in job.get("pdfs", []):
            caminho = os.path.join(PASTA_CERT, rel.replace("/", os.sep))
            if os.path.exists(caminho):
                zf.write(caminho, rel)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"certidoes_{jid}.zip",
                     mimetype="application/zip")

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
