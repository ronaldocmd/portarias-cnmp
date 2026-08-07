#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interface web (Flask) para geração automatizada de Portarias da Presidência do CNMP.

Fluxo:
  - O usuário informa apenas o NÚMERO e o ANO da portaria.
  - O backend executa `gerar_portaria` (busca no DOU, baixa a versão certificada,
    extrai o texto, busca links das portarias mencionadas no portal do CNMP e
    gera o .docx formatado).
  - A geração roda em uma thread; o front-end faz polling do progresso e, ao
    final, oferece o download do .docx (e do PDF certificado, quando houver).

Uso:
    python3 app.py            # sobe em 0.0.0.0:3000
"""

import io
import os
import threading
import traceback
import uuid
from contextlib import redirect_stdout
from datetime import datetime

from flask import (
    Flask, jsonify, render_template, request, send_file, abort
)

import portaria_formatter as pf

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)

# Estado dos jobs em memória. { job_id: {status, log, result, error, started} }
JOBS = {}
JOBS_LOCK = threading.Lock()


class _JobLogWriter(io.StringIO):
    """Captura o stdout do gerador e acumula linhas no estado do job."""

    def __init__(self, job_id):
        super().__init__()
        self.job_id = job_id

    def write(self, s):
        super().write(s)
        if s and s.strip():
            with JOBS_LOCK:
                job = JOBS.get(self.job_id)
                if job is not None:
                    job["log"].append(s.rstrip("\n"))
        return len(s)


def _run_job(job_id, numero, ano):
    writer = _JobLogWriter(job_id)
    try:
        with redirect_stdout(writer):
            result = pf.gerar_portaria(numero, ano, output_dir=OUTPUT_DIR)
        with JOBS_LOCK:
            job = JOBS[job_id]
            job["status"] = "done"
            job["result"] = result
    except Exception as exc:  # noqa: BLE001
        tb = traceback.format_exc()
        with JOBS_LOCK:
            job = JOBS[job_id]
            job["status"] = "error"
            job["error"] = str(exc)
            job["log"].append("ERRO: " + str(exc))
        print(tb)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/gerar", methods=["POST"])
def gerar():
    data = request.get_json(silent=True) or request.form
    try:
        numero = int(str(data.get("numero", "")).strip())
        ano = int(str(data.get("ano", "")).strip())
    except (TypeError, ValueError):
        return jsonify({"error": "Informe número e ano válidos."}), 400

    if numero <= 0 or ano < 2000 or ano > 2100:
        return jsonify({"error": "Número ou ano fora do intervalo esperado."}), 400

    job_id = uuid.uuid4().hex
    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "running",
            "log": [],
            "result": None,
            "error": None,
            "numero": numero,
            "ano": ano,
            "started": datetime.now().isoformat(timespec="seconds"),
        }

    t = threading.Thread(target=_run_job, args=(job_id, numero, ano), daemon=True)
    t.start()
    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job is None:
            return jsonify({"error": "job não encontrado"}), 404
        payload = {
            "status": job["status"],
            "log": job["log"][-40:],
            "error": job["error"],
            "numero": job["numero"],
            "ano": job["ano"],
        }
        result = job.get("result")
        if result:
            payload["result"] = {
                "titulo": result.get("titulo"),
                "publicado_dou": result.get("publicado_dou"),
                "publicacao_campo": result.get("publicacao_campo"),
                "dou_link_impreciso": result.get("dou_link_impreciso"),
                "dou_indisponivel": result.get("dou_indisponivel"),
                "dou_cert_url": result.get("dou_cert_url"),
                "dou_page_url": result.get("dou_page_url"),
                "norma_url": result.get("norma_url"),
                "n_paragrafos_corpo": result.get("n_paragrafos_corpo"),
                "assinaturas": result.get("assinaturas"),
                "has_docx": bool(result.get("docx_path") and os.path.exists(result.get("docx_path"))),
                "has_pdf": bool(result.get("pdf_path") and os.path.exists(result.get("pdf_path") or "")),
            }
    return jsonify(payload)


@app.route("/download/<job_id>/<tipo>")
def download(job_id, tipo):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job is None or not job.get("result"):
            abort(404)
        result = job["result"]
    if tipo == "docx":
        path = result.get("docx_path")
    elif tipo == "pdf":
        path = result.get("pdf_path")
    else:
        abort(404)
    if not path or not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "3000"))
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)

@app.route("/attach_cert_url", methods=["POST"])
def attach_cert_url():
    """
    Recebe JSON/form com: job_id, cert_url
    Tenta baixar o PDF certificado a partir da URL fornecida, extrai texto
    (quando possível) e re-gera o .docx usando o PDF certificado — atualiza o JOB.
    """
    data_in = request.get_json(silent=True) or request.form
    job_id = data_in.get("job_id")
    cert_url = (data_in.get("cert_url") or "").strip()

    if not job_id or not cert_url:
        return jsonify({"error": "job_id e cert_url são obrigatórios."}), 400

    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({"error": "job não encontrado."}), 404
        numero = job["numero"]
        ano = job["ano"]

    try:
        session = pf.make_session()
        dou = pf.DOUClient(session)
        cnmp = pf.CNMPClient(session)

        # Se o usuário colou diretamente o servlet do PDF, usamos tal qual;
        # se colou a URL 'visualiza', construímos o servlet com _build_pdf_url.
        if "INPDFViewer" in cert_url or "servlet/INPDFViewer" in cert_url:
            servlet = cert_url
            cert = cert_url if "visualiza" in cert_url else ""
        else:
            # tenta construir servlet a partir da visualiza
            servlet = dou._build_pdf_url(cert_url)
            cert = cert_url

        # Monta PortariaData mínimo
        pdata = pf.PortariaData(numero=numero, ano=ano)
        pdata.dou_cert_url = cert or ""
        pdata.dou_pdf_servlet_url = servlet or ""

        # Caminho onde salvar o PDF
        pdf_name = f"{ano}.Portaria-CNMP-PRESI.{numero}-DOU-certificada.pdf"
        pdf_path = os.path.join(OUTPUT_DIR, pdf_name)

        # Baixa o PDF
        res_pdf = dou.download_certified_pdf(pdata, pdf_path)
        if not res_pdf:
            return jsonify({"error": "Falha ao baixar o PDF certificado."}), 500

        # Tenta extrair texto do PDF para preencher título/corpo (melhora o .docx)
        texto = pf.DOUClient.extract_pdf_text(pdf_path)
        if texto:
            parsed = pf.parse_portaria_text(texto, numero, ano)
            pdata.titulo = parsed.get("titulo") or pdata.titulo
            pdata.ementa = parsed.get("ementa", "")
            pdata.preambulo = parsed.get("preambulo", "")
            pdata.corpo = parsed.get("corpo", [])
            pdata.assinaturas = parsed.get("assinaturas", [])
            pdata.cargos = parsed.get("cargos", [])
            pdata.fonte_texto = "DOU-PDF"

        # Re-gera o .docx usando o template (garante que o título terá o link correto)
        builder = pf.PortariaDocBuilder(pf.TEMPLATE_PATH, cnmp)
        builder.build(pdata)
        docx_name = f"{ano}.Portaria-CNMP-PRESI.{numero}.docx"
        docx_path = os.path.join(OUTPUT_DIR, docx_name)
        builder.save(docx_path)

        # Atualiza o job em memória
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if job.get("result") is None:
                job["result"] = {}
            job["result"].update({
                "dou_cert_url": pdata.dou_cert_url,
                "dou_pdf_servlet_url": pdata.dou_pdf_servlet_url,
                "dou_page_url": pdata.dou_page_url,
                "pdf_path": pdf_path,
                "docx_path": docx_path,
            })

        return jsonify({"ok": True, "pdf_path": pdf_path, "docx_path": docx_path})

    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        log(f"attach_cert_url: exceção: {exc}\n{tb}")
        return jsonify({"error": "erro interno ao anexar a URL do DOU."}), 500
