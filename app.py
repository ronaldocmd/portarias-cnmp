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
    Tenta baixar o PDF certificado a partir da URL fornecida e atualiza o JOB.
    """
    data_in = request.get_json(silent=True) or request.form
    job_id = data_in.get("job_id")
    cert_url = (data_in.get("cert_url") or "").strip()

    if not job_id or not cert_url:
        return jsonify({"error": "job_id e cert_url são obrigatórios."}), 400

    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job or not job.get("result"):
            return jsonify({"error": "job não encontrado ou sem resultado ainda."}), 404
        numero = job["numero"]
        ano = job["ano"]

    try:
        # Usa a mesma sessão / cliente do portaria_formatter
        session = pf.make_session()
        dou = pf.DOUClient(session)

        # Se o usuário colou diretamente o servlet INPDFViewer, usamos tal qual;
        # se colou a URL 'visualiza', construímos o servlet com _build_pdf_url.
        if "INPDFViewer" in cert_url or "servlet/INPDFViewer" in cert_url:
            servlet = cert_url
            cert = cert_url if "visualiza" in cert_url else ""  # pode ser servlet direto
        else:
            servlet = dou._build_pdf_url(cert_url)
            cert = cert_url

        # Monta um PortariaData mínimo para passar ao download_certified_pdf
        pdata = pf.PortariaData(numero=numero, ano=ano)
        pdata.dou_cert_url = cert or ""
        pdata.dou_pdf_servlet_url = servlet

        # Caminho onde salvar o PDF (mesma convenção do gerador)
        pdf_name = f"{ano}.Portaria-CNMP-PRESI.{numero}-DOU-certificada.pdf"
        pdf_path = os.path.join(OUTPUT_DIR, pdf_name)

        res = dou.download_certified_pdf(pdata, pdf_path)
        if res:
            with JOBS_LOCK:
                job = JOBS.get(job_id)
                job["result"]["dou_cert_url"] = cert or pdata.dou_cert_url
                job["result"]["dou_pdf_servlet_url"] = servlet
                job["result"]["pdf_path"] = pdf_path
                job["log"].append(f"PDF certificado baixado manualmente: {pdf_path}")
            return jsonify({"ok": True, "pdf_path": pdf_path})
        else:
            with JOBS_LOCK:
                job = JOBS.get(job_id)
                job["log"].append("Falha ao baixar PDF usando a URL fornecida.")
            return jsonify({"error": "falha ao baixar o PDF com a URL fornecida."}), 500

    except Exception as exc:  # noqa: BLE001
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if job:
                job["log"].append("ERRO ao baixar PDF com url manual: " + str(exc))
        return jsonify({"error": str(exc)}), 500
