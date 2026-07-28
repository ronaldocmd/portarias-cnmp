#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
portaria_formatter.py
=====================

Automação completa para gerar Portarias da Presidência do CNMP formatadas
conforme o padrão institucional (Manual de padronização de atos do CNMP).

Dado apenas o NÚMERO e o ANO de uma Portaria CNMP-PRESI, o script:

  1. Busca a portaria no Diário Oficial da União (www.in.gov.br), localiza a
     página web da matéria e a URL da *versão certificada* do DOU.
  2. Baixa o PDF da versão certificada do DOU.
  3. Extrai o texto integral da portaria (título, ementa, preâmbulo, artigos,
     assinatura) a partir da página oficial do DOU.
  4. Busca no portal do CNMP (www.cnmp.mp.br/portal/atos-e-normas/) os links
     das portarias eventualmente mencionadas no corpo do texto.
  5. Gera um arquivo .docx formatado a partir do modelo institucional,
     preservando cabeçalho (brasão + "Conselho Nacional do Ministério
     Público"), rodapé, margens e estilos ("Cabeçalho Presi" e "Corpo Presi"),
     inserindo os hiperlinks apropriados:
        - título -> hyperlink azul para a versão certificada do DOU
          (ou texto vermelho, sem link, quando publicada apenas no Diário
          Eletrônico do CNMP);
        - portarias mencionadas -> hyperlink para a página da norma no portal
          do CNMP.
  6. Salva o .docx (e o PDF certificado) na pasta output/.

Uso:
    python3 portaria_formatter.py 164 2026
    python3 portaria_formatter.py --numero 164 --ano 2026

Dependências: python-docx, pdfplumber, requests, beautifulsoup4, lxml
"""

import argparse
import os
import re
import sys
import time
import unicodedata
import urllib.parse
from dataclasses import dataclass, field
from typing import List, Optional

import requests
import urllib3
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# python-docx
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --------------------------------------------------------------------------- #
# Configurações gerais
# --------------------------------------------------------------------------- #

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(HERE, "template_model.docx")
OUTPUT_DIR = os.path.join(HERE, "output")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT, "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8"}

# Cores
COLOR_LINK_BLUE = "0563C1"   # azul padrão de hyperlink do Word / template
COLOR_CNMP_RED = "EE0000"    # vermelho usado quando publicada no Diário Eletrônico do CNMP
COLOR_BLACK = "000000"

# Categorias de atos no portal do CNMP (atos-e-normas)
CNMP_CATEGORIES = {
    "PRESI": "389",   # Portarias da Presidência
    "CN": "477",      # Portarias da Corregedoria
    "SG": "391",      # Portarias da Secretaria Geral
}


def log(msg: str) -> None:
    print(f"[portaria] {msg}", flush=True)


def make_session() -> requests.Session:
    """Cria uma sessão HTTP com retentativas automáticas e backoff."""
    session = requests.Session()
    session.headers.update(HEADERS)
    retry = Retry(
        total=1,
        connect=1,
        read=1,
        status=1,
        backoff_factor=0.5,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET", "HEAD"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def get_with_retry(session: requests.Session, url: str, *, attempts: int = 2,
                   **kwargs) -> Optional[requests.Response]:
    """GET com poucas retentativas manuais e backoff curto.

    Objetivo: falhar rápido quando um servidor (em especial o in.gov.br)
    estiver instável, sem travar a experiência do usuário.
    """
    last_exc = None
    for i in range(attempts):
        try:
            r = session.get(url, **kwargs)
            if r.status_code < 500:
                return r
            last_exc = requests.HTTPError(f"{r.status_code} para {url}")
        except requests.RequestException as exc:
            last_exc = exc
        if i < attempts - 1:
            wait = 1.5 * (i + 1)
            log(f"HTTP {url}: tentativa {i + 1}/{attempts} falhou ({last_exc}). "
                f"Aguardando {wait:.0f}s...")
            time.sleep(wait)
        else:
            log(f"HTTP {url}: tentativa {i + 1}/{attempts} falhou ({last_exc}).")
    return None


def strip_accents_upper(text: str) -> str:
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).upper()



# --------------------------------------------------------------------------- #
# Estruturas de dados
# --------------------------------------------------------------------------- #

@dataclass
class PortariaData:
    """Conteúdo extraído de uma portaria."""
    numero: int
    ano: int
    titulo: str = ""                       # epígrafe (ex.: "PORTARIA CNMP-PRESI N° 164 DE 28 DE MAIO DE 2026")
    ementa: str = ""                       # opcional ("Dispõe sobre ...")
    notas: List[str] = field(default_factory=list)  # avisos "Vide/Revogada" após o título
    preambulo: str = ""                    # parágrafo que termina em RESOLVE:
    corpo: List[str] = field(default_factory=list)   # artigos / parágrafos
    assinaturas: List[str] = field(default_factory=list)  # nomes
    cargos: List[str] = field(default_factory=list)       # cargos/funções
    # DOU
    publicado_dou: bool = False
    dou_page_url: str = ""                 # página da matéria no in.gov.br
    dou_cert_url: str = ""                 # página web da versão certificada
    dou_pdf_servlet_url: str = ""          # URL do PDF certificado (servlet)
    pub_date: str = ""                     # dd/mm/aaaa da publicação no DOU
    dou_indisponivel: bool = False         # True se a busca do DOU falhou por erro de servidor
    dou_link_impreciso: bool = False       # link aponta para a edição (não a página exata)
    # CNMP
    cnmp_pdf_url: str = ""                 # URL do PDF da portaria no portal do CNMP
    norma_url: str = ""                    # página da norma no portal do CNMP
    publicacao_campo: str = ""             # texto do campo "Publicação:" da norma
    fonte_texto: str = ""                  # de onde veio o texto: "CNMP-PDF" | "DOU-HTML" | ""


# --------------------------------------------------------------------------- #
# Cliente do DOU (in.gov.br)
# --------------------------------------------------------------------------- #

class DOUClient:
    """Busca e extração de portarias publicadas no Diário Oficial da União."""

    SEARCH_URL = "https://www.in.gov.br/consulta/-/buscar/dou"
    MATERIA_URL = "https://www.in.gov.br/web/dou/-/{url_title}"
    SCRIPT_ID = "_br_com_seatecnologia_in_buscadou_BuscaDouPortlet_params"

    def __init__(self, session: Optional[requests.Session] = None):
        self.session = session or make_session()
        self.server_error = False   # True se o in.gov.br falhou por erro de servidor

    # -- busca ------------------------------------------------------------- #
    def _raw_search(self, query: str) -> List[dict]:
        params = {
            "q": query,
            "delta": "20",
            "currentPage": "1",
            "exactDate": "all",
            "sortType": "0",
        }
        # A busca do in.gov.br sofre 502 transitórios; como o resultado define a
        # cor/link do título, vale a pena insistir um pouco mais aqui.
        r = get_with_retry(self.session, self.SEARCH_URL, params=params,
                           timeout=25, attempts=4)
        if r is None or r.status_code != 200:
            self.server_error = True
            log(f"Falha na busca do DOU (query={query!r}).")
            return []
        m = re.search(
            r'<script id="%s" type="application/json">\s*(\{.*?\})\s*</script>'
            % re.escape(self.SCRIPT_ID),
            r.text,
            re.DOTALL,
        )
        if not m:
            return []
        import json
        try:
            data = json.loads(m.group(1))
        except json.JSONDecodeError:
            return []
        return data.get("jsonArray", []) or []

    @staticmethod
    def _clean_html_text(text: str) -> str:
        return re.sub(r"<[^>]+>", "", text or "").strip()

    def find(self, numero: int, ano: int) -> Optional[dict]:
        """Localiza a entrada do DOU correspondente à portaria numero/ano."""
        self.server_error = False
        queries = [
            f'"PORTARIA CNMP-PRESI N° {numero}"',
            f'PORTARIA CNMP-PRESI {numero} {ano}',
        ]
        seen = set()
        for q in queries:
            for item in self._raw_search(q):
                key = item.get("urlTitle")
                if not key or key in seen:
                    continue
                seen.add(key)
                title = self._clean_html_text(item.get("title", ""))
                up = strip_accents_upper(title)
                if "CNMP-PRESI" not in up:
                    continue
                # número (após N° / Nº)
                mnum = re.search(r"N[°ºO]\s*(\d+)", up)
                if not mnum or int(mnum.group(1)) != int(numero):
                    continue
                # ano (4 dígitos no título)
                years = [int(y) for y in re.findall(r"\b(19\d{2}|20\d{2})\b", title)]
                if not years or ano not in years:
                    continue
                log(f"DOU: encontrada matéria '{title}' (pub. {item.get('pubDate')}, "
                    f"{item.get('pubName')})")
                return item
        return None

    # -- leiturajornal (índice completo do DOU por edição/seção) ---------- #
    # Códigos de "jornal" usados pelo pesquisa.in.gov.br por seção do DOU.
    JORNAL_POR_SECAO = {1: 515, 2: 529, 3: 530}
    LEITURA_URL = "https://www.in.gov.br/leiturajornal"

    def find_via_leiturajornal(self, numero: int, ano: int, pub_date: str,
                               secao: int) -> Optional[dict]:
        """Localiza a portaria no índice completo do DOU (leiturajornal).

        Diferente da busca textual do in.gov.br (que é incompleta para 2026),
        o leiturajornal lista TODAS as matérias de uma edição/seção. A partir
        dele obtemos o ``urlTitle`` e, principalmente, a PÁGINA no DOU
        (``numberPage``), necessária para montar o link da versão certificada.

        ``pub_date`` no formato ``DD/MM/AAAA``; ``secao`` = 1, 2 ou 3.
        Retorna o item (dict) do jsonArray, ou ``None``.
        """
        if not pub_date or not secao:
            return None
        d = pub_date.replace("/", "-")
        url = f"{self.LEITURA_URL}?data={d}&secao=do{secao}"
        # A edição pode ser grande (vários MB) e o in.gov.br devolve 502
        # transitórios; leitura longa + algumas tentativas.
        r = get_with_retry(self.session, url, timeout=(15, 120), attempts=4)
        if r is None or r.status_code != 200:
            self.server_error = True
            log(f"DOU: falha no leiturajornal ({url}).")
            return None
        m = re.search(
            r'<script id="params" type="application/json">(.*?)</script>',
            r.text, re.DOTALL,
        )
        if not m:
            return None
        import json
        try:
            obj = json.loads(m.group(1))
        except json.JSONDecodeError:
            return None
        for it in obj.get("jsonArray", []) or []:
            title = strip_accents_upper(
                re.sub(r"<[^>]+>", "", it.get("title", "") or ""))
            if "CNMP-PRESI" not in title:
                continue
            mn = re.search(r"N[°ºO]\s*(\d+)", title)
            if mn and int(mn.group(1)) == int(numero):
                log(f"DOU: portaria localizada no leiturajornal "
                    f"(página {it.get('numberPage')}, {it.get('pubName')}).")
                return it
        return None

    def build_cert_urls(self, pub_date: str, secao: int, pagina) -> tuple:
        """Monta (cert_url, servlet_pdf_url) da versão certificada do DOU."""
        jornal = self.JORNAL_POR_SECAO.get(int(secao), 529)
        cert = (f"https://pesquisa.in.gov.br/imprensa/jsp/visualiza/index.jsp?"
                f"data={pub_date}&jornal={jornal}&pagina={pagina}")
        servlet = (f"https://pesquisa.in.gov.br/imprensa/servlet/INPDFViewer?"
                   f"jornal={jornal}&pagina={pagina}&data={pub_date}"
                   f"&captchafield=firstAccess")
        return cert, servlet

    # -- página da matéria ------------------------------------------------- #
    def fetch_materia(self, item: dict, data: PortariaData, *,
                      only_cert: bool = False) -> None:
        """Extrai o conteúdo da página da matéria e localiza a versão certificada.

        Se ``only_cert`` for True, apenas o título (epígrafe) e o link da versão
        certificada são atualizados; o corpo/preâmbulo/assinaturas já obtidos de
        outra fonte (PDF do CNMP) são preservados.
        """
        url_title = item.get("urlTitle")
        page_url = self.MATERIA_URL.format(url_title=url_title)
        data.dou_page_url = page_url
        data.pub_date = item.get("pubDate", "")
        r = get_with_retry(self.session, page_url, timeout=40)
        if r is None or r.status_code != 200:
            self.server_error = True
            log("DOU: falha ao carregar a página da matéria.")
            return
        soup = BeautifulSoup(r.text, "lxml")

        # Título / epígrafe (autoritativo no DOU)
        ident = soup.find(class_="identifica")
        if ident:
            data.titulo = self._normalize_titulo(ident.get_text(" ", strip=True))

        if not only_cert:
            # Ementa (opcional)
            ementa = soup.find(class_="ementa")
            if ementa:
                data.ementa = ementa.get_text(" ", strip=True)

            # Parágrafos do corpo (preâmbulo + artigos)
            paras = [p.get_text(" ", strip=True) for p in soup.find_all(class_="dou-paragraph")]
            paras = [p for p in paras if p]
            if paras:
                data.preambulo = paras[0]
                data.corpo = paras[1:]

            # Assinaturas e cargos
            data.assinaturas = [a.get_text(" ", strip=True)
                                for a in soup.find_all(class_="assina") if a.get_text(strip=True)]
            data.cargos = [c.get_text(" ", strip=True)
                           for c in soup.find_all(class_="cargo") if c.get_text(strip=True)]

        # Link da versão certificada
        cert_url = None
        for a in soup.find_all("a", href=True):
            href = a["href"]
            txt = strip_accents_upper(a.get_text(" ", strip=True))
            if "pesquisa.in.gov.br" in href and ("visualiza" in href or "CERTIFICAD" in txt):
                cert_url = href
                break
        if not cert_url:
            for a in soup.find_all("a", href=True):
                if "pesquisa.in.gov.br" in a["href"]:
                    cert_url = a["href"]
                    break
        if cert_url:
            data.publicado_dou = True
            data.dou_cert_url = cert_url
            data.dou_pdf_servlet_url = self._build_pdf_url(cert_url)
        log(f"DOU: página da matéria e versão certificada localizadas "
            f"({'com' if cert_url else 'sem'} link certificado).")

    @staticmethod
    def _normalize_titulo(titulo: str) -> str:
        """Remove a vírgula após o número (padroniza como 'N° 164 DE ...')."""
        t = re.sub(r"\s+", " ", titulo).strip()
        t = re.sub(r"(N[°ºO]\s*\d+)\s*,", r"\1", t)
        return t

    @staticmethod
    def _build_pdf_url(cert_url: str) -> str:
        """Constrói a URL do servlet do PDF a partir da URL 'visualiza'."""
        parsed = urllib.parse.urlparse(cert_url)
        qs = urllib.parse.parse_qs(parsed.query)
        jornal = qs.get("jornal", [""])[0]
        pagina = qs.get("pagina", [""])[0]
        data_ = qs.get("data", [""])[0]
        base = f"{parsed.scheme}://{parsed.netloc}"
        return (f"{base}/imprensa/servlet/INPDFViewer?jornal={jornal}"
                f"&pagina={pagina}&data={data_}&captchafield=firstAccess")

    # -- download do PDF certificado -------------------------------------- #
    def download_certified_pdf(self, data: PortariaData, dest_path: str) -> Optional[str]:
        if not data.dou_pdf_servlet_url:
            log("DOU: sem URL de PDF certificado para baixar.")
            return None
        r = get_with_retry(
            self.session, data.dou_pdf_servlet_url,
            headers={"Referer": data.dou_cert_url}, timeout=90, attempts=4,
        )
        if r is None or r.status_code != 200:
            log("DOU: falha ao baixar PDF certificado.")
            return None
        if r.content[:4] != b"%PDF":
            log("DOU: resposta do servlet não é um PDF válido.")
            return None
        with open(dest_path, "wb") as fh:
            fh.write(r.content)
        log(f"DOU: PDF certificado salvo em {dest_path} ({len(r.content)} bytes).")
        return dest_path



    # -- extração de texto do PDF certificado (fallback) ------------------- #
    @staticmethod
    def extract_pdf_text(pdf_path: str) -> str:
        """Extrai o texto completo do PDF (fallback, caso a página web falhe)."""
        try:
            import pdfplumber
        except ImportError:
            return ""
        parts = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    parts.append(page.extract_text() or "")
        except Exception as exc:  # noqa: BLE001
            log(f"DOU: falha ao extrair texto do PDF: {exc}")
        return "\n".join(parts)


# --------------------------------------------------------------------------- #
# Cliente do portal do CNMP (atos-e-normas)
# --------------------------------------------------------------------------- #

class CNMPClient:
    """Busca de links de normas no portal do CNMP."""

    BUSCA_URL = "https://www.cnmp.mp.br/portal/atos-e-normas-busca"
    # Formato de link majoritário nos documentos de referência (atos-e-normas-busca)
    NORMA_URL = "https://www.cnmp.mp.br/portal/atos-e-normas-busca/norma/{id}"

    def __init__(self, session: Optional[requests.Session] = None):
        self.session = session or make_session()
        self._cache = {}

    def find_norma_url(self, numero: int, ano: int, tipo: str = "PRESI") -> Optional[str]:
        """Retorna a URL canônica da norma no portal do CNMP, ou None."""
        cache_key = (numero, ano, tipo)
        if cache_key in self._cache:
            return self._cache[cache_key]

        categoria = CNMP_CATEGORIES.get(tipo.upper(), CNMP_CATEGORIES["PRESI"])
        params = {
            "filter[numero]": str(numero),
            "filter[ano]": str(ano),
            "ano": str(ano),
            "filter[categoria][]": categoria,
            "task": "",
            "boxchecked": "0",
        }
        url = None
        r = get_with_retry(self.session, self.BUSCA_URL, params=params,
                           timeout=40, verify=False)
        if r is not None and r.status_code == 200:
            url = self._parse_results(r.text, numero, ano)
        else:
            log(f"CNMP: falha na busca ({numero}/{ano}).")

        self._cache[cache_key] = url
        return url

    @staticmethod
    def _parse_results(html: str, numero: int, ano: int) -> Optional[str]:
        soup = BeautifulSoup(html, "lxml")
        candidates = []
        for rt in soup.find_all(class_="result-title"):
            a = rt.find("a", href=True)
            if not a:
                continue
            title = rt.get_text(" ", strip=True)
            mid = re.search(r"/norma/(\d+)", a["href"])
            if not mid:
                continue
            norma_id = mid.group(1)
            up = strip_accents_upper(title)
            mnum = re.search(r"N[°ºO]\s*(\d+)", up)
            years = [int(y) for y in re.findall(r"\b(19\d{2}|20\d{2})\b", title)]
            if mnum and int(mnum.group(1)) == int(numero) and ano in years:
                candidates.append(norma_id)
        if candidates:
            return CNMPClient.NORMA_URL.format(id=candidates[0])
        return None

    # -- campo "Publicação:" da norma (sinal confiável DOU x Eletrônico) --- #
    _PUB_RE = re.compile(
        r"Publica[çc][ãa]o:\s*(.*?)(?:Categoria:|Status:|Refer[êe]ncia|"
        r"Situa[çc][ãa]o|Assunto:|$)",
        re.IGNORECASE | re.DOTALL,
    )
    _SECAO_RE = re.compile(r"Se[çc][ãa]o\s*(\d)", re.IGNORECASE)
    _DATA_RE = re.compile(r"(\d{2}/\d{2}/\d{4})")

    def get_publicacao_info(self, numero: int, ano: int, tipo: str = "PRESI"):
        """Lê o campo "Publicação:" da página da norma no portal do CNMP.

        Esse campo é a fonte MAIS confiável para decidir se a portaria foi
        publicada no Diário Oficial da União (título azul, com link) ou apenas
        no Diário Eletrônico do CNMP (título vermelho, sem link).

        Retorna um dicionário::

            {"norma_url", "publicado_dou": bool, "pub_date": "DD/MM/AAAA"|"",
             "secao": 1|2|3|None, "publicacao": "<texto do campo>"}

        ou ``None`` se a norma não for encontrada no portal.
        """
        norma_url = self.find_norma_url(numero, ano, tipo)
        if not norma_url:
            return None
        r = get_with_retry(self.session, norma_url, timeout=40, verify=False)
        if r is None or r.status_code != 200:
            log(f"CNMP: falha ao carregar a página da norma ({numero}/{ano}).")
            return None
        texto = re.sub(r"<[^>]+>", " ", r.text)
        texto = re.sub(r"\s+", " ", texto)
        m = self._PUB_RE.search(texto)
        campo = (m.group(1).strip() if m else "")[:200]
        campo_up = strip_accents_upper(campo)
        publicado_dou = "DIARIO OFICIAL DA UNIAO" in campo_up
        secao = None
        ms = self._SECAO_RE.search(campo)
        if ms:
            secao = int(ms.group(1))
        elif publicado_dou:
            secao = 2  # portarias da Presidência saem, por padrão, na Seção 2
        md = self._DATA_RE.search(campo)
        pub_date = md.group(1) if md else ""
        log(f"CNMP: norma {numero}/{ano} — Publicação: {campo!r} "
            f"(DOU={publicado_dou}, seção={secao}, data={pub_date}).")
        return {
            "norma_url": norma_url,
            "publicado_dou": publicado_dou,
            "pub_date": pub_date,
            "secao": secao,
            "publicacao": campo,
        }

    # -- PDF da portaria (fonte primária e confiável de texto) ------------- #
    PDF_URL_TEMPLATE = (
        "https://www.cnmp.mp.br/portal/images/Portarias_Presidencia_nova_versao/"
        "{ano}/{ano}.Portaria-CNMP-PRESI.{num:03d}.pdf"
    )

    def pdf_candidate_urls(self, numero: int, ano: int) -> List[str]:
        """URLs candidatas do PDF da portaria no portal do CNMP."""
        base = ("https://www.cnmp.mp.br/portal/images/"
                "Portarias_Presidencia_nova_versao/{ano}/{ano}.Portaria-CNMP-PRESI.")
        base = base.format(ano=ano)
        return [
            f"{base}{numero:03d}.pdf",
            f"{base}{numero}.pdf",
            f"{base}{numero:03d}-1.pdf",
        ]

    def download_pdf(self, numero: int, ano: int, dest_path: Optional[str] = None):
        """Baixa o PDF da portaria no portal do CNMP.

        Retorna (conteudo_bytes, url) ou (None, None) se não encontrado.
        """
        for url in self.pdf_candidate_urls(numero, ano):
            r = get_with_retry(self.session, url, timeout=60, verify=False)
            if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
                log(f"CNMP: PDF da portaria localizado ({url}).")
                if dest_path:
                    with open(dest_path, "wb") as fh:
                        fh.write(r.content)
                return r.content, url
        log(f"CNMP: PDF da portaria {numero}/{ano} não localizado no portal.")
        return None, None

    def fetch_text(self, numero: int, ano: int):
        """Baixa o PDF da portaria no CNMP e devolve (texto, url)."""
        content, url = self.download_pdf(numero, ano)
        if content is None:
            return None, None
        try:
            import io
            import pdfplumber
            parts = []
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for page in pdf.pages:
                    parts.append(page.extract_text() or "")
            return "\n".join(parts), url
        except Exception as exc:  # noqa: BLE001
            log(f"CNMP: falha ao extrair texto do PDF: {exc}")
            return None, url


# --------------------------------------------------------------------------- #
# Parsing do texto da portaria (a partir do PDF do CNMP)
# --------------------------------------------------------------------------- #

_HEADER_LINE = "CONSELHO NACIONAL DO MINISTERIO PUBLICO"
_TITLE_RE = re.compile(r"PORTARIA\s+CNMP-PRESI\s+N", re.IGNORECASE)
_NOTE_PREFIXES = ("VIDE", "REVOGADA", "REVOGADO", "TEXTO COMPILADO",
                  "(REVOGAD", "(REDACAO", "(INCLUID")
# início de novo parágrafo no corpo
_PARA_START_RE = re.compile(
    r"^(Art\.\s|§\s|Par[áa]grafo\s|Considerando|CONSIDERANDO|"
    r"CAP[ÍI]TULO\b|T[ÍI]TULO\b|SE[ÇC][ÃA]O\b|"
    r"[IVXLCDM]+\s*[–\-]\s|[a-z]\)\s)"
)
_HEADING_RE = re.compile(r"^(CAP[ÍI]TULO|T[ÍI]TULO|SE[ÇC][ÃA]O)\b", re.IGNORECASE)


def _is_header_or_page(line: str) -> bool:
    up = strip_accents_upper(line).strip()
    if not up:
        return True
    if up == _HEADER_LINE:
        return True
    if re.fullmatch(r"[-–\s]*\d+[-–\s]*", up):   # número de página isolado
        return True
    if re.fullmatch(r"FLS?\.?\s*\d+", up):
        return True
    return False


def _normalize_titulo(titulo: str) -> str:
    t = re.sub(r"\s+", " ", titulo).strip()
    t = re.sub(r"(N[°ºO]\s*\d+)\s*,", r"\1", t)      # remove vírgula após o número
    return t


_FOOTER_EPIGRAFE_RE = re.compile(
    r"^\s*PORTARIA\s+CNMP-PRESI\s+N[°ºO]?\s*(\d+)", re.IGNORECASE)


def _is_footer_epigrafe(line: str, numero: int) -> bool:
    """True se a linha é a epígrafe da PRÓPRIA portaria repetida no rodapé.

    Só considera a epígrafe que COMEÇA a linha (rodapé), com o mesmo número
    da portaria atual — assim, menções a outras portarias no meio das frases
    do corpo são preservadas.
    """
    m = _FOOTER_EPIGRAFE_RE.match(strip_accents_upper(line))
    return bool(m) and int(m.group(1)) == int(numero)


def parse_portaria_text(text: str, numero: int, ano: int) -> dict:
    """Interpreta o texto extraído do PDF em campos estruturados.

    Retorna dict com: titulo, ementa, notas (list), preambulo, corpo (list),
    assinaturas (list), cargos (list).
    """
    out = {"titulo": "", "ementa": "", "notas": [], "preambulo": "",
           "corpo": [], "assinaturas": [], "cargos": []}
    if not text:
        return out

    # 1) linhas limpas (remove cabeçalho repetido e números de página)
    raw_lines = [ln.strip() for ln in text.splitlines()]
    lines = [ln for ln in raw_lines if not _is_header_or_page(ln)]
    if not lines:
        return out

    # 2) título
    t_idx = None
    for i, ln in enumerate(lines):
        if _TITLE_RE.search(ln):
            t_idx = i
            break
    if t_idx is None:
        out["titulo"] = f"PORTARIA CNMP-PRESI N° {numero} DE {ano}"
        body_lines = lines
    else:
        out["titulo"] = _normalize_titulo(lines[t_idx])
        body_lines = lines[t_idx + 1:]

    # remove a epígrafe repetida nos rodapés das páginas seguintes.
    # Importante: só remove a epígrafe DESTA portaria (linha que COMEÇA com
    # "PORTARIA CNMP-PRESI N° {numero}"), preservando menções a outras
    # portarias que aparecem no meio das frases do corpo.
    body_lines = [ln for ln in body_lines
                  if not _is_footer_epigrafe(ln, numero)]

    # 3) localizar início do preâmbulo ("O/A PRESIDENTE ...")
    pre_start = None
    for i, ln in enumerate(body_lines):
        up = strip_accents_upper(ln)
        if up.startswith("O PRESIDENTE") or up.startswith("A PRESIDENTE"):
            pre_start = i
            break

    # 3a) bloco anterior ao preâmbulo -> ementa e/ou notas
    if pre_start is not None:
        pre_block = body_lines[:pre_start]
        after = body_lines[pre_start:]
    else:
        pre_block = []
        after = body_lines

    # agrupa pre_block em parágrafos (linha em branco não existe mais; usa notas)
    ementa_parts = []
    for ln in pre_block:
        up = strip_accents_upper(ln)
        if any(up.startswith(p) for p in _NOTE_PREFIXES):
            out["notas"].append(ln)
        else:
            ementa_parts.append(ln)
    out["ementa"] = " ".join(ementa_parts).strip()

    # 4) preâmbulo: da linha "O PRESIDENTE" até a que contém "RESOLVE"
    corpo_start = 0
    if after:
        pre_lines = []
        resolve_found = False
        idx = 0
        for idx, ln in enumerate(after):
            pre_lines.append(ln)
            if re.search(r"RESOLVE\s*:?\s*$", ln) or "RESOLVE:" in strip_accents_upper(ln):
                resolve_found = True
                break
        if resolve_found:
            out["preambulo"] = re.sub(r"\s+", " ", " ".join(pre_lines)).strip()
            corpo_start = idx + 1
        else:
            # sem "RESOLVE": trata a 1ª linha como preâmbulo e o resto como corpo
            out["preambulo"] = after[0]
            corpo_start = 1
        rest = after[corpo_start:]
    else:
        rest = []

    # 5) corpo até "Brasília,"; depois assinaturas/cargos
    sig_idx = None
    for i, ln in enumerate(rest):
        if strip_accents_upper(ln).startswith("BRASILIA"):
            sig_idx = i
            break
    corpo_lines = rest[:sig_idx] if sig_idx is not None else rest
    sig_lines = rest[sig_idx + 1:] if sig_idx is not None else []

    # 5a) agrupa corpo em parágrafos
    paragraphs = []
    cur = ""
    for ln in corpo_lines:
        if _PARA_START_RE.match(ln) or _HEADING_RE.match(ln):
            if cur.strip():
                paragraphs.append(cur.strip())
            cur = ln
        else:
            cur = (cur + " " + ln).strip() if cur else ln
    if cur.strip():
        paragraphs.append(cur.strip())
    out["corpo"] = [re.sub(r"\s+", " ", p).strip() for p in paragraphs if p.strip()]

    # 5b) assinaturas (linhas em CAIXA ALTA) e cargos (demais)
    for ln in sig_lines:
        if not ln.strip():
            continue
        letters = [c for c in ln if c.isalpha()]
        is_caps = letters and all(c.isupper() for c in letters)
        if is_caps and len(ln.split()) >= 2:
            out["assinaturas"].append(ln.strip())
        elif out["assinaturas"]:
            out["cargos"].append(ln.strip())
    if not out["assinaturas"]:
        out["assinaturas"] = ["PAULO GUSTAVO GONET BRANCO"]

    return out


# --------------------------------------------------------------------------- #
# Detecção de menções a portarias no corpo do texto
# --------------------------------------------------------------------------- #

# Ex.: "Portaria CNMP-PRESI nº 71/2026", "Portaria CNMP-PRESI n° 71, de 3 de ... de 2026",
#      "Portaria CNMP-CN nº 95, de 19 de maio de 2026"
MENTION_RE = re.compile(
    r"Portaria\s+CNMP-(?P<tipo>PRESI|CN|SG)\s*n[.\u00ba\u00b0o]*\s*(?P<num>\d+)"
    r"(?:\s*/\s*(?P<ano1>\d{4})|.{0,40}?\bde\s+(?P<ano2>\d{4}))",
    re.IGNORECASE,
)


@dataclass
class Segment:
    text: str
    bold: bool = False
    url: Optional[str] = None


def build_segments(text: str,
                   cnmp: Optional["CNMPClient"] = None,
                   bold_phrases: Optional[List[str]] = None) -> List[Segment]:
    """Divide o texto em segmentos (texto simples, negrito, hyperlink).

    - Menções a portarias do CNMP viram segmentos com URL (link).
    - Trechos em ``bold_phrases`` viram segmentos em negrito.
    """
    bold_phrases = bold_phrases or []
    spans = []  # (start, end, kind, payload)

    # 1) menções a portarias -> links
    if cnmp is not None:
        for m in MENTION_RE.finditer(text):
            tipo = (m.group("tipo") or "PRESI").upper()
            num = int(m.group("num"))
            ano = int(m.group("ano1") or m.group("ano2"))
            url = cnmp.find_norma_url(num, ano, tipo)
            if url:
                spans.append((m.start(), m.end(), "link", url))
                log(f"Corpo: menção a 'Portaria CNMP-{tipo} nº {num}/{ano}' "
                    f"-> {url}")

    # 2) trechos em negrito
    for phrase in bold_phrases:
        start = text.find(phrase)
        if start >= 0:
            spans.append((start, start + len(phrase), "bold", None))

    if not spans:
        return [Segment(text=text)]

    # ordena e remove sobreposições (link tem prioridade)
    spans.sort(key=lambda s: (s[0], 0 if s[2] == "link" else 1))
    filtered = []
    last_end = -1
    for s in spans:
        if s[0] >= last_end:
            filtered.append(s)
            last_end = s[1]

    segments = []
    cursor = 0
    for start, end, kind, payload in filtered:
        if start > cursor:
            segments.append(Segment(text=text[cursor:start]))
        if kind == "link":
            segments.append(Segment(text=text[start:end], url=payload))
        elif kind == "bold":
            segments.append(Segment(text=text[start:end], bold=True))
        cursor = end
    if cursor < len(text):
        segments.append(Segment(text=text[cursor:]))
    return segments



# --------------------------------------------------------------------------- #
# Construtor do .docx
# --------------------------------------------------------------------------- #

# Estilos (ids) presentes no modelo institucional
STYLE_CABECALHO = "CabealhoPresi"   # "Cabeçalho Presi"
STYLE_CORPO = "CorpoPresi"          # "Corpo Presi"

PRESIDENTE_BOLD = "PRESIDENTE DO CONSELHO NACIONAL DO MINISTÉRIO PÚBLICO"


def _set(el, tag, **attrs):
    child = OxmlElement(tag)
    for k, v in attrs.items():
        child.set(qn(k), v)
    el.append(child)
    return child


class PortariaDocBuilder:
    """Gera o .docx da portaria a partir do modelo institucional."""

    def __init__(self, template_path: str, cnmp: Optional[CNMPClient] = None):
        self.doc = Document(template_path)
        self.cnmp = cnmp
        self.body = self.doc.element.body
        self.sectPr = self.body.find(qn("w:sectPr"))

    # -- infraestrutura de parágrafos ------------------------------------- #
    def _new_paragraph(self, style_id: Optional[str] = None) -> Paragraph:
        """Cria um <w:p> imediatamente antes do sectPr e devolve o Paragraph."""
        p = OxmlElement("w:p")
        self.sectPr.addprevious(p)
        para = Paragraph(p, self.doc._body)
        if style_id:
            ppr = p.get_or_add_pPr()
            _set(ppr, "w:pStyle", **{"w:val": style_id})
        return para

    def _clear_body(self) -> None:
        for child in list(self.body):
            if child.tag == qn("w:sectPr"):
                continue
            self.body.remove(child)

    # -- propriedades de run ---------------------------------------------- #
    @staticmethod
    def _tnr_rpr(run_el, *, color=COLOR_BLACK, bold=False, link=False):
        """Aplica as propriedades de run (Times New Roman, cor, negrito, link)."""
        rpr = run_el.get_or_add_rPr()
        if link:
            _set(rpr, "w:rStyle", **{"w:val": "Hyperlink"})
        rfonts = _set(rpr, "w:rFonts", **{
            "w:ascii": "Times New Roman", "w:eastAsia": "Times New Roman",
            "w:hAnsi": "Times New Roman", "w:cs": "Times New Roman"})
        if bold:
            _set(rpr, "w:b")
            _set(rpr, "w:bCs")
        if link:
            _set(rpr, "w:color", **{"w:val": COLOR_LINK_BLUE})
            _set(rpr, "w:u", **{"w:val": "single"})
        elif color:
            _set(rpr, "w:color", **{"w:val": color})
        return rpr

    def _add_text_run(self, para: Paragraph, text: str, *, color=COLOR_BLACK,
                      bold=False):
        r = OxmlElement("w:r")
        self._tnr_rpr(r, color=color, bold=bold)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        para._p.append(r)
        return r

    def _add_hyperlink(self, para: Paragraph, text: str, url: str, *,
                       bold=False):
        """Adiciona um hyperlink externo (azul, sublinhado) ao parágrafo."""
        r_id = self.doc.part.relate_to(url, RT.HYPERLINK, is_external=True)
        hyper = OxmlElement("w:hyperlink")
        hyper.set(qn("r:id"), r_id)
        r = OxmlElement("w:r")
        self._tnr_rpr(r, bold=bold, link=True)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        hyper.append(r)
        para._p.append(hyper)
        return hyper

    def _add_segments(self, para: Paragraph, segments: List[Segment]):
        for seg in segments:
            if seg.url:
                self._add_hyperlink(para, seg.text, seg.url, bold=seg.bold)
            else:
                self._add_text_run(para, seg.text, bold=seg.bold)

    # -- parágrafos utilitários ------------------------------------------- #
    def _blank_paragraph(self, first_line: Optional[int] = None):
        """Parágrafo em branco no formato usado pelo modelo (espaçamento 80)."""
        para = self._new_paragraph()
        ppr = para._p.get_or_add_pPr()
        _set(ppr, "w:suppressAutoHyphens", **{"w:val": "0"})
        _set(ppr, "w:autoSpaceDN")
        if first_line is not None:
            _set(ppr, "w:ind", **{"w:firstLine": str(first_line)})
        else:
            _set(ppr, "w:spacing", **{"w:before": "80", "w:after": "80"})
        _set(ppr, "w:jc", **{"w:val": "both"})
        _set(ppr, "w:textAlignment", **{"w:val": "auto"})
        self._add_text_run(para, " ")
        return para

    def _blank_corpo(self):
        para = self._new_paragraph(STYLE_CORPO)
        self._add_text_run(para, " ")
        return para



    # -- título ------------------------------------------------------------ #
    def _add_titulo(self, data: PortariaData):
        para = self._new_paragraph(STYLE_CABECALHO)
        if data.publicado_dou and data.dou_cert_url:
            # título como hyperlink azul para a versão certificada do DOU
            self._add_hyperlink(para, data.titulo, data.dou_cert_url)
        else:
            # publicada no Diário Eletrônico do CNMP -> vermelho, sem link
            ppr = para._p.get_or_add_pPr()
            rpr_mark = OxmlElement("w:rPr")
            _set(rpr_mark, "w:color", **{"w:val": COLOR_CNMP_RED})
            ppr.append(rpr_mark)
            self._add_text_run(para, data.titulo, color=COLOR_CNMP_RED)
        return para

    # -- ementa ------------------------------------------------------------ #
    def _add_ementa(self, texto: str):
        para = self._new_paragraph()
        ppr = para._p.get_or_add_pPr()
        _set(ppr, "w:autoSpaceDN")
        _set(ppr, "w:ind", **{"w:firstLine": "3969"})  # 7 cm
        _set(ppr, "w:jc", **{"w:val": "both"})
        segments = build_segments(texto, self.cnmp)
        self._add_segments(para, segments)
        return para

    # -- notas ("Vide" / "Revogada") -------------------------------------- #
    def _add_nota(self, texto: str):
        """Nota informativa (ex.: '(Revogada pela Portaria ...)') abaixo do título."""
        para = self._new_paragraph()
        ppr = para._p.get_or_add_pPr()
        _set(ppr, "w:jc", **{"w:val": "center"})
        # itálico para diferenciar do texto normativo
        segments = build_segments(texto, self.cnmp)
        for seg in segments:
            if seg.url:
                self._add_hyperlink(para, seg.text, seg.url)
            else:
                r = self._add_text_run(para, seg.text)
                rpr = r.get_or_add_rPr()
                _set(rpr, "w:i")
                _set(rpr, "w:iCs")
        return para

    # -- assinatura -------------------------------------------------------- #
    def _add_assinatura_nome(self, nome: str):
        para = self._new_paragraph()
        ppr = para._p.get_or_add_pPr()
        _set(ppr, "w:suppressAutoHyphens", **{"w:val": "0"})
        _set(ppr, "w:autoSpaceDN")
        _set(ppr, "w:jc", **{"w:val": "center"})
        _set(ppr, "w:textAlignment", **{"w:val": "auto"})
        self._add_text_run(para, nome)
        return para

    def _add_cargo(self, cargo: str):
        para = self._new_paragraph()
        ppr = para._p.get_or_add_pPr()
        _set(ppr, "w:jc", **{"w:val": "center"})
        self._add_text_run(para, cargo)
        return para

    # -- rodapé (epígrafe) ------------------------------------------------- #
    def _update_footer_epigrafe(self, titulo: str):
        """Atualiza a epígrafe exibida no rodapé das páginas seguintes."""
        try:
            section = self.doc.sections[0]
            ftr = section.footer._element
        except Exception:  # noqa: BLE001
            return
        for p in ftr.iter(qn("w:p")):
            texts = "".join(t.text or "" for t in p.iter(qn("w:t")))
            if "PORTARIA" in strip_accents_upper(texts):
                for r in p.findall(qn("w:r")):
                    p.remove(r)
                r = OxmlElement("w:r")
                rpr = r.get_or_add_rPr()
                _set(rpr, "w:rFonts", **{"w:cs": "Times New Roman"})
                _set(rpr, "w:sz", **{"w:val": "16"})
                _set(rpr, "w:szCs", **{"w:val": "16"})
                t = OxmlElement("w:t")
                t.set(qn("xml:space"), "preserve")
                t.text = titulo
                r.append(t)
                p.append(r)
                log("Rodapé: epígrafe atualizada.")
                return

    # -- montagem completa ------------------------------------------------- #
    @staticmethod
    def _normalize_resolve(preambulo: str) -> str:
        return re.sub(
            r"(,?\s*)resolve\s*:\s*$",
            lambda m: (m.group(1) or " ") + "RESOLVE:",
            preambulo.strip(),
            flags=re.IGNORECASE,
        )

    def build(self, data: PortariaData) -> None:
        self._clear_body()

        # 1) título
        self._add_titulo(data)

        # Observação: as anotações "Vide/Revogada" que o portal do CNMP insere
        # após o título são metadados do sistema de normas (não fazem parte do
        # texto oficial da portaria) e, por isso, não são reproduzidas — assim
        # como nos documentos formatados manualmente pela equipe.

        # 2) espaço / ementa
        self._blank_paragraph()
        if data.ementa:
            self._add_ementa(data.ementa)
        else:
            self._blank_paragraph()
        self._blank_paragraph(first_line=1699)

        # 3) preâmbulo (com "PRESIDENTE..." em negrito e "RESOLVE:")
        preambulo = self._normalize_resolve(data.preambulo) if data.preambulo else ""
        if preambulo:
            para = self._new_paragraph(STYLE_CORPO)
            segments = build_segments(preambulo, self.cnmp,
                                      bold_phrases=[PRESIDENTE_BOLD])
            self._add_segments(para, segments)

        # 4) corpo (artigos / parágrafos)
        if data.corpo:
            self._blank_corpo()
            for texto in data.corpo:
                para = self._new_paragraph(STYLE_CORPO)
                segments = build_segments(texto, self.cnmp)
                self._add_segments(para, segments)

        # 5) local e data + assinatura
        self._blank_corpo()
        para = self._new_paragraph(STYLE_CORPO)
        self._add_text_run(para, "Brasília, data da assinatura eletrônica.")
        self._blank_paragraph(first_line=1699)

        if data.assinaturas:
            for i, nome in enumerate(data.assinaturas):
                self._add_assinatura_nome(nome)
                if i < len(data.cargos):
                    self._add_cargo(data.cargos[i])
        # parágrafo final vazio (Normal)
        self._new_paragraph()

        # 6) rodapé
        self._update_footer_epigrafe(data.titulo)

    def save(self, path: str) -> None:
        self.doc.save(path)
        log(f"DOCX gerado: {path}")



# --------------------------------------------------------------------------- #
# Orquestração
# --------------------------------------------------------------------------- #

def gerar_portaria(numero: int, ano: int, *, template_path: str = TEMPLATE_PATH,
                   output_dir: str = OUTPUT_DIR) -> dict:
    """Executa o fluxo completo para a portaria numero/ano.

    Retorna um dicionário com os caminhos gerados e metadados.
    """
    os.makedirs(output_dir, exist_ok=True)
    session = make_session()

    dou = DOUClient(session)
    cnmp = CNMPClient(session)

    data = PortariaData(numero=numero, ano=ano)

    log(f"Iniciando geração da PORTARIA CNMP-PRESI Nº {numero} DE {ano}.")

    # 1) Fonte primária do texto: PDF da portaria no portal do CNMP.
    #    É confiável e existe tanto para portarias do DOU quanto do Diário
    #    Eletrônico do CNMP (o in.gov.br costuma limitar/negar requisições).
    texto_cnmp, cnmp_url = cnmp.fetch_text(numero, ano)
    if texto_cnmp:
        data.cnmp_pdf_url = cnmp_url or ""
        parsed = parse_portaria_text(texto_cnmp, numero, ano)
        data.titulo = parsed["titulo"]
        data.ementa = parsed["ementa"]
        data.notas = parsed["notas"]
        data.preambulo = parsed["preambulo"]
        data.corpo = parsed["corpo"]
        data.assinaturas = parsed["assinaturas"]
        data.cargos = parsed["cargos"]
        data.fonte_texto = "CNMP-PDF"
        log(f"Texto obtido do portal do CNMP "
            f"({len(data.corpo)} parágrafos de corpo).")
    else:
        log("CNMP: PDF da portaria indisponível no portal; "
            "o texto será obtido do DOU, se possível.")

    # 2) Determina DOU x Diário Eletrônico pelo campo "Publicação:" da norma no
    #    portal do CNMP. Essa é a fonte MAIS confiável dessa informação (a busca
    #    textual do in.gov.br é incompleta para 2026).
    pub = cnmp.get_publicacao_info(numero, ano)
    if pub:
        data.norma_url = pub["norma_url"]
        data.publicacao_campo = pub["publicacao"]
        if pub["publicado_dou"]:
            # Publicada no DOU → título azul com link para a versão certificada.
            data.publicado_dou = True
            data.pub_date = pub["pub_date"] or data.pub_date
            secao = pub["secao"] or 2
            # Página no DOU (via índice completo leiturajornal) → link certificado.
            item = dou.find_via_leiturajornal(numero, ano, pub["pub_date"], secao)
            if item and item.get("numberPage") and pub["pub_date"]:
                cert, servlet = dou.build_cert_urls(
                    pub["pub_date"], secao, item["numberPage"])
                data.dou_cert_url = cert
                data.dou_pdf_servlet_url = servlet
                if item.get("urlTitle"):
                    data.dou_page_url = DOUClient.MATERIA_URL.format(
                        url_title=item["urlTitle"])
                log(f"DOU: versão certificada montada (seção {secao}, "
                    f"página {item['numberPage']}, {pub['pub_date']}).")
            else:
                # Fallback 1: busca textual do in.gov.br (às vezes funciona).
                fitem = dou.find(numero, ano)
                if fitem:
                    dou.fetch_materia(fitem, data, only_cert=True)
                # Fallback 2: link para a edição da seção (não a página exata).
                if not data.dou_cert_url and pub["pub_date"]:
                    d = pub["pub_date"].replace("/", "-")
                    data.dou_cert_url = (f"{DOUClient.LEITURA_URL}"
                                         f"?data={d}&secao=do{secao}")
                    data.dou_link_impreciso = True
                    log("DOU: página exata indisponível no momento; o link do "
                        "título aponta para a edição da seção no DOU.")
        else:
            # Publicada apenas no Diário Eletrônico do CNMP → vermelho, sem link.
            data.publicado_dou = False
            log("Norma publicada no Diário Eletrônico do CNMP "
                "(título em vermelho, sem link).")
    else:
        # Norma ainda não catalogada no portal do CNMP (ex.: muito recente).
        # Recorre à busca direta do DOU como melhor esforço.
        log("CNMP: norma não localizada no portal; tentando busca direta no DOU.")
        item = dou.find(numero, ano)
        if item:
            dou.fetch_materia(item, data, only_cert=bool(data.fonte_texto))
            if not data.fonte_texto and data.corpo:
                data.fonte_texto = "DOU-HTML"
        elif dou.server_error:
            data.dou_indisponivel = True
            log("DOU: serviço indisponível — não foi possível confirmar a "
                "publicação no Diário Oficial da União.")
        else:
            log("DOU: portaria não localizada — provavelmente publicada apenas "
                "no Diário Eletrônico do CNMP (título em vermelho, sem link).")

    # 3) baixar o PDF certificado do DOU (somente se publicada no DOU)
    pdf_path = None
    if data.publicado_dou:
        pdf_name = f"{ano}.Portaria-CNMP-PRESI.{numero}-DOU-certificada.pdf"
        pdf_path = os.path.join(output_dir, pdf_name)
        dou.download_certified_pdf(data, pdf_path)

    if not data.titulo:
        # título mínimo caso nada tenha sido encontrado
        data.titulo = f"PORTARIA CNMP-PRESI N° {numero} DE {ano}"

    # 4) gerar o .docx
    builder = PortariaDocBuilder(template_path, cnmp)
    builder.build(data)

    docx_name = f"{ano}.Portaria-CNMP-PRESI.{numero}.docx"
    docx_path = os.path.join(output_dir, docx_name)
    builder.save(docx_path)

    return {
        "numero": numero,
        "ano": ano,
        "titulo": data.titulo,
        "publicado_dou": data.publicado_dou,
        "dou_indisponivel": data.dou_indisponivel,
        "dou_link_impreciso": data.dou_link_impreciso,
        "publicacao_campo": data.publicacao_campo,
        "fonte_texto": data.fonte_texto,
        "dou_page_url": data.dou_page_url,
        "dou_cert_url": data.dou_cert_url,
        "dou_pdf_servlet_url": data.dou_pdf_servlet_url,
        "cnmp_pdf_url": data.cnmp_pdf_url,
        "norma_url": data.norma_url,
        "pdf_path": pdf_path,
        "docx_path": docx_path,
        "n_paragrafos_corpo": len(data.corpo),
        "assinaturas": data.assinaturas,
    }


def _parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Gera Portarias CNMP-PRESI formatadas a partir do DOU.")
    parser.add_argument("numero", nargs="?", type=int, help="Número da portaria")
    parser.add_argument("ano", nargs="?", type=int, help="Ano da portaria")
    parser.add_argument("--numero", dest="numero_opt", type=int)
    parser.add_argument("--ano", dest="ano_opt", type=int)
    parser.add_argument("--template", default=TEMPLATE_PATH)
    parser.add_argument("--output", default=OUTPUT_DIR)
    args = parser.parse_args(argv)
    numero = args.numero_opt or args.numero
    ano = args.ano_opt or args.ano
    if not numero or not ano:
        parser.error("Informe o número e o ano da portaria. Ex.: python3 "
                     "portaria_formatter.py 164 2026")
    return numero, ano, args.template, args.output


def main(argv=None):
    numero, ano, template, output = _parse_args(argv)
    result = gerar_portaria(numero, ano, template_path=template, output_dir=output)
    print("\n" + "=" * 70)
    print("RESUMO")
    print("=" * 70)
    for k, v in result.items():
        print(f"  {k}: {v}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
