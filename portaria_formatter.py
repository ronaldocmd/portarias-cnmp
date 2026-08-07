#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
portaria_formatter.py
====

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

# ---- #
# Configurações gerais
# ---- #

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
COLOR_BLACK = "0000"

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



# ---- #
# Estruturas de dados
# ---- #

@dataclass
class PortariaData:
    """Conteúdo extraído de uma portaria."""
    numero: int
    ano: int
    titulo: str = ""                    # epígrafe (ex.: "PORTARIA CNMP-PRESI N° 164 DE 28 DE MAIO DE 2026")
    ementa: str = ""                    # opcional ("Dispõe sobre ...")
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
    pub_date: str = ""                    # dd/mm/aaaa da publicação no DOU
    dou_indisponivel: bool = False         # True se a busca do DOU falhou por erro de servidor
    dou_link_impreciso: bool = False       # link aponta para a edição (não a página exata)
    # CNMP
    cnmp_pdf_url: str = ""                 # URL do PDF da portaria no portal do CNMP
    norma_url: str = ""                    # página da norma no portal do CNMP
    publicacao_campo: str = ""             # texto do campo "Publicação:" da norma
    fonte_texto: str = ""                  # de onde veio o texto: "CNMP-PDF" | "DOU-HTML" | ""


# ---- #
# Cliente do DOU (in.gov.br)
# ---- #

class DOUClient:
    """Busca e extração de portarias publicadas no Diário Oficial da União."""

    SEARCH_URL = "https://www.in.gov.br/consulta/-/buscar/dou"
    MATERIA_URL = "https://www.in.gov.br/web/dou/-/{url_title}"
    SCRIPT_ID = "_br_com_seatecnologia_in_buscadou_BuscaDouPortlet_params"

    def __init__(self, session: Optional[requests.Session] = None):
        self.session = session or make_session()
        self.server_error = False   # True se o in.gov.br falhou por erro de servidor

    # -- busca ---- #
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

    # -- leiturajornal (índice completo do DOU por edição/seção) ---- #
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

    # -- página da matéria ---- #
    def fetch_materia(self, item: dict, data: PortariaData, ,
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
