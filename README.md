# Formatador de Portarias CNMP-PRESI

Ferramenta web para formatação automática de Portarias da Presidência do CNMP em `.docx` padronizado, conforme modelo institucional.

## Funcionalidades

- Busca automática do texto oficial no portal do CNMP
- Identificação da fonte de publicação (DOU ou Diário Eletrônico do CNMP)
- Portarias do DOU: título em **azul** com hiperlink para versão certificada
- Portarias do Diário Eletrônico: título em **vermelho**
- Hiperlinks automáticos para portarias mencionadas no corpo do texto
- Interface web simples — sem necessidade de conhecimento técnico

## Deploy rápido no Render (gratuito)

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy)

1. Acesse [render.com](https://render.com) e crie uma conta gratuita
2. Clique em **New → Web Service → Connect a repository**
3. Selecione este repositório
4. O Render detecta o `render.yaml` automaticamente — clique em **Deploy**

## Uso local

```bash
pip install -r requirements.txt
python app.py
# Acesse http://localhost:3000
```

## Tecnologias

- Python 3.12 · Flask · python-docx · pdfplumber · BeautifulSoup
