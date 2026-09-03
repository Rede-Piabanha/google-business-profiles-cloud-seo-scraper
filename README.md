# Google Business Profiles Cloud SEO Scraper

## Visão geral

`Google Business Profiles Cloud SEO Scraper` é um script em Python para mapear empresas locais a partir do Google Places e gerar uma planilha de prospecção com dados comerciais, presença digital, sinais técnicos de SEO e indicadores aproximados de visibilidade de marca.

A ferramenta consulta empresas por segmento e cidade, coleta dados do perfil público, verifica site próprio, telefone, WhatsApp, redes sociais, blog, AMP, PageSpeed Insights, indexação aproximada, busca de marca e entidade no Knowledge Graph.

O objetivo é apoiar análises comerciais e priorização de leads locais com base em sinais objetivos de presença digital e oportunidade de melhoria.

## Funcionalidades

- Busca empresas locais pela Google Places API usando termo-base e cidade.
- Coleta nome, endereço, telefone, site, nota, avaliações, status, tipos de empresa e link do Google Maps.
- Filtra empresas sem site ou telefone.
- Ignora sites classificados como redes sociais ou aplicativos de delivery.
- Analisa o site informado no perfil da empresa.
- Verifica HTTPS, WhatsApp no site, blog, AMP e possíveis erros de acesso.
- Detecta links para Facebook, Instagram, YouTube, TikTok, LinkedIn, Kwai, Messenger, Telegram, Pinterest e X/Twitter.
- Consulta PageSpeed Insights em mobile para performance, SEO, LCP, CLS, INP e categoria de experiência.
- Usa Google Programmable Search Engine para estimar páginas indexadas por domínio.
- Verifica a posição aproximada do site em busca de marca com nome da empresa e cidade.
- Consulta Knowledge Graph para identificar entidade, descrição e tipos associados.
- Calcula scores de oportunidade local, SEO técnico, visibilidade de marca e prioridade geral do lead.
- Gera planilha XLSX com links clicáveis.
- Usa caches locais em JSON para reduzir consultas repetidas e consumo de APIs.

## Quando usar

Use este script para levantar empresas locais com presença digital ativa e sinais de oportunidade comercial, especialmente em prospecção de SEO, desenvolvimento de sites, presença local, conteúdo, performance e consultoria digital.

A ferramenta foi pensada para triagem estratégica. Ela não substitui auditoria técnica completa, análise manual de negócio, validação comercial, diagnóstico jurídico de uso de dados ou conferência individual dos resultados.

## Pré-requisitos

- Python 3.10 ou superior.
- Projeto no Google Cloud com as APIs necessárias habilitadas.
- Chave de API do Google Cloud.
- Programmable Search Engine CX, caso as análises de indexação e busca de marca sejam usadas.
- Dependências listadas em `requirements.txt`.

Instale as dependências com:

```sh
python -m pip install -r requirements.txt
```

## APIs utilizadas

Conforme os recursos habilitados, o script pode usar:

- Places API.
- PageSpeed Insights API.
- Custom Search API.
- Knowledge Graph Search API.

O uso dessas APIs pode consumir cotas e gerar custos no projeto Google Cloud. Revise limites, faturamento, restrições de chave e permissões antes de coletas amplas.

## Configuração do ambiente

Copie o arquivo de exemplo:

```sh
cp .env.example .env
```

No Windows, você pode simplesmente duplicar `.env.example` e renomear a cópia para `.env`.

Preencha apenas o `.env` local:

```env
GOOGLE_CLOUD_API_KEY=sua_chave_google_cloud
GOOGLE_PROGRAMMABLE_SEARCH_ENGINE_CX=seu_cx_do_programmable_search_engine
USE_PLACES_API=true
USE_PAGESPEED_INSIGHTS_API=true
USE_PROGRAMMABLE_SEARCH_ENGINE_API=true
USE_KNOWLEDGE_GRAPH_API=true
```

As variáveis `USE_*` aceitam valores como `true`, `false`, `1`, `0`, `sim` e `não`.

O `.gitignore` exclui `.env`, planilhas geradas e caches locais. Nunca publique uma chave real de API, mesmo que o arquivo seja removido posteriormente do commit mais recente.

## Configuração do script

Abra o arquivo `google-business-profiles-cloud-seo-scraper.py` e ajuste as variáveis principais:

```python
BASE_QUERY = "local search"
CITIES = [
    "City, ST",
]
OUTPUT_XLSX = "google-business-profiles.xlsx"
```

Exemplo:

```python
BASE_QUERY = "escola particular"
CITIES = [
    "Petrópolis, RJ",
    "Teresópolis, RJ",
]
```

Também é possível controlar o consumo das APIs complementares:

```python
MAX_PAGESPEED_REQUESTS = 200
MAX_CSE_REQUESTS = 200
MAX_KG_REQUESTS = 200
```

## Execução

No terminal, execute:

```sh
python google-business-profiles-cloud-seo-scraper.py
```

Durante a execução, o script registra o andamento da coleta, das análises complementares, do cálculo dos scores e da gravação dos caches.

## Arquivos gerados

- `google-business-profiles.xlsx`: planilha final com empresas, dados de contato, site, redes sociais, métricas técnicas, sinais de SEO e scores de prioridade.
- `known_places.json`: cache de empresas já processadas por Place ID.
- `pagespeed_cache.json`: cache de consultas ao PageSpeed Insights por domínio.
- `cse_site_cache.json`: cache de estimativas de indexação por domínio.
- `cse_brand_cache.json`: cache de consultas de busca de marca.
- `kg_cache.json`: cache de consultas ao Knowledge Graph.

Esses arquivos são ignorados pelo Git porque podem conter dados coletados, resultados de APIs ou informações de prospecção que não devem ser publicados por acidente.

## Principais colunas da planilha

A planilha inclui, entre outros campos:

- Nome, cidade, estado e endereço.
- Site, telefone e link de WhatsApp.
- Nota, número de avaliações e status no Google.
- HTTPS, blog, AMP e erro no site.
- Links sociais encontrados.
- Performance mobile, SEO mobile, LCP, CLS e INP.
- Categoria de experiência.
- Posição aproximada na busca de marca.
- Páginas indexadas aproximadas.
- Entidade no Knowledge Graph.
- Scores local, técnico, visibilidade de marca e prioridade do lead.

## Testes e CI

O repositório inclui testes unitários para funções utilitárias e um workflow de GitHub Actions que executa compilação e testes em Python 3.10, 3.11, 3.12 e 3.13.

Para validar localmente:

```sh
python -m py_compile google-business-profiles-cloud-seo-scraper.py
python -m unittest discover -s tests -v
```

Os testes não exigem credenciais reais nem fazem chamadas às APIs do Google.

## Segurança, privacidade e uso responsável

Use a ferramenta apenas de forma compatível com os termos das APIs e serviços utilizados e com as regras aplicáveis ao tratamento, armazenamento e uso dos dados coletados.

Não inclua chaves de API, credenciais, dados pessoais, listas de leads, planilhas exportadas ou caches de respostas em issues, pull requests ou commits públicos.

Os dados públicos retornados por APIs ou encontrados em sites podem continuar sujeitos a termos de serviço, direitos de terceiros, regras de privacidade e outras obrigações. A existência pública de um dado não implica autorização irrestrita para qualquer finalidade.

Para reportar vulnerabilidades, consulte `SECURITY.md`.

## Observações

- O script coleta apenas empresas com site próprio e telefone identificados no perfil.
- Empresas cujo site informado seja rede social ou app de delivery são descartadas.
- Empresas sem WhatsApp detectado no site podem ser descartadas quando o site responde sem erro.
- A estimativa de páginas indexadas depende do Google Programmable Search Engine e não deve ser tratada como contagem oficial de indexação.
- A posição de busca de marca é aproximada e limitada aos resultados retornados pela Custom Search API.
- Os scores servem para priorização interna e devem ser revisados antes de qualquer abordagem comercial.
- Este projeto não é afiliado, patrocinado ou endossado pelo Google. Google, Google Cloud, Google Maps e demais marcas citadas pertencem aos seus respectivos titulares.

## Contribuindo

Contribuições são bem-vindas. Consulte `CONTRIBUTING.md` antes de abrir um pull request.

## Licença

Este projeto é distribuído sob a MIT License.

Copyright (c) 2025-2026 Rede Piabanha.

Consulte `LICENSE` para o texto completo.
