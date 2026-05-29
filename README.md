# Google Business Profiles Cloud SEO Scraper

## Descrição

Este script, `google-business-profiles-cloud-seo-scraper.py`, foi desenvolvido para coletar empresas locais a partir do Google Places e gerar uma planilha de análise comercial e SEO. A ferramenta cruza dados de presença local, site próprio, WhatsApp, redes sociais, PageSpeed Insights, indexação aproximada no Google, busca de marca e Knowledge Graph.

O objetivo é identificar empresas com presença digital ativa, mas com possíveis oportunidades de melhoria em SEO técnico, conteúdo, visibilidade de marca e estrutura comercial.

## Funcionalidades

- **Buscar empresas locais no Google Places**: Consulta empresas por termo-base e cidade, usando a API do Google Places.
- **Coletar dados do perfil da empresa**: Obtém nome, endereço, telefone, site, nota, número de avaliações, status, tipos de empresa e link do Google Maps.
- **Filtrar oportunidades comerciais**: Prioriza empresas com site próprio e telefone, ignorando sites que são apenas redes sociais ou aplicativos de delivery.
- **Analisar o site da empresa**: Verifica HTTPS, presença de WhatsApp, blog, AMP, links sociais e possíveis erros de acesso.
- **Identificar canais sociais**: Detecta links para Facebook, Instagram, YouTube, TikTok, LinkedIn, Kwai, Messenger, Telegram, Pinterest e X/Twitter.
- **Analisar performance e SEO técnico**: Consulta o PageSpeed Insights para obter Performance mobile, SEO mobile, LCP, CLS, INP e categoria de experiência.
- **Verificar indexação aproximada**: Usa o Google Programmable Search Engine para estimar páginas indexadas por domínio.
- **Analisar busca de marca**: Verifica a posição aproximada do site da empresa em buscas pelo nome da marca e cidade.
- **Consultar o Knowledge Graph**: Identifica se a empresa possui entidade reconhecida no Knowledge Graph e coleta descrição e tipos de entidade.
- **Calcular scores de prioridade**: Gera scores de oportunidade local, SEO técnico, visibilidade de marca e prioridade geral do lead.
- **Gerar planilha Excel**: Exporta os resultados para `google-business-profiles.xlsx`, com links clicáveis para site, Google Maps, WhatsApp e redes sociais.
- **Salvar caches locais**: Mantém arquivos JSON de cache para evitar consultas repetidas e reduzir consumo de APIs.

## Utilidade

Este script é útil para profissionais de SEO, agências digitais, equipes comerciais e consultores que precisam mapear empresas locais com potencial de contratação de serviços digitais.

A planilha gerada ajuda a identificar negócios que já possuem presença online, mas apresentam sinais de oportunidade, como baixa performance técnica, pouca visibilidade de marca, presença fraca no Google, ausência de blog, problemas de HTTPS, baixa reputação local ou estrutura digital incompleta.

## Pré-requisitos

- Python 3.x
- Bibliotecas Python: requests, python-dotenv, openpyxl
- Projeto no Google Cloud com as APIs necessárias habilitadas:
  - Places API
  - PageSpeed Insights API
  - Custom Search API
  - Knowledge Graph Search API
- Chave de API do Google Cloud
- Programmable Search Engine CX, caso deseje usar as análises de indexação e busca de marca

## Como Usar

1. **Configurar o ambiente**: Certifique-se de ter o Python 3.x instalado e instale as bibliotecas necessárias:
   ```sh
   pip install requests python-dotenv openpyxl
   ```

2. **Criar o arquivo `.env`**: Na raiz do projeto, crie um arquivo `.env` com as credenciais e opções de uso:
   ```env
   GOOGLE_CLOUD_API_KEY=sua_chave_google_cloud
   GOOGLE_PROGRAMMABLE_SEARCH_ENGINE_CX=seu_cx_do_programmable_search_engine
   USE_PLACES_API=true
   USE_PAGESPEED_INSIGHTS_API=true
   USE_PROGRAMMABLE_SEARCH_ENGINE_API=true
   USE_KNOWLEDGE_GRAPH_API=true
   ```

3. **Editar o script**: Atualize as variáveis principais no arquivo `google-business-profiles-cloud-seo-scraper.py`:
   - `BASE_QUERY`: termo-base da busca local, como `escola particular`, `restaurante`, `clínica odontológica` ou outro segmento.
   - `CITIES`: lista de cidades e estados que serão pesquisados.
   - `OUTPUT_XLSX`: nome do arquivo final de saída.
   - `MAX_PAGESPEED_REQUESTS`, `MAX_CSE_REQUESTS` e `MAX_KG_REQUESTS`: limites de consultas para controlar consumo de API.

4. **Executar o script**: No terminal ou prompt de comando, execute:
   ```sh
   python google-business-profiles-cloud-seo-scraper.py
   ```

5. **Verificar os resultados**: Ao final da execução, os dados serão salvos em `google-business-profiles.xlsx`.

## Arquivos Gerados

- `google-business-profiles.xlsx`: planilha final com empresas, dados de contato, site, redes sociais, métricas técnicas, sinais de SEO e scores de prioridade.
- `known_places.json`: cache de empresas já processadas, evitando duplicações entre execuções.
- `pagespeed_cache.json`: cache de consultas ao PageSpeed Insights por domínio.
- `cse_site_cache.json`: cache de consultas de indexação aproximada por domínio.
- `cse_brand_cache.json`: cache de consultas de busca de marca.
- `kg_cache.json`: cache de consultas ao Knowledge Graph.

## Observações

- O uso das APIs do Google pode gerar custos ou consumir cotas do projeto no Google Cloud.
- As métricas de indexação e busca de marca são aproximações baseadas no Google Programmable Search Engine.
- O script foi pensado para análise estratégica e prospecção comercial, não para auditoria definitiva de SEO.
- Antes de executar coletas amplas, ajuste os limites de requisições e revise as cotas disponíveis no Google Cloud.
- Não publique o arquivo `.env` nem exponha a chave da API em repositórios públicos.
