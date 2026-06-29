# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Cockpit de Revenue Marketing Inbound — dashboard HTML single-file conectado a um `dados.json` gerado a partir de planilha xlsx com dados mockados.

## Commands

```bash
# Gerar dados.json a partir do xlsx
python gerar_json.py

# Dependências necessárias (Python 3.x)
python -m pip install openpyxl pandas
```

## Architecture

| Arquivo | Papel |
|---|---|
| `dash_teste_tabelas (1).xlsx` | Fonte de dados: 9 tabelas (fato + dim) |
| `gerar_json.py` | ETL: xlsx → `dados.json` |
| `dados.json` | Contrato de dados consumido pelo cockpit |
| `cockpit-inbound-teste.html` | Dashboard HTML (atualmente com dados mockados inline) |
| `spec_anonimizada.docx` | Especificação técnica completa |

### Tabelas do xlsx

**Fato:**
- `fato_funil_diario` — espinha dorsal; granularidade: data × motion × carteira × fonte_mql × etapa
- `fato_paginas_lm` — pageviews e form_fills das páginas MOFU/BOFU (Levantada de Mão)
- `fato_paginas_passivos` — pageviews e form_fills dos materiais TOFU (Passivos)
- `fato_seo_kws_semanal` — posicionamento orgânico por keyword (semanal)
- `fato_aeo_prompts_semanal` — presença em LLMs por prompt monitorado (semanal)
- `fato_midia_paga_diaria` — performance de campanhas (Google Ads, Meta, LinkedIn)
- `fato_midia_criativos` — performance de criativos ativos

**Dimensão:**
- `dim_metas` — metas por mês × motion × etapa × carteira
- `dim_alertas_funil` — limites inferior/superior para alertas automáticos

### Dimensões-chave do funil

- **motion**: `LM` (Levantada de Mão) | `Passivos`
- **carteira**: `Medias` | `Enterprise` | `GrandesContas`
- **fonte_mql**: `Pago` | `Organico` | `Misto`
- **etapas LM**: FormFill → ContatoT → ContatoQ → MQL → MQLcPerfil → SQL → Oportunidade
- **etapas Passivos**: FormFill → ContatoMKT → MQLPassivo → SQL → Oportunidade

### Estrutura do dados.json

```
metadata              — período de referência e timestamp
funil_principal
  hero                — KPIs agregados do mês de referência
  alertas             — disparados quando taxa/volume sai da faixa dim_alertas_funil
  funil_lm            — volume por etapa (motion=LM)
  funil_passivos      — volume por etapa (motion=Passivos)
  carteiras           — breakdown por Médias / Enterprise / GrandesContas
  tendencia_12m       — séries mensais de MQL por carteira e fonte
canais_aquisicao
  mofu_bofu           — fato_paginas_lm agregado + top páginas
  tofu                — fato_paginas_passivos + SEO (última semana) + AEO (última semana)
midia_paga            — resumo + por plataforma + top campanhas + top criativos
metas                 — metas do mês de referência (lista plana)
```

### Notas de implementação

- `gerar_json.py` detecta automaticamente o mês de referência (max date de `fato_funil_diario`)
- Deltas MoM ficam `null` quando o xlsx contém apenas um mês de dados
- Alertas são suprimidos quando dentro da faixa normal; `critical` quando abaixo do limite inferior
- A `tendencia_12m` retroage 12 meses a partir do mês de referência; meses sem dados ficam `0`
