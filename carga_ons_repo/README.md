# 📊 Carga Efetiva ONS — Download Automático

Script Python que baixa automaticamente os dados diários de **carga efetiva de energia elétrica** (MWmed) do [Boletim Diário da Operação da ONS](http://sdro.ons.org.br/SDRO/DIARIO/index.htm) e gera planilhas Excel formatadas.

## O que o script faz

1. Detecta automaticamente os **dois últimos meses completos**
2. Faz o download diário via scraping do SDRO (ONS)
3. Extrai a tabela de **MWmed por subsistema** (SE/CO, Sul, Nordeste, Norte e SIN)
4. Gera arquivos Excel com formatação profissional, média mensal e gráfico de linha

## Arquivos gerados

| Arquivo | Conteúdo |
|---|---|
| `Carga_Energia_ONS_AAAA_MM.xlsx` | Um arquivo por mês com dados diários |
| `Carga_Energia_ONS_Consolidado.xlsx` | Série completa + abas mensais + gráfico |

## Como usar

### 1. Instalar dependências

```bash
pip install -r requirements.txt
```

### 2. Rodar o script

```bash
python carga_ons.py
```

O script imprime o progresso no terminal e salva os arquivos na pasta atual.

## Fonte dos dados

- **ONS SDRO**: `https://sdro.ons.org.br/SDRO/DIARIO/YYYY_MM_DD/HTML/15_CargaDiariaSubmercado.html`
- **ONS Dados Abertos** (séries históricas anuais): `https://dados.ons.org.br/dataset/carga-energia`

## Subsistemas

| Sigla | Nome completo |
|---|---|
| SE/CO | Sudeste / Centro-Oeste |
| S | Sul |
| NE | Nordeste |
| N | Norte |
| SIN | Sistema Interligado Nacional |

## Licença

Dados: [Creative Commons Atribuição (CC-BY)](http://www.opendefinition.org/licenses/cc-by) — ONS  
Código: MIT
