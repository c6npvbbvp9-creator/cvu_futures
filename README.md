# CME Futures Database — Automação Diária

Coleta automática dos **preços de fechamento (settlement prices)** de todos os contratos futuros disponíveis no CME Group para as seguintes commodities energéticas:

| Aba no Excel | Commodity                           | Exchange | Ticker CME |
|--------------|-------------------------------------|----------|------------|
| HH           | Henry Hub Natural Gas               | NYMEX    | NG         |
| Brent        | Brent Crude Oil                     | NYMEX    | BZ         |
| NBP          | UK NBP Natural Gas Calendar Month   | NYMEX    | UKG        |
| JKM          | LNG Japan Korea Marker (Platts)     | NYMEX    | JKM        |
| TTF          | Dutch TTF Natural Gas (Platts ENDEX)| NYMEX    | TTF        |
| Coal_API2    | Coal API2 CIF ARA (Argus/McCloskey) | NYMEX    | MTF        |

---

## Como funciona

1. O **GitHub Actions** roda automaticamente de **segunda a sexta às 18h (Brasília)** / 21h UTC.
2. O script acessa o endpoint público de settlements do CME, que retorna **todos os contratos ativos** (inclusive os de 2038).
3. Os preços são salvos no arquivo `data/CME_Futures_Database.xlsx`, com **uma aba por commodity**.
4. O arquivo Excel atualizado é commitado automaticamente de volta no repositório.
5. Para o **TTF** (menor liquidez no CME), contratos sem preço usam fallback: **média dos últimos 12 meses** do histórico (metodologia de contingência do GT CVU Estrutural).

---

## Estrutura do projeto

```
├── main.py                            # Ponto de entrada
├── scraper.py                         # Coleta via API de settlements do CME
├── update_excel.py                    # Salva no Excel com formatação
├── debug_api.py                       # Diagnóstico: inspeciona a API localmente
├── requirements.txt                   # Dependências Python
├── data/
│   └── CME_Futures_Database.xlsx     # Database gerado automaticamente
└── .github/
    └── workflows/
        └── daily_update.yml           # GitHub Actions — execução diária
```

---

## Estrutura do Excel

Cada aba contém as colunas:

| Data de Coleta | Contrato | Data de Vencimento | Preço de Fechamento (USD) | Origem |
|----------------|----------|--------------------|--------------------------|--------|
| 2026-06-19     | JUL 26   | 2026-07-28         | 3.245                    | CME    |
| 2026-06-19     | AUG 26   | 2026-08-27         | 3.310                    | CME    |
| ...            | ...      | ...                | ...                      | ...    |
| 2026-06-19     | DEC 38   | 2038-12-28         | 4.100                    | CME    |

Linhas com `Origem = CONTINGÊNCIA (média N meses)` ficam destacadas em **amarelo** — aplicável ao TTF quando o CME não retorna preço para um contrato específico.

---

## Configuração inicial

### 1. Crie o repositório no GitHub (privado recomendado)

### 2. Suba todos os arquivos

```bash
git clone https://github.com/SEU_USUARIO/cme-futures-database.git
cd cme-futures-database
# copie todos os arquivos do projeto aqui
git add .
git commit -m "feat: setup inicial"
git push origin main
```

### 3. Ative o GitHub Actions

- Acesse seu repositório → **Actions**
- O workflow **Daily CME Futures Update** aparece automaticamente
- Para rodar imediatamente: **Run workflow** → **Run workflow**

> Nenhuma configuração adicional é necessária. O `GITHUB_TOKEN` é gerado automaticamente pelo GitHub.

---

## Diagnóstico local

Para inspecionar a API do CME e confirmar quantos contratos estão sendo retornados:

```bash
pip install -r requirements.txt
python debug_api.py
```

Isso testa cada endpoint e imprime o número de contratos, campos disponíveis, primeiro e último contrato da curva.

---

## Observações

- **Anti-duplicatas:** se o script rodar mais de uma vez no mesmo dia, não adiciona linhas repetidas.
- **Horário:** 21h UTC = 18h de Brasília no horário de verão americano (março–novembro). No inverno americano, o mercado fecha às 22h UTC — o cron pode ser ajustado para `0 22 * * 1-5` se necessário.
- **Fallback TTF:** ativo desde o primeiro dia, mas a média só é calculada após haver histórico no Excel. Nos primeiros dias sem histórico, contratos sem preço são ignorados.
- **Repositório privado:** recomendado, pois o Excel com dados de mercado ficará versionado no GitHub.
