# 📆 Gerenciador de Dias Úteis com Python

Este projeto automatiza a atualização de uma planilha Excel que controla os dias úteis de cada mês. É uma solução simples, mas poderosa, para quem trabalha com KPIs que dependem de dias úteis — como produtividade, SLAs e metas operacionais.

## 🎯 Objetivo

Atualizar automaticamente os campos:

- `"Dia útil do mês"`
- `"Dias úteis restantes"`

Com base na data atual e no primeiro dia útil do mês.

## ⚙️ Como Funciona

1. O script lê uma planilha Excel com uma aba chamada `"Dias uteis"`.
2. Verifica a data atual e identifica se é o **primeiro dia útil do mês**.
3. Aplica a lógica de incremento nos campos, dependendo do mês e da posição da data.
4. Salva a planilha com os novos valores atualizados.

## 🛠️ Bibliotecas Utilizadas

```python
from datetime import datetime
import locale
import pandas as pd
from pandas.tseries.offsets import BDay
