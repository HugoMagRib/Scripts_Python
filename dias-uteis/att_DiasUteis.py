# Importações agrupadas
from datetime import datetime
import locale
import pandas as pd
from pandas.tseries.offsets import BDay

# Configurar a localidade
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')


def exec_alteracoes_Diasuteis():
    data_hoje = datetime.now().strftime("%d/%m/%y")
    # Caminho do arquivo Excel e nome da planilha
    file_path = f"H:/Bi/Solution/Base de Dados 2Tech/Dias Uteis.xlsx"  
    sheet_name = "Dias uteis"        
    
    try:
        # Carregar a planilha existente
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Data atual
        hoje = datetime.now().date()
        # current_date_str = hoje.strftime("%Y-%m-%d")  # não será mais necessário

        # Removemos a verificação do log para que o script execute sempre

        # Determinar o primeiro dia útil do mês atual
        primeiro_dia_util = (hoje.replace(day=1) + BDay(0)).date()

        if hoje == primeiro_dia_util:
            if hoje.month == 1:
                mes_alvo = 12
                ano_alvo = hoje.year - 1
            else:
                mes_alvo = hoje.month - 1
                ano_alvo = hoje.year
        else:
            mes_alvo = hoje.month
            ano_alvo = hoje.year

        # Monta o filtro para selecionar a linha a ser atualizada
        filter_condition = (df[' Mês '] == mes_alvo) & (df['Ano '] == ano_alvo)

        # Calcular quantos dias úteis se passaram desde a última execução.
        incremento = 1

        if hoje == primeiro_dia_util and incremento > 0:
            incremento -= 1

        if incremento > 0 and filter_condition.any():
            df.loc[filter_condition, 'Dia util do mês'] += incremento
            df.loc[filter_condition, 'Dias ultes restantes'] -= incremento

        # Salvar a planilha alterada
        df.to_excel(file_path, sheet_name=sheet_name, index=False)

        print(f"Alterações realizadas e execução registrada.{data_hoje}")
    except Exception as e:
        print(f"Erro ao executar: {e}")

if __name__ == "__main__":
    exec_alteracoes_Diasuteis()