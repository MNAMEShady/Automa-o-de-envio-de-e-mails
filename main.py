import pandas as pd
import win32com.client as win32
import datetime
import pytz

assunto = "Vencimentos RF"

def processar_planilha_e_criar_compromissos(caminho_planilha, nome_aba=0):

    try:
        df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)

        print(f"Planilha '{caminho_planilha}' lida com sucesso. Processando {len(df)} compromissos...")

        utc_timezone = pytz.utc

        for index, row in df.iterrows():
            cliente = row["Cliente"]
            nome_do_cliente = row["Nome do Cliente"]
            ativo = row["Ativo"]
            inicio_original = row["Data Prevista do Evento"]
            financeiro_a_liquidar = row["Financeiro a liquidar"]
            assessor = row["Assessor"]

            participantes_obrigatorios = []
            if pd.notna(row["E-mail"]):
                participantes_obrigatorios = [email.strip() for email in
                                              str(row["E-mail"]).split(';') if email.strip()]

            inicio_utc = None

            if pd.isna(inicio_original):
                print(f"Aviso: 'Data Prevista do Evento' está vazia para o registro {index}. Pulando este compromisso.")
                continue

            try:
                temp_datetime = pd.to_datetime(inicio_original)

                if temp_datetime.tz is None:
                    sp_timezone = pytz.timezone('America/Sao_Paulo')
                    aware_datetime_sp = sp_timezone.localize(temp_datetime)
                    inicio_utc = aware_datetime_sp.astimezone(utc_timezone)
                else:
                    # Se já for aware, apenas converta para UTC
                    inicio_utc = temp_datetime.astimezone(utc_timezone)

            except Exception as e:
                print(f"Erro ao processar 'Data Prevista do Evento' para o registro {index}: {inicio_original} - {e}. Pulando este compromisso.")
                continue

            outlook = win32.Dispatch("Outlook.application")
            mail = outlook.CreateItem(1)
            mail.Subject = assunto
            mail.Body = f"{cliente} | {nome_do_cliente} | {ativo} | {inicio_utc.strftime('%Y-%m-%d')} | R$ {financeiro_a_liquidar:,.2f} | {assessor} | {', '.join(participantes_obrigatorios)}"
            mail.Start = inicio_utc # Passa a data/hora em UTC
            mail.ReminderSet = True
            mail.Recipients.Add(participantes_obrigatorios).Type = 1 # gustavo.aquila@tauariinvestimentos.com.br
            mail.Recipients.Add('operacional@tauariinvestimentos.com.br').Type = 2
            mail.Display()
            mail.Send()
            # mail.Save()

    except FileNotFoundError:
        print(f"Erro: O arquivo Excel '{caminho_planilha}' não foi encontrado.")
    except KeyError as e:
        print(f"Erro: Coluna '{e}' não encontrada na planilha. Verifique se os nomes das colunas estão corretos.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")

if __name__ == "__main__":
    caminho_da_minha_planilha = "Vctos RF - Python.xlsx"
    processar_planilha_e_criar_compromissos(caminho_da_minha_planilha, nome_aba=0)