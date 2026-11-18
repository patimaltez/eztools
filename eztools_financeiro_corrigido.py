import pandas as pd
from datetime import datetime
# executado em 08/06/2026
# Caminhos dos arquivos, baixar relatório completo sem filtro de data da eztools, salvar eztools_fc com a data antiga no nome e 
# remover da pasta para o bkp
caminho_entrada = r"M:\Business Intelligence\FONTES DE DADOS\eztools financeiro\eztools_financeiro.xlsx"
caminho_saida = r"M:\Business Intelligence\FONTES DE DADOS\eztools financeiro\Arquivo tratado\eztools_fc.xlsx"
caminho_feriados = r"M:\Business Intelligence\FONTES DE DADOS\Feriados.xlsx"

# Leitura da planilha original
df = pd.read_excel(caminho_entrada, dtype=str)

# Carregar feriados
df_feriados = pd.read_excel(caminho_feriados, sheet_name="Feriados")
lista_feriados = pd.to_datetime(df_feriados['FERIADOS'], errors='coerce').dropna().dt.normalize().tolist()

# Filtros de 'Solicitação' e 'Status'
df_filtrado = df[
    df["Solicitação"].isin(["Alteração", "Novo Rateio"]) &
    df["Status"].eq("Concluído")
]

# Conversão temporária da coluna "Criado Em" para datetime
df_filtrado_com_data = df_filtrado.copy()
df_filtrado_com_data["Criado Em"] = pd.to_datetime(df_filtrado_com_data["Criado Em"], errors="coerce", dayfirst=True)

# Aplicar filtro de data
data_inicial = pd.Timestamp("2025-01-01")
data_final = pd.Timestamp("2026-12-31")
df_filtrado_com_data = df_filtrado_com_data[
    (df_filtrado_com_data["Criado Em"] >= data_inicial) &
    (df_filtrado_com_data["Criado Em"] <= data_final)
]

# Índices válidos para manter os dados originais
df_filtrado_final = df_filtrado.loc[df_filtrado_com_data.index].copy()

# Função para extrair última data válida da coluna "StatusPendente"
def obter_data_status_pendente(status_str, criado_em):
    if pd.isna(status_str) or status_str.strip() == "":
        return criado_em
    partes = status_str.split(",")
    for val in reversed(partes):
        val = val.strip()
        try:
            return pd.to_datetime(val, errors="raise")
        except:
            continue
    return criado_em  # fallback se nenhuma data for válida

# Função para calcular tempo útil em dias decimais (minuto a minuto)
def calcular_tempo_util(data_inicial, data_final, feriados):
    if pd.isna(data_inicial) or pd.isna(data_final):
        return None

    total_minutos_uteis = 0
    intervalo_minutos = int((data_final - data_inicial).total_seconds() / 60)

    for i in range(intervalo_minutos):
        instante = data_inicial + pd.Timedelta(minutes=i)
        if instante.weekday() >= 5:
            continue  # sábado ou domingo
        if instante.normalize() in feriados:
            continue  # feriado
        total_minutos_uteis += 1

    return total_minutos_uteis / 1440  # minutos por dia

# Calcular tempo de resposta considerando tempo útil
datas_calculadas = []
for i, row in df_filtrado_final.iterrows():
    criado_em = pd.to_datetime(row["Criado Em"], errors="coerce", dayfirst=True)
    ultima_atualizacao = pd.to_datetime(row["Últ. Atualização"], errors="coerce", dayfirst=True)
    data_base = obter_data_status_pendente(row.get("StatusPendente", ""), criado_em)

    if pd.notna(data_base) and pd.notna(ultima_atualizacao):
        dias = calcular_tempo_util(data_base, ultima_atualizacao, lista_feriados)
    else:
        dias = None

    datas_calculadas.append(round(dias, 2) if dias is not None else "")

# Adicionar nova coluna com o cálculo
df_filtrado_final["Tempo de resposta"] = datas_calculadas

# Criar coluna 'prazo_meta' com base em 'Tempo de resposta'
def classificar_prazo(valor):
    try:
        if valor == "":
            return ""
        elif float(valor) <= 1:
            return "No prazo"
        else:
            return "Fora do prazo"
    except:
        return ""

df_filtrado_final["prazo_meta"] = df_filtrado_final["Tempo de resposta"].apply(classificar_prazo)

# Remover colunas desnecessárias (se existirem)
colunas_para_excluir = [
    "Unnamed: 0", "ID", "Status", "Solicitação", "Cobrança", "Histórico",
    "Forma de Rateio", "Emissão", "Criado Por", "StatusConluído"
]
df_filtrado_final = df_filtrado_final.drop(columns=[col for col in colunas_para_excluir if col in df_filtrado_final.columns])

# Adicionar colunas fixas
df_filtrado_final["Funcionário"] = "JENIFER BARBOSA REIS PINTO"
df_filtrado_final["Sub-Setor"] = "PASTAS"
        
# Reordenar colunas
colunas_finais = [
    "Condominio",
    "Criado Em",
    "Últ. Atualização",
    "Funcionário",
    "Sub-Setor",
    "Tempo de resposta",
    "prazo_meta"
]
df_filtrado_final = df_filtrado_final[[col for col in colunas_finais if col in df_filtrado_final.columns]]

# Reconverter colunas de data para datetime no DataFrame final
df_filtrado_final["Criado Em"] = pd.to_datetime(df_filtrado_final["Criado Em"], errors="coerce", dayfirst=True)
df_filtrado_final["Últ. Atualização"] = pd.to_datetime(df_filtrado_final["Últ. Atualização"], errors="coerce", dayfirst=True)

# Exportar com formatação de data e hora no Excel
with pd.ExcelWriter(caminho_saida, engine='openpyxl', datetime_format="dd/mm/yyyy hh:mm:ss") as writer:
    df_filtrado_final.to_excel(writer, index=False)

print(f"Arquivo salvo com {len(df_filtrado_final)} linhas e datas com hora formatadas.")

df_fc = pd.read_excel(caminho_saida)

# Garantir que datas estejam como datetime (especialmente porque ao ler do Excel podem vir como texto)
df_fc["Criado Em"] = pd.to_datetime(df_fc["Criado Em"], errors="coerce", dayfirst=True)
df_fc["Últ. Atualização"] = pd.to_datetime(df_fc["Últ. Atualização"], errors="coerce", dayfirst=True)

# Ordenar do mais antigo para o mais recente
df_fc = df_fc.sort_values(by=["Criado Em", "Últ. Atualização"], ascending=[True, True]).reset_index(drop=True)

# Inserir IDs após ordenar
df_fc.insert(0, "id_item", range(1, len(df_fc) + 1))

# Remover registros "Fora do prazo":
# - Ano de 2025: julho a dezembro
# - Ano de 2026: julho
mask_fora_do_prazo = (
    df_fc["prazo_meta"]
    .astype(str)
    .str.strip()
    .str.lower()
    .eq("fora do prazo")
)

mask_2025 = (
    df_fc["Criado Em"].dt.year.eq(2025) &
    df_fc["Criado Em"].dt.month.isin([7, 8, 9, 10, 11, 12])
)

mask_2026 = (
    df_fc["Criado Em"].dt.year.eq(2026) &
    df_fc["Criado Em"].dt.month.isin([7])
)

mask_remover = mask_fora_do_prazo & (mask_2025 | mask_2026)

# Identificar registros que serão removidos
df_removidos = df_fc[mask_remover].copy()

print("\n=== REGISTROS REMOVIDOS ===")
for _, row in df_removidos.iterrows():
    print(
        f"id_item={row['id_item']} | "
        f"Condominio={row['Condominio']} | "
        f"Criado Em={row['Criado Em']} | "
        f"Últ. Atualização={row['Últ. Atualização']} | "
        f"Tempo de resposta={row['Tempo de resposta']} | "
        f"prazo_meta={row['prazo_meta']}"
    )

qtd_antes = len(df_fc)

# Remover efetivamente
df_fc = df_fc[~mask_remover].copy()

# --------------------------------------------------
# Excluir registros específicos de rateios excepcionais
# A exclusão não depende do id_item, pois ele pode mudar
# --------------------------------------------------

exclusoes_manuais = [
    ("1521 - Paysandu 23", "2026-01-02 15:49:00", "2026-01-09 09:16:00"),
    ("0306 - Capibaribe", "2026-01-05 08:37:00", "2026-01-09 09:19:00"),
    ("0306 - Capibaribe", "2026-01-05 08:38:00", "2026-01-09 09:19:00"),
    ("0306 - Capibaribe", "2026-01-05 08:39:00", "2026-01-09 09:20:00"),
    ("0306 - Capibaribe", "2026-01-05 08:50:00", "2026-01-09 09:21:00"),
    ("1135 - La Spezia", "2026-01-05 17:28:00", "2026-01-09 09:56:00"),
    ("1595 - Residencial Solar dos Araujos", "2026-01-06 12:26:00", "2026-01-09 09:57:00"),
    ("1521 - Paysandu 23", "2026-01-06 12:44:00", "2026-01-09 10:43:00"),
    ("2600 - Barão de Marajó", "2026-01-06 14:25:00", "2026-01-09 10:45:00"),
    ("1576 - Quinta do Conde", "2026-01-26 10:28:00", "2026-02-09 10:18:00"),
    ("1025 - José Cyntrybaum", "2026-02-02 21:31:00", "2026-02-09 10:19:00"),
    ("1521 - Paysandu 23", "2026-02-04 17:14:00", "2026-02-09 10:24:00"),
]

df_exclusoes = pd.DataFrame(
    exclusoes_manuais,
    columns=["Condominio", "Criado Em", "Últ. Atualização"]
)

df_exclusoes["Condominio"] = df_exclusoes["Condominio"].astype(str).str.strip()
df_exclusoes["Criado Em"] = pd.to_datetime(df_exclusoes["Criado Em"])
df_exclusoes["Últ. Atualização"] = pd.to_datetime(df_exclusoes["Últ. Atualização"])

df_fc["Condominio"] = df_fc["Condominio"].astype(str).str.strip()
df_fc["Criado Em"] = pd.to_datetime(df_fc["Criado Em"], errors="coerce", dayfirst=True)
df_fc["Últ. Atualização"] = pd.to_datetime(df_fc["Últ. Atualização"], errors="coerce", dayfirst=True)

qtd_antes_exclusoes = len(df_fc)

df_fc = df_fc.merge(
    df_exclusoes.assign(excluir_rateio_manual=True),
    on=["Condominio", "Criado Em", "Últ. Atualização"],
    how="left"
)

df_removidos_manuais = df_fc[df_fc["excluir_rateio_manual"].eq(True)].copy()

print("\n=== REGISTROS REMOVIDOS POR EXCEÇÃO MANUAL ===")
for _, row in df_removidos_manuais.iterrows():
    print(
        f"id_item={row.get('id_item', '')} | "
        f"Condominio={row['Condominio']} | "
        f"Criado Em={row['Criado Em']} | "
        f"Últ. Atualização={row['Últ. Atualização']} | "
        f"Tempo de resposta={row.get('Tempo de resposta', '')} | "
        f"prazo_meta={row.get('prazo_meta', '')}"
    )

df_fc = df_fc[df_fc["excluir_rateio_manual"].ne(True)].copy()
df_fc = df_fc.drop(columns=["excluir_rateio_manual"])

qtd_depois_final = len(df_fc)

print(f"Registros removidos por regra de julho a dezembro: {qtd_antes - qtd_antes_exclusoes}")
print(f"Registros removidos por exceção manual: {qtd_antes_exclusoes - qtd_depois_final}")
print(f"Total de registros removidos da eztools_fc: {qtd_antes - qtd_depois_final}")

# Salvar novamente
df_fc.to_excel(caminho_saida, index=False)