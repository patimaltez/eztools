import pandas as pd
from datetime import timedelta
import re
# executado 08/06/2026
# ========================
# Parâmetros e caminhos baixar o relatório da eztools de 1/01/2024 até o dia atual e substituir o arquivo de entrada
# ========================
caminho_entrada  = r"M:\Business Intelligence\FONTES DE DADOS\eztools DP\eztools_dp_2025.xlsx"
caminho_saida    = r"M:\Business Intelligence\FONTES DE DADOS\eztools DP\Arquivo tratado\eztools_dp.xlsx"
caminho_feriados = r"M:\Business Intelligence\FONTES DE DADOS\Feriados.xlsx"

# Janela de datas (INCLUSIVA) para filtrar "Criado Em"
janela_inicial_str = "01/01/2023"  # dd/mm/yyyy
janela_final_str   = "31/12/2026"  # dd/mm/yyyy

print("[1/12] Iniciando execução...")
print(f"[2/12] Caminhos:\n  Entrada: {caminho_entrada}\n  Saída:   {caminho_saida}\n  Feriados:{caminho_feriados}")

# ========================
# Helpers de data (defina antes de usar)
# ========================
from datetime import datetime, timedelta, time as dtime

# Janela de trabalho (ajuste se necessário)
HORA_INICIO = dtime(8, 0, 0)   # 08:00
HORA_FIM    = dtime(17, 30, 0)  # 17:30

def segundos_uteis_entre(dt_ini: pd.Timestamp,
                         dt_fim: pd.Timestamp,
                         feriados_norm: set[pd.Timestamp]) -> float:
    """
    Soma os segundos dentro do intervalo [dt_ini, dt_fim] que caem
    em dias úteis (seg-sex), fora dos feriados e dentro de HORA_INICIO..HORA_FIM.
    """
    if pd.isna(dt_ini) or pd.isna(dt_fim) or dt_fim <= dt_ini:
        return 0.0

    total = 0.0
    dia = dt_ini.normalize()
    ultimo_dia = dt_fim.normalize()

    while dia <= ultimo_dia:
        # pular fins de semana e feriados
        if (dia.weekday() < 5) and (dia not in feriados_norm):
            j_ini = datetime.combine(dia.to_pydatetime().date(), HORA_INICIO)
            j_fim = datetime.combine(dia.to_pydatetime().date(), HORA_FIM)

            # janela do dia corrente recortada pelo intervalo pedido
            start_eff = max(dt_ini.to_pydatetime(), j_ini)
            end_eff   = min(dt_fim.to_pydatetime(), j_fim)

            if end_eff > start_eff:
                total += (end_eff - start_eff).total_seconds()

        dia += pd.Timedelta(days=1)

    return total


def parse_datetime_misto(txt):
    """Parse robusto: detecta ISO (YYYY-MM-DD ...) vs BR (DD/MM/YYYY ...)."""
    if pd.isna(txt):
        return pd.NaT
    s = str(txt).strip()
    if not s:
        return pd.NaT
    # ISO 8601-like -> NÃO usar dayfirst
    if re.search(r"\b\d{4}-\d{2}-\d{2}\b", s):
        return pd.to_datetime(s, errors="coerce", dayfirst=False)
    # BR/PT-BR -> usar dayfirst
    if re.search(r"\b\d{2}/\d{2}/\d{4}\b", s):
        return pd.to_datetime(s, errors="coerce", dayfirst=True)
    # fallback
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def primeira_data_em_DatasRespostas(celula: str):
    """Retorna a 1ª data/hora válida em 'DatasRespostas' (ou NaT), tratando formatos mistos."""
    if pd.isna(celula):
        return pd.NaT
    for p in (x.strip() for x in str(celula).split(",")):
        if not p:
            continue
        dt = parse_datetime_misto(p)
        if pd.notna(dt):
            return dt
    return pd.NaT

def classificar_quando_foi_respondido(dt_resp, lista_feriados_norm):
    """Classifica 'Qnd foi respondido' usando o MESMO timestamp calculado (dt_resp)."""
    dt = parse_datetime_misto(dt_resp)
    if pd.isna(dt):
        return "Sem resposta"
    if (dt.weekday() >= 5) or (dt.normalize() in lista_feriados_norm):
        return "Após horario comercial"
    # 08:00:00 a 17:30:00
    inicio = datetime.combine(dt.date(), HORA_INICIO)
    fim    = datetime.combine(dt.date(), HORA_FIM)

    if inicio <= dt <= fim:
        return "Dentro do horario comercial"
    return "Após horario comercial"

# ========================
# Leitura de bases
# ========================
print("[3/12] Lendo planilha principal...")
df = pd.read_excel(caminho_entrada, dtype=str)
print(f"[3/12] Linhas lidas: {len(df)}")

print("[4/12] Lendo feriados...")
df_feriados = pd.read_excel(caminho_feriados, sheet_name="Feriados")
lista_feriados = pd.to_datetime(df_feriados["FERIADOS"], errors="coerce").dropna().dt.normalize().tolist()
feriados_norm = set(lista_feriados)  # <<< use set para O(1)
print(f"[4/12] Feriados carregados: {len(feriados_norm)}")

# ========================
# Filtro por Status
# ========================
print("[5/12] Filtrando Status == 'Finalizado'...")
df_filtrado = df[df["Status"].eq("Finalizado")].copy()
print(f"[5/12] Linhas após Status=Finalizado: {len(df_filtrado)}")

# ========================
# Conversões de data p/ filtro
# ========================
print("[6/12] Convertendo 'Criado Em' para datetime (apenas para filtrar)...")
df_filtrado["_CriadoEm_dt"] = df_filtrado["Criado Em"].apply(parse_datetime_misto)

# Janela inclusiva: início 00:00:00 do dia inicial, fim 23:59:59.999999 do dia final
janela_inicial = pd.to_datetime(janela_inicial_str, dayfirst=True)
janela_final   = pd.to_datetime(janela_final_str,   dayfirst=True) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)

print(f"[6/12] Aplicando janela inclusiva: {janela_inicial_str} 00:00:00  até  {janela_final_str} 23:59:59")
mask_janela = (df_filtrado["_CriadoEm_dt"] >= janela_inicial) & (df_filtrado["_CriadoEm_dt"] <= janela_final)
df_filtrado = df_filtrado[mask_janela].copy()
print(f"[6/12] Linhas após janela de datas: {len(df_filtrado)}")

# ========================
# Cálculo do Tempo de Resposta (dias decimais)
# ========================
print("[7/12] Calculando 'Tempo de resposta' (dias decimais, úteis 08:00–18:00, sem fds/feriados + regra do almoço)...")

cri = df_filtrado["Criado Em"].apply(parse_datetime_misto)
dr1 = df_filtrado["DatasRespostas"].apply(primeira_data_em_DatasRespostas)
ult = df_filtrado["Últ. Atualização"].apply(parse_datetime_misto)

resp = dr1.where(dr1.notna(), ult)

seg_uteis_series = []
lunch_deducted = 0
lunch_skipped_short = 0

for i in range(len(df_filtrado)):
    dt_ini = cri.iloc[i]
    dt_fim = resp.iloc[i]

    seg = segundos_uteis_entre(dt_ini, dt_fim, feriados_norm)   # usa o set (O(1) membership)

    # --- Regra do almoço ---
    # Se Criado Em ∈ [11:00, 15:00) em dia útil não feriado:
    #   - se tempo útil > 1h => desconta 1h
    #   - se tempo útil menor ou igual 1h => não desconta nada
    if pd.notna(dt_ini):
        is_weekday = dt_ini.weekday() < 5
        not_holiday = (dt_ini.normalize() not in feriados_norm)
        in_lunch_window = dtime(11, 0) <= dt_ini.time() < dtime(15, 0)

        if is_weekday and not_holiday and in_lunch_window and seg is not None:
            if seg > 3600.0:
                seg -= 3600.0
                lunch_deducted += 1
            else:
                lunch_skipped_short += 1

    seg_uteis_series.append(seg if (seg is not None and seg > 0) else None)

# Série de segundos úteis -> dias decimais (base 24h)
seg_uteis = pd.Series(seg_uteis_series, index=df_filtrado.index)
dias_dec = (seg_uteis / 86400.0).astype(float).round(4)

df_filtrado["_TimestampResposta"] = resp
df_filtrado["Tempo de resposta"] = dias_dec

print(f"[7/12] Cálculo concluído. Descontos de almoço aplicados: {lunch_deducted} | Casos <= 1h (sem desconto): {lunch_skipped_short}")


# ========================
# Classificação "Qnd foi respondido"
# ========================
print("[8/12] Classificando 'Qnd foi respondido' com base no timestamp calculado...")
df_filtrado["Qnd foi respondido"] = df_filtrado["_TimestampResposta"].apply(
    lambda x: classificar_quando_foi_respondido(x, feriados_norm)
)

# ========================
# Organizar colunas
# ========================
print("[9/12] Reorganizando colunas para saída...")
colunas_para_excluir = [
    "Unnamed: 0", "Solicitação", "Cobrança", "Histórico",
    "Forma de Rateio", "Emissão", "StatusConluído", "T",
    "_CriadoEm_dt", "_TimestampResposta"
]
df_saida = df_filtrado.drop(columns=[c for c in colunas_para_excluir if c in df_filtrado.columns]).copy()

colunas_finais = [
    "ID","Status","Urgente","Condominio","Assunto","Atividade","Responsável",
    "Criado Por","Criado Em","Últ. Atualização",
    "Tempo de resposta","Qnd foi respondido","DatasRespostas","Usuário Atendente","Duplicar"
]
df_saida = df_saida[[c for c in colunas_finais if c in df_saida.columns]]

# ========================
# Converter datas p/ Excel (usar o MESMO parser)
# ========================
print("[10/12] Ajustando tipos datetime para manter formatação no Excel...")
for col in ["Criado Em", "Últ. Atualização"]:
    if col in df_saida.columns:
        df_saida[col] = df_saida[col].apply(parse_datetime_misto)

# ========================
# Exportar Excel
# ========================
print("[11/12] Salvando arquivo Excel...")
with pd.ExcelWriter(caminho_saida, engine="openpyxl", datetime_format="dd/mm/yyyy hh:mm:ss") as writer:
    df_saida.to_excel(writer, index=False)

print(f"[12/12] Concluído. Arquivo salvo em: {caminho_saida}")
print(f"Total de linhas exportadas: {len(df_saida)}")
