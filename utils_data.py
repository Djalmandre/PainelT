import pandas as pd
import numpy as np
import re

HOJE = pd.Timestamp.now().normalize()

# ── Mapeamento REAL das colunas PTM ────
COL_MAP_PTM = {
    "Remessa":                              "REMESSA",
    "Pedido":                               "PEDIDO",
    "Item":                                 "ITEM",
    "NM":                                   "NM",
    "Denominação de Item":                  "DENOMINACAO",
    "Qtde Remessa":                         "QTDE",
    "Unidade de Medida":                    "UNIDADE",
    "Valor NF":                             "VALOR_NF",
    "Dt Remessa":                           "DT_REMESSA",
    "Dt Necessidade":                       "DT_NECESSIDADE",
    "Dt Baixa":                             "DT_BAIXA",
    "Status":                               "STATUS",
    "Local de expedição":                   "LOCAL_EXPEDICAO",
    "Local do recebedor da mercadoria":     "LOCAL_RECEBEDOR",
    "Centro":                               "CENTRO",
    "DSM":                                  "DSM",
    "Sonof":                                "SONOF",
    "Data NF Solicitada":                   "DT_NF_SOLICITADA",
    "LEP Solicitação NF":                   "LEP_SOLIC_NF",
    "Dt NF":                                "DT_NF",
    "NF":                                   "NF",
    "Dt Lep":                               "DT_LEP",
    "LEP":                                  "LEP",
    "Solicitante LEP":                      "SOLICITANTE_LEP",
    "DTM":                                  "DTM",
    "Tipo Transporte":                      "TIPO_TRANSPORTE",
    "Transportadora":                       "TRANSPORTADORA",
    "Prazo da coleta":                      "PRAZO_COLETA",
    "Prazo de entrega":                     "PRAZO_ENTREGA",
    "Dt - Coleta":                          "DT_COLETA",
    "Piorização Transporte":                "PRIORIZACAO",
    "Retorno LOEP":                         "RETORNO_LOEP",
    "Cobrança PTM":                         "COBRANCA_PTM",
    "Parecer medição":                      "PARECER_MEDICAO",
    "Pedido/Item":                          "PEDIDO_ITEM",
    "Power BI":                             "FASE_POWER_BI",
    "Chamado":                              "CHAMADO",
    "Solicitante":                          "SOLICITANTE",
    "Obs - Chamado 1":                      "OBS_CHAMADO",
    "Posição no Depósito":                  "POSICAO_DEPOSITO",
}

DATE_COLS = ["DT_REMESSA","DT_NECESSIDADE","DT_BAIXA","DT_NF","DT_LEP",
             "DT_NF_SOLICITADA","PRAZO_COLETA","PRAZO_ENTREGA","DT_COLETA"]

ID_COLS = ["REMESSA","PEDIDO","ITEM","NM","SONOF","DSM","NF","LEP",
           "RETORNO_LOEP","PEDIDO_ITEM","CENTRO","DTM","CHAMADO"]


def _find_sheet(xl: pd.ExcelFile, name: str) -> str:
    sheets = xl.sheet_names
    if name in sheets:
        return name
    nl = name.strip().lower()
    for s in sheets:
        if s.strip().lower() == nl:
            return s
    for s in sheets:
        if nl in s.strip().lower():
            return s
    raise ValueError(f"Aba '{name}' nao encontrada. Abas: {sheets}")


def load_sheet(path: str, sheet: str) -> pd.DataFrame:
    xl = pd.ExcelFile(path, engine="openpyxl")
    real = _find_sheet(xl, sheet)
    df = xl.parse(real, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _to_date(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def _to_num(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
         .str.replace(r"[\s]", "", regex=True)
         .str.replace(".", "", regex=False)
         .str.replace(",", ".", regex=False)
         .pipe(pd.to_numeric, errors="coerce")
    )


def _clean_id(s: pd.Series) -> pd.Series:
    if not isinstance(s, pd.Series):
        return s
    return s.astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True).fillna("")


def _faixa(dias):
    if pd.isna(dias):
        return "Sem data"
    d = int(dias)
    if d < 0:   return "Futuro"
    if d <= 7:  return "0–7 dias"
    if d <= 15: return "8–15 dias"
    if d <= 30: return "16–30 dias"
    if d <= 60: return "31–60 dias"
    if d <= 90: return "61–90 dias"
    return "> 90 dias"


def standardize_ptm(df: pd.DataFrame) -> pd.DataFrame:
    df = df.loc[:, ~df.columns.duplicated(keep="first")]

    rename = {k: v for k, v in COL_MAP_PTM.items() if k in df.columns}
    df = df.rename(columns=rename)
    df = df.loc[:, ~df.columns.duplicated(keep="first")]

    # Datas
    for c in DATE_COLS:
        if c in df.columns:
            df[c] = _to_date(df[c])

    # Numérico
    for c in ["VALOR_NF","QTDE"]:
        if c in df.columns:
            df[c] = _to_num(df[c])

    # IDs
    for c in ID_COLS:
        if c in df.columns and isinstance(df[c], pd.Series):
            df[c] = _clean_id(df[c])

    # Strings
    str_cols = ["STATUS","FASE_POWER_BI","PRIORIZACAO","TIPO_TRANSPORTE",
                "TRANSPORTADORA","LOCAL_RECEBEDOR","LOCAL_EXPEDICAO",
                "COBRANCA_PTM","DENOMINACAO","CENTRO","SOLICITANTE",
                "OBS_CHAMADO","RETORNO_LOEP","PARECER_MEDICAO"]
    for c in str_cols:
        if c in df.columns and isinstance(df[c], pd.Series):
            df[c] = df[c].astype(str).str.strip().fillna("")
            df[c] = df[c].replace({"nan":"","None":"","NaT":""})

    # ── Derivados ────
    if "DT_REMESSA" in df.columns:
        df["DIAS_ABERTO"] = (HOJE - df["DT_REMESSA"]).dt.days.clip(lower=0)
        df["FAIXA_ABERTO"] = df["DIAS_ABERTO"].apply(_faixa)

    if "DT_NECESSIDADE" in df.columns and "DT_REMESSA" in df.columns:
        df["LEAD_TIME_DIAS"] = (df["DT_NECESSIDADE"] - df["DT_REMESSA"]).dt.days

    if "DT_NECESSIDADE" in df.columns:
        df["ATRASO_DIAS"] = (HOJE - df["DT_NECESSIDADE"]).dt.days

    # Pendente = sem Dt Baixa
    df["PENDENTE"] = df["DT_BAIXA"].isna() if "DT_BAIXA" in df.columns else True

    # Atrasado = pendente E passou da necessidade
    df["FLAG_ATRASADO"] = False
    if "ATRASO_DIAS" in df.columns and "PENDENTE" in df.columns:
        df["FLAG_ATRASADO"] = (df["PENDENTE"] == True) & (df["ATRASO_DIAS"] > 0)

    # Emergencial — busca em Priorização e Tipo Transporte
    df["IS_EMERGENCIA"] = False
    for c in ["PRIORIZACAO","TIPO_TRANSPORTE"]:
        if c in df.columns and isinstance(df[c], pd.Series):
            mask = df[c].astype(str).str.upper().str.contains("EMERG", na=False)
            if mask.any():
                df["IS_EMERGENCIA"] = df["IS_EMERGENCIA"] | mask

    # Flags documentais
    df["TEM_LEP"] = (df["LEP"].astype(str).str.strip() != "") if "LEP" in df.columns else False
    df["TEM_NF"]  = (df["NF"].astype(str).str.strip()  != "") if "NF"  in df.columns else False
    df["TEM_DSM"] = (df["DSM"].astype(str).str.strip()  != "") if "DSM" in df.columns else False

    # Fase numérica para ordenação
    if "FASE_POWER_BI" in df.columns:
        def _fase_num(s):
            m = re.search(r"(\d+)", str(s))
            return int(m.group(1)) if m else 99
        df["FASE_NUM"] = df["FASE_POWER_BI"].apply(_fase_num)

    return df


def standardize_generic(df: pd.DataFrame) -> pd.DataFrame:
    df = df.loc[:, ~df.columns.duplicated(keep="first")]
    df.columns = [str(c).strip() for c in df.columns]
    for c in df.columns:
        if any(k in c.lower() for k in ["data","dt","date"]):
            df[c] = pd.to_datetime(df[c], dayfirst=True, errors="coerce")
    return df
