import requests
import pandas as pd

# =========================================
# CONFIGURAÇÕES
# =========================================
URL_BASE = "https://www.comdinheiro.com.br/Clientes/API/EndPoint001.php"

USERNAME = "douro.capital"
PASSWORD = "Douro@2022"
balanco = "31122025"
balanco_anterior = "31122024"
data_final_carteiras = 31032026

# =========================================
# FUNÇÃO PRINCIPAL (EXTRAÇÃO + TRATAMENTO)
# =========================================
def extrair_e_tratar(payload_url_encoded, tab_name, nome_df=None):

    body = (
        f"username={USERNAME}"
        f"&password={PASSWORD}"
        f"&URL={payload_url_encoded}"
        f"&format=json3"
    )

    response = requests.post(
        URL_BASE,
        params={"code": "import_data"},
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data=body
    )

    print(f"\n--- DEBUG {nome_df} ---")
    print("Status:", response.status_code)
    print("Resposta (preview):", response.text[:300])

    if not response.text.strip():
        raise ValueError(f"{nome_df} retornou resposta vazia")

    try:
        data = response.json()
    except Exception:
        raise ValueError(f"{nome_df} NÃO retornou JSON válido")

    tables = data["tables"]

    # Detecta aviso de cota/erro da API antes de tentar parsear
    warnings = data.get("warnings") or data.get("errors") or {}
    if isinstance(warnings, dict):
        for msg in warnings.values():
            if msg:
                print(f"  [AVISO API] {str(msg)[:300]}")

    # A API pode retornar "tables" como dict {"tab0": {...}} ou como lista [{...}]
    if isinstance(tables, dict):
        tab = tables[tab_name]
    elif isinstance(tables, list):
        if len(tables) == 0:
            aviso = next(iter(warnings.values()), "sem mensagem de erro") if warnings else "sem mensagem de erro"
            raise ValueError(
                f"{nome_df}: a API retornou 'tables' vazio.\n"
                f"  Motivo reportado pela API: {str(aviso)[:300]}"
            )
        idx = int(tab_name.replace("tab", "")) if tab_name.startswith("tab") else 0
        tab = tables[idx]
    else:
        raise ValueError(f"{nome_df}: formato inesperado em 'tables' (tipo: {type(tables).__name__})")

    # "tab" pode ser dict {"0": {...}, "1": {...}, ...} ou lista de dicts
    if isinstance(tab, dict):
        registros = list(tab.values())
    elif isinstance(tab, list):
        registros = tab
    else:
        raise ValueError(f"{nome_df}: formato inesperado na tabela '{tab_name}' (tipo: {type(tab).__name__})")
    df = pd.DataFrame(registros)

    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    df.columns = [str(col).strip().lower() for col in df.columns]

    if nome_df:
        print(f"{nome_df} carregado: {df.shape[0]} linhas")

    return df

# =========================================
# PAYLOADS (MANTENDO ENCODING ORIGINAL)
# =========================================

payload_ranking_cp = (
    "HistoricoIndicadoresFundamentalistas001.php%3F%26data_ini%3D"
    + balanco_anterior + "%26data_fim%3D"
    + balanco +
    "%26trailing%3D12%26conv%3DMIXED%26moeda%3DMOEDA_ORIGINAL%26c_c%3Dconsolidado_preferencialmente%26m_m%3D1000000%26n_c%3D2%26f_v%3D1%26papel%3DRDOR3%2BEQTL3%2BRAIZ4%2BKLBN11%2BJSLG3%2BENGI3%2BAXIA3%2BBEEF3%2BALOS3%2BPETR4%2BRAIL3%2BISAE4%2BEGIE3%2BENEV3%2BVAMO3%2BMDIA3%2BJBSS3%2BRENT3%2BMOVI3%2BSUZB3%2BTAEE11%2BHAPV3%2BSMTO3%2BVBBR3%2BPRIO3%2BBRKM3%2BALUP3%2BCSMG3%2BDASA3%2BASAI3%2BBHIA3%2BVALE3%2BBRKP18%2BUGPA3%2BMOTV3%2BAURE3%2BCMIG3%2BSRNA3%2BCPFE3%2BCPLE3%2BCSAN3%2BCSNA3%2BDIRR3%2BSIMH3%2BECOR3%2BSAPR3%2BGMAT3%2BHBSA3%2BSBSP3%2BIGTA3%2BRANI3%2BLIGT3%2BMBRF3%2BMRVE3%2BONCO3%2BNEOE3%2BCURY3%2BDXCO3%2BCAML3%2BBTEL23%2BMULT3%2BBISA3%2BJHSF3%2BMRSA3B%2B02339%2B02718%2B900685%2B904540%2B915264%2B909101%26indic%3DNOME_EMPRESA%2Bdivida_liquida~ebitda%2Bebitda~DESPESA_FINANCEIRA%2BLC%2BML%2BROIC%2BCAIXA_E_EQUIVALENTES%2BINVESTIMENTOS%2BDEBENTURES_CP%2BEMP_FIN_CP%2BDIVIDA_BRUTA%2BPL%2BTITULOS_PARA_NEGOCIACAO_CP%2BDISPONIBILIDADES%26periodicidade%3Dano%26graf_tab%3Dtabela_vi%26desloc_data_analise%3D1%26flag_transpor%3D0%26c_d%3Dd%26enviar_email%3D0%26enviar_email_log%3D0%26cabecalho_excel%3Dmodo1%26relat_alias_automatico%3Dcmd_alias_01&format=json3)"
)
# (mantive truncado só pra não poluir visual — pode manter o seu completo)


# =========================================
# EXTRAÇÕES
# =========================================
dados_ranking_cp = extrair_e_tratar(payload_ranking_cp, "tab0", "dados_ranking_cp")

# =========================================
# DEBUG
# =========================================

print("\nCarteiras Crédito:")
print(dados_ranking_cp.head())


