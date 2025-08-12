
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Balanço Neo Bio (Completo)", layout="wide")

# ---- Try to import openpyxl, but keep working without it ----
try:
    from openpyxl import load_workbook
    OPENPYXL_OK = True
except Exception:
    OPENPYXL_OK = False

PRIMARY = "#0B7A75"
BG_APP  = "#F6F7F9"
BG_CARD = "#FFFFFF"

st.markdown(f"""
<style>
.stApp, section.main, div[data-testid='stAppViewContainer'], div.block-container {{
  background: {BG_APP} !important;
}}
.card {{
  background:{BG_CARD}; border:1px solid #eef0f2; border-radius:16px; padding:14px 16px; box-shadow:0 1px 3px rgba(0,0,0,.05);
}}
.grid3 {{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}}
</style>
""", unsafe_allow_html=True)

st.title("Balanço Neo Bio — Aba Completo")

# ------------- Load defaults -------------
DEFAULTS_HARDCODED = {{
    "C4_Cana": 0.0, "C5_K_cana": 0.0, "C6_vazao_vinho": 100.0, "C8_ds": 8.5, "C9_gl": 14.5,
    "C19_vazao": 0.0, "C20_ds": 0.0, "C21_gl": 0.0, "H8_consumo_to_be": 1.65
}}

def read_defaults_from_xlsx(file):
    wb = load_workbook(file, data_only=False)
    ws = wb["Completo"]
    v = lambda r,c: ws.cell(r,c).value
    return {{
        "C4_Cana": v(4,3) or 0.0,
        "C5_K_cana": v(5,3) or 0.0,
        "C6_vazao_vinho": v(6,3) or 0.0,
        "C8_ds": v(8,3) or 0.0,
        "C9_gl": v(9,3) or 0.0,
        "C19_vazao": v(19,3) or 0.0,
        "C20_ds": v(20,3) or 0.0,
        "C21_gl": v(21,3) or 0.0,
        "H8_consumo_to_be": v(8,8) or 0.0,
    }}

defaults = DEFAULTS_HARDCODED.copy()

st.sidebar.header("📄 Planilha (opcional)")
if OPENPYXL_OK:
    up = st.sidebar.file_uploader("Carregar 'Balanço Neo Bio (1).xlsx' (aba Completo)", type=["xlsx"])
    if up is not None:
        try:
            defaults = read_defaults_from_xlsx(up)
            st.sidebar.success("Valores padrão lidos da planilha.")
        except Exception as e:
            st.sidebar.error(f"Não consegui ler a planilha: {e}")
else:
    st.sidebar.warning("openpyxl não está instalado. O app usa valores padrão embutidos.\nAdicione 'openpyxl' no requirements.txt para ler a planilha automaticamente.")

# ---------- Inputs / Outputs (resumo) ----------
def fmt(x, nd=3):
    try:
        return f"{float(x):,.{nd}f}".replace(",","X").replace(".",",").replace("X",".")
    except:
        return str(x)

st.header("1) Dados da Bio")
col1,col2,col3,col4 = st.columns(4)
with col1:
    C4 = float(defaults["C4_Cana"]); st.write(f"**Cana (fixo)**: {fmt(C4)}")
with col2:
    C5 = float(defaults["C5_K_cana"]); st.write(f"**K Cana (fixo)**: {fmt(C5)}")
with col3:
    C6 = st.number_input("Vazão vinho C6 (variar)", min_value=0.0, value=float(defaults["C6_vazao_vinho"]), step=1.0, format="%.3f")
with col4:
    st.empty()
col5,col6 = st.columns(2)
with col5:
    C8 = st.number_input("%Ds C8 (variar)", min_value=0.0, value=float(defaults["C8_ds"]), step=0.1, format="%.3f")
with col6:
    C9 = st.number_input("Conc GL C9 (variar)", min_value=0.0, value=float(defaults["C9_gl"]), step=0.1, format="%.3f")

# Fórmulas
C7  = (C4*C5/C6) if C6 else 0.0
C10 = (-0.244*C9 + 4.564)
C11 = (C6*C9/96.0*C10)
C12 = (C6*C9/96.0)
C13 = (C12*24.0)
C14 = (C12*1.2)
C15 = (C6 - C12*0.789 + C11 - C14)
C16 = (C6 / C8 / C15) if (C8 and C15) else 0.0
C17 = (C5*C4/C15) if C15 else 0.0

st.markdown('<div class="grid3">', unsafe_allow_html=True)
st.markdown(f'<div class="card"><b>K vinho (C7)</b><div style="font-size:1.3rem">{fmt(C7)}</div><div>Fórmula: C4*C5/C6</div></div>', unsafe_allow_html=True)
st.markdown(f'<div class="card"><b>V1 total as is (C11)</b><div style="font-size:1.3rem">{fmt(C11)}</div><div>Fórmula: C6*C9/96*C10</div></div>', unsafe_allow_html=True)
st.markdown(f'<div class="card"><b>Vinhaça (C15)</b><div style="font-size:1.3rem">{fmt(C15)}</div><div>Fórmula: C6 - C12*0.789 + C11 - C14</div></div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

st.header("2) Dados Neo (Volante Neo - Vinho)")
colN1,colN2,colN3 = st.columns(3)
with colN1:
    C19 = st.number_input("Vazão C19 (variar)", min_value=0.0, value=float(defaults["C19_vazao"]), step=1.0, format="%.3f")
with colN2:
    C20 = st.number_input("%Ds C20 (variar)", min_value=0.0, value=float(defaults["C20_ds"]), step=0.1, format="%.3f")
with colN3:
    C21 = st.number_input("Conc GL C21 (variar)", min_value=0.0, value=float(defaults["C21_gl"]), step=0.1, format="%.3f")

st.header("3) Dados da Mistura")
H5 = C19 + C6
H6 = ((C19*C20 + C6*C8) / H5) if H5 else 0.0
H7 = ((C19*C21 + C6*C9) / H5) if H5 else 0.0

H8 = st.number_input("Consumo específico to be H8 (variar)", min_value=0.0, value=float(defaults["H8_consumo_to_be"]), step=0.1, format="%.3f")

H9  = H5*H7/96.0*H8
H10 = H9 - C11
H11 = H5*H7/96.0
H12 = H11*24.0
H13 = H11*1.2
H14 = H5 - H11*0.786 + H9 - H13
H15 = (H5*H6/H14) if H14 else 0.0
H16 = H14 - C15
H17 = (C4*C5/H14) if H14 else 0.0

st.markdown('<div class="grid3">', unsafe_allow_html=True)
st.markdown(f'<div class="card"><b>H11 Etanol hidratado (m³/h)</b><div style="font-size:1.3rem">{fmt(H11)}</div><div>Fórmula: H5*H7/96</div></div>', unsafe_allow_html=True)
st.markdown(f'<div class="card"><b>H14 Vinhaça (m³/h)</b><div style="font-size:1.3rem">{fmt(H14)}</div><div>Fórmula: H5 - H11*0.786 + H9 - H13</div></div>', unsafe_allow_html=True)
st.markdown(f'<div class="card"><b>H15 %Ds</b><div style="font-size:1.3rem">{fmt(H15*100,2)} %</div><div>Fórmula: H5*H6/H14</div></div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

st.header("4) Produções & Financeiro")
preco_etanol = st.number_input("Preço etanol (R$/m³)", min_value=0.0, value=2800.0, step=50.0, format="%.2f")
producao_as_is_m3dia = C13
producao_to_be_m3dia = H12
delta_producao = producao_to_be_m3dia - producao_as_is_m3dia
st.markdown(f"<div class='card'><b>Δ Produção (m³/dia):</b> {fmt(delta_producao)}</div>", unsafe_allow_html=True)
st.markdown(f"<div class='card'><b>Δ Receita (R$/dia):</b> {fmt(delta_producao*preco_etanol,2)}</div>", unsafe_allow_html=True)
