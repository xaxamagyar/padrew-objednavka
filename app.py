import io
import re
import urllib.request
import pandas as pd
import streamlit as st

# Importy pro ReportLab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="Objednávky & Štítky", page_icon="📦", layout="wide")
st.title("📦 Automatizace Objednávky & Štítky")

# ==============================================================================
# 🔤 NAČTENÍ FONTŮ PŘÍMO Z TVÉHO GITHUBŪ
# ==============================================================================
@st.cache_resource
def nacist_fonty_z_vlastniho_githubu():
    try:
        url_normal = "https://raw.githubusercontent.com/xaxamagyar/padrew-objednavka/main/ARIAL.TTF"
        url_bold = "https://raw.githubusercontent.com/xaxamagyar/padrew-objednavka/main/ARIALBD.TTF"
        
        req_normal = urllib.request.urlopen(url_normal)
        req_bold = urllib.request.urlopen(url_bold)
        
        pdfmetrics.registerFont(TTFont('Padrew-Arial', io.BytesIO(req_normal.read())))
        pdfmetrics.registerFont(TTFont('Padrew-Arial-Bold', io.BytesIO(req_bold.read())))
        return 'Padrew-Arial', 'Padrew-Arial-Bold'
    except Exception as e:
        st.warning(f"⚠ Nepodařilo se načíst fonty z GitHubu ({e}). Použije se Helvetica.")
        return 'Helvetica', 'Helvetica-Bold'

FONT_NORMAL, FONT_BOLD = nacist_fonty_z_vlastniho_githubu()

# 🔗 RAW ODKAZ NA EXCEL Z TVÉHO GITHUBŪ
GITHUB_EXCEL_URL = "https://raw.githubusercontent.com/xaxamagyar/padrew-objednavka/main/translate-PL.xlsx"

# --- 1. VÝBĚR DODAVATELE ---
st.sidebar.header("⚙️ Nastavení dodavatele")
DODAVATEL = st.sidebar.selectbox("Vyberte dodavatele:", ["PADREW", "KOPER"])

st.write(f"Aplikace čerpá překlady pro dodavatele **{DODAVATEL}** z GitHubu. Nahrajte objednávky.")

# Inicializace vnitřní paměti pro VŠECHNY listy najednou
if "all_sheets" not in st.session_state:
    st.session_state.all_sheets = {}

# --- AUTOMATICKÉ NAČTENÍ KOMPLETNÍHO EXCELU JAKO ČISTÝ TEXT ---
if not st.session_state.all_sheets:
    try:
        with st.spinner("🔄 Načítám kompletní překladovou databázi z GitHubu..."):
            # Otevřeme soubor přes ExcelFile, abychom zmapovali strukturu
            excel_file = pd.ExcelFile(GITHUB_EXCEL_URL)
            for sheet in excel_file.sheet_names:
                # ⭐ KLÍČOVÁ OPRAVA: Vynutíme dtype=str pro VŠECHNY sloupce při načítání, což smaže chybu float64
                df_sheet = pd.read_excel(GITHUB_EXCEL_URL, sheet_name=sheet, dtype=str)
                
                # Vyčistíme prázdné hodnoty na čisté textové řetězce
                df_sheet = df_sheet.fillna("").astype(str)
                for col in df_sheet.columns:
                    df_sheet[col] = df_sheet[col].str.strip()
                
                st.session_state.all_sheets[sheet] = df_sheet
            st.toast("✔ Databáze kompletně načtena z GitHubu!", icon="📥")
    except Exception as e:
        st.error(f"❌ Chyba při stahování Excelu z GitHubu: {e}")
        st.stop()

# Dynamické přiřazení listů podle vybraného dodavatele
sheet_data = f"DATA-{DODAVATEL}"
sheet_var = f"VAR-{DODAVATEL}"

# Pojistka pro případ, že by záložky v Excelu chyběly (vytvoří prázdné s textovými sloupci)
if sheet_data not in st.session_state.all_sheets:
    st.session_state.all_sheets[sheet_data] = pd.DataFrame(columns=["název", "PL"]).astype(str)
if sheet_var not in st.session_state.all_sheets:
    st.session_state.all_sheets[sheet_var] = pd.DataFrame(columns=["VAR", "PL"]).astype(str)
if "LABELS" not in st.session_state.all_sheets:
    st.session_state.all_sheets["LABELS"] = pd.DataFrame(columns=["NAME", "PCS"]).astype(str)

# --- BOČNÍ PANEL: NAHRÁNÍ OBJEDNÁVEK ---
st.sidebar.markdown("---")
st.sidebar.header("📂 2. Nahrání objednávek")
uploaded_orders1 = st.sidebar.file_uploader("Vyberte orders.xlsx (E-shop 1)", type=["xlsx"])
uploaded_orders2 = st.sidebar.file_uploader("Vyberte orders-v.xlsx (E-shop 2) - nepovinné", type=["xlsx"])

if uploaded_orders1:
    try:
        # Spojení e-shopů
        seznam_df_orders = [pd.read_excel(uploaded_orders1)]
        if uploaded_orders2:
            seznam_df_orders.append(pd.read_excel(uploaded_orders2))
        df_orders = pd.concat(seznam_df_orders, ignore_index=True)

        # DYNAMICKÁ FILTRACE PODLE VYBRANÉHO DODAVATELE
        filtr = (
            (df_orders["orderItemType"] == "product")
            & (df_orders["orderItemSupplier"] == DODAVATEL)
            & (df_orders["orderItemStatusName"] != "Stornována")
        )
        df_filtrovane = df_orders[filtr].copy()

        if df_filtrovane.empty:
            st.warning(f"ℹ Žádné nevyřízené objednávky pro dodavatele: {DODAVATEL}")
        else:
            # Sloučení šuplíků k postelím
            je_samostatny_suplik = df_filtrovane["orderItemName"].str.contains(r"Zásuvka|šuplík", case=False, na=False)
            objednavky_se_suplikem = df_filtrovane[je_samostatny_suplik]["code"].unique()
            radky_ke_smazani = []

            for kod_objednavky in objednavky_se_suplikem:
                df_objednavka = df_filtrovane[df_filtrovane["code"] == kod_objednavky]
                indexy_posteli = df_objednavka[~df_objednavka["orderItemName"].str.contains(r"Zásuvka|šuplík", case=False, na=False)].index
                indexy_supliku = df_objednavka[df_objednavka["orderItemName"].str.contains(r"Zásuvka|šuplík", case=False, na=False)].index

                if len(indexy_posteli) > 0 and len(indexy_supliku) > 0:
                    puvodni_nazev = df_filtrovane.loc[indexy_posteli[0], "orderItemName"]
                    if "ŠUPLÍK" not in puvodni_nazev.upper():
                        df_filtrovane.loc[indexy_posteli[0], "orderItemName"] = puvodni_nazev + "  + ŠUPLÍKY"
                    radky_ke_smazani.extend(indexy_supliku)

            if radky_ke_smazani:
                df_filtrovane = df_filtrovane.drop(index=radky_ke_smazani)

            # Příprava dat pro kontrolu překladů
            df_check = df_filtrovane[["code", "orderItemName", "orderItemAmount", "orderItemVariantName"]].copy()
            df_check["orderItemName"] = df_check["orderItemName"].astype(str).str.strip()
            df_check["orderItemVariantName"] = df_check["orderItemVariantName"].fillna("").astype(str).str.strip()

            # Vytvoření čistě textových slovníků z paměti aplikace
            slovnik_prod = dict(zip(st.session_state.all_sheets[sheet_data]["název"], st.session_state.all_sheets[sheet_data]["PL"]))
            slovnik_var = dict(zip(st.session_state.all_sheets[sheet_var]["VAR"], st.session_state.all_sheets[sheet_var]["PL"]))
            slovnik_lab = dict(zip(st.session_state.all_sheets["LABELS"]["NAME"], st.session_state.all_sheets["LABELS"]["PCS"]))

            df_check["orderItemName_PL"] = df_check["orderItemName"].map(slovnik_prod)
            df_check["orderItemVariantName_PL"] = df_check["orderItemVariantName"].map(slovnik_var)
            df_check.loc[df_check["orderItemVariantName"] == "", "orderItemVariantName_PL"] = ""

            # Ochrana prázdných hodnot po mapování
            df_check["orderItemName_PL"] = df_check["orderItemName_PL"].replace("", None).fillna("")
            df_check["orderItemVariantName_PL"] = df_check["orderItemVariantName_PL"].replace("", None).fillna("")

            # Detekce chybějících položek
            chybi_prod = df_check[df_check["orderItemName_PL"] == ""]["orderItemName"].unique()
            chybi_var = df_check[(df_check["orderItemVariantName_PL"] == "") & (df_check["orderItemVariantName"] != "")]["orderItemVariantName"].unique()

            df_check["orderItemName_PL_clean"] = df_check["orderItemName_PL"]
            df_check["orderItemVariantName_PL_clean"] = df_check["orderItemVariantName_PL"]
            
            obsahuje_sz_pl = df_check["orderItemVariantName_PL_clean"].str.contains(r"\+?\s*SZ", case=False, na=False)
            df_check.loc[obsahuje_sz_pl & ~df_check["orderItemName_PL_clean"].str.contains(r"\+?\s*SZ", case=False), "orderItemName_PL_clean"] = (df_check["orderItemName_PL_clean"] + " + SZ")
            
            df_check["baliky_pcs"] = df_check["orderItemName_PL_clean"].map(slovnik_lab).fillna("")
            chybi_lab = df_check[(df_check["baliky_pcs"] == "") & (df_check["orderItemName_PL_clean"] != "")]["orderItemName_PL_clean"].unique()

            # --- FORMULÁŘE PRO DOPLNĚNÍ (ZDE TO PADALO, NYNÍ OŠETŘENO) ---
            if len(chybi_prod) > 0 or len(chybi_var) > 0 or len(chybi_lab) > 0:
                st.error(f"🛑 V online databázi ({DODAVATEL}) chybí položky z nových objednávek! Vyplňte je zde:")
                
                if len(chybi_prod) > 0:
                    st.warning("➕ Chybějící PŘEKLADY PRODUKTŮ:")
                    for i, p_cz in enumerate(chybi_prod):
                        with st.form(key=f"form_prod_{i}"):
                            st.write(f"Produkt z e-shopu: **{p_cz}**")
                            p_pl = st.text_input("Zadejte polský kód/překlad:", key=f"in_prod_{i}")
                            if st.form_submit_button("Přidat produkt"):
                                if p_pl:
                                    novy_radek = pd.DataFrame([{"název": str(p_cz), "PL": str(p_pl.strip())}]).astype(str)
                                    st.session_state.all_sheets[sheet_data] = pd.concat([st.session_state.all_sheets[sheet_data], novy_radek], ignore_index=True)
                                    st.rerun()

                if len(chybi_var) > 0:
                    st.warning("➕ Chybějící PŘEKLADY VARIANT:")
                    for i, v_cz in enumerate(chybi_var):
                        with st.form(key=f"form_var_{i}"):
                            st.write(f"Varianta z e-shopu: **{v_cz}**")
                            v_pl = st.text_input("Zadejte polský překlad varianty:", key=f"in_var_{i}")
                            if st.form_submit_button("Přidat variantu"):
                                if v_pl:
                                    novy_radek = pd.DataFrame([{"VAR": str(v_cz), "PL": str(v_pl.strip())}]).astype(str)
                                    st.session_state.all_sheets[sheet_var] = pd.concat([st.session_state.all_sheets[sheet_var], novy_radek], ignore_index=True)
                                    st.rerun()

                if len(chybi_lab) > 0 and len(chybi_prod) == 0:
                    st.warning("➕ Chybějící POČTY BALÍKŮ (Společná záložka LABELS):")
                    for i, l_pl in enumerate(chybi_lab):
                        with st.form(key=f"form_lab_{i}"):
                            st.write(f"Polský název produktu: **{l_pl}**")
                            l_pcs = st.number_input("Počet krabic (PCS):", min_value=1, max_value=20, value=2, key=f"in_lab_{i}")
                            if st.form_submit_button("Přidat počet balíků"):
                                novy_radek = pd.DataFrame([{"NAME": str(l_pl), "PCS": str(int(l_pcs))}]).astype(str)
                                st.session_state.all_sheets["LABELS"] = pd.concat([st.session_state.all_sheets["LABELS"], novy_radek], ignore_index=True)
                                st.rerun()
                st.stop()

            # --- GENERACE VÝSTUPŮ ---
            df_final = df_check.copy()
            df_final["baliky_pcs"] = pd.to_numeric(df_final["baliky_pcs"], errors='coerce').fillna(1).astype(int)
            df_final["orderItemVariantName_PL_clean"] = df_final["orderItemVariantName_PL_clean"].str.replace(r",?\s*\+?\s*SZ", "", case=False, regex=True).str.strip(", ").str.strip()

            df_final = df_final.sort_values(by=["orderItemName_PL_clean", "orderItemVariantName_PL_clean"], ascending=True)

            def sestav_radek(row):
                kod, ks, nazev, var = str(row["code"]), f"{row['orderItemAmount']} szt.", row["orderItemName_PL_clean"], row["orderItemVariantName_PL_clean"]
                return f"{kod} - {ks} {nazev} ({var})" if var else f"{kod} - {ks} {nazev}"
            df_final["Radek_pro_dodavatele"] = df_final.apply(sestav_radek, axis=1)

            st.subheader(f"📋 Přehled řádků pro e-mail dodavateli: {DODAVATEL}")
            st.code("\n".join(df_final["Radek_pro_dodavatele"].tolist()))

            col_down1, col_down2 = st.columns(2)
            with col_down1:
                excel_buffer = io.BytesIO()
                df_final[["code", "orderItemName", "orderItemAmount", "orderItemVariantName", "Radek_pro_dodavatele"]].to_excel(excel_buffer, index=False)
                st.download_button(label="🟢 Stáhnout e-mailový přehled (Excel)", data=excel_buffer.getvalue(), file_name=f"vystup_{DODAVATEL}.xlsx")

            with col_down2:
                pdf_buffer = io.BytesIO()
                doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, leftMargin=3*mm, rightMargin=3*mm, topMargin=0*mm, bottomMargin=0*mm)
                
                styles = getSampleStyleSheet()
                style_top_left = ParagraphStyle('TopLeft', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=10, leading=11)
                style_top_right = ParagraphStyle('TopRight', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=10, leading=11, alignment=2)
                style_main = ParagraphStyle('Main', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=11, leading=13)
                style_pl = ParagraphStyle('PL', parent=styles['Normal'], fontName=FONT_NORMAL, fontSize=8.5, leading=10, textColor=colors.HexColor('#555555'))

                fraze_ke_smazani = ["se zásuvkou", "s úložným prostorem", "včetně roštu", " matrace"]
                
                bunky_stitku = []
                for _, row in df_final.iterrows():
                    kod, cz_nazev, cz_var, pl_nazev = str(row["code"]), row["orderItemName"], row["orderItemVariantName"], row["orderItemName_PL_clean"]
                    mnozstvi_objednano, baliku_na_produkt = int(row["orderItemAmount"]), int(row["baliky_pcs"])
                    
                    for i_postele in range(1, mnozstvi_objednano + 1):
                        for a_balik in range(1, baliku_na_produkt + 1):
                            
                            cz_var_cista = cz_var
                            for fraze in fraze_ke_smazani:
                                cz_var_cista = re.sub(re.escape(fraze), "", cz_var_cista, flags=re.IGNORECASE)
                            cz_var_cista = cz_var_cista.replace(", ,", ",").strip(", ").strip()
                            
                            cz_text = f"{cz_nazev} - {cz_var_cista}" if cz_var_cista else cz_nazev
                            pl_text = f"[PL: {pl_nazev}]"
                            
                            horni_tab = Table([[Paragraph(f"OBJEDNÁVKA: {kod}", style_top_left), Paragraph(f"Balík: {a_balik} z {baliku_na_produkt}", style_top_right)]], colWidths=[48*mm, 44*mm])
                            horni_tab.setStyle(TableStyle([('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0)]))
                            
                            bunky_stitku.append([
                                horni_tab, 
                                Spacer(1, 1.5*mm), 
                                Paragraph(cz_text, style_main), 
                                Spacer(1, 1*mm), 
                                Paragraph(pl_text, style_pl)
                            ])

                data_mrizky = []
                docasny_radek = []
                while len(bunky_stitku) % 14 != 0: 
                    bunky_stitku.append("")
                for stitek in bunky_stitku:
                    docasny_radek.append(stitek)
                    if len(docasny_radek) == 2:
                        data_mrizky.append(docasny_radek)
                        docasny_radek = []
                        
                mrizka_tabulka = Table(data_mrizky, colWidths=[102*mm, 102*mm], rowHeights=[41.7*mm]*len(data_mrizky))
                t_style = [('VALIGN', (0,0), (-1,-1), 'TOP'), ('LEFTPADDING', (0,0), (-1,-1), 2.5*mm), ('RIGHTPADDING', (0,0), (-1,-1), 2.5*mm), ('TOPPADDING', (0,0), (-1,-1), 2.5*mm), ('BOTTOMPADDING', (0,0), (-1,-1), 1.5*mm), ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#DDDDDD')), ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#DDDDDD'))]
                for r in range(7, len(data_mrizky), 7): 
                    t_style.append(('PAGEBREAK', (0, r-1), (-1, r-1)))
                mrizka_tabulka.setStyle(TableStyle(t_style))
                doc.build([mrizka_tabulka])

                st.download_button(label=f"🔵 Stáhnout štítky {DODAVATEL} 2x7 (PDF)", data=pdf_buffer.getvalue(), file_name=f"stitky_{DODAVATEL}.pdf")
                st.balloons()

    except Exception as e:
        st.error(f"❌ Chyba při zpracování objednávek: {e}")

# --- ŽIVÁ SPRÁVA DATABÁZE ---
st.markdown("---")
st.subheader(f"⚙️ Živá správa překladů pro list: {sheet_data} a {sheet_var}")

tab1, tab2, tab3 = st.tabs([f"🛒 Překlady produktů ({sheet_data})", f"🎨 Překlady variant ({sheet_var})", "📦 Počty balíků (Společný list LABELS)"])
with tab1:
    st.session_state.all_sheets[sheet_data] = st.data_editor(st.session_state.all_sheets[sheet_data], num_rows="dynamic", use_container_width=True, key=f"edit_prod_{DODAVATEL}")
with tab2:
    st.session_state.all_sheets[sheet_var] = st.data_editor(st.session_state.all_sheets[sheet_var], num_rows="dynamic", use_container_width=True, key=f"edit_var_{DODAVATEL}")
with tab3:
    st.session_state.all_sheets["LABELS"] = st.data_editor(st.session_state.all_sheets["LABELS"], num_rows="dynamic", use_container_width=True, key="edit_lab_global")

st.markdown("### 💾 Stažení kompletního aktualizovaného Excelu")
st.write("Po úpravách stáhněte soubor a nahrajte jej zpět na GitHub repozitář. Budou v něm stoprocentně zachováni všichni dodavatelé.")

export_buffer = io.BytesIO()
with pd.ExcelWriter(export_buffer, engine='openpyxl') as writer:
    # ⭐ BEZPEČNÝ ZÁPIS: Uložíme všechny listy z vnitřní paměti aplikace najednou, aby se žádný nesmazal
    for s_name, df_s in st.session_state.all_sheets.items():
        df_s.to_excel(writer, sheet_name=s_name, index=False)
        
st.download_button(
    label="📥 STÁHNOUT SOUBOR TRANSLATE-PL.XLSX SE ZMÊNAMI",
    data=export_buffer.getvalue(),
    file_name="translate-PL.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)