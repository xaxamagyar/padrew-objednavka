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

st.set_page_config(page_title="PADREW Automatizace", page_icon="📦", layout="wide")
st.title("📦 PADREW: Objednávky, Štítky & Online Databáze")

# ==============================================================================
# 🔤 NAČTENÍ NAHRANÝCH FONTŮ PŘÍMO Z TVÉHO GITHUBŪ (SROVNÁNO NA VELKÁ PÍSMENA)
# ==============================================================================
@st.cache_resource
def nacist_fonty_z_vlastniho_githubu():
    try:
        # Odkazy směřují přesně na soubory ARIAL.TTF a ARIALBD.TTF ve tvém repozitáři
        url_normal = "https://raw.githubusercontent.com/xaxamagyar/padrew-objednavka/main/ARIAL.TTF"
        url_bold = "https://raw.githubusercontent.com/xaxamagyar/padrew-objednavka/main/ARIALBD.TTF"
        
        req_normal = urllib.request.urlopen(url_normal)
        req_bold = urllib.request.urlopen(url_bold)
        
        pdfmetrics.registerFont(TTFont('Padrew-Arial', io.BytesIO(req_normal.read())))
        pdfmetrics.registerFont(TTFont('Padrew-Arial-Bold', io.BytesIO(req_bold.read())))
        return 'Padrew-Arial', 'Padrew-Arial-Bold'
    except Exception as e:
        st.warning(f"⚠ Nepodařilo se načíst fonty z tvého GitHubu ({e}). Použije se Helvetica (může tvořit bloky).")
        return 'Helvetica', 'Helvetica-Bold'

FONT_NORMAL, FONT_BOLD = nacist_fonty_z_vlastniho_githubu()

# ==============================================================================
# 🔗 RAW ODKAZ NA EXCEL Z TVÉHO GITHUBŪ
# ==============================================================================
GITHUB_EXCEL_URL = "https://raw.githubusercontent.com/xaxamagyar/padrew-objednavka/main/translate-PL.xlsx"
# ==============================================================================

st.write("Aplikace automaticky čerpá překladovou databázi z GitHubu. Stačí nahrát objednávky.")

# Inicializace vnitřní paměti pro editovatelné databáze
if "df_trans_prod" not in st.session_state:
    st.session_state.df_trans_prod = None
if "df_trans_var" not in st.session_state:
    st.session_state.df_trans_var = None
if "df_labels" not in st.session_state:
    st.session_state.df_labels = None

# --- AUTOMATICKÉ NAČTENÍ DATABÁZE Z GITHUBŪ PŘI SPUŠTĚNÍ ---
if st.session_state.df_trans_prod is None:
    try:
        with st.spinner("🔄 Načítám aktuální překladovou databázi z GitHubu..."):
            st.session_state.df_trans_prod = pd.read_excel(GITHUB_EXCEL_URL, sheet_name="DATA-PAWEL")
            st.session_state.df_trans_var = pd.read_excel(GITHUB_EXCEL_URL, sheet_name="VAR-PAWEL")
            st.session_state.df_labels = pd.read_excel(GITHUB_EXCEL_URL, sheet_name="LABELS")
            st.toast("✔ Databáze z GitHubu úspěšně načtena!", icon="📥")
    except Exception as e:
        st.error(f"❌ Nepodařilo se stáhnout databázi z GitHubu. Zkontrolujte odkaz GITHUB_EXCEL_URL. Chyba: {e}")
        st.stop()

# --- BOČNÍ PANEL: OBJEDNÁVKY ---
st.sidebar.header("📂 1. Nahrání objednávek")
uploaded_orders1 = st.sidebar.file_uploader("Vyberte orders.xlsx (E-shop 1)", type=["xlsx"])
uploaded_orders2 = st.sidebar.file_uploader("Vyberte orders-v.xlsx (E-shop 2) - nepovinné", type=["xlsx"])

# Zpracování objednávek
if uploaded_orders1 and st.session_state.df_trans_prod is not None:
    try:
        seznam_df_orders = [pd.read_excel(uploaded_orders1)]
        if uploaded_orders2:
            seznam_df_orders.append(pd.read_excel(uploaded_orders2))
        df_orders = pd.concat(seznam_df_orders, ignore_index=True)

        filtr = (
            (df_orders["orderItemType"] == "product")
            & (df_orders["orderItemSupplier"] == "PADREW")
            & (df_orders["orderItemStatusName"] != "Stornována")
        )
        df_filtrovane = df_orders[filtr].copy()

        if df_filtrovane.empty:
            st.warning("ℹ Žádné nevyřízené objednávky pro PADREW.")
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

            df_check = df_filtrovane[["code", "orderItemName", "orderItemAmount", "orderItemVariantName"]].copy()
            df_check["orderItemName"] = df_check["orderItemName"].astype(str).str.strip()
            df_check["orderItemVariantName"] = df_check["orderItemVariantName"].fillna("").astype(str).str.strip()

            slovnik_prod = dict(zip(st.session_state.df_trans_prod["název"].astype(str).str.strip(), st.session_state.df_trans_prod["PL"]))
            slovnik_var = dict(zip(st.session_state.df_trans_var["VAR"].astype(str).str.strip(), st.session_state.df_trans_var["PL"]))
            slovnik_lab = dict(zip(st.session_state.df_labels["NAME"].astype(str).str.strip(), st.session_state.df_labels["PCS"]))

            df_check["orderItemName_PL"] = df_check["orderItemName"].map(slovnik_prod)
            df_check["orderItemVariantName_PL"] = df_check["orderItemVariantName"].map(slovnik_var)
            df_check.loc[df_check["orderItemVariantName"] == "", "orderItemVariantName_PL"] = ""

            chybi_prod = df_check[df_check["orderItemName_PL"].isna()]["orderItemName"].unique()
            chybi_var = df_check[df_check["orderItemVariantName_PL"].isna() & (df_check["orderItemVariantName"] != "")]["orderItemVariantName"].unique()

            df_check["orderItemName_PL_clean"] = df_check["orderItemName_PL"].fillna("")
            df_check["orderItemVariantName_PL_clean"] = df_check["orderItemVariantName_PL"].fillna("")
            
            obsahuje_sz_pl = df_check["orderItemVariantName_PL_clean"].str.contains(r"\+?\s*SZ", case=False, na=False)
            df_check.loc[obsahuje_sz_pl & ~df_check["orderItemName_PL_clean"].str.contains(r"\+?\s*SZ", case=False), "orderItemName_PL_clean"] = (df_check["orderItemName_PL_clean"] + " + SZ")
            
            df_check["baliky_pcs"] = df_check["orderItemName_PL_clean"].map(slovnik_lab)
            chybi_lab = df_check[df_check["baliky_pcs"].isna() & (df_check["orderItemName_PL_clean"] != "")]["orderItemName_PL_clean"].unique()

            # --- DYNAMICKÉ FORMULÁŘE PRO CHYBĚJÍCÍ POLOŽKY ---
            if len(chybi_prod) > 0 or len(chybi_var) > 0 or len(chybi_lab) > 0:
                st.error("🛑 V databázi chybí položky! Vyplňte je zde:")
                
                if len(chybi_prod) > 0:
                    st.warning("➕ Chybějící PŘEKLADY PRODUKTŮ:")
                    for i, p_cz in enumerate(chybi_prod):
                        with st.form(key=f"form_prod_{i}"):
                            st.write(f"Produkt z e-shopu: **{p_cz}**")
                            p_pl = st.text_input("Zadejte polský kód (např. 8X8 3A):", key=f"in_prod_{i}")
                            if st.form_submit_button("Přidat produkt"):
                                if p_pl:
                                    novy_radek = pd.DataFrame([{"název": p_cz, "PL": p_pl.strip()}])
                                    st.session_state.df_trans_prod = pd.concat([st.session_state.df_trans_prod, novy_radek], ignore_index=True)
                                    st.rerun()

                if len(chybi_var) > 0:
                    st.warning("➕ Chybějící PŘEKLADY VARIANT:")
                    for i, v_cz in enumerate(chybi_var):
                        with st.form(key=f"form_var_{i}"):
                            st.write(f"Varianta z e-shopu: **{v_cz}**")
                            v_pl = st.text_input("Zadejte polský překlad varianty:", key=f"in_var_{i}")
                            if st.form_submit_button("Přidat variantu"):
                                if v_pl:
                                    novy_radek = pd.DataFrame([{"VAR": v_cz, "PL": v_pl.strip()}])
                                    st.session_state.df_trans_var = pd.concat([st.session_state.df_trans_var, novy_radek], ignore_index=True)
                                    st.rerun()

                if len(chybi_lab) > 0 and len(chybi_prod) == 0:
                    st.warning("➕ Chybějící POČTY BALÍKŮ (LABELS):")
                    for i, l_pl in enumerate(chybi_lab):
                        with st.form(key=f"form_lab_{i}"):
                            st.write(f"Polský název produktu: **{l_pl}**")
                            l_pcs = st.number_input("Počet krabic (PCS):", min_value=1, max_value=20, value=2, key=f"in_lab_{i}")
                            if st.form_submit_button("Přidat počet balíků"):
                                novy_radek = pd.DataFrame([{"NAME": l_pl, "PCS": int(l_pcs)}])
                                st.session_state.df_labels = pd.concat([st.session_state.df_labels, novy_radek], ignore_index=True)
                                st.rerun()
                st.stop()

            # ==============================================================================
            # GENERACE VÝSTUPŮ (VŠE OK)
            # ==============================================================================
            df_final = df_check.copy()
            df_final["baliky_pcs"] = df_final["baliky_pcs"].astype(int)
            df_final["orderItemVariantName_PL_clean"] = df_final["orderItemVariantName_PL_clean"].str.replace(r",?\s*\+?\s*SZ", "", case=False, regex=True).str.strip(", ").str.strip()

            df_final = df_final.sort_values(by=["orderItemName_PL_clean", "orderItemVariantName_PL_clean"], ascending=True)

            def sestav_radek(row):
                kod, ks, nazev, var = str(row["code"]), f"{row['orderItemAmount']} szt.", row["orderItemName_PL_clean"], row["orderItemVariantName_PL_clean"]
                return f"{kod} - {ks} {nazev} ({var})" if var else f"{kod} - {ks} {nazev}"
            df_final["Radek_pro_dodavatele"] = df_final.apply(sestav_radek, axis=1)

            st.subheader("📋 Přehled řádků pro e-mail dodavateli")
            st.code("\n".join(df_final["Radek_pro_dodavatele"].tolist()))

            col_down1, col_down2 = st.columns(2)
            with col_down1:
                excel_buffer = io.BytesIO()
                df_final[["code", "orderItemName", "orderItemAmount", "orderItemVariantName", "Radek_pro_dodavatele"]].to_excel(excel_buffer, index=False)
                st.download_button(label="🟢 Stáhnout e-mailový přehled (Excel)", data=excel_buffer.getvalue(), file_name="vystup_kontrola.xlsx")

            with col_down2:
                pdf_buffer = io.BytesIO()
                doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, leftMargin=3*mm, rightMargin=3*mm, topMargin=0*mm, bottomMargin=0*mm)
                
                # Použití tvých ověřených a načtených fontů z GitHubu
                styles = getSampleStyleSheet()
                style_top_left = ParagraphStyle('TopLeft', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=10, leading=11)
                style_top_right = ParagraphStyle('TopRight', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=10, leading=11, alignment=2)
                style_main = ParagraphStyle('Main', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=11, leading=13)
                style_pl = ParagraphStyle('PL', parent=styles['Normal'], fontName=FONT_NORMAL, fontSize=8.5, leading=10, textColor=colors.HexColor('#555555'))

                # OČIŠTĚNÍ OD BALASTU (FRÁZE KE SMAZÁNÍ)
                fraze_ke_smazani = ["se zásuvkou", "s úložným prostorem", "včetně roštu", " matrace"]
                
                bunky_stitku = []
                for _, row in df_final.iterrows():
                    kod, cz_nazev, cz_var, pl_nazev = str(row["code"]), row["orderItemName"], row["orderItemVariantName"], row["orderItemName_PL_clean"]
                    mnozstvi_objednano, baliku_na_produkt = int(row["orderItemAmount"]), int(row["baliky_pcs"])
                    
                    for i_postele in range(1, mnozstvi_objednano + 1):
                        for a_balik in range(1, baliku_na_produkt + 1):
                            
                            # Odmazání nežádoucích slov z varianty pro čistý tisk
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

                st.download_button(label="🔵 Stáhnout tiskové štítky 2x7 (PDF)", data=pdf_buffer.getvalue(), file_name="stitky_k_tisku.pdf")
                st.balloons()

    except Exception as e:
        st.error(f"❌ Chyba při zpracování objednávek: {e}")

# --- ŽIVÁ SPRÁVA DATABÁZE PŘÍMO NA WEBU ---
st.markdown("---")
st.subheader("⚙️ Živá správa překladové databáze z GitHubu")

tab1, tab2, tab3 = st.tabs(["🛒 Překlady produktů", "🎨 Překlady variant", "📦 Počty balíků"])
with tab1:
    st.session_state.df_trans_prod = st.data_editor(st.session_state.df_trans_prod, num_rows="dynamic", use_container_width=True, key="edit_prod")
with tab2:
    st.session_state.df_trans_var = st.data_editor(st.session_state.df_trans_var, num_rows="dynamic", use_container_width=True, key="edit_var")
with tab3:
    st.session_state.df_labels = st.data_editor(st.session_state.df_labels, num_rows="dynamic", use_container_width=True, key="edit_lab")

st.markdown("### 💾 Stažení aktualizované databáze k nahrání na GitHub")
st.write("Pokud jste upravovali tabulky nebo doplňovali chybějící data, stáhněte si tento soubor a přetáhněte ho zpět na GitHub.")

export_buffer = io.BytesIO()
with pd.ExcelWriter(export_buffer, engine='openpyxl') as writer:
    st.session_state.df_trans_prod.to_excel(writer, sheet_name="DATA-PAWEL", index=False)
    st.session_state.df_trans_var.to_excel(writer, sheet_name="VAR-PAWEL", index=False)
    st.session_state.df_labels.to_excel(writer, sheet_name="LABELS", index=False)
    
st.download_button(
    label="📥 STÁHNOUT AKTUALIZOVANÝ SOUBOR TRANSLATE-PL.XLSX",
    data=export_buffer.getvalue(),
    file_name="translate-PL.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)