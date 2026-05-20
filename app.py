import io
import os
import sys
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

# Nastavení vzhledu Streamlit stránky
st.set_page_config(page_title="PADREW Automatizace", page_icon="📦", layout="centered")
st.title("📦 PADREW: Objednávky & Štítky")
st.write("Nahrajte exporty z e-shopů a překladový soubor pro vygenerování podkladů.")

# Pokus o načtení fontu (na serveru GitHubu/Streamlitu Helvetica, lokálně na Windows Arial)
try:
    if os.path.exists(os.path.join(os.environ.get('WINDIR', ''), 'Fonts', 'arial.ttf')):
        arial_path = os.path.join(os.environ['WINDIR'], 'Fonts', 'arial.ttf')
        arial_bold_path = os.path.join(os.environ['WINDIR'], 'Fonts', 'arialbd.ttf')
        arial_italic_path = os.path.join(os.environ['WINDIR'], 'Fonts', 'ariali.ttf')
        pdfmetrics.registerFont(TTFont('Arial', arial_path))
        pdfmetrics.registerFont(TTFont('Arial-Bold', arial_bold_path))
        pdfmetrics.registerFont(TTFont('Arial-Italic', arial_italic_path))
        FONT_NAME, FONT_BOLD, FONT_ITALIC = 'Arial', 'Arial-Bold', 'Arial-Italic'
    else:
        FONT_NAME, FONT_BOLD, FONT_ITALIC = 'Helvetica', 'Helvetica-Bold', 'Helvetica-Oblique'
except Exception:
    FONT_NAME, FONT_BOLD, FONT_ITALIC = 'Helvetica', 'Helvetica-Bold', 'Helvetica-Oblique'

# --- 1. UFOAD SOUBORŮ PŘES WEBOVÉ ROZHRANÍ ---
st.sidebar.header("📂 Nahrání souborů")
uploaded_orders1 = st.sidebar.file_uploader("Vyberte orders.xlsx (E-shop 1)", type=["xlsx"])
uploaded_orders2 = st.sidebar.file_uploader("Vyberte orders-v.xlsx (E-shop 2) - nepovinné", type=["xlsx"])
uploaded_translate = st.sidebar.file_uploader("Vyberte translate-PL.xlsx (Překlady)", type=["xlsx"])

if uploaded_orders1 and uploaded_translate:
    try:
        # Načtení dat z paměti prohlížeče
        seznam_df_orders = [pd.read_excel(uploaded_orders1)]
        if uploaded_orders2:
            seznam_df_orders.append(pd.read_excel(uploaded_orders2))
            
        df_orders = pd.concat(seznam_df_orders, ignore_index=True)
        
        df_trans_prod = pd.read_excel(uploaded_translate, sheet_name="DATA-PAWEL")
        df_trans_var = pd.read_excel(uploaded_translate, sheet_name="VAR-PAWEL")
        df_labels = pd.read_excel(uploaded_translate, sheet_name="LABELS")
        
        st.success("✔ Všechny soubory úspěšně nahrány do paměti aplikace.")

        # --- 2. FILTRACE DAT ---
        filtr = (
            (df_orders["orderItemType"] == "product")
            & (df_orders["orderItemSupplier"] == "PADREW")
            & (df_orders["orderItemStatusName"] != "Stornována")
        )
        df_filtrovane = df_orders[filtr].copy()

        if df_filtrovane.empty:
            st.warning("ℹ Žádné objednávky k vyřízení od PADREW.")
        else:
            # --- 3. SLOUČENÍ ZÁSUVEK ---
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
                st.info(f"✔ Sloučeno a očištěno {len(radky_ke_smazani)} samostatných šuplíků přímo k postelím.")

            # Očištění textů
            df_final = df_filtrovane[["code", "orderItemName", "orderItemAmount", "orderItemVariantName"]].copy()
            df_final["orderItemName"] = df_final["orderItemName"].astype(str).str.strip()
            df_final["orderItemVariantName"] = df_final["orderItemVariantName"].fillna("").astype(str).str.strip()

            # --- 4. PŘEKLAD DO POLŠTINY ---
            slovnik_produktů = dict(zip(df_trans_prod["název"].astype(str).str.strip(), df_trans_prod["PL"]))
            slovnik_variant = dict(zip(df_trans_var["VAR"].astype(str).str.strip(), df_trans_var["PL"]))

            df_final["orderItemName_PL"] = df_final["orderItemName"].map(slovnik_produktů)
            df_final["orderItemVariantName_PL"] = df_final["orderItemVariantName"].map(slovnik_variant)
            df_final.loc[df_final["orderItemVariantName"] == "", "orderItemVariantName_PL"] = ""

            # 🛑 OCHRANA 1: Chybějící překlady
            chybejici_produkty = df_final[df_final["orderItemName_PL"].isna()]["orderItemName"].unique()
            chybejici_varianty = df_final[df_final["orderItemVariantName_PL"].isna() & (df_final["orderItemVariantName"] != "")]["orderItemVariantName"].unique()

            if len(chybejici_produkty) > 0 or len(chybejici_varianty) > 0:
                st.error("🛑 STOP: V překladovém souboru chybí data!")
                if len(chybejici_produkty) > 0:
                    st.write("**Doplň tyto PRODUKTY do 'DATA-PAWEL':**")
                    st.write(chybejici_produkty)
                if len(chybejici_varianty) > 0:
                    st.write("**Doplň tyto VARIANTY do 'VAR-PAWEL':**")
                    st.write(chybejici_varianty)
                st.stop()

            # Úprava + SZ
            df_final["orderItemName_PL"] = df_final["orderItemName_PL"].astype(str).str.strip()
            df_final["orderItemVariantName_PL"] = df_final["orderItemVariantName_PL"].astype(str).str.strip()
            obsahuje_sz_pl = df_final["orderItemVariantName_PL"].str.contains(r"\+?\s*SZ", case=False, na=False)
            df_final.loc[obsahuje_sz_pl & ~df_final["orderItemName_PL"].str.contains(r"\+?\s*SZ", case=False), "orderItemName_PL"] = (df_final["orderItemName_PL"] + " + SZ")
            df_final["orderItemVariantName_PL"] = df_final["orderItemVariantName_PL"].str.replace(r",?\s*\+?\s*SZ", "", case=False, regex=True).str.strip(", ").str.strip()

            # --- 5. POČTY BALÍKŮ (LABELS) ---
            slovnik_baliku = dict(zip(df_labels["NAME"].astype(str).str.strip(), df_labels["PCS"]))
            df_final["baliky_pcs"] = df_final["orderItemName_PL"].map(slovnik_baliku)

            # 🛑 OCHRANA 2: Chybějící balíky
            chybejici_definice_baliku = df_final[df_final["baliky_pcs"].isna()]["orderItemName_PL"].unique()
            if len(chybejici_definice_baliku) > 0:
                st.error("🛑 STOP: V záložce 'LABELS' chybí definice počtu balíků!")
                st.write("**Doplň tyto polské názvy do sloupce 'NAME':**")
                st.write(chybejici_definice_baliku)
                st.stop()

            df_final["baliky_pcs"] = df_final["baliky_pcs"].astype(int)

            # Řazení a tvorba řádků
            df_final = df_final.sort_values(by=["orderItemName_PL", "orderItemVariantName_PL"], ascending=True)
            
            def sestav_radek(row):
                kod, ks, nazev, var = str(row["code"]), f"{row['orderItemAmount']} szt.", row["orderItemName_PL"], row["orderItemVariantName_PL"]
                return f"{kod} - {ks} {nazev} ({var})" if var else f"{kod} - {ks} {nazev}"
            
            df_final["Radek_pro_dodavatele"] = df_final.apply(sestav_radek, axis=1)

            # --- ZOBRAZENÍ VÝSLEDKŮ NA WEBU ---
            st.subheader("📋 Přehled řádků pro e-mail dodavateli")
            st.code("\n".join(df_final["Radek_pro_dodavatele"].tolist()))

            # --- TLAČÍTKO PRO STAŽENÍ EXCELU ---
            excel_buffer = io.BytesIO()
            df_final.to_excel(excel_buffer, index=False)
            st.download_button(
                label="🟢 Stáhnout hotový Excel (vystup_kontrola.xlsx)",
                data=excel_buffer.getvalue(),
                file_name="vystup_kontrola.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # --- 6. GENERACE ŠTÍTKŮ DO PDF DO PAMĚTI ---
            pdf_buffer = io.BytesIO()
            doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, leftMargin=3*mm, rightMargin=3*mm, topMargin=0*mm, bottomMargin=0*mm)
            
            styles = getSampleStyleSheet()
            style_top_left = ParagraphStyle('TopLeft', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=10, leading=11)
            style_top_right = ParagraphStyle('TopRight', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=10, leading=11, alignment=2)
            style_main = ParagraphStyle('Main', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=11, leading=13, textColor=colors.black)
            style_pl = ParagraphStyle('PL', parent=styles['Normal'], fontName=FONT_ITALIC, fontSize=8.5, leading=10, textColor=colors.HexColor('#555555'))

            fraze_ke_smazani = ["Zvolte barvu:: ", "Zvolte rozměr:: ", "Zvolte variantu:: ", "Barva: ", "Rozměr: ",]
            bunky_stitku = []

            for _, row in df_final.iterrows():
                kod, cz_nazev, cz_var, pl_nazev = str(row["code"]), row["orderItemName"], row["orderItemVariantName"], row["orderItemName_PL"]
                mnozstvi_objednano, baliku_na_produkt = int(row["orderItemAmount"]), int(row["baliky_pcs"])
                
                for index_postele in range(1, mnozstvi_objednano + 1):
                    for aktualni_balik in range(1, baliku_na_produkt + 1):
                        cz_var_cista = cz_var
                        for fraze in fraze_ke_smazani:
                            import re
                            cz_var_cista = re.sub(re.escape(fraze), "", cz_var_cista, flags=re.IGNORECASE)
                        cz_var_cista = cz_var_cista.replace(", ,", ",").strip(", ").strip()
                        
                        cz_text = f"{cz_nazev} - {cz_var_cista}" if cz_var_cista else cz_nazev
                        pl_text = f"[{pl_nazev}]"
                        
                        horni_radek_data = [[Paragraph(f"OBJEDNÁVKA: {kod}", style_top_left), Paragraph(f"Balík: {aktualni_balik} z {baliku_na_produkt}", style_top_right)]]
                        horni_radek_tabulka = Table(horni_radek_data, colWidths=[48*mm, 44*mm])
                        horni_radek_tabulka.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0)]))
                        
                        obsah_stitku = [horni_radek_tabulka, Spacer(1, 1.5*mm), Paragraph(cz_text, style_main), Spacer(1, 1*mm), Paragraph(pl_text, style_pl)]
                        bunky_stitku.append(obsah_stitku)

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

            # --- TLAČÍTKO PRO STAŽENÍ PDF ---
            st.download_button(
                label="🔵 Stáhnout tiskové štítky PDF (stitky_k_tisku.pdf)",
                data=pdf_buffer.getvalue(),
                file_name="stitky_k_tisku.pdf",
                mime="application/pdf"
            )
            st.balloons() # Malá oslavná animace při úspěchu!

    except Exception as e:
        st.error(f"❌ Došlo k chybě při zpracování: {e}")
else:
    st.info("💡 Pro spuštění aplikace nahrajte v levém panelu minimálně soubor objednávek (E-shop 1) a soubor s překlady.")