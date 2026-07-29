# -*- coding: utf-8 -*-
"""
Fabrika Puantaj / Devam Analiz Sistemi
----------------------------------------
Streamlit + Pandas + Openpyxl + Plotly
"""

import io
import re
import random
import calendar
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# =========================================================================================
# 1) SABİTLER
# =========================================================================================

st.set_page_config(
    page_title="Fabrika Puantaj / Devam Analiz Sistemi",
    page_icon="🏭",
    layout="wide",
)

TURKCE_AYLAR = {
    "ocak": 1, "şubat": 2, "subat": 2, "mart": 3, "nisan": 4,
    "mayıs": 5, "mayis": 5, "haziran": 6, "temmuz": 7,
    "ağustos": 8, "agustos": 8, "eylül": 9, "eylul": 9,
    "ekim": 10, "kasım": 11, "kasim": 11, "aralık": 12, "aralik": 12,
}

RESMI_TATILLER_2026 = {
    (1, 1): "Yılbaşı",
    (3, 19): "Ramazan Bayramı Arifesi (Yarım Gün)",
    (3, 20): "Ramazan Bayramı 1. Gün",
    (3, 21): "Ramazan Bayramı 2. Gün",
    (3, 22): "Ramazan Bayramı 3. Gün",
    (4, 23): "23 Nisan Ulusal Egemenlik ve Çocuk Bayramı",
    (5, 1): "1 Mayıs Emek ve Dayanışma Günü",
    (5, 19): "19 Mayıs Atatürk'ü Anma Gençlik ve Spor Bayramı",
    (5, 26): "Kurban Bayramı Arifesi (Yarım Gün)",
    (5, 27): "Kurban Bayramı 1. Gün",
    (5, 28): "Kurban Bayramı 2. Gün",
    (5, 29): "Kurban Bayramı 3. Gün",
    (5, 30): "Kurban Bayramı 4. Gün",
    (7, 15): "15 Temmuz Demokrasi ve Millî Birlik Günü",
    (8, 30): "30 Ağustos Zafer Bayramı",
    (10, 28): "29 Cumhuriyet Bayramı Arifesi (Yarım Gün)",
    (10, 29): "29 Cumhuriyet Bayramı",
}

GELDI_SAYILAN_DURUMLAR = ["Geldi", "Kozmetik", "Ambalaj", "Refakatçi"]

DURUM_RENK = {
    "Geldi": "background-color:#d4edda",
    "Kozmetik": "background-color:#cce5ff",
    "Ambalaj": "background-color:#fff3cd",
    "Refakatçi": "background-color:#e2d6f3",
    "İstirahatli": "background-color:#f8d7da",
    "Ücretsiz İzin": "background-color:#f5c6cb",
    "Yıllık İzinli": "background-color:#d1ecf1",
    "Tatil": "background-color:#e9ecef; color:#6c757d",
    "Diğer": "background-color:#fce8b2",
}

# =========================================================================================
# 2) YARDIMCI FONKSİYONLAR
# =========================================================================================

def parse_ay_yil(sheet_name: str):
    s = tr_lower(sheet_name.strip())
    yil_match = re.search(r"(20\d{2})", s)
    yil = int(yil_match.group(1)) if yil_match else 2026

    ay_no = None
    for ay_adi, no in TURKCE_AYLAR.items():
        if ay_adi in s:
            ay_no = no
            break
    return ay_no, yil


def is_tatil_gunu(yil: int, ay: int, gun: int):
    try:
        d = date(yil, ay, gun)
    except ValueError:
        return False, None
    if d.weekday() == 6:  
        return True, "Pazar (Haftalık Tatil)"
    if (ay, gun) in RESMI_TATILLER_2026:
        return True, RESMI_TATILLER_2026[(ay, gun)]
    return False, None


def parse_saat(deger):
    if deger is None or isinstance(deger, bool):
        return None
    if isinstance(deger, (int, float)):
        if pd.isna(deger):
            return None
        return float(deger)
    s = str(deger).strip().replace(",", ".")
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def tr_lower(s: str) -> str:
    return s.replace("İ", "i").replace("I", "ı").lower()


def classify_durum(raw_value):
    if raw_value is None or (isinstance(raw_value, float) and pd.isna(raw_value)):
        return None, None

    s_raw = str(raw_value).strip()
    if s_raw == "" or s_raw.lower() == "nan":
        return None, None

    saat = parse_saat(raw_value)
    if saat is not None:
        return "Geldi", saat

    s = tr_lower(s_raw)

    if s in {"kozmetik", "kozmetık"}:
        return "Kozmetik", None
    if s in {"amb", "ambalaj"}:
        return "Ambalaj", None
    if s in {"refakatçi", "refakatci"}:
        return "Refakatçi", None
    if s == "geldi":
        return "Geldi", None
    if s in {"ist", "istirahatli", "rapor"}:
        return "İstirahatli", None
    if s in {"izin", "ücretsiz izin", "ucretsiz izin"}:
        return "Ücretsiz İzin", None
    if s in {"yıllık izin", "yıllık izinli", "yillik izin", "yillik izinli"}:
        return "Yıllık İzinli", None
    if s in {"resmi tatil", "pazar", "tatil", "arife"} or "arife" in s or "tatil" in s:
        return "Tatil", None

    return "Diğer", None


def load_excel_to_long(file_like) -> pd.DataFrame:
    xls = pd.ExcelFile(file_like)
    all_rows = []

    for sheet in xls.sheet_names:
        ay_no, yil = parse_ay_yil(sheet)
        if ay_no is None:
            continue

        df = xls.parse(sheet, header=None)
        gun_sayisi_ay = calendar.monthrange(yil, ay_no)[1]

        day_cols = []
        for col_idx in range(df.shape[1]):
            val = df.iloc[0, col_idx]
            try:
                d = int(float(str(val).strip()))
                if 1 <= d <= gun_sayisi_ay:
                    day_cols.append((col_idx, d))
            except (ValueError, TypeError):
                continue

        if not day_cols:
            for col_idx in range(df.shape[1]):
                for r_idx in [0, 1]:
                    val = df.iloc[r_idx, col_idx]
                    try:
                        d = int(float(str(val).strip()))
                        if 1 <= d <= gun_sayisi_ay:
                            if (col_idx, d) not in day_cols:
                                day_cols.append((col_idx, d))
                    except (ValueError, TypeError):
                        continue

        if not day_cols:
            continue

        for r_idx in range(len(df)):
            row = df.iloc[r_idx]
            col0 = row.iloc[0]
            col1 = row.iloc[1] if len(row) > 1 else None

            try:
                sicil_val = float(col0)
                if pd.isna(sicil_val) or sicil_val < 1000 or sicil_val > 99999:
                    continue
                sicil = str(int(sicil_val))
            except (ValueError, TypeError):
                continue

            ad = str(col1).strip() if pd.notna(col1) else ""
            soyad = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""

            if any(kw in ad.upper() for kw in ["AEROSOL", "NOT", "ÖNEMLİ", "PAZAR", "VARDİYA"]):
                continue

            ad_soyad = f"{ad} {soyad}".strip()
            if ad_soyad == "":
                ad_soyad = sicil

            for col_idx, gun_no in day_cols:
                raw_val = row.iloc[col_idx]
                tatil, tatil_sebep = is_tatil_gunu(yil, ay_no, gun_no)
                durum, saat = classify_durum(raw_val)

                if tatil:
                    durum_final = "Tatil"
                else:
                    durum_final = durum

                all_rows.append({
                    "Sicil No": sicil,
                    "Ad Soyad": ad_soyad,
                    "Ay": sheet.strip(),
                    "AyNo": ay_no,
                    "Yil": yil,
                    "Gun": gun_no,
                    "Tarih": date(yil, ay_no, gun_no),
                    "Durum": durum_final,
                    "Saat": saat,
                    "TatilSebep": tatil_sebep,
                })

    return pd.DataFrame(all_rows)


def aggregate_calisan_ay(long_df: pd.DataFrame) -> pd.DataFrame:
    if long_df.empty:
        return pd.DataFrame()

    agg = long_df.groupby(["Sicil No", "Ad Soyad", "Ay", "AyNo", "Yil"]).agg(
        Toplam_Takvim_Gunu=("Gun", "nunique"),
        Tatil_Gunu=("Durum", lambda s: (s == "Tatil").sum()),
        Geldi_Gunu=("Durum", lambda s: s.isin(GELDI_SAYILAN_DURUMLAR).sum()),
        Kozmetik_Gunu=("Durum", lambda s: (s == "Kozmetik").sum()),
        Ambalaj_Gunu=("Durum", lambda s: (s == "Ambalaj").sum()),
        Refakatci_Gunu=("Durum", lambda s: (s == "Refakatçi").sum()),
        Istirahat_Gunu=("Durum", lambda s: (s == "İstirahatli").sum()),
        Ucretsiz_Izin_Gunu=("Durum", lambda s: (s == "Ücretsiz İzin").sum()),
        Yillik_Izin_Gunu=("Durum", lambda s: (s == "Yıllık İzinli").sum()),
        Diger_Gunu=("Durum", lambda s: (s == "Diğer").sum()),
        Toplam_Saat=("Saat", "sum"),
    ).reset_index()

    agg["P_toplam"] = agg["Toplam_Takvim_Gunu"] - agg["Tatil_Gunu"]
    agg["Devam_Orani"] = np.where(
        agg["P_toplam"] > 0, 
        np.clip(agg["Geldi_Gunu"] / agg["P_toplam"] * 100, 0, 100), 
        np.nan
    )
    return agg


@st.cache_data(show_spinner="Excel dosyası işleniyor...")
def process_excel_bytes(file_bytes: bytes) -> pd.DataFrame:
    return load_excel_to_long(io.BytesIO(file_bytes))


# =========================================================================================
# 3) SIDEBAR
# =========================================================================================

st.sidebar.header("⚙️ Ayarlar & Hedefler")
uploaded_file = st.sidebar.file_uploader("📁 Puantaj Excel Dosyası Yükleyin (.xlsx)", type=["xlsx"])

hedef_calisan_sayisi = st.sidebar.number_input(
    "Gerekli / Hedef Kadro Çalışan Sayısı", min_value=1, value=90, step=1,
    help="Fabrikanın operasyonel olarak ihtiyaç duyduğu tam kadro sayısı."
)
hedef_devam_orani = st.sidebar.slider(
    "Hedef Devam Performans Oranı (%)", min_value=50, max_value=100, value=90,
    help="Bu oranın altına düşüldüğünde ek personel ihtiyacı doğar."
)

st.sidebar.markdown("---")
st.sidebar.markdown("Geldi sayılanlar: Geldi, Kozmetik, Ambalaj, Refakatçi.")
st.sidebar.markdown("Haftalık Tatil, gün sayımlarından hariç tutulur.")

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.getvalue()
        long_df = process_excel_bytes(file_bytes)
    except Exception as e:
        st.sidebar.error(f"Hata: {e}")
        long_df = pd.DataFrame()
else:
    try:
        with open("işçilik-2026.xlsx", "rb") as f:
            long_df = process_excel_bytes(f.read())
    except Exception:
        try:
            with open("işçilik-2026.xlsx", "rb") as f:
                long_df = process_excel_bytes(f.read())
        except Exception:
            long_df = pd.DataFrame()

if long_df.empty:
    st.title("🏭 Fabrika Puantaj / Devam Analiz Sistemi")
    st.warning("⚠️ Lütfen sol menüden Excel puantaj dosyanızı (.xlsx) yükleyin veya dosyanın klasörde olduğundan emin olun.")
    st.stop()

agg = aggregate_calisan_ay(long_df)
ay_sira = agg[["Ay", "AyNo"]].drop_duplicates().sort_values("AyNo")["Ay"].tolist()


# =========================================================================================
# 4) BAŞLIK VE AYLIK ÖZET METRİKLER
# =========================================================================================

st.title("🏭 Fabrika Puantaj / Devam Analiz Sistemi")

# Genel özet için ay seçme filtresi
secilen_genel_ay = st.selectbox("🗓️ Genel Özet & Metrikler İçin Ay Seçin", ["Tüm Yıl (Genel Toplam)"] + ay_sira)

if secilen_genel_ay == "Tüm Yıl (Genel Toplam)":
    filt_agg = agg
    filt_long = long_df
    donem_etiketi = "Tüm Yıl"
else:
    filt_agg = agg[agg["Ay"] == secilen_genel_ay]
    filt_long = long_df[long_df["Ay"] == secilen_genel_ay]
    donem_etiketi = secilen_genel_ay

toplam_calisan = filt_agg["Sicil No"].nunique()
genel_p_toplam = filt_agg["P_toplam"].sum()
genel_p_geldi = filt_agg["Geldi_Gunu"].sum()
fabrika_katilim_orani = (genel_p_geldi / genel_p_toplam * 100) if genel_p_toplam > 0 else np.nan
toplam_rapor = filt_agg["Istirahat_Gunu"].sum()

st.caption(f"📌 Seçilen Dönem: **{donem_etiketi}** | Toplam kayıt satırı: {len(filt_long):,}")

col1, col2, col3 = st.columns(3)
col1.metric("👥 Toplam Aktif Çalışan", f"{toplam_calisan}", help=f"Hedef Kadro: {hedef_calisan_sayisi}")
col2.metric("✅ Fabrika Geneli Katılım Oranı", f"%{fabrika_katilim_orani:.1f}" if not np.isnan(fabrika_katilim_orani) else "N/A", help=f"Hedef: %{hedef_devam_orani}")
col3.metric("🩺 Toplam Rapor/İstirahat Günü", f"{int(toplam_rapor)}")

# --- Ek Personel / Takviye İhtiyacı Analizi ---
st.markdown("### 📋 Ek Personel & Takviye İhtiyacı Analizi")
if not np.isnan(fabrika_katilim_orani):
    # Ortalama günlük gelen kişi sayısı vs hedef çalışan üzerinden açık hesaplama
    ortalama_gunluk_gelen = genel_p_geldi / filt_agg["P_toplam"].max() if filt_agg["P_toplam"].max() > 0 else 0
    gerekli_etkin_calisan = hedef_calisan_sayisi * (hedef_devam_orani / 100.0)
    
    if fabrika_katilim_orani < hedef_devam_orani or toplam_calisan < hedef_calisan_sayisi:
        eksik_oran = hedef_devam_orani - fabrika_katilim_orani
        tahmini_ek_kisi = max(int(np.ceil((hedef_calisan_sayisi - toplam_calisan) + (eksik_oran / 100 * hedef_calisan_sayisi))), 1)
        st.warning(
            f"⚠️ **EK PERSONEL İHTİYACI VAR!** ({donem_etiketi})\n\n"
            f"* **Mevcut Durum:** Aktif çalışan sayısı {toplam_calisan} (Hedef: {hedef_calisan_sayisi}), Katılım Oranı: %{fabrika_katilim_orani:.1f} (Hedef: %{hedef_devam_orani}).\n"
            f"* **Öneri:** Raporlar ve izinler nedeniyle düşen iş gücünü kompanse etmek için **yaklaşık {tahmini_ek_kisi} ek personel (takviye işçi)** görevlendirilmesi önerilir."
        )
    else:
        st.success(
            f"✅ **EK PERSONEL İHTİYACI YOK.** ({donem_etiketi})\n\n"
            f"* Mevcut kadro ({toplam_calisan} çalışan) ve katılım oranımız (%{fabrika_katilim_orani:.1f}), belirlenen hedef performans eşiğinin (%{hedef_devam_orani}) üzerindedir."
        )

st.divider()


# =========================================================================================
# 5) GRAFİKLER (AYLIK FİLTRELİ)
# =========================================================================================

st.subheader("🏆 En Yüksek & En Düşük Performanslı Çalışanlar")

secilen_grafik_ayi = st.selectbox("📅 Sıralama Grafikleri İçin Ay Seçin", ["Tüm Yıl Ortalama"] + ay_sira)

if secilen_grafik_ayi == "Tüm Yıl Ortalama":
    kisi_ozet = agg.groupby(["Sicil No", "Ad Soyad"]).agg(
        P_toplam=("P_toplam", "sum"),
        Geldi_Gunu=("Geldi_Gunu", "sum"),
        Istirahat_Gunu=("Istirahat_Gunu", "sum"),
        Ucretsiz_Izin_Gunu=("Ucretsiz_Izin_Gunu", "sum"),
        Yillik_Izin_Gunu=("Yillik_Izin_Gunu", "sum"),
    ).reset_index()
    baslik_eki = "(Tüm Yıl Ortalaması)"
else:
    kisi_ozet = agg[agg["Ay"] == secilen_grafik_ayi].copy()
    baslik_eki = f"({secilen_grafik_ayi})"

kisi_ozet["Devam_Orani"] = np.where(
    kisi_ozet["P_toplam"] > 0, 
    np.clip(kisi_ozet["Geldi_Gunu"] / kisi_ozet["P_toplam"] * 100, 0, 100), 
    np.nan
)

top10 = kisi_ozet.dropna(subset=["Devam_Orani"]).sort_values("Devam_Orani", ascending=False).head(10)
bottom10 = kisi_ozet.dropna(subset=["Devam_Orani"]).sort_values("Devam_Orani", ascending=True).head(10)

col_a, col_b = st.columns(2)
with col_a:
    fig_top = px.bar(
        top10.sort_values("Devam_Orani"), x="Devam_Orani", y="Ad Soyad", orientation="h",
        title=f"Top 10 — En Yüksek Katılım {baslik_eki}", text="Devam_Orani",
        color_discrete_sequence=["#2E8B57"],
    )
    fig_top.update_traces(texttemplate="%{text:.1f}%")
    fig_top.update_layout(xaxis_title="Katılım Oranı (%)", yaxis_title="")
    st.plotly_chart(fig_top, use_container_width=True)

with col_b:
    fig_bottom = px.bar(
        bottom10.sort_values("Devam_Orani", ascending=False), x="Devam_Orani", y="Ad Soyad", orientation="h",
        title=f"Bottom 10 — En Düşük Katılım {baslik_eki}", text="Devam_Orani",
        color_discrete_sequence=["#C0392B"],
    )
    fig_bottom.update_traces(texttemplate="%{text:.1f}%")
    fig_bottom.update_layout(xaxis_title="Katılım Oranı (%)", yaxis_title="")
    st.plotly_chart(fig_bottom, use_container_width=True)

ay_ozet = agg.groupby(["Ay", "AyNo"]).agg(
    Istirahat_Gunu=("Istirahat_Gunu", "sum"),
    P_toplam=("P_toplam", "sum"),
    Geldi_Gunu=("Geldi_Gunu", "sum"),
).reset_index().sort_values("AyNo")

ay_ozet["Devam_Orani"] = np.where(
    ay_ozet["P_toplam"] > 0, 
    np.clip(ay_ozet["Geldi_Gunu"] / ay_ozet["P_toplam"] * 100, 0, 100), 
    np.nan
)

col_c, col_d = st.columns(2)
with col_c:
    fig_rapor = px.bar(
        ay_ozet, x="Ay", y="Istirahat_Gunu", title="🩺 Aylara Göre Rapor / İstirahat Trendi",
        text="Istirahat_Gunu", color_discrete_sequence=["#5DADE2"],
    )
    fig_rapor.update_layout(xaxis_title="", yaxis_title="İstirahat/Rapor Gün Sayısı")
    st.plotly_chart(fig_rapor, use_container_width=True)

with col_d:
    fig_devam = px.line(
        ay_ozet, x="Ay", y="Devam_Orani", markers=True,
        title="📉 Aylara Göre Fabrika Geneli Devam Oranı Trendi",
    )
    fig_devam.update_traces(line_color="#8E44AD")
    fig_devam.update_layout(xaxis_title="", yaxis_title="Devam Oranı (%)")
    st.plotly_chart(fig_devam, use_container_width=True)

st.divider()

# =========================================================================================
# 6) BİREYSEL PERSONEL SORGULAMA
# =========================================================================================

st.header("👤 Bireysel Personel Sorgulama")

personel_liste = agg[["Sicil No", "Ad Soyad"]].drop_duplicates().copy()
personel_liste["Etiket"] = personel_liste["Sicil No"].astype(str) + " - " + personel_liste["Ad Soyad"]
personel_liste = personel_liste.sort_values("Ad Soyad")

col_f1, col_f2 = st.columns(2)
with col_f1:
    secili_etiket = st.selectbox(
        "🔍 Personel Seçimi",
        options=personel_liste["Etiket"].tolist(),
    )
with col_f2:
    secili_ay = st.selectbox("🗓️ Ay Seçimi", ["Tüm Aylar"] + ay_sira)

secili_sicil = personel_liste.loc[personel_liste["Etiket"] == secili_etiket, "Sicil No"].iloc[0]

if secili_ay == "Tüm Aylar":
    kisi_agg = agg[agg["Sicil No"] == secili_sicil]
    kisi_long = long_df[long_df["Sicil No"] == secili_sicil]
else:
    kisi_agg = agg[(agg["Sicil No"] == secili_sicil) & (agg["Ay"] == secili_ay)]
    kisi_long = long_df[(long_df["Sicil No"] == secili_sicil) & (long_df["Ay"] == secili_ay)]

p_toplam_k = kisi_agg["P_toplam"].sum()
geldi_k = kisi_agg["Geldi_Gunu"].sum()
istirahat_k = kisi_agg["Istirahat_Gunu"].sum()
ucretsiz_k = kisi_agg["Ucretsiz_Izin_Gunu"].sum()
yillik_k = kisi_agg["Yillik_Izin_Gunu"].sum()
devam_orani_k = (geldi_k / p_toplam_k * 100) if p_toplam_k > 0 else np.nan

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("✅ Geldiği Gün", f"{int(geldi_k)}")
kpi2.metric("🩺 İstirahat / Rapor", f"{int(istirahat_k)}")
kpi3.metric("🌴 Ücretsiz / Yıllık İzin", f"{int(ucretsiz_k + yillik_k)}")
kpi4.metric("📈 Devam Orani", f"%{devam_orani_k:.1f}" if not np.isnan(devam_orani_k) else "N/A")

detay = kisi_long[["AyNo", "Ay", "Gun", "Tarih", "Durum", "Saat", "TatilSebep"]].sort_values(["AyNo", "Gun"]).copy()
detay = detay.drop(columns=["AyNo"]).rename(columns={
    "Gun": "Gün", "Tarih": "Tarih", "Durum": "Durum", "Saat": "Çalışılan Saat", "TatilSebep": "Tatil Açıklaması",
})

st.dataframe(detay, use_container_width=True, height=420)
