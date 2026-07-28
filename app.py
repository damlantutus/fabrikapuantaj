# -*- coding: utf-8 -*-
"""
Fabrika Puantaj / Devam Analiz Sistemi
Streamlit + Pandas + Openpyxl + Plotly
"""

import io
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ----------------------------------------------------------------------------
# SABİTLER
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Puantaj Analiz Sistemi", layout="wide", page_icon="🏭")

AY_SIRA = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz",
           "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]

GELDI_SAYILAN = {"Geldi", "Kozmetik", "Ambalaj", "Refakatçi"}

STATUS_MAP = {
    "kozmetik": "Kozmetik", "kozmetık": "Kozmetik",
    "amb": "Ambalaj", "ambalaj": "Ambalaj",
    "geldi": "Geldi",
    "refakatçi": "Refakatçi", "refakatci": "Refakatçi",
    "istirahatli": "İstirahatli", "i̇stirahatli": "İstirahatli", "rapor": "İstirahatli",
    "haftalık tatil": "Haftalık Tatil", "haftalik tatil": "Haftalık Tatil",
    "ücretsiz izin": "Ücretsiz İzin", "ucretsiz izin": "Ücretsiz İzin",
    "yıllık izinli": "Yıllık İzinli", "yillik izinli": "Yıllık İzinli",
    "resmi tatil": "Resmi Tatil",
}

STATUS_COLORS = {
    "Geldi": "background-color: #c6efce",
    "Kozmetik": "background-color: #bee7e8",
    "Ambalaj": "background-color: #bee7e8",
    "Refakatçi": "background-color: #bee7e8",
    "Haftalık Tatil": "background-color: #e2e3e5",
    "İstirahatli": "background-color: #f8c9c9",
    "Ücretsiz İzin": "background-color: #ffe8a1",
    "Yıllık İzinli": "background-color: #ffe8a1",
    "Resmi Tatil": "background-color: #cfe2ff",
}


# ----------------------------------------------------------------------------
# NORMALİZASYON
# ----------------------------------------------------------------------------
def normalize_status(val):
    if pd.isna(val):
        return None
    s = str(val).strip().lower()
    if s == "" or s == "nan" or s == "nat":
        return None
    if s in STATUS_MAP:
        return STATUS_MAP[s]
    if "bayram" in s or "resmi tatil" in s or "tatili" in s:
        return "Resmi Tatil"
    # Bilinmeyen / tanımsız bir metin ise, orijinal halini (title-case) koru
    return str(val).strip()


# ----------------------------------------------------------------------------
# ETL: EXCEL -> UZUN FORMAT DATAFRAME
# ----------------------------------------------------------------------------
def load_and_transform(uploaded_file):
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheet_names = [s for s in xls.sheet_names if s.strip().lower() not in ("özet", "ozet")]

    all_frames = []
    for sheet in sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet)
        raw.columns = [str(c).strip() for c in raw.columns]

        if raw.shape[1] < 3:
            continue

        id_col = raw.columns[0]
        name_col = raw.columns[1]

        # Son 2 sütun özet sütunları (Ücretsiz İzin (gün) / Oranı) ise hariç tut
        last_cols = [str(c).lower() for c in raw.columns[-2:]]
        if len(raw.columns) > 4 and any("izin" in c for c in last_cols):
            day_cols = list(raw.columns[2:-2])
        else:
            day_cols = list(raw.columns[2:])

        if not day_cols:
            continue

        melted = raw.melt(
            id_vars=[id_col, name_col],
            value_vars=day_cols,
            var_name="Gün/Tarih",
            value_name="Durum",
        )
        melted = melted.rename(columns={id_col: "Sicil No", name_col: "Ad Soyad"})
        melted["Ay"] = sheet.strip()

        # Sıralama için orijinal sütun konumunu sakla (gün sırasını korumak için)
        order_map = {col: i for i, col in enumerate(day_cols)}
        melted["Sıra"] = melted["Gün/Tarih"].map(order_map)

        all_frames.append(melted)

    if not all_frames:
        raise ValueError("Excel dosyasında işlenebilir ay sekmesi bulunamadı.")

    df = pd.concat(all_frames, ignore_index=True)
    df["Sicil No"] = df["Sicil No"].astype(str).str.strip()
    df["Ad Soyad"] = df["Ad Soyad"].astype(str).str.strip()
    df["Gün/Tarih"] = df["Gün/Tarih"].astype(str).str.strip()
    df["Durum_Norm"] = df["Durum"].apply(normalize_status)
    df = df.dropna(subset=["Durum_Norm"]).reset_index(drop=True)
    df = df[df["Sicil No"].str.lower() != "nan"]
    return df


# ----------------------------------------------------------------------------
# MOCK VERİ ÜRETİMİ
# ----------------------------------------------------------------------------
@st.cache_data
def generate_mock_data():
    np.random.seed(42)
    months = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz"]
    days_in_month = {"Ocak": 31, "Şubat": 28, "Mart": 31, "Nisan": 30,
                      "Mayıs": 31, "Haziran": 30, "Temmuz": 31}

    isimler = [
        "Ahmet Yılmaz", "Mehmet Demir", "Ayşe Kaya", "Fatma Şahin", "Mustafa Çelik",
        "Emine Yıldız", "Hüseyin Aydın", "Zeynep Arslan", "Ali Doğan", "Hatice Kılıç",
        "İbrahim Aslan", "Elif Çetin", "Hasan Kurt", "Sultan Özdemir", "Kemal Koç",
        "Aylin Kaplan", "Serkan Er", "Derya Şen", "Bora Aksoy", "Nur Polat",
    ]
    employees = [(str(1000 + i), ad) for i, ad in enumerate(isimler)]

    pool_normal = (["Geldi"] * 70 + ["Haftalık Tatil"] * 15 + ["İstirahatli"] * 4 +
                   ["Ücretsiz İzin"] * 3 + ["Yıllık İzinli"] * 3 + ["Kozmetik"] * 2 +
                   ["Ambalaj"] * 2 + ["Refakatçi"] * 1)
    pool_yuksek = ["Geldi"] * 85 + ["Haftalık Tatil"] * 13 + ["İstirahatli"] * 1 + ["Kozmetik"] * 1
    pool_dusuk = (["Geldi"] * 55 + ["Haftalık Tatil"] * 13 + ["İstirahatli"] * 15 +
                  ["Ücretsiz İzin"] * 10 + ["Yıllık İzinli"] * 5 + ["Ambalaj"] * 2)

    rows = []
    for sicil, ad in employees:
        bias = np.random.choice(["yuksek", "orta", "dusuk"], p=[0.3, 0.5, 0.2])
        pool = pool_yuksek if bias == "yuksek" else (pool_dusuk if bias == "dusuk" else pool_normal)
        for ay in months:
            n = days_in_month[ay]
            for gun in range(1, n + 1):
                durum = np.random.choice(pool)
                rows.append({
                    "Sicil No": sicil, "Ad Soyad": ad, "Ay": ay,
                    "Gün/Tarih": str(gun), "Sıra": gun, "Durum": durum,
                })

    df = pd.DataFrame(rows)
    df["Durum_Norm"] = df["Durum"].apply(normalize_status)
    df = df.dropna(subset=["Durum_Norm"]).reset_index(drop=True)
    return df


# ----------------------------------------------------------------------------
# METRİK HESAPLAMALARI
# ----------------------------------------------------------------------------
def calc_stats(df: pd.DataFrame) -> dict:
    if df.empty:
        return dict(toplam_kayitli=0, haftalik_tatil=0, p_toplam=0, geldi=0,
                    istirahat=0, ucretsiz_izin=0, yillik_izin=0, resmi_tatil=0,
                    standart=0, kozmetik=0, ambalaj=0, refakatci=0, katilim_orani=0.0)

    toplam_kayitli = len(df)
    haftalik = int((df["Durum_Norm"] == "Haftalık Tatil").sum())
    p_toplam = toplam_kayitli - haftalik
    geldi = int(df["Durum_Norm"].isin(GELDI_SAYILAN).sum())
    istirahat = int((df["Durum_Norm"] == "İstirahatli").sum())
    ucretsiz = int((df["Durum_Norm"] == "Ücretsiz İzin").sum())
    yillik = int((df["Durum_Norm"] == "Yıllık İzinli").sum())
    resmi = int((df["Durum_Norm"] == "Resmi Tatil").sum())
    standart = int((df["Durum_Norm"] == "Geldi").sum())
    kozmetik = int((df["Durum_Norm"] == "Kozmetik").sum())
    ambalaj = int((df["Durum_Norm"] == "Ambalaj").sum())
    refakatci = int((df["Durum_Norm"] == "Refakatçi").sum())
    oran = round(geldi / p_toplam * 100, 2) if p_toplam > 0 else 0.0

    return dict(toplam_kayitli=toplam_kayitli, haftalik_tatil=haftalik, p_toplam=p_toplam,
                geldi=geldi, istirahat=istirahat, ucretsiz_izin=ucretsiz, yillik_izin=yillik,
                resmi_tatil=resmi, standart=standart, kozmetik=kozmetik, ambalaj=ambalaj,
                refakatci=refakatci, katilim_orani=oran)


def per_employee_stats(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for (sicil, ad), g in df.groupby(["Sicil No", "Ad Soyad"]):
        s = calc_stats(g)
        s["Sicil No"] = sicil
        s["Ad Soyad"] = ad
        rows.append(s)
    return pd.DataFrame(rows)


def style_durum(val):
    return STATUS_COLORS.get(val, "")


# ----------------------------------------------------------------------------
# VERİ YÜKLEME (SIDEBAR)
# ----------------------------------------------------------------------------
st.sidebar.title("🏭 Puantaj Analiz Sistemi")
st.sidebar.markdown("---")
uploaded_file = st.sidebar.file_uploader(
    "Puantaj Excel Dosyasını Yükleyin (.xlsx)", type=["xlsx"]
)

data_source = "mock"
df = None
error_msg = None

if uploaded_file is not None:
    try:
        df = load_and_transform(uploaded_file)
        data_source = "excel"
    except Exception as e:
        error_msg = str(e)

if df is None:
    df = generate_mock_data()
    data_source = "mock"

if data_source == "excel":
    st.sidebar.success(f"✅ Excel yüklendi: {df['Ay'].nunique()} ay, {df['Sicil No'].nunique()} personel")
else:
    if error_msg:
        st.sidebar.error(f"Excel işlenemedi, örnek veri kullanılıyor.\n\nHata: {error_msg}")
    else:
        st.sidebar.info("📊 Dosya yüklenmedi — örnek (mock) veri gösteriliyor.")

st.sidebar.markdown("---")
st.sidebar.caption(
    "Geldi sayılanlar: Geldi, Kozmetik, Ambalaj, Refakatçi.\n\n"
    "Haftalık Tatil, gün sayımlarından hariç tutulur."
)

# Ay sırası (veride bulunanlara göre)
mevcut_aylar = [a for a in AY_SIRA if a in df["Ay"].unique()]
diger_aylar = [a for a in df["Ay"].unique() if a not in AY_SIRA]
mevcut_aylar = mevcut_aylar + diger_aylar

# ----------------------------------------------------------------------------
# BAŞLIK
# ----------------------------------------------------------------------------
st.title("🏭 Fabrika Puantaj / Devam Analiz Sistemi")

tab1, tab2 = st.tabs(["👤 Bireysel Personel Sorgulama", "🏭 Fabrika Geneli Dashboard"])

# ============================================================================
# SEKME 1 — BİREYSEL PERSONEL SORGULAMA
# ============================================================================
with tab1:
    personel_listesi = (
        df[["Sicil No", "Ad Soyad"]]
        .drop_duplicates()
        .sort_values("Ad Soyad")
    )
    personel_listesi["Etiket"] = personel_listesi["Sicil No"] + " — " + personel_listesi["Ad Soyad"]

    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        secili_etiket = st.selectbox(
            "Personel Seçimi (Sicil No veya Ad Soyad ile arayabilirsiniz)",
            options=personel_listesi["Etiket"].tolist(),
        )
    with col_f2:
        ay_secimi = st.selectbox("Ay Seçimi", options=["Tüm Aylar"] + mevcut_aylar)

    secili_sicil = secili_etiket.split(" — ")[0]

    df_person = df[df["Sicil No"] == secili_sicil].copy()
    if ay_secimi != "Tüm Aylar":
        df_person = df_person[df_person["Ay"] == ay_secimi]

    stats = calc_stats(df_person)

    st.markdown("### 📌 Performans Özeti")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric(
        "Geldiği Gün Sayısı",
        stats["geldi"],
        help=(f"Standart Geldi: {stats['standart']} | Kozmetik: {stats['kozmetik']} | "
              f"Ambalaj: {stats['ambalaj']} | Refakatçi: {stats['refakatci']}"),
    )
    k2.metric("İstirahat / Rapor Gün Sayısı", stats["istirahat"])
    k3.metric(
        "Ücretsiz / Yıllık İzin Sayısı",
        stats["ucretsiz_izin"] + stats["yillik_izin"],
        help=f"Ücretsiz İzin: {stats['ucretsiz_izin']} | Yıllık İzin: {stats['yillik_izin']}",
    )
    k4.metric("Devam Performans Oranı (%)", f"%{stats['katilim_orani']}")

    st.caption(
        f"Toplam Kayıtlı Gün: {stats['toplam_kayitli']} | Haftalık Tatil (hariç): {stats['haftalik_tatil']} "
        f"| Net İş Günü (P_toplam): {stats['p_toplam']}"
    )

    st.markdown("---")
    st.markdown("### 📅 Çalışma Detay Kartı")

    col_d1, col_d2 = st.columns([1.4, 1])

    with col_d1:
        st.markdown("**Gün Gün Katılım Durumu**")
        gorunum = (
            df_person.rename(columns={"Durum_Norm": "Durum"})
            [["Ay", "Gün/Tarih", "Durum"]]
        )
        if "Sıra" in df_person.columns:
            gorunum["Sıra"] = df_person["Sıra"].values
            gorunum["_ay_sira"] = gorunum["Ay"].apply(lambda a: mevcut_aylar.index(a) if a in mevcut_aylar else 99)
            gorunum = gorunum.sort_values(["_ay_sira", "Sıra"]).drop(columns=["_ay_sira", "Sıra"])

        if gorunum.empty:
            st.info("Seçilen periyot için veri bulunamadı.")
        else:
            styled = gorunum.style.applymap(style_durum, subset=["Durum"])
            st.dataframe(styled, use_container_width=True, hide_index=True, height=420)

    with col_d2:
        st.markdown("**Özel Çalışma Türü Dağılımı**")
        dagilim = pd.DataFrame({
            "Tür": ["Standart Geldi", "Kozmetik", "Ambalaj", "Refakatçi"],
            "Gün Sayısı": [stats["standart"], stats["kozmetik"], stats["ambalaj"], stats["refakatci"]],
        })
        dagilim = dagilim[dagilim["Gün Sayısı"] > 0]
        if dagilim.empty:
            st.info("Bu periyotta özel çalışma türü kaydı yok.")
        else:
            fig_dagilim = px.bar(
                dagilim, x="Tür", y="Gün Sayısı", color="Tür", text="Gün Sayısı",
                title="Çalışma Türü Kırılımı",
            )
            fig_dagilim.update_layout(showlegend=False)
            st.plotly_chart(fig_dagilim, use_container_width=True)

# ============================================================================
# SEKME 2 — FABRİKA GENELİ DASHBOARD
# ============================================================================
with tab2:
    st.markdown("### 🏭 Fabrika Geneli Özet")

    ay_secimi_dash = st.multiselect(
        "Ay Filtresi (boş bırakılırsa tüm aylar dahil edilir)",
        options=mevcut_aylar, default=[],
    )
    df_dash = df if not ay_secimi_dash else df[df["Ay"].isin(ay_secimi_dash)]

    genel_stats = calc_stats(df_dash)
    toplam_calisan = df_dash["Sicil No"].nunique()

    g1, g2, g3 = st.columns(3)
    g1.metric("Toplam Aktif Çalışan Sayısı", toplam_calisan)
    g2.metric("Fabrika Geneli Katılım Oranı (%)", f"%{genel_stats['katilim_orani']}")
    g3.metric("Toplam Alınan Rapor / İstirahat Günü", genel_stats["istirahat"])

    st.markdown("---")

    emp_stats = per_employee_stats(df_dash)

    st.markdown("#### 🏆 En Yüksek & En Düşük Performanslı Çalışanlar")
    col_top, col_bottom = st.columns(2)

    if not emp_stats.empty:
        emp_sirali = emp_stats.sort_values("katilim_orani", ascending=False)
        top10 = emp_sirali.head(10).sort_values("katilim_orani")
        bottom10 = emp_sirali.tail(10).sort_values("katilim_orani")

        with col_top:
            fig_top = px.bar(
                top10, x="katilim_orani", y="Ad Soyad", orientation="h",
                text="katilim_orani", title="Top 10 — En Yüksek Katılım Oranı (%)",
                labels={"katilim_orani": "Katılım Oranı (%)", "Ad Soyad": ""},
                color_discrete_sequence=["#2ca02c"],
            )
            st.plotly_chart(fig_top, use_container_width=True)

        with col_bottom:
            fig_bottom = px.bar(
                bottom10, x="katilim_orani", y="Ad Soyad", orientation="h",
                text="katilim_orani", title="Bottom 10 — En Düşük Katılım Oranı (%)",
                labels={"katilim_orani": "Katılım Oranı (%)", "Ad Soyad": ""},
                color_discrete_sequence=["#d62728"],
            )
            st.plotly_chart(fig_bottom, use_container_width=True)
    else:
        st.info("Seçilen filtrelere uygun veri bulunamadı.")

    st.markdown("---")
    st.markdown("#### 📈 Aylara Göre Rapor / İstirahat Trendi")

    aylik_trend = (
        df_dash[df_dash["Durum_Norm"] == "İstirahatli"]
        .groupby("Ay").size().reindex(mevcut_aylar).fillna(0).reset_index(name="İstirahat Gün Sayısı")
    )
    fig_trend = px.line(
        aylik_trend, x="Ay", y="İstirahat Gün Sayısı", markers=True,
        title="Aylık İstirahat / Rapor Gün Sayısı Trendi",
    )
    st.plotly_chart(fig_trend, use_container_width=True)

    st.markdown("---")
    st.markdown("#### 🥯 Devamsızlık ve İzin Dağılımı")

    dagilim_pie = pd.DataFrame({
        "Kategori": ["İstirahatli", "Ücretsiz İzin", "Yıllık İzinli", "Resmi Tatil"],
        "Gün Sayısı": [genel_stats["istirahat"], genel_stats["ucretsiz_izin"],
                        genel_stats["yillik_izin"], genel_stats["resmi_tatil"]],
    })
    dagilim_pie = dagilim_pie[dagilim_pie["Gün Sayısı"] > 0]

    if dagilim_pie.empty:
        st.info("Bu periyotta devamsızlık/izin kaydı bulunamadı.")
    else:
        fig_pie = px.pie(
            dagilim_pie, names="Kategori", values="Gün Sayısı", hole=0.5,
            title="Devamsızlık / İzin Türü Dağılımı",
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("---")
    with st.expander("📋 Detaylı Personel Tablosu"):
        if not emp_stats.empty:
            gosterim = emp_stats.rename(columns={
                "Sicil No": "Sicil No", "Ad Soyad": "Ad Soyad",
                "geldi": "Geldi", "istirahat": "İstirahatli",
                "ucretsiz_izin": "Ücretsiz İzin", "yillik_izin": "Yıllık İzinli",
                "p_toplam": "Net İş Günü", "katilim_orani": "Katılım Oranı (%)",
            })[["Sicil No", "Ad Soyad", "Geldi", "İstirahatli", "Ücretsiz İzin",
                "Yıllık İzinli", "Net İş Günü", "Katılım Oranı (%)"]]
            st.dataframe(gosterim.sort_values("Katılım Oranı (%)", ascending=False),
                         use_container_width=True, hide_index=True)
