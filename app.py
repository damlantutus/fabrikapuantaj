# -*- coding: utf-8 -*-
"""
Fabrika Puantaj / Devam Analiz Sistemi
----------------------------------------
Streamlit + Pandas + Openpyxl + Plotly

Çalıştırmak için:
    pip install streamlit pandas openpyxl plotly numpy
    streamlit run app.py
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

# 2026 yılı resmi tatil ve arife günleri (Diyanet/Resmi Gazete duyurularına göre tahmini/onaylı tarihler).
# Gerekirse bu sözlüğü güncelleyerek kendi fabrikanızın resmi tatil takvimine uyarlayabilirsiniz.
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
    (10, 28): "29 Ekim Cumhuriyet Bayramı Arifesi (Yarım Gün)",
    (10, 29): "29 Ekim Cumhuriyet Bayramı",
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
# 2) YARDIMCI FONKSİYONLAR (ETL / SINIFLANDIRMA)
# =========================================================================================

def parse_ay_yil(sheet_name: str):
    """Sekme isminden (örn. 'OCAK-2026') ay numarasını ve yılı çıkarır."""
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
    """Verilen tarih Pazar günü mü ya da resmi tatil/arife mi kontrol eder."""
    try:
        d = date(yil, ay, gun)
    except ValueError:
        return False, None
    if d.weekday() == 6:  # Pazartesi=0 ... Pazar=6
        return True, "Pazar (Haftalık Tatil)"
    if (ay, gun) in RESMI_TATILLER_2026:
        return True, RESMI_TATILLER_2026[(ay, gun)]
    return False, None


def parse_saat(deger):
    """Hücre değerini sayısal saate çevirmeye çalışır. Başarısızsa None döner."""
    if deger is None:
        return None
    if isinstance(deger, bool):
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
    """Türkçe'ye özgü İ/I harflerini doğru şekilde küçük harfe çevirir.
    Standart Python .lower() metodu 'İ' harfini 'i̇' (noktalı bileşik) karaktere
    çevirdiğinden, "KOZMETİK" gibi büyük harfli Türkçe girdilerin eşleşmesini bozar."""
    return s.replace("İ", "i").replace("I", "ı").lower()


def classify_durum(raw_value):
    """
    Ham hücre değerini standart bir Durum etiketine çevirir.
    Dönüş: (durum_str veya None, saat_degeri veya None)
    """
    if raw_value is None or (isinstance(raw_value, float) and pd.isna(raw_value)):
        return None, None

    s_raw = str(raw_value).strip()
    if s_raw == "" or s_raw.lower() == "nan":
        return None, None

    # 1) Sayısal saat kontrolü -> kaç saat olursa olsun "Geldi"
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


def gecerli_personel_mi(sicil: str, ad_soyad: str) -> bool:
    """
    Puantaj tablosundaki gerçek personel satırlarını, "SAYIM GÜNÜ",
    "1 MAYIS İŞÇİ BAYRAMI" gibi personel olmayan açıklama/etkinlik satırlarından ayırt eder.
    """
    # 1) Sicil No sayısal olmalı (gerçek personel sicil numaraları sayısaldır).
    try:
        float(str(sicil).strip().replace(",", "."))
    except (ValueError, TypeError):
        return False

    # 2) Ad Soyad, sicil no'nun fallback'i değilse (yani gerçek bir isim ise) rakam içermemeli.
    #    ("1 MAYIS İŞÇİ BAYRAMI" gibi tarih içeren notlar bu şekilde elenir.)
    if ad_soyad != sicil and any(ch.isdigit() for ch in str(ad_soyad)):
        return False

    # 3) Bilinen özel gün / sayım notu anahtar kelimeleri.
    ozel_kelimeler = {"bayram", "sayım", "sayim", "tatil", "arife", "kapanış", "kapanis"}
    ad_soyad_kucuk = tr_lower(str(ad_soyad))
    if any(kelime in ad_soyad_kucuk for kelime in ozel_kelimeler):
        return False

    return True


def load_excel_to_long(file_like) -> pd.DataFrame:
    """Yüklenen Excel dosyasını okuyup uzun (long) formatlı DataFrame'e dönüştürür."""
    xls = pd.ExcelFile(file_like)
    all_rows = []

    for sheet in xls.sheet_names:
        ay_no, yil = parse_ay_yil(sheet)
        if ay_no is None:
            # Ay sekmesi olarak tanınamayan sekmeler (örn. "Özet") atlanır.
            continue

        df = xls.parse(sheet, dtype=object)
        df.columns = [str(c).strip() for c in df.columns]
        cols = list(df.columns)

        # İlk sütun her zaman Sicil No kabul edilir (başlığı ne olursa olsun: "Sicil No", "Unnamed: 0" vb.)
        sicil_col = cols[0]

        gun_sayisi_ay = calendar.monthrange(yil, ay_no)[1]

        # Gün sütunlarını tespit et: başlığı 1..ayın_gun_sayisi arasında bir tamsayıya çevrilebilen sütunlar.
        #
        # ÖNEMLİ HATA DÜZELTMESİ: Excel'de yinelenen sütun başlıkları varsa (örn. iki adet "1" günü
        # sütunu), pandas bunları otomatik olarak "1.1", "2.1" gibi ondalıklı adlarla yeniden adlandırır.
        # Eskiden int(float("1.1")) = 1 olduğu için bu sütun yanlışlıkla "1. gün" ile eşleşip o günün
        # verisi mükerrer sayılıyordu (Geldi_Gunu ikiye katlanıyor ama P_toplam katlanmıyordu, bu da
        # %200 gibi imkânsız Devam Oranlarına yol açıyordu). Bu yüzden sadece TAM SAYI değerli
        # (ondalık kısmı olmayan) başlıkları gün sütunu olarak kabul ediyoruz.
        day_cols = []
        gorulen_gunler = set()
        for c in cols:
            c_str = str(c).strip()
            try:
                d_float = float(c_str)
            except (ValueError, TypeError):
                continue
            if not d_float.is_integer():
                continue  # "1.1", "2.3" gibi pandas'ın otomatik eklediği yinelenen sütun son ekleri
            d = int(d_float)
            if 1 <= d <= gun_sayisi_ay and d not in gorulen_gunler:
                day_cols.append((c, d))
                gorulen_gunler.add(d)

        if not day_cols:
            continue

        day_col_names = {c for c, _ in day_cols}
        # Gün sütunlarından ilkinin index'i: Sicil No ile ilk gün sütunu arasındaki her şey isim sütunu kabul edilir.
        # (Örn: "AEROSOL ÜRETİM", "Unnamed: 2" gibi ayrı Ad/Soyad sütunları da bu şekilde otomatik yakalanır.)
        first_day_idx = min(cols.index(c) for c in day_col_names)
        name_cols = [c for c in cols[1:first_day_idx] if c not in day_col_names]


        for _, row in df.iterrows():
            sicil_raw = row.get(sicil_col)
            if sicil_raw is None or (isinstance(sicil_raw, float) and pd.isna(sicil_raw)) or str(sicil_raw).strip() == "":
                continue
            sicil = str(sicil_raw).strip()

            ad_soyad_parcalari = []
            for c in name_cols:
                deger = row.get(c)
                if deger is None or (isinstance(deger, float) and pd.isna(deger)):
                    continue
                deger_str = str(deger).strip()
                if deger_str == "" or deger_str.lower() == "nan":
                    continue
                ad_soyad_parcalari.append(deger_str)
            ad_soyad = " ".join(ad_soyad_parcalari).strip()
            if ad_soyad == "":
                ad_soyad = sicil  # son çare: sicil no'yu isim olarak kullan

            # Gerçek personel olmayan (özel gün/sayım notu vb.) satırları tamamen atla.
            if not gecerli_personel_mi(sicil, ad_soyad):
                continue

            for col, gun_no in day_cols:
                raw_val = row.get(col)
                tatil, tatil_sebep = is_tatil_gunu(yil, ay_no, gun_no)
                durum, saat = classify_durum(raw_val)

                if tatil:
                    durum_final = "Tatil"
                else:
                    durum_final = durum

                all_rows.append({
                    "Sicil No": sicil,
                    "Ad Soyad": ad_soyad,
                    "Ay": sheet,
                    "AyNo": ay_no,
                    "Yil": yil,
                    "Gun": gun_no,
                    "Tarih": date(yil, ay_no, gun_no),
                    "Durum": durum_final,
                    "Saat": saat,
                    "TatilSebep": tatil_sebep,
                })

    sonuc_df = pd.DataFrame(all_rows)
    if not sonuc_df.empty:
        # Ek güvenlik katmanı: aynı personelin aynı ay/gün için (veri girişi hatası, tekrarlanan
        # satır vb. nedenlerle) birden fazla kaydı varsa, her gün yalnızca BİR kez sayılmalıdır.
        # Aksi halde Geldi_Gunu gibi sayaçlar mükerrer artıp %100'ü aşan devam oranlarına yol açar.
        sonuc_df = sonuc_df.drop_duplicates(subset=["Sicil No", "Ay", "Gun"], keep="first")
    return sonuc_df


def generate_mock_long_df(n_calisan: int = 25, yil: int = 2026) -> pd.DataFrame:
    """Excel yüklenmediğinde kullanılacak örnek/mock veri seti üretir."""
    random.seed(42)

    isimler = [
        "Ahmet Yılmaz", "Mehmet Demir", "Ayşe Kaya", "Fatma Şahin", "Ali Çelik",
        "Zeynep Arslan", "Mustafa Doğan", "Elif Aydın", "Hüseyin Yıldız", "Emine Öztürk",
        "İbrahim Aslan", "Hatice Koç", "Yusuf Kurt", "Meryem Şimşek", "Ömer Aksoy",
        "Rabia Güneş", "Murat Bulut", "Sultan Polat", "Kemal Erdoğan", "Nur Aktaş",
        "Ferhat Yalçın", "Gül Türkmen", "Cemal Öz", "Aslı Kara", "Serkan Bozkurt",
    ]
    calisanlar = [(f"{1000 + i}", isimler[i % len(isimler)]) for i in range(n_calisan)]

    aylar = [
        (1, "OCAK-2026"), (2, "ŞUBAT-2026"), (3, "MART-2026"), (4, "NİSAN-2026"),
        (5, "MAYIS-2026"), (6, "HAZİRAN-2026"), (7, "TEMMUZ-2026"),
    ]

    durum_havuzu = (
        ["Geldi"] * 42 + ["İstirahatli"] * 4 + ["Ücretsiz İzin"] * 2 +
        ["Yıllık İzinli"] * 3 + ["Kozmetik"] * 4 + ["Ambalaj"] * 4 + ["Refakatçi"] * 2
    )
    saat_havuzu = [7.5, 7.5, 7.5, 7.5, 8, 8, 11, 5]

    rows = []
    for ay_no, ay_adi in aylar:
        gun_sayisi = calendar.monthrange(yil, ay_no)[1]
        for sicil, ad_soyad in calisanlar:
            # Her çalışan için hafif farklı bir performans eğilimi (bazıları daha düşük katılımlı)
            egilim = random.choice(["yuksek"] * 6 + ["orta"] * 3 + ["dusuk"] * 1)
            for gun in range(1, gun_sayisi + 1):
                tatil, tatil_sebep = is_tatil_gunu(yil, ay_no, gun)
                if tatil:
                    rows.append({
                        "Sicil No": sicil, "Ad Soyad": ad_soyad, "Ay": ay_adi, "AyNo": ay_no,
                        "Yil": yil, "Gun": gun, "Tarih": date(yil, ay_no, gun),
                        "Durum": "Tatil", "Saat": None, "TatilSebep": tatil_sebep,
                    })
                    continue

                if egilim == "yuksek":
                    secim = random.choices(durum_havuzu, weights=None, k=1)[0] if random.random() < 0.95 else "İstirahatli"
                elif egilim == "orta":
                    secim = random.choice(durum_havuzu)
                else:
                    secim = random.choice(["Geldi"] * 25 + ["İstirahatli"] * 10 + ["Ücretsiz İzin"] * 6 + ["Yıllık İzinli"] * 4)

                if secim == "Geldi":
                    saat_val = random.choice(saat_havuzu)
                    durum_final, saat = classify_durum(saat_val)
                else:
                    durum_final, saat = secim, None

                rows.append({
                    "Sicil No": sicil, "Ad Soyad": ad_soyad, "Ay": ay_adi, "AyNo": ay_no,
                    "Yil": yil, "Gun": gun, "Tarih": date(yil, ay_no, gun),
                    "Durum": durum_final, "Saat": saat, "TatilSebep": None,
                })

    return pd.DataFrame(rows)


def aggregate_calisan_ay(long_df: pd.DataFrame) -> pd.DataFrame:
    """Çalışan x Ay kırılımında P_toplam, P_geldi ve tüm sayaçları hesaplar."""
    if long_df.empty:
        return pd.DataFrame(columns=[
            "Sicil No", "Ad Soyad", "Ay", "AyNo", "Yil", "Toplam_Takvim_Gunu", "Tatil_Gunu",
            "Geldi_Gunu", "Kozmetik_Gunu", "Ambalaj_Gunu", "Refakatci_Gunu", "Istirahat_Gunu",
            "Ucretsiz_Izin_Gunu", "Yillik_Izin_Gunu", "Diger_Gunu", "Toplam_Saat", "P_toplam", "Devam_Orani",
        ])

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
    agg["Devam_Orani"] = np.where(agg["P_toplam"] > 0, agg["Geldi_Gunu"] / agg["P_toplam"] * 100, np.nan)
    return agg


@st.cache_data(show_spinner="Excel dosyası işleniyor...")
def process_excel_bytes(file_bytes: bytes) -> pd.DataFrame:
    return load_excel_to_long(io.BytesIO(file_bytes))



# =========================================================================================
# 3) SIDEBAR - VERİ YÜKLEME VE AYARLAR
# =========================================================================================

st.sidebar.header("⚙️ Ayarlar")
uploaded_file = st.sidebar.file_uploader("📁 Puantaj Excel Dosyası Yükleyin (.xlsx)", type=["xlsx"])

st.sidebar.markdown("---")
st.sidebar.caption(
    "ℹ️ 2026 yılına ait resmi tatil, bayram ve arife günleri otomatik olarak hesaba katılmıştır. "
    "Pazar günleri de haftalık tatil olarak kabul edilip net iş gününden düşülür."
)

# --- Veri Yükleme ---
kaynak = ""
if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.getvalue()
        long_df = process_excel_bytes(file_bytes)
        if long_df.empty:
            st.sidebar.warning("⚠️ Dosyada okunabilir bir ay sekmesi bulunamadı. Örnek veri gösteriliyor.")
            long_df = generate_mock_long_df()
            kaynak = "Örnek (Mock) Veri"
        else:
            kaynak = f"Yüklenen Dosya: {uploaded_file.name}"
    except Exception as e:
        st.sidebar.error(f"❌ Excel dosyası işlenirken hata oluştu: {e}")
        long_df = generate_mock_long_df()
        kaynak = "Örnek (Mock) Veri"
else:
    long_df = generate_mock_long_df()
    kaynak = "Örnek (Mock) Veri (Demo)"

agg = aggregate_calisan_ay(long_df)
ay_sira = agg[["Ay", "AyNo"]].drop_duplicates().sort_values("AyNo")["Ay"].tolist()


# =========================================================================================
# 4) BAŞLIK
# =========================================================================================

st.title("🏭 Fabrika Puantaj / Devam Analiz Sistemi")
st.caption(f"📌 Veri kaynağı: **{kaynak}**  |  Toplam kayıt satırı: {len(long_df):,}")

if agg.empty:
    st.error("Hesaplanacak veri bulunamadı. Lütfen Excel dosyanızın yapısını kontrol edin.")
    st.stop()


# =========================================================================================
# 5) BİREYSEL PERSONEL SORGULAMA (EN ÜSTTE)
# =========================================================================================

st.header("👤 Bireysel Personel Sorgulama")

personel_liste = agg[["Sicil No", "Ad Soyad"]].drop_duplicates().copy()
personel_liste["Ad Soyad"] = personel_liste["Ad Soyad"].astype(str).str.strip()
personel_liste["Etiket"] = personel_liste["Sicil No"].astype(str) + " - " + personel_liste["Ad Soyad"]
personel_liste = personel_liste.sort_values("Ad Soyad")

col_f1, col_f2 = st.columns(2)
with col_f1:
    secili_etiket = st.selectbox(
        "🔍 Personel Seçimi (Ad Soyad / Sicil No ile arayabilirsiniz)",
        options=personel_liste["Etiket"].tolist(),
    )
with col_f2:
    secili_aylar = st.multiselect(
        "🗓️ Ay / Dönem Filtresi (Bir veya Birden Fazla Ay Seçebilirsiniz)",
        options=ay_sira,
        default=ay_sira,
        help="Seçtiğiniz ay(lar) aşağıdaki tüm KPI kartlarını, gün gün detay tablosunu ve çalışma türü grafiğini anında filtreler.",
    )

secili_sicil = personel_liste.loc[personel_liste["Etiket"] == secili_etiket, "Sicil No"].iloc[0]

# Seçilen personel ve seçilen ay(lar) birebir aynı filtreyle hem KPI'lara hem detay tablosuna uygulanır;
# böylece "Tüm Aylar" (hiçbir ay çıkarılmamış hali) dahil her durumda tam senkronizasyon sağlanır.
kisi_agg = agg[(agg["Sicil No"] == secili_sicil) & (agg["Ay"].isin(secili_aylar))]
kisi_long = long_df[(long_df["Sicil No"] == secili_sicil) & (long_df["Ay"].isin(secili_aylar))]

if not secili_aylar:
    donem_etiketi = "Seçim Yok"
elif set(secili_aylar) == set(ay_sira):
    donem_etiketi = "Tüm Aylar"
elif len(secili_aylar) <= 3:
    donem_etiketi = ", ".join(secili_aylar)
else:
    donem_etiketi = f"{len(secili_aylar)} Ay Seçili"

if not secili_aylar:
    st.warning("⚠️ Lütfen en az bir ay seçiniz. Aşağıdaki metrikler ve tablo boş görünecektir.")

p_toplam_k = kisi_agg["P_toplam"].sum()
geldi_k = kisi_agg["Geldi_Gunu"].sum()
istirahat_k = kisi_agg["Istirahat_Gunu"].sum()
ucretsiz_k = kisi_agg["Ucretsiz_Izin_Gunu"].sum()
yillik_k = kisi_agg["Yillik_Izin_Gunu"].sum()
kozmetik_k = kisi_agg["Kozmetik_Gunu"].sum()
ambalaj_k = kisi_agg["Ambalaj_Gunu"].sum()
refakatci_k = kisi_agg["Refakatci_Gunu"].sum()
standart_geldi_k = max(geldi_k - kozmetik_k - ambalaj_k - refakatci_k, 0)
toplam_saat_k = kisi_agg["Toplam_Saat"].sum()
devam_orani_k = (geldi_k / p_toplam_k * 100) if p_toplam_k > 0 else np.nan

st.markdown(f"**Seçilen Personel:** {secili_etiket}  |  **Dönem:** {donem_etiketi}")

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric(
    "✅ Geldiği Gün Sayısı", f"{int(geldi_k)}",
    help=(f"Standart Geldi: {int(standart_geldi_k)} | Kozmetik: {int(kozmetik_k)} | "
          f"Ambalaj: {int(ambalaj_k)} | Refakatçi: {int(refakatci_k)} | Toplam Kayıtlı Saat: {toplam_saat_k:.1f} sa"),
)
kpi2.metric("🩺 İstirahat / Rapor Gün Sayısı", f"{int(istirahat_k)}")
kpi3.metric(
    "🌴 Ücretsiz / Yıllık İzin Sayısı", f"{int(ucretsiz_k + yillik_k)}",
    help=f"Ücretsiz İzin: {int(ucretsiz_k)} | Yıllık İzinli: {int(yillik_k)}",
)
kpi4.metric("📈 Devam Performans Oranı", f"%{devam_orani_k:.1f}" if p_toplam_k > 0 else "N/A")

st.subheader("📅 Gün Gün Katılım / Saat Detay Tablosu")

detay = kisi_long[["AyNo", "Ay", "Gun", "Tarih", "Durum", "Saat", "TatilSebep"]].sort_values(["AyNo", "Gun"]).copy()
detay = detay.drop(columns=["AyNo"]).rename(columns={
    "Gun": "Gün", "Tarih": "Tarih", "Durum": "Durum", "Saat": "Çalışılan Saat", "TatilSebep": "Tatil Açıklaması",
})


def _renk_uygula(val):
    return DURUM_RENK.get(val, "")


styled_detay = detay.style.map(_renk_uygula, subset=["Durum"])
st.dataframe(styled_detay, use_container_width=True, height=420)

st.subheader("📊 Özel Çalışma Türü Dağılımı")

tur_dagilim = pd.DataFrame({
    "Tür": ["Standart Geldi", "Kozmetik", "Ambalaj", "Refakatçi"],
    "Gün": [standart_geldi_k, kozmetik_k, ambalaj_k, refakatci_k],
})
fig_tur = px.bar(
    tur_dagilim, x="Tür", y="Gün", text="Gün", color="Tür",
    title=f"{secili_etiket} — Çalışma Türü Dağılımı ({donem_etiketi})",
    color_discrete_sequence=["#2E8B57", "#3498DB", "#F39C12", "#9B59B6"],
)
fig_tur.update_layout(showlegend=False, xaxis_title="", yaxis_title="Gün Sayısı")
st.plotly_chart(fig_tur, use_container_width=True)

st.divider()


# =========================================================================================
# 6) ANALİZ DÖNEMİ FİLTRESİ + FABRİKA GENELİ ÖZET METRİKLER
# =========================================================================================

st.header("📊 Fabrika Geneli Özet")

secili_donem_aylar = st.multiselect(
    "📅 Analiz Dönemi (Bir veya Birden Fazla Ay Seçebilirsiniz)",
    options=ay_sira,
    default=ay_sira,
    key="genel_ozet_donem",
    help="Buradan seçtiğiniz dönem; aşağıdaki Fabrika Geneli Özet metriklerini, Top 5 İstirahatli/İzinli listesini, "
         "Top 10/Bottom 10 grafiklerini ve izin dağılımı grafiğini birlikte günceller.",
)

if not secili_donem_aylar:
    st.warning("⚠️ Lütfen en az bir ay seçiniz.")
    agg_donem = agg.iloc[0:0]
else:
    agg_donem = agg[agg["Ay"].isin(secili_donem_aylar)]

donem_etiketi_ozet = "Tüm Aylar" if (secili_donem_aylar and set(secili_donem_aylar) == set(ay_sira)) else (
    "Seçim Yok" if not secili_donem_aylar else (
        ", ".join(secili_donem_aylar) if len(secili_donem_aylar) <= 3 else f"{len(secili_donem_aylar)} Ay Seçili"
    )
)

# NOT: "Toplam Aktif Çalışan" ve bağlı metrikler, tüm ayların toplamından değil,
# yalnızca seçilen döneme (agg_donem) ait gerçek kayıtlardan hesaplanır.
toplam_calisan = agg_donem["Sicil No"].nunique()
genel_p_toplam = agg_donem["P_toplam"].sum()
genel_p_geldi = agg_donem["Geldi_Gunu"].sum()
fabrika_katilim_orani = (genel_p_geldi / genel_p_toplam * 100) if genel_p_toplam > 0 else np.nan
toplam_rapor = agg_donem["Istirahat_Gunu"].sum()

col1, col2, col3 = st.columns(3)
col1.metric("👥 Toplam Aktif Çalışan", f"{toplam_calisan}", help=f"Dönem: {donem_etiketi_ozet}")
col2.metric("✅ Fabrika Geneli Katılım Oranı", f"%{fabrika_katilim_orani:.1f}")
col3.metric("🩺 Toplam Rapor/İstirahat Günü", f"{int(toplam_rapor)}")

st.subheader("🏅 En Çok İstirahatli / İzinli Olan Top 5 Personel")
st.caption(f"Dönem: {donem_etiketi_ozet}")

top5_kategori = st.selectbox(
    "Kategori",
    ["Toplam (İstirahatli + İzin)", "İstirahatli", "Ücretsiz İzin", "Yıllık İzinli"],
    key="top5_kategori",
)

if agg_donem.empty:
    st.info("Top 5 listesini görüntülemek için en az bir ay seçiniz.")
else:
    kisi_ozet_rapor = agg_donem.groupby(["Sicil No", "Ad Soyad"]).agg(
        Istirahat_Gunu=("Istirahat_Gunu", "sum"),
        Ucretsiz_Izin_Gunu=("Ucretsiz_Izin_Gunu", "sum"),
        Yillik_Izin_Gunu=("Yillik_Izin_Gunu", "sum"),
    ).reset_index()
    kisi_ozet_rapor["Ad Soyad"] = kisi_ozet_rapor["Ad Soyad"].astype(str).str.strip()
    # Grafik Y ekseninde asla sicil no tek başına (sayısal) gösterilmesin diye metinsel Etiket kullanılır.
    kisi_ozet_rapor["Etiket"] = kisi_ozet_rapor["Sicil No"].astype(str) + " - " + kisi_ozet_rapor["Ad Soyad"]

    if top5_kategori == "Toplam (İstirahatli + İzin)":
        kisi_ozet_rapor["Gun_Sayisi"] = (
            kisi_ozet_rapor["Istirahat_Gunu"] + kisi_ozet_rapor["Ucretsiz_Izin_Gunu"] + kisi_ozet_rapor["Yillik_Izin_Gunu"]
        )
    elif top5_kategori == "İstirahatli":
        kisi_ozet_rapor["Gun_Sayisi"] = kisi_ozet_rapor["Istirahat_Gunu"]
    elif top5_kategori == "Ücretsiz İzin":
        kisi_ozet_rapor["Gun_Sayisi"] = kisi_ozet_rapor["Ucretsiz_Izin_Gunu"]
    else:
        kisi_ozet_rapor["Gun_Sayisi"] = kisi_ozet_rapor["Yillik_Izin_Gunu"]

    top5_rapor = kisi_ozet_rapor[kisi_ozet_rapor["Gun_Sayisi"] > 0].sort_values("Gun_Sayisi", ascending=False).head(5)

    if top5_rapor.empty:
        st.info("Seçili kriterlere uyan bir kayıt bulunmuyor.")
    else:
        top5_rapor_sirali = top5_rapor.sort_values("Gun_Sayisi")
        fig_top5_rapor = px.bar(
            top5_rapor_sirali, x="Gun_Sayisi", y="Etiket", orientation="h",
            title=None, text="Gun_Sayisi",
            color_discrete_sequence=["#E67E22"],
        )
        fig_top5_rapor.update_traces(texttemplate="%{text:.0f} gün")
        fig_top5_rapor.update_layout(xaxis_title="Gün Sayısı", yaxis_title="", height=320)
        fig_top5_rapor.update_yaxes(type="category", categoryorder="array", categoryarray=top5_rapor_sirali["Etiket"].tolist())
        st.plotly_chart(fig_top5_rapor, use_container_width=True)

st.divider()


# =========================================================================================
# 7) FABRİKA GENELİ ANALİZ GRAFİKLERİ
# =========================================================================================

st.subheader("📈 Fabrika Geneli Analiz Grafikleri")
st.caption(f"Top 10/Bottom 10 performans ve izin dağılımı grafikleri seçili döneme ({donem_etiketi_ozet}) göre hesaplanır.")

kisi_ozet = agg_donem.groupby(["Sicil No", "Ad Soyad"]).agg(
    P_toplam=("P_toplam", "sum"),
    Geldi_Gunu=("Geldi_Gunu", "sum"),
    Istirahat_Gunu=("Istirahat_Gunu", "sum"),
    Ucretsiz_Izin_Gunu=("Ucretsiz_Izin_Gunu", "sum"),
    Yillik_Izin_Gunu=("Yillik_Izin_Gunu", "sum"),
).reset_index()
kisi_ozet["Devam_Orani"] = np.where(kisi_ozet["P_toplam"] > 0, kisi_ozet["Geldi_Gunu"] / kisi_ozet["P_toplam"] * 100, np.nan)

# ÖNEMLİ: Grafiklerin Y ekseninde ASLA sicil no tek başına (sayısal) gösterilmemeli.
# "Sicil No - Ad Soyad" formatında, kesinlikle metinsel (string) bir Etiket sütunu oluşturuyoruz.
kisi_ozet["Ad Soyad"] = kisi_ozet["Ad Soyad"].astype(str).str.strip()
kisi_ozet["Etiket"] = kisi_ozet["Sicil No"].astype(str) + " - " + kisi_ozet["Ad Soyad"]

top10 = kisi_ozet.dropna(subset=["Devam_Orani"]).sort_values("Devam_Orani", ascending=False).head(10)
bottom10 = kisi_ozet.dropna(subset=["Devam_Orani"]).sort_values("Devam_Orani", ascending=True).head(10)

col_a, col_b = st.columns(2)
with col_a:
    top10_sirali = top10.sort_values("Devam_Orani")
    fig_top = px.bar(
        top10_sirali, x="Devam_Orani", y="Etiket", orientation="h",
        title="🏆 En Yüksek Performanslı Çalışanlar (Top 10)", text="Devam_Orani",
        color_discrete_sequence=["#2E8B57"],
    )
    fig_top.update_traces(texttemplate="%{text:.1f}%")
    fig_top.update_layout(xaxis_title="Devam Oranı (%)", yaxis_title="")
    fig_top.update_yaxes(type="category", categoryorder="array", categoryarray=top10_sirali["Etiket"].tolist())
    st.plotly_chart(fig_top, use_container_width=True)

with col_b:
    bottom10_sirali = bottom10.sort_values("Devam_Orani", ascending=False)
    fig_bottom = px.bar(
        bottom10_sirali, x="Devam_Orani", y="Etiket", orientation="h",
        title="⬇️ En Düşük Performanslı Çalışanlar (Bottom 10)", text="Devam_Orani",
        color_discrete_sequence=["#C0392B"],
    )
    fig_bottom.update_traces(texttemplate="%{text:.1f}%")
    fig_bottom.update_layout(xaxis_title="Devam Oranı (%)", yaxis_title="")
    fig_bottom.update_yaxes(type="category", categoryorder="array", categoryarray=bottom10_sirali["Etiket"].tolist())
    st.plotly_chart(fig_bottom, use_container_width=True)

dagilim = pd.DataFrame({
    "Kategori": ["İstirahatli / Rapor", "Ücretsiz İzin", "Yıllık İzinli"],
    "Gün": [agg_donem["Istirahat_Gunu"].sum(), agg_donem["Ucretsiz_Izin_Gunu"].sum(), agg_donem["Yillik_Izin_Gunu"].sum()],
})
fig_pie = px.pie(
    dagilim, names="Kategori", values="Gün", hole=0.45,
    title=f"🥧 Devamsızlık ve İzin Dağılımı ({donem_etiketi_ozet})",
    color_discrete_sequence=["#E74C3C", "#F39C12", "#3498DB"],
)
st.plotly_chart(fig_pie, use_container_width=True)

st.divider()


# =========================================================================================
# 8) AYLARA GÖRE TREND GRAFİKLERİ (Dinamik Ay Aralığı Filtresi)
# =========================================================================================

st.subheader("📉 Aylara Göre Trend Grafikleri")

trend_ay_secimi = st.multiselect(
    "🗓️ Trend Grafikleri için Ay Aralığı Seçiniz",
    options=ay_sira,
    default=ay_sira,
    help="Buradan seçtiğiniz aylara göre aşağıdaki iki trend grafiği dinamik olarak daraltılır/güncellenir.",
)

if not trend_ay_secimi:
    st.info("Trend grafiklerini görüntülemek için en az bir ay seçiniz.")
else:
    ay_ozet = agg[agg["Ay"].isin(trend_ay_secimi)].groupby(["Ay", "AyNo"]).agg(
        Istirahat_Gunu=("Istirahat_Gunu", "sum"),
        P_toplam=("P_toplam", "sum"),
        Geldi_Gunu=("Geldi_Gunu", "sum"),
    ).reset_index().sort_values("AyNo")
    ay_ozet["Devam_Orani"] = np.where(ay_ozet["P_toplam"] > 0, ay_ozet["Geldi_Gunu"] / ay_ozet["P_toplam"] * 100, np.nan)

    col_c, col_d = st.columns(2)
    with col_c:
        fig_rapor = px.bar(
            ay_ozet, x="Ay", y="Istirahat_Gunu", title="🩺 Aylara Göre Rapor/İstirahat Trendi",
            text="Istirahat_Gunu", color_discrete_sequence=["#5DADE2"],
        )
        fig_rapor.update_layout(xaxis_title="", yaxis_title="İstirahat/Rapor Gün Sayısı")
        st.plotly_chart(fig_rapor, use_container_width=True)

    with col_d:
        fig_devam = px.line(
            ay_ozet, x="Ay", y="Devam_Orani", markers=True,
            title="📈 Aylara Göre Fabrika Geneli Devam Oranı Trendi",
        )
        fig_devam.update_traces(line_color="#8E44AD")
        fig_devam.update_layout(xaxis_title="", yaxis_title="Devam Oranı (%)")
        st.plotly_chart(fig_devam, use_container_width=True)
