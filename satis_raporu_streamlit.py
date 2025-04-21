import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Personel Satış Raporu", layout="wide")
st.title("📊 Genişletilmiş Personel Satış Raporu")

uploaded_file = st.file_uploader("Excel dosyasını yükleyin (doğru formatta):", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df["DSL Hız Bilgisi"] = pd.to_numeric(df["DSL Hız Bilgisi"], errors="coerce")
        df_aktif = df[df["Durum"] == "Aktif"]

        columns = [
            "Mobil Faturalı", "Faturalı Yıldızlı Günler", "Mobil Faturasız", "Faturasız Turist", "DSL Aktivasyon", "DSL 100 MBPS Altı", "DSL 100 MBPS Üstü",
            "12 MBPS", "16 MBPS", "24 MBPS", "35 MBPS", "50 MBPS", "75 MBPS", "100 MBPS", "200 MBPS", "500 MBPS", "1000 MBPS",
            "Tivi Uydu", "IPTV", "IPTV Sinema Spor Super", "Odak Cihaz", "Akıllı Premium Cihaz", "Akıllı Standart Cihaz", "Akıllı PK",
            "Diğer Cihaz", "Terminal Servisler", "Tablet Temlik", "Tablet PK", "YNA PK", "YNA Standart", "YNA Premium",
            "Mobil Taahhüt", "Mobil Upsell", "DSL Pure Upsell", "DSL Taahhüt", "MAP",
            "Paket 150", "Paket 200", "Paket 250", "Paket 300", "Paket 350", "Paket 400", "Paket 500", "Paket 650", "Paket 750",
            "Paket 1000", "Paket 1250", "Paket 1500", "Paket 1750", "Paket 2500", "Uç Cihaz"
        ]

        group_keys = ["Pos Kodu", "İşlemi Yapan Kullanıcı Kodu", "İşlemi Yapan Kullanıcı"]
        report_df = pd.DataFrame(columns=group_keys + columns)

        gruplar = df_aktif.groupby(group_keys)

        def kategori_say(grup, filtre):
            try:
                return len(grup.query(filtre)) if not grup.empty else 0
            except Exception:
                return 0

        for keys, grup in gruplar:
            row = dict(zip(group_keys, keys))
            row["Mobil Faturalı"] = kategori_say(grup, '`İşlem Tipi` in ["MBB", "SES"] and `Ürün Tipi` in ["FATURALI-NT", "FATURALI-Yeni"]')
            row["Faturalı Yıldızlı Günler"] = 0
            row["Mobil Faturasız"] = kategori_say(grup, '`İşlem Tipi` in ["MBB", "SES"] and `Ürün Tipi` in ["FATURASIZ-NT", "FATURASIZ-Yeni"]')
            row["Faturasız Turist"] = 0
            row["DSL Aktivasyon"] = kategori_say(grup, '`İşlem Tipi` == "DSL"')
            row["DSL 100 MBPS Altı"] = kategori_say(grup, '`İşlem Tipi` == "DSL" and `DSL Hız Bilgisi` in [12288, 16384, 24576, 35840, 51200, 76800]')
            row["DSL 100 MBPS Üstü"] = kategori_say(grup, '`İşlem Tipi` == "DSL" and `DSL Hız Bilgisi` in [102400, 204800, 512000, 1048576]')
            row["12 MBPS"] = kategori_say(grup, '`DSL Hız Bilgisi` == 12288')
            row["16 MBPS"] = kategori_say(grup, '`DSL Hız Bilgisi` == 16384')
            row["24 MBPS"] = kategori_say(grup, '`DSL Hız Bilgisi` == 24576')
            row["35 MBPS"] = kategori_say(grup, '`DSL Hız Bilgisi` == 35840')
            row["50 MBPS"] = kategori_say(grup, '`DSL Hız Bilgisi` == 51200')
            row["75 MBPS"] = kategori_say(grup, '`DSL Hız Bilgisi` == 76800')
            row["100 MBPS"] = kategori_say(grup, '`DSL Hız Bilgisi` == 102400')
            row["200 MBPS"] = kategori_say(grup, '`DSL Hız Bilgisi` == 204800')
            row["500 MBPS"] = kategori_say(grup, '`DSL Hız Bilgisi` == 512000')
            row["1000 MBPS"] = kategori_say(grup, '`DSL Hız Bilgisi` == 1048576')
            row["Tivi Uydu"] = kategori_say(grup, '`İşlem Tipi` == "TİVİBU UYDU"')
            row["IPTV"] = kategori_say(grup, '`İşlem Tipi` == "IPTV"')
            row["IPTV Sinema Spor Super"] = 0
            row["Odak Cihaz"] = 0
            row["Akıllı Premium Cihaz"] = 0
            row["Akıllı Standart Cihaz"] = kategori_say(grup, '`İşlem Tipi` == "AKILLI" and `Kampanya Tipi` in ["Temlikli Cihaz", "TFŞ"]')
            row["Akıllı PK"] = kategori_say(grup, '`İşlem Tipi` == "AKILLI" and `Kampanya Tipi` == "Peşine Kontrat"')
            row["Diğer Cihaz"] = kategori_say(grup, '`İşlem Tipi` in ["Aksesuar", "Mobil Aksesuar", "PC", "TABLET", "Terminal Servisler", "Uç Cihaz"]')
            row["Terminal Servisler"] = kategori_say(grup, '`İşlem Tipi` == "Terminal Servisler"')
            row["Tablet Temlik"] = kategori_say(grup, '`İşlem Tipi` == "TABLET" and `Kampanya Tipi` in ["Temlikli Cihaz", "TFŞ"]')
            row["Tablet PK"] = kategori_say(grup, '`İşlem Tipi` == "TABLET" and `Kampanya Tipi` == "Peşine Kontrat"')
            row["YNA PK"] = kategori_say(grup, '`İşlem Tipi` == "Aksesuar" and `Kampanya Tipi` == "Peşine Kontrat"')
            row["YNA Standart"] = kategori_say(grup, '`İşlem Tipi` == "Aksesuar" and `Kampanya Tipi` in ["Temlikli Cihaz", "TFŞ"]')
            row["YNA Premium"] = 0
            row["Mobil Taahhüt"] = kategori_say(grup, '`İşlem Tipi` == "Mobil Taahhüt"')
            row["Mobil Upsell"] = kategori_say(grup, '`İşlem Tipi` == "Mobil Upsell"')
            row["DSL Pure Upsell"] = kategori_say(grup, '`İşlem Tipi` == "DSL Pure Upsell"')
            row["DSL Taahhüt"] = kategori_say(grup, '`İşlem Tipi` == "DSL Taahhüt"')
            row["MAP"] = kategori_say(grup, '`Model`.isin(["Paket150", "Paket200", "Paket250", "Paket300", "Paket350", "Paket400", "Paket500", "Paket650", "Paket750", "1000", "1250", "1500", "1750", "2500"])')
            for paket in ["150", "200", "250", "300", "350", "400", "500", "650", "750", "1000", "1250", "1500", "1750", "2500"]:
                row[f"Paket {paket}"] = kategori_say(grup, f'`Model` == "Paket{paket}"' if not paket.isdigit() else f'`Model` == "{paket}"')
            row["Uç Cihaz"] = kategori_say(grup, '`İşlem Tipi` == "Uç Cihaz"')

            report_df = pd.concat([report_df, pd.DataFrame([row])], ignore_index=True)

        st.success("✅ Rapor başarıyla oluşturuldu.")
        st.dataframe(report_df)

        # Excel çıktısı
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            report_df.to_excel(writer, index=False, sheet_name="Rapor")
        buffer.seek(0)

        st.download_button(
            label="📥 Raporu Excel olarak indir",
            data=buffer,
            file_name=f"genis_satis_raporu_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Hata oluştu: {e}")
