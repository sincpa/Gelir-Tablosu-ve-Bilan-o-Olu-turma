"""
app.py
------
Mizandan Gelir Tablosu ve Bilanco olusturan web uygulamasi.

CALISTIRMA (kendi bilgisayarinda test icin):
    streamlit run app.py

YAYINLAMA (Streamlit Community Cloud):
    1. Bu klasoru GitHub'a public repo olarak yukle
    2. share.streamlit.io -> New app -> repo'yu sec -> Deploy
    3. Cikan linki paylas
"""

import traceback
from datetime import datetime

import streamlit as st

from isleyici import MIZAN_CARI, MIZAN_ONCEKI, isle

st.set_page_config(
    page_title="Mizandan Mali Tablo",
    page_icon="📊",
    layout="wide",
)

# ----------------------------------------------------------------------
# BASLIK
# ----------------------------------------------------------------------
st.title("📊 Mizandan Gelir Tablosu ve Bilanço")
st.caption(
    "Finansal analiz raporu şablonunuzu yükleyin; program yalnızca "
    "**Gelir Tablosu** ve **Bilanço** sayfalarını doldurup dosyayı aynen geri verir."
)

# ----------------------------------------------------------------------
# KULLANIM TALIMATI
# ----------------------------------------------------------------------
with st.expander("ℹ️  Nasıl kullanılır? (ilk kullanımda okuyun)", expanded=False):
    st.markdown(
        f"""
**1. Şablonunuzu Excel'de açın.**

**2. En sona iki sayfa ekleyin ve mizanları yapıştırın:**

| Sayfa adı | İçerik |
|---|---|
| `{MIZAN_ONCEKI}` | **Geçen yılın** mizanı |
| `{MIZAN_CARI}` | **Bu yılın** mizanı |

Sayfa adı tam olarak bu şekilde, **küçük harfle ve boşluksuz** yazılmalıdır.
`Mizan1`, `MIZAN2`, `mizan 1` gibi yazımlar kabul edilmez.

**3. Mizanda şu sütun başlıkları bulunmalıdır:**
`Hesap Kodu` · `Hesap Adı` · `Borç Bakiye` · `Alacak Bakiye`

**4. Dosyayı buraya yükleyin ve sonucu indirin.**

---

**Program ne yapar?**
- 7/A maliyet hesaplarındaki yansıtılmamış tutarları 6'lı hesaplara aktarır
- Dönem kâr/zararını hesaplayıp 590 veya 591'e yazar
- Aktif = Pasif denkliğini kontrol eder
- `{MIZAN_ONCEKI}` ve `{MIZAN_CARI}` sayfalarını siler (dosya doğrudan mükellefe gidebilir)

**Program ne yapmaz?**
- **Vergi karşılığı hesaplamaz.** Bunlar geçici vergi dönemi tablolarıdır;
  vergi hesabı ayrı sayfada elle yapılır.
- Üretim/stok maliyeti hesaplamaz. Mizandaki bakiyeler neyse o kullanılır.
- Şablonun biçimlendirmesine dokunmaz. Renk, yazı tipi, formüller korunur.
"""
    )

st.divider()

# ----------------------------------------------------------------------
# DOSYA YUKLEME
# ----------------------------------------------------------------------
yuklenen = st.file_uploader(
    "Şablon Excel dosyanızı seçin",
    type=["xlsx", "xlsm"],
    help=f"İçinde '{MIZAN_ONCEKI}' ve/veya '{MIZAN_CARI}' sayfaları bulunmalıdır.",
)

if yuklenen is None:
    st.info("Başlamak için yukarıdan Excel dosyanızı yükleyin.")
    st.stop()

baytlar = yuklenen.getvalue()

if not st.button("▶️  Tabloları Doldur", type="primary", use_container_width=True):
    st.stop()

# ----------------------------------------------------------------------
# ISLEM
# ----------------------------------------------------------------------
try:
    with st.spinner("Mizanlar okunuyor, tablolar oluşturuluyor..."):
        cikti, sonuc = isle(baytlar)
except ValueError as hata:
    st.error(str(hata))
    st.stop()
except Exception:
    st.error("Beklenmeyen bir hata oluştu. Ayrıntı aşağıda:")
    st.code(traceback.format_exc())
    st.stop()

st.success(f"Tamamlandı. {sonuc.yazilan_hucre} hücre yazıldı.")

# ----------------------------------------------------------------------
# INDIRME
# ----------------------------------------------------------------------
ad = yuklenen.name.rsplit(".", 1)[0]
damga = datetime.now().strftime("%Y%m%d_%H%M")
st.download_button(
    "⬇️  Doldurulmuş dosyayı indir",
    data=cikti,
    file_name=f"{ad}_DOLDURULDU_{damga}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

st.divider()

# ----------------------------------------------------------------------
# UYARILAR
# ----------------------------------------------------------------------
if sonuc.uyarilar:
    st.subheader("⚠️ Uyarılar")
    for u in sonuc.uyarilar:
        st.warning(u)

# ----------------------------------------------------------------------
# OZET
# ----------------------------------------------------------------------
st.subheader("Cari dönem özeti")
k1, k2, k3 = st.columns(3)

etiket = "Dönem Kârı" if sonuc.donem_sonucu >= 0 else "Dönem Zararı"
hedef_hesap = "590" if sonuc.donem_sonucu >= 0 else "591"
k1.metric(
    f"{etiket} ({hedef_hesap})",
    f"{abs(sonuc.donem_sonucu):,.2f} ₺",
)
k2.metric("Aktif Toplam", f"{sonuc.aktif_toplam:,.2f} ₺")

fark = round(sonuc.aktif_toplam - sonuc.pasif_toplam, 2)
k3.metric(
    "Pasif Toplam",
    f"{sonuc.pasif_toplam:,.2f} ₺",
    delta=f"Fark: {fark:,.2f}" if abs(fark) > 0.5 else "Denk ✓",
    delta_color="inverse" if abs(fark) > 0.5 else "off",
)

# ----------------------------------------------------------------------
# ADIM 1 - TESHIS
# ----------------------------------------------------------------------
if sonuc.teshisler:
    st.subheader("Adım 1 — 7/A Maliyet Hesapları Teşhisi")
    satirlar = []
    for t in sonuc.teshisler:
        satirlar.append(
            {
                "Dönem": getattr(t, "donem", ""),
                "Hesap": t.kod,
                "Açıklama": t.ad,
                "7'li Bakiye": f"{t.yedili_bakiye:,.2f}",
                "Yansıtma": f"{t.yansitma_bakiye:,.2f}",
                "Kalan": f"{t.kalan:,.2f}",
                "Hedef": t.hedef,
                "Durum": t.durum,
            }
        )
    st.dataframe(satirlar, use_container_width=True, hide_index=True)
    st.caption(
        "Kalan = 7'li borç bakiyesi − yansıtma hesabının alacak bakiyesi. "
        "Yansıtma tam yapılmışsa kalan sıfır çıkar ve aktarım yapılmaz "
        "(mükerrer kayıt oluşmaz)."
    )

# ----------------------------------------------------------------------
# ADIM 2 - AKTARIM
# ----------------------------------------------------------------------
if sonuc.aktarimlar:
    st.subheader("Adım 2 — Yapılan Aktarımlar")
    st.dataframe(
        [
            {
                "Dönem": a.get("donem", ""),
                "Kaynak": a["kaynak"],
                "Hedef": a["hedef"],
                "Tutar": f"{a['tutar']:,.2f}",
                "Açıklama": a["aciklama"],
            }
            for a in sonuc.aktarimlar
        ],
        use_container_width=True,
        hide_index=True,
    )
else:
    st.info("Aktarılacak 7/A bakiyesi bulunmadı (yansıtmalar tam yapılmış).")

# ----------------------------------------------------------------------
# ESLESMEYEN HESAPLAR
# ----------------------------------------------------------------------
if sonuc.eslesmeyen:
    st.subheader("❗ Şablonda karşılığı bulunamayan hesaplar")
    st.error(
        "Aşağıdaki hesaplar mizanda bakiyeli ancak şablonda satırı yok. "
        "Bu tutarlar **tabloya yazılamadı** — kontrol ediniz."
    )
    st.dataframe(
        [
            {
                "Mizan": e["donem"],
                "Hesap": e["kod"],
                "Hesap Adı": e["ad"],
                "Tutar": f"{e['tutar']:,.2f}",
            }
            for e in sonuc.eslesmeyen
        ],
        use_container_width=True,
        hide_index=True,
    )

st.divider()
st.caption(
    "Bu araç mizandaki bakiyeleri tabloya aktarır. Sonuçların mevzuata "
    "uygunluğunun kontrolü kullanıcının sorumluluğundadır."
)
