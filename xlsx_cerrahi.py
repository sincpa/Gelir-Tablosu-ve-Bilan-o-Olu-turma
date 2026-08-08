"""
xlsx_cerrahi.py
---------------
Excel dosyasina XML seviyesinde cerrahi mudahale yapar.

NEDEN BOYLE BIR SEY GEREKLI?
openpyxl ile bir dosyayi acip kaydettiginizde; resimler, grafikler, hucre
aciklamalari (comment), yazici ayarlari ve bazi bicimlendirmeler KAYBOLUR.
Bu modul dosyayi bir ZIP arsivi olarak acar, sadece degistirmemiz gereken
hucrelerin XML'ini duzenler, geri kalan her seyi byte byte aynen korur.

Sonuc: renk, yazi tipi, kalinlik, kenarlik, sutun genisligi, resim, aciklama
       -> hicbiri degismez. Sadece sayilar degisir.
"""

import re
import shutil
import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

# Excel'in ana XML isim alani (namespace)
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

# ET'nin ciktiya "ns0:" gibi onekler eklemesini engelliyoruz
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)


def _q(tag: str) -> str:
    """Ana isim alanindaki etiket adini tam haline cevirir."""
    return f"{{{NS_MAIN}}}{tag}"


def sutun_harfi_to_no(harf: str) -> int:
    """'A' -> 1, 'B' -> 2, 'AA' -> 27"""
    sonuc = 0
    for ch in harf:
        sonuc = sonuc * 26 + (ord(ch.upper()) - ord("A") + 1)
    return sonuc


def hucre_ref_parcala(ref: str):
    """'C15' -> ('C', 15)"""
    m = re.match(r"([A-Za-z]+)(\d+)", ref)
    if not m:
        raise ValueError(f"Gecersiz hucre adresi: {ref}")
    return m.group(1).upper(), int(m.group(2))


class ExcelCerrah:
    """
    Bir .xlsx dosyasini ZIP olarak tutar, secili hucreleri duzenler,
    istenen sayfalari siler ve dosyayi geri verir.
    """

    def __init__(self, kaynak):
        """kaynak: dosya yolu (str) veya bytes/BytesIO"""
        if isinstance(kaynak, (bytes, bytearray)):
            kaynak = BytesIO(kaynak)

        # Arsivdeki her parcayi hafizaya aliyoruz
        self.parcalar = {}
        self.sira = []  # orijinal sira korunsun
        with zipfile.ZipFile(kaynak, "r") as z:
            for isim in z.namelist():
                self.parcalar[isim] = z.read(isim)
                self.sira.append(isim)

        # Sayfa adi -> XML parca yolu eslemesi
        self._sayfa_haritasi = self._sayfa_haritasi_kur()

        # Duzenlenmek uzere acilmis sayfa XML agaclari (onbellek)
        self._acik_sayfalar = {}

    # ------------------------------------------------------------------
    # Sayfa haritasi
    # ------------------------------------------------------------------
    def _sayfa_haritasi_kur(self):
        """workbook.xml + rels dosyalarindan sayfa adi -> xml yolu esler."""
        wb = ET.fromstring(self.parcalar["xl/workbook.xml"])
        rels = ET.fromstring(self.parcalar["xl/_rels/workbook.xml.rels"])

        # rId -> hedef dosya
        rid_hedef = {}
        for rel in rels:
            rid_hedef[rel.get("Id")] = rel.get("Target")

        harita = {}
        sheets = wb.find(_q("sheets"))
        if sheets is None:
            return harita
        for sh in sheets:
            ad = sh.get("name")
            rid = sh.get(f"{{{NS_REL}}}id")
            hedef = rid_hedef.get(rid, "")
            if hedef.startswith("/"):
                yol = hedef.lstrip("/")
            else:
                yol = "xl/" + hedef.lstrip("./")
            harita[ad] = yol
        return harita

    def sayfa_adlari(self):
        return list(self._sayfa_haritasi.keys())

    def sayfa_var_mi(self, ad: str) -> bool:
        """Buyuk/kucuk harfe DUYARLI kontrol."""
        return ad in self._sayfa_haritasi

    # ------------------------------------------------------------------
    # Sayfa XML'ini ac / kaydet
    # ------------------------------------------------------------------
    def _sayfa_ac(self, sayfa_adi: str):
        if sayfa_adi in self._acik_sayfalar:
            return self._acik_sayfalar[sayfa_adi]
        yol = self._sayfa_haritasi[sayfa_adi]
        kok = ET.fromstring(self.parcalar[yol])
        self._acik_sayfalar[sayfa_adi] = (kok, yol)
        return kok, yol

    def _acik_sayfalari_yaz(self):
        for sayfa_adi, (kok, yol) in self._acik_sayfalar.items():
            self.parcalar[yol] = ET.tostring(kok, encoding="UTF-8", xml_declaration=True)

    # ------------------------------------------------------------------
    # Hucre okuma / yazma
    # ------------------------------------------------------------------
    def _satir_bul_veya_olustur(self, kok, satir_no: int):
        sheet_data = kok.find(_q("sheetData"))
        if sheet_data is None:
            sheet_data = ET.SubElement(kok, _q("sheetData"))

        hedef = None
        ekleme_indeksi = len(sheet_data)
        for i, satir in enumerate(sheet_data):
            r = satir.get("r")
            if r is None:
                continue
            r = int(r)
            if r == satir_no:
                hedef = satir
                break
            if r > satir_no:
                ekleme_indeksi = i
                break

        if hedef is None:
            hedef = ET.Element(_q("row"), {"r": str(satir_no)})
            sheet_data.insert(ekleme_indeksi, hedef)
        return hedef

    def _stil_tahmin_et(self, kok, sutun: str, satir_no: int):
        """
        Hucre XML'de hic yoksa olustururken, ayni sutundaki EN YAKIN
        hucrenin stilini kullaniriz. Boylece sayi formati (binlik ayirac,
        ondalik, renk) bozulmaz.
        """
        sheet_data = kok.find(_q("sheetData"))
        if sheet_data is None:
            return None
        en_yakin_s, en_yakin_mesafe = None, None
        for satir in sheet_data:
            r = satir.get("r")
            if r is None:
                continue
            r = int(r)
            mesafe = abs(r - satir_no)
            if en_yakin_mesafe is not None and mesafe >= en_yakin_mesafe:
                continue
            for c in satir:
                ref = c.get("r")
                if not ref:
                    continue
                h, _ = hucre_ref_parcala(ref)
                if h == sutun and c.get("s") is not None:
                    en_yakin_s = c.get("s")
                    en_yakin_mesafe = mesafe
                    break
        return en_yakin_s

    def _hucre_bul_veya_olustur(self, kok, ref: str):
        sutun, satir_no = hucre_ref_parcala(ref)
        hedef_no = sutun_harfi_to_no(sutun)
        satir = self._satir_bul_veya_olustur(kok, satir_no)

        ekleme_indeksi = len(satir)
        for i, c in enumerate(satir):
            c_ref = c.get("r")
            if not c_ref:
                continue
            c_sutun, _ = hucre_ref_parcala(c_ref)
            c_no = sutun_harfi_to_no(c_sutun)
            if c_no == hedef_no:
                return c
            if c_no > hedef_no:
                ekleme_indeksi = i
                break

        yeni = ET.Element(_q("c"), {"r": ref})
        stil = self._stil_tahmin_et(kok, sutun, satir_no)
        if stil is not None:
            yeni.set("s", stil)
        satir.insert(ekleme_indeksi, yeni)
        return yeni

    def hucre_oku_ham(self, sayfa_adi: str, ref: str):
        """
        Hucrenin ham degerini dondurur. Formul varsa formul metnini
        ('=' isaretsiz) dondurur. Sadece kontrol amaclidir.
        """
        kok, _ = self._sayfa_ac(sayfa_adi)
        sutun, satir_no = hucre_ref_parcala(ref)
        sheet_data = kok.find(_q("sheetData"))
        if sheet_data is None:
            return None
        for satir in sheet_data:
            if satir.get("r") != str(satir_no):
                continue
            for c in satir:
                if c.get("r") == ref:
                    f = c.find(_q("f"))
                    if f is not None:
                        return "=" + (f.text or "")
                    v = c.find(_q("v"))
                    return v.text if v is not None else None
        return None

    def hucre_yaz(self, sayfa_adi: str, ref: str, deger):
        """
        Hucreye SAYI yazar. deger None ise hucreyi bosaltir.

        ONEMLI: Bu islem hucrenin 's' (stil) niteligine dokunmaz.
        Yani renk, yazi tipi, kalinlik, sayi formati aynen kalir.
        """
        kok, _ = self._sayfa_ac(sayfa_adi)
        c = self._hucre_bul_veya_olustur(kok, ref)

        # Mevcut icerigi temizle (formul, deger, satir ici metin)
        for cocuk in list(c):
            c.remove(cocuk)
        # Metin tipi isaretini kaldir (sayi yazacagiz)
        if "t" in c.attrib:
            del c.attrib["t"]

        if deger is None:
            return  # bos hucre: sadece stil kalir

        v = ET.SubElement(c, _q("v"))
        # Cok kucuk kalintilari sifirla, gereksiz ondalik uretme
        if isinstance(deger, float) and abs(deger) < 0.005:
            deger = 0.0
        v.text = repr(round(float(deger), 2)) if isinstance(deger, float) else str(deger)

    # ------------------------------------------------------------------
    # Sayfa silme
    # ------------------------------------------------------------------
    def sayfa_sil(self, sayfa_adi: str):
        """Sayfayi workbook.xml, rels, content-types ve arsivden kaldirir."""
        if sayfa_adi not in self._sayfa_haritasi:
            return False

        yol = self._sayfa_haritasi[sayfa_adi]

        # Acik onbellekten dus
        self._acik_sayfalar.pop(sayfa_adi, None)

        # 1) workbook.xml icindeki <sheet> kaydini kaldir
        wb = ET.fromstring(self.parcalar["xl/workbook.xml"])
        sheets = wb.find(_q("sheets"))
        silinen_rid = None
        silinen_idx = None
        for i, sh in enumerate(list(sheets)):
            if sh.get("name") == sayfa_adi:
                silinen_rid = sh.get(f"{{{NS_REL}}}id")
                silinen_idx = i
                sheets.remove(sh)
                break

        # 1b) <definedNames> icindeki bu sayfaya ait kayitlari temizle.
        # localSheetId, <sheets> listesindeki 0-tabanli sira numarasidir.
        # Temizlenmezse silinen sayfayi isaret eden veya artik yanlis
        # sayfayi isaret eden (kaymis) kayitlar kalir; Excel bunu
        # "dosya bozuk" olarak degerlendirip acmayi reddedebilir.
        if silinen_idx is not None:
            defined_names = wb.find(_q("definedNames"))
            if defined_names is not None:
                for dn in list(defined_names):
                    lsid = dn.get("localSheetId")
                    if lsid is None:
                        continue
                    lsid = int(lsid)
                    if lsid == silinen_idx:
                        defined_names.remove(dn)
                    elif lsid > silinen_idx:
                        dn.set("localSheetId", str(lsid - 1))
                if len(defined_names) == 0:
                    wb.remove(defined_names)

        # Excel'in acilista tum formulleri yeniden hesaplamasini saglar
        calc = wb.find(_q("calcPr"))
        if calc is None:
            calc = ET.SubElement(wb, _q("calcPr"))
        calc.set("fullCalcOnLoad", "1")

        self.parcalar["xl/workbook.xml"] = ET.tostring(
            wb, encoding="UTF-8", xml_declaration=True
        )

        # 2) workbook.xml.rels icindeki iliskiyi kaldir
        if silinen_rid:
            rels_xml = self.parcalar["xl/_rels/workbook.xml.rels"]
            rels = ET.fromstring(rels_xml)
            for rel in list(rels):
                if rel.get("Id") == silinen_rid:
                    rels.remove(rel)
                    break
            ET.register_namespace("", NS_PKG_REL)
            self.parcalar["xl/_rels/workbook.xml.rels"] = ET.tostring(
                rels, encoding="UTF-8", xml_declaration=True
            )
            ET.register_namespace("", NS_MAIN)

        # 3) [Content_Types].xml icindeki kaydi kaldir
        ct = ET.fromstring(self.parcalar["[Content_Types].xml"])
        for ov in list(ct):
            if ov.get("PartName") == "/" + yol:
                ct.remove(ov)
                break
        ET.register_namespace("", NS_CT)
        self.parcalar["[Content_Types].xml"] = ET.tostring(
            ct, encoding="UTF-8", xml_declaration=True
        )
        ET.register_namespace("", NS_MAIN)

        # 4) Sayfa dosyasini ve varsa kendi rels dosyasini arsivden cikar
        for silinecek in [yol, yol.replace("worksheets/", "worksheets/_rels/") + ".rels"]:
            self.parcalar.pop(silinecek, None)
            if silinecek in self.sira:
                self.sira.remove(silinecek)

        del self._sayfa_haritasi[sayfa_adi]
        return True

    # ------------------------------------------------------------------
    # Kaydet
    # ------------------------------------------------------------------
    def _calc_chain_temizle(self):
        """
        calcChain.xml, formullerin hesaplanma sirasini tutar. Biz formul
        iceren hucrelere yazdiysak bu dosya tutarsiz kalir ve Excel
        'dosya bozuk' uyarisi verebilir. Silmek guvenlidir; Excel yeniden
        olusturur.
        """
        yol = "xl/calcChain.xml"
        if yol not in self.parcalar:
            return
        self.parcalar.pop(yol, None)
        if yol in self.sira:
            self.sira.remove(yol)

        ct = ET.fromstring(self.parcalar["[Content_Types].xml"])
        for ov in list(ct):
            if ov.get("PartName") == "/" + yol:
                ct.remove(ov)
                break
        ET.register_namespace("", NS_CT)
        self.parcalar["[Content_Types].xml"] = ET.tostring(
            ct, encoding="UTF-8", xml_declaration=True
        )
        ET.register_namespace("", NS_MAIN)

    def tam_hesaplama_iste(self):
        """Excel dosyayi acinca tum formulleri yeniden hesaplasin."""
        wb = ET.fromstring(self.parcalar["xl/workbook.xml"])
        calc = wb.find(_q("calcPr"))
        if calc is None:
            calc = ET.SubElement(wb, _q("calcPr"))
        calc.set("fullCalcOnLoad", "1")
        self.parcalar["xl/workbook.xml"] = ET.tostring(
            wb, encoding="UTF-8", xml_declaration=True
        )

    def kaydet_bytes(self) -> bytes:
        self._acik_sayfalari_yaz()
        self._calc_chain_temizle()
        self.tam_hesaplama_iste()

        cikti = BytesIO()
        with zipfile.ZipFile(cikti, "w", zipfile.ZIP_DEFLATED) as z:
            for isim in self.sira:
                if isim in self.parcalar:
                    z.writestr(isim, self.parcalar[isim])
            # Sirada olmayan yeni parca varsa (nadiren) sona ekle
            for isim, veri in self.parcalar.items():
                if isim not in self.sira:
                    z.writestr(isim, veri)
        return cikti.getvalue()

    def kaydet(self, hedef_yol: str):
        with open(hedef_yol, "wb") as f:
            f.write(self.kaydet_bytes())
