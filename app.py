#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LİDERFORM ANALİZ WEB UYGULAMASI
Flask tabanlı web arayüzü — URL gir → tarayıcıda sonuç gör
"""

import re, sys, os
from flask import Flask, render_template, request, jsonify

# Core modülü import et
sys.path.insert(0, os.path.dirname(__file__))
import liderform_core as core

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "liderform2026")

# ─── Yardımcı: core'dan web verisi üret ──────────────────────
def analiz_yap(url):
    """
    Mevcut core mantığını kullanarak analiz yapar,
    Excel yerine dict/list döndürür.
    """
    import re, time

    if "liderform.com.tr" not in url:
        raise ValueError("Geçerli bir liderform.com.tr linki girin.")

    meta = core.url_meta(url)
    if not meta:
        raise ValueError("Link formatı hatalı. Örnek: https://liderform.com.tr/program/2026-03-28/istanbul/1")

    # 1. Program sayfası
    prog_url = (f"https://liderform.com.tr/program/"
                f"{meta['tarih_iso']}/{meta['sehir_raw']}/{meta['kosu_no']}")
    soup   = core.fetch_soup(prog_url)
    rinfo  = core.race_info(soup)
    horses = core.parse_ana_sayfa(soup)

    if not horses:
        raise ValueError("At verisi bulunamadı. Koşu henüz yayınlanmamış olabilir.")

    # 2. Son koşu linkleri
    son_kosu_map = core.son_kosu_linklerini_cek(
        meta["tarih_iso"], meta["sehir_raw"], meta["kosu_no"]
    )

    # 3. Accurace verisi
    acc_data = core.fetch_accurace_per_horse(horses, son_kosu_map)

    # 4. Analiz sonuçlarını üret (build_excel'in içindeki mantık)
    result = _build_web_result(meta, rinfo, horses, acc_data, son_kosu_map)
    return result


def _build_web_result(meta, rinfo, horses, acc_data, son_kosu_map):
    """
    build_excel yerine web için dict döndürür.
    Core'daki hesap mantığını tekrar kullanır.
    """
    import re

    kosu_mesafe_str = rinfo.get("mesafe", "1200 m").replace(" m","").strip()
    try:
        kosu_mesafe = int(kosu_mesafe_str)
    except:
        kosu_mesafe = 1200

    _tempo_mes = "1000m" if kosu_mesafe >= 1600 else "800m"

    # ── ACCURACE veri normalize (core mantığı) ───────────────
    def basari_skoru(veri):
        sm = veri.get("son_mesafe", "")
        sureler = veri.get("sureler", {})
        final_sira = sureler.get(sm, {}).get("sira", 99)
        if not isinstance(final_sira, int): final_sira = 99
        s400_str = veri.get("son400", "") or ""
        s400_sn = core.sure_to_sec(s400_str) if s400_str else 99.0
        return (final_sira, -(s400_sn or 99.0))

    for horse in horses:
        at_adi = horse["adi"]
        acc    = acc_data.get(at_adi, {})
        # son400/son600 hesapla
        for kt in ["son", "onceki"]:
            veri = acc.get(kt, {})
            sureler   = veri.get("sureler", {})
            son_mes   = veri.get("son_mesafe", "")
            s400, s600 = core.hesapla_son_400_600(sureler, son_mes)
            veri["son400"] = s400
            veri["son600"] = s600
        # en iyi koşu seç
        son_veri    = acc.get("son", {})
        onceki_veri = acc.get("onceki", {})
        skor_son    = basari_skoru(son_veri)
        skor_onceki = basari_skoru(onceki_veri)
        son_veri_bos = (not son_veri.get("sureler") and
                        not son_veri.get("son400") and
                        not son_veri.get("son600"))
        if son_veri_bos and onceki_veri.get("sureler"):
            en_iyi_veri = onceki_veri
            en_iyi_etik = "ÖNCEKİ (son boş)"
        elif skor_son <= skor_onceki:
            en_iyi_veri = son_veri
            en_iyi_etik = "SON"
        else:
            en_iyi_veri = onceki_veri
            en_iyi_etik = "ÖNCEKİ"
        acc["en_iyi"]      = en_iyi_veri
        acc["en_iyi_etik"] = en_iyi_etik

    # ── TEMPO ön hesabı (core._tempo_karar mantığı) ──────────
    _TEMPO_MES = "1000m" if kosu_mesafe >= 1600 else "800m"
    _TEMPO_TOL = 1.5 if _TEMPO_MES == "1000m" else 1.0
    _BASARI_SIRA = 4

    at_tempo_pre = []
    for horse in horses:
        an = horse["adi"]
        ac = acc_data.get(an, {})
        so = ac.get("son", {})
        on = ac.get("onceki", {})

        def gs(v, m): return v.get("sureler",{}).get(m,{}).get("sira","")
        def gsn(v, m):
            s = v.get("sureler",{}).get(m,{}).get("sure","")
            return core.sure_to_sec(s) if s else None

        s8s = gs(so, _TEMPO_MES); s8o = gs(on, _TEMPO_MES)
        su8s = gsn(so, _TEMPO_MES); su8o = gsn(on, _TEMPO_MES)
        sf_son = so.get("sureler",{}).get(so.get("son_mesafe",""),{}).get("sira","")
        s400s = core.sure_to_sec(so.get("son400","") or "") if so.get("son400") else None
        oyun_k = (s8s == 1 and s8o == 1)
        pot_k  = (s8s == 1 or s8o == 1)
        sure_baz = (not pot_k and isinstance(s8s, int) and s8s <= 3 and su8s is not None)
        on_grup = (isinstance(s8s, int) and s8s <= 4)
        geride  = (isinstance(s8s, int) and s8s >= 5)
        best = min([s for s in [su8s, su8o] if s is not None], default=None)
        at_tempo_pre.append({
            "adi": an, "no": horse["no"],
            "s8_son": s8s, "s8_onc": s8o,
            "sf_son": sf_son,
            "sure8_son": su8s, "sure8_onc": su8o, "best_sure": best,
            "son400_sn": s400s,
            "oyun_kurucu": oyun_k, "pot_oyun_kurucu": pot_k,
            "sure_bazli_aday": sure_baz,
            "on_grup": on_grup, "geride": geride,
            "sprinter": False,
        })

    # Tahmini tempo
    onde = []
    for t in at_tempo_pre:
        for su, si in [(t["sure8_son"], t["s8_son"]), (t["sure8_onc"], t["s8_onc"])]:
            if isinstance(si, int) and si <= 3 and su is not None:
                onde.append(su)
        if t["sure_bazli_aday"] and t["sure8_son"] is not None:
            onde.append(t["sure8_son"])
    tahmini_tempo_sn = sum(onde)/len(onde) if onde else None
    tahmini_tempo_str = core.sec_to_str(tahmini_tempo_sn) if tahmini_tempo_sn else "?"

    # Sprinter
    s400_lst = [t["son400_sn"] for t in at_tempo_pre if t["son400_sn"]]
    ort_s400 = sum(s400_lst)/len(s400_lst) if s400_lst else None
    for t in at_tempo_pre:
        t["sprinter"] = t["geride"] and bool(t["son400_sn"] and ort_s400 and t["son400_sn"] < ort_s400)

    oyun_k_list = [t for t in at_tempo_pre if t["oyun_kurucu"]]
    pot_k_list  = [t for t in at_tempo_pre if t["pot_oyun_kurucu"] and not t["oyun_kurucu"]]
    sure_baz_lst = sorted([t for t in at_tempo_pre if t["sure_bazli_aday"]],
                          key=lambda x: x["sure8_son"] if x["sure8_son"] else 999)
    if not oyun_k_list and not pot_k_list and sure_baz_lst:
        for t in at_tempo_pre:
            if t["adi"] == sure_baz_lst[0]["adi"]:
                t["pot_oyun_kurucu"] = True
        pot_k_list = [sure_baz_lst[0]]

    if len(oyun_k_list) >= 2:
        tempo_karar = "HIZLI"
        tempo_aciklama = f"{', '.join(t['adi'] for t in oyun_k_list)} birbirini sıkıştıracak → yüksek tempo."
    elif len(oyun_k_list) == 1:
        rakipler = [t for t in pot_k_list if isinstance(t["s8_son"], int) and t["s8_son"] <= 3]
        if rakipler:
            tempo_karar = "HIZLI"
            tempo_aciklama = f"{oyun_k_list[0]['adi']} öne gider, {', '.join(t['adi'] for t in rakipler)} baskı yapar → tempo yükselir."
        else:
            tempo_karar = "YAVAS"
            tempo_aciklama = f"{oyun_k_list[0]['adi']} tek başına önde gidecek, rakipsiz → kontrollü tempo."
    elif len(pot_k_list) >= 2:
        tempo_karar = "HIZLI"
        tempo_aciklama = f"Birden fazla muhtemel kurucu: {', '.join(t['adi'] for t in pot_k_list[:3])} → yüksek tempo."
    elif len(pot_k_list) == 1:
        tempo_karar = "YAVAS"
        tempo_aciklama = f"{pot_k_list[0]['adi']} muhtemel kurucu, rakipsiz gidebilir → kontrollü tempo."
    else:
        tempo_karar = "BELIRSIZ"
        tempo_aciklama = "Önde gidecek at belirlenemiyor — veri yetersiz."

    at_roller = {t["adi"]: t for t in at_tempo_pre}

    # ── TEMPO UYUM skoru ─────────────────────────────────────
    tempo_uyum = {}
    for horse in horses:
        an = horse["adi"]
        ac = acc_data.get(an, {})
        uyum_skoru = 0
        uyumsuz_kez = 0
        for kt in ["son", "onceki"]:
            veri = ac.get(kt, {})
            sure8_str = veri.get("sureler", {}).get(_tempo_mes, {}).get("sure", "")
            sf_m = veri.get("son_mesafe", "")
            sf = veri.get("sureler", {}).get(sf_m, {}).get("sira", "")
            sure8 = core.sure_to_sec(sure8_str) if sure8_str else None
            if sure8 is None or tahmini_tempo_sn is None:
                continue
            basarili = isinstance(sf, int) and sf <= _BASARI_SIRA
            kaldirabilir = sure8 <= tahmini_tempo_sn + _TEMPO_TOL
            if kaldirabilir:
                uyum_skoru += 2 if basarili else 1
            else:
                uyumsuz_kez += 1
        tempo_uyum[an] = {
            "uyum_skoru": uyum_skoru,
            "uyumlu": uyum_skoru >= 2,
            "uyumsuz": uyumsuz_kez >= 2,
        }

    # ── SINIF AVANTAJI ────────────────────────────────────────
    is_arap    = "ARAP"    in rinfo.get("tip","").upper()
    is_ingiliz = not is_arap

    bugun_hp = sorted(
        [h.get("hp",0) for h in horses if isinstance(h.get("hp"),(int,float)) and h.get("hp",0)>0],
        reverse=True
    )[:5]
    bugun_ort_hp = sum(bugun_hp)/len(bugun_hp) if bugun_hp else 0

    sinif_avantaj = {}
    for horse in horses:
        an = horse["adi"]
        hp = horse.get("hp",0) or 0
        try: hp = float(hp)
        except: hp = 0
        if bugun_ort_hp > 0 and hp > 0:
            fark = hp - bugun_ort_hp
            if fark >= 3:
                sinif_avantaj[an] = {"var": True, "lbl": f"Sınıf üstü (+{fark:.0f}hp)", "fark_hp": fark}
            elif fark <= -3:
                sinif_avantaj[an] = {"var": False, "lbl": f"Sınıf altı ({fark:.0f}hp)", "seviye": "dezavantaj", "fark_hp": fark}
            else:
                sinif_avantaj[an] = {"var": False, "lbl": "Sınıf benzer", "seviye": "normal", "fark_hp": fark}
        else:
            sinif_avantaj[an] = {"var": False, "lbl": "-", "seviye": "bilinmiyor", "fark_hp": 0}

    # ── KARMA CİNSİYET ────────────────────────────────────────
    disi_sayisi  = sum(1 for h in horses if h.get("cinsiyet") == "dişi")
    erkek_sayisi = sum(1 for h in horses if h.get("cinsiyet") == "erkek")
    karma_cinsiyet = is_ingiliz and disi_sayisi > 0 and erkek_sayisi > 0

    # ── PUANLAMA ─────────────────────────────────────────────
    W_SON400 = 30; W_SON600 = 20; W_FORM = 20; W_MESAFE = 15; W_STIL = 15
    KGS_ESIK = 100

    # Ortalama son400
    s400_lst2 = []
    for horse in horses:
        ac = acc_data.get(horse["adi"],{})
        ei = ac.get("en_iyi",{})
        s4 = core.sure_to_sec(ei.get("son400","") or "")
        if s4: s400_lst2.append(s4)
    ort_400 = sum(s400_lst2)/len(s400_lst2) if s400_lst2 else None

    s600_lst = []
    for horse in horses:
        ac = acc_data.get(horse["adi"],{})
        ei = ac.get("en_iyi",{})
        s6 = core.sure_to_sec(ei.get("son600","") or "")
        if s6: s600_lst.append(s6)
    ort_600 = sum(s600_lst)/len(s600_lst) if s600_lst else None

    sonuclar = []
    for horse in horses:
        at_adi = horse["adi"]
        acc    = acc_data.get(at_adi, {})
        en_iyi = acc.get("en_iyi", {})

        veri_eksik = not en_iyi.get("sureler") and not en_iyi.get("son400")
        puan = 0
        detaylar = []
        eleme_nedenleri = []

        # Son400 puanı
        s4 = core.sure_to_sec(en_iyi.get("son400","") or "")
        if s4 and ort_400:
            fark = s4 - ort_400
            if fark <= -0.5:
                puan += W_SON400
                detaylar.append(f"✅ Son400 çok hızlı ({en_iyi.get('son400','')})")
            elif fark <= 0.3:
                puan += round(W_SON400 * 0.7)
                detaylar.append(f"➡️ Son400 ortalamanın altında ({en_iyi.get('son400','')})")
            elif fark <= 1.0:
                puan += round(W_SON400 * 0.35)
                detaylar.append(f"⚠️ Son400 yavaş ({en_iyi.get('son400','')})")
                eleme_nedenleri.append("Son400 yavaş")
            else:
                puan += 0
                detaylar.append(f"❌ Son400 çok yavaş ({en_iyi.get('son400','')})")
                eleme_nedenleri.append("Son400 çok yavaş")
        elif veri_eksik:
            puan += W_SON400 // 2

        # Son600 puanı
        s6 = core.sure_to_sec(en_iyi.get("son600","") or "")
        if s6 and ort_600:
            fark6 = s6 - ort_600
            if fark6 <= -0.5:
                puan += W_SON600
                detaylar.append(f"✅ Son600 çok hızlı ({en_iyi.get('son600','')})")
            elif fark6 <= 0.5:
                puan += round(W_SON600 * 0.7)
                detaylar.append(f"➡️ Son600 ortalamanın altında ({en_iyi.get('son600','')})")
            elif fark6 <= 1.5:
                puan += round(W_SON600 * 0.35)
                detaylar.append(f"⚠️ Son600 yavaş ({en_iyi.get('son600','')})")
                eleme_nedenleri.append("Son600 yavaş")
            else:
                puan += 0
                detaylar.append(f"❌ Son600 çok yavaş ({en_iyi.get('son600','')})")
                eleme_nedenleri.append("Son600 çok yavaş")
        elif veri_eksik:
            puan += W_SON600 // 2

        # Form puanı
        acc_son    = acc.get("son",{})
        acc_onceki = acc.get("onceki",{})
        sf_son_m = acc_son.get("son_mesafe","")
        sf_onc_m = acc_onceki.get("son_mesafe","") or acc_onceki.get("mesafe","")
        sira_son = acc_son.get("sureler",{}).get(sf_son_m,{}).get("sira","")
        sira_onc = acc_onceki.get("sureler",{}).get(sf_onc_m,{}).get("sira","")
        if isinstance(sira_son, int) and isinstance(sira_onc, int):
            fark_form = sira_onc - sira_son  # pozitif = iyileşme
            if fark_form >= 2:
                puan += W_FORM
                genel_form = f"📈 YÜKSELİŞ ({sira_onc}→{sira_son})"
                detaylar.append(f"✅ Form yükselen ({sira_onc}→{sira_son})")
            elif fark_form >= 0:
                puan += round(W_FORM * 0.7)
                genel_form = f"➡️ STABİL ({sira_onc}→{sira_son})"
                detaylar.append(f"➡️ Form stabil ({sira_onc}→{sira_son})")
            elif fark_form >= -2:
                puan += round(W_FORM * 0.4)
                genel_form = f"📉 DÜŞÜŞ ({sira_onc}→{sira_son})"
                detaylar.append(f"⚠️ Form düşen ({sira_onc}→{sira_son})")
                eleme_nedenleri.append("Form düşüşte")
            else:
                puan += 0
                genel_form = f"📉 SERT DÜŞÜŞ ({sira_onc}→{sira_son})"
                detaylar.append(f"❌ Form sert düşüş ({sira_onc}→{sira_son})")
                eleme_nedenleri.append("Form sert düşüş")
        else:
            puan += W_FORM // 2
            genel_form = "⚪ VERİ YOK"

        # Mesafe uyumu
        mesafe_uyumsuz = False
        if is_arap:
            en_iyi_mes_str = en_iyi.get("son_mesafe","") or en_iyi.get("mesafe","")
            if en_iyi_mes_str:
                en_iyi_m = int(re.sub(r"[^0-9]","",en_iyi_mes_str))
                fark_mes = abs(en_iyi_m - kosu_mesafe)
                if fark_mes > 200:
                    eleme_nedenleri.append(f"Arap atı: en iyi koşu mesafesi çok farklı ({en_iyi_mes_str} vs {kosu_mesafe}m)")
                    mesafe_uyumsuz = True

        # Stil tespiti
        rol = at_roller.get(at_adi, {})
        is_sprinter  = rol.get("sprinter", False)
        is_geride_at = rol.get("geride", False)
        is_on_grup   = rol.get("on_grup", False)
        is_oyun_kur  = rol.get("oyun_kurucu", False) or rol.get("pot_oyun_kurucu", False)

        if is_oyun_kur:   stil_lbl = "OYN.KUR."
        elif is_sprinter: stil_lbl = "SPRİNTER"
        elif is_geride_at: stil_lbl = "GERİDE"
        elif is_on_grup:  stil_lbl = "ÖN GRUP"
        else:              stil_lbl = "BELİRSİZ"

        # Stil/tempo puan
        if tempo_karar == "HIZLI":
            if is_sprinter or is_geride_at:
                puan += W_STIL
                detaylar.append(f"✅ HIZLI tempoda avantajlı: {stil_lbl} (+{W_STIL}p)")
            elif is_on_grup or is_oyun_kur:
                puan += round(W_STIL * 0.4)
                detaylar.append(f"⚠️ HIZLI tempoda zor: {stil_lbl}")
            else:
                puan += W_STIL // 2
        elif tempo_karar == "YAVAS":
            if is_on_grup or is_oyun_kur:
                puan += W_STIL
                detaylar.append(f"✅ YAVAS tempoda avantajlı: {stil_lbl} (+{W_STIL}p)")
            elif is_sprinter or is_geride_at:
                puan += 0
                detaylar.append(f"❌ YAVAS tempoda dezavantajlı: {stil_lbl}")
            else:
                puan += W_STIL // 2
        else:
            puan += W_STIL // 2

        # Yavaş tempo + geride eleme
        yavas_geride_eleme = False
        if tempo_karar == "YAVAS" and (is_geride_at or is_sprinter):
            tu = tempo_uyum.get(at_adi, {"uyumlu": False})
            if not tu.get("uyumlu", False):
                yavas_geride_eleme = True
                eleme_nedenleri.append(f"YAVAS tempo + {stil_lbl} karakter + tempo uyumsuz → DİREKT ELENDİ")

        # Tempo uyum bonus
        tu = tempo_uyum.get(at_adi, {})
        tempo_uyumlu  = tu.get("uyumlu", False)
        tempo_uyumsuz = tu.get("uyumsuz", False)
        if tempo_uyumlu:
            puan += 20
            detaylar.append("✅ TEMPO UYUMLU+ (+20 bonus)")
        elif tempo_uyumsuz:
            detaylar.append("❌ TEMPO UYUMSUZ")

        # Sınıf avantajı
        sb = sinif_avantaj.get(at_adi, {})
        if sb.get("var"):
            detaylar.append(f"⭐ {sb['lbl']}")
        elif sb.get("seviye") == "dezavantaj":
            detaylar.append(f"⚠️ {sb['lbl']}")

        # KGS
        kgs = 0
        try: kgs = int(str(horse.get("kgs","0") or "0").replace("+",""))
        except: pass
        at_kg = 0; kg_artis = 0
        try: at_kg = float(str(horse.get("kg",0) or 0).replace(",","."))
        except: pass

        # Eleme kararı
        elendi = False
        if kgs >= KGS_ESIK:
            elendi = True
            eleme_nedenleri.append(f"KGS {kgs} gün (100+ gün start yok)")
        elif is_arap and any("en iyi koşu mesafesi çok farklı" in n for n in eleme_nedenleri):
            elendi = True
        elif is_ingiliz and karma_cinsiyet and horse.get("cinsiyet") == "dişi":
            elendi = True
            eleme_nedenleri.append("İngiliz dişi at, karma cinsiyet koşusu")
        elif kosu_mesafe >= 1600 and at_kg > 60 and kg_artis >= 2:
            elendi = True
            eleme_nedenleri.append(f"1600m+ koşuda {at_kg}kg (+{kg_artis}kg artış)")
        elif yavas_geride_eleme:
            elendi = True
        elif tempo_uyumsuz:
            elendi = True

        # Mesafe uyumsuz + tempo yok → eleme
        elif mesafe_uyumsuz and not tempo_uyumlu:
            elendi = True
            eleme_nedenleri.append("Mesafe uyumsuz + tempo uyumu yok")
        else:
            son400_yavas = any("Son400" in n for n in eleme_nedenleri)
            son600_yavas = any("Son600" in n for n in eleme_nedenleri)
            if son400_yavas and son600_yavas and len(eleme_nedenleri) >= 2 and not tempo_uyumlu:
                elendi = True
            elif len(eleme_nedenleri) >= 3 and not tempo_uyumlu:
                elendi = True
            elif s4 and ort_400 and (s4 - ort_400) >= 1.0 and not tempo_uyumlu:
                elendi = True
                eleme_nedenleri.append(f"Son400 çok yavaş (ort.+{s4-ort_400:.1f}sn)")

        tempo_yildiz = tu.get("uyum_skoru", 0) >= 3

        sonuclar.append({
            "no":             horse["no"],
            "adi":            at_adi,
            "puan":           round(puan, 1),
            "elendi":         elendi,
            "eleme_nedenleri": eleme_nedenleri,
            "detaylar":       detaylar,
            "s400_str":       en_iyi.get("son400","") or "-",
            "s600_str":       en_iyi.get("son600","") or "-",
            "s400_sn":        s4,
            "s600_sn":        s6,
            "genel_form":     genel_form,
            "stil_lbl":       stil_lbl,
            "tempo_uyumlu":   tempo_uyumlu,
            "tempo_uyumsuz":  tempo_uyumsuz,
            "tempo_yildiz":   tempo_yildiz,
            "sinif_lbl":      sb.get("lbl","-"),
            "veri_eksik":     veri_eksik,
            "jokey":          horse.get("jokey",""),
            "gp":             horse.get("gp",""),
            "hp":             horse.get("hp",""),
            "kg":             horse.get("kg",""),
            "kgs":            kgs,
            "agf":            horse.get("agf_pct",""),
            "cinsiyet":       horse.get("cinsiyet",""),
            "s400_en_iyi":    False,
            "s600_kotu":      False,
        })

    # Puana göre sırala
    sonuclar.sort(key=lambda x: x["puan"], reverse=True)
    for i, s in enumerate(sonuclar, 1):
        s["sira"] = i

    # Son600 en kötü 2 → eleme
    gec_s600 = [(i, s["s600_sn"]) for i, s in enumerate(sonuclar)
                if s["s600_sn"] is not None and not s["elendi"]]
    kotu_s600 = sorted(gec_s600, key=lambda x: x[1], reverse=True)[:2]
    for idx, _ in kotu_s600:
        if not sonuclar[idx]["elendi"]:
            sonuclar[idx]["elendi"] = True
            sonuclar[idx]["s600_kotu"] = True
            sonuclar[idx]["eleme_nedenleri"].append(
                f"Son600 en kötü 2 at içinde ({sonuclar[idx]['s600_str']}) → ELENDİ"
            )

    # Son400 en iyi 3 → işaretle
    gec_s400 = [(i, s["s400_sn"]) for i, s in enumerate(sonuclar) if s["s400_sn"] is not None]
    iyi_s400 = sorted(gec_s400, key=lambda x: x[1])[:3]
    for idx, _ in iyi_s400:
        sonuclar[idx]["s400_en_iyi"] = True

    # TEMPO sayfası verileri
    oyun_kurucu_adlar = [f"{t['no']}-{t['adi']}" for t in at_tempo_pre if t["oyun_kurucu"]]
    sprinter_adlar    = [f"{t['no']}-{t['adi']}" for t in at_tempo_pre if t["sprinter"]]
    on_grup_adlar     = [f"{t['no']}-{t['adi']}" for t in at_tempo_pre if t["on_grup"] and not t["oyun_kurucu"]]
    geride_adlar      = [f"{t['no']}-{t['adi']}" for t in at_tempo_pre if t["geride"] and not t["sprinter"]]

    return {
        "meta":            meta,
        "rinfo":           rinfo,
        "kosu_mesafe":     kosu_mesafe,
        "sonuclar":        sonuclar,
        "tempo_karar":     tempo_karar,
        "tempo_aciklama":  tempo_aciklama,
        "tahmini_tempo":   tahmini_tempo_str,
        "tempo_mes":       _TEMPO_MES,
        "is_arap":         is_arap,
        "is_ingiliz":      is_ingiliz,
        "karma_cinsiyet":  karma_cinsiyet,
        "disi_sayisi":     disi_sayisi,
        "erkek_sayisi":    erkek_sayisi,
        "oyun_kurucu_adlar": oyun_kurucu_adlar,
        "sprinter_adlar":    sprinter_adlar,
        "on_grup_adlar":     on_grup_adlar,
        "geride_adlar":      geride_adlar,
        "at_sayisi":       len(horses),
        "elenen_sayisi":   sum(1 for s in sonuclar if s["elendi"]),
        "kalan_sayisi":    sum(1 for s in sonuclar if not s["elendi"]),
    }


# ─── ROUTES ──────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analiz", methods=["POST"])
def analiz():
    data = request.get_json()
    url  = (data or {}).get("url", "").strip()
    if not url:
        return jsonify({"hata": "URL boş olamaz."}), 400
    try:
        sonuc = analiz_yap(url)
        return jsonify({"sonuc": sonuc})
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print(tb)
        return jsonify({"hata": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
