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
    build_excel(web_mode=True) ile tam analizi yapar,
    Excel yazmak yerine dict döndürür.
    """
    import re

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

    # 4. Tam analiz — Excel yazmadan, web_mode=True
    raw = core.build_excel(
        meta, rinfo, horses, acc_data, [],
        fname="",
        son_kosu_map=son_kosu_map,
        web_mode=True
    )

    eleme  = raw["eleme"]
    tempo  = raw["tempo"]

    kosu_mesafe_str = rinfo.get("mesafe","1200 m").replace(" m","").strip()
    try:    kosu_mesafe = int(kosu_mesafe_str)
    except: kosu_mesafe = 1200

    is_arap    = "ARAP"    in rinfo.get("tip","").upper()
    is_ingiliz = not is_arap
    disi_sayisi  = sum(1 for h in horses if h.get("cinsiyet") == "dişi")
    erkek_sayisi = sum(1 for h in horses if h.get("cinsiyet") == "erkek")
    karma_cinsiyet = is_ingiliz and disi_sayisi > 0 and erkek_sayisi > 0

    return {
        "meta":            meta,
        "rinfo":           rinfo,
        "kosu_mesafe":     kosu_mesafe,
        "sonuclar":        eleme,
        "tempo_karar":     tempo["karar"],
        "tempo_aciklama":  tempo["aciklama"],
        "tahmini_tempo":   tempo["tahmini"],
        "tempo_mes":       tempo["tempo_mes"],
        "tempo_at_detay":  tempo["at_detay"],
        "sanslilar":       tempo["sanslilar"],
        "dezavantajlilar": tempo["dezavantajlilar"],
        "oyun_kurucu_adlar": tempo["oyun_kurucular"],
        "sprinter_adlar":    tempo["sprinterlar"],
        "on_grup_adlar":     tempo["on_gruplar"],
        "geride_adlar":      tempo["geride_adlar"],
        "is_arap":         is_arap,
        "is_ingiliz":      is_ingiliz,
        "karma_cinsiyet":  karma_cinsiyet,
        "disi_sayisi":     disi_sayisi,
        "erkek_sayisi":    erkek_sayisi,
        "at_sayisi":       len(horses),
        "elenen_sayisi":   sum(1 for s in eleme if s["elendi"]),
        "kalan_sayisi":    sum(1 for s in eleme if not s["elendi"]),
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
