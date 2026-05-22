# LİDERFORM ANALİZ — WEB KURULUM REHBERİ

## Klasör Yapısı
```
liderform_web/
├── app.py              ← Flask uygulaması (ana dosya)
├── liderform_core.py   ← Analiz motoru (v9.8)
├── templates/
│   └── index.html      ← Web arayüzü
├── requirements.txt    ← Python paketleri
├── Procfile            ← Başlatma komutu
├── runtime.txt         ← Python versiyonu
├── railway.json        ← Railway ayarları
└── render.yaml         ← Render ayarları
```

---

## SEÇENEK 1: RAILWAY (Önerilen — Ücretsiz)

### 1. GitHub'a yükle
```bash
cd liderform_web
git init
git add .
git commit -m "liderform analiz v9.8"
# GitHub'da yeni repo oluştur, sonra:
git remote add origin https://github.com/KULLANICI/liderform-analiz.git
git push -u origin main
```

### 2. Railway'e deploy et
1. https://railway.app adresine git
2. "New Project" → "Deploy from GitHub Repo"
3. Reponuzu seçin
4. Otomatik deploy başlar (~2 dakika)
5. "Settings" → "Domains" → ücretsiz domain alın
   (örn: `liderform-analiz.up.railway.app`)

**Ücretsiz plan:** Ayda 500 saat (tek proje için yeterli)

---

## SEÇENEK 2: RENDER (Alternatif — Ücretsiz)

1. https://render.com adresine git
2. "New" → "Web Service"
3. GitHub reponuzu bağlayın
4. Ayarlar otomatik okunur (render.yaml sayesinde)
5. "Create Web Service" tıklayın
6. Ücretsiz URL alırsınız: `liderform-analiz.onrender.com`

**Not:** Render ücretsiz planda 15 dakika aktivite yoksa uyur,
ilk istek ~30 saniye gecikebilir.

---

## YEREL TEST (bilgisayarınızda)

```bash
cd liderform_web
pip install -r requirements.txt
python app.py
```
Tarayıcıda: http://localhost:5000

---

## SIKÇA SORULAN SORULAR

**S: Analiz ne kadar sürer?**
A: Her koşu için 30-60 saniye. Liderform'dan veri çekildiği için
internet hızına bağlıdır.

**S: Çok kişi aynı anda kullanabilir mi?**
A: Evet, gunicorn 2 worker ile çalışır. Çok yoğun kullanım için
worker sayısını artırabilirsiniz.

**S: Veriler kaydediliyor mu?**
A: Hayır. Her analiz bellekte işlenir, disk'e hiçbir şey yazılmaz.

**S: Timeout hatası alırsam?**
A: Render/Railway'de timeout 120 saniye. Liderform yavaş dönerse
hata alabilirsiniz. Bir daha deneyin.
