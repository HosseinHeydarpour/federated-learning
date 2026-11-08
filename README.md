# Federated Learning — حریم خصوصی در IoT

ارائهٔ فارسی «یادگیری فدرال برای حفظ حریم خصوصی در اینترنت اشیا» به همراه راهنمای ارائه‌دهنده.

- **سایت راهنما:** https://hosseinheydarpour.github.io/federated-learning/
- **پاورپوینت:** [`presenter-guide/downloads/FL_IoT_Privacy_FA.pptx`](presenter-guide/downloads/FL_IoT_Privacy_FA.pptx)
- **فونت‌ها:** [`presenter-guide/downloads/fonts.zip`](presenter-guide/downloads/fonts.zip)
- **PDF چاپی راهنما:** [`presenter-guide/presenter-guide.pdf`](presenter-guide/presenter-guide.pdf)

## ساختار

```text
build_deck.py                  # سازندهٔ پاورپوینت
assets/                        # تصاویر و پس‌زمینه‌های ارائه
fonts/                         # B Nazanin و B Titr
presenter-guide/               # سایت راهنما (همین پوشه روی Pages منتشر می‌شود)
├─ index.html                  # خروجی make_guide.py
├─ styles.css, app.js
├─ slides/slide-01..20.png     # تصاویر اسلایدها
├─ downloads/                  # فایل‌های دانلودی سایت
└─ make_guide.py               # سازندهٔ سایت و PDF
.github/workflows/pages.yml    # انتشار خودکار روی GitHub Pages
```

## بازسازی

```powershell
python build_deck.py          # ساخت پاورپوینت
python presenter-guide\make_guide.py   # ساخت index.html
```

پس از هر تغییر در ارائه: تصاویر `preview/Slide*.PNG` را در `presenter-guide/slides/` کپی کنید، `make_guide.py` را اجرا کنید، فایل PPTX را در `presenter-guide/downloads/` به‌روز کنید و PDF را دوباره بسازید.

## انتشار

هر push روی `main` به‌صورت خودکار با GitHub Actions روی GitHub Pages منتشر می‌شود.
اگر Pages تنظیم نشده باشد: Settings → Pages → Source = **GitHub Actions**.
