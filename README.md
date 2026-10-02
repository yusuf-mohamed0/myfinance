# MyFinance

نظام إدارة مالية شخصي، محلي بالكامل: بيحلل رسائل الـSMS البنكية، بيطلع خطة شهرية
بالجنيه، وبيبني موقع dashboard كامل تقدر تنشره على GitHub Pages.

A fully local personal-finance system: parses your bank SMS, builds an exact
EGP monthly plan, and generates a self-contained dashboard you can publish to
GitHub Pages.

---

## المميزات / Features

| الميزة | الوصف |
|---|---|
| تحليل SMS | بيفصل وارد/صادر، بيصنّف كل عملية في 12 تصنيف، وبيحسب المصروف الحقيقي |
| خطة الجنيه | بيوزّع الدخل على أظرفEnvelope بسقف أسبوعي، كل رقم بالجنيه مش نسبة |
| تقارير Excel | كراسة تحليل كامل + خطة الشهر + كراسة_master، كلها بأوراق عربية |
| تقرير Word | دليل الإدارة المالية بصيغة docx |
| موقع Dashboard | واجهة iOS/Shopify hybrid، موبايل أول، RTL كامل، صفر إيموجي |
| قفل البوابة | SHA-256 gate passcode، والتوكنات مش قابلة للتخمين |
| مساعد مالي | شخصية SVG دايمة + ردود متغيّرة حسب أرقامك اللحظية |
| AI اختياري | 20 مزوّد (OpenAI/Gemini/Groq/MiMo...) بمفتاحك أنت، مش متخزّن في السيرفر |
| مزامنة | سيرفر محلي يخلّي موبايلك وجهازك نسخة واحدة على نفس الواي فاي |
| أصوات | Web Audio إجرائي (صفر ملفات)،Dependency صفر |
| Market pulse | أسعار البنزين والدولار والتضخم، بتتحدّث أسبوعياً كل سبت |

---

## التشغيل السريع / Quick start

### المتطلبات / Requirements

- Python 3.9 أو أحدث
- الحزم: `openpyxl` (Excel) و `python-docx` (Word) — اختيارية حسب اللي عايز تعمله

```bash
git clone https://github.com/yusuf-mohamed0/myfinance.git
cd myfinance
python init_project.py          # ينشئ البنية + config.json
pip install openpyxl python-docx  # اختياري
```

### 1. حط رسائل البنك / Put your bank SMS

الصق رسائل الـSMS في:

```
01_Data/sms_raw_2026-01-28_to_2026-09-19.txt
```

كل سطر رسالة، بأي صيغة البنك بيكتبها.

### 2. اضبط الإعدادات / Configure

افتح `03_System/config.json`:

```json
{
  "name": "MyFinance",
  "passcode": "your-own-passcode",
  "income": 17480,
  "currency": "EGP",
  "city": "Cairo",
  "github_owner": "",
  "github_repo": "",
  "weekly_cap_fallback": 0
}
```

- `passcode`: قفل الموقع + بذرة روابط التحميل. **غيّرها قبل أي حاجة.**
- `income`: دخلك الشهري بالجنيه.
- `github_owner/repo`: لو هتنشر على Pages، اكتبهم هنا.

### 3. ابنِ واقرأ / Build and verify

```bash
python 04_Source/analyze_sms.py      # يحلل SMS ويكتب CSV/Excel/JSON/dashboard
python 04_Source/plan_month.py       # يبني خطة الشهر
python 04_Source/build_web.py        # يبني الموقع في 06_Web
python 04_Source/verify_system.py    # 56 فحص — لازم يعدي كلهم
```

### 4. انشر / Publish (اختياري)

```bash
python 04_Source/publish_web.py
```

بيعمل commit + push وينتظر build بتاع GitHub Pages.

---

## البنية / Structure

```
01_Data/          رسائل الـSMS الخام + daily_log.csv
02_Reports/       المخرجات: CSV/Excel/JSON/dashboard/plan + market_watch.json
03_System/        config.json + كراسة الـmaster + source_manifest.json (موقّع)
04_Source/        كل السكربتات + site_template.html + art/
05_Docs/          Word + README (غير متتبَّع)
06_Web/           الموقع المولَّد (غير متتبَّع)
```

`site_template.html` هو قالب HTML بـ26 placeholder. `build_web.py` يملاه ببياناتك
ويحطّ ناتج `safe_substitute`. أي تعديل على شكل الصفحة = تعديل في القالب، مش في
السكربت.

---

## الأمان / Security

- **البوابة**: SHA-256 من `passcode` في `config.json`. اللي معاه الباسكود بس
  يقدر يفتح الموقع.
- **روابط التحميل**: كل اسم ملف بيت prepend له `sha256(passcode + ":" + name)[:10]`.
  حد تاني يقدر يخترع الروابط دي؟ لأ.
- **الملفات**: `01_Data`, `03_System`, `05_Docs`, `06_Web` مستبعدة من git
  (ماعدا `.gitkeep` وملفSigned واحد). بياناتك ما بتترفعش.
- **الـmanifest**: `03_System/source_manifest.json` موقّع RSA-2048/SHA-256.
  لو عدّلت أي سكربت، شغّل `sign_template_manifest.py` تاني.
- **الـACL**: `verify_system.py` بيتأكد إن مجلد المشروع ما يفتحش غير لحسابك
  بتاع ويندوز.

---

## الاختبارات / Tests

```bash
python 04_Source/verify_system.py
# RESULT: 56 passed, 0 failed
# ALL GREEN - system fully wired
```

الـ56 فحص: البنية، الملفات، البيانات، الاتساق، NTFS ACL، الموقع، الرموز.

---

## المساهمات / Contributing

MIT License. Issues و PRs مرحب بيهم. لو عدّلت سكربت، شغّل `verify_system.py`
+a`sign_template_manifest.py` قبل PR.

---

<div dir="rtl">

## أسئلة شائعة

**هل أقدر أشغّلها من غير نت؟** أيوا، كل حاجة محلية. مفيش CDN ولا API إجباري.

**ليش الأرقام مش ظاهرة كاملة؟** ده ضمptive: فيه زرار العين بيعمل blur للأرقام
الشخصية بس (الأسعار العامة في السوق تفضل واضحة).

**ليه تغيير الرسايل مشتغل؟** قفل بابي؛ الرسايل اللي في `01_Data` هي المصدر.
لو ضفت رسالة جديدة، شغّل `analyze_sms.py` تاني.

**الموبايل يشتغل مع الجهاز؟** أيوا: شغّل `python 04_Source/sync_server.py`،
افتح `http://<IP>:8765` على الموبايل من نفس الواي فاي.

</div>