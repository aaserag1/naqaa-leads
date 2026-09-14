# Naqaa | نقاء

أداة Streamlit عربية باتجاه RTL لتنظيف قوائم العملاء المحتملين من الملفات، بواجهة شبيهة بتطبيقات SaaS. هذه نسخة لمعالجة الملفات وتصدير النتائج، وليست خدمة CRM أو نظام اشتراكات متعدد المستأجرين.

> **قبل استخدام بيانات حقيقية:** الوضع الافتراضي `AUTH_MODE=open` مخصص للعرض ببيانات اصطناعية. فعّل OIDC مع قائمة السماح، أو قيّد المشاهدين فعلياً لدى المضيف. تصميم الواجهة لا يعني وجود عزل إنتاجي بين مؤسسات متعددة. راجع [DEPLOYMENT.md](DEPLOYMENT.md) قبل النشر.

## Product identity

الاسم النهائي المعتمد للمنصة هو **Naqaa | نقاء**، والمعرّف التقني المقترح للمشروع هو `naqaa`. يظهر الاسم في عنوان المتصفح ورأس الواجهة وشاشة الدخول. تخصيص **اسم الشركة وشعارها** مستقل تماماً: يظل ظاهراً في معاينة الهوية وتقارير Excel وPDF، دون فرض اسم المنصة أو شعارها على عنوان التقرير. راجع [BRANDING.md](BRANDING.md).

## Scope

- إدخال `xlsx` عبر `openpyxl`، و`xls` القديم عبر `xlrd`، و`csv` مع مساعدة `charset-normalizer` في التعرف على الترميز.
- مساعدة في توحيد أرقام الهواتف عبر `phonenumbers` ومقارنة النصوص عبر `rapidfuzz`. النتيجة تحتاج مراجعة بشرية؛ التشابه لا يثبت أن سجلين للشخص نفسه.
- تصدير النتائج وتقارير PDF عربية عبر **WeasyPrint** مع الخط المحلي `assets/Cairo.ttf` وترخيصه `assets/OFL-Cairo.txt`؛ لا حاجة إلى طلب خط من Google Fonts أثناء التشغيل.
- الملفات المرفوعة والبيانات المقروءة تبقى في ذاكرة جلسة Streamlit المعنية، لا في قاعدة بيانات أو ذاكرة تخزين مؤقت عامة مشتركة بين الجلسات. لا تحفظ بيانات العملاء في المستودع.
- لا توجد حسابات فوترة أو ربط CRM أو قاعدة بيانات أو عمليات سحب لقوائم العملاء عبر API. تسجيل OIDC، إذا فُعّل، يتصل بمزوّد الهوية للمصادقة فقط.
- روابط الاتصال وWhatsApp إجراءات يفتحها المستخدم يدوياً؛ لا توجد خدمة إرسال جماعي أو إرسال تلقائي للرسائل. فتح رابط خارجي يشارك المعلومات الموجودة فيه مع التطبيق الخارجي وفق سياساته.

## Runtime limits

| الحد | القيمة الافتراضية المستهدفة |
| --- | --- |
| حجم الملف الواحد | 20 MB؛ يضبطه `server.maxUploadSize` أيضاً |
| عدد الملفات | 5 في عملية الإدخال |
| عدد الصفوف | 25,000 ضمن عملية المعالجة |
| عدد الأعمدة | 200 |
| تقرير PDF | حتى 500 عميل في الدفعة، مع اختيار الدفعات/التقسيم إلى صفحات بشكل صريح |

حدّ PDF يخص **الدفعة، وليس الصفحة**؛ صدّر الدفعات المتبقية بدلاً من افتراض أن التقرير الأول يحتوي كل النتائج. هذه حدود وقائية وليست ضماناً بأن رفع أكبر خمسة ملفات معاً أو تشغيل عدة جلسات سيناسب ذاكرة استضافة مجانية. يتضخم استهلاك Excel وPDF بعد فك الضغط والتحليل. تُراجع الحدود في `app.py` والاختبارات وهذا الجدول معاً عند تغييرها.

## Quick start — Docker

ثبّت Docker على جهازك، ثم نفّذ الأوامر من المجلد الذي يحتوي `app.py` و`Dockerfile`:

```bash
docker build --pull -t naqaa:local .
docker run --rm -p 127.0.0.1:8501:8501 -e AUTH_MODE=open naqaa:local
```

افتح <http://localhost:8501>. هذا الربط متاح على جهازك فقط. توقف الحاوية ينهي جلساتها ولا يحفظ الملفات المرفوعة. صورة Docker تستخدم Python 3.12 وDebian Bookworm، وتثبّت مكتبات WeasyPrint، وتعمل بمستخدم غير root. يقرأ أمر التشغيل `PORT` عند وجوده وإلا يستخدم `8501`.

لا تُمرّر الأسرار كوسائط بناء Docker ولا تضفها إلى الصورة. تشغيل OIDC داخل الحاوية موضح في [دليل النشر](DEPLOYMENT.md#oidc-configuration).

## Local development — Linux

يلزم Python **3.12** مع دعم `venv` وبيئة Debian/Ubuntu تحتوي مكتبات النظام. Python 3.10 هدف توافق في CI فقط ما دامت الحزم المثبّتة تدعمه؛ المرجع الإنتاجي هو 3.12.

```bash
python3.12 -m venv .venv
. .venv/bin/activate
sudo apt-get update
sudo xargs -r -a packages.txt apt-get install -y --no-install-recommends
python -m pip install -r requirements.txt
python -m pip check
python -m weasyprint --info
python -m streamlit run app.py
```

على macOS أو Windows، Docker هو المسار الأكثر اتساقاً لهذه الحزمة. للتثبيت الأصلي اتبع [دليل WeasyPrint الرسمي](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html)، خصوصاً متطلبات Pango على Windows؛ تثبيت حزمة Python وحدها لا يضمن عمل PDF.

شغّل الأوامر دائماً من جذر المشروع لكي تُقرأ `.streamlit/config.toml` والمسارات المحلية. مراقبة الملفات معطلة في إعداد النشر؛ للتطوير فقط يمكن إضافة `--server.fileWatcherType=auto` إلى أمر Streamlit.

## Project files

```text
naqaa/
├── app.py
├── requirements.txt
├── requirements-dev.txt
├── packages.txt
├── Dockerfile
├── .dockerignore
├── .gitignore
├── render.yaml
├── README.md
├── DEPLOYMENT.md
├── START_HERE.md
├── BRANDING.md
├── VALIDATION.md
├── pytest.ini
├── assets/
│   ├── Cairo.ttf
│   └── OFL-Cairo.txt
├── tests/
│   └── ...
├── .streamlit/
│   └── config.toml
└── .github/
    └── workflows/
        └── ci.yml
```

لا يتضمن هذا المخطط `.streamlit/secrets.toml`: إن أنشأته محلياً فهو ملف خاص لا يُرفع. يحمي `.dockerignore` سياق البناء بقائمة سماح محدودة، ويستبعد `.gitignore` صيغ بيانات العملاء والتقارير الشائعة. **رفع الملفات بمتصفح GitHub لا يعتمد على `.gitignore`**؛ اختر الملفات يدوياً بعناية.

خدمة الملفات الثابتة مفعلة في الإعدادات، لكنها تخدم `static/` فقط، وليس `assets/` تلقائياً. يجهز Docker نسخة عامة من `Cairo.ttf` داخل `static/`. في التشغيل المباشر/Community Cloud يمكن للتطبيق تضمين بايتات الخط المحلي أو تقديم نسخة منه في `static/Cairo.ttf`. لا تضع في `static/` غير الخط العام وترخيصه؛ المحتوى هناك ليس محمياً ببوابة OIDC الخاصة بواجهة التطبيق. راجع [توثيق Streamlit](https://docs.streamlit.io/develop/concepts/configuration/serving-static-files).

## Authentication and privacy

- `AUTH_MODE=open`: عرض فقط ببيانات اصطناعية؛ لا يتطلب أسراراً.
- `AUTH_MODE=oidc`: يستخدم `st.login()` و`st.user`، ويجب أن تتوافر إعدادات `[auth]` وقائمة `ALLOWED_EMAILS` غير فارغة؛ غياب الإعدادات يجب أن يمنع الوصول، لا أن يعيد التطبيق إلى الوضع المفتوح.
- القبول يتطلب هوية مسجّلة، والبريد الكامل ضمن قائمة السماح المفصولة بفواصل، وادعاء `email_verified` بالقيمة المنطقية `true`. غياب هذا الادعاء أو عدم تحققه سبب للرفض حتى لو نجح تسجيل الدخول لدى المزوّد.
- يمكن توفير `AUTH_MODE` و`ALLOWED_EMAILS` بمتغيرات البيئة أو بمفاتيح جذرية في `st.secrets`. لا تكرر قيماً متعارضة بين المصدرين. إعدادات OIDC نفسها توضع في `[auth]` داخل Secrets.
- المصادقة تحدد من يدخل، لكنها لا تضيف قاعدة بيانات للمؤسسات أو صلاحيات متعددة المستأجرين أو اشتراكات. للإنتاج متعدد المؤسسات يلزم تصميم أمني وعزل وتدقيق إضافي.
- بيانات الجلسة ليست تخزيناً دائماً: قد تُفقد بعد إعادة التحميل أو انتهاء الجلسة أو إعادة التشغيل أو السكون. تنزيل الملفات يحفظ نسخة على جهاز المستخدم تقع مسؤولية حمايتها عليه.
- عبارة «في الذاكرة» تصف سلوك التطبيق، وليست وعداً بمحو جنائي فوري لذاكرة نظام التشغيل أو بعدم امتلاك مزوّد الاستضافة سجلات بنية تحتية. راجع سياسات المضيف ومزوّد الهوية، وتجنب تسجيل محتوى العملاء أو الأسرار في logs.

## Tests

من البيئة المحلية المفعلة، وببيانات اصطناعية فقط:

```bash
python -m pip install -r requirements.txt -r requirements-dev.txt
python -m pip check
python -m compileall -q app.py tests
python -m streamlit config show
python -m weasyprint --info
python -m pytest -q
```

ينفذ GitHub Actions اختبارات Python 3.12 و3.10، ويتحقق من إعدادات الحماية والأصول، ثم يبني Docker ويفحص تشغيله على `PORT=10000` بغير root وتوليد PDF في الذاكرة. لا يرفع صورة إلى سجل حاويات ولا ينشر تطبيقاً. لا تُعتبر الفحوص ناجحة إلا بعد تشغيلها فعلياً؛ لا تكفي إضافة ملف workflow.

## Deployment

اتبع [DEPLOYMENT.md](DEPLOYMENT.md) لرفع المشروع إلى **حساب GitHub الخاص بك** من المتصفح أو سطر الأوامر، مع الحفاظ على المجلدات المخفية، ثم اختر:

1. **Streamlit Community Cloud**: Python 3.12، `app.py`، `requirements.txt`، `packages.txt`، والأسرار في إعدادات المنصة.
2. **Render — Docker**: `Dockerfile`، متغير المنفذ `PORT`، وفحص `/_stcore/health`، وملف `secrets.toml` خاص وقت التشغيل.

المستودع المعتمد للمشروع هو [aaserag1/naqaa-leads](https://github.com/aaserag1/naqaa-leads). وجود الكود على GitHub لا يعني نشر التطبيق حيّاً أو إعداد أسرار تسجيل الدخول؛ يلزم ربط الاستضافة وضبط الوصول وفق دليل النشر.

## Official references

- [Streamlit — Community Cloud deployment](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)
- [Streamlit — Python and apt dependencies](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies)
- [Streamlit — st.login / OIDC](https://docs.streamlit.io/develop/api-reference/user/st.login)
- [WeasyPrint — Installation and security](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html)
- [Render — Docker deployment](https://render.com/docs/docker)
- [Render — Free-instance limits](https://render.com/docs/free)
