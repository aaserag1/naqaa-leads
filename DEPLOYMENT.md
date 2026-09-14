# Naqaa | نقاء — Deployment Guide

المستودع المعتمد هو [aaserag1/naqaa-leads](https://github.com/aaserag1/naqaa-leads). هذا دليل تشغيل على **حساباتك أنت**؛ وجود المستودع لا يعني نشر التطبيق حيّاً. أقسام إنشاء مستودع جديد اختيارية لمن يريد نسخة مستقلة؛ استخدم المستودع المعتمد للنشر الحالي. لا تُرسل كلمة مرور أو رمز دخول أو أسرار OIDC في المحادثة أو داخل رابط Git.

## 1. Deployment decision

| المسار | متى تختاره | ما ينبغي معرفته |
| --- | --- | --- |
| Streamlit Community Cloud | تجربة أو أداة داخلية محدودة مع ضوابط وصول مناسبة | يثبت `requirements.txt` و`packages.txt` من GitHub. اختر Python 3.12 صراحةً. ليس ضمان استمرارية أو سعة إنتاجية. |
| Render — Docker | تشغيل يمكن التحكم في مكتبات نظامه بصورة أوضح | يثبت Docker مكتبات Pango والخطوط ويعمل بغير root. `render.yaml` يختار Free للتجربة فقط؛ اختر خطة إنتاج ملائمة بعد مراجعة التكلفة والذاكرة. |
| Docker محلي | اختبار بلا نشر عام، أو مراجعة قبل اتخاذ قرار الاستضافة | يظل التطبيق محلياً عند ربط المنفذ بـ`127.0.0.1`. لا ترفع بيانات حقيقية إلى مستودع المصدر. |

النسخة الحالية لا تحتوي فوترة أو اشتراكات أو قاعدة بيانات أو حسابات CRM أو سحب بيانات عملاء عبر API. المخرجات تنزيلات ملفية وروابط اتصال/WhatsApp يفتحها المستخدم يدوياً؛ لا يوجد إرسال جماعي. OIDC اتصال خاص بالهوية وليس تصريحاً للتعامل مع CRM نيابة عن المستخدم.

**حماية الوصول شرط للإنتاج:** فعّل `AUTH_MODE=oidc` مع بريد موثّق وقائمة سماح، أو قيّد جميع المشاهدين فعلياً لدى المضيف/بوابة دخول موثوقة. `AUTH_MODE=open` للعرض ببيانات اصطناعية؛ لا يحول حتى مع تقييد المضيف التطبيق إلى منتج متعدد المستأجرين. خصوصية مستودع GitHub وحدها ليست بديلاً عن التحقق من إعداد وصول رابط التطبيق، وخصوصاً في Render.

## 2. Preflight

1. استخرج ملفات المشروع إلى مجلد على جهازك. يجب أن يكون `app.py` و`requirements.txt` و`packages.txt` و`Dockerfile` في جذر المستودع، وليس داخل مجلد إضافي غير مقصود.
2. تحقق من وجود `requirements-dev.txt` و`tests/` وأصلَي الخط `assets/Cairo.ttf.b64` و`assets/OFL-Cairo.txt`، والملفات المخفية في [مخطط README](README.md#project-files).
3. احذف من نسخة الرفع أي بيانات عملاء أو تقارير أو ملفات `.env` أو `secrets.toml` أو قواعد بيانات. اختبر ببيانات اصطناعية فقط. `.gitignore` يقلل الأخطاء لكنه ليس فحصاً أمنياً ولا يحذف ملفات سبق تتبعها.
4. نفّذ اختبار Docker المحلي من README، ثم اختبارات `pytest`. نجاح تشغيل صفحة Streamlit وحده لا يثبت صحة PDF أو قواعد التنظيف.
5. راجع الحدود: 20 MB/ملف، 5 ملفات، 25,000 صف/عملية، 200 عمود، و500 عميل/دفعة PDF. يجب أن تطابق الثوابت الفعلية والاختبارات. إذا خفضتها للذاكرة المتاحة، حدّث الوثائق أيضاً.
6. قرر نطاق المستخدمين وسياسة التعامل مع بياناتهم ومكان الاستضافة. لا تستخدم Free للتعهد بزمن استجابة أو توفر مضمون.

## 3. GitHub — browser upload

هذا المسار لا يحتاج Git مثبتاً ولا وصول المساعد إلى GitHub. تعليمات GitHub الرسمية تدعم رفع ملفات/مجلدات حتى **100 ملف في العملية و25 MiB للملف**؛ لا ترفع أرشيف المشروع فقط وتتوقع من الاستضافة فك ضغطه. [GitHub upload documentation](https://docs.github.com/en/repositories/working-with-files/managing-files/adding-a-file-to-a-repository)

1. افتح <https://github.com> وسجّل الدخول بنفسك. اختر **New repository**، وحدد المالك واسم مستودع جديد، مثل `naqaa`، واختر الخصوصية المناسبة. لا تهيئ README أو `.gitignore` مختلفين إذا كنت سترفع الملفين الموجودين هنا.
2. افتح المستودع. في المستودع الفارغ استخدم **uploading an existing file**؛ وبعد إنشاء ملف استخدم **Add file → Upload files**.
3. ارفع **محتويات** مجلد المشروع لا الغلاف الخارجي `naqaa/`. ارفع ملفات الجذر المرئية، ومجلدَي `assets/` و`tests/` مع الحفاظ على مساراتهما. راجع كل اسم في قائمة الرفع قبل التأكيد.
4. فعّل إظهار الملفات المخفية في مدير ملفات جهازك عند الحاجة. لا تعتمد على سحب المجلد وحده لإدراج ملفات النقطة؛ قد يخفيها نظام التشغيل أو منتقي الملفات. **لا تختَر `.git/` أو `.venv/` أو ملفات الأسرار**.
5. لضمان وجود الملفات المخفية، استخدم من جذر المستودع **Add file → Create new file** لكل مسار مما يلي، والصق محتوى الملف الأصلي كاملاً دون تغيير. كتابة `/` في الاسم تنشئ المجلدات اللازمة؛ لا تضف لاحقة `.txt`:

   ```text
   .gitignore
   .dockerignore
   .streamlit/config.toml
   .github/workflows/ci.yml
   ```

   إذا رُفعت هذه الملفات بالفعل فلا تنشئ نسخاً إضافية؛ افتحها وتحقق من مساراتها. يدعم محرر GitHub إنشاء المجلدات بهذه الطريقة رسمياً. [GitHub create-file documentation](https://docs.github.com/en/repositories/working-with-files/managing-files/creating-new-files)

6. **لا تنشئ `.streamlit/secrets.toml` في GitHub**، حتى إذا كان المستودع خاصاً. تُدخل الأسرار في المنصة فقط. رفع المتصفح قد يقبل ملفات يستبعدها `.gitignore`؛ المسؤولية هنا مراجعة الملفات المختارة.
7. اكتب رسالة commit واضحة ثم احفظ التغييرات. في مستودع قائم ومحمي، ارفع إلى فرع وافتح Pull Request ثم ادمجه بعد المراجعة. في المستودع الجديد راجع اسم الفرع النهائي؛ التعليمات أدناه تفترض `main`.
8. افتح كل مسار من هذه المسارات في GitHub قبل الانتقال إلى الاستضافة:

   ```text
   app.py
   requirements.txt
   requirements-dev.txt
   packages.txt
   Dockerfile
   render.yaml
   assets/Cairo.ttf.b64
   assets/OFL-Cairo.txt
   tests/
   .streamlit/config.toml
   .github/workflows/ci.yml
   ```

9. افتح تبويب **Actions** وراقب workflow المسمى **Naqaa CI**. لا تضف أسرار النشر إليه؛ الاختبارات تستخدم بيانات اصطناعية والوضع المفتوح المحلي. قد تحتاج حسابات المؤسسات إلى السماح بتشغيل Actions وفق سياسة المؤسسة.

## 4. GitHub — CLI on your own machine

نفّذ هذه الأوامر **على جهازك أنت فقط** بعد إنشاء مستودع GitHub فارغ، ومن نسخة مشروع جديدة غير مرتبطة بمستودع آخر. استبدل اسم الحساب والمستودع في الرابط بالقيم الحقيقية. إذا كان المستودع يحتوي تاريخاً أو ملفات مسبقاً، استخدم `git clone` لذلك المستودع ثم انقل ملفات المشروع إليه وافتح فرعاً/طلب دمج؛ لا تستخدم force push.

```bash
cd "/path/to/naqaa"
git init -b main
git add app.py requirements.txt requirements-dev.txt packages.txt pytest.ini Dockerfile render.yaml README.md DEPLOYMENT.md START_HERE.md BRANDING.md VALIDATION.md .gitignore .dockerignore .streamlit/config.toml .github/workflows/ci.yml assets tests
git status --short
git diff --cached --stat
```

توقف هنا وراجع القائمة والمحتوى قبل commit؛ يجب ألا يظهر أي ملف عملاء أو تقرير أو سر. لا تستخدم `git add -f` لتجاوز ملفات البيانات المستبعدة. بعد المراجعة فقط:

```bash
git commit -m "Add Naqaa application and deployment files"
git remote add origin https://github.com/aaserag1/naqaa-leads.git
git push -u origin main
```

استخدم آلية المصادقة الآمنة المعتادة في Git على جهازك، مثل مدير بيانات الاعتماد أو SSH المعدّ مسبقاً. لا تضع PAT أو كلمة مرور في عنوان `https://...` أو في ملف المصدر. إذا فشل الدخول، أصلح اتصال جهازك بدلاً من مشاركة بيانات الدخول مع المساعد.

إذا سبق رفع سر بالخطأ، ألغِه/دوّره فوراً لدى المزوّد وعالج تاريخ المستودع وفق دليل GitHub؛ حذف الملف من آخر commit لا يجعله سرياً من جديد.

## OIDC configuration

هذه الإعدادات مشتركة بين طريقتي الاستضافة. تستخدم `st.login()` مزوّد **OpenID Connect**، وليس أي مزوّد OAuth عام. يتطلب Streamlit تثبيت `Authlib>=1.3.2`؛ يجب أن يتضمن `requirements.txt` إصداراً متوافقاً ومثبتاً منه أو تثبيت `streamlit[auth]` بإصدار مثبت مع ضبط اعتماداته. راجع [st.login](https://docs.streamlit.io/develop/api-reference/user/st.login) و[Authentication concepts](https://docs.streamlit.io/develop/concepts/connections/authentication).

### Identity provider

1. أنشئ تطبيق OAuth/OIDC من نوع Web application لدى مزوّد الهوية. مثال واضح لهذه السياسة هو Google، الذي يوفر ادعاء `email_verified`. جهّز شاشة الموافقة والجمهور المسموح ومستخدمي الاختبار إن كان التطبيق في وضع Testing. [Google OIDC documentation](https://developers.google.com/identity/openid-connect/openid-connect)
2. سجل **عنوان إعادة التوجيه المطابق تماماً** في إعدادات المزوّد وفي Secrets. اختر ما يلائم الوجهة:

   ```text
   http://localhost:8501/oauth2callback
   https://YOUR-APP.streamlit.app/oauth2callback
   https://YOUR-SERVICE.onrender.com/oauth2callback
   ```

   استخدم HTTPS للوجهة العامة. إذا تغير النطاق أو أضفت نطاقاً مخصصاً، عدّل الإعدادين معاً. يمكن أن يوجد أكثر من عنوان مصرح لدى المزوّد، لكن قيمة `redirect_uri` الفعالة في البيئة الواحدة يجب أن تكون عنوان تلك البيئة.
3. احتفظ بـ`client_id` و`client_secret` خارج Git. أنشئ cookie secret عشوائياً على جهازك؛ لا تستخدم نص المثال أدناه ولا تنشر ناتج الأمر:

   ```bash
   python -c "import secrets; print(secrets.token_urlsafe(48))"
   ```

4. اختر عناوين البريد الكاملة للمستخدمين المصرح لهم. `ALLOWED_EMAILS` سلسلة مفصولة بفواصل، لا قائمة نطاقات عامة ولا wildcard. مثال الصيغة: `owner@example.com,reviewer@example.com`، مع استبدالها بعناوينك الفعلية.

### Secrets template — never commit

المثال التالي **قالب فقط بلا أسرار حقيقية**. استبدل كل `REPLACE_...` و`YOUR-APP` والبريد النموذجي قبل الاستخدام. محلياً احفظه في `.streamlit/secrets.toml` المستبعد من Git، وفي Community Cloud الصقه في واجهة Secrets. لا تحط النص بأسوار Markdown داخل حقل Secrets.

```toml
AUTH_MODE = "oidc"
ALLOWED_EMAILS = "owner@example.com,reviewer@example.com"

[auth]
redirect_uri = "https://YOUR-APP.streamlit.app/oauth2callback"
cookie_secret = "REPLACE_WITH_LOCALLY_GENERATED_RANDOM_SECRET"
client_id = "REPLACE_WITH_OIDC_CLIENT_ID"
client_secret = "REPLACE_WITH_OIDC_CLIENT_SECRET"
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
```

المفاتيح العليا `AUTH_MODE` و`ALLOWED_EMAILS` توضع **قبل** `[auth]`؛ وضعها بعده دون جدول جديد يجعلها مفاتيح داخل `auth` بالخطأ. القالب يستخدم مزوّداً افتراضياً واحداً في `[auth]` لكي يناسب `st.login()` دون اسم مزوّد. لا تضف `expose_tokens` ولا صلاحيات APIs إضافية؛ هذا التطبيق لا يحتاج عرض access tokens أو الوصول إلى بيانات حسابات أخرى.

يمكنك اختيار مزوّد OIDC آخر، لكن يجب أن يوفر هوية تحقق شروط التطبيق. **إذا غاب `email_verified` أو لم يكن boolean true، يجب رفض الوصول**؛ لا تتجاوز هذا الشرط باستخدام قيمة البريد وحدها. ظهور شاشة اختيار الحساب ليس اختبار تفويض ناجحاً: اختبر مستخدماً داخل القائمة وآخر خارجها، وقائمة فارغة، وإعدادات ناقصة.

### Fail-closed policy

- في `AUTH_MODE=oidc` يجب أن يمنع التطبيق الوصول إلى الرفع والبيانات والتصدير عند غياب `[auth]` أو نقص إعداداته أو فراغ قائمة السماح أو عدم التحقق من البريد.
- لا تحوّل `AUTH_MODE` إلى `open` لمعالجة خطأ إنتاجي في تسجيل الدخول. أصلح الأسرار/عنوان العودة، مع بقاء البوابة مغلقة.
- إذا استخدمت متغيرات بيئة لـ`AUTH_MODE` و`ALLOWED_EMAILS`، لا تضع قيماً متعارضة لها أيضاً في Secrets. في Render يُفضّل وضع هذين المتغيرين في Environment، وملف الأسرار يحوي `[auth]` فقط.
- `.streamlit/config.toml` يبحث عن الأسرار بالترتيب: `~/.streamlit/secrets.toml` ثم `.streamlit/secrets.toml` ثم `/etc/secrets/secrets.toml`. الملف اللاحق يتقدم على السابق. لا يضع هذا الإعداد أي أسرار في المستودع. [Streamlit configuration](https://docs.streamlit.io/develop/api-reference/configuration/config.toml)
- تسجيل الدخول والخروج قد ينشئ جلسة Streamlit جديدة. نزّل النتائج المسموح بها قبل الخروج. لا تعتمد على إبقاء جلسة المتصفح مفتوحة كسياسة صلاحيات أو مدة جلسة إنتاجية.

### Local Docker with OIDC

بعد إعداد ملف الأسرار المحلي بعنوان العودة المحلي، شغّل الحاوية بتركيب **للقراءة فقط**. يتطلب الأمر التالي shell من نوع Bash؛ على PowerShell استخدم مسار الملف المطلق الموافق لنظامك.

```bash
docker build --pull -t naqaa:local .
docker run --rm -p 127.0.0.1:8501:8501 -e AUTH_MODE=oidc --mount "type=bind,source=$(pwd)/.streamlit/secrets.toml,target=/etc/secrets/secrets.toml,readonly" naqaa:local
```

تأكد من أن الملف موجود وأن مستخدم الحاوية يستطيع قراءته دون جعله عاماً. لا تُدرج محتواه في Dockerfile أو `--build-arg`. عضوية المجموعة `1000` في الصورة مخصصة كذلك للوصول إلى Secret Files على Render وفق توثيقه؛ لا تشغّل الحاوية كـroot لحل خطأ صلاحيات.

## 5. Streamlit Community Cloud

التوثيق الرسمي يذكر Python 3.12 كالإصدار الافتراضي عند المراجعة، لكن **اختره صراحةً** لتطابق بيئة Docker. قد تتغير الإصدارات المدعومة بمرور الوقت. [Deployment reference](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)

1. افتح <https://share.streamlit.io> وسجّل الدخول واربط حساب GitHub الذي يملك المستودع، وامنح التطبيق الوصول لذلك المستودع فقط قدر الإمكان. قد تحتاج المؤسسة إلى اعتماد الوصول.
2. اختر **Create app** ثم **Yup, I have an app**.
3. املأ **Repository** باسم `aaserag1/naqaa-leads`، و**Branch** بـ`main` أو فرعك الحقيقي، و**Main file path** بـ`app.py`. اختر subdomain مناسباً إن رغبت، واحفظ عنوانه لإعداد OIDC.
4. افتح **Advanced settings** واختر **Python version → 3.12**.
5. للإنتاج/البيانات المقيدة، الصق قالب Secrets بعد استكماله، مع `AUTH_MODE="oidc"` و`ALLOWED_EMAILS` و`[auth]` و`redirect_uri` المطابق للنطاق المختار، ثم **Save**. للتجربة الاصطناعية فقط يمكنك ترك Secrets فارغاً والعمل بالوضع المفتوح.
6. اضغط **Deploy** وراقب سجلات البناء. لا تحتاج Build command أو Dockerfile في هذا المسار؛ Community Cloud يثبت Python من `requirements.txt` ويثبت apt من `packages.txt` في الجذر. لا تضف ملف إدارة اعتماد آخر يسبق `requirements.txt` دون قصد. [Dependencies](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies)
7. افتح التطبيق. أكمل اختبار OIDC والتحميل والتصدير ببيانات اصطناعية. إذا احتجت تحديث الأسرار بعد معرفة الرابط النهائي، افتح **App settings → Secrets** وعدّلها ثم انتظر إعادة تشغيل التطبيق. [Secrets management](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management)
8. للتحكم لدى المضيف أيضاً، افتح **App settings → Sharing → Who can view this app → Only specific people can view this app**، وأضف المشاهدين المقصودين ثم احفظ. راجع حدود الحساب وإمكانات الدعوات؛ التوثيق الحالي يذكر تطبيقاً خاصاً واحداً في الوقت نفسه. جرّب الرابط من نافذة خاصة غير مصرح لها ولا تعتمد فقط على كون مستودع GitHub خاصاً. [Sharing reference](https://docs.streamlit.io/deploy/streamlit-community-cloud/share-your-app)
9. لا تغيّر `enableCORS` أو `enableXsrfProtection` إلى `false` لحل مشكلات رفع/دخول. صحح النطاق وعنوان العودة وإعداد الوكيل أو أعد تشغيل الجلسة.

**الخطوط:** أبقِ `assets/Cairo.ttf.b64` وملف OFL في GitHub. هذا هو الخط الأصلي نفسه بترميز نصي قابل للعكس، وليس خطاً بديلاً. يفك التطبيق ترميزه في الذاكرة ويتحقق من بصمته قبل تضمينه في الواجهة وPDF. لا تحتاج Community Cloud إلى تجهيز يدوي للخط أو Docker. Docker يصنع النسخة العامة `static/Cairo.ttf` أثناء البناء. لا تضع بيانات جلسة أو أسراراً في static. [Static-file reference](https://docs.streamlit.io/develop/concepts/configuration/serving-static-files)

**السكون:** توثيق Community Cloud يذكر نوم التطبيقات بعد 12 ساعة بلا زيارات وإمكانية إيقاظها من صفحة التطبيق. لا تعد المستخدم بأن الاستضافة المجانية تعمل بلا توقف أو SLA. إعادة التشغيل/السكون ليست وسيلة لحفظ الجلسات. [App hibernation](https://docs.streamlit.io/deploy/streamlit-community-cloud/manage-your-app)

## 6. Render — manual Docker service

اختيار **Docker** مقصود لأن PDF يحتاج مكتبات نظام، وليس `pip install` فقط. توثيق Render يدعم البناء من Dockerfile ويتطلب الاستماع على `0.0.0.0` والمنفذ `PORT`، والذي يكون افتراضياً `10000`. [Docker](https://render.com/docs/docker) · [Port binding](https://render.com/docs/web-services#port-binding)

1. افتح <https://dashboard.render.com> وسجّل الدخول بنفسك.
2. اختر **New → Web Service → Git Provider**، ثم اربط حساب GitHub واختر مستودعك. لا يحتاج المساعد أو هذه الحزمة إلى إنشاء أي خدمة نيابة عنك.
3. عيّن اسم الخدمة والمنطقة المناسبة للبيانات و**Branch** الصحيح، واترك **Root Directory** فارغاً إذا كان المشروع في جذر المستودع.
4. اختر **Language → Docker** و**Dockerfile Path → ./Dockerfile** وسياق البناء جذر المشروع. اترك **Docker Command** فارغاً لاستخدام `CMD` الموجود؛ لا تستبدله بأمر Python native.
5. اختر **Free** للتجربة فقط إن كان متاحاً في حسابك. راجع السعر/الذاكرة الفعلية قبل التأكيد، أو اختر خطة مدفوعة مناسبة باختيار واعٍ. لا تضف قاعدة بيانات أو قرصاً دائماً؛ التطبيق لا يحتاجهما ولا يحفظ ملفات العملاء.
6. في **Advanced / Environment** أضف `AUTH_MODE=oidc` و`ALLOWED_EMAILS` بعناوينك المفصولة بفواصل. لا تضف مفتاحاً عاماً يفتح التطبيق تلقائياً عند فشل OIDC.
7. عيّن **Health Check Path** إلى `/_stcore/health`. لا تثبت `PORT=8501` في Render: أمر Docker يقرأ المنفذ الذي توفره المنصة. الرقم في `EXPOSE 8501` مجرد تعريف للصورة وليس إعداد منفذ Render.
8. إذا كانت واجهة الإنشاء تعرض **Secret Files**، أضف ملفاً اسمه `secrets.toml` بمحتوى `[auth]` من القالب بعد استكماله بعنوان `https://YOUR-SERVICE.onrender.com/oauth2callback`. وإلا أكمل إنشاء الخدمة بـOIDC المغلق، ثم أضف الملف من إعداداتها في الخطوة التالية.
9. بعد مراجعة الخطة والإعدادات اضغط **Create Web Service / Deploy**. هذه النقرة هي التي تنشئ المورد فعلياً؛ إنشاء الملفات المحلية لا يفعل ذلك.
10. افتح الخدمة ثم **Environment → Secret Files → Add Secret File** إذا لم تُضف الملف بعد. استخدم **Filename: `secrets.toml`**، والصق `[auth]` فقط؛ `AUTH_MODE` و`ALLOWED_EMAILS` موجودان في Environment. اضغط **Save Changes**. توفر Render الملف داخل الحاوية في `/etc/secrets/secrets.toml`، والإعداد الموجود يقرأه مباشرةً، دون نسخه إلى المستودع أو الصورة. [Secret files](https://render.com/docs/configure-environment-variables#secret-files) · [Docker secret permissions](https://render.com/docs/docker-secrets#accessing-secret-files-at-runtime)
11. استخدم رابط الخدمة الفعلي الظاهر في لوحة التحكم لتحديث `redirect_uri` لدى المزوّد وفي ملف الأسرار، ثم أعد النشر/التشغيل عند الحاجة. سيبقى الوصول مغلقاً إلى أن تصح الإعدادات؛ لا تغيّر الوضع إلى `open` أثناء ذلك.
12. اختبر من نافذة خاصة: الدخول المسموح، البريد خارج القائمة، الخروج، رفع ملفات اصطناعية وتوليد PDF. راقب الذاكرة والسجلات وأخطاء الاستدعاء العكسي دون عرض بيانات العملاء فيها.
13. إن أردت النشر اليدوي فقط، عطّل **Auto-Deploy** في إعدادات الخدمة. وعند تحديث نسخة مختبرة اختر **Manual Deploy → Deploy latest commit**. لا تفترض أن نجاح GitHub CI يمنع النشر التلقائي ما لم تضبط سياسة النشر صراحةً.

### Blueprint alternative

لا تستخدم هذا المسار لإنشاء نسخة ثانية بالخطأ إن كنت قد أنشأت الخدمة يدوياً بالفعل. ملف `render.yaml` قالب يشغّل خدمة Docker واحدة، ولا يحتوي أسراراً أو مورد فوترة/CRM/قاعدة بيانات.

1. راجع `render.yaml`: `plan: free` للتجربة، و`branch: main`، و`AUTH_MODE: oidc`، و`ALLOWED_EMAILS` بقيمة `sync: false` كي تدخلها أنت في لوحة Render. لا تُدخل البريد الحقيقي أو أسرار المزوّد في YAML.
2. افتح **New → Blueprint**، ثم **Connect** للمستودع الصحيح. حدد اسم Blueprint وفرعك ومسار **Blueprint Path: `render.yaml`**.
3. راجع قائمة الموارد والتكلفة والمنطقة. عدم تحديد `region` في القالب يعني اختيار Render الافتراضي؛ اختر منطقة ملائمة قبل اعتماد نشر بيانات حقيقية.
4. أدخل قائمة البريد عندما يطلبها Render. بعد المراجعة فقط اضغط **Deploy Blueprint**. يبدأ ذلك الإنشاء والنشر الأول؛ `autoDeployTrigger: "off"` لا يمنع الإنشاء الأول.
5. أضف Secret File `secrets.toml` من لوحة الخدمة كما في المسار اليدوي. لا يحتوي Blueprint على الأسرار، ويجب أن تبقى بوابة OIDC مغلقة قبل توفيرها.
6. `autoDeployTrigger: "off"` يوقف النشر التلقائي المعتاد للخدمة عند دفع الكود. **Auto Sync الخاص بـBlueprint إعداد منفصل**: إذا أردت موافقة يدوية على تغييرات البنية أيضاً، افتح **Blueprint → Settings → Auto Sync → No** ثم استخدم **Manual Sync** عند الحاجة. تعديل `render.yaml` قد يعيد تهيئة الموارد إذا كان Auto Sync مفعلاً.

مرجع [Blueprint setup and Auto Sync](https://render.com/docs/infrastructure-as-code) و[Blueprint YAML schema](https://render.com/docs/blueprint-spec).

### Free hosting is not production capacity

توثيق Render الحالي يتيح **Free web services** ويقول صراحةً إنها ليست للإنتاج. تدخل الخدمة السكون بعد **15 دقيقة** بلا حركة واردة، بما فيها رسائل WebSocket؛ الاستيقاظ قد يستغرق نحو دقيقة. يمكن إعادة تشغيلها في أي وقت، وتخضع لساعات شهرية وباندويث ودقائق بناء وحدود حساب. قد تُعلّق الخدمة عند استنفاد الحصص، وقد تترتب رسوم استخدام إضافي وفق الخطة وطريقة الدفع. راجع لوحة الفوترة وحدود الإنفاق بنفسك. [Render Free documentation](https://render.com/docs/free)

لا تطبق وسائل ping لإيهام المستخدم بضمان «لا تنام». للبيانات الحقيقية اختر موارد وخطة ودعماً يناسب متطلباتك، واختبر الذروة ومدة PDF. حد رفع 20 MB لا يعني أن الذاكرة المستهلكة 20 MB؛ رفع خمسة ملفات وتحليلها وتوليد عدة تقارير بالتوازي قد يتجاوز ذاكرة خطة صغيرة بكثير.

## 7. Operations and troubleshooting

| العرض | الفحص والإجراء |
| --- | --- |
| `libpango` / `libpangoft2` / `libharfbuzz-subset` غير موجود | تأكد أن `packages.txt` في الجذر على Community Cloud، أو أن الخدمة على Render تستخدم Docker فعلاً. أعد البناء بعد تثبيت المكتبات؛ لا تغيّر PDF إلى محرك وهمي. |
| فشل بناء Pillow/CFFI من المصدر | تستخدم الحزمة `libjpeg-dev` و`libopenjp2-7-dev` و`libffi-dev` احتياطياً. إن لم تتوفر wheels لمنصتك فقد تحتاج compiler وحزم التطوير طبقاً لدليل WeasyPrint؛ لا تغيّر الإصدارات عشوائياً. |
| التطبيق لا يستمع على Render | لا تستخدم `localhost` داخل أمر الخادم؛ يجب أن يكون `0.0.0.0:$PORT`. اترك Docker Command الافتراضي وتحقق من سجلات اكتشاف المنفذ. |
| فحص الصحة ناجح لكن الدخول/التقرير لا يعمل | `/_stcore/health` يؤكد حياة خادم Streamlit فقط؛ لا يختبر أسرار OIDC ولا صلاحيات المستخدم ولا صحة PDF. نفّذ الاختبار التفاعلي أيضاً. |
| `redirect_uri_mismatch` | طابق البروتوكول والنطاق والمنفذ والمسار `/oauth2callback` تماماً في الجهتين، واستخدم عنوان الإنتاج لا localhost. |
| دخول المزوّد ناجح ثم الرفض | تحقق من `email_verified is True` ومن البريد الكامل ضمن `ALLOWED_EMAILS`. مزوّد لا يرجع الادعاء المطلوب غير مناسب دون تغيير مدروس للسياسة؛ لا تتجاوز البوابة. |
| `Permission denied` لملف Render السري | الاسم `secrets.toml` والمسار `/etc/secrets/secrets.toml`، ومستخدم الصورة ضمن المجموعة `1000`. لا تسجّل المحتوى ولا تحول العملية إلى root. |
| العربية مربعات أو خط غير مطابق | تحقق من `assets/Cairo.ttf.b64` وترخيصه ومن تحميل الخط المحلي في الواجهة وPDF. إذا استُخدم عنوان static، يجب وجود `static/Cairo.ttf`؛ `assets/` ليس مساراً عاماً. |
| نفاد الذاكرة أو بطء PDF | جرّب بيانات أصغر ودفعات PDF أصغر وراجع التزامن ثم ارفع الموارد عند الحاجة. 500 عميل/دفعة سقف لا ضمان أداء. راجع التقسيم الصريح للدفعات ولا تقطع السجلات بصمت. |
| فقدان البيانات بعد السكون/إعادة التشغيل | متوقع: بيانات الجلسة في الذاكرة فقط. يجب تنزيل النتائج قبل إنهاء الجلسة؛ لا توجد قاعدة بيانات لاستعادتها. |
| الرفع يفشل بـ403 أو XSRF | صحح عنوان التطبيق وإعداد الوكيل/HTTPS وCookies وأعد تسجيل الدخول؛ لا تعطل CORS أو XSRF كحل. |

المكتبات الأساسية في `packages.txt` هي `libpango-1.0-0` و`libpangoft2-1.0-0` و`libharfbuzz-subset0`، مع `fontconfig` و`fonts-dejavu-core` والخط المرفق. يذكر دليل WeasyPrint الرسمي هذه المكتبات لمسار Debian، ويضيف مكتبات JPEG/OpenJPEG/FFI عندما لا تُستخدم wheels. يتطلب الدليل الحالي Python 3.10 أو أحدث؛ لا يكفي ذلك وحده لإثبات توافق **جميع** إصدارات الحزم المثبتة مع 3.10. [WeasyPrint installation](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html#installation)

## 8. Production checklist

- [ ] اجتاز `requirements.txt` و`requirements-dev.txt` الاختبارات الفعلية و`pip check` على Python 3.12؛ إن توقف دعم 3.10 فحدّث مصفوفة CI والتوثيق ولا تتجاهل فشل الاعتمادات.
- [ ] اختُبر Docker وفحص الصحة، واختُبر PDF عربي فعلياً بصرياً وببيانات اصطناعية، لا بمجرد وجود الملف أو نجاح استيراد المكتبة.
- [ ] كل إصدارات Python المباشرة المطلوبة مثبتة. لا تعني `==` للحزم أن البناء bit-for-bit ثابت: اعتماداتها الانتقالية، وحزم apt، وصورة `python:3.12-slim-bookworm` قد تتغير. للإنتاج المحكوم اعتمد lock/constraints مع hashes وصورة أساس مثبتة بالـdigest، وراقب تحديثات الأمان؛ لا تُجمّد الثغرات إلى الأبد.
- [ ] فُعّل OIDC مع allowlist وبريد موثّق أو اختُبرت قيود المشاهدة لدى المضيف من نافذة غير مصرح لها؛ مستودع خاص أو رابط يصعب تخمينه ليسا حماية كافية.
- [ ] اختُبر الفشل المغلق عند غياب الأسرار/قائمة السماح وعدم تحقق البريد، ولا تظهر شاشة الرفع أو روابط تنزيل حساسة قبل السماح.
- [ ] عولجت دورة حياة الجلسة وإلغاء الوصول والخروج وفق سياسة المؤسسة؛ OIDC ليس نظام صلاحيات متعدد المستأجرين.
- [ ] لا توجد ملفات عملاء أو تقارير أو أسرار في Git أو الصورة أو `static/`؛ لا توجد ملفات مرفوعة/بيانات مقروءة في global cache. الاختبارات تستخدم توليد بيانات اصطناعية في الذاكرة.
- [ ] بقي `enableCORS=true` و`enableXsrfProtection=true` و`gatherUsageStats=false`، ولم تُسجل بيانات العملاء أو رموز الهوية في logs.
- [ ] تم اختبار حدود الصفوف/الأعمدة وعدد الملفات والحجم والذاكرة المتزامنة ووقت PDF ودفعاته. قيود الملفات لا تغني عن حدود CPU/ذاكرة/مهلة على المضيف ومراقبة تشغيلية.
- [ ] قوالب PDF ثابتة ومدخلات العملاء مهربة، ولا يُقبل HTML/CSS تعسفي أو موارد خارجية قادمة من بيانات المستخدم. راجع خطر قراءة ملفات محلية/شبكية وإساءة استهلاك الموارد في [WeasyPrint security](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html#security).
- [ ] الخط يأتي من `assets/Cairo.ttf.b64` محلياً في العرض والتصدير ولا يطلب خدمة خطوط خارجية وقت التشغيل؛ بقي ملف OFL مصاحباً له.
- [ ] تم توضيح موافقات الاتصال وسياسة الخصوصية والتنزيلات ومسؤولية الأجهزة للمستخدمين. روابط الهاتف/WhatsApp يفتحها المستخدم بنفسه ولا يوجد إرسال تلقائي أو جماعي.
- [ ] روجعت الخطة والفاتورة والمنطقة وسياسة السجلات وحدود المنصة؛ لا يُقدَّم وعد باستضافة مجانية بلا سكون أو SLA.
- [ ] تم اختبار جلسات مستقلة ومستخدمين مختلفين. قبل إطلاق SaaS لمؤسسات متعددة، أُنجز تصميم العزل والتفويض والتدقيق وإدارة الحسابات والفوترة المطلوبة فعلاً؛ هذه الوظائف ليست مضمّنة هنا.
- [ ] عُيّن مالك للتحديثات الأمنية وخطة rollback لنسخة مختبرة وإجراء لاستجابة الحوادث وتدوير الأسرار؛ تحديث السر يكون من المنصة لا من Git.

## Official source index

المراجع التالية هي المصادر الرسمية للتحقق من الخطوات؛ أسماء الأزرار والخطط والحدود قد تتغير، فراجع الصفحة المعنية ولوحة حسابك وقت النشر.

- [GitHub — Upload existing files](https://docs.github.com/en/repositories/working-with-files/managing-files/adding-a-file-to-a-repository)
- [GitHub — Create files and directories](https://docs.github.com/en/repositories/working-with-files/managing-files/creating-new-files)
- [Streamlit — File organization](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/file-organization)
- [Streamlit — App dependencies / packages.txt](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies)
- [Streamlit — Deploy, Python version, and Advanced settings](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)
- [Streamlit — Secrets management](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management)
- [Streamlit — Sharing and viewer access](https://docs.streamlit.io/deploy/streamlit-community-cloud/share-your-app)
- [Streamlit — Hibernation](https://docs.streamlit.io/deploy/streamlit-community-cloud/manage-your-app)
- [Streamlit — st.login](https://docs.streamlit.io/develop/api-reference/user/st.login)
- [Streamlit — config.toml](https://docs.streamlit.io/develop/api-reference/configuration/config.toml)
- [Streamlit — Static file serving](https://docs.streamlit.io/develop/concepts/configuration/serving-static-files)
- [Google — OpenID Connect](https://developers.google.com/identity/openid-connect/openid-connect)
- [WeasyPrint — First steps and security](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html)
- [Render — Docker](https://render.com/docs/docker)
- [Render — Web services and PORT](https://render.com/docs/web-services)
- [Render — Health checks](https://render.com/docs/health-checks)
- [Render — Environment variables and secret files](https://render.com/docs/configure-environment-variables)
- [Render — Docker secret access](https://render.com/docs/docker-secrets)
- [Render — Blueprint setup / Auto Sync](https://render.com/docs/infrastructure-as-code)
- [Render — Blueprint YAML reference](https://render.com/docs/blueprint-spec)
- [Render — Free plan limitations](https://render.com/docs/free)
