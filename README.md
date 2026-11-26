# Smart Sales CRM Pro

تطبيق سطح مكتب لإدارة العملاء ومراحل المبيعات (CRM Pipeline) مبني بـ Python و CustomTkinter و SQLite.

## المميزات
- لوحة تحكم تفاعلية بحساب نسبة التحويل ومتابعات اليوم المستحقة.
- إدارة مراحل المبيعات (جديد، تم التواصل، مهتم، معاينة، تم البيع، ملغي).
- قاعدة بيانات SQLite محلية سريعة مع استيراد وتصدير Excel بأزرار واتساب وإيميل جاهزة.
- مكتبة قوالب رسائل قابلة للتخصيص بمتغيرات تلقائية (`{name}`, `{category}`, `{budget}`, `{location}`, `{sender_name}`).
- إرسال حملات Gmail SMTP مع وضع التجربة (DRY RUN) وتأخير عشوائي.
- دعم الوضع الليلي والنهاري (Dark / Light Mode).

## التشغيل محلياً
```bash
pip install -r requirements.txt
python smart_crm_pro.py
```

## بناء ملف EXE لويندوز
```bash
pip install pyinstaller
pyinstaller --noconfirm --onedir --windowed --name "SmartCRM_Pro" --collect-all customtkinter smart_crm_pro.py
```

## البناء التلقائي عبر GitHub Actions
عند رفع الملفات إلى فرع `main`، يعمل سير العمل `.github/workflows/build-windows.yml` تلقائياً لبناء نسخة ويندوز مضغوطة يمكن تحميلها من تبويب **Actions > Artifacts**.
