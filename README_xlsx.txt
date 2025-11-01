\
تشغيل نظام المخزون القائم على Excel

هذا الإصدار يقرأ ويكتب مباشرة داخل:
D:\Project ILM\Tool store\المخزون.xlsx

التحضير:
1) أنشئ بيئة (اختياري):
   py -3 -m venv .venv
   .\.venv\Scripts\Activate.ps1
2) ثبّت المتطلبات:
   pip install -r requirements_xlsx.txt
3) التشغيل:
   streamlit run app_xlsx.py

ملاحظات:
- إذا لم يكن الملف موجودًا، سيُنشئ التطبيق ملفًا بأوراق: Stock, MinLevels, Transactions.
- صفحة "تحرير البيانات" تتيح تعديل Stock و MinLevels مباشرة ثم "حفظ التغييرات".
- جميع العمليات (استلام/صرف/تحويل/تسوية) تُسجّل في ورقة Transactions ويُحدّث Stock.
