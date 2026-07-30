# School Inventory - Excel Macro-Enabled Workbook (Assembly Instructions)

هذا المستودع يحتوي على ملفات داعمة لمصنف Excel جاهز لإدارة مخازن مدرسة أهلية.
الملفات تشمل: بيانات تجريبية بصيغة CSV، وحدات VBA (.bas)، وكود UserForm.

ملف الهدف: School_Inventory.xlsm (تحتاج لتجميعه محليًا عبر الخطوات التالية)

خطوات تجميع المصنف (.xlsm) محليًا:
1. افتح Microsoft Excel وابدأ مصنفًا جديدًا.
2. لكل ملف CSV في المجلد release/data/:
   - افتح ملف CSV، انسخ كل المحتوى (ابدأ من A1)، والصقه في ورقة جديدة في المصنف.
   - عدّل اسم الورقة ليطابق الاسم (Users, Categories, Warehouses, Suppliers, Items, PurchaseOrders, PO_Lines, StockTransactions, Transfers, AuditLog).
   - حدّد أي خلية داخل النطاق ثم Insert > Table، تأكد من تفعيل "My table has headers".
   - من Design > Table Name غيّر الاسم إلى: UsersTbl, CategoriesTbl, WarehousesTbl, SuppliersTbl, ItemsTbl, POsTbl, POLinesTbl, TxTbl, TransfersTbl, AuditTbl.
3. في TxTbl أضف عمودًا باسم "كمية_موقعة" إن لم يكن موجودًا، ثم يمكن تعبئته عبر الماكرو RecalculateQtySigned أو بالصيغة الموضحة في README.
4. افتح محرر VBA (Alt+F11) ثم:
   - Insert > Module، ألصق محتوى release/vba/Module_StockHelpers.bas
   - Insert > Module، ألصق محتوى release/vba/Module_LoginHandler.bas
   - Insert > UserForm، سمّه frmLogin، أضف عناصر تحكم: txtUsername, txtPassword (PasswordChar="*"), btnLogin, btnCancel.
     ثم ألصق كود release/vba/UserForm_frmLogin.txt في كود الفورم.
5. اختبر الماكروز: من Excel Developer > Macros شغّل RecalculateQtySigned وAddStockTransactionSimple.
6. أضف ورقة Inventory أو أنشئ PivotTables لعرض الرصيد الحالي حسب الصنف والمخزن (يمكن استخدام TxTbl -> Sum of كمية_موقعة).
7. احفظ الملف: File > Save As > Excel Macro-Enabled Workbook (*.xlsm) باسم School_Inventory.xlsm

ملاحظات نهائية:
- لتمكين ا��ماكروز تأكد من إعدادات أمان الماكروز في Excel.
- عند رفع الملف إلى Google Drive، اطلب من المستخدمين تنزيل الملف وفتحه في Excel لعمل الماكروز وواجهة المستخدم.
- إن رغبت أستطيع تجهيز ملف .xlsm ورفعه تلقائيًا إلى هذا الريبو إن منحتني إذن الكتابة (الذي قمت بتوفيره عبر الربط الحالي). هذا الالتزام سيضاف هنا ويكون متاحًا للتحميل.

إذا رغبت أرفع الآن ملف .xlsm جاهز إلى الريبو فأكد وأبدأ في تجميعه ورفعه.