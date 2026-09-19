
py generatePdfFromMd.py "RoadMap.md"  

py generatePdfTablesAndIndexingFromMd.py "RoadMap.md"  





مفيش طريقة في الـ Python/DOCX إني "أجبر" الإعداد ده يبقى مفعّل تلقائيًا لما تعمل الحفظ يدويًا — لأنه مش بيتقرأ من الملف، Word بس بيفتكر آخر اختيار عملته إنت في الحوار.
الحل (مرة واحدة بس، وبعدها هيفضل شغال دايمًا)
افتح أي DOCX من نفس الـ output بتاع السكربت.
File → Save As (أو Export → Create PDF/XPS).
اختار PDF واضغط زرار Options جنب اسم الملف.
فعّل الاختيار: Create bookmarks using → اختار Headings.
اعمل OK واحفظ.

من بعدها، Word هيفتكر الإعداد ده كإعداد افتراضي لأي Save As PDF تاني تعمله من نفس الجهاز — سواء لنفس الملف أو لغيره — لحد ما تلغي التفعيل بنفسك.