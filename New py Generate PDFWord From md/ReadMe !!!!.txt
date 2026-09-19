## specific file.md

py WithIndexWithYouglishPronounce.py "RoadMap.md" 
py WithIndexWithYouglishPronounce.py "GemTable_PPL_WithVideos.md" 
 
## Folder All in 
py WithIndexWithYouglishPronounce.py "with pronounce"



مفيش طريقة في الـ Python/DOCX إني "أجبر" الإعداد ده يبقى مفعّل تلقائيًا لما تعمل الحفظ يدويًا — لأنه مش بيتقرأ من الملف، Word بس بيفتكر آخر اختيار عملته إنت في الحوار.
الحل (مرة واحدة بس، وبعدها هيفضل شغال دايمًا)
افتح أي DOCX من نفس الـ output بتاع السكربت.
File → Save As (أو Export → Create PDF/XPS).
اختار PDF واضغط زرار Options جنب اسم الملف.
فعّل الاختيار: Create bookmarks using → اختار Headings.
اعمل OK واحفظ.

من بعدها، Word هيفتكر الإعداد ده كإعداد افتراضي لأي Save As PDF تاني تعمله من نفس الجهاز — سواء لنفس الملف أو لغيره — لحد ما تلغي التفعيل بنفسك.


==========
=============

=============

تمام، لقيت الإعداد الرسمي بالظبط — Firefox فيه pref حقيقي جوه about:config اسمه pdfjs.externalLinkTarget، وده بيتحكم في سلوك أي رابط جوه الـ PDF viewer المدمج.

طريقة الضبط في Firefox

الإعداد اللي هيحل المشكلة: في about:config، دور على pdfjs.externalLinkTarget وغيّر قيمته من 0 (الافتراضي، بيفتح في نفس التاب) لـ 2 (يفتح في تاب جديد/blank).

هنا الفرق المهم عن Chrome و Edge: بما إن Firefox بيستخدم PDF.js كـ viewer افتراضي بتاعه، فده pref حقيقي موجود جوه المتصفح نفسه ومتحقق منه (مش تخمين)، ومفيش حاجة زيه في Chrome أو Edge لأن الـ viewer عندهم native binary مش PDF.js.

بالنسبة للروابط الجاية من Word لـ Firefox

ده pref تاني في نفس about:config:

browser.link.open_newwindow.override.external
غيّر قيمته لـ 3 → معناها "open external links in a new tab in the last active window"

يعني Firefox — على عكس Chrome و Edge — بيديك تحكم كامل في الحالتين عن طريق about:config، من غير ما تحتاج extensions أو حلول جانبية.