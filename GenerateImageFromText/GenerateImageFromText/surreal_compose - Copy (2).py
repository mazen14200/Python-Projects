#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
surreal_compose.py
-------------------
ياخد مجموعة كلمات من الـ command line بالشكل:

    python surreal_compose.py '{"desk","cat","car"}'

ولكل كلمة:
  1) يبحث في الإنترنت عن صورة (باستخدام DuckDuckGo، بدون أي API key).
  2) ينزّل الصورة في مجلد مؤقت.
  3) يشيل الخلفية منها (rembg).
  4) يلزقها بحجم/زاوية/موقع عشوائي فوق كانفاس واحد خلفيته بني فاتح (بيج).
  5) بعد إنشاء الصورة النهائية، يحذف كل الصور المؤقتة.

الناتج: final_composition.png (1024x1024) في نفس مجلد التشغيل.

المكتبات المطلوبة (يتم تثبيتها مرة واحدة فقط):
    pip install duckduckgo_search requests pillow rembg onnxruntime --break-system-packages
"""

import sys
import os
import ast
import re
import time
import random
import shutil
import requests
from datetime import datetime

from PIL import Image

# ============ إعدادات عامة ============
CANVAS_SIZE = (1024, 1024)
BACKGROUND_COLOR = (245, 230, 211)  # بني فاتح / بيج
TEMP_DIR = "temp_images"
# اسم ملف الناتج بيتحدد وقت التشغيل باسم فيه timestamp (شوف main) لتفادي أي تعارض أسماء

# نطاق الحجم العشوائي لكل عنصر (كنسبة من عرض الكانفاس)
MIN_SCALE = 0.30
MAX_SCALE = 0.55

# نطاق زاوية الميل العشوائية (بالدرجات)
MIN_ROTATION = -25
MAX_ROTATION = 25


def parse_input_words(raw_arg: str):
    """
    يحول النص المُدخل من الشكل:  {"desk","cat","car"}
    إلى قائمة كلمات.

    يحاول أولاً تفسيره كـ literal بايثون صحيح (set/list/tuple).
    لو فشلت المحاولة دي (غالبًا لأن Windows cmd بيمسح علامات
    التنصيص الداخلية لو اتكتبت من غير escape)، بيرجع لطريقة تفسير
    أبسط: يشيل الأقواس {} لو موجودة، وبيقسّم النص على الفواصل،
    وبيشيل أي علامات تنصيص متبقية من كل كلمة.
    """
    text = raw_arg.strip()

    # المحاولة 1: تفسير دقيق كـ Python literal (الحالة المثالية)
    try:
        parsed = ast.literal_eval(text)
        if isinstance(parsed, (set, list, tuple)):
            words = [str(w).strip() for w in parsed if str(w).strip()]
            if words:
                return words
    except (ValueError, SyntaxError):
        pass

    # المحاولة 2: تفسير متسامح (يتحمل غياب علامات التنصيص)
    stripped = text
    if stripped.startswith("{") and stripped.endswith("}"):
        stripped = stripped[1:-1]

    # لو فيه كلمات بين علامات تنصيص، ناخدها زي ما هي
    quoted_words = re.findall(r'"([^"]+)"|\'([^\']+)\'', stripped)
    words = []
    if quoted_words:
        for a, b in quoted_words:
            w = (a or b).strip()
            if w:
                words.append(w)
    else:
        # مفيش علامات تنصيص خالص -> نقسم على الفاصلة
        for part in stripped.split(","):
            w = part.strip().strip('"').strip("'")
            if w:
                words.append(w)

    if not words:
        raise ValueError(
            f"صيغة المدخل غير صحيحة أو فاضية. المتوقع مثل: {{\"desk\",\"cat\",\"car\"}}"
        )

    return words


def _get_ddgs_class():
    """
    يجيب كلاس DDGS من المكتبة الجديدة (ddgs) لو موجودة،
    ولو مش موجودة يرجع للقديمة (duckduckgo_search) للتوافق.
    """
    try:
        from ddgs import DDGS
        return DDGS
    except ImportError:
        from duckduckgo_search import DDGS
        return DDGS


def search_image_url(query: str, max_retries: int = 4):
    """
    يبحث عن صورة واحدة مناسبة لكل كلمة باستخدام DuckDuckGo Images.
    يعيد المحاولة تلقائيًا مع تأخير متزايد لو حصل Ratelimit (403).
    يرجع أول رابط صورة صالح.
    """
    DDGS = _get_ddgs_class()

    for attempt in range(1, max_retries + 1):
        try:
            with DDGS() as ddgs:
                try:
                    # المكتبة الجديدة (ddgs) بتستخدم query
                    results = ddgs.images(query=query, max_results=10)
                except TypeError:
                    # المكتبة القديمة (duckduckgo_search) بتستخدم keywords
                    results = ddgs.images(keywords=query, max_results=10)
                for r in results:
                    url = r.get("image")
                    if url:
                        return url
            return None
        except Exception as e:
            error_name = type(e).__name__
            if "Ratelimit" in error_name or "403" in str(e):
                wait = attempt * 5 + random.uniform(1, 3)
                print(
                    f"  [تنبيه] تم تقييد الطلبات (Ratelimit)، "
                    f"إعادة المحاولة بعد {wait:.1f} ثانية "
                    f"(محاولة {attempt}/{max_retries})..."
                )
                time.sleep(wait)
                continue
            print(f"  [خطأ] فشل البحث عن '{query}': {e}")
            return None

    print(f"  [خطأ] فشل البحث عن '{query}' بعد {max_retries} محاولات (Ratelimit مستمر).")
    return None


def download_image(url: str, save_path: str) -> bool:
    """ينزّل صورة من رابط معين ويحفظها في المسار المطلوب."""
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        with open(save_path, "wb") as f:
            f.write(resp.content)
        return True
    except Exception as e:
        print(f"  [تحذير] فشل تنزيل الصورة من {url}: {e}")
        return False


def remove_background(input_path: str, output_path: str) -> bool:
    """يشيل خلفية الصورة باستخدام rembg ويحفظها PNG بخلفية شفافة."""
    try:
        from rembg import remove

        with open(input_path, "rb") as f_in:
            input_bytes = f_in.read()

        output_bytes = remove(input_bytes)

        with open(output_path, "wb") as f_out:
            f_out.write(output_bytes)
        return True
    except Exception as e:
        print(f"  [تحذير] فشل إزالة الخلفية لـ {input_path}: {e}")
        return False


def fetch_and_prepare(word: str, index: int):
    """
    يبحث عن صورة للكلمة، ينزّلها، يشيل خلفيتها،
    ويرجع مسار الصورة النهائية (بخلفية شفافة) أو None لو فشل.
    """
    print(f"[{index}] بحث عن صورة لـ: {word}")
    url = search_image_url(word)
    if not url:
        print(f"  [خطأ] لم يتم إيجاد صورة لـ '{word}', سيتم تجاهلها.")
        return None

    raw_path = os.path.join(TEMP_DIR, f"{index}_{word}_raw.jpg")
    if not download_image(url, raw_path):
        return None

    # التأكد إن الصورة قابلة للفتح فعلاً (أحياناً السيرفر يرجع HTML بدل صورة)
    try:
        with Image.open(raw_path) as test_img:
            test_img.verify()
    except Exception:
        print(f"  [خطأ] الملف المُنزّل لـ '{word}' مش صورة صالحة.")
        return None

    nobg_path = os.path.join(TEMP_DIR, f"{index}_{word}_nobg.png")
    print(f"  إزالة الخلفية...")
    if not remove_background(raw_path, nobg_path):
        return None

    return nobg_path


def build_composition(image_paths, output_path):
    """
    يركّب كل الصور (بخلفياتها الشفافة) بشكل عشوائي فني
    فوق كانفاس واحد بخلفية بني فاتح.
    """
    canvas = Image.new("RGBA", CANVAS_SIZE, BACKGROUND_COLOR + (255,))

    canvas_w, canvas_h = CANVAS_SIZE

    for path in image_paths:
        try:
            element = Image.open(path).convert("RGBA")
        except Exception as e:
            print(f"  [تحذير] تعذر فتح {path}: {e}")
            continue

        # تحديد حجم عشوائي مع الحفاظ على نسبة العرض للطول
        scale = random.uniform(MIN_SCALE, MAX_SCALE)
        target_w = int(canvas_w * scale)
        ratio = target_w / element.width
        target_h = int(element.height * ratio)
        element = element.resize((target_w, target_h), Image.LANCZOS)

        # تدوير عشوائي (expand=True عشان منقصش من حواف العنصر)
        angle = random.uniform(MIN_ROTATION, MAX_ROTATION)
        element = element.rotate(angle, expand=True)

        # موقع عشوائي (نسمح بتداخل بسيط، لكن نتجنب الخروج الكامل بره الكانفاس)
        max_x = max(canvas_w - element.width, 0)
        max_y = max(canvas_h - element.height, 0)
        pos_x = random.randint(0, max_x) if max_x > 0 else 0
        pos_y = random.randint(0, max_y) if max_y > 0 else 0

        canvas.paste(element, (pos_x, pos_y), element)

    final_path = os.path.abspath(output_path)
    try:
        canvas.convert("RGB").save(final_path, "PNG")
        print(f"\n✅ تم إنشاء الصورة النهائية: {final_path}")
    except OSError as e:
        print(f"\n[خطأ] فشل حفظ الصورة في: {final_path}")
        print(f"  التفاصيل: {e}")
        print(f"  مجلد العمل الحالي: {os.getcwd()}")
        if os.path.exists(final_path):
            kind = "مجلد" if os.path.isdir(final_path) else "ملف"
            print(f"  ملاحظة: يوجد {kind} بنفس الاسم بالفعل في المسار ده، ممكن يكون سبب المشكلة.")

        # محاولة أخيرة: الحفظ في مجلد المستخدم الرئيسي بدل مجلد العمل الحالي
        fallback_path = os.path.join(os.path.expanduser("~"), os.path.basename(output_path))
        try:
            canvas.convert("RGB").save(fallback_path, "PNG")
            print(f"✅ تم الحفظ بنجاح في مسار بديل: {fallback_path}")
        except OSError as e2:
            print(f"[خطأ] فشلت المحاولة البديلة أيضًا: {e2}")
            raise


def cleanup():
    """يحذف مجلد الصور المؤقتة بالكامل."""
    if os.path.isdir(TEMP_DIR):
        shutil.rmtree(TEMP_DIR)
        print("🧹 تم حذف الصور المؤقتة.")


def main():
    if len(sys.argv) < 2:
        print('الاستخدام: python surreal_compose.py \'{"desk","cat","car"}\'')
        sys.exit(1)

    raw_arg = " ".join(sys.argv[1:])

    try:
        words = parse_input_words(raw_arg)
    except ValueError as e:
        print(f"[خطأ في المدخل] {e}")
        sys.exit(1)

    print(f"الكلمات المستخرجة: {words}\n")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"final_composition_{timestamp}.png"

    os.makedirs(TEMP_DIR, exist_ok=True)

    prepared_images = []
    try:
        for i, word in enumerate(words, start=1):
            result = fetch_and_prepare(word, i)
            if result:
                prepared_images.append(result)
            if i < len(words):
                time.sleep(random.uniform(2, 4))  # تهدئة بين الطلبات لتفادي Ratelimit

        if not prepared_images:
            print("\n[خطأ] لم يتم تجهيز أي صورة بنجاح. تم إيقاف العملية.")
            sys.exit(1)

        # ترتيب عشوائي للطبقات (أيهم فوق أيهم)
        random.shuffle(prepared_images)

        build_composition(prepared_images, output_path)

    finally:
        cleanup()


if __name__ == "__main__":
    main()