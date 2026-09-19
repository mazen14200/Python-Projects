#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py
-------
واجهة ويب محلية بسيطة (Flask) لسكريبت surreal_compose.py.

بتاخد الكلمات من textarea بأي صيغة (مفصولة بفواصل / مسافات / أسطر،
بعلامات تنصيص أو من غيرها، جوه أقواس {} أو من غيرها)، تحوّلها للصيغة
الأساسية اللي السكريبت فاهمها {"word1","word2",...}, وتشغّل بيها
surreal_compose.py فعليًا كـ subprocess، وتتابع سطور مخرجاته لحظيًا
عشان تحسب نسبة تقدم تقريبية، وتلتقط اسماء الصورتين النهائيتين أو أي
أخطاء حصلت.

التشغيل:
    pip install -r requirements.txt --break-system-packages
    python app.py
    وبعدين افتح المتصفح على: http://127.0.0.1:5000
"""

import os
import re
import sys
import uuid
import threading
import subprocess

from flask import Flask, request, jsonify, Response, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(BASE_DIR, "surreal_compose.py")

app = Flask(__name__)

INDEX_HTML = """<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>قصّ ولصق — صانع الكولاج السريالي</title>
<style>
:root {
  --bg: #efe3d0;
  --paper: #fbf6ec;
  --paper-soft: #fffdf8;
  --ink: #2b2620;
  --muted: #7a6f5d;
  --line: #d9c9ae;
  --clay: #c1502e;
  --clay-dark: #a3401f;
  --teal: #3f6357;
  --gold: #caa04a;
}

* { box-sizing: border-box; }

html, body {
  margin: 0;
  min-height: 100%;
  background: var(--bg);
  color: var(--ink);
}

body {
  font-family: -apple-system, "Segoe UI", Tahoma, Arial, sans-serif;
  direction: rtl;
  padding: 56px 20px 72px;
  display: flex;
  justify-content: center;
}

.page {
  width: 100%;
  max-width: 1040px;
}

/* ---------- Header ---------- */

header.hero {
  text-align: center;
  margin-bottom: 40px;
}

.chips {
  display: flex;
  justify-content: center;
  gap: 10px;
  margin-bottom: 20px;
}

.chip {
  width: 20px;
  height: 20px;
  border-radius: 50%;
  display: inline-block;
}
.chip:nth-child(1) { background: var(--clay); transform: rotate(-8deg); }
.chip:nth-child(2) { background: var(--teal); transform: rotate(6deg); }
.chip:nth-child(3) { background: var(--gold); transform: rotate(-4deg); }

h1 {
  font-family: Georgia, "Times New Roman", serif;
  font-size: clamp(34px, 5vw, 54px);
  margin: 0 0 10px;
  font-weight: 700;
  letter-spacing: -0.01em;
}

.subtitle {
  color: var(--muted);
  font-size: 17px;
  line-height: 1.7;
  max-width: 560px;
  margin: 0 auto;
}

/* ---------- Layout ---------- */

.grid {
  display: grid;
  grid-template-columns: 1.05fr 1fr;
  gap: 24px;
  align-items: start;
}

@media (max-width: 820px) {
  .grid { grid-template-columns: 1fr; }
}

.card {
  background: var(--paper);
  border: 1px solid var(--line);
  border-radius: 18px;
  padding: 28px;
}

.card h2 {
  font-family: Georgia, "Times New Roman", serif;
  font-size: 20px;
  margin: 0 0 8px;
}

.card .hint {
  color: var(--muted);
  font-size: 14px;
  line-height: 1.8;
  margin: 0 0 16px;
}

/* ---------- Input card ---------- */

textarea {
  width: 100%;
  min-height: 150px;
  resize: vertical;
  border: 1px solid var(--line);
  border-radius: 12px;
  background: var(--paper-soft);
  padding: 14px 16px;
  font-size: 15px;
  color: var(--ink);
  font-family: inherit;
  line-height: 1.7;
}

textarea:focus {
  outline: none;
  border-color: var(--clay);
  box-shadow: 0 0 0 3px rgba(193, 80, 46, 0.12);
}

.examples {
  margin-top: 14px;
  font-size: 13px;
  color: var(--muted);
  direction: ltr;
  text-align: left;
  background: #f4ecdd;
  border: 1px dashed var(--line);
  border-radius: 10px;
  padding: 10px 14px;
  line-height: 1.9;
}

.examples code { color: var(--ink); }

button#go {
  margin-top: 20px;
  width: 100%;
  padding: 14px 20px;
  border: none;
  border-radius: 12px;
  background: var(--clay);
  color: #fff9f2;
  font-size: 16px;
  font-weight: 600;
  cursor: pointer;
  transition: background 0.15s ease, transform 0.05s ease;
}

button#go:hover:not(:disabled) { background: var(--clay-dark); }
button#go:active:not(:disabled) { transform: scale(0.99); }
button#go:disabled {
  background: #cbbfa8;
  cursor: not-allowed;
}

/* ---------- Progress card ---------- */

.progress-shell {
  height: 14px;
  background: #f0e6d3;
  border: 1px solid var(--line);
  border-radius: 999px;
  overflow: hidden;
  margin-bottom: 12px;
}

.progress-fill {
  height: 100%;
  width: 0%;
  background: linear-gradient(90deg, var(--gold), var(--clay));
  transition: width 0.4s ease;
}

.status-line {
  font-size: 14px;
  color: var(--muted);
  margin-bottom: 14px;
  min-height: 20px;
}

.log {
  direction: ltr;
  text-align: left;
  background: var(--ink);
  color: #f2e9d8;
  font-family: "Cascadia Code", Consolas, "SFMono-Regular", monospace;
  font-size: 12.5px;
  line-height: 1.75;
  border-radius: 12px;
  padding: 14px 16px;
  height: 220px;
  overflow-y: auto;
  white-space: pre-wrap;
  word-break: break-word;
}

.log .err { color: #ff9d80; }
.log .ok { color: #a8dcc4; }

.result { display: none; margin-top: 18px; }
.result.show { display: block; }

.done-badge {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  background: rgba(63, 99, 87, 0.12);
  color: var(--teal);
  font-weight: 700;
  padding: 8px 14px;
  border-radius: 999px;
  font-size: 15px;
  margin-bottom: 14px;
}

.image-item {
  display: flex;
  align-items: center;
  gap: 12px;
  border: 1px solid var(--line);
  border-radius: 12px;
  padding: 10px;
  margin-bottom: 10px;
  background: var(--paper-soft);
}

.image-item img {
  width: 64px;
  height: 64px;
  object-fit: cover;
  border-radius: 8px;
  border: 1px solid var(--line);
  background: var(--bg);
  flex-shrink: 0;
}

.image-item .name {
  direction: ltr;
  text-align: left;
  font-family: Consolas, monospace;
  font-size: 12.5px;
  color: var(--ink);
  word-break: break-all;
}

.image-item a.open {
  margin-right: auto;
  font-size: 13px;
  color: var(--clay);
  text-decoration: none;
  font-weight: 600;
  white-space: nowrap;
}
.image-item a.open:hover { text-decoration: underline; }

.error-box {
  display: none;
  margin-top: 16px;
  background: rgba(193, 80, 46, 0.08);
  border: 1px solid rgba(193, 80, 46, 0.35);
  color: var(--clay-dark);
  padding: 12px 14px;
  border-radius: 10px;
  font-size: 14px;
  line-height: 1.8;
  white-space: pre-wrap;
}
.error-box.show { display: block; }

footer.note {
  text-align: center;
  color: var(--muted);
  font-size: 13px;
  margin-top: 36px;
}
</style>
</head>
<body>
<div class="page">

  <header class="hero">
    <div class="chips">
      <span class="chip"></span><span class="chip"></span><span class="chip"></span>
    </div>
    <h1>قصّ ولصق</h1>
    <p class="subtitle">
      اكتب أي كلمات إنجليزية بأي شكل تحبه، وهيّا يدوّر على صورها، يفصلها عن
      خلفيتها، ويلزقها في مشهد سريالي واحد.
    </p>
  </header>

  <div class="grid">

    <section class="card">
      <h2>الكلمات</h2>
      <p class="hint">
        اكتب الكلمات بأي صيغة — بفواصل، بعلامة +، بمسافات، أو كل كلمة
        في سطر، بعلامات تنصيص أو من غيرها، جوه أقواس {} أو [] أو () أو
        من غيرها خالص. هنحوّلها تلقائيًا للصيغة اللي السكريبت فاهمها.
      </p>

      <textarea id="words" placeholder="desk, cat, car"></textarea>

      <div class="examples">
        أمثلة مقبولة كلها:<br>
        <code>desk + cat + car</code><br>
        <code>(desk + cat + car)</code><br>
        <code>{desk + cat + car}</code><br>
        <code>[desk + cat + car]</code><br>
        <code>desk, cat, car</code><br>
        <code>"desk","cat","car"</code><br>
        <code>{"desk","cat","car"}</code><br>
        <code>desk<br>cat<br>car</code>
      </div>

      <button id="go">أنشئ الصورتين</button>
    </section>

    <section class="card">
      <h2>حالة العملية</h2>

      <div class="progress-shell">
        <div class="progress-fill" id="progressFill"></div>
      </div>
      <div class="status-line" id="statusLine">في انتظار البدء...</div>

      <div class="log" id="logBox"></div>

      <div class="result" id="result">
        <div class="done-badge">✅ Done — تم بنجاح</div>
        <div id="imagesList"></div>
      </div>

      <div class="error-box" id="errorBox"></div>
    </section>

  </div>

  <footer class="note">الصورتين بيتحفظوا جنب surreal_compose.py على جهازك.</footer>
</div>

<script>
const wordsEl = document.getElementById("words");
const goBtn = document.getElementById("go");
const progressFill = document.getElementById("progressFill");
const statusLine = document.getElementById("statusLine");
const logBox = document.getElementById("logBox");
const result = document.getElementById("result");
const imagesList = document.getElementById("imagesList");
const errorBox = document.getElementById("errorBox");

let pollTimer = null;
let renderedLogCount = 0;

function resetUI() {
  progressFill.style.width = "0%";
  statusLine.textContent = "جاري البدء...";
  logBox.innerHTML = "";
  renderedLogCount = 0;
  result.classList.remove("show");
  imagesList.innerHTML = "";
  errorBox.classList.remove("show");
  errorBox.textContent = "";
}

function appendLogLines(lines) {
  for (const line of lines) {
    const div = document.createElement("div");
    if (line.startsWith("[خطأ") || line.includes("Traceback")) {
      div.className = "err";
    } else if (line.startsWith("✅")) {
      div.className = "ok";
    }
    div.textContent = line;
    logBox.appendChild(div);
  }
  logBox.scrollTop = logBox.scrollHeight;
}

function renderImages(names) {
  imagesList.innerHTML = "";
  for (const name of names) {
    const url = "/images/" + encodeURIComponent(name);

    const row = document.createElement("div");
    row.className = "image-item";

    const img = document.createElement("img");
    img.src = url;
    img.alt = name;

    const nameDiv = document.createElement("div");
    nameDiv.className = "name";
    nameDiv.textContent = name;

    const link = document.createElement("a");
    link.className = "open";
    link.href = url;
    link.target = "_blank";
    link.rel = "noopener";
    link.textContent = "فتح";

    row.appendChild(img);
    row.appendChild(nameDiv);
    row.appendChild(link);
    imagesList.appendChild(row);
  }
}

async function startGeneration() {
  const text = wordsEl.value.trim();
  if (!text) {
    statusLine.textContent = "اكتب كلمة واحدة على الأقل الأول.";
    return;
  }

  resetUI();
  goBtn.disabled = true;
  goBtn.textContent = "جارٍ التنفيذ...";

  let data;
  try {
    const res = await fetch("/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ words_text: text }),
    });
    data = await res.json();
    if (!res.ok) {
      throw new Error(data.error || "حصل خطأ غير متوقع.");
    }
  } catch (err) {
    statusLine.textContent = "فشل الإرسال.";
    errorBox.textContent = err.message;
    errorBox.classList.add("show");
    goBtn.disabled = false;
    goBtn.textContent = "أنشئ الصورتين";
    return;
  }

  statusLine.textContent = "تم الإرسال: " + data.normalized;
  pollTimer = setInterval(() => pollStatus(data.job_id), 900);
}

async function pollStatus(jobId) {
  let data;
  try {
    const res = await fetch(`/status/${jobId}`);
    data = await res.json();
  } catch (err) {
    return; // فشل مؤقت في الشبكة، هنعيد المحاولة في الدورة الجاية
  }

  progressFill.style.width = (data.progress || 0) + "%";

  const newLines = data.log.slice(renderedLogCount);
  if (newLines.length) {
    appendLogLines(newLines);
    renderedLogCount = data.log.length;
  }

  if (data.status === "running") {
    statusLine.textContent = `جارٍ التنفيذ... (${data.progress || 0}%)`;
    return;
  }

  clearInterval(pollTimer);
  goBtn.disabled = false;
  goBtn.textContent = "أنشئ الصورتين";

  if (data.status === "done") {
    statusLine.textContent = "اكتملت العملية بنجاح.";
    progressFill.style.width = "100%";
    result.classList.add("show");
    renderImages(data.images);
  } else {
    statusLine.textContent = "حصل خطأ أثناء التنفيذ.";
    errorBox.textContent = data.error || "فشلت العملية لسبب غير معروف.";
    errorBox.classList.add("show");
  }
}

goBtn.addEventListener("click", startGeneration);
</script>
</body>
</html>
"""


jobs = {}
jobs_lock = threading.Lock()


# ============ تطبيع مدخلات الـ textarea لأي صيغة ============

def normalize_words(raw_text: str):
    """
    يستخرج قائمة كلمات نظيفة من أي صيغة يكتبها المستخدم في الـ textarea:
    فواصل / أسطر / مسافات / علامة +، بعلامات تنصيص أو من غيرها، جوه
    {} أو [] أو () أو من غيرها خالص.
    """
    text = raw_text.strip()

    if len(text) >= 2 and text[0] in "{[(" and text[-1] in "}])":
        text = text[1:-1]

    quoted = re.findall(r'"([^"]+)"|\'([^\']+)\'', text)
    words = []
    if quoted:
        for a, b in quoted:
            w = (a or b).strip()
            if w:
                words.append(w)
    else:
        for part in re.split(r'[,\n+]+', text):
            for w in part.strip().split():
                w = w.strip().strip('"\'')
                if w:
                    words.append(w)

    # إزالة أي تكرار (بنفس الحروف بالظبط) مع الحفاظ على ترتيب الظهور
    seen = set()
    unique_words = []
    for w in words:
        if w not in seen:
            seen.add(w)
            unique_words.append(w)

    return unique_words


def build_set_literal(words) -> str:
    """يبني الصيغة الأساسية اللي بياخدها سكريبت surreal_compose.py."""
    inner = ",".join(f'"{w}"' for w in words)
    return "{" + inner + "}"


# ============ تشغيل السكريبت ومتابعة تقدمه ============

def run_job(job_id: str, words):
    job = jobs[job_id]
    literal_arg = build_set_literal(words)

    with jobs_lock:
        job["log"].append(f"$ python surreal_compose.py {literal_arg}")
        job["progress"] = 3

    total_words = max(len(words), 1)
    saved_paths = []
    search_index = 0

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    try:
        proc = subprocess.Popen(
            [sys.executable, SCRIPT_PATH, literal_arg],
            cwd=BASE_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=env,
        )

        for line in proc.stdout:
            line = line.rstrip("\n")
            if not line:
                continue

            with jobs_lock:
                job["log"].append(line)

                if line.startswith("الكلمات المستخرجة"):
                    job["progress"] = max(job["progress"], 5)
                elif re.match(r"^\[\d+\] بحث عن صورة لـ:", line):
                    search_index += 1
                    pct = 5 + (search_index / total_words) * 55
                    job["progress"] = max(job["progress"], int(pct))
                elif "إزالة الخلفية" in line:
                    pct = 5 + (search_index / total_words) * 55 + 5
                    job["progress"] = max(job["progress"], int(min(pct, 65)))
                elif "ترتيب العناصر وضمان نسبة الظهور" in line:
                    job["progress"] = max(job["progress"], 72)
                elif re.match(r"^\s*(✅|⚠️) عنصر", line):
                    job["progress"] = max(job["progress"], 85)
                elif line.startswith("✅ تم الحفظ:"):
                    path = line.split("✅ تم الحفظ:", 1)[1].strip()
                    saved_paths.append(path)
                    job["progress"] = 92 if len(saved_paths) == 1 else 98
                elif line.startswith("[خطأ") or "Traceback" in line:
                    job["error_lines"].append(line)

        proc.wait()

        with jobs_lock:
            if proc.returncode == 0 and len(saved_paths) >= 1:
                job["status"] = "done"
                job["progress"] = 100
                job["images"] = [os.path.basename(p) for p in saved_paths]
            else:
                job["status"] = "error"
                job["error"] = (
                    "\n".join(job["error_lines"][-6:])
                    if job["error_lines"]
                    else "فشلت العملية لسبب غير معروف. راجع السجل بالأسفل."
                )

    except Exception as e:
        with jobs_lock:
            job["status"] = "error"
            job["error"] = f"تعذر تشغيل السكريبت: {e}"


# ============ مسارات (routes) الويب ============

@app.route("/")
def index():
    return Response(INDEX_HTML, mimetype="text/html")


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True) or {}
    raw_text = data.get("words_text", "")

    words = normalize_words(raw_text)
    if not words:
        return jsonify({"error": "اكتب كلمة واحدة على الأقل."}), 400

    if not os.path.exists(SCRIPT_PATH):
        return jsonify({
            "error": "لم يتم إيجاد surreal_compose.py في نفس مجلد app.py."
        }), 400

    job_id = uuid.uuid4().hex
    jobs[job_id] = {
        "status": "running",
        "progress": 0,
        "log": [],
        "error_lines": [],
        "error": None,
        "images": [],
    }

    thread = threading.Thread(target=run_job, args=(job_id, words), daemon=True)
    thread.start()

    return jsonify({"job_id": job_id, "normalized": build_set_literal(words)})


@app.route("/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "job غير موجود"}), 404

    with jobs_lock:
        return jsonify({
            "status": job["status"],
            "progress": job["progress"],
            "log": list(job["log"]),
            "images": list(job["images"]),
            "error": job["error"],
        })


@app.route("/images/<path:filename>")
def serve_image(filename):
    return send_from_directory(BASE_DIR, filename)


if __name__ == "__main__":
    print("افتح المتصفح على: http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
