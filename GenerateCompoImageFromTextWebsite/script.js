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
