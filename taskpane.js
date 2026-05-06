var isListening = false;
var ghostCC = null;
var ghostText = "";
var pollTimer = null;
var idleTimer = null;
var isGenerating = false;
var lastDocText = "";
var ignoreNextChange = false;
var generationId = 0;

var POLL_MS = 500;
var IDLE_MS = 2000;
var MAX_SUGGESTION_LEN = 30;

// ── Settings ──
function loadSettings() {
  document.getElementById("api-url").value = localStorage.getItem("ai_url") || "";
  document.getElementById("api-key").value = localStorage.getItem("ai_key") || "";
  document.getElementById("api-model").value = localStorage.getItem("ai_model") || "deepseek-chat";
}

function saveSettings() {
  localStorage.setItem("ai_url", document.getElementById("api-url").value);
  localStorage.setItem("ai_key", document.getElementById("api-key").value);
  localStorage.setItem("ai_model", document.getElementById("api-model").value);
  log("设置已保存");
}

function fillDeepSeek() {
  document.getElementById("api-url").value = "https://api.deepseek.com/v1/chat/completions";
  document.getElementById("api-model").value = "deepseek-chat";
}

function fillOpenAI() {
  document.getElementById("api-url").value = "https://api.openai.com/v1/chat/completions";
  document.getElementById("api-model").value = "gpt-3.5-turbo";
}

function fillZhipu() {
  document.getElementById("api-url").value = "https://open.bigmodel.cn/api/paas/v4/chat/completions";
  document.getElementById("api-model").value = "glm-4-flash";
}

// ── Log ──
function log(msg) {
  var area = document.getElementById("log-area");
  var ts = new Date().toLocaleTimeString();
  area.innerHTML = "<div>[" + ts + "] " + msg + "</div>" + area.innerHTML;
  if (area.children.length > 30) area.removeChild(area.lastChild);
}

// ── Status UI ──
function setStatus(on) {
  document.getElementById("status-indicator").className = "indicator " + (on ? "running" : "stopped");
  document.getElementById("status-text").textContent = on ? "运行中" : "未启动";
  document.getElementById("btn-start").disabled = on;
  document.getElementById("btn-stop").disabled = !on;
}

// ── Start / Stop ──
function startListening() {
  if (isListening) return;
  isListening = true;
  lastDocText = "";
  generationId = 0;
  setStatus(true);
  log("续写已启动");

  getDocText().then(function (t) { lastDocText = t; });

  pollTimer = setInterval(pollDocument, POLL_MS);
}

function stopListening() {
  isListening = false;
  setStatus(false);
  clearInterval(pollTimer);
  pollTimer = null;
  clearTimeout(idleTimer);
  idleTimer = null;
  removeGhost();
  log("续写已停止");
}

// ── Poll: detect doc changes & ghost state ──
async function pollDocument() {
  if (!isListening) return;

  try {
    // Check ghost CC state first
    if (ghostCC) {
      await checkGhostState();
      return; // don't check doc changes while ghost is active
    }

    if (isGenerating) return;

    var currentText = await getDocText();
    if (!currentText || currentText === lastDocText) return;

    if (ignoreNextChange) {
      lastDocText = currentText;
      ignoreNextChange = false;
      return;
    }

    lastDocText = currentText;

    // Reset idle timer
    clearTimeout(idleTimer);
    idleTimer = setTimeout(function () {
      generateSuggestion();
    }, IDLE_MS);

  } catch (_) {}
}

// ── Check ghost CC state ──
async function checkGhostState() {
  try {
    await Word.run(async function (ctx) {
      var cc = ghostCC;
      cc.load("text");
      await ctx.sync();

      if (!cc.text) {
        // CC was deleted by user
        ghostCC = null;
        ghostText = "";
        log("建议已删除");
        return;
      }

      if (cc.text !== ghostText) {
        // User edited inside the ghost text → delete it
        var range = cc.getRange(Word.RangeLocation.whole);
        range.delete(Word.DeleteMode.delete);
        await ctx.sync();
        ghostCC = null;
        ghostText = "";
        lastDocText = await getDocTextRaw(ctx);
        log("建议已清除");
        return;
      }

      // Text unchanged → user typed AFTER the ghost → accept
      cc.font.color = "black";
      cc.font.italic = false;
      cc.removeContentControl();
      await ctx.sync();

      var accepted = ghostText;
      ghostCC = null;
      ghostText = "";
      lastDocText = await getDocTextRaw(ctx);
      ignoreNextChange = true;
      log("已自动采纳: " + accepted);
    });
  } catch (err) {
    // CC may have been invalidated
    ghostCC = null;
    ghostText = "";
  }
}

// ── Get document text ──
async function getDocText() {
  return Word.run(function (ctx) {
    return getDocTextRaw(ctx);
  });
}

function getDocTextRaw(ctx) {
  var body = ctx.document.body;
  body.load("text");
  return ctx.sync().then(function () {
    return body.text || "";
  });
}

// ── Extract recent context ──
function getRecentContext(fullText) {
  var paragraphs = fullText.split(/\n|\r/);
  var lastPara = "";
  for (var i = paragraphs.length - 1; i >= 0; i--) {
    var p = paragraphs[i].trim();
    if (p) { lastPara = p; break; }
  }
  if (lastPara.length > 120) lastPara = lastPara.slice(-120);

  var prevPara = "";
  for (var j = paragraphs.length - 2; j >= 0; j--) {
    var pp = paragraphs[j].trim();
    if (pp) { prevPara = pp; break; }
  }
  if (prevPara.length > 80) prevPara = prevPara.slice(-80);

  return prevPara ? prevPara + "\n" + lastPara : lastPara;
}

// ── Generate suggestion ──
async function generateSuggestion() {
  var url = document.getElementById("api-url").value.trim();
  var key = document.getElementById("api-key").value.trim();
  var model = document.getElementById("api-model").value.trim();

  if (!url || !key) {
    log("请填写 API 地址和 Key");
    return;
  }

  var myGenId = ++generationId;
  isGenerating = true;
  log("正在生成...");

  try {
    var context = getRecentContext(lastDocText);

    var resp = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + key,
      },
      body: JSON.stringify({
        model: model || "deepseek-chat",
        messages: [
          {
            role: "system",
            content: "你是中文写作续写助手。根据用户最近输入的文字，自然地续写接下来的内容。要求：续写不超过30个字，语意连贯自然，不加引号，不要换行，不要重复用户已写的内容。只输出续写文字。",
          },
          { role: "user", content: context },
        ],
        max_tokens: 80,
        temperature: 0.7,
      }),
    });

    // Check if this generation is still current
    if (myGenId !== generationId) return;

    if (!resp.ok) {
      var errText = await resp.text();
      throw new Error("API " + resp.status + ": " + errText.slice(0, 200));
    }

    var data = await resp.json();
    var text = data.choices[0].message.content.trim();

    text = text.replace(/^["'"「【（\(]+/, "");
    text = text.replace(/["'"」】）\)]+$/, "");
    text = text.replace(/^[\s,，。、；：！？]+/, "");

    if (text.length > MAX_SUGGESTION_LEN) text = text.slice(0, MAX_SUGGESTION_LEN);

    if (text && myGenId === generationId) {
      await insertGhost(text);
      log("已生成: " + text);
    }
  } catch (err) {
    log("生成失败: " + err.message);
  } finally {
    isGenerating = false;
  }
}

// ── Insert ghost text into document ──
async function insertGhost(text) {
  await removeGhost();

  try {
    await Word.run(async function (ctx) {
      var sel = ctx.document.getSelection();
      sel.insertText(text, Word.InsertLocation.end);
      var cc = sel.insertContentControl();
      cc.tag = "ai-ghost";
      cc.appearance = Word.ContentControlAppearance.hidden;
      cc.font.color = "gray";
      cc.font.italic = true;

      ghostCC = cc;
      ghostText = text;
      ignoreNextChange = true;

      await ctx.sync();
    });
  } catch (err) {
    log("插入失败: " + err.message);
  }
}

// ── Remove ghost text from document ──
async function removeGhost() {
  if (!ghostCC) return;
  var cc = ghostCC;
  ghostCC = null;
  ghostText = "";
  try {
    await Word.run(async function (ctx) {
      cc.load("text");
      await ctx.sync();
      if (cc.text) {
        var range = cc.getRange(Word.RangeLocation.whole);
        range.delete(Word.DeleteMode.delete);
        await ctx.sync();
        lastDocText = await getDocTextRaw(ctx);
      }
    });
  } catch (_) {}
}

// ── Init ──
Office.onReady(function () {
  loadSettings();
  log("插件已就绪");
});
