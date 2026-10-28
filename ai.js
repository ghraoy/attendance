// One interface over ChatGPT (OpenAI), Gemini (Google) and Claude (Anthropic).
// Keys stay on this device only; they are never synced to Drive.
window.AI = (() => {
  const NAMES = { openai: "ChatGPT", gemini: "Gemini", anthropic: "Claude" };

  async function check(r) {
    if (r.ok) return r.json();
    let msg = r.status + "";
    try { const j = await r.json(); msg = j.error?.message || j.error?.type || JSON.stringify(j).slice(0, 200); } catch (e) {}
    if (r.status === 401 || r.status === 403) msg = "المفتاح غير صحيح أو لا يملك صلاحية (" + msg + ")";
    if (r.status === 429) msg = "تجاوزت حد الاستخدام أو الرصيد في حساب المزوّد (" + msg + ")";
    throw new Error(msg);
  }

  async function listModels(provider, key) {
    if (provider === "openai") {
      const j = await check(await fetch("https://api.openai.com/v1/models", { headers: { Authorization: "Bearer " + key } }));
      return j.data.map(m => m.id).filter(id => /^(gpt|o\d|chatgpt)/.test(id) && !/(audio|realtime|tts|transcribe|search|image|embedding)/.test(id)).sort().reverse();
    }
    if (provider === "gemini") {
      const j = await check(await fetch("https://generativelanguage.googleapis.com/v1beta/models?pageSize=200&key=" + encodeURIComponent(key)));
      return j.models.filter(m => (m.supportedGenerationMethods || []).includes("generateContent") && /gemini/.test(m.name) && !/(tts|image|embedding|audio)/.test(m.name))
        .map(m => m.name.replace(/^models\//, "")).sort().reverse();
    }
    if (provider === "anthropic") {
      const j = await check(await fetch("https://api.anthropic.com/v1/models?limit=100", { headers: { "x-api-key": key, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" } }));
      return j.data.map(m => m.id);
    }
    return [];
  }

  // Pick a sensible default from the provider's own list: a fast, inexpensive model.
  function suggest(provider, models) {
    const pref = { openai: [/mini/, /^gpt/], gemini: [/flash(?!.*lite)/, /flash/, /pro/], anthropic: [/sonnet/, /haiku/, /opus/] }[provider] || [];
    for (const re of pref) { const m = models.find(x => re.test(x) && !/preview|exp/.test(x)) || models.find(x => re.test(x)); if (m) return m; }
    return models[0] || "";
  }

  // ask({system, text, image:{mime,b64}?, json:bool}) -> string
  async function ask(cfg, { system, text, image, maxTokens = 3000 }) {
    const p = cfg.provider, key = (cfg.keys || {})[p], model = (cfg.models || {})[p];
    if (!key) throw new Error("لم تُدخل مفتاح " + (NAMES[p] || "") + " في الإعدادات");
    if (!model) throw new Error("اختر نموذجاً لـ " + NAMES[p] + " في الإعدادات");
    if (p === "openai") {
      const content = [{ type: "text", text }];
      if (image) content.push({ type: "image_url", image_url: { url: `data:${image.mime};base64,${image.b64}` } });
      const j = await check(await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST", headers: { "Content-Type": "application/json", Authorization: "Bearer " + key },
        body: JSON.stringify({ model, messages: [{ role: "system", content: system }, { role: "user", content }] })
      }));
      return j.choices?.[0]?.message?.content || "";
    }
    if (p === "gemini") {
      const parts = [{ text }];
      if (image) parts.push({ inline_data: { mime_type: image.mime, data: image.b64 } });
      const j = await check(await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(key)}`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ systemInstruction: { parts: [{ text: system }] }, contents: [{ role: "user", parts }] })
      }));
      return (j.candidates?.[0]?.content?.parts || []).map(x => x.text || "").join("");
    }
    if (p === "anthropic") {
      const content = [];
      if (image) content.push({ type: "image", source: { type: "base64", media_type: image.mime, data: image.b64 } });
      content.push({ type: "text", text });
      const j = await check(await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
        body: JSON.stringify({ model, max_tokens: maxTokens, system, messages: [{ role: "user", content }] })
      }));
      return (j.content || []).filter(b => b.type === "text").map(b => b.text).join("");
    }
    throw new Error("اختر مزوّد الذكاء الاصطناعي في الإعدادات");
  }

  // Shrink a photo so uploads are fast and within provider limits.
  function imageFromFile(file, maxSide = 1800) {
    return new Promise((res, rej) => {
      const img = new Image(), url = URL.createObjectURL(file);
      img.onload = () => {
        const s = Math.min(1, maxSide / Math.max(img.width, img.height));
        const c = document.createElement("canvas"); c.width = Math.round(img.width * s); c.height = Math.round(img.height * s);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
        res({ mime: "image/jpeg", b64: c.toDataURL("image/jpeg", 0.85).split(",")[1] });
      };
      img.onerror = () => rej(new Error("تعذّر فتح الصورة"));
      img.src = url;
    });
  }

  function extractJSON(t) {
    const m = t.match(/\{[\s\S]*\}/); if (!m) throw new Error("لم يُرجع النموذج بيانات مفهومة");
    return JSON.parse(m[0]);
  }

  return { NAMES, listModels, suggest, ask, imageFromFile, extractJSON };
})();
