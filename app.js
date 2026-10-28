/* ===== state ===== */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const STATUS = { P: "حاضر", A: "غائب", N: "غير مطالب", E: "مستوفية" };
const fmtDate = d => { const [y, m, dd] = d.split("-"); return `${+dd}/${+m}/${y}`; };
const dayName = d => { try { return new Date(d + "T12:00:00").toLocaleDateString("ar-IQ", { weekday: "long" }); } catch (e) { return ""; } };
const today = () => { const t = new Date(); return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`; };
const uid = () => Math.random().toString(36).slice(2, 9);
const short = n => { const w = String(n).split(" "); return /^(عبد|ابو|أبو)$/.test(w[0]) && w[1] ? w[0] + " " + w[1] : w[0]; };
const pct = x => (x * 100).toFixed(1).replace(/\.0$/, "") + "%";
const clone = o => JSON.parse(JSON.stringify(o));

const KEY = "mtu-attendance-v1", CFG = "mtu-config-v1";
const load = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } };
const store = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } };

let S = load(KEY, null) || clone(window.SEED);
let C = load(CFG, { provider: "gemini", keys: {}, models: {}, drive: { clientId: "" }, lastSync: null });
C.keys ||= {}; C.models ||= {}; C.drive ||= { clientId: "" };
const DEFAULT_CLIENT_ID = "629045560897-1ritqqkb0qcul2km6ifik0mrq315ru3s.apps.googleusercontent.com";
if (!C.drive.clientId && location.hostname === "ghraoy.github.io") C.drive.clientId = DEFAULT_CLIENT_ID;
const ui = { subject: "", student: "", onlyAbs: false, preview: null, answer: "", asking: false, letter: "", photoBusy: false, modelLists: {} };

function saveCfg() { store(CFG, C); }
function commit(msg) {
  S.updatedAt = new Date().toISOString();
  if (!store(KEY, S)) toast("تعذّر الحفظ على هذا الجهاز — مساحة المتصفح ممتلئة أو محجوبة");
  render(); if (msg) toast(msg);
  scheduleSync();
}

/* ===== attendance logic ===== */
const WEEKDAYS = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"];
const offDays = () => S.settings.offDays ?? [3, 4]; // default: no duty on Wednesday and Thursday
const isOffDay = d => offDays().includes(new Date(d + "T12:00:00").getDay());
function status(sess, stu) {
  if (sess.date < stu.start) return "N";
  if ((stu.exempt || []).includes(sess.subject)) return "E";
  const absent = (sess.absent || []).includes(stu.id);
  if (absent && isOffDay(sess.date)) return "N"; // absence on a no-duty day is never counted
  return absent ? "A" : "P";
}
const subj = id => S.subjects.find(s => s.id === id);
function stats(stuId, subId) {
  const stu = S.students.find(s => s.id === stuId);
  let req = 0, abs = 0, hreq = 0, habs = 0; const dates = [];
  for (const se of S.sessions) {
    if (subId && se.subject !== subId) continue;
    const st = status(se, stu); if (st === "N" || st === "E") continue;
    const h = +(subj(se.subject)?.hours || 0);
    req++; hreq += h; if (st === "A") { abs++; habs += h; dates.push(se.date); }
  }
  // planned semester hours (official calendar): weeks this student is required × lectures/week × hours
  let htot = 0;
  for (const su of (subId ? [subj(subId)] : S.subjects)) if (su && !(stu.exempt || []).includes(su.id)) htot += studentWeeks(stu) * (+su.perWeek || 1) * (+su.hours || 0);
  return { req, abs, hreq, habs, htot, rate: htot ? habs / htot : 0, left1: Math.max(0, Math.floor(htot * W().w1 / 100 - habs)), left2: Math.max(0, Math.floor(htot * W().w2 / 100 - habs)), dates: dates.sort() };
}
const term = () => ({ start: S.settings.termStart || "2026-09-20", end: S.settings.termEnd || "2026-12-31", weeks: +S.settings.termWeeks || 15 });
function studentWeeks(stu) {
  const t = term(); if (!stu.start || stu.start <= t.start) return t.weeks;
  const days = (new Date(t.end + "T12:00:00") - new Date(stu.start + "T12:00:00")) / 86400e3 + 1;
  return Math.max(0, Math.min(t.weeks, Math.ceil(days / 7)));
}
const W = () => ({ w1: +S.settings.warn1 || 5, w2: +S.settings.warn2 || 7 }); // الإنذار الأولي / الإنذار الثاني (% من ساعات الفصل)
const level = rate => { const { w1, w2 } = W(); return rate * 100 >= w2 ? "d" : rate * 100 >= w1 ? "w" : ""; };
const levelText = rate => ({ d: "تجاوز الإنذار الثاني", w: "تجاوز الإنذار الأولي" }[level(rate)] || "");

/* ===== helpers ===== */
function toast(t) { const el = $("#toast"); el.textContent = t; el.hidden = false; clearTimeout(el._t); el._t = setTimeout(() => el.hidden = true, 2600); }
function twoStep(btn, fn, label = "حذف") { if (btn.dataset.arm) { fn(); return; } btn.dataset.arm = 1; const old = btn.textContent; btn.textContent = "تأكيد " + label; setTimeout(() => { if (btn.isConnected) { delete btn.dataset.arm; btn.textContent = old; } }, 3000); }
function saveFile(name, data, type) {
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([data], { type })); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
async function copy(text) { try { await navigator.clipboard.writeText(text); toast("نُسخ النص"); } catch (e) { toast("تعذّر النسخ — حدّد النص وانسخه يدوياً"); } }
const aiReady = () => !!(C.keys[C.provider] && C.models[C.provider]);

/* ===== views ===== */
const VIEWS = ["overview", "log", "import", "ai", "setup"];
const view = () => { const h = location.hash.slice(1); return VIEWS.includes(h) ? h : "overview"; };
function render() {
  document.querySelectorAll("nav a").forEach(a => a.setAttribute("aria-current", a.getAttribute("href") === "#" + view() ? "page" : "false"));
  $("#termLine").textContent = [S.settings.program, S.settings.term].filter(Boolean).join(" · ");
  const v = view();
  $("#main").innerHTML = { overview, log: logView, import: importView, ai: aiView, setup: setupView }[v]();
  bind(v); syncBadge();
}

function overview() {
  const { w1, w2 } = W(), scale = w2 * 1.5;
  if (!S.students.length || !S.subjects.length) return `<div class="panel" style="text-align:center;justify-items:center;padding:32px"><h2 style="margin:0">ابدأ بملف Excel</h2><p class="muted" style="margin:0;max-width:52ch">السجل فارغ على هذا الجهاز. استورد ملف الحضور الذي عندك (ورقة لكل مادة، أو جدول بأعمدة التاريخ والمادة والطالب والحالة) فتُنشأ المواد والطلبة والمحاضرات تلقائياً. أو أضفها يدوياً من «الإعدادات».</p><div class="row" style="justify-content:center"><a class="btn primary" href="#import">استيراد ملف Excel</a><a class="btn" href="#setup">إضافة يدوية</a></div><p class="muted small" style="margin:0">إذا ربطت Google Drive على جهاز آخر، اضغط «اتصال» في الإعدادات فيُسحب السجل.</p></div>`;
  const cards = S.students.map(st => {
    const a = stats(st.id), lv = level(a.rate);
    let worst = null; for (const su of S.subjects) { const b = stats(st.id, su.id); if (b.req && (!worst || b.rate > worst.rate)) worst = { ...b, name: su.name }; }
    return `<div class="card">
      <div class="name">${esc(st.name)}</div>
      <div class="start">يُحسب الدوام من ${fmtDate(st.start)}</div>
      <div class="row" style="align-items:baseline"><span class="big">${a.habs}</span><span class="muted">ساعة غياب من ${a.htot} ساعة مقررة للفصل</span></div>
      <div class="bar" title="الإنذار الأولي ${w1}% · الثاني ${w2}%"><span class="${lv}" style="width:${Math.min(100, a.rate * 100 / scale * 100)}%"></span><b style="right:${w1 / scale * 100}%"></b><b style="right:${w2 / scale * 100}%"></b></div>
      <div class="kv"><span>عدد الغيابات</span><strong>${a.abs} من ${a.req} محاضرة أُعطيت</strong></div>
      <div class="kv"><span>باقي حتى الإنذار الأولي (${w1}%)</span><strong class="${a.left1 ? "" : "w"}">${a.left1 ? a.left1 + " ساعة" : "تجاوزه"}</strong></div>
      <div class="kv"><span>باقي حتى الإنذار الثاني (${w2}%)</span><strong class="${a.left2 ? "" : "d"}">${a.left2 ? a.left2 + " ساعة" : "تجاوزه"}</strong></div>
      <div class="kv"><span>نسبة الغياب الكلية</span><strong class="${lv}">${pct(a.rate)}</strong></div>
      ${worst && worst.abs ? `<div class="kv"><span>أعلى مادة</span><span class="${level(worst.rate)}">${esc(worst.name)} · ${pct(worst.rate)}</span></div>` : ""}
    </div>`;
  }).join("");
  const head = S.students.map(s => `<th>${esc(short(s.name))}</th>`).join("");
  const rows = S.subjects.map(su => `<tr><td class="l"><strong>${esc(su.name)}</strong><br><small class="muted">${su.hours} س/محاضرة</small></td>${S.students.map(st => {
    if ((st.exempt || []).includes(su.id)) return `<td><span class="pill p-E">مستوفية</span></td>`;
    const b = stats(st.id, su.id); if (!b.req) return `<td class="muted">—</td>`;
    return `<td><span class="cell ${level(b.rate)}"><strong>${b.abs} / ${b.req}</strong><small>${b.habs} من ${b.htot} س · ${pct(b.rate)}</small></span></td>`;
  }).join("")}</tr>`).join("");
  const foot = S.students.map(st => { const a = stats(st.id); return `<td>${a.habs} / ${a.htot} س</td>`; }).join("");
  return `<section><div class="row"><h2>الطلبة</h2><span class="spacer"></span><span class="muted small">الإنذار الأولي ${w1}% · الثاني ${w2}% من ساعات الفصل</span></div><div class="cards">${cards}</div></section>
  <section><div class="row"><h2>الغياب حسب المادة</h2><span class="spacer"></span><button class="btn small" id="xlsx" type="button">تنزيل Excel</button><button class="btn small" id="xlsxDrive" type="button">رفع نسخة إلى Drive</button></div>
  <div class="scroll"><table><thead><tr><th class="l">المادة</th>${head}</tr></thead><tbody>${rows}</tbody><tfoot><tr><td class="l">ساعات الغياب / المطالب بها</td>${foot}</tr></tfoot></table></div>
  <p class="muted small" style="margin:8px 0 0">في كل خانة: الغيابات / المحاضرات التي أُعطيت، وتحتها ساعات الغياب من مجموع ساعات الفصل المقررة (${term().weeks} أسبوعاً حسب التقويم) والنسبة. البرتقالي = تجاوز الإنذار الأولي، والأحمر = تجاوز الإنذار الثاني.</p></section>`;
}

function logView() {
  const subOpts = `<option value="">كل المواد</option>` + S.subjects.map(s => `<option value="${s.id}" ${ui.subject === s.id ? "selected" : ""}>${esc(s.name)}</option>`).join("");
  const stuOpts = `<option value="">كل الطلبة</option>` + S.students.map(s => `<option value="${s.id}" ${ui.student === s.id ? "selected" : ""}>${esc(s.name)}</option>`).join("");
  let list = [...S.sessions].sort((a, b) => b.date.localeCompare(a.date) || (subj(a.subject)?.name || "").localeCompare(subj(b.subject)?.name || ""));
  if (ui.subject) list = list.filter(s => s.subject === ui.subject);
  if (ui.onlyAbs) list = list.filter(se => S.students.some(st => (!ui.student || st.id === ui.student) && status(se, st) === "A"));
  const groups = {}; list.forEach(s => (groups[s.date] = groups[s.date] || []).push(s));
  const body = Object.keys(groups).map(d => `<div class="day"><div class="dayhead">${fmtDate(d)} <span>${dayName(d)}</span></div>${groups[d].map(se => {
    const su = subj(se.subject) || { name: "مادة محذوفة", teacher: "", hours: 0 };
    const chips = S.students.filter(st => !ui.student || st.id === ui.student).map(st => { const x = status(se, st); return `<span class="pill p-${x}">${esc(short(st.name))}: ${STATUS[x]}</span>`; }).join("");
    return `<button class="sess" type="button" data-id="${se.id}"><div><div class="t">${esc(su.name)}</div><div class="m">${esc(su.teacher)}${su.teacher ? " · " : ""}${su.hours} ساعة</div></div><span class="m">تعديل</span><div class="chips">${chips}</div>${se.note ? `<div class="note" style="grid-column:1/-1">${esc(se.note)}</div>` : ""}</button>`;
  }).join("")}</div>`).join("");
  return `<section><h2>أرشيف المحاضرات</h2>
   <div class="filters"><label class="field">المادة<select id="fSub">${subOpts}</select></label>
   <label class="field">الطالب<select id="fStu">${stuOpts}</select></label>
   <label class="check" style="padding-bottom:8px"><input type="checkbox" id="fAbs" ${ui.onlyAbs ? "checked" : ""}> الغيابات فقط</label>
   <span class="spacer"></span><span class="muted small" style="padding-bottom:8px">${list.length} محاضرة</span></div></section>
   <section style="display:grid;gap:18px">${body || `<div class="empty">لا توجد محاضرات بهذا التصفية. اضغط «+ محاضرة» لإضافة محاضرة.</div>`}</section>`;
}

function importView() {
  const pv = ui.preview;
  const preview = pv ? `<div class="panel" style="border-color:var(--accent)"><h3>معاينة قبل الحفظ</h3>
    <div class="small">${pv.created} محاضرة جديدة · ${pv.updated} محاضرة موجودة ستُحدَّث · ${pv.absences} غياب مسجّل في الملف</div>
    ${pv.newSubjects.length ? `<div class="banner">مواد جديدة ستُضاف (راجع ساعاتها من الإعدادات): ${pv.newSubjects.map(s => esc(s.name)).join("، ")}</div>` : ""}
    ${pv.newStudents.length ? `<div class="banner">طلبة جدد سيُضافون: ${pv.newStudents.map(s => esc(s.name) + " (من " + fmtDate(s.start) + ")").join("، ")}. إن كان الاسم مكتوباً بشكل مختلف لطالب موجود، صحّحه في الملف وأعد الاستيراد.</div>` : ""}
    ${(pv.newExempt || []).length ? `<div class="banner ok">مستوفون سيُسجَّلون: ${pv.newExempt.map(esc).join("، ")}</div>` : ""}
    ${!pv.created && !pv.updated ? `<div class="banner err">لم أجد أي محاضرة في الملف. تأكد أن فيه عموداً باسم «التاريخ».</div>` : ""}
    <div class="row"><button class="btn primary" id="applyImp" type="button" ${!pv.created && !pv.updated ? "disabled" : ""}>حفظ في السجل</button><button class="btn" id="cancelImp" type="button">إلغاء</button></div></div>` : "";
  return `<section><h2>إدخال البيانات</h2><p class="muted small" style="margin:0">أربع طرق: محاضرة محاضرة بزر «+ محاضرة»، أو ملف Excel، أو لصق جدول، أو صورة ورقة الحضور.</p></section>
  ${preview}
  <div class="panels">
    <div class="panel"><h3>صورة ورقة الحضور</h3>
      <p class="muted small" style="margin:0">صوّر الورقة الموقّعة، فيقرأ ${esc(AI.NAMES[C.provider] || "الذكاء الاصطناعي")} التاريخ والمادة والأسماء ويعبّئ المحاضرة، ثم تراجعها قبل الحفظ.</p>
      ${aiReady() ? "" : `<div class="banner">أضف مفتاح ChatGPT أو Gemini أو Claude من «الإعدادات» لتفعيل هذه الطريقة.</div>`}
      <label class="btn primary" style="justify-self:start;${aiReady() && !ui.photoBusy ? "" : "opacity:.5;pointer-events:none"}">${ui.photoBusy ? "جارٍ قراءة الورقة…" : "التقاط أو اختيار صورة"}<input type="file" id="photo" accept="image/*" capture="environment" hidden></label>
    </div>
    <div class="panel"><h3>ملف Excel</h3>
      <p class="muted small" style="margin:0">يقبل ملف السجل الذي صدّرته سابقاً (ورقة لكل مادة)، أو جدولاً بأعمدة: التاريخ، المادة، الطالب، الحالة. الحالة تُكتب «حاضر» أو «غائب».</p>
      <div class="row"><label class="btn primary">اختيار ملف<input type="file" id="xfile" accept=".xlsx,.xls,.csv" hidden></label><button class="btn" id="tpl" type="button">تنزيل نموذج فارغ</button></div>
    </div>
    <div class="panel"><h3>لصق جدول</h3>
      <label class="field">انسخ الخلايا من Excel أو Google Sheets والصقها هنا مع سطر العناوين<textarea id="paste" placeholder="التاريخ	المادة	الطالب	الحالة&#10;2026-11-01	الوبائيات	محمد علي حسن	غائب"></textarea></label>
      <label class="field">المادة إن لم يكن في الجدول عمود «المادة»<select id="pasteSub">${S.subjects.map(s => `<option value="${esc(s.name)}">${esc(s.name)}</option>`).join("")}</select></label>
      <button class="btn primary" id="pasteGo" type="button" style="justify-self:start">قراءة الجدول</button>
    </div>
    <div class="panel"><h3>نسخة احتياطية</h3>
      <p class="muted small" style="margin:0">ملف واحد فيه كل السجل، تحفظه أو تنقله لجهاز آخر يدوياً.</p>
      <div class="row"><button class="btn" id="bkSave" type="button">حفظ نسخة</button><label class="btn">استرجاع نسخة<input type="file" id="bkLoad" accept=".json" hidden></label></div>
    </div>
  </div>`;
}

function aiView() {
  const quick = ["من تجاوز خط الإنذار أو اقترب منه؟", "لخّص غيابات كل طالب حسب المادة", "ما المواد التي فيها أكثر غياب؟", "كم ساعة يستطيع كل طالب أن يغيب بعد قبل الإنذار؟"];
  return `<section><h2>المساعد</h2><p class="muted small" style="margin:0">${aiReady() ? `يعمل الآن بـ ${esc(AI.NAMES[C.provider])} (${esc(C.models[C.provider])}). تُرسل إليه أرقام السجل مع سؤالك فقط.` : "أضف مفتاح ChatGPT أو Gemini أو Claude من «الإعدادات» لتفعيل الأسئلة وصياغة الكتب بالذكاء الاصطناعي. قالب كتاب الإنذار يعمل بدون مفتاح."}</p></section>
  <div class="panels">
    <div class="panel"><h3>اسأل عن السجل</h3>
      <div class="row">${quick.map(q => `<button class="chipbtn" type="button" data-q="${esc(q)}">${esc(q)}</button>`).join("")}</div>
      <label class="field">سؤالك<textarea id="q" style="min-height:80px" placeholder="مثلاً: كم غياب على الطالب في الوبائيات ومتى؟"></textarea></label>
      <button class="btn primary" id="askGo" type="button" style="justify-self:start" ${aiReady() && !ui.asking ? "" : "disabled"}>${ui.asking ? "جارٍ التفكير…" : "اسأل"}</button>
      ${ui.answer ? `<div class="answer" id="ans">${esc(ui.answer)}</div><div class="row"><button class="btn small" id="ansCopy" type="button">نسخ</button></div>` : ""}
    </div>
    <div class="panel"><h3>كتاب إنذار</h3>
      <div class="grid2">
        <label class="field">الطالب<select id="lStu">${S.students.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join("")}</select></label>
        <label class="field">المادة<select id="lSub"><option value="">كل المواد</option>${S.subjects.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join("")}</select></label>
        <label class="field">النوع<select id="lType"><option>إنذار أول</option><option>إنذار ثانٍ</option><option>إنذار نهائي</option><option>إشعار حرمان</option></select></label>
        <label class="field">جهة الإصدار<input id="lFrom" value="${esc(S.settings.letterFrom || "رئاسة قسم تقنيات صحة المجتمع")}"></label>
      </div>
      <div class="row"><button class="btn" id="lTpl" type="button">قالب جاهز</button><button class="btn primary" id="lAi" type="button" ${aiReady() ? "" : "disabled"}>صياغة بالذكاء الاصطناعي</button></div>
      ${ui.letter ? `<label class="field">الكتاب (يمكنك تعديله)<textarea id="letter" style="min-height:280px">${esc(ui.letter)}</textarea></label>
      <div class="row"><button class="btn small" id="lCopy" type="button">نسخ</button><button class="btn small" id="lPrint" type="button">طباعة</button><button class="btn small" id="lDoc" type="button">حفظ Word</button></div>` : ""}
    </div>
  </div>`;
}

function setupView() {
  const subRows = S.subjects.map(s => `<tr data-sub="${s.id}"><td><input data-k="name" value="${esc(s.name)}" aria-label="اسم المادة"></td><td><input data-k="teacher" value="${esc(s.teacher)}" aria-label="الأستاذ"></td><td><input data-k="hours" type="number" min="0" step="0.5" value="${s.hours}" aria-label="الساعات"></td><td><input data-k="perWeek" type="number" min="0" step="1" value="${s.perWeek || 1}" aria-label="محاضرات في الأسبوع"></td><td class="muted">${term().weeks * (s.perWeek || 1) * s.hours} س</td><td><button class="btn small danger" type="button" data-delsub="${s.id}">حذف</button></td></tr>`).join("");
  const stuRows = S.students.map(st => `<tr data-stu="${st.id}"><td><input data-k="name" value="${esc(st.name)}" aria-label="اسم الطالب"></td><td><input data-k="start" type="date" value="${st.start}" aria-label="تاريخ بدء الدوام"></td><td class="l"><div class="ex">${S.subjects.map(su => `<label><input type="checkbox" data-ex="${su.id}" ${(st.exempt || []).includes(su.id) ? "checked" : ""}>${esc(su.name)}</label>`).join("")}</div></td><td><button class="btn small danger" type="button" data-delstu="${st.id}">حذف</button></td></tr>`).join("");
  const p = C.provider, list = ui.modelLists[p] || [];
  const models = [...new Set([C.models[p], ...list].filter(Boolean))];
  return `<section class="setup panel"><div class="row"><h2 style="margin:0">المواد</h2><span class="spacer"></span><button class="btn small" id="addSub" type="button">+ مادة</button></div>
   <div class="scroll"><table><thead><tr><th class="l">المادة</th><th class="l">الأستاذ</th><th>ساعات المحاضرة</th><th>محاضرات/أسبوع</th><th>ساعات الفصل</th><th></th></tr></thead><tbody>${subRows}</tbody></table></div></section>
   <section class="setup panel"><div class="row"><h2 style="margin:0">الطلبة</h2><span class="spacer"></span><button class="btn small" id="addStu" type="button">+ طالب</button></div>
   <div class="scroll"><table><thead><tr><th class="l">الاسم</th><th>يُحسب من</th><th class="l">مواد مستوفاة</th><th></th></tr></thead><tbody>${stuRows}</tbody></table></div></section>
   <section class="panel"><h2 style="margin:0">عام</h2><div class="grid2">
   <label class="field">الإنذار الأولي (% من ساعات الفصل)<input id="w1" type="number" min="0" max="100" step="0.5" value="${W().w1}"></label>
   <label class="field">الإنذار الثاني (% من ساعات الفصل)<input id="w2" type="number" min="0" max="100" step="0.5" value="${W().w2}"></label>
   <div class="field" style="grid-column:1/-1">أيام لا دوام فيها (الغياب فيها لا يُحسب، والحضور يُحسب)<div class="ex">${WEEKDAYS.map((d, i) => `<label><input type="checkbox" data-off="${i}" ${offDays().includes(i) ? "checked" : ""}>${d}</label>`).join("")}</div></div>
   <label class="field">بداية الفصل (حسب التقويم)<input id="tStart" type="date" value="${term().start}"></label>
   <label class="field">نهاية الفصل<input id="tEnd" type="date" value="${term().end}"></label>
   <label class="field">عدد أسابيع الفصل<input id="tWeeks" type="number" min="1" max="30" value="${term().weeks}"></label>
   <p class="muted small" style="grid-column:1/-1;margin:0">نسبة الغياب = ساعات الغياب ÷ ساعات الفصل المقررة (الأسابيع × محاضرات الأسبوع × ساعات المحاضرة). الطالب الذي بدأ بعد بداية الفصل تُحسب له الأسابيع من تاريخ بدئه.</p>
   <label class="field">الفصل الدراسي<input id="term" value="${esc(S.settings.term)}"></label>
   <label class="field" style="grid-column:1/-1">البرنامج والقسم<input id="prog" value="${esc(S.settings.program)}"></label></div></section>

   <section class="panel"><h2 style="margin:0">الذكاء الاصطناعي</h2>
   <p class="muted small" style="margin:0">المفاتيح تُحفظ على هذا الجهاز فقط ولا تُرفع إلى Drive. أدخلها على كل جهاز تستخدمه.</p>
   <div class="grid2">
     <label class="field">المزوّد<select id="prov">${Object.entries(AI.NAMES).map(([k, v]) => `<option value="${k}" ${k === p ? "selected" : ""}>${v}</option>`).join("")}</select></label>
     <label class="field">النموذج<select id="model">${models.length ? models.map(m => `<option ${m === C.models[p] ? "selected" : ""}>${esc(m)}</option>`).join("") : `<option value="">أدخل المفتاح ثم اضغط «جلب النماذج»</option>`}</select></label>
     <label class="field" style="grid-column:1/-1">مفتاح ${AI.NAMES[p]} (API key)<input id="key" type="password" autocomplete="off" value="${esc(C.keys[p] || "")}" placeholder="${{ openai: "sk-…", gemini: "AIza…", anthropic: "sk-ant-…" }[p]}"></label>
   </div>
   <div class="row"><button class="btn" id="fetchModels" type="button">جلب النماذج واختبار المفتاح</button><span class="muted small" id="aiMsg"></span></div>
   <details class="small"><summary>من أين أحصل على المفتاح؟</summary><ol class="steps" style="margin-top:8px">
     <li>ChatGPT: <code>platform.openai.com/api-keys</code> ← Create new secret key (يحتاج رصيداً مدفوعاً في حساب API).</li>
     <li>Gemini: <code>aistudio.google.com/apikey</code> ← Create API key (فيه استخدام مجاني محدود).</li>
     <li>Claude: <code>console.anthropic.com/settings/keys</code> ← Create Key (يحتاج رصيداً).</li>
   </ol></details></section>

   <section class="panel"><h2 style="margin:0">Google Drive</h2>
   <div class="${Drive.connected() ? "banner ok" : "banner"}">${Drive.connected() ? "متصل. يُرفع السجل تلقائياً بعد كل تعديل إلى مجلد «" + Drive.FOLDER + "»." : "غير متصل. يحتاج إعداداً لمرة واحدة (الخطوات أدناه)، ثم تضغط «اتصال»."}${C.lastSync ? "<br>آخر مزامنة: " + new Date(C.lastSync).toLocaleString("ar-IQ") : ""}</div>
   <label class="field">Google OAuth Client ID<input id="cid" value="${esc(C.drive.clientId || "")}" placeholder="xxxxxxxx.apps.googleusercontent.com" style="direction:ltr"></label>
   <div class="row"><button class="btn primary" id="driveGo" type="button">${Drive.connected() ? "مزامنة الآن" : "اتصال"}</button>${Drive.connected() ? `<button class="btn" id="driveOut" type="button">قطع الاتصال</button>` : ""}</div>
   <details class="small" ${C.drive.clientId ? "" : "open"}><summary>خطوات الإعداد (مرة واحدة، حوالي 10 دقائق)</summary><ol class="steps" style="margin-top:8px">
     <li>افتح <code>console.cloud.google.com</code> بحساب Google الذي تريد الحفظ فيه، وأنشئ مشروعاً جديداً (New Project) باسم «حضور الماجستير».</li>
     <li>من القائمة: APIs &amp; Services ← Library، ابحث عن <b>Google Drive API</b> واضغط Enable.</li>
     <li>APIs &amp; Services ← OAuth consent screen (أو Google Auth Platform ← Branding): اختر External، اكتب اسم التطبيق وبريدك، ثم من Audience أضف بريدك في Test users.</li>
     <li>Credentials ← Create Credentials ← OAuth client ID ← النوع <b>Web application</b>.</li>
     <li>في Authorized JavaScript origins أضف هذا العنوان بالضبط: <code>${esc(location.origin)}</code></li>
     <li>اضغط Create وانسخ الـ Client ID، الصقه في الخانة أعلاه، ثم اضغط «اتصال» ووافق على الصلاحية.</li>
     <li>كرّر الضغط على «اتصال» على الموبايل بنفس الحساب، فيسحب السجل نفسه.</li>
   </ol></details></section>

   <section class="panel"><h2 style="margin:0">التحديثات</h2>
   <p class="muted small" style="margin:0">التطبيق يتحدّث تلقائياً عند فتحه إذا وُجدت نسخة جديدة، ولا تفقد بياناتك. الإصدار الحالي: <b>${APP_VERSION}</b></p>
   <div class="row"><button class="btn" id="updBtn" type="button">البحث عن تحديث الآن</button></div></section>
   <section class="panel"><h2 style="margin:0">إعادة الضبط</h2>
   <div class="row"><button class="btn danger" id="wipeBtn" type="button">مسح كل البيانات والبدء من جديد</button></div>
   <p class="muted small" style="margin:0">احفظ نسخة احتياطية قبل أي منهما.</p></section>`;
}

/* ===== bindings ===== */
function bind(v) {
  if (v === "overview" && $("#xlsx")) { $("#xlsx").onclick = () => saveFile(`حضور الماجستير ${today()}.xlsx`, buildXlsx(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); $("#xlsxDrive").onclick = sheetToDrive; }
  if (v === "log") {
    $("#fSub").onchange = e => { ui.subject = e.target.value; render(); };
    $("#fStu").onchange = e => { ui.student = e.target.value; render(); };
    $("#fAbs").onchange = e => { ui.onlyAbs = e.target.checked; render(); };
    document.querySelectorAll(".sess").forEach(b => b.onclick = () => openSess(b.dataset.id));
  }
  if (v === "import") {
    $("#photo").onchange = e => { const f = e.target.files[0]; e.target.value = ""; if (f) readPhoto(f); };
    $("#xfile").onchange = async e => {
      const f = e.target.files[0]; e.target.value = ""; if (!f) return;
      try { const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true }); ui.preview = Importer.plan(S, Importer.fromWorkbook(wb)); render(); }
      catch (err) { toast("تعذّر قراءة الملف: " + err.message); }
    };
    $("#tpl").onclick = () => saveFile("نموذج إدخال الحضور.xlsx", Importer.template(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    $("#pasteGo").onclick = () => { const t = $("#paste").value; if (!t.trim()) { toast("الصق الجدول أولاً"); return; } ui.preview = Importer.plan(S, Importer.fromText(t, $("#pasteSub").value)); render(); };
    if ($("#applyImp")) { $("#applyImp").onclick = () => { const pv = ui.preview; S = pv.next; ui.preview = null; commit(`حُفظ: ${pv.created} محاضرة جديدة و${pv.updated} محدّثة`); }; $("#cancelImp").onclick = () => { ui.preview = null; render(); }; }
    $("#bkSave").onclick = () => saveFile(`نسخة سجل الحضور ${today()}.json`, JSON.stringify(S, null, 1), "application/json");
    $("#bkLoad").onchange = async e => {
      const f = e.target.files[0]; e.target.value = ""; if (!f) return;
      try { const j = JSON.parse(await f.text()); if (!j.students || !j.sessions) throw new Error("الملف ليس نسخة من هذا السجل"); S = j; commit("استُرجعت النسخة"); } catch (err) { toast(err.message); }
    };
  }
  if (v === "ai") {
    document.querySelectorAll("[data-q]").forEach(b => b.onclick = () => { $("#q").value = b.dataset.q; askAI(); });
    $("#askGo").onclick = askAI;
    if ($("#ansCopy")) $("#ansCopy").onclick = () => copy(ui.answer);
    $("#lFrom").onchange = e => { S.settings.letterFrom = e.target.value; commit(); };
    $("#lTpl").onclick = () => { ui.letter = letterTemplate(); render(); };
    $("#lAi").onclick = letterAI;
    if ($("#letter")) {
      $("#letter").oninput = e => ui.letter = e.target.value;
      $("#lCopy").onclick = () => copy(ui.letter);
      $("#lPrint").onclick = () => { $("#printArea").textContent = ui.letter; window.print(); };
      $("#lDoc").onclick = () => saveFile(`كتاب ${$("#lType").value}.doc`, `<html dir="rtl"><head><meta charset="utf-8"></head><body style="font-family:Arial;font-size:14pt;line-height:2;white-space:pre-wrap">${esc(ui.letter)}</body></html>`, "application/msword");
    }
  }
  if (v === "setup") {
    document.querySelectorAll("tr[data-sub] input").forEach(i => i.onchange = () => { const s = subj(i.closest("tr").dataset.sub); s[i.dataset.k] = ["hours", "perWeek"].includes(i.dataset.k) ? Math.max(0, +i.value || 0) : i.value.trim(); commit(); });
    document.querySelectorAll("tr[data-stu] input[data-k]").forEach(i => i.onchange = () => { const s = S.students.find(x => x.id === i.closest("tr").dataset.stu); s[i.dataset.k] = i.value.trim() || s[i.dataset.k]; commit(); });
    document.querySelectorAll("tr[data-stu] input[data-ex]").forEach(i => i.onchange = () => { const s = S.students.find(x => x.id === i.closest("tr").dataset.stu); s.exempt = (s.exempt || []).filter(x => x !== i.dataset.ex); if (i.checked) s.exempt.push(i.dataset.ex); commit(); });
    document.querySelectorAll("[data-delsub]").forEach(b => b.onclick = () => twoStep(b, () => { S.subjects = S.subjects.filter(s => s.id !== b.dataset.delsub); commit(); }));
    document.querySelectorAll("[data-delstu]").forEach(b => b.onclick = () => twoStep(b, () => { S.students = S.students.filter(s => s.id !== b.dataset.delstu); commit(); }));
    $("#addSub").onclick = () => { S.subjects.push({ id: "s" + uid(), name: "مادة جديدة", teacher: "", hours: 2 }); commit(); };
    $("#addStu").onclick = () => { S.students.push({ id: "t" + uid(), name: "طالب جديد", start: today(), exempt: [] }); commit(); };
    document.querySelectorAll("[data-off]").forEach(c => c.onchange = () => { S.settings.offDays = [...document.querySelectorAll("[data-off]:checked")].map(x => +x.dataset.off); commit("حُدّثت أيام الدوام"); });
    $("#tStart").onchange = e => { S.settings.termStart = e.target.value; commit(); };
    $("#tEnd").onchange = e => { S.settings.termEnd = e.target.value; commit(); };
    $("#tWeeks").onchange = e => { S.settings.termWeeks = Math.max(1, +e.target.value || 15); commit(); };
    $("#w1").onchange = e => { S.settings.warn1 = Math.min(100, Math.max(0, +e.target.value || 5)); commit(); };
    $("#w2").onchange = e => { S.settings.warn2 = Math.min(100, Math.max(0, +e.target.value || 7)); commit(); };
    $("#term").onchange = e => { S.settings.term = e.target.value; commit(); };
    $("#prog").onchange = e => { S.settings.program = e.target.value; commit(); };
    $("#prov").onchange = e => { C.provider = e.target.value; saveCfg(); render(); };
    $("#model").onchange = e => { C.models[C.provider] = e.target.value; saveCfg(); toast("اختير النموذج"); };
    $("#key").onchange = e => { C.keys[C.provider] = e.target.value.trim(); saveCfg(); };
    $("#fetchModels").onclick = async () => {
      C.keys[C.provider] = $("#key").value.trim(); saveCfg();
      const msg = $("#aiMsg"); msg.textContent = "جارٍ الاتصال…"; msg.className = "muted small";
      try {
        const list = await AI.listModels(C.provider, C.keys[C.provider]); ui.modelLists[C.provider] = list;
        if (!C.models[C.provider] || !list.includes(C.models[C.provider])) C.models[C.provider] = AI.suggest(C.provider, list);
        saveCfg(); render(); toast(`المفتاح يعمل — ${list.length} نموذج متاح`);
      } catch (err) { msg.textContent = err.message; msg.className = "small d"; }
    };
    $("#cid").onchange = e => { C.drive.clientId = e.target.value.trim(); saveCfg(); };
    $("#driveGo").onclick = () => { C.drive.clientId = $("#cid").value.trim(); saveCfg(); syncNow(true); };
    if ($("#driveOut")) $("#driveOut").onclick = () => { Drive.signOut(); render(); toast("قُطع الاتصال بـ Drive"); };
    $("#updBtn").onclick = async () => { toast("جارٍ البحث عن تحديث…"); try { await window._swReg?.update(); } catch (e) {} setTimeout(() => location.reload(), 1500); };
    $("#wipeBtn").onclick = e => twoStep(e.target, () => { S = { students: [], subjects: [], sessions: [], settings: { ...S.settings } }; commit("بدأ سجل فارغ"); }, "المسح");
  }
}

/* ===== session editor ===== */
let draft = null;
function openSess(id, preset, banner) {
  if (!S.subjects.length) { toast("أضف مادة أولاً من «الإعدادات»"); return; }
  const ex = id && S.sessions.find(s => s.id === id);
  draft = preset || (ex ? clone(ex) : { id: "x" + uid(), date: today(), subject: ui.subject || S.subjects[0].id, note: "", absent: [] });
  draft._banner = banner || "";
  drawSess(!!ex); if (!$("#dlg").open) $("#dlg").showModal();
}
function drawSess(isEdit) {
  const f = $("#sessForm");
  const att = S.students.map(st => {
    const x = status(draft, st);
    const ctrl = (x === "N" || x === "E") ? `<span class="pill p-${x}">${STATUS[x]}</span>` :
      `<span class="seg"><button type="button" class="P" data-st="${st.id}" data-v="P" aria-pressed="${x === "P"}">حاضر</button><button type="button" class="A" data-st="${st.id}" data-v="A" aria-pressed="${x === "A"}">غائب</button></span>`;
    return `<div class="attrow"><span>${esc(st.name)}</span>${ctrl}</div>`;
  }).join("");
  const dup = S.sessions.find(s => s.id !== draft.id && s.date === draft.date && s.subject === draft.subject);
  f.innerHTML = `<h2 style="margin:0">${isEdit ? "تعديل محاضرة" : "محاضرة جديدة"}</h2>
   ${draft._banner ? `<div class="banner">${esc(draft._banner)}</div>` : ""}
   <div class="grid2"><label class="field">التاريخ<input type="date" id="dDate" value="${draft.date}" required></label>
   <label class="field">المادة<select id="dSub">${S.subjects.map(s => `<option value="${s.id}" ${s.id === draft.subject ? "selected" : ""}>${esc(s.name)}</option>`).join("")}</select></label></div>
   ${dup ? `<div class="banner">هذه المادة مسجّلة في نفس اليوم. الحفظ سيُدمج الحضور مع المحاضرة الموجودة.</div>` : ""}
   <div class="att">${att}</div>
   <label class="field">ملاحظة (اختياري)<input id="dNote" value="${esc(draft.note)}" placeholder="مثلاً: امتحان"></label>
   <div class="row">${isEdit ? `<button type="button" class="btn danger" id="dDel">حذف</button>` : ""}<span class="spacer"></span>
   <button type="button" class="btn" id="dCancel">إلغاء</button><button type="button" class="btn primary" id="dSave">حفظ</button></div>`;
  $("#dDate").onchange = e => { draft.date = e.target.value || draft.date; drawSess(isEdit); };
  $("#dSub").onchange = e => { draft.subject = e.target.value; drawSess(isEdit); };
  $("#dNote").oninput = e => draft.note = e.target.value;
  f.querySelectorAll("[data-st]").forEach(b => b.onclick = () => { draft.absent = draft.absent.filter(x => x !== b.dataset.st); if (b.dataset.v === "A") draft.absent.push(b.dataset.st); drawSess(isEdit); });
  $("#dCancel").onclick = () => $("#dlg").close();
  $("#dSave").onclick = () => {
    const d = clone(draft); delete d._banner; d.note = $("#dNote").value.trim();
    const i = S.sessions.findIndex(s => s.id === d.id);
    if (dup && i < 0) { dup.absent = d.absent; if (d.note) dup.note = d.note; }
    else if (dup && i >= 0) { dup.absent = d.absent; if (d.note) dup.note = d.note; S.sessions.splice(i, 1); }
    else if (i >= 0) S.sessions[i] = d; else S.sessions.push(d);
    $("#dlg").close(); commit(i >= 0 ? "عُدّلت المحاضرة" : "أُضيفت المحاضرة");
  };
  if (isEdit) $("#dDel").onclick = () => twoStep($("#dDel"), () => { S.sessions = S.sessions.filter(s => s.id !== draft.id); $("#dlg").close(); commit("حُذفت المحاضرة"); });
}
$("#addBtn").onclick = () => openSess(null);

/* ===== AI features ===== */
async function readPhoto(file) {
  ui.photoBusy = true; render();
  try {
    const image = await AI.imageFromFile(file);
    const system = "أنت تقرأ صور أوراق حضور جامعية عراقية مكتوبة بخط اليد وتُرجع JSON فقط بدون أي نص آخر.";
    const text = `هذه صورة ورقة حضور لطلبة الدراسات العليا. في أعلاها: اليوم، التاريخ (يوم/شهر/سنة)، مدرس المادة، المادة. وفيها جدول بعمود «الاسم» وعمود «التوقيع». الطالب الحاضر هو من كُتب اسمه في الجدول. قد تحتوي الصورة نسختين متطابقتين من الورقة (أعلى وأسفل) — اعتبرهما ورقة واحدة.
الطلبة المسجلون: ${JSON.stringify(S.students.map(s => ({ id: s.id, name: s.name })))}
المواد: ${JSON.stringify(S.subjects.map(s => ({ id: s.id, name: s.name, teacher: s.teacher })))}
أرجع JSON بهذا الشكل بالضبط:
{"date":"YYYY-MM-DD أو null","subject_id":"معرّف المادة من القائمة أو null","subject_text":"اسم المادة كما هو مكتوب","teacher_text":"اسم المدرس كما هو مكتوب","present_ids":["معرّفات الطلبة الذين وردت أسماؤهم"],"unmatched_names":["أسماء مكتوبة لم تطابق أي طالب"],"confidence":"high أو medium أو low","notes":"أي ملاحظة مثل امتحان أو كتابة غير واضحة"}`;
    const out = AI.extractJSON(await AI.ask(C, { system, text, image }));
    const date = Importer.toYMD(out.date) || today();
    const subject = subj(out.subject_id) ? out.subject_id : (Importer.matchBy(S.subjects, out.subject_text || "")?.id || S.subjects[0].id);
    const present = new Set(out.present_ids || []);
    const preset = { id: "x" + uid(), date, subject, note: out.notes && /امتحان/.test(out.notes) ? "امتحان" : "", absent: [] };
    S.students.forEach(st => { const x = status(preset, st); if (x === "P" && !present.has(st.id)) preset.absent.push(st.id); });
    const warn = [`قُرئت من الصورة بواسطة ${AI.NAMES[C.provider]} — راجع التاريخ والمادة والحضور قبل الحفظ.`];
    if (out.subject_text) warn.push(`المادة المكتوبة: «${out.subject_text}».`);
    if ((out.unmatched_names || []).length) warn.push(`أسماء لم أطابقها: ${out.unmatched_names.join("، ")}.`);
    if (out.confidence === "low") warn.push("الكتابة غير واضحة في أجزاء من الورقة.");
    if (out.notes) warn.push(out.notes);
    ui.photoBusy = false; render();
    openSess(null, preset, warn.join(" "));
  } catch (err) { ui.photoBusy = false; render(); toast("تعذّر قراءة الصورة: " + err.message); }
}

function dataForAI() {
  return {
    الإعدادات: { الإنذار_الأولي: W().w1 + "% من ساعات الفصل", الإنذار_الثاني: W().w2 + "% من ساعات الفصل", الفصل: S.settings.term, اليوم: today() },
    المواد: S.subjects.map(s => ({ المادة: s.name, المدرس: s.teacher, ساعات_المحاضرة: s.hours })),
    الطلبة: S.students.map(st => ({
      الاسم: st.name, يحسب_الدوام_من: st.start, مستوفي_لـ: (st.exempt || []).map(id => subj(id)?.name).filter(Boolean),
      الإجمالي: (a => ({ محاضرات_مطالب_بها: a.req, غيابات: a.abs, ساعات_الفصل_المقررة: a.htot, ساعات_المحاضرات_المعطاة: a.hreq, ساعات_غياب: a.habs, نسبة: pct(a.rate) }))(stats(st.id)),
      حسب_المادة: S.subjects.filter(su => !(st.exempt || []).includes(su.id)).map(su => { const b = stats(st.id, su.id); return { المادة: su.name, محاضرات: b.req, غيابات: b.abs, ساعات_غياب: b.habs, من_ساعات_الفصل: b.htot, نسبة: pct(b.rate), تواريخ_الغياب: b.dates }; })
    }))
  };
}
async function askAI() {
  const q = $("#q").value.trim(); if (!q) { toast("اكتب سؤالك أولاً"); return; }
  ui.asking = true; ui.answer = ""; render(); $("#q").value = q;
  try {
    ui.answer = await AI.ask(C, { system: "أنت مساعد إداري في قسم أكاديمي عراقي. أجب بالعربية الفصحى المختصرة اعتماداً على بيانات السجل المرفقة فقط، واذكر الأرقام بدقة. لا تخترع بيانات غير موجودة. لا تستخدم تنسيق Markdown.", text: `بيانات السجل:\n${JSON.stringify(dataForAI())}\n\nالسؤال: ${q}` });
  } catch (err) { ui.answer = "تعذّر الحصول على إجابة: " + err.message; }
  ui.asking = false; render(); const qq = $("#q"); if (qq) qq.value = q;
}
function letterFacts() {
  const st = S.students.find(s => s.id === $("#lStu").value), subId = $("#lSub").value, su = subj(subId);
  const a = stats(st.id, subId || undefined);
  return { st, su, a, type: $("#lType").value, from: $("#lFrom").value };
}
function letterTemplate() {
  const { st, su, a, type, from } = letterFacts();
  const t = new Date().toLocaleDateString("ar-IQ", { year: "numeric", month: "long", day: "numeric" });
  return `${from}
العدد:
التاريخ: ${t}

إلى/ الطالب ${st.name}
${S.settings.program || ""}

م/ ${type}

نظراً لتغيّبك عن ${su ? `مادة «${su.name}»` : "المحاضرات"} ${a.abs} ${a.abs === 1 ? "مرة" : "مرات"} بواقع ${a.habs} ساعة من أصل ${a.htot} ساعة مقررة، أي بنسبة غياب قدرها ${pct(a.rate)}${a.dates.length ? `، وذلك في التواريخ: ${a.dates.map(fmtDate).join("، ")}` : ""}، وحيث إن تعليمات الدراسات العليا تحدد توجيه الإنذار الأولي عند تجاوز ${W().w1}% والإنذار الثاني عند تجاوز ${W().w2}% من الساعات المقررة، لذا تقرر توجيه ${type} إليك، مع التأكيد على ضرورة الالتزام بالدوام، وبخلافه ستُتخذ الإجراءات القانونية وفق التعليمات النافذة.

مع التقدير.

رئيس القسم
التوقيع:

نسخة منه إلى:
- لجنة الدراسات العليا في القسم
- الحفظ`;
}
async function letterAI() {
  const btn = $("#lAi"); btn.disabled = true; btn.textContent = "جارٍ الصياغة…";
  const { st, su, a, type, from } = letterFacts();
  try {
    ui.letter = (await AI.ask(C, {
      system: "أنت كاتب مراسلات رسمية في جامعة عراقية. اكتب الكتب الرسمية بالأسلوب الإداري العراقي المعتاد (الجهة، العدد، التاريخ، إلى، م/، المتن، التوقيع، نسخة منه إلى). نص عادي بدون Markdown.",
      text: `اكتب ${type} رسمياً للطالب «${st.name}» (${S.settings.program}، ${S.settings.term}) صادراً عن «${from}».
الوقائع (لا تغيّرها): ${su ? `المادة: ${su.name}، المدرس: ${su.teacher}` : "كل المواد"}؛ عدد الغيابات ${a.abs} من ${a.req} محاضرة؛ ساعات الغياب ${a.habs} من ${a.htot} ساعة؛ النسبة ${pct(a.rate)}؛ حد الإنذار الأولي ${W().w1}% وحد الإنذار الثاني ${W().w2}% من ساعات الفصل؛ تواريخ الغياب: ${a.dates.map(fmtDate).join("، ") || "لا يوجد"}؛ تاريخ اليوم ${fmtDate(today())}.
اترك خانة العدد فارغة.`
    })).trim();
  } catch (err) { toast("تعذّرت الصياغة: " + err.message); }
  render();
}

/* ===== Excel export ===== */
function buildXlsx() {
  const wb = XLSX.utils.book_new();
  const sum = [["المادة", ...S.students.flatMap(s => [s.name + " — غياب/محاضرات أُعطيت", s.name + " — ساعات الغياب/ساعات الفصل"])]];
  for (const su of S.subjects) sum.push([su.name, ...S.students.flatMap(st => { if ((st.exempt || []).includes(su.id)) return ["مستوفية", "مستوفية"]; const b = stats(st.id, su.id); return [`${b.abs} / ${b.req}`, `${b.habs} / ${b.htot}`]; })]);
  sum.push(["المجموع", ...S.students.flatMap(st => { const a = stats(st.id); return [`${a.abs} / ${a.req}`, `${a.habs} / ${a.htot}`]; })]);
  const ws = XLSX.utils.aoa_to_sheet(sum); ws["!cols"] = [{ wch: 24 }, ...S.students.flatMap(() => [{ wch: 18 }, { wch: 22 }])]; ws["!views"] = [{ RTL: true }];
  XLSX.utils.book_append_sheet(wb, ws, "الملخص");
  const used = new Set(["الملخص"]);
  for (const su of S.subjects) {
    const rows = [["التاريخ", "اليوم", ...S.students.map(s => s.name), "ملاحظات"]];
    S.sessions.filter(s => s.subject === su.id).sort((a, b) => a.date.localeCompare(b.date)).forEach(se => rows.push([fmtDate(se.date), dayName(se.date), ...S.students.map(st => STATUS[status(se, st)]), se.note || ""]));
    const w = XLSX.utils.aoa_to_sheet(rows); w["!cols"] = [{ wch: 12 }, { wch: 10 }, ...S.students.map(() => ({ wch: 22 })), { wch: 30 }]; w["!views"] = [{ RTL: true }];
    let name = su.name.replace(/[\\\/\?\*\[\]:]/g, " ").slice(0, 31); while (used.has(name)) name = name.slice(0, 29) + "_" + used.size; used.add(name);
    XLSX.utils.book_append_sheet(wb, w, name);
  }
  return XLSX.write(wb, { type: "array", bookType: "xlsx" });
}

/* ===== Drive sync ===== */
let syncing = false, syncTimer = null;
function syncBadge(state, text) {
  const b = $("#syncBtn"), on = Drive.connected();
  state = state || (on ? "" : "off");
  text = text || (on ? "متزامن مع Drive" : (C.drive.clientId ? "اضغط للمزامنة مع Drive" : "محفوظ على هذا الجهاز"));
  b.className = "save " + state; b.querySelector("span").textContent = text;
}
function scheduleSync() { if (!Drive.connected()) return; syncBadge("busy", "سيُرفع إلى Drive…"); clearTimeout(syncTimer); syncTimer = setTimeout(() => syncNow(false), 3000); }
async function syncNow(interactive) {
  if (syncing) return; syncing = true; syncBadge("busy", "جارٍ المزامنة…");
  try {
    await Drive.signIn(C.drive.clientId, interactive);
    const remote = await Drive.pull(C.drive);
    const localChanged = !C.lastSync || (S.updatedAt || "") > C.lastSync;
    const remoteNewer = remote && (remote.updatedAt || "") > (S.updatedAt || "");
    if (remote && remoteNewer && !localChanged) { S = remote; store(KEY, S); toast("سُحب أحدث سجل من Drive"); }
    else if (remote && remoteNewer && localChanged) {
      // both sides changed since last sync: keep the newer copy, back up the other inside Drive's version history
      if (interactive) { await Drive.push(C.drive, S); toast("رُفعت نسختك. النسخة السابقة محفوظة في سجل إصدارات الملف على Drive"); }
      else { S = remote; store(KEY, S); toast("سُحب سجل أحدث من جهاز آخر"); }
    }
    else if (!remote || (S.updatedAt || "") !== (remote.updatedAt || "")) await Drive.push(C.drive, S);
    C.lastSync = new Date().toISOString(); saveCfg(); render();
  } catch (err) { syncBadge("err", "تعذّرت المزامنة"); if (interactive) toast(err.message); }
  syncing = false;
}
$("#syncBtn").onclick = () => { if (!C.drive.clientId) { location.hash = "setup"; toast("أضف Client ID أولاً — الخطوات في الأسفل"); return; } syncNow(true); };
async function sheetToDrive() {
  if (!C.drive.clientId) { location.hash = "setup"; toast("اربط Google Drive أولاً من الإعدادات"); return; }
  try { await Drive.signIn(C.drive.clientId, true); const link = await Drive.uploadSheet(C.drive, `حضور الماجستير ${today()}`, buildXlsx()); saveCfg(); toast("رُفعت نسخة Google Sheets إلى مجلد «" + Drive.FOLDER + "»"); if (link) window.open(link, "_blank"); }
  catch (err) { toast(err.message); }
}

/* ===== install + boot ===== */
let installEvt = null;
window.addEventListener("beforeinstallprompt", e => { e.preventDefault(); installEvt = e; $("#installBtn").hidden = false; });
$("#installBtn").onclick = async () => { if (!installEvt) return; installEvt.prompt(); await installEvt.userChoice; installEvt = null; $("#installBtn").hidden = true; };
window.addEventListener("appinstalled", () => toast("ثُبّت التطبيق"));
const APP_VERSION = "v6";
// Updates install themselves: a new service worker takes over and the page reloads once.
if ("serviceWorker" in navigator && location.protocol === "https:") {
  const hadController = !!navigator.serviceWorker.controller; let reloaded = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => { if (hadController && !reloaded) { reloaded = true; location.reload(); } });
  navigator.serviceWorker.register("sw.js").then(r => { window._swReg = r; setInterval(() => r.update().catch(() => {}), 30 * 60e3); }).catch(() => {});
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") window._swReg?.update().catch(() => {}); });
}
window.addEventListener("hashchange", render);
store(KEY, S);
render();
