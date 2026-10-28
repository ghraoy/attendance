/* ===== helpers ===== */
const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const STATUS = { P: "حاضر", A: "غائب", X: "مجاز", N: "غير مطالب", E: "مستوفية" };
const fmtDate = d => { if (!d) return "—"; const [y, m, dd] = d.split("-"); return `${+dd}/${+m}/${y}`; };
const dayName = d => Core.WEEKDAYS[Core.dow(d)] || "";
const today = () => { const t = new Date(); return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`; };
const uid = Core.uid, clone = Core.clone;
const short = n => { const w = String(n).split(" "); return /^(عبد|ابو|أبو)$/.test(w[0]) && w[1] ? w[0] + " " + w[1] : w[0]; };
const pct = x => (x * 100).toFixed(1).replace(/\.0$/, "") + "%";
const num = v => { const n = +v; return isNaN(n) ? 0 : n; };

/* ===== storage ===== */
const KEY = "mtu-attendance-v2", OLDKEY = "mtu-attendance-v1", CFG = "mtu-config-v1";
const load = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } };
const store = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } };

let DB = load(KEY, null);
if (!DB || !(DB.registers || []).length) DB = Core.migrate(load(OLDKEY, null)) || DB || Core.emptyDB(); // carry the v1 masters register over
let C = load(CFG, { provider: "gemini", keys: {}, models: {}, drive: { clientId: "" }, lastSync: null });
C.keys ||= {}; C.models ||= {}; C.drive ||= { clientId: "" };
const DEFAULT_CLIENT_ID = "629045560897-1ritqqkb0qcul2km6ifik0mrq315ru3s.apps.googleusercontent.com";
if (!C.drive.clientId && location.hostname === "ghraoy.github.io") C.drive.clientId = DEFAULT_CLIENT_ID;
const R = () => DB.registers.find(r => r.id === DB.current) || DB.registers[0] || null;
const ui = { subject: "", student: "", onlyAbs: false, preview: null, answer: "", asking: false, letter: "", photoBusy: false, modelLists: {}, day: today(), wiz: null, tplPreview: null };

function saveCfg() { store(CFG, C); }
function commit(msg) {
  DB.updatedAt = new Date().toISOString();
  if (C.drive.clientId && !C.pendingSync) { C.pendingSync = true; saveCfg(); } // queued for upload until a sync succeeds
  if (!store(KEY, DB)) toast("تعذّر الحفظ على هذا الجهاز — مساحة المتصفح ممتلئة أو محجوبة");
  render(); if (msg) toast(msg);
  scheduleSync();
}

/* ===== shortcuts over Core for the current register ===== */
const subj = id => R().subjects.find(s => s.id === id);
const status = (se, stu) => Core.status(R(), se, stu);
const stats = (stuId, subId) => Core.stats(R(), stuId, subId);
const level = rate => Core.level(R(), rate);
const LV = { w1: "w", w2: "d", w3: "d" };   // css class per level
const levelText = lv => Core.LEVEL_TEXT[lv] || "";

/* ===== small UI helpers ===== */
function toast(t) { const el = $("#toast"); el.textContent = t; el.hidden = false; clearTimeout(el._t); el._t = setTimeout(() => el.hidden = true, 2800); }
function twoStep(btn, fn, label = "حذف") { if (btn.dataset.arm) { fn(); return; } btn.dataset.arm = 1; const old = btn.textContent; btn.textContent = "تأكيد " + label; setTimeout(() => { if (btn.isConnected) { delete btn.dataset.arm; btn.textContent = old; } }, 3000); }
async function saveFile(name, data, type) { try { data = await data; } catch (e) { toast("تعذّر إنشاء الملف: " + e.message); return; } const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([data], { type })); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000); }
async function copy(text) { try { await navigator.clipboard.writeText(text); toast("نُسخ النص"); } catch (e) { toast("تعذّر النسخ — حدّد النص وانسخه يدوياً"); } }
const aiReady = () => !!(C.keys[C.provider] && C.models[C.provider]);
const opt = (v, cur, label) => `<option value="${esc(v)}" ${String(v) === String(cur) ? "selected" : ""}>${esc(label ?? v)}</option>`;
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/* ===== routing ===== */
const VIEWS = ["overview", "today", "log", "timetable", "import", "ai", "setup", "new"];
const view = () => { const h = location.hash.slice(1); return VIEWS.includes(h) ? h : "overview"; };
function render() {
  const v = !R() && !["new", "import", "setup"].includes(view()) ? "welcome" : view();
  $$("nav a").forEach(a => a.setAttribute("aria-current", a.getAttribute("href") === "#" + v ? "page" : "false"));
  // register switcher
  const sw = $("#regSel");
  sw.innerHTML = DB.registers.map(r => opt(r.id, DB.current, Core.registerName(r))).join("") + `<option value="__new">+ سجل جديد…</option>`;
  sw.hidden = !DB.registers.length;
  $("#termLine").textContent = R() ? [R().meta.college, R().meta.year && "العام " + R().meta.year].filter(Boolean).join(" · ") : "سجل الحضور والغياب الجامعي";
  $("#addBtn").hidden = !R();
  const fn = { welcome, overview, today: todayView, log: logView, timetable: timetableView, import: importView, ai: aiView, setup: setupView, new: wizardView }[v];
  $("#main").innerHTML = fn();
  bind(v); syncBadge();
}

/* ===== welcome (no register yet) ===== */
function welcome() {
  return `<div class="panel" style="text-align:center;justify-items:center;padding:36px 20px">
    <h2 style="margin:0;font-size:1.3rem">سجل الحضور والغياب الجامعي</h2>
    <p class="muted" style="margin:0;max-width:56ch">لأي قسم ولأي مرحلة: دراسات أولية (صباحي أو مسائي) أو دراسات عليا. تُعرّف المواد وجدولها الأسبوعي والتقويم وقواعد الإنذار، ثم تسجّل الحضور فيحسب البرنامج الساعات والنسب والإنذارات.</p>
    <div class="row" style="justify-content:center"><a class="btn primary" href="#new">إنشاء سجل جديد خطوة بخطوة</a><a class="btn" href="#import">استيراد نموذج Excel معبّأ</a></div>
    <p class="muted small" style="margin:0">إذا كان لديك سجل على جهاز آخر مربوط بـ Google Drive، اضغط زر المزامنة في الأعلى.</p></div>`;
}

/* ===== overview ===== */
function overview() {
  const r = R(), rules = r.rules;
  if (!r.students.length || !r.subjects.length) return `<div class="empty">السجل «${esc(Core.registerName(r))}» بلا ${!r.subjects.length ? "مواد" : "طلبة"} بعد. أكملها من <a href="#setup">الإعدادات</a> أو استورد ملف Excel من <a href="#import">إدخال البيانات</a>.</div>`;
  const scale = Math.max(rules.warn3 || rules.warn2, rules.warn2) * 1.4;
  const basisTxt = rules.basis === "subject" ? "لكل مادة على حدة" : "من مجموع ساعات كل المواد";
  const cards = r.students.map(st => {
    const a = stats(st.id), gov = Core.studentLevel(r, st), lvc = LV[gov.lv] || "";
    return `<div class="card">
      <div class="name">${esc(st.name)}</div>
      <div class="start">يُحسب الدوام من ${fmtDate(st.start || r.calendar.start)}${(st.exempt || []).length ? " · مستوفٍ " + st.exempt.length + " مادة" : ""}</div>
      <div class="row" style="align-items:baseline"><span class="big">${a.habs}</span><span class="muted">ساعة غياب من ${a.htot} ساعة مقررة</span></div>
      <div class="bar" title="الإنذارات: ${rules.warn1}% · ${rules.warn2}%${rules.warn3 ? " · " + rules.warn3 + "%" : ""}"><span class="${LV[level(a.rate)] || ""}" style="width:${Math.min(100, a.rate * 100 / scale * 100)}%"></span>${[rules.warn1, rules.warn2, rules.warn3].filter(Boolean).map(w => `<b style="right:${w / scale * 100}%"></b>`).join("")}</div>
      <div class="kv"><span>النسبة الكلية</span><strong class="${LV[level(a.rate)] || ""}">${pct(a.rate)}</strong></div>
      <div class="kv"><span>الغيابات</span><strong>${a.abs} من ${a.req} محاضرة أُعطيت${a.exc ? ` · ${a.exc} بعذر` : ""}</strong></div>
      ${rules.basis === "total" ? `<div class="kv"><span>باقي حتى الإنذار الأولي</span><strong class="${a.left1 ? "" : "w"}">${a.left1 ? a.left1 + " ساعة" : "تجاوزه"}</strong></div>
      <div class="kv"><span>باقي حتى الإنذار الثاني</span><strong class="${a.left2 ? "" : "d"}">${a.left2 ? a.left2 + " ساعة" : "تجاوزه"}</strong></div>` : gov.subject ? `<div class="kv"><span>أعلى مادة</span><span class="${lvc}">${esc(gov.subject.name)} · ${pct(gov.rate)}</span></div>` : ""}
      ${gov.lv ? `<div class="banner ${gov.lv === "w1" ? "" : "err"}" style="padding:4px 10px">${levelText(gov.lv)}${gov.subject ? " في " + esc(gov.subject.name) : ""}</div>` : ""}
    </div>`;
  }).join("");
  const head = r.students.map(s => `<th>${esc(short(s.name))}</th>`).join("");
  const rows = r.subjects.map(su => `<tr><td class="l"><strong>${esc(su.name)}</strong><br><small class="muted">${Core.subjectWeeklyHours(su)} س/أسبوع</small></td>${r.students.map(st => {
    if ((st.exempt || []).includes(su.id)) return `<td><span class="pill p-E">مستوفية</span></td>`;
    const b = stats(st.id, su.id); if (!b.htot && !b.req) return `<td class="muted">—</td>`;
    const cls = rules.basis === "subject" ? (LV[level(b.rate)] || "") : "";
    return `<td><span class="cell ${cls}"><strong>${b.abs} / ${b.req}</strong><small>${b.habs} من ${b.htot} س · ${pct(b.rate)}</small></span></td>`;
  }).join("")}</tr>`).join("");
  const foot = r.students.map(st => { const a = stats(st.id), lv = level(a.rate); return `<td><span class="cell ${rules.basis === "total" ? LV[lv] || "" : ""}"><strong>${pct(a.rate)}</strong><small>${a.habs} / ${a.htot} س</small></span></td>`; }).join("");
  return `<section><div class="row"><h2>${esc(Core.registerName(r))}</h2><span class="spacer"></span><span class="muted small">النسبة ${basisTxt} · إنذار أولي ${rules.warn1}% · ثانٍ ${rules.warn2}%${rules.warn3 ? " · حرمان " + rules.warn3 + "%" : ""}</span></div><div class="cards">${cards}</div></section>
  <section><div class="row"><h2>الغياب حسب المادة</h2><span class="spacer"></span><button class="btn small" id="xlsx" type="button">تنزيل Excel</button><button class="btn small" id="xlsxDrive" type="button">رفع نسخة إلى Drive</button></div>
  <div class="scroll"><table><thead><tr><th class="l">المادة</th>${head}</tr></thead><tbody>${rows}</tbody><tfoot><tr><td class="l">النسبة الكلية (كل المواد)</td>${foot}</tr></tfoot></table></div>
  <p class="muted small" style="margin:8px 0 0">في كل خانة: الغيابات / المحاضرات المُعطاة، وتحتها ساعات الغياب من ساعات الفصل المقررة (${Core.termWeeks(r)} أسبوعاً${r.calendar.holidays.length ? " مع خصم العطل" : ""}). ${rules.basis === "subject" ? "الإنذارات هنا تُحسب لكل مادة على حدة." : "الإنذارات تُحسب من النسبة الكلية."}</p></section>`;
}

/* ===== today: quick marking from the timetable ===== */
function todayView() {
  const r = R(), d = ui.day, sched = Core.scheduledOn(r, d);
  const held = r.sessions.filter(s => s.date === d);
  const holiday = Core.isHoliday(r, d), duty = Core.isDuty(r, d);
  const rowsFor = (su, slot) => {
    const se = held.find(s => s.subject === su.id);
    const marks = r.students.map(st => {
      const x = se ? status(se, st) : (st.start && d < st.start ? "N" : (st.exempt || []).includes(su.id) ? "E" : "P");
      const fixed = x === "N" || x === "E";
      return `<div class="attrow"><span>${esc(st.name)}</span>${fixed ? `<span class="pill p-${x}">${STATUS[x]}</span>` :
        `<span class="seg">${["P", "A", "X"].map(v => `<button type="button" class="${v}" data-q="${su.id}" data-st="${st.id}" data-v="${v}" aria-pressed="${x === v}">${STATUS[v]}</button>`).join("")}</span>`}</div>`;
    }).join("");
    return `<div class="panel"><div class="row"><h3>${esc(su.name)}</h3><span class="muted small">${slot ? slot.from + " – " + slot.to : ""} · ${esc(su.teacher)}</span><span class="spacer"></span>${se ? `<span class="pill p-P">مُسجّلة</span>` : `<span class="pill p-N">لم تُسجّل بعد</span>`}</div>
      <div class="att">${marks}</div>
      <div class="row">${se ? `<button class="btn small" type="button" data-open="${se.id}">تفاصيل وملاحظة</button>` : `<button class="btn small primary" type="button" data-allp="${su.id}">الجميع حاضر — حفظ</button>`}</div></div>`;
  };
  const extra = held.filter(s => !sched.some(x => x.subject.id === s.subject));
  return `<section><h2>محاضرات اليوم</h2>
    <div class="filters"><label class="field">التاريخ<input type="date" id="dayPick" value="${d}"></label>
    <button class="btn small" id="dPrev" type="button">اليوم السابق</button><button class="btn small" id="dToday" type="button">اليوم</button><button class="btn small" id="dNext" type="button">اليوم التالي</button>
    <span class="spacer"></span><span class="muted small" style="padding-bottom:8px">${dayName(d)} ${fmtDate(d)}</span></div>
    ${holiday ? `<div class="banner" style="margin-top:10px">هذا اليوم ضمن العطل المسجّلة في التقويم.</div>` : !duty ? `<div class="banner" style="margin-top:10px">${dayName(d)} ليس يوم دوام إلزامي: الغياب فيه لا يُحسب، والحضور يُحسب.</div>` : ""}</section>
    ${sched.length || extra.length ? `<div class="panels">${sched.map(x => rowsFor(x.subject, x.slot)).join("")}${extra.map(se => rowsFor(subj(se.subject) || { id: se.subject, name: "؟", teacher: "" }, null)).join("")}</div>`
      : `<div class="empty">لا محاضرات في جدول ${dayName(d)}. ${r.subjects.some(s => (s.slots || []).length) ? "" : "أضف أيام وأوقات المحاضرات لكل مادة من «الإعدادات» ليظهر الجدول هنا تلقائياً. "}يمكنك إضافة محاضرة بزر «+ محاضرة».</div>`}`;
}
function markQuick(subId, stuId, v) {
  const r = R(), d = ui.day;
  let se = r.sessions.find(s => s.date === d && s.subject === subId);
  if (!se) { se = { id: "x" + uid(), date: d, subject: subId, note: "", absent: [], excused: [] }; r.sessions.push(se); }
  se.absent = (se.absent || []).filter(x => x !== stuId); se.excused = (se.excused || []).filter(x => x !== stuId);
  if (v === "A") se.absent.push(stuId); if (v === "X") se.excused.push(stuId);
  commit();
}

/* ===== timetable ===== */
function timetableView() {
  const r = R(), days = r.calendar.dutyDays.length ? [...new Set([...r.calendar.dutyDays, ...r.subjects.flatMap(s => (s.slots || []).map(x => +x.day))])].sort() : [0, 1, 2, 3, 4];
  const cols = days.map(d => {
    const items = r.subjects.flatMap(su => (su.slots || []).filter(s => +s.day === d).map(s => ({ su, s }))).sort((a, b) => Core.toMin(a.s.from) - Core.toMin(b.s.from));
    return `<div class="panel" style="align-content:start"><h3>${Core.WEEKDAYS[d]}${r.calendar.dutyDays.includes(d) ? "" : ` <span class="pill p-N">ليس يوم دوام</span>`}</h3>
      ${items.length ? items.map(({ su, s }) => `<div class="attrow" style="display:grid;gap:2px"><strong>${esc(su.name)}</strong><span class="muted small">${s.from} – ${s.to} · ${Core.slotHours(s)} س · ${esc(su.teacher)}${su.kind ? " · " + esc(su.kind) : ""}</span></div>`).join("") : `<span class="muted small">لا محاضرات</span>`}</div>`;
  }).join("");
  const c = r.calendar;
  return `<section><div class="row"><h2>الجدول الأسبوعي</h2><span class="spacer"></span><a class="btn small" href="#setup">تعديل الجدول</a></div>
    <p class="muted small" style="margin:0">الفصل ${esc(r.meta.semester || "")}: ${fmtDate(c.start)} – ${fmtDate(c.end)} · ${Core.termWeeks(r)} أسبوعاً · أيام الدوام: ${c.dutyDays.map(d => Core.WEEKDAYS[d]).join("، ") || "—"}${c.holidays.length ? " · العطل: " + c.holidays.map(h => (h.name || "عطلة") + " (" + fmtDate(h.from) + (h.to && h.to !== h.from ? "–" + fmtDate(h.to) : "") + ")").join("، ") : ""}</p></section>
    <div class="panels" style="grid-template-columns:repeat(auto-fit,minmax(200px,1fr))">${cols}</div>
    <section class="scroll"><table><thead><tr><th class="l">المادة</th><th>التدريسي</th><th>النوع</th><th>ساعات/أسبوع</th><th>ساعات الفصل (للطالب المنتظم)</th></tr></thead><tbody>${r.subjects.map(su => `<tr><td class="l">${esc(su.name)}${su.code ? ` <small class="muted">${esc(su.code)}</small>` : ""}</td><td>${esc(su.teacher)}</td><td>${esc(su.kind || "")}</td><td>${Core.subjectWeeklyHours(su)}</td><td>${Core.plannedHours(r, { start: c.start, exempt: [] }, su)}</td></tr>`).join("")}</tbody></table></section>`;
}

/* ===== archive ===== */
function logView() {
  const r = R();
  const subOpts = `<option value="">كل المواد</option>` + r.subjects.map(s => opt(s.id, ui.subject, s.name)).join("");
  const stuOpts = `<option value="">كل الطلبة</option>` + r.students.map(s => opt(s.id, ui.student, s.name)).join("");
  let list = [...r.sessions].sort((a, b) => b.date.localeCompare(a.date) || (subj(a.subject)?.name || "").localeCompare(subj(b.subject)?.name || ""));
  if (ui.subject) list = list.filter(s => s.subject === ui.subject);
  if (ui.onlyAbs) list = list.filter(se => r.students.some(st => (!ui.student || st.id === ui.student) && ["A", "X"].includes(status(se, st))));
  const groups = {}; list.forEach(s => (groups[s.date] = groups[s.date] || []).push(s));
  const body = Object.keys(groups).map(d => `<div class="day"><div class="dayhead">${fmtDate(d)} <span>${dayName(d)}${Core.isDuty(r, d) ? "" : " · ليس يوم دوام"}</span></div>${groups[d].map(se => {
    const su = subj(se.subject) || { name: "مادة محذوفة", teacher: "" };
    const chips = r.students.filter(st => !ui.student || st.id === ui.student).map(st => { const x = status(se, st); return `<span class="pill p-${x}">${esc(short(st.name))}: ${STATUS[x]}</span>`; }).join("");
    return `<button class="sess" type="button" data-id="${se.id}"><div><div class="t">${esc(su.name)}</div><div class="m">${esc(su.teacher)}${su.teacher ? " · " : ""}${Core.sessionHours(r, se)} ساعة</div></div><span class="m">تعديل</span><div class="chips">${chips}</div>${se.note ? `<div class="note" style="grid-column:1/-1">${esc(se.note)}</div>` : ""}</button>`;
  }).join("")}</div>`).join("");
  return `<section><h2>أرشيف المحاضرات</h2>
   <div class="filters"><label class="field">المادة<select id="fSub">${subOpts}</select></label>
   <label class="field">الطالب<select id="fStu">${stuOpts}</select></label>
   <label class="check" style="padding-bottom:8px"><input type="checkbox" id="fAbs" ${ui.onlyAbs ? "checked" : ""}> الغيابات فقط</label>
   <span class="spacer"></span><span class="muted small" style="padding-bottom:8px">${list.length} محاضرة</span></div></section>
   <section style="display:grid;gap:18px">${body || `<div class="empty">لا محاضرات بهذا التصفية. سجّل من «اليوم» أو بزر «+ محاضرة».</div>`}</section>`;
}

/* ===== import ===== */
function importView() {
  const pv = ui.preview, tp = ui.tplPreview;
  const tpl = tp ? `<div class="panel" style="border-color:var(--accent)"><h3>سجل جديد من نموذج Excel</h3>
    <div class="small"><b>${esc(Core.registerName(tp.R))}</b><br>${esc(tp.R.meta.university)} · ${esc(tp.R.meta.college)}<br>
    ${tp.R.subjects.length} مادة · ${tp.R.students.length} طالب · ${tp.plan.created} محاضرة مسجّلة · ${tp.plan.absences} غياب · الفصل ${fmtDate(tp.R.calendar.start)} – ${fmtDate(tp.R.calendar.end)}</div>
    ${tp.R.subjects.filter(s => !(s.slots || []).length).length ? `<div class="banner">${tp.R.subjects.filter(s => !(s.slots || []).length).length} مادة بلا جدول أسبوعي؛ ستُحسب ساعاتها من «ساعات المحاضرة × محاضرات الأسبوع».</div>` : ""}
    <div class="row"><button class="btn primary" id="tplApply" type="button">إنشاء السجل</button>${R() ? `<button class="btn" id="tplMerge" type="button">دمج الحضور فقط في السجل الحالي</button>` : ""}<button class="btn" id="tplCancel" type="button">إلغاء</button></div></div>` : "";
  const preview = pv ? `<div class="panel" style="border-color:var(--accent)"><h3>معاينة قبل الحفظ</h3>
    <div class="small">${pv.created} محاضرة جديدة · ${pv.updated} محاضرة موجودة ستُحدَّث · ${pv.absences} غياب</div>
    ${pv.newSubjects.length ? `<div class="banner">مواد جديدة ستُضاف (راجع ساعاتها وجدولها من الإعدادات): ${pv.newSubjects.map(s => esc(s.name)).join("، ")}</div>` : ""}
    ${pv.newStudents.length ? `<div class="banner">طلبة جدد: ${pv.newStudents.map(s => esc(s.name) + (s.start && s.start < "9999" ? " (من " + fmtDate(s.start) + ")" : "")).join("، ")}. إن كان الاسم مكتوباً بشكل مختلف لطالب موجود، صحّحه في الملف وأعد الاستيراد.</div>` : ""}
    ${(pv.newExempt || []).length ? `<div class="banner ok">مستوفون: ${pv.newExempt.map(esc).join("، ")}</div>` : ""}
    ${!pv.created && !pv.updated ? `<div class="banner err">لم أجد محاضرات في الملف. تأكد أن فيه عموداً باسم «التاريخ».</div>` : ""}
    <div class="row"><button class="btn primary" id="applyImp" type="button" ${!pv.created && !pv.updated ? "disabled" : ""}>حفظ في السجل</button><button class="btn" id="cancelImp" type="button">إلغاء</button></div></div>` : "";
  return `<section><h2>إدخال البيانات</h2><p class="muted small" style="margin:0">طرق الإدخال: «اليوم» حسب الجدول، أو زر «+ محاضرة»، أو نموذج Excel، أو لصق جدول، أو صورة ورقة الحضور.</p></section>
  ${tpl}${preview}
  <div class="panels">
    <div class="panel"><h3>نموذج Excel الكامل</h3>
      <p class="muted small" style="margin:0">نموذج موحّد لأي قسم ومرحلة: الإعدادات، العطل، المواد وجدولها، الطلبة، و<b>كشف حضور لكل مادة</b> بتواريخ المحاضرات تُولَّد تلقائياً من الجدول (يُكتب فيه غ / م / ح)، مع ملخص يحسب الساعات والنسب والإنذارات. نزّله فارغاً واملأه، أو صدّر السجل الحالي بنفس الشكل.</p>
      <div class="row"><button class="btn" id="tplBlank" type="button">تنزيل النموذج فارغاً</button>${R() ? `<button class="btn" id="tplExport" type="button">تصدير السجل الحالي بالنموذج</button>` : ""}</div>
      <label class="btn primary" style="justify-self:start">استيراد ملف Excel<input type="file" id="xfile" accept=".xlsx,.xls,.csv" hidden></label>
      <p class="muted small" style="margin:0">يقبل النموذج الموحّد (يُنشئ سجلاً كاملاً)، أو أي ملف فيه أعمدة: التاريخ، المادة، الطالب، الحالة (يُضاف للسجل الحالي).</p>
    </div>
    ${R() ? `<div class="panel"><h3>صورة ورقة الحضور</h3>
      <p class="muted small" style="margin:0">صوّر الورقة الموقّعة، فيقرأ ${esc(AI.NAMES[C.provider] || "الذكاء الاصطناعي")} التاريخ والمادة والأسماء، ثم تراجعها قبل الحفظ.</p>
      ${aiReady() ? "" : `<div class="banner">أضف مفتاح Gemini أو ChatGPT أو Claude من «الإعدادات» لتفعيل هذه الطريقة.</div>`}
      <label class="btn primary" style="justify-self:start;${aiReady() && !ui.photoBusy ? "" : "opacity:.5;pointer-events:none"}">${ui.photoBusy ? "جارٍ قراءة الورقة…" : "التقاط أو اختيار صورة"}<input type="file" id="photo" accept="image/*" capture="environment" hidden></label></div>
    <div class="panel"><h3>لصق جدول</h3>
      <label class="field">انسخ الخلايا من Excel أو Google Sheets والصقها مع سطر العناوين<textarea id="paste" placeholder="التاريخ	المادة	الطالب	الحالة&#10;2026-11-01	الوبائيات	محمد علي حسن	غائب"></textarea></label>
      <label class="field">المادة إن لم يكن في الجدول عمود «المادة»<select id="pasteSub">${R().subjects.map(s => opt(s.name, "", s.name)).join("")}</select></label>
      <button class="btn primary" id="pasteGo" type="button" style="justify-self:start">قراءة الجدول</button></div>` : ""}
    <div class="panel"><h3>نسخة احتياطية</h3>
      <p class="muted small" style="margin:0">ملف واحد فيه كل السجلات، تحفظه أو تنقله لجهاز آخر.</p>
      <div class="row"><button class="btn" id="bkSave" type="button">حفظ نسخة</button><label class="btn">استرجاع نسخة<input type="file" id="bkLoad" accept=".json" hidden></label></div></div>
  </div>`;
}

/* ===== AI ===== */
function aiView() {
  const r = R();
  const quick = ["من عليه إنذار أو اقترب منه؟", "لخّص غيابات كل طالب حسب المادة", "ما المواد التي فيها أكثر غياب؟", "كم ساعة يستطيع كل طالب أن يغيب بعد قبل الإنذار؟"];
  return `<section><h2>المساعد</h2><p class="muted small" style="margin:0">${aiReady() ? `يعمل بـ ${esc(AI.NAMES[C.provider])} (${esc(C.models[C.provider])}). تُرسل إليه أرقام السجل الحالي مع سؤالك فقط.` : "أضف مفتاح Gemini أو ChatGPT أو Claude من «الإعدادات» لتفعيل الأسئلة والصياغة. قالب كتاب الإنذار يعمل بدون مفتاح."}</p></section>
  <div class="panels">
    <div class="panel"><h3>اسأل عن السجل</h3>
      <div class="row">${quick.map(q => `<button class="chipbtn" type="button" data-ask="${esc(q)}">${esc(q)}</button>`).join("")}</div>
      <label class="field">سؤالك<textarea id="q" style="min-height:80px" placeholder="مثلاً: كم غياب على الطالب في مادة كذا ومتى؟"></textarea></label>
      <button class="btn primary" id="askGo" type="button" style="justify-self:start" ${aiReady() && !ui.asking ? "" : "disabled"}>${ui.asking ? "جارٍ التفكير…" : "اسأل"}</button>
      ${ui.answer ? `<div class="answer">${esc(ui.answer)}</div><div class="row"><button class="btn small" id="ansCopy" type="button">نسخ</button></div>` : ""}</div>
    <div class="panel"><h3>كتاب إنذار</h3>
      <div class="grid2">
        <label class="field">الطالب<select id="lStu">${r.students.map(s => opt(s.id, "", s.name)).join("")}</select></label>
        <label class="field">المادة<select id="lSub"><option value="">كل المواد</option>${r.subjects.map(s => opt(s.id, "", s.name)).join("")}</select></label>
        <label class="field">النوع<select id="lType"><option>إنذار أولي</option><option>إنذار ثانٍ</option><option>إنذار نهائي</option><option>إشعار حرمان</option></select></label>
        <label class="field">جهة الإصدار<input id="lFrom" value="${esc(r.meta.letterFrom || ("رئاسة قسم " + (r.meta.department || "")))}"></label>
      </div>
      <div class="row"><button class="btn" id="lTpl" type="button">قالب جاهز</button><button class="btn primary" id="lAi" type="button" ${aiReady() ? "" : "disabled"}>صياغة بالذكاء الاصطناعي</button></div>
      ${ui.letter ? `<label class="field">الكتاب (يمكنك تعديله)<textarea id="letter" style="min-height:280px">${esc(ui.letter)}</textarea></label>
      <div class="row"><button class="btn small" id="lCopy" type="button">نسخ</button><button class="btn small" id="lPrint" type="button">طباعة</button><button class="btn small" id="lDoc" type="button">حفظ Word</button></div>` : ""}</div>
  </div>`;
}

/* ===== shared form pieces (setup + wizard) ===== */
const LEVELS = ["أولية", "عليا"], SHIFTS = ["صباحي", "مسائي"], DEGREES = ["دبلوم عالي", "ماجستير", "دكتوراه"], SYSTEMS = ["فصلي", "سنوي", "بولونيا"], SEMESTERS = ["الأول", "الثاني", "الصيفي", "سنوي"];
const UG_STAGES = ["الأولى", "الثانية", "الثالثة", "الرابعة", "الخامسة", "السادسة"], PG_STAGES = ["السنة التحضيرية (الكورسات)", "مرحلة البحث"];
function metaForm(m, p = "m") {
  const ug = m.level === "أولية";
  return `<div class="grid2">
    <label class="field">الجامعة<input data-${p}="university" value="${esc(m.university)}" placeholder="مثلاً: الجامعة التقنية الوسطى"></label>
    <label class="field">الكلية / المعهد<input data-${p}="college" value="${esc(m.college)}"></label>
    <label class="field">القسم<input data-${p}="department" value="${esc(m.department)}"></label>
    <label class="field">العام الدراسي<input data-${p}="year" value="${esc(m.year)}" placeholder="2026-2027"></label>
    <label class="field">نوع الدراسة<select data-${p}="level">${LEVELS.map(x => opt(x, m.level, "الدراسات " + x)).join("")}</select></label>
    ${ug ? `<label class="field">الدوام<select data-${p}="shift">${SHIFTS.map(x => opt(x, m.shift)).join("")}</select></label>` : `<label class="field">الشهادة<select data-${p}="degree">${DEGREES.map(x => opt(x, m.degree)).join("")}</select></label>`}
    <label class="field">${ug ? "المرحلة" : "المرحلة / السنة"}<input data-${p}="stage" list="stages-${p}" value="${esc(m.stage)}"><datalist id="stages-${p}">${(ug ? UG_STAGES : PG_STAGES).map(x => `<option value="${x}">`).join("")}</datalist></label>
    <label class="field">الشعبة (اختياري)<input data-${p}="section" value="${esc(m.section)}" placeholder="أ، ب …"></label>
    <label class="field">النظام الدراسي<select data-${p}="system">${SYSTEMS.map(x => opt(x, m.system)).join("")}</select></label>
    <label class="field">الفصل الدراسي<select data-${p}="semester">${SEMESTERS.map(x => opt(x, m.semester)).join("")}</select></label>
  </div>`;
}
function calendarForm(c, p = "c") {
  return `<div class="grid2">
    <label class="field">بداية الفصل<input type="date" data-${p}="start" value="${c.start}"></label>
    <label class="field">نهاية الفصل (قبل الامتحانات)<input type="date" data-${p}="end" value="${c.end}"></label>
    <label class="field">عدد الأسابيع (اتركه 0 ليُحسب من التاريخين)<input type="number" min="0" max="52" data-${p}="weeks" value="${+c.weeks || 0}"></label>
    <div class="field">أيام الدوام الإلزامي (الغياب في غيرها لا يُحسب)<div class="ex">${Core.WEEKDAYS.map((d, i) => `<label><input type="checkbox" data-${p}day="${i}" ${c.dutyDays.includes(i) ? "checked" : ""}>${d}</label>`).join("")}</div></div>
  </div>
  <div class="field">العطل الرسمية والانقطاعات (تُخصم من الساعات المقررة)</div>
  <div class="scroll"><table><thead><tr><th class="l">المناسبة</th><th>من</th><th>إلى</th><th></th></tr></thead><tbody>
    ${c.holidays.map((h, i) => `<tr><td><input data-${p}hol="${i}" data-k="name" value="${esc(h.name)}" placeholder="مثلاً: عطلة رسمية"></td><td><input type="date" data-${p}hol="${i}" data-k="from" value="${h.from}"></td><td><input type="date" data-${p}hol="${i}" data-k="to" value="${h.to}"></td><td><button class="btn small danger" type="button" data-${p}holdel="${i}">حذف</button></td></tr>`).join("")}
  </tbody></table></div><button class="btn small" type="button" data-${p}holadd="1" style="justify-self:start">+ عطلة</button>
  <p class="muted small" style="margin:0">المدة الحالية: ${Core.termWeeks({ calendar: c })} أسبوعاً.</p>`;
}
function rulesForm(rl, p = "r") {
  return `<div class="grid2">
    <label class="field">أساس حساب النسبة<select data-${p}="basis">${opt("total", rl.basis, "من مجموع ساعات كل المواد")}${opt("subject", rl.basis, "لكل مادة على حدة")}</select></label>
    <label class="field">الغياب بعذر (مجاز)<select data-${p}="countExcused">${opt("false", String(!!rl.countExcused), "لا يُحسب ضمن النسبة")}${opt("true", String(!!rl.countExcused), "يُحسب ضمن النسبة")}</select></label>
    <label class="field">الإنذار الأولي عند (%)<input type="number" min="0" max="100" step="0.5" data-${p}="warn1" value="${rl.warn1}"></label>
    <label class="field">الإنذار الثاني عند (%)<input type="number" min="0" max="100" step="0.5" data-${p}="warn2" value="${rl.warn2}"></label>
    <label class="field">الحرمان / الرسوب بالغياب عند (%) — 0 = لا يوجد<input type="number" min="0" max="100" step="0.5" data-${p}="warn3" value="${rl.warn3 || 0}"></label>
  </div><p class="muted small" style="margin:0">عدّل النسب حسب تعليمات جامعتك وقسمك؛ القيم الافتراضية مقترحة فقط.</p>`;
}
function subjectsForm(list, p = "s") {
  return `<div class="scroll"><table><thead><tr><th class="l">المادة</th><th class="l">التدريسي</th><th>الرمز</th><th>النوع</th><th>الجدول الأسبوعي (اليوم · من · إلى)</th><th>ساعات/أسبوع</th><th></th></tr></thead><tbody>
    ${list.map((s, i) => `<tr><td><input data-${p}="${i}" data-k="name" value="${esc(s.name)}" placeholder="اسم المادة"></td><td><input data-${p}="${i}" data-k="teacher" value="${esc(s.teacher)}"></td><td><input data-${p}="${i}" data-k="code" value="${esc(s.code || "")}" style="min-width:60px"></td>
      <td><select data-${p}="${i}" data-k="kind">${["نظري", "عملي", "نظري وعملي", "سريري", "حلقة نقاشية"].map(k => opt(k, s.kind || "نظري")).join("")}</select></td>
      <td class="l"><div style="display:grid;gap:4px">${(s.slots || []).map((sl, j) => `<div class="row" style="gap:4px;flex-wrap:nowrap"><select data-${p}slot="${i}:${j}" data-k="day">${Core.WEEKDAYS.map((d, di) => opt(di, sl.day, d)).join("")}</select><input type="time" data-${p}slot="${i}:${j}" data-k="from" value="${sl.from}" style="min-width:90px"><input type="time" data-${p}slot="${i}:${j}" data-k="to" value="${sl.to}" style="min-width:90px"><button class="btn small" type="button" data-${p}slotdel="${i}:${j}" aria-label="حذف الموعد">×</button></div>`).join("")}
        <button class="btn small" type="button" data-${p}slotadd="${i}" style="justify-self:start">+ موعد</button>
        ${(s.slots || []).length ? "" : `<div class="row small muted" style="gap:4px">بدون جدول: <input type="number" min="0" step="0.5" data-${p}="${i}" data-k="hours" value="${s.hours || 2}" style="width:60px;min-width:0"> س × <input type="number" min="0" data-${p}="${i}" data-k="perWeek" value="${s.perWeek || 1}" style="width:50px;min-width:0"> محاضرة/أسبوع</div>`}</div></td>
      <td>${Core.subjectWeeklyHours(s)}</td><td><button class="btn small danger" type="button" data-${p}del="${i}">حذف</button></td></tr>`).join("")}
  </tbody></table></div><button class="btn small" type="button" data-${p}add="1" style="justify-self:start">+ مادة</button>`;
}
function studentsForm(list, subjects, defStart, p = "t") {
  return `<div class="scroll"><table><thead><tr><th class="l">اسم الطالب</th><th>يُحسب دوامه من</th><th class="l">مواد مستوفاة / معفى منها</th><th></th></tr></thead><tbody>
    ${list.map((t, i) => `<tr><td><input data-${p}="${i}" data-k="name" value="${esc(t.name)}"></td><td><input type="date" data-${p}="${i}" data-k="start" value="${t.start || defStart || ""}"></td>
      <td class="l"><div class="ex">${subjects.map(su => `<label><input type="checkbox" data-${p}ex="${i}" data-sub="${su.id}" ${(t.exempt || []).includes(su.id) ? "checked" : ""}>${esc(su.name || "؟")}</label>`).join("") || `<span class="muted small">أضف المواد أولاً</span>`}</div></td>
      <td><button class="btn small danger" type="button" data-${p}del="${i}">حذف</button></td></tr>`).join("")}
  </tbody></table></div>
  <div class="row"><button class="btn small" type="button" data-${p}add="1">+ طالب</button></div>
  <label class="field">أو الصق قائمة أسماء (اسم في كل سطر)<textarea data-${p}paste="1" style="min-height:70px"></textarea></label><button class="btn small" type="button" data-${p}pastego="1" style="justify-self:start">إضافة الأسماء</button>`;
}
// binds the shared forms to an object graph {meta, calendar, rules, subjects, students}; onChange(rerender?)
function bindForms(T, onChange, k = { m: "m", c: "c", r: "r", s: "s", t: "t" }) {
  $$(`[data-${k.m}]`).forEach(i => i.onchange = () => { T.meta[i.dataset[k.m]] = i.value.trim(); if (i.dataset[k.m] === "level") { Object.assign(T.rules, Core.PRESETS[i.value]); } onChange(i.dataset[k.m] === "level"); });
  $$(`[data-${k.c}]`).forEach(i => i.onchange = () => { const f = i.dataset[k.c]; T.calendar[f] = f === "weeks" ? Math.max(0, num(i.value)) : i.value; onChange(true); });
  $$(`[data-${k.c}day]`).forEach(i => i.onchange = () => { T.calendar.dutyDays = $$(`[data-${k.c}day]:checked`).map(x => +x.dataset[k.c + "day"]); onChange(); });
  $$(`[data-${k.c}hol]`).forEach(i => i.onchange = () => { T.calendar.holidays[+i.dataset[k.c + "hol"]][i.dataset.k] = i.value; onChange(); });
  $$(`[data-${k.c}holdel]`).forEach(b => b.onclick = () => { T.calendar.holidays.splice(+b.dataset[k.c + "holdel"], 1); onChange(true); });
  $$(`[data-${k.c}holadd]`).forEach(b => b.onclick = () => { T.calendar.holidays.push({ name: "", from: T.calendar.start || today(), to: T.calendar.start || today() }); onChange(true); });
  $$(`[data-${k.r}]`).forEach(i => i.onchange = () => { const f = i.dataset[k.r]; T.rules[f] = f === "basis" ? i.value : f === "countExcused" ? i.value === "true" : Math.max(0, num(i.value)); onChange(); });
  $$(`[data-${k.s}]`).forEach(i => i.onchange = () => { const s = T.subjects[+i.dataset[k.s]], f = i.dataset.k; s[f] = ["hours", "perWeek"].includes(f) ? Math.max(0, num(i.value)) : i.value.trim(); onChange(["hours", "perWeek", "name"].includes(f)); });
  $$(`[data-${k.s}slot]`).forEach(i => i.onchange = () => { const [a, b] = i.dataset[k.s + "slot"].split(":").map(Number); const sl = T.subjects[a].slots[b]; sl[i.dataset.k] = i.dataset.k === "day" ? +i.value : i.value; onChange(true); });
  $$(`[data-${k.s}slotdel]`).forEach(x => x.onclick = () => { const [a, b] = x.dataset[k.s + "slotdel"].split(":").map(Number); T.subjects[a].slots.splice(b, 1); onChange(true); });
  $$(`[data-${k.s}slotadd]`).forEach(x => x.onclick = () => { const s = T.subjects[+x.dataset[k.s + "slotadd"]]; s.slots ||= []; s.slots.push({ day: T.calendar.dutyDays[0] ?? 0, from: "08:30", to: "10:30" }); onChange(true); });
  $$(`[data-${k.s}del]`).forEach(x => x.onclick = () => twoStep(x, () => { const s = T.subjects.splice(+x.dataset[k.s + "del"], 1)[0]; T.students.forEach(t => t.exempt = (t.exempt || []).filter(e => e !== s.id)); onChange(true); }));
  $$(`[data-${k.s}add]`).forEach(x => x.onclick = () => { T.subjects.push({ id: "s" + uid(), name: "", teacher: "", code: "", kind: "نظري", slots: [], hours: 2, perWeek: 1 }); onChange(true); });
  $$(`[data-${k.t}]`).forEach(i => i.onchange = () => { const t = T.students[+i.dataset[k.t]]; t[i.dataset.k] = i.value.trim() || t[i.dataset.k]; onChange(); });
  $$(`[data-${k.t}ex]`).forEach(i => i.onchange = () => { const t = T.students[+i.dataset[k.t + "ex"]]; t.exempt = (t.exempt || []).filter(e => e !== i.dataset.sub); if (i.checked) t.exempt.push(i.dataset.sub); onChange(); });
  $$(`[data-${k.t}del]`).forEach(x => x.onclick = () => twoStep(x, () => { T.students.splice(+x.dataset[k.t + "del"], 1); onChange(true); }));
  $$(`[data-${k.t}add]`).forEach(x => x.onclick = () => { T.students.push({ id: "t" + uid(), name: "طالب جديد", start: T.calendar.start || today(), exempt: [] }); onChange(true); });
  $$(`[data-${k.t}pastego]`).forEach(x => x.onclick = () => { const ta = $(`[data-${k.t}paste]`); const names = ta.value.split(/\r?\n/).map(s => s.replace(/^\s*\d+[\s.\-)]*/, "").trim()).filter(Boolean); names.forEach(n => T.students.push({ id: "t" + uid(), name: n, start: T.calendar.start || today(), exempt: [] })); onChange(true); toast(`أُضيف ${names.length} طالباً`); });
}

/* ===== setup (current register) ===== */
function setupView() {
  const r = R();
  const p = C.provider, list = ui.modelLists[p] || [], models = [...new Set([C.models[p], ...list].filter(Boolean))];
  const reg = r ? `
   <section class="panel"><div class="row"><h2 style="margin:0">معلومات السجل</h2><span class="spacer"></span><a class="btn small" href="#new">+ سجل جديد</a></div>${metaForm(r.meta)}</section>
   <section class="panel"><h2 style="margin:0">التقويم الدراسي</h2>${calendarForm(r.calendar)}</section>
   <section class="panel"><h2 style="margin:0">قواعد الغياب والإنذار</h2>${rulesForm(r.rules)}</section>
   <section class="panel setup"><h2 style="margin:0">المواد والجدول الأسبوعي</h2>${subjectsForm(r.subjects)}</section>
   <section class="panel setup"><h2 style="margin:0">الطلبة</h2>${studentsForm(r.students, r.subjects, r.calendar.start)}</section>` : `<div class="empty">لا يوجد سجل بعد. <a href="#new">أنشئ سجلاً جديداً</a>.</div>`;
  return reg + `
   <section class="panel"><h2 style="margin:0">الذكاء الاصطناعي</h2>
   <p class="muted small" style="margin:0">المفاتيح تُحفظ على هذا الجهاز فقط ولا تُرفع إلى Drive. Gemini يُستخرج مفتاحه مجاناً.</p>
   <div class="grid2">
     <label class="field">المزوّد<select id="prov">${Object.entries(AI.NAMES).map(([k, v]) => opt(k, p, v)).join("")}</select></label>
     <label class="field">النموذج<select id="model">${models.length ? models.map(m => opt(m, C.models[p])).join("") : `<option value="">أدخل المفتاح ثم اضغط «جلب النماذج»</option>`}</select></label>
     <label class="field" style="grid-column:1/-1">مفتاح ${AI.NAMES[p]} (API key)<input id="key" type="password" autocomplete="off" value="${esc(C.keys[p] || "")}" placeholder="${{ openai: "sk-…", gemini: "AIza…", anthropic: "sk-ant-…" }[p]}"></label>
   </div>
   <div class="row"><button class="btn" id="fetchModels" type="button">جلب النماذج واختبار المفتاح</button><span class="muted small" id="aiMsg"></span></div>
   <details class="small"><summary>من أين أحصل على المفتاح؟</summary><ol class="steps" style="margin-top:8px">
     <li>Gemini (مجاني): <code>aistudio.google.com/apikey</code> ← Create API key.</li>
     <li>ChatGPT: <code>platform.openai.com/api-keys</code> (مدفوع حسب الاستخدام).</li>
     <li>Claude: <code>console.anthropic.com/settings/keys</code> (مدفوع حسب الاستخدام).</li></ol></details></section>
   <section class="panel"><h2 style="margin:0">Google Drive</h2>
   <div class="${Drive.connected() ? "banner ok" : "banner"}">${Drive.connected() ? "متصل. تُرفع كل السجلات تلقائياً بعد كل تعديل إلى مجلد «" + Drive.FOLDER + "»." : "غير متصل. اضغط «اتصال» واختر حسابك."}${C.lastSync ? "<br>آخر مزامنة: " + new Date(C.lastSync).toLocaleString("ar-IQ") : ""}${C.pendingSync ? "<br><b>توجد تغييرات بانتظار الرفع.</b>" : ""}</div>
   <p class="muted small" style="margin:0">التطبيق يعمل بالكامل بدون إنترنت: التسجيل والحساب والتصدير. أي تغيير يُحفظ على الجهاز فوراً، ويُرفع إلى Drive تلقائياً عند عودة الاتصال. إذا انتهت جلسة Google (بعد ساعة تقريباً) يظهر في الأعلى «اضغط للرفع».</p>
   <label class="field">Google OAuth Client ID<input id="cid" value="${esc(C.drive.clientId || "")}" style="direction:ltr"></label>
   <div class="row"><button class="btn primary" id="driveGo" type="button">${Drive.connected() ? "مزامنة الآن" : "اتصال"}</button>${Drive.connected() ? `<button class="btn" id="driveOut" type="button">قطع الاتصال</button>` : ""}</div></section>
   <section class="panel"><h2 style="margin:0">التحديثات</h2>
   <p class="muted small" style="margin:0">التطبيق يتحدّث تلقائياً عند فتحه، ولا تفقد بياناتك. الإصدار الحالي: <b>${APP_VERSION}</b></p>
   <div class="row"><button class="btn" id="updBtn" type="button">البحث عن تحديث الآن</button></div></section>
   ${r ? `<section class="panel"><h2 style="margin:0">حذف</h2><div class="row"><button class="btn danger" id="delReg" type="button">حذف هذا السجل</button><button class="btn danger" id="wipeBtn" type="button">مسح كل السجلات</button></div>
   <p class="muted small" style="margin:0">احفظ نسخة احتياطية من «إدخال البيانات» قبل الحذف.</p></section>` : ""}`;
}

/* ===== new-register wizard ===== */
const WIZ_STEPS = ["الدراسة", "التقويم", "قواعد الغياب", "المواد والجدول", "الطلبة", "مراجعة"];
function wizardView() {
  if (!ui.wiz) ui.wiz = { step: 0, T: Core.blankRegister({ level: "أولية", ...(R() ? { university: R().meta.university, college: R().meta.college, department: R().meta.department, year: R().meta.year } : {}) }) };
  const { step, T } = ui.wiz;
  const nav = `<div class="row" style="gap:6px">${WIZ_STEPS.map((s, i) => `<button class="chipbtn" type="button" data-wstep="${i}" style="${i === step ? "border-color:var(--accent);font-weight:700" : ""}">${i + 1}. ${s}</button>`).join("")}</div>`;
  const nWeeks = Core.termWeeks(T);
  const body = [
    () => `<h3>نوع الدراسة والجهة</h3><p class="muted small" style="margin:0">اختر الدراسات الأولية (صباحي أو مسائي) أو العليا، ثم القسم والمرحلة. تتغير قواعد الإنذار الافتراضية حسب نوع الدراسة.</p>${metaForm(T.meta, "wm")}`,
    () => `<h3>التقويم الدراسي</h3><p class="muted small" style="margin:0">حسب التقويم الجامعي لوزارة التعليم العالي. مثال 2026-2027: المراحل ما فوق الأولى من 20/9 إلى 31/12/2026 (15 أسبوعاً)، والمرحلة الأولى من 15/11/2026.</p>
      <div class="row"><button class="btn small" type="button" id="calPg">تعبئة: ما فوق الأولى 2026-2027</button><button class="btn small" type="button" id="calUg1">تعبئة: المرحلة الأولى 2026-2027</button></div>${calendarForm(T.calendar, "wc")}`,
    () => `<h3>قواعد الغياب والإنذار</h3>${rulesForm(T.rules, "wr")}`,
    () => `<h3>المواد والجدول الأسبوعي</h3><p class="muted small" style="margin:0">لكل مادة أضف مواعيدها في الأسبوع (اليوم ووقت البداية والنهاية)، فتُحسب ساعاتها المقررة من الجدول والتقويم والعطل تلقائياً، وتظهر في شاشة «اليوم» للتسجيل السريع.</p>
      <div class="row"><label class="field" style="display:flex;gap:6px;align-items:center">عدد المواد <input type="number" id="wN" min="0" max="40" value="${T.subjects.length}" style="width:70px"></label><button class="btn small" id="wNgo" type="button">إنشاء الصفوف</button></div>${subjectsForm(T.subjects, "ws")}`,
    () => `<h3>الطلبة</h3><p class="muted small" style="margin:0">الطالب الذي التحق متأخراً غيّر تاريخ بدء دوامه، فلا يُحسب عليه ما قبله. وحدّد المواد المستوفاة لكل طالب.</p>${studentsForm(T.students, T.subjects, T.calendar.start, "wt")}`,
    () => `<h3>مراجعة</h3><div class="small" style="display:grid;gap:6px">
      <div><b>${esc(Core.registerName(T))}</b></div>
      <div>${esc(T.meta.university)} · ${esc(T.meta.college)} · قسم ${esc(T.meta.department)} · ${esc(T.meta.year)}</div>
      <div>الفصل: ${fmtDate(T.calendar.start)} – ${fmtDate(T.calendar.end)} · ${nWeeks} أسبوعاً · أيام الدوام: ${T.calendar.dutyDays.map(d => Core.WEEKDAYS[d]).join("، ")} · ${T.calendar.holidays.length} عطلة</div>
      <div>النسبة ${T.rules.basis === "subject" ? "لكل مادة" : "من مجموع المواد"} · إنذار أولي ${T.rules.warn1}% · ثانٍ ${T.rules.warn2}%${T.rules.warn3 ? " · حرمان " + T.rules.warn3 + "%" : ""}</div>
      <div>${T.subjects.filter(s => s.name).length} مادة · ${T.subjects.reduce((a, s) => a + Core.subjectWeeklyHours(s), 0)} ساعة أسبوعياً · ${T.students.length} طالب</div></div>
      ${!T.subjects.filter(s => s.name).length ? `<div class="banner">لم تُضف مواد بعد؛ يمكنك إضافتها لاحقاً من الإعدادات.</div>` : ""}
      ${!T.calendar.start || !T.calendar.end ? `<div class="banner">حدّد بداية الفصل ونهايته لتُحسب الساعات المقررة.</div>` : ""}`
  ][step]();
  return `<section class="panel"><div class="row"><h2 style="margin:0">سجل جديد</h2><span class="spacer"></span><button class="btn small" id="wCancel" type="button">إلغاء</button></div>${nav}</section>
    <section class="panel setup">${body}
    <div class="row"><button class="btn" id="wPrev" type="button" ${step ? "" : "disabled"}>السابق</button><span class="spacer"></span>${step < WIZ_STEPS.length - 1 ? `<button class="btn primary" id="wNext" type="button">التالي</button>` : `<button class="btn primary" id="wDone" type="button">إنشاء السجل</button>`}</div></section>`;
}

/* ===== bindings ===== */
function bind(v) {
  if (v === "overview" && $("#xlsx")) { $("#xlsx").onclick = () => saveFile(`${Core.registerName(R())} ${today()}.xlsx`, Template.build(R()), XLSX_TYPE); $("#xlsxDrive").onclick = sheetToDrive; }
  if (v === "today") {
    $("#dayPick").onchange = e => { ui.day = e.target.value || today(); render(); };
    $("#dPrev").onclick = () => { ui.day = Core.addDays(ui.day, -1); render(); };
    $("#dNext").onclick = () => { ui.day = Core.addDays(ui.day, 1); render(); };
    $("#dToday").onclick = () => { ui.day = today(); render(); };
    $$("[data-q]").forEach(b => b.onclick = () => markQuick(b.dataset.q, b.dataset.st, b.dataset.v));
    $$("[data-allp]").forEach(b => b.onclick = () => { R().sessions.push({ id: "x" + uid(), date: ui.day, subject: b.dataset.allp, note: "", absent: [], excused: [] }); commit("سُجّلت المحاضرة والجميع حاضر"); });
    $$("[data-open]").forEach(b => b.onclick = () => openSess(b.dataset.open));
  }
  if (v === "log") {
    $("#fSub").onchange = e => { ui.subject = e.target.value; render(); };
    $("#fStu").onchange = e => { ui.student = e.target.value; render(); };
    $("#fAbs").onchange = e => { ui.onlyAbs = e.target.checked; render(); };
    $$(".sess").forEach(b => b.onclick = () => openSess(b.dataset.id));
  }
  if (v === "import") bindImport();
  if (v === "ai") bindAI();
  if (v === "setup") bindSetup();
  if (v === "new") bindWizard();
}
function bindImport() {
  $("#xfile").onchange = async e => {
    const f = e.target.files[0]; e.target.value = ""; if (!f) return;
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true });
      if (Importer.isTemplate(wb)) { ui.tplPreview = Importer.fromTemplate(wb); ui.preview = null; }
      else if (!R()) { toast("الملف ليس بالنموذج الموحّد. أنشئ سجلاً أولاً ثم استورد الحضور، أو استخدم النموذج."); return; }
      else { ui.preview = Importer.plan(R(), Importer.fromWorkbook(wb)); ui.tplPreview = null; }
      render();
    } catch (err) { toast("تعذّر قراءة الملف: " + err.message); }
  };
  $("#tplBlank").onclick = () => saveFile("نموذج سجل الحضور الجامعي.xlsx", Template.build(null), XLSX_TYPE);
  if ($("#tplExport")) $("#tplExport").onclick = () => saveFile(`${Core.registerName(R())}.xlsx`, Template.build(R()), XLSX_TYPE);
  if ($("#tplApply")) {
    $("#tplApply").onclick = () => { const nr = ui.tplPreview.R; DB.registers.push(nr); DB.current = nr.id; ui.tplPreview = null; location.hash = "overview"; commit("أُنشئ السجل من النموذج"); };
    if ($("#tplMerge")) $("#tplMerge").onclick = () => { const t = ui.tplPreview.R; const recs = []; for (const se of t.sessions) { const su = t.subjects.find(s => s.id === se.subject); const marks = {}; t.students.forEach(st => { marks[st.name] = (se.excused || []).includes(st.id) ? "X" : (se.absent || []).includes(st.id) ? "A" : "P"; }); recs.push({ date: se.date, subjectName: su.name, marks, note: se.note }); } ui.preview = Importer.plan(R(), recs); ui.tplPreview = null; render(); };
    $("#tplCancel").onclick = () => { ui.tplPreview = null; render(); };
  }
  if ($("#photo")) $("#photo").onchange = e => { const f = e.target.files[0]; e.target.value = ""; if (f) readPhoto(f); };
  if ($("#pasteGo")) $("#pasteGo").onclick = () => { const t = $("#paste").value; if (!t.trim()) { toast("الصق الجدول أولاً"); return; } ui.preview = Importer.plan(R(), Importer.fromText(t, $("#pasteSub").value)); render(); };
  if ($("#applyImp")) { $("#applyImp").onclick = () => { const pv = ui.preview, r = R(); Object.assign(r, { subjects: pv.next.subjects, students: pv.next.students, sessions: pv.next.sessions }); ui.preview = null; commit(`حُفظ: ${pv.created} محاضرة جديدة و${pv.updated} محدّثة`); }; $("#cancelImp").onclick = () => { ui.preview = null; render(); }; }
  $("#bkSave").onclick = () => saveFile(`نسخة سجلات الحضور ${today()}.json`, JSON.stringify(DB, null, 1), "application/json");
  $("#bkLoad").onchange = async e => {
    const f = e.target.files[0]; e.target.value = ""; if (!f) return;
    try { const j = Core.migrate(JSON.parse(await f.text())); if (!j || !j.registers) throw new Error("الملف ليس نسخة من هذا التطبيق"); DB = j; commit("استُرجعت النسخة"); } catch (err) { toast(err.message); }
  };
}
function bindAI() {
  $$("[data-ask]").forEach(b => b.onclick = () => { $("#q").value = b.dataset.ask; askAI(); });
  $("#askGo").onclick = askAI;
  if ($("#ansCopy")) $("#ansCopy").onclick = () => copy(ui.answer);
  $("#lFrom").onchange = e => { R().meta.letterFrom = e.target.value; commit(); };
  $("#lTpl").onclick = () => { ui.letter = letterTemplate(); render(); };
  $("#lAi").onclick = letterAI;
  if ($("#letter")) {
    $("#letter").oninput = e => ui.letter = e.target.value;
    $("#lCopy").onclick = () => copy(ui.letter);
    $("#lPrint").onclick = () => { $("#printArea").textContent = ui.letter; window.print(); };
    $("#lDoc").onclick = () => saveFile(`كتاب ${$("#lType").value}.doc`, `<html dir="rtl"><head><meta charset="utf-8"></head><body style="font-family:Arial;font-size:14pt;line-height:2;white-space:pre-wrap">${esc(ui.letter)}</body></html>`, "application/msword");
  }
}
function bindSetup() {
  if (R()) {
    bindForms(R(), rerender => commit());
    $("#delReg").onclick = e => twoStep(e.target, () => { DB.registers = DB.registers.filter(r => r.id !== DB.current); DB.current = DB.registers[0]?.id || null; commit("حُذف السجل"); }, "الحذف");
    $("#wipeBtn").onclick = e => twoStep(e.target, () => { DB = Core.emptyDB(); commit("مُسحت كل السجلات"); }, "المسح");
  }
  $("#prov").onchange = e => { C.provider = e.target.value; saveCfg(); render(); };
  $("#model").onchange = e => { C.models[C.provider] = e.target.value; saveCfg(); toast("اختير النموذج"); };
  $("#key").onchange = e => { C.keys[C.provider] = e.target.value.trim(); saveCfg(); };
  $("#fetchModels").onclick = async () => {
    C.keys[C.provider] = $("#key").value.trim(); saveCfg();
    const msg = $("#aiMsg"); msg.textContent = "جارٍ الاتصال…"; msg.className = "muted small";
    try { const list = await AI.listModels(C.provider, C.keys[C.provider]); ui.modelLists[C.provider] = list; if (!C.models[C.provider] || !list.includes(C.models[C.provider])) C.models[C.provider] = AI.suggest(C.provider, list); saveCfg(); render(); toast(`المفتاح يعمل — ${list.length} نموذج متاح`); }
    catch (err) { msg.textContent = err.message; msg.className = "small d"; }
  };
  $("#cid").onchange = e => { C.drive.clientId = e.target.value.trim(); saveCfg(); };
  $("#driveGo").onclick = () => { C.drive.clientId = $("#cid").value.trim(); saveCfg(); syncNow(true); };
  if ($("#driveOut")) $("#driveOut").onclick = () => { Drive.signOut(); render(); toast("قُطع الاتصال بـ Drive"); };
  $("#updBtn").onclick = async () => { toast("جارٍ البحث عن تحديث…"); try { await window._swReg?.update(); } catch (e) {} setTimeout(() => location.reload(), 1500); };
}
function bindWizard() {
  const W = ui.wiz, T = W.T, keep = () => { const y = scrollY; render(); scrollTo(0, y); };
  bindForms(T, () => keep(), { m: "wm", c: "wc", r: "wr", s: "ws", t: "wt" });
  $$("[data-wstep]").forEach(b => b.onclick = () => { W.step = +b.dataset.wstep; render(); });
  $("#wPrev").onclick = () => { W.step--; render(); };
  if ($("#wNext")) $("#wNext").onclick = () => { W.step++; render(); scrollTo(0, 0); };
  $("#wCancel").onclick = () => { ui.wiz = null; location.hash = R() ? "overview" : ""; render(); };
  if ($("#calPg")) { $("#calPg").onclick = () => { Object.assign(T.calendar, { start: "2026-09-20", end: "2026-12-31", weeks: 0 }); T.meta.year ||= "2026-2027"; keep(); }; $("#calUg1").onclick = () => { Object.assign(T.calendar, { start: "2026-11-15", end: "2027-02-25", weeks: 0 }); T.meta.year ||= "2026-2027"; keep(); }; }
  if ($("#wNgo")) $("#wNgo").onclick = () => { const n = Math.max(0, Math.min(40, num($("#wN").value))); while (T.subjects.length < n) T.subjects.push({ id: "s" + uid(), name: "", teacher: "", code: "", kind: "نظري", slots: [], hours: 2, perWeek: 1 }); T.subjects.length = n; keep(); };
  if ($("#wDone")) $("#wDone").onclick = () => { T.subjects = T.subjects.filter(s => s.name); DB.registers.push(T); DB.current = T.id; ui.wiz = null; location.hash = "overview"; commit("أُنشئ السجل"); };
}

/* ===== session editor ===== */
let draft = null;
function openSess(id, preset, banner) {
  const r = R(); if (!r) return;
  if (!r.subjects.length) { toast("أضف مادة أولاً من «الإعدادات»"); return; }
  const ex = id && r.sessions.find(s => s.id === id);
  draft = preset || (ex ? clone(ex) : { id: "x" + uid(), date: ui.day || today(), subject: ui.subject || r.subjects[0].id, note: "", absent: [], excused: [] });
  draft.excused ||= []; draft._banner = banner || "";
  drawSess(!!ex); if (!$("#dlg").open) $("#dlg").showModal();
}
function drawSess(isEdit) {
  const r = R(), f = $("#sessForm");
  const att = r.students.map(st => {
    const x = status(draft, st);
    const ctrl = (x === "N" || x === "E") && !(draft.absent.includes(st.id)) ? `<span class="pill p-${x}">${STATUS[x]}</span>` :
      `<span class="seg">${["P", "A", "X"].map(v => `<button type="button" class="${v}" data-st="${st.id}" data-v="${v}" aria-pressed="${(v === "A" && draft.absent.includes(st.id)) || (v === "X" && draft.excused.includes(st.id)) || (v === "P" && !draft.absent.includes(st.id) && !draft.excused.includes(st.id))}">${STATUS[v]}</button>`).join("")}</span>`;
    return `<div class="attrow"><span>${esc(st.name)}</span>${ctrl}</div>`;
  }).join("");
  const dup = r.sessions.find(s => s.id !== draft.id && s.date === draft.date && s.subject === draft.subject);
  f.innerHTML = `<h2 style="margin:0">${isEdit ? "تعديل محاضرة" : "محاضرة جديدة"}</h2>
   ${draft._banner ? `<div class="banner">${esc(draft._banner)}</div>` : ""}
   <div class="grid2"><label class="field">التاريخ<input type="date" id="dDate" value="${draft.date}" required></label>
   <label class="field">المادة<select id="dSub">${r.subjects.map(s => opt(s.id, draft.subject, s.name)).join("")}</select></label>
   <label class="field">ساعات هذه المحاضرة<input type="number" id="dHrs" min="0" step="0.5" value="${draft.hours || Core.sessionHours(r, { ...draft, hours: 0 })}"></label></div>
   ${!Core.isDuty(r, draft.date) ? `<div class="banner">${dayName(draft.date)} ليس يوم دوام إلزامي: الغياب فيه لا يُحسب.</div>` : ""}
   ${dup ? `<div class="banner">هذه المادة مسجّلة في نفس اليوم. الحفظ سيُدمج الحضور مع المحاضرة الموجودة.</div>` : ""}
   <div class="att">${att}</div>
   <label class="field">ملاحظة (اختياري)<input id="dNote" value="${esc(draft.note)}" placeholder="مثلاً: امتحان"></label>
   <div class="row">${isEdit ? `<button type="button" class="btn danger" id="dDel">حذف</button>` : ""}<span class="spacer"></span>
   <button type="button" class="btn" id="dCancel">إلغاء</button><button type="button" class="btn primary" id="dSave">حفظ</button></div>`;
  $("#dDate").onchange = e => { draft.date = e.target.value || draft.date; drawSess(isEdit); };
  $("#dSub").onchange = e => { draft.subject = e.target.value; drawSess(isEdit); };
  $("#dNote").oninput = e => draft.note = e.target.value;
  f.querySelectorAll("[data-st]").forEach(b => b.onclick = () => { const id = b.dataset.st; draft.absent = draft.absent.filter(x => x !== id); draft.excused = draft.excused.filter(x => x !== id); if (b.dataset.v === "A") draft.absent.push(id); if (b.dataset.v === "X") draft.excused.push(id); drawSess(isEdit); });
  $("#dCancel").onclick = () => $("#dlg").close();
  $("#dSave").onclick = () => {
    const d = clone(draft); delete d._banner; d.note = $("#dNote").value.trim();
    const h = num($("#dHrs").value), def = Core.sessionHours(r, { ...d, hours: 0 }); if (h && h !== def) d.hours = h; else delete d.hours;
    const i = r.sessions.findIndex(s => s.id === d.id);
    if (dup) { Object.assign(dup, { absent: d.absent, excused: d.excused }); if (d.note) dup.note = d.note; if (i >= 0) r.sessions.splice(i, 1); }
    else if (i >= 0) r.sessions[i] = d; else r.sessions.push(d);
    $("#dlg").close(); commit(i >= 0 ? "عُدّلت المحاضرة" : "أُضيفت المحاضرة");
  };
  if (isEdit) $("#dDel").onclick = () => twoStep($("#dDel"), () => { r.sessions = r.sessions.filter(s => s.id !== draft.id); $("#dlg").close(); commit("حُذفت المحاضرة"); });
}

/* ===== AI features ===== */
async function readPhoto(file) {
  const r = R(); ui.photoBusy = true; render();
  try {
    const image = await AI.imageFromFile(file);
    const system = "أنت تقرأ صور أوراق حضور جامعية عراقية مكتوبة بخط اليد وتُرجع JSON فقط بدون أي نص آخر.";
    const text = `هذه صورة ورقة حضور. في أعلاها عادةً: اليوم، التاريخ (يوم/شهر/سنة)، مدرس المادة، المادة. وفيها جدول بعمود «الاسم» وعمود «التوقيع». الطالب الحاضر هو من كُتب اسمه في الجدول. قد تحتوي الصورة نسختين متطابقتين من الورقة — اعتبرهما ورقة واحدة.
الطلبة المسجلون: ${JSON.stringify(r.students.map(s => ({ id: s.id, name: s.name })))}
المواد: ${JSON.stringify(r.subjects.map(s => ({ id: s.id, name: s.name, teacher: s.teacher })))}
أرجع JSON بهذا الشكل بالضبط:
{"date":"YYYY-MM-DD أو null","subject_id":"معرّف المادة من القائمة أو null","subject_text":"اسم المادة كما هو مكتوب","present_ids":["معرّفات الطلبة الذين وردت أسماؤهم"],"unmatched_names":["أسماء لم تطابق أي طالب"],"confidence":"high أو medium أو low","notes":"ملاحظات مثل امتحان أو كتابة غير واضحة"}`;
    const out = AI.extractJSON(await AI.ask(C, { system, text, image }));
    const date = Importer.toYMD(out.date) || today();
    const subject = subj(out.subject_id) ? out.subject_id : (Importer.matchBy(r.subjects, out.subject_text || "")?.id || r.subjects[0].id);
    const present = new Set(out.present_ids || []);
    const preset = { id: "x" + uid(), date, subject, note: out.notes && /امتحان/.test(out.notes) ? "امتحان" : "", absent: [], excused: [] };
    r.students.forEach(st => { if (status(preset, st) === "P" && !present.has(st.id)) preset.absent.push(st.id); });
    const warn = [`قُرئت من الصورة بواسطة ${AI.NAMES[C.provider]} — راجع التاريخ والمادة والحضور قبل الحفظ.`];
    if (out.subject_text) warn.push(`المادة المكتوبة: «${out.subject_text}».`);
    if ((out.unmatched_names || []).length) warn.push(`أسماء لم أطابقها: ${out.unmatched_names.join("، ")}.`);
    if (out.confidence === "low") warn.push("الكتابة غير واضحة في أجزاء من الورقة.");
    ui.photoBusy = false; render(); openSess(null, preset, warn.join(" "));
  } catch (err) { ui.photoBusy = false; render(); toast("تعذّر قراءة الصورة: " + err.message); }
}
function dataForAI() {
  const r = R();
  return {
    السجل: Core.registerName(r), الجهة: r.meta, التقويم: { ...r.calendar, الأسابيع: Core.termWeeks(r) },
    القواعد: { الأساس: r.rules.basis === "subject" ? "لكل مادة" : "من مجموع كل المواد", الإنذار_الأولي: r.rules.warn1 + "%", الإنذار_الثاني: r.rules.warn2 + "%", الحرمان: r.rules.warn3 ? r.rules.warn3 + "%" : "لا يوجد" }, اليوم: today(),
    المواد: r.subjects.map(s => ({ المادة: s.name, التدريسي: s.teacher, ساعات_أسبوعية: Core.subjectWeeklyHours(s), الجدول: (s.slots || []).map(x => Core.WEEKDAYS[x.day] + " " + x.from + "-" + x.to) })),
    الطلبة: r.students.map(st => {
      const a = stats(st.id), g = Core.studentLevel(r, st);
      return { الاسم: st.name, يحسب_من: st.start, مستوفي_لـ: (st.exempt || []).map(id => subj(id)?.name).filter(Boolean), الحالة: levelText(g.lv) || "ضمن الحد",
        الإجمالي: { محاضرات_معطاة: a.req, غيابات: a.abs, بعذر: a.exc, ساعات_غياب: a.habs, ساعات_الفصل_المقررة: a.htot, نسبة: pct(a.rate) },
        حسب_المادة: r.subjects.filter(su => !(st.exempt || []).includes(su.id)).map(su => { const b = stats(st.id, su.id); return { المادة: su.name, غيابات: b.abs, ساعات_غياب: b.habs, من_ساعات: b.htot, نسبة: pct(b.rate), تواريخ_الغياب: b.dates }; }) };
    })
  };
}
async function askAI() {
  const q = $("#q").value.trim(); if (!q) { toast("اكتب سؤالك أولاً"); return; }
  ui.asking = true; ui.answer = ""; render(); $("#q").value = q;
  try { ui.answer = await AI.ask(C, { system: "أنت مساعد إداري في قسم أكاديمي عراقي. أجب بالعربية الفصحى المختصرة اعتماداً على بيانات السجل المرفقة فقط، واذكر الأرقام بدقة. لا تخترع بيانات. لا تستخدم Markdown.", text: `بيانات السجل:\n${JSON.stringify(dataForAI())}\n\nالسؤال: ${q}` }); }
  catch (err) { ui.answer = "تعذّر الحصول على إجابة: " + err.message; }
  ui.asking = false; render(); const qq = $("#q"); if (qq) qq.value = q;
}
function letterFacts() { const r = R(), st = r.students.find(s => s.id === $("#lStu").value), subId = $("#lSub").value; return { r, st, su: subj(subId), a: stats(st.id, subId || undefined), type: $("#lType").value, from: $("#lFrom").value }; }
function letterTemplate() {
  const { r, st, su, a, type, from } = letterFacts();
  const t = new Date().toLocaleDateString("ar-IQ", { year: "numeric", month: "long", day: "numeric" });
  return `${r.meta.university ? r.meta.university + "\n" : ""}${r.meta.college ? r.meta.college + "\n" : ""}${from}
العدد:
التاريخ: ${t}

إلى/ الطالب ${st.name}
${Core.registerName(r)}

م/ ${type}

نظراً لتغيّبك عن ${su ? `مادة «${su.name}»` : "المحاضرات"} ${a.abs} ${a.abs === 1 ? "مرة" : "مرات"} بواقع ${a.habs} ساعة من أصل ${a.htot} ساعة مقررة، أي بنسبة غياب قدرها ${pct(a.rate)}${a.dates.length ? `، وذلك في التواريخ: ${a.dates.map(fmtDate).join("، ")}` : ""}، وحيث إن التعليمات النافذة تقضي بتوجيه الإنذار الأولي عند تجاوز ${r.rules.warn1}% والإنذار الثاني عند تجاوز ${r.rules.warn2}%${r.rules.warn3 ? ` والحرمان عند تجاوز ${r.rules.warn3}%` : ""} من الساعات المقررة، لذا تقرر توجيه ${type} إليك، مع التأكيد على ضرورة الالتزام بالدوام، وبخلافه ستُتخذ الإجراءات القانونية وفق التعليمات.

مع التقدير.

رئيس القسم
التوقيع:

نسخة منه إلى:
- ${r.meta.level === "عليا" ? "لجنة الدراسات العليا في القسم" : "مقرر القسم"}
- الحفظ`;
}
async function letterAI() {
  const btn = $("#lAi"); btn.disabled = true; btn.textContent = "جارٍ الصياغة…";
  const { r, st, su, a, type, from } = letterFacts();
  try {
    ui.letter = (await AI.ask(C, {
      system: "أنت كاتب مراسلات رسمية في جامعة عراقية. اكتب الكتب الرسمية بالأسلوب الإداري العراقي المعتاد (الجهة، العدد، التاريخ، إلى، م/، المتن، التوقيع، نسخة منه إلى). نص عادي بدون Markdown.",
      text: `اكتب ${type} رسمياً للطالب «${st.name}» (${Core.registerName(r)}، ${r.meta.college}، ${r.meta.university}) صادراً عن «${from}».
الوقائع (لا تغيّرها): ${su ? `المادة: ${su.name}، التدريسي: ${su.teacher}` : "كل المواد"}؛ الغيابات ${a.abs} من ${a.req} محاضرة؛ ساعات الغياب ${a.habs} من ${a.htot} ساعة مقررة؛ النسبة ${pct(a.rate)}؛ الإنذار الأولي ${r.rules.warn1}% والثاني ${r.rules.warn2}%؛ تواريخ الغياب: ${a.dates.map(fmtDate).join("، ") || "لا يوجد"}؛ تاريخ اليوم ${fmtDate(today())}. اترك خانة العدد فارغة.`
    })).trim();
  } catch (err) { toast("تعذّرت الصياغة: " + err.message); }
  render();
}

/* ===== Drive sync (all registers in one file) ===== */
let syncing = false, syncTimer = null;
function syncBadge(state, text) {
  const b = $("#syncBtn"), on = Drive.connected(), online = navigator.onLine !== false;
  if (!state && !text) {
    if (!online) { state = "off"; text = C.pendingSync ? "بدون إنترنت — التغييرات محفوظة وستُرفع عند الاتصال" : "بدون إنترنت — يعمل من الجهاز"; }
    else if (C.pendingSync && C.drive.clientId) { state = "busy"; text = on ? "تغييرات بانتظار الرفع…" : "تغييرات بانتظار الرفع — اضغط للرفع"; }
    else { state = on ? "" : "off"; text = on ? "متزامن مع Drive" : (C.drive.clientId ? "اضغط للمزامنة مع Drive" : "محفوظ على هذا الجهاز"); }
  }
  b.className = "save " + state; b.querySelector("span").textContent = text;
  b.title = text;
}
function scheduleSync() {
  syncBadge();
  if (!Drive.connected() || navigator.onLine === false) return; // stays queued; uploaded when back online
  clearTimeout(syncTimer); syncTimer = setTimeout(() => syncNow(false), 3000);
}
// back online, or app reopened: flush the queue
function flushQueue() {
  syncBadge();
  if (!C.pendingSync || !C.drive.clientId || navigator.onLine === false) return;
  if (Drive.connected()) syncNow(false);
}
window.addEventListener("online", () => { toast("عاد الاتصال بالإنترنت"); flushQueue(); });
window.addEventListener("offline", () => { syncBadge(); toast("انقطع الإنترنت — استمر بالعمل، وسيُرفع كل شيء عند الاتصال"); });
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") flushQueue(); });
setInterval(() => { if (C.pendingSync && Drive.connected() && navigator.onLine !== false && !syncing) syncNow(false); }, 60e3);
// the Google session lasts about an hour; when it has expired, the next tap on the badge renews it and uploads the queue
async function syncNow(interactive) {
  if (syncing) return; syncing = true; syncBadge("busy", "جارٍ المزامنة…");
  try {
    await Drive.signIn(C.drive.clientId, interactive);
    const remote = Core.migrate(await Drive.pull(C.drive));
    const localChanged = !C.lastSync || (DB.updatedAt || "") > C.lastSync;
    const remoteNewer = remote && (remote.updatedAt || "") > (DB.updatedAt || "");
    if (remote && remoteNewer && !localChanged) { DB = remote; store(KEY, DB); toast("سُحب أحدث نسخة من Drive"); }
    else if (remote && remoteNewer && localChanged) {
      if (interactive) { await Drive.push(C.drive, DB); toast("رُفعت نسختك. النسخة السابقة محفوظة في سجل إصدارات الملف على Drive"); }
      else { DB = remote; store(KEY, DB); toast("سُحبت نسخة أحدث من جهاز آخر"); }
    } else if (!remote || (DB.updatedAt || "") !== (remote.updatedAt || "")) await Drive.push(C.drive, DB);
    C.lastSync = new Date().toISOString(); C.pendingSync = false; saveCfg(); render();
  } catch (err) {
    if (navigator.onLine === false) syncBadge(); else syncBadge("err", C.pendingSync ? "تعذّر الرفع — التغييرات محفوظة، اضغط للمحاولة" : "تعذّرت المزامنة");
    if (interactive) toast(navigator.onLine === false ? "لا يوجد إنترنت الآن — التغييرات محفوظة على الجهاز" : err.message);
  }
  syncing = false;
}
$("#syncBtn").onclick = () => {
  if (!C.drive.clientId) { location.hash = "setup"; toast("أضف Client ID أولاً"); return; }
  if (navigator.onLine === false) { toast("لا يوجد إنترنت الآن — التغييرات محفوظة على الجهاز وستُرفع عند الاتصال"); return; }
  syncNow(true);
};
async function sheetToDrive() {
  if (!C.drive.clientId) { location.hash = "setup"; toast("اربط Google Drive أولاً من الإعدادات"); return; }
  try { await Drive.signIn(C.drive.clientId, true); const link = await Drive.uploadSheet(C.drive, `${Core.registerName(R())} ${today()}`, await Template.build(R())); saveCfg(); toast("رُفعت نسخة Google Sheets إلى مجلد «" + Drive.FOLDER + "»"); if (link) window.open(link, "_blank"); }
  catch (err) { toast(err.message); }
}

/* ===== header + install + boot ===== */
$("#addBtn").onclick = () => openSess(null);
$("#regSel").onchange = e => { if (e.target.value === "__new") { location.hash = "new"; render(); return; } DB.current = e.target.value; ui.subject = ""; ui.student = ""; commit(); };
let installEvt = null;
window.addEventListener("beforeinstallprompt", e => { e.preventDefault(); installEvt = e; $("#installBtn").hidden = false; });
$("#installBtn").onclick = async () => { if (!installEvt) return; installEvt.prompt(); await installEvt.userChoice; installEvt = null; $("#installBtn").hidden = true; };
window.addEventListener("appinstalled", () => toast("ثُبّت التطبيق"));
const APP_VERSION = "v10";
if ("serviceWorker" in navigator && location.protocol === "https:") {
  const hadController = !!navigator.serviceWorker.controller; let reloaded = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => { if (hadController && !reloaded) { reloaded = true; location.reload(); } });
  navigator.serviceWorker.register("sw.js").then(r => { window._swReg = r; setInterval(() => r.update().catch(() => {}), 30 * 60e3); }).catch(() => {});
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") window._swReg?.update().catch(() => {}); });
}
window.addEventListener("hashchange", render);
if (DB.registers.length && !DB.current) DB.current = DB.registers[0].id;
store(KEY, DB);
render();
flushQueue();
