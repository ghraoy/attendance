// Reads attendance from Excel / pasted tables into a preview of changes.
// Understands two layouts:
//  1) Long:  التاريخ | المادة | الطالب | الحالة   (one row per student per lecture)
//  2) Wide:  التاريخ | [المادة] | <student 1> | <student 2> ... | ملاحظات   (one row per lecture; sheet name = subject when no المادة column)
window.Importer = (() => {
  const norm = s => String(s ?? "").replace(/[ً-ْـ]/g, "").replace(/[أإآ]/g, "ا").replace(/ة/g, "ه").replace(/ى/g, "ي").replace(/\s+/g, " ").trim().toLowerCase();
  const H = { date: ["التاريخ", "تاريخ", "date"], subject: ["الماده", "ماده", "subject", "course"], student: ["الطالب", "الاسم", "اسم الطالب", "student", "name"], status: ["الحاله", "الحضور", "status"], note: ["ملاحظات", "ملاحظه", "notes", "note"], skip: ["ت", "#", "اليوم", "day", "التسلسل"] };
  const isH = (cell, k) => H[k].includes(norm(cell));
  const P = ["حاضر", "حاضره", "ح", "p", "present", "1", "✓", "✔", "نعم", "yes"].map(norm);
  const A = ["غائب", "غائبه", "غ", "a", "absent", "0", "✗", "✘", "x", "لا", "no"].map(norm);
  const E = ["مستوفيه", "مستوفي", "معفي", "معفي من الماده", "exempt"].map(norm);
  const X = ["مجاز", "مجازه", "بعذر", "غائب بعذر", "اجازه", "excused"].map(norm);
  const statusOf = v => { const n = norm(v); if (!n) return null; if (P.includes(n)) return "P"; if (A.includes(n)) return "A"; if (X.includes(n)) return "X"; if (E.includes(n)) return "E"; return null; };

  function toYMD(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date && !isNaN(v)) { const d = new Date(v.getTime() + 12 * 3600e3); return d.toISOString().slice(0, 10); }
    if (typeof v === "number" && v > 30000 && v < 80000) { const d = new Date(Date.UTC(1899, 11, 30) + v * 86400e3); return d.toISOString().slice(0, 10); }
    const s = String(v).replace(/[٠-٩]/g, c => "٠١٢٣٤٥٦٧٨٩".indexOf(c)).trim();
    let m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
    if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
    return null;
  }

  function matchBy(list, name, key = "name") {
    const n = norm(name); if (!n) return null;
    return list.find(x => norm(x[key]) === n)
      || list.find(x => { const a = norm(x[key]); return a.startsWith(n) || n.startsWith(a); })
      || list.find(x => norm(x[key]).split(" ").slice(0, 2).join(" ") === n.split(" ").slice(0, 2).join(" "));
  }

  // rows: array of arrays. Returns array of {date, subjectName, marks:{studentName:'P'|'A'}, note}
  function readTable(rows, sheetName) {
    const out = []; let hi = rows.findIndex(r => (r || []).some(c => isH(c, "date")));
    if (hi < 0) return out;
    // optional sheet meta above the header (as in the exported register): "… — <teacher>" title and a lecture-hours cell
    const meta = {};
    rows.slice(0, hi).forEach(r => (r || []).forEach((c, i) => {
      const t = String(c ?? "");
      if (t.includes(" — ") && !meta.teacher) meta.teacher = t.split(" — ").pop().trim();
      if (/ساعات/.test(t)) { const n = (r || []).slice(i + 1).find(x => typeof x === "number" || /^\d+(\.\d+)?$/.test(String(x).trim())); if (n !== undefined && meta.hours == null) meta.hours = +n; }
    }));
    const head = rows[hi].map(c => String(c ?? "").trim());
    const col = k => head.findIndex(c => isH(c, k));
    const cDate = col("date"), cSub = col("subject"), cStu = col("student"), cStat = col("status"), cNote = col("note");
    const long = cStu >= 0 && cStat >= 0;
    const studentCols = long ? [] : head.map((c, i) => ({ c, i })).filter(({ c, i }) => c && i !== cDate && i !== cSub && i !== cNote && !Object.values(H).flat().includes(norm(c)));
    for (const r of rows.slice(hi + 1)) {
      if (!r) continue;
      const date = toYMD(r[cDate]); if (!date) continue;
      const subjectName = cSub >= 0 ? String(r[cSub] ?? "").trim() : sheetName;
      if (!subjectName) continue;
      const note = cNote >= 0 ? String(r[cNote] ?? "").trim() : "";
      if (long) { const st = statusOf(r[cStat]); const nm = String(r[cStu] ?? "").trim(); if (st && nm) out.push({ date, subjectName, marks: { [nm]: st }, note, meta }); }
      else { const marks = {}; studentCols.forEach(({ c, i }) => { const st = statusOf(r[i]); if (st) marks[c] = st; }); if (Object.keys(marks).length) out.push({ date, subjectName, marks, note, meta }); }
    }
    return out;
  }

  function fromWorkbook(wb) {
    let recs = [];
    for (const name of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: "" });
      recs = recs.concat(readTable(rows, name.trim()));
    }
    return recs;
  }
  function fromText(text, fallbackSubject) {
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    const sep = lines.some(l => l.includes("\t")) ? "\t" : lines.some(l => l.includes(";")) ? ";" : ",";
    const rows = lines.map(l => l.split(sep).map(c => c.trim()));
    return readTable(rows, fallbackSubject);
  }

  // Turn records into a plan against the current state (nothing is changed until apply()).
  function plan(S, recs) {
    S = { ...S, subjects: S.subjects || [], students: S.students || [], sessions: S.sessions || [] };
    const newSubjects = [], newStudents = [], uid = () => Math.random().toString(36).slice(2, 9);
    const subjects = [...S.subjects], students = S.students.map(x => ({ ...x, exempt: [...(x.exempt || [])] }));
    const sub = (n, m = {}) => { let s = matchBy(subjects, n); if (!s) { s = { id: "s" + uid(), name: n, teacher: m.teacher || "", hours: m.hours > 0 ? m.hours : 2 }; subjects.push(s); newSubjects.push(s); } return s; };
    const stu = (n, date, noStart) => { let s = matchBy(students, n); if (!s) { s = { id: "t" + uid(), name: n, start: noStart ? "9999-12-31" : date, exempt: [] }; students.push(s); newStudents.push(s); } else if (!noStart && newStudents.includes(s) && date < s.start) s.start = date; return s; };
    const exempt = [];
    const byKey = new Map();
    for (const r of recs) {
      const s = sub(r.subjectName, r.meta), k = r.date + "|" + s.id;
      const e = byKey.get(k) || { date: r.date, subject: s.id, marks: {}, note: r.note };
      for (const [n, st] of Object.entries(r.marks)) { if (st === "E") { const x = stu(n, r.date, true); exempt.push([x, s.id]); } else e.marks[stu(n, r.date).id] = st; }
      if (r.note && !e.note) e.note = r.note;
      byKey.set(k, e);
    }
    const newExempt = [];
    for (const [st, sid] of exempt) if (!st.exempt.includes(sid)) { st.exempt.push(sid); newExempt.push(st.name + " ← " + subjects.find(x => x.id === sid).name); }
    let created = 0, updated = 0, absences = 0; const sessions = S.sessions.map(x => ({ ...x, absent: [...(x.absent || [])], excused: [...(x.excused || [])] }));
    for (const e of byKey.values()) {
      let s = sessions.find(x => x.date === e.date && x.subject === e.subject);
      if (s) updated++; else { s = { id: "x" + uid(), date: e.date, subject: e.subject, note: e.note || "", absent: [], excused: [] }; sessions.push(s); created++; }
      for (const [id, st] of Object.entries(e.marks)) { s.absent = s.absent.filter(x => x !== id); s.excused = s.excused.filter(x => x !== id); if (st === "A") { s.absent.push(id); absences++; } if (st === "X") s.excused.push(id); }
      if (e.note && !s.note) s.note = e.note;
    }
    return { created, updated, absences, newSubjects, newStudents, newExempt: [...new Set(newExempt)], next: { ...S, subjects, students, sessions } };
  }

  /* ----- full register template (sheets: الإعدادات، العطل، المواد، الطلبة، الحضور) ----- */
  const DAYS = ["الاحد", "الاثنين", "الثلاثاء", "الاربعاء", "الخميس", "الجمعه", "السبت"];
  const dayIndex = v => { const n = norm(v); if (!n) return null; const i = DAYS.findIndex(d => n === d || n.startsWith(d) || d.startsWith(n)); return i >= 0 ? i : null; };
  function toHM(v) {
    if (v == null || v === "") return "";
    if (typeof v === "number") { const m = Math.round((v % 1) * 1440); return String(Math.floor(m / 60)).padStart(2, "0") + ":" + String(m % 60).padStart(2, "0"); }
    if (v instanceof Date) return String(v.getHours()).padStart(2, "0") + ":" + String(v.getMinutes()).padStart(2, "0");
    const m = String(v).replace(/[٠-٩]/g, c => "٠١٢٣٤٥٦٧٨٩".indexOf(c)).match(/(\d{1,2})[:.](\d{2})/); if (!m) { const h = String(v).match(/^\s*(\d{1,2})\s*$/); return h ? h[1].padStart(2, "0") + ":00" : ""; }
    return m[1].padStart(2, "0") + ":" + m[2];
  }
  const yes = v => ["نعم", "yes", "1", "true", "✓", "صح"].includes(norm(v));
  const SETTINGS = [
    ["university", "الجامعه"], ["college", "الكليه"], ["department", "القسم"], ["level", "نوع الدراسه"], ["degree", "الشهاده"], ["shift", "الدراسه"],
    ["stage", "المرحله"], ["section", "الشعبه"], ["system", "النظام الدراسي"], ["semester", "الفصل الدراسي"], ["year", "العام الدراسي"],
    ["start", "تاريخ بدايه الفصل"], ["end", "تاريخ نهايه الفصل"], ["weeks", "عدد الاسابيع"], ["basis", "اساس حساب النسبه"],
    ["warn1", "الانذار الاولي"], ["warn2", "الانذار الثاني"], ["warn3", "الحرمان"], ["countExcused", "احتساب الغياب بعذر"]];
  const rowsOf = (wb, name) => { const sh = wb.Sheets[name]; return sh ? XLSX.utils.sheet_to_json(sh, { header: 1, raw: true, defval: "" }) : null; };
  const isTemplate = wb => wb.SheetNames.includes("الإعدادات") && wb.SheetNames.includes("المواد");
  function fromTemplate(wb) {
    const R = Core.blankRegister();
    const set = {}; const duty = [];
    for (const r of rowsOf(wb, "الإعدادات") || []) {
      const lab = norm(r[0]); if (!lab) continue;
      const day = DAYS.findIndex(d => lab === "دوام يوم " + d || lab === "يوم " + d || lab === d);
      if (day >= 0) { if (yes(r[1])) duty.push(day); continue; }
      const k = SETTINGS.find(([, l]) => lab.startsWith(l)); if (k && set[k[0]] === undefined) set[k[0]] = r[1];
    }
    const txt = v => String(v ?? "").trim();
    const level = norm(set.level).includes("اولي") ? "أولية" : "عليا";
    Object.assign(R, Core.blankRegister({ level }));
    for (const k of ["university", "college", "department", "degree", "shift", "stage", "section", "system", "semester", "year"]) if (txt(set[k])) R.meta[k] = txt(set[k]);
    R.meta.level = level;
    R.calendar.start = toYMD(set.start) || ""; R.calendar.end = toYMD(set.end) || "";
    R.calendar.weeks = /^\d+$/.test(txt(set.weeks)) ? +set.weeks : 0;
    if (duty.length) R.calendar.dutyDays = duty.sort();
    if (txt(set.basis)) R.rules.basis = norm(set.basis).includes("لكل ماده") ? "subject" : "total";
    const pctv = v => { const n = parseFloat(String(v).replace("%", "")); return isNaN(n) ? null : (n > 0 && n < 1 ? n * 100 : n); };
    for (const k of ["warn1", "warn2", "warn3"]) { const n = pctv(set[k]); if (n != null) R.rules[k] = n; }
    if (txt(set.countExcused)) R.rules.countExcused = yes(set.countExcused);
    for (const r of (rowsOf(wb, "العطل") || []).slice(1)) { const from = toYMD(r[0]); if (from) R.calendar.holidays.push({ from, to: toYMD(r[1]) || from, name: txt(r[2]) }); }
    const sr = rowsOf(wb, "المواد") || []; const sh = sr.findIndex(r => r.some(c => norm(c) === "اسم الماده"));
    for (const r of sr.slice(sh + 1)) {
      const name = txt(r[1]); if (!name) continue;
      const slots = [];
      for (let j = 5; j + 2 < 14 + 1; j += 3) { const d = dayIndex(r[j]); const from = toHM(r[j + 1]), to = toHM(r[j + 2]); if (d != null && from && to) slots.push({ day: d, from, to }); }
      const hrs = parseFloat(r[16]), pw = parseFloat(r[17]);
      R.subjects.push({ id: "s" + Core.uid(), name, code: txt(r[2]), teacher: txt(r[3]), kind: txt(r[4]) || "نظري", slots, hours: hrs > 0 ? hrs : slots[0] ? Core.slotHours(slots[0]) : 2, perWeek: slots.length || (pw > 0 ? pw : 1) });
    }
    const tr = rowsOf(wb, "الطلبة") || []; const th = tr.findIndex(r => r.some(c => norm(c) === "اسم الطالب"));
    for (const r of tr.slice(th + 1)) {
      const name = txt(r[1]); if (!name) continue;
      const exempt = txt(r[3]).split(/[،,;\n]+/).map(x => matchBy(R.subjects, x)).filter(Boolean).map(x => x.id);
      R.students.push({ id: "t" + Core.uid(), name, start: toYMD(r[2]) || R.calendar.start || "", exempt });
    }
    const ar = rowsOf(wb, "الحضور");
    const p = plan(R, ar ? readTable(ar, "") : []);
    return { R: { ...p.next, meta: R.meta, calendar: R.calendar, rules: R.rules, id: R.id, createdAt: R.createdAt }, plan: p };
  }

  function template() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([["التاريخ", "المادة", "الطالب", "الحالة", "ملاحظات"], ["2026-11-01", "الوبائيات", "زينب احمد كاظم", "حاضر", ""], ["2026-11-01", "الوبائيات", "محمد علي حسن", "غائب", "مثال — احذف هذين السطرين"]]);
    ws["!cols"] = [{ wch: 12 }, { wch: 22 }, { wch: 26 }, { wch: 10 }, { wch: 28 }]; ws["!views"] = [{ RTL: true }];
    XLSX.utils.book_append_sheet(wb, ws, "الحضور");
    return XLSX.write(wb, { type: "array", bookType: "xlsx" });
  }
  return { fromWorkbook, fromText, plan, template, norm, matchBy, toYMD, isTemplate, fromTemplate, toHM, dayIndex };
})();
