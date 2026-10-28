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
  const statusOf = v => { const n = norm(v); if (!n) return null; if (P.includes(n)) return "P"; if (A.includes(n)) return "A"; if (E.includes(n)) return "E"; return null; };

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
    let created = 0, updated = 0, absences = 0; const sessions = S.sessions.map(x => ({ ...x, absent: [...(x.absent || [])] }));
    for (const e of byKey.values()) {
      let s = sessions.find(x => x.date === e.date && x.subject === e.subject);
      if (s) updated++; else { s = { id: "x" + uid(), date: e.date, subject: e.subject, note: e.note || "", absent: [] }; sessions.push(s); created++; }
      for (const [id, st] of Object.entries(e.marks)) { s.absent = s.absent.filter(x => x !== id); if (st === "A") { s.absent.push(id); absences++; } }
      if (e.note && !s.note) s.note = e.note;
    }
    return { created, updated, absences, newSubjects, newStudents, newExempt: [...new Set(newExempt)], next: { ...S, subjects, students, sessions } };
  }

  function template() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([["التاريخ", "المادة", "الطالب", "الحالة", "ملاحظات"], ["2026-11-01", "الوبائيات", "زينب احمد كاظم", "حاضر", ""], ["2026-11-01", "الوبائيات", "محمد علي حسن", "غائب", "مثال — احذف هذين السطرين"]]);
    ws["!cols"] = [{ wch: 12 }, { wch: 22 }, { wch: 26 }, { wch: 10 }, { wch: 28 }]; ws["!views"] = [{ RTL: true }];
    XLSX.utils.book_append_sheet(wb, ws, "الحضور");
    return XLSX.write(wb, { type: "array", bookType: "xlsx" });
  }
  return { fromWorkbook, fromText, plan, template, norm, matchBy, toYMD };
})();
