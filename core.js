/* ===== core: data model, migration, attendance rules (no UI) =====
 DB = { version: 2, current, registers: [R], updatedAt }
 R  = { id, meta, calendar, rules, subjects, students, sessions, createdAt }
   meta:     university, college, department, level ("أولية"|"عليا"), degree, shift, stage, section, system, semester, year
   calendar: start, end, weeks? (manual override), dutyDays [0..6], holidays [{from,to,name}]
   rules:    basis ("total"|"subject"), warn1, warn2, warn3 (0 = off), countExcused
   subject:  id, name, code, teacher, kind, hours (per lecture, fallback), perWeek (fallback), slots [{day, from:"HH:MM", to:"HH:MM"}]
   student:  id, name, start, exempt [subjectId]
   session:  id, date, subject, note, absent [studentId], excused [studentId], hours? (override)
*/
const Core = (() => {
  const WEEKDAYS = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"];
  const uid = () => Math.random().toString(36).slice(2, 9);
  const clone = o => JSON.parse(JSON.stringify(o));
  const dow = d => new Date(d + "T12:00:00").getDay();
  const addDays = (d, n) => { const x = new Date(d + "T12:00:00"); x.setDate(x.getDate() + n); return x.toISOString().slice(0, 10); };
  const toMin = t => { const [h, m] = String(t || "0:0").split(":").map(Number); return (h || 0) * 60 + (m || 0); };
  const slotHours = s => Math.max(0, (toMin(s.to) - toMin(s.from)) / 60);

  const PRESETS = {
    "عليا": { basis: "total", warn1: 5, warn2: 7, warn3: 0, countExcused: false },
    "أولية": { basis: "subject", warn1: 5, warn2: 10, warn3: 15, countExcused: false }
  };

  function blankRegister(meta = {}) {
    const level = meta.level || "عليا";
    return {
      id: "r" + uid(), createdAt: new Date().toISOString(),
      meta: { university: "", college: "", department: "", level, degree: level === "عليا" ? "ماجستير" : "", shift: level === "أولية" ? "صباحي" : "", stage: "", section: "", system: "فصلي", semester: "الأول", year: "", ...meta },
      calendar: { start: "", end: "", weeks: 0, dutyDays: [0, 1, 2, 3, 4], holidays: [] },
      rules: { ...PRESETS[level] },
      subjects: [], students: [], sessions: []
    };
  }

  function registerName(R) {
    const m = R.meta;
    const who = m.level === "عليا" ? [m.degree || "دراسات عليا", m.stage].filter(Boolean).join(" ") : ["المرحلة " + (m.stage || "?"), m.shift, m.section ? "شعبة " + m.section : ""].filter(Boolean).join(" · ");
    return [m.department, who, m.semester ? "الفصل " + m.semester : ""].filter(Boolean).join(" — ") || "سجل بلا اسم";
  }

  // v1 (single masters register) → v2
  function migrate(old) {
    if (!old) return null;
    if (old.version === 2) return old;
    const st = old.settings || {};
    const off = st.offDays ?? [3, 4];
    const R = {
      id: "r" + uid(), createdAt: new Date().toISOString(),
      meta: { university: "الجامعة التقنية الوسطى", college: "الكلية التقنية الصحية والطبية / بغداد", department: "تقنيات صحة المجتمع", level: "عليا", degree: "ماجستير", shift: "", stage: "", section: "", system: "فصلي", semester: "الأول", year: "2026-2027" },
      calendar: { start: st.termStart || "2026-09-20", end: st.termEnd || "2026-12-31", weeks: +st.termWeeks || 15, dutyDays: [0, 1, 2, 3, 4, 5, 6].filter(d => !off.includes(d) && d !== 5 && d !== 6), holidays: [] },
      rules: { basis: "total", warn1: +st.warn1 || 5, warn2: +st.warn2 || 7, warn3: 0, countExcused: false },
      subjects: (old.subjects || []).map(s => ({ code: "", kind: "نظري", slots: [], perWeek: 1, hours: 2, ...s })),
      students: old.students || [], sessions: (old.sessions || []).map(s => ({ excused: [], ...s }))
    };
    return { version: 2, current: R.id, registers: [R], updatedAt: old.updatedAt || new Date().toISOString() };
  }
  const emptyDB = () => ({ version: 2, current: null, registers: [], updatedAt: new Date().toISOString() });

  /* ----- calendar ----- */
  const isHoliday = (R, d) => (R.calendar.holidays || []).some(h => h.from && d >= h.from && d <= (h.to || h.from));
  const isDuty = (R, d) => (R.calendar.dutyDays || [0, 1, 2, 3, 4]).includes(dow(d));
  function termWeeks(R) {
    const c = R.calendar; if (+c.weeks) return +c.weeks;
    if (!c.start || !c.end) return 0;
    const days = (new Date(c.end + "T12:00:00") - new Date(c.start + "T12:00:00")) / 864e5 + 1;
    return Math.max(0, Math.ceil(days / 7));
  }
  function studentWeeks(R, stu) {
    const c = R.calendar, w = termWeeks(R);
    if (!stu.start || !c.start || stu.start <= c.start) return w;
    const days = (new Date(c.end + "T12:00:00") - new Date(stu.start + "T12:00:00")) / 864e5 + 1;
    return Math.max(0, Math.min(w, Math.ceil(days / 7)));
  }
  const subjectWeeklyHours = su => (su.slots || []).length ? su.slots.reduce((a, s) => a + slotHours(s), 0) : (+su.perWeek || 1) * (+su.hours || 0);
  // hours planned for one student in one subject over the term
  function plannedHours(R, stu, su) {
    if (!su || (stu.exempt || []).includes(su.id)) return 0;
    const c = R.calendar;
    if ((su.slots || []).length && c.start && c.end && !+c.weeks) {
      let h = 0; const from = stu.start && stu.start > c.start ? stu.start : c.start;
      for (let d = from; d <= c.end; d = addDays(d, 1)) {
        if (isHoliday(R, d)) continue;
        const wd = dow(d); for (const s of su.slots) if (+s.day === wd) h += slotHours(s);
      }
      return h;
    }
    return studentWeeks(R, stu) * subjectWeeklyHours(su);
  }
  function sessionHours(R, se) {
    if (se.hours) return +se.hours;
    const su = R.subjects.find(s => s.id === se.subject); if (!su) return 0;
    const slot = (su.slots || []).find(s => +s.day === dow(se.date));
    return slot ? slotHours(slot) : +su.hours || 0;
  }

  /* ----- attendance ----- */
  // P present · A absent (counted) · X excused · N not required · E exempt
  function status(R, se, stu) {
    if (stu.start && se.date < stu.start) return "N";
    if ((stu.exempt || []).includes(se.subject)) return "E";
    if ((se.excused || []).includes(stu.id)) return R.rules.countExcused && isDuty(R, se.date) && !isHoliday(R, se.date) ? "A" : "X";
    const absent = (se.absent || []).includes(stu.id);
    if (absent && (!isDuty(R, se.date) || isHoliday(R, se.date))) return "N"; // absence on a no-duty day or a holiday is not counted
    return absent ? "A" : "P";
  }
  function stats(R, stuId, subId) {
    const stu = R.students.find(s => s.id === stuId);
    let req = 0, abs = 0, exc = 0, hreq = 0, habs = 0; const dates = [];
    for (const se of R.sessions) {
      if (subId && se.subject !== subId) continue;
      const st = status(R, se, stu); if (st === "N" || st === "E") continue;
      const h = sessionHours(R, se); req++; hreq += h;
      if (st === "A") { abs++; habs += h; dates.push(se.date); } else if (st === "X") exc++;
    }
    let htot = 0;
    for (const su of (subId ? [R.subjects.find(s => s.id === subId)] : R.subjects)) htot += plannedHours(R, stu, su);
    const rate = htot ? habs / htot : 0, r = R.rules;
    const left = p => Math.max(0, Math.floor(htot * p / 100 - habs));
    return { req, abs, exc, hreq, habs, htot, rate, left1: left(r.warn1), left2: left(r.warn2), left3: r.warn3 ? left(r.warn3) : null, dates: dates.sort() };
  }
  // level for a rate: "" | "w1" | "w2" | "w3"
  function level(R, rate) { const r = R.rules, p = rate * 100; return r.warn3 && p >= r.warn3 ? "w3" : p >= r.warn2 ? "w2" : p >= r.warn1 ? "w1" : ""; }
  const LEVEL_TEXT = { w1: "إنذار أولي", w2: "إنذار ثانٍ", w3: "حرمان / رسوب بالغياب" };
  // the student's governing status under the register's basis
  function studentLevel(R, stu) {
    if (R.rules.basis === "subject") {
      let worst = { lv: "", rate: 0, subject: null };
      for (const su of R.subjects) { const b = stats(R, stu.id, su.id); if (!b.htot) continue; const lv = level(R, b.rate); if (rank(lv) > rank(worst.lv) || (rank(lv) === rank(worst.lv) && b.rate > worst.rate)) worst = { lv, rate: b.rate, subject: su }; }
      return worst;
    }
    const a = stats(R, stu.id); return { lv: level(R, a.rate), rate: a.rate, subject: null };
  }
  const rank = lv => ({ "": 0, w1: 1, w2: 2, w3: 3 }[lv]);

  // lectures scheduled on a given date (from the timetable), for quick marking
  function scheduledOn(R, d) {
    const wd = dow(d), out = [];
    for (const su of R.subjects) for (const s of su.slots || []) if (+s.day === wd) out.push({ subject: su, slot: s });
    return out.sort((a, b) => toMin(a.slot.from) - toMin(b.slot.from));
  }

  // Merge two copies of an account (this device vs Google Drive): per register the newer copy wins,
  // deletions are kept as tombstones so a register deleted on one device does not come back.
  function mergeDB(a, b) {
    if (!b || !(b.registers || b.deleted)) return a;
    if (!a) return b;
    const stamp = r => r.updatedAt || r.createdAt || "";
    const dead = new Map();
    for (const t of [...(a.deleted || []), ...(b.deleted || [])]) if (!dead.has(t.id) || dead.get(t.id) < t.at) dead.set(t.id, t.at);
    const byId = new Map(), order = [];
    for (const r of [...(a.registers || []), ...(b.registers || [])]) { if (!byId.has(r.id)) order.push(r.id); const x = byId.get(r.id); if (!x || stamp(r) > stamp(x)) byId.set(r.id, r); }
    const registers = order.map(id => byId.get(id)).filter(r => !(dead.has(r.id) && dead.get(r.id) >= stamp(r)));
    const has = id => registers.some(r => r.id === id);
    return { version: 2, registers, deleted: [...dead].map(([id, at]) => ({ id, at })),
      current: has(a.current) ? a.current : has(b.current) ? b.current : registers[0]?.id || null,
      updatedAt: [a.updatedAt || "", b.updatedAt || ""].sort().pop() };
  }
  const sameData = (a, b) => !!a && !!b && JSON.stringify([a.registers, a.deleted || []]) === JSON.stringify([b.registers, b.deleted || []]);

  return { mergeDB, sameData, WEEKDAYS, PRESETS, uid, clone, dow, addDays, toMin, slotHours, blankRegister, registerName, migrate, emptyDB, isHoliday, isDuty, termWeeks, studentWeeks, subjectWeeklyHours, plannedHours, sessionHours, status, stats, level, LEVEL_TEXT, studentLevel, scheduledOn };
})();
