/* Unified Excel template (blank or filled from a register), built in the browser with ExcelJS.
   Sheets: اقرأني · الإعدادات · العطل · المواد · الطلبة · الحضور · الملخص
   The importer reads only the plain input cells, so a workbook exported here re-imports cleanly. */
window.Template = (() => {
  const FONT = "Arial", INK = "FF152B3C", HEAD = "FF0E6A66", SUBH = "FFDCEBE9", INPUT = "FFFFF7CC", MUTED = "FF5A6B72";
  const DAYS = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"];
  const S_ROWS = { university: 4, college: 5, department: 6, level: 7, degree: 8, shift: 9, stage: 10, section: 11, system: 12, semester: 13, year: 14, start: 15, end: 16, weeks: 17, weeksUsed: 18, basis: 19, warn1: 20, warn2: 21, warn3: 22, countExcused: 23, day0: 24 };
  const ref = (k) => `'الإعدادات'!$B$${S_ROWS[k]}`;
  const thin = { style: "thin", color: { argb: "FFC9D3D1" } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };
  const ymdToDate = d => d ? new Date(d + "T00:00:00Z") : null;
  const hmToFrac = t => { if (!t) return null; const [h, m] = t.split(":").map(Number); return ((h || 0) * 60 + (m || 0)) / 1440; };

  function sheet(wb, name, cols) {
    const ws = wb.addWorksheet(name, { views: [{ rightToLeft: true, state: "frozen", ySplit: 4 }], properties: { defaultRowHeight: 18 } });
    ws.columns = cols.map(w => ({ width: w }));
    return ws;
  }
  function title(ws, text, sub, span) {
    ws.mergeCells(1, 1, 1, span); ws.mergeCells(2, 1, 2, span);
    Object.assign(ws.getCell(1, 1), { value: text }); ws.getCell(1, 1).font = { name: FONT, size: 14, bold: true, color: { argb: HEAD } };
    Object.assign(ws.getCell(2, 1), { value: sub }); ws.getCell(2, 1).font = { name: FONT, size: 10, italic: true, color: { argb: MUTED } };
    ws.getRow(1).height = 26;
  }
  function header(ws, row, labels) {
    labels.forEach((l, i) => { const c = ws.getCell(row, i + 1); c.value = l; c.font = { name: FONT, bold: true, color: { argb: "FFFFFFFF" } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEAD } }; c.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; c.border = box; });
    ws.getRow(row).height = 32;
  }
  function put(ws, r, c, v, o = {}) {
    const x = ws.getCell(r, c); x.value = v ?? null; x.font = { name: FONT, bold: !!o.bold, color: { argb: o.color || INK } };
    x.alignment = { horizontal: o.align || "center", vertical: "middle", wrapText: true }; x.border = box;
    if (o.fill) x.fill = { type: "pattern", pattern: "solid", fgColor: { argb: o.fill } };
    if (o.fmt) x.numFmt = o.fmt; if (o.list) x.dataValidation = { type: "list", allowBlank: true, formulae: [o.list], showErrorMessage: false };
    return x;
  }
  const inp = o => ({ fill: INPUT, ...o });

  async function build(R) {
    const wb = new ExcelJS.Workbook(); wb.creator = "سجل الحضور"; wb.created = new Date();
    const m = R?.meta || {}, cal = R?.calendar || { dutyDays: [0, 1, 2, 3, 4], holidays: [] }, rl = R?.rules || Core.PRESETS["أولية"];
    const subjects = R?.subjects || [], students = R?.students || [];
    const NS = Math.max(15, subjects.length + 5), NT = Math.max(60, students.length + 20);
    const SUB0 = 5, STU0 = 5;

    /* ---- اقرأني ---- */
    const rd = wb.addWorksheet("اقرأني", { views: [{ rightToLeft: true }] }); rd.columns = [{ width: 110 }];
    const lines = [
      ["نموذج سجل الحضور والغياب الجامعي", "t"],
      ["صالح لأي قسم وأي مرحلة: الدراسات الأولية (صباحي / مسائي) والدراسات العليا (دبلوم عالي / ماجستير / دكتوراه).", ""],
      ["الخانات الصفراء هي فقط ما يُملأ. الخانات البيضاء محسوبة بمعادلات.", "b"],
      ["1) ورقة «الإعدادات»: الجامعة والكلية والقسم، نوع الدراسة والمرحلة والشعبة، الفصل وتواريخه، قواعد الإنذار، وأيام الدوام (نعم/لا).", ""],
      ["2) ورقة «العطل»: العطل الرسمية والانقطاعات (تُخصم في التطبيق تلقائياً؛ في Excel خفّض «عدد الأسابيع» إن أردت خصمها).", ""],
      ["3) ورقة «المواد»: اسم المادة والتدريسي، ومواعيدها في الأسبوع (حتى 3 مواعيد: اليوم ومن وإلى). إن لم يكن لها جدول اكتب ساعات المحاضرة وعدد المحاضرات في الأسبوع.", ""],
      ["4) ورقة «الطلبة»: الاسم، وتاريخ بدء الدوام (للملتحق متأخراً)، والمواد المستوفاة أو المعفى منها (اكتب أسماءها مفصولة بفاصلة).", ""],
      ["5) ورقة «الحضور»: سطر لكل طالب في كل محاضرة: التاريخ، المادة، الطالب، الحالة (حاضر / غائب / مجاز). اليوم والساعات تُملأ تلقائياً.", ""],
      ["6) ورقة «الملخص»: ساعات الغياب والساعات المقررة والنسبة لكل طالب في كل مادة، والنسبة الكلية والحالة (ضمن الحد / إنذار أولي / إنذار ثانٍ / حرمان).", ""],
      ["الغياب في يوم ليس يوم دوام لا يُحسب. الغياب بعذر (مجاز) لا يُحسب إلا إذا اخترت ذلك في الإعدادات.", ""],
      ["التطبيق: https://ghraoy.github.io/attendance/ — استورد هذا الملف من «إدخال البيانات» فيُنشأ السجل كاملاً، ويمكنك تصديره بنفس الشكل.", ""]
    ];
    lines.forEach(([t, k], i) => { const c = rd.getCell(i + 1, 1); c.value = t; c.alignment = { horizontal: "right", vertical: "middle", wrapText: true }; c.font = { name: FONT, size: k === "t" ? 15 : 11, bold: k !== "", color: { argb: k === "t" ? HEAD : INK } }; rd.getRow(i + 1).height = k === "t" ? 30 : 22; });

    /* ---- الإعدادات ---- */
    const st = sheet(wb, "الإعدادات", [44, 34, 52]);
    title(st, "إعدادات السجل", "املأ الخانات الصفراء. القوائم المنسدلة تظهر عند الضغط على الخانة.", 3);
    header(st, 3, ["البند", "القيمة", "ملاحظة"]);
    const rows = [
      ["university", "الجامعة", m.university, ""], ["college", "الكلية / المعهد", m.college, ""], ["department", "القسم", m.department, ""],
      ["level", "نوع الدراسة (أولية / عليا)", m.level, "", '"أولية,عليا"'], ["degree", "الشهادة (للدراسات العليا)", m.degree, "دبلوم عالي / ماجستير / دكتوراه", '"دبلوم عالي,ماجستير,دكتوراه"'],
      ["shift", "الدراسة (صباحي / مسائي) — للأولية", m.shift, "", '"صباحي,مسائي"'], ["stage", "المرحلة / السنة", m.stage, "الأولى، الثانية … أو السنة التحضيرية"], ["section", "الشعبة", m.section, "اختياري"],
      ["system", "النظام الدراسي", m.system || "فصلي", "", '"فصلي,سنوي,بولونيا"'], ["semester", "الفصل الدراسي", m.semester || "الأول", "", '"الأول,الثاني,الصيفي,سنوي"'], ["year", "العام الدراسي", m.year, "مثال: 2026-2027"],
      ["start", "تاريخ بداية الفصل", ymdToDate(cal.start), "حسب التقويم الجامعي", null, "yyyy/mm/dd"], ["end", "تاريخ نهاية الفصل", ymdToDate(cal.end), "آخر يوم دوام قبل الامتحانات", null, "yyyy/mm/dd"],
      ["weeks", "عدد الأسابيع (اتركه فارغاً ليُحسب)", +cal.weeks || null, "اكتب رقماً لتجاوز الحساب التلقائي"],
      ["weeksUsed", "عدد الأسابيع المعتمد", { formula: `IF(N(B17)>0,B17,IF(AND(ISNUMBER(B15),ISNUMBER(B16)),ROUNDUP((B16-B15+1)/7,0),0))` }, "محسوب"],
      ["basis", "أساس حساب النسبة", rl.basis === "subject" ? "لكل مادة على حدة" : "من مجموع كل المواد", "", '"من مجموع كل المواد,لكل مادة على حدة"'],
      ["warn1", "الإنذار الأولي (%)", rl.warn1, "نسبة من الساعات المقررة"], ["warn2", "الإنذار الثاني (%)", rl.warn2, ""], ["warn3", "الحرمان / الرسوب بالغياب (%)", rl.warn3 || 0, "0 = لا يوجد"],
      ["countExcused", "احتساب الغياب بعذر", rl.countExcused ? "نعم" : "لا", "مجاز = غياب بعذر رسمي", '"نعم,لا"']
    ];
    for (const [k, label, v, note, list, fmt] of rows) {
      const r = S_ROWS[k]; put(st, r, 1, label, { bold: true, align: "right", fill: SUBH });
      put(st, r, 2, v, k === "weeksUsed" ? { bold: true } : inp({ list, fmt })); put(st, r, 3, note, { align: "right", color: MUTED });
    }
    DAYS.forEach((d, i) => { const r = S_ROWS.day0 + i; put(st, r, 1, "دوام يوم " + d, { bold: true, align: "right", fill: SUBH }); put(st, r, 2, cal.dutyDays.includes(i) ? "نعم" : "لا", inp({ list: '"نعم,لا"' })); put(st, r, 3, i === 0 ? "الغياب في يوم «لا» لا يُحسب" : "", { align: "right", color: MUTED }); });

    /* ---- العطل ---- */
    const hs = sheet(wb, "العطل", [16, 16, 40]);
    title(hs, "العطل الرسمية والانقطاعات", "تُخصم أيامها من الساعات المقررة في التطبيق.", 3);
    header(hs, 4, ["من تاريخ", "إلى تاريخ", "المناسبة"]);
    for (let i = 0; i < Math.max(15, cal.holidays.length + 5); i++) { const h = cal.holidays[i] || {}; put(hs, 5 + i, 1, ymdToDate(h.from), inp({ fmt: "yyyy/mm/dd" })); put(hs, 5 + i, 2, ymdToDate(h.to), inp({ fmt: "yyyy/mm/dd" })); put(hs, 5 + i, 3, h.name || null, inp({ align: "right" })); }

    /* ---- المواد ---- */
    const ms = sheet(wb, "المواد", [5, 28, 10, 24, 12, 10, 8, 8, 10, 8, 8, 10, 8, 8, 11, 11, 11, 11]);
    title(ms, "المواد والجدول الأسبوعي", "لكل مادة حتى 3 مواعيد أسبوعية. إن لم يوجد جدول اكتب ساعات المحاضرة وعدد المحاضرات أسبوعياً (آخر عمودين).", 18);
    header(ms, 4, ["ت", "اسم المادة", "الرمز", "التدريسي", "النوع", "اليوم 1", "من", "إلى", "اليوم 2", "من", "إلى", "اليوم 3", "من", "إلى", "ساعات الأسبوع", "ساعات الفصل", "ساعات المحاضرة", "محاضرات/أسبوع (بلا جدول)"]);
    const dayList = `"${DAYS.join(",")}"`;
    for (let i = 0; i < NS; i++) {
      const r = SUB0 + i, s = subjects[i], sl = s?.slots || [];
      put(ms, r, 1, i + 1); put(ms, r, 2, s?.name || null, inp({ align: "right" })); put(ms, r, 3, s?.code || null, inp()); put(ms, r, 4, s?.teacher || null, inp({ align: "right" }));
      put(ms, r, 5, s ? s.kind || "نظري" : null, inp({ list: '"نظري,عملي,نظري وعملي,سريري,حلقة نقاشية"' }));
      for (let j = 0; j < 3; j++) { const x = sl[j]; put(ms, r, 6 + 3 * j, x ? DAYS[x.day] : null, inp({ list: dayList })); put(ms, r, 7 + 3 * j, x ? hmToFrac(x.from) : null, inp({ fmt: "hh:mm" })); put(ms, r, 8 + 3 * j, x ? hmToFrac(x.to) : null, inp({ fmt: "hh:mm" })); }
      put(ms, r, 15, { formula: `IF(B${r}="","",IF((N(H${r})-N(G${r}))+(N(K${r})-N(J${r}))+(N(N${r})-N(M${r}))>0,ROUND(((N(H${r})-N(G${r}))+(N(K${r})-N(J${r}))+(N(N${r})-N(M${r})))*24,2),N(Q${r})*MAX(1,N(R${r}))))` }, { bold: true });
      put(ms, r, 16, { formula: `IF(B${r}="","",O${r}*${ref("weeksUsed")})` }, { bold: true });
      put(ms, r, 17, s ? (sl[0] ? Core.slotHours(sl[0]) : +s.hours || 2) : null, inp());
      put(ms, r, 18, s && !sl.length ? +s.perWeek || 1 : null, inp());
    }

    /* ---- الطلبة ---- */
    const ts = sheet(wb, "الطلبة", [5, 34, 16, 44, 24, 14]);
    title(ts, "الطلبة", "تاريخ بدء الدوام: للملتحق بعد بداية الفصل. المواد المستوفاة: اكتب أسماءها كما في ورقة المواد مفصولة بفاصلة.", 6);
    header(ts, 4, ["ت", "اسم الطالب", "تاريخ بدء الدوام", "المواد المستوفاة / المعفى منها", "ملاحظات", "أسابيع الدوام المطلوبة"]);
    for (let i = 0; i < NT; i++) {
      const r = STU0 + i, t = students[i];
      put(ts, r, 1, i + 1); put(ts, r, 2, t?.name || null, inp({ align: "right" }));
      put(ts, r, 3, t && t.start && t.start !== cal.start ? ymdToDate(t.start) : null, inp({ fmt: "yyyy/mm/dd" }));
      put(ts, r, 4, t ? (t.exempt || []).map(id => subjects.find(s => s.id === id)?.name).filter(Boolean).join("، ") || null : null, inp({ align: "right" }));
      put(ts, r, 5, null, inp({ align: "right" }));
      put(ts, r, 6, { formula: `IF(B${r}="","",IF(OR(NOT(ISNUMBER(C${r})),C${r}<=${ref("start")}),${ref("weeksUsed")},MAX(0,MIN(${ref("weeksUsed")},ROUNDUP((${ref("end")}-C${r}+1)/7,0)))))` }, { bold: true });
    }

    /* ---- الحضور ---- */
    const as = sheet(wb, "الحضور", [13, 11, 28, 32, 12, 10, 14, 30]);
    title(as, "سجل الحضور", "سطر لكل طالب في كل محاضرة. الحالة: حاضر / غائب / مجاز. اليوم والساعات وساعات الغياب المحسوبة تُملأ تلقائياً.", 8);
    header(as, 4, ["التاريخ", "اليوم", "المادة", "الطالب", "الحالة", "الساعات", "ساعات الغياب المحسوبة", "ملاحظات"]);
    const entries = [];
    if (R) for (const se of [...R.sessions].sort((a, b) => a.date.localeCompare(b.date))) {
      const su = subjects.find(s => s.id === se.subject); if (!su) continue;
      for (const t of students) {
        const x = Core.status(R, se, t), absent = (se.absent || []).includes(t.id), excused = (se.excused || []).includes(t.id);
        // write what happened; the sheet's own formulas decide what counts (duty days, start date, excused rule)
        const txt = x === "E" ? "مستوفية" : excused ? "مجاز" : absent ? (t.start && se.date < t.start ? "غير مطالب" : "غائب") : x === "N" ? "غير مطالب" : "حاضر";
        entries.push([se.date, su.name, t.name, txt, Core.sessionHours(R, se), se.note || ""]);
      }
    }
    const subList = `'المواد'!$B$${SUB0}:$B$${SUB0 + NS - 1}`, stuList = `'الطلبة'!$B$${STU0}:$B$${STU0 + NT - 1}`;
    const NA = entries.length + 500;
    for (let i = 0; i < NA; i++) {
      const r = 5 + i, e = entries[i];
      put(as, r, 1, e ? ymdToDate(e[0]) : null, inp({ fmt: "yyyy/mm/dd" }));
      put(as, r, 2, { formula: `IF(A${r}="","",CHOOSE(WEEKDAY(A${r}),"الأحد","الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت"))` });
      put(as, r, 3, e ? e[1] : null, inp({ align: "right", list: subList }));
      put(as, r, 4, e ? e[2] : null, inp({ align: "right", list: stuList }));
      put(as, r, 5, e ? e[3] : null, inp({ list: '"حاضر,غائب,مجاز,غير مطالب,مستوفية"' }));
      put(as, r, 6, e ? e[4] : { formula: `IF(C${r}="","",IFERROR(VLOOKUP(C${r},'المواد'!$B:$Q,16,0),2))` }, e ? inp() : {});
      // counted only when the student is required: on a duty day and not before his start date
      put(as, r, 7, { formula: `IF(D${r}="","",IF(AND(OR(E${r}="غائب",AND(E${r}="مجاز",${ref("countExcused")}="نعم")),IFERROR(VLOOKUP("دوام يوم "&B${r},'الإعدادات'!$A:$B,2,0),"نعم")="نعم",IFERROR(A${r}>=IF(ISNUMBER(VLOOKUP(D${r},'الطلبة'!$B:$C,2,0)),VLOOKUP(D${r},'الطلبة'!$B:$C,2,0),0),TRUE)),N(F${r}),0))` }, { bold: true });
      put(as, r, 8, e ? e[5] || null : null, inp({ align: "right" }));
    }
    as.addConditionalFormatting({ ref: `E5:E${4 + NA}`, rules: [
      { type: "cellIs", operator: "equal", formulae: ['"غائب"'], style: { font: { color: { argb: "FFB3262F" }, bold: true }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFBE5E6" } } } },
      { type: "cellIs", operator: "equal", formulae: ['"مجاز"'], style: { font: { color: { argb: "FF2A5D8F" } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFE2EDF8" } } } }] });

    /* ---- الملخص ---- */
    const sm = sheet(wb, "الملخص", [5, 30, ...Array(NS).fill(11), 13, 13, 11, 12, 16]);
    const lastSub = 2 + NS, cTotA = lastSub + 1, cTotP = lastSub + 2, cRate = lastSub + 3, cMax = lastSub + 4, cState = lastSub + 5;
    const col = c => sm.getColumn(c).letter;
    title(sm, "ملخص الحضور والغياب", R ? Core.registerName(R) : "يُملأ تلقائياً من الأوراق الأخرى", cState);
    const subHdr = j => ({ formula: `IF('المواد'!$B$${SUB0 + j}="","",'المواد'!$B$${SUB0 + j})` });
    const block = (top, caption) => {
      sm.mergeCells(top, 1, top, cState); put(sm, top, 1, caption, { bold: true, align: "right", fill: SUBH });
      header(sm, top + 1, ["ت", "اسم الطالب"]); for (let j = 0; j < NS; j++) { const c = put(sm, top + 1, 3 + j, subHdr(j), { bold: true, fill: HEAD, color: "FFFFFFFF" }); }
      return top + 2;
    };
    const A0 = block(4, "ساعات الغياب المحسوبة لكل طالب في كل مادة");
    ["مجموع ساعات الغياب", "مجموع الساعات المقررة", "النسبة الكلية", "أعلى نسبة في مادة", "الحالة"].forEach((h, i) => put(sm, A0 - 1, cTotA + i, h, { bold: true, fill: HEAD, color: "FFFFFFFF" }));
    const B0 = block(A0 + NT + 2, "الساعات المقررة لكل طالب في كل مادة (ساعات الأسبوع × أسابيع دوام الطالب؛ 0 للمستوفي)");
    const C0 = block(B0 + NT + 2, "نسبة الغياب في كل مادة");
    for (let i = 0; i < NT; i++) {
      const name = `'الطلبة'!$B$${STU0 + i}`;
      for (const top of [A0, B0, C0]) { put(sm, top + i, 1, { formula: `IF(${name}="","",${i + 1})` }); put(sm, top + i, 2, { formula: `IF(${name}="","",${name})` }, { align: "right" }); }
      for (let j = 0; j < NS; j++) {
        const c = col(3 + j), h = `${c}$${A0 - 1}`, nB = `$B${A0 + i}`;
        put(sm, A0 + i, 3 + j, { formula: `IF(OR(${nB}="",${h}=""),"",SUMIFS('الحضور'!$G:$G,'الحضور'!$D:$D,${nB},'الحضور'!$C:$C,${h}))` });
        put(sm, B0 + i, 3 + j, { formula: `IF(OR(${nB}="",${h}=""),"",IF(ISNUMBER(SEARCH(${h},VLOOKUP(${nB},'الطلبة'!$B:$F,3,0)&"")),0,IFERROR(VLOOKUP(${h},'المواد'!$B:$O,14,0),0)*IFERROR(VLOOKUP(${nB},'الطلبة'!$B:$F,5,0),0)))` });
        put(sm, C0 + i, 3 + j, { formula: `IF(OR(${nB}="",${h}=""),"",IF(N(${c}${B0 + i})=0,"مستوفية",N(${c}${A0 + i})/${c}${B0 + i}))` }, { fmt: "0.0%" });
      }
      const r = A0 + i, first = col(3), last = col(lastSub);
      put(sm, r, cTotA, { formula: `IF($B${r}="","",SUM(${first}${r}:${last}${r}))` }, { bold: true });
      put(sm, r, cTotP, { formula: `IF($B${r}="","",SUM(${first}${B0 + i}:${last}${B0 + i}))` }, { bold: true });
      put(sm, r, cRate, { formula: `IF(OR($B${r}="",N(${col(cTotP)}${r})=0),"",${col(cTotA)}${r}/${col(cTotP)}${r})` }, { bold: true, fmt: "0.0%" });
      put(sm, r, cMax, { formula: `IF($B${r}="","",MAX(${first}${C0 + i}:${last}${C0 + i}))` }, { fmt: "0.0%" });
      const rate = `IF(${ref("basis")}="لكل مادة على حدة",N(${col(cMax)}${r}),N(${col(cRate)}${r}))*100`;
      put(sm, r, cState, { formula: `IF($B${r}="","",IF(AND(N(${ref("warn3")})>0,${rate}>=${ref("warn3")}),"حرمان",IF(${rate}>=${ref("warn2")},"إنذار ثانٍ",IF(${rate}>=${ref("warn1")},"إنذار أولي","ضمن الحد"))))` }, { bold: true });
    }
    const stRef = `${col(cState)}${A0}:${col(cState)}${A0 + NT - 1}`;
    sm.addConditionalFormatting({ ref: stRef, rules: [
      { type: "cellIs", operator: "equal", formulae: ['"حرمان"'], style: { font: { color: { argb: "FFFFFFFF" }, bold: true }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFB3262F" } } } },
      { type: "cellIs", operator: "equal", formulae: ['"إنذار ثانٍ"'], style: { font: { color: { argb: "FFB3262F" }, bold: true }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFBE5E6" } } } },
      { type: "cellIs", operator: "equal", formulae: ['"إنذار أولي"'], style: { font: { color: { argb: "FF9A5B00" }, bold: true }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFDF0D8" } } } }] });
    sm.views = [{ rightToLeft: true, state: "frozen", xSplit: 2, ySplit: A0 - 1 }];
    wb.views = [{ activeTab: R ? 6 : 1 }];
    return wb.xlsx.writeBuffer();
  }
  return { build };
})();
