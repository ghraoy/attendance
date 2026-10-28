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
    const wb = new ExcelJS.Workbook(); wb.creator = "م.م. سجاد مهدي الغراوي"; wb.lastModifiedBy = "سجل الحضور الجامعي"; wb.company = "سجل الحضور الجامعي — برمجة وتصميم م.م. سجاد مهدي الغراوي"; wb.created = new Date();
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
      ["5) أوراق «كشف 1، كشف 2 …»: كشف لكل مادة بنفس ترتيب ورقة المواد. الطلبة في الصفوف وتواريخ المحاضرات في الأعمدة، تُولَّد تلقائياً من جدول المادة وتاريخي الفصل. اكتب غ للغائب، م للمجاز، ح للحاضر (أو اتركها فارغة).", ""],
      ["6) ورقة «الملخص»: ساعات الغياب والساعات المقررة والنسبة لكل طالب في كل مادة، والنسبة الكلية والحالة (ضمن الحد / إنذار أولي / إنذار ثانٍ / حرمان).", ""],
      ["الغياب في يوم ليس يوم دوام لا يُحسب. الغياب بعذر (مجاز) لا يُحسب إلا إذا اخترت ذلك في الإعدادات.", ""],
      ["التطبيق: https://ghraoy.github.io/attendance/ — استورد هذا الملف من «إدخال البيانات» فيُنشأ السجل كاملاً، ويمكنك تصديره بنفس الشكل.", ""],
      ["", ""],
      ["برمجة وتصميم: م.م. سجاد مهدي الغراوي — © 2026 جميع الحقوق محفوظة. الاستخدام مجاني للأغراض التعليمية والإدارية مع بقاء هذا الإسناد.", "c"]
    ];
    lines.forEach(([t, k], i) => { const c = rd.getCell(i + 1, 1); c.value = t; c.alignment = { horizontal: "right", vertical: "middle", wrapText: true }; c.font = { name: FONT, size: k === "t" ? 15 : 11, bold: k !== "", color: { argb: k === "t" || k === "c" ? HEAD : INK } }; rd.getRow(i + 1).height = k === "t" ? 30 : 22; });

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

    /* ---- كشوف الحضور: ورقة لكل مادة (الكشف j ↔ صف المادة j في ورقة المواد) ---- */
    const DAYS_ARR = `{${DAYS.map(d => `"${d}"`).join(",")}}`;
    const CHOOSE_DAY = x => `CHOOSE(WEEKDAY(${x}),${DAYS.map(d => `"${d}"`).join(",")})`;
    const KS = []; // per subject: {name, colTotHours, colPlanned}
    const NK = R ? Math.max(subjects.length, 1) : NS;
    for (let j = 0; j < NK; j++) {
      const R0 = SUB0 + j, su = subjects[j], sheetName = `كشف ${j + 1}`;
      // lecture dates: held sessions + remaining timetable dates (filled export) or formulas (blank)
      let dates = [];
      if (R && su) {
        const held = R.sessions.filter(x => x.subject === su.id).map(x => x.date);
        const last = held.length ? held.reduce((m, d) => d > m ? d : m) : null;
        const future = [];
        if ((su.slots || []).length && cal.start && cal.end) for (let d = Core.addDays(last || Core.addDays(cal.start, -1), 1); d <= cal.end; d = Core.addDays(d, 1)) if (!Core.isHoliday(R, d) && su.slots.some(x => +x.day === Core.dow(d))) future.push(d);
        dates = [...new Set([...held, ...future])].sort();
      }
      const NC = Math.max(40, dates.length + 6), C0 = 4, CL = C0 + NC - 1;
      const ks = wb.addWorksheet(sheetName, { views: [{ rightToLeft: true, state: "frozen", xSplit: 3, ySplit: 7 }], properties: { defaultRowHeight: 18 } });
      ks.columns = [{ width: 5 }, { width: 28 }, { width: 11 }, ...Array(NC).fill({ width: 6.5 }), { width: 9 }, { width: 10 }, { width: 8 }, { width: 10 }, { width: 10 }];
      ks.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
      const L = c => ks.getColumn(c).letter, fc = L(C0), lc = L(CL);
      const tA = CL + 1, tH = CL + 2, tX = CL + 3, tP = CL + 4, tR = CL + 5;
      // row 3 (hidden helpers): subject name, slot weekdays, slot hours, default hours, term dates
      const nameF = `IF('المواد'!$B$${R0}="","",'المواد'!$B$${R0})`;
      put(ks, 3, 1, "مساعد"); put(ks, 3, 3, su ? { formula: nameF, result: su.name } : { formula: nameF });
      [["F", 4], ["I", 5], ["L", 6]].forEach(([c, k]) => put(ks, 3, k, { formula: `IFERROR(MATCH('المواد'!$${c}$${R0},${DAYS_ARR},0),"")` }));
      [["G", "H", 7], ["J", "K", 8], ["M", "N", 9]].forEach(([f, t, k]) => put(ks, 3, k, { formula: `IFERROR(ROUND((N('المواد'!$${t}$${R0})-N('المواد'!$${f}$${R0}))*24,2),0)` }));
      put(ks, 3, 10, { formula: `IF(N('المواد'!$Q$${R0})>0,'المواد'!$Q$${R0},2)` });
      put(ks, 3, 11, { formula: ref("start") }); put(ks, 3, 12, { formula: ref("end") });
      ks.getRow(3).hidden = true;
      ks.mergeCells(1, 1, 1, tR); ks.mergeCells(2, 1, 2, tR);
      put(ks, 1, 1, { formula: `"كشف حضور مادة: "&IF($C$3="","(اكتب اسم المادة في الصف ${j + 1} من ورقة المواد)",$C$3)&"  —  "&IFERROR('المواد'!$D$${R0},"")`, result: su ? `كشف حضور مادة: ${su.name}  —  ${su.teacher || ""}` : undefined }, { bold: true, align: "right", color: HEAD });
      ks.getCell(1, 1).font = { name: FONT, size: 14, bold: true, color: { argb: HEAD } };
      put(ks, 2, 1, "اكتب في خانة كل طالب: غ = غائب · م = مجاز (غياب بعذر) · ح = حاضر (أو اتركها فارغة). التواريخ تُولَّد من جدول المادة والتقويم، ويمكنك الكتابة فوقها. العمود الرمادي = عطلة أو ليس يوم دوام.", { align: "right", color: MUTED });
      ks.getRow(1).height = 26; ks.getRow(2).height = 30;
      // rows 4-7: date / weekday / hours / duty-day
      header(ks, 4, ["ت", "اسم الطالب", "يبدأ دوامه"]);
      put(ks, 5, 2, "اليوم", { bold: true, fill: SUBH }); put(ks, 6, 2, "ساعات المحاضرة", { bold: true, fill: SUBH }); put(ks, 7, 2, "يوم دوام؟", { bold: true, fill: SUBH });
      const holiday = x => `COUNTIFS('العطل'!$A:$A,"<="&${x},'العطل'!$B:$B,">="&${x})>0`;
      const next = x => `${x}+MIN(IF($D$3="",99,MOD($D$3-WEEKDAY(${x})-1,7)+1),IF($E$3="",99,MOD($E$3-WEEKDAY(${x})-1,7)+1),IF($F$3="",99,MOD($F$3-WEEKDAY(${x})-1,7)+1))`;
      for (let k = 0; k < NC; k++) {
        const c = C0 + k, X = L(c), P = k ? L(c - 1) : null, d = dates[k];
        const f = k === 0 ? `IF(AND(ISNUMBER($K$3),COUNT($D$3:$F$3)>0),IF(${next("($K$3-1)")}>$L$3,"",${next("($K$3-1)")}),"")` : `IF(NOT(ISNUMBER(${P}4)),"",IF(COUNT($D$3:$F$3)=0,"",IF(${next(P + "4")}>$L$3,"",${next(P + "4")})))`;
        put(ks, 4, c, R ? (d ? ymdToDate(d) : null) : { formula: f }, { bold: true, fill: INPUT, fmt: "d/m", color: INK });
        ks.getCell(4, c).alignment = { horizontal: "center", vertical: "middle", textRotation: 90 };
        put(ks, 5, c, { formula: `IF(NOT(ISNUMBER(${X}4)),"",IF(${holiday(X + "4")},"عطلة",${CHOOSE_DAY(X + "4")}))` }, { color: MUTED });
        ks.getCell(5, c).alignment = { horizontal: "center", vertical: "middle", textRotation: 90 };
        const se = R && su && d ? R.sessions.find(x => x.subject === su.id && x.date === d) : null;
        put(ks, 6, c, se && se.hours ? +se.hours : { formula: `IF(NOT(ISNUMBER(${X}4)),0,IF(${holiday(X + "4")},0,IF(AND(WEEKDAY(${X}4)=N($D$3),$G$3>0),$G$3,IF(AND(WEEKDAY(${X}4)=N($E$3),$H$3>0),$H$3,IF(AND(WEEKDAY(${X}4)=N($F$3),$I$3>0),$I$3,$J$3)))))` }, { fmt: "0.##;;;" });
        put(ks, 7, c, { formula: `IF(NOT(ISNUMBER(${X}4)),"",IF(${holiday(X + "4")},"لا",IFERROR(VLOOKUP("دوام يوم "&${CHOOSE_DAY(X + "4")},'الإعدادات'!$A:$B,2,0),"نعم")))` }, { color: MUTED });
      }
      ["الغيابات", "ساعات الغياب المحسوبة", "مجاز", "الساعات المقررة", "النسبة"].forEach((h, i) => put(ks, 4, tA + i, h, { bold: true, fill: HEAD, color: "FFFFFFFF" }));
      for (const rr of [5, 6, 7]) for (let i = 0; i < 5; i++) put(ks, rr, tA + i, null, { fill: SUBH });
      // student rows
      const S0 = 8;
      for (let i = 0; i < NT; i++) {
        const r = S0 + i, t = students[i], nm = `'الطلبة'!$B$${STU0 + i}`;
        put(ks, r, 1, { formula: `IF($B${r}="","",${i + 1})` });
        put(ks, r, 2, t ? { formula: `IF(${nm}="","",${nm})`, result: t.name } : { formula: `IF(${nm}="","",${nm})` }, { align: "right" });
        put(ks, r, 3, { formula: `IF($B${r}="","",IFERROR(IF(ISNUMBER(VLOOKUP($B${r},'الطلبة'!$B:$C,2,0)),VLOOKUP($B${r},'الطلبة'!$B:$C,2,0),${ref("start")}),${ref("start")}))` }, { fmt: "d/m/yyyy", color: MUTED });
        for (let k = 0; k < NC; k++) {
          const d = dates[k]; let v = null;
          if (R && su && t && d) { const se = R.sessions.find(x => x.subject === su.id && x.date === d); if (se) { if ((t.exempt || []).includes(su.id) || (t.start && d < t.start)) v = null; else v = (se.excused || []).includes(t.id) ? "م" : (se.absent || []).includes(t.id) ? "غ" : "ح"; } }
          put(ks, r, C0 + k, v, { fill: INPUT, list: '"غ,م,ح"' });
        }
        const rng = `${fc}${r}:${lc}${r}`, hrs = `$${fc}$6:$${lc}$6`, duty = `($${fc}$7:$${lc}$7="نعم")`, after = `($${fc}$4:$${lc}$4>=$C${r})`;
        const exempt = `ISNUMBER(SEARCH($C$3,IFERROR(VLOOKUP($B${r},'الطلبة'!$B:$D,3,0),"")&""))`;
        put(ks, r, tA, { formula: `IF($B${r}="","",COUNTIF(${rng},"غ"))` }, { bold: true });
        put(ks, r, tH, { formula: `IF($B${r}="","",SUMPRODUCT((${rng}="غ")*${hrs}*${duty}*${after})+IF(${ref("countExcused")}="نعم",SUMPRODUCT((${rng}="م")*${hrs}*${duty}*${after}),0))` }, { bold: true });
        put(ks, r, tX, { formula: `IF($B${r}="","",COUNTIF(${rng},"م"))` });
        // planned: with a timetable, the sum of this grid's lecture hours from the student's start (holidays are 0); otherwise weekly hours × weeks
        put(ks, r, tP, { formula: `IF(OR($B${r}="",$C$3=""),"",IF(${exempt},0,IF(COUNT($D$3:$F$3)>0,SUMPRODUCT(${hrs}*${after}),N('المواد'!$O$${R0})*IFERROR(VLOOKUP($B${r},'الطلبة'!$B:$F,5,0),0))))` });
        put(ks, r, tR, { formula: `IF(OR($B${r}="",$C$3=""),"",IF(N(${L(tP)}${r})=0,"مستوفية",N(${L(tH)}${r})/${L(tP)}${r}))` }, { bold: true, fmt: "0.0%" });
      }
      const grid = `${fc}${S0}:${lc}${S0 + NT - 1}`;
      ks.addConditionalFormatting({ ref: grid, rules: [
        { type: "cellIs", operator: "equal", formulae: ['"غ"'], style: { font: { color: { argb: "FFB3262F" }, bold: true }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFBE5E6" } } } },
        { type: "cellIs", operator: "equal", formulae: ['"م"'], style: { font: { color: { argb: "FF2A5D8F" }, bold: true }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFE2EDF8" } } } },
        { type: "expression", formulae: [`${fc}$7="لا"`], style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFE4E7E7" } } } },
        { type: "expression", formulae: [`AND($B${S0}<>"",$C$3<>"",ISNUMBER(SEARCH($C$3,IFERROR(VLOOKUP($B${S0},'الطلبة'!$B:$D,3,0),"")&"")))`], style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFE2EDF8" } } } }] });
      KS.push({ sheetName, colH: tH - 1, colP: tP - 1 }); // VLOOKUP column offsets from column B
    }

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
        const K = KS[j];
        put(sm, A0 + i, 3 + j, K ? { formula: `IF(OR(${nB}="",${h}=""),"",IFERROR(N(VLOOKUP(${nB},'${K.sheetName}'!$B:$${sm.getColumn(K.colH + 1).letter},${K.colH},0)),0))` } : "");
        put(sm, B0 + i, 3 + j, K ? { formula: `IF(OR(${nB}="",${h}=""),"",IFERROR(N(VLOOKUP(${nB},'${K.sheetName}'!$B:$${sm.getColumn(K.colP + 1).letter},${K.colP},0)),0))` } : "");
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
    wb.views = [{ activeTab: R ? wb.worksheets.length - 1 : 1 }];
    return wb.xlsx.writeBuffer();
  }
  return { build };
})();
