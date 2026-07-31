/**
 * KPI スライド ジェネレーター（16:9, PowerPoint native charts）
 *
 * 使い方（1社1ファイル）:
 *   node build_deck.js data.json [out.pptx]
 *
 * 使い方（複数社を1つのPowerPointにまとめる。グループ会社を束ねて経営会議資料にする場合）:
 *   node build_deck.js bundle.json [out.pptx]
 *   bundle.json は { "companies": [data1, data2, ...] } の形。各要素は通常の data.json と同じ
 *   スキーマ。会社の数だけスライドが追加され（extraSection.layout==="page" の会社はさらに1枚
 *   追加）、1つの.pptxファイルとして書き出される。
 *
 * data.json のスキーマは references/design-system.md 末尾のサンプル、または
 * このリポジトリの sample.data.json を参照。会社名・グラフの数値・所見文は
 * 呼び出し側（Claude）が references/calc-rules.md のルールに従って
 * 事前に計算し、この JSON に詰めてから渡すこと。このスクリプト自身は集計を一切行わない
 * ——見た目とレイアウトの責務だけを持つ。
 *
 * kpis / charts は会社ごとに数が変わってよい（例: 受注データが無ければ受注高カードと
 * 受注高を含むチャートを外して3カード・2チャートにする。存在しないデータを埋めない）。
 *
 * extraSection（任意）は会社特有の指標（受注TOP10、稼働率、不良率など）を表現する汎用枠。
 * 中身は「表」か「箇条書き」のどちらかで自由に組み立てる（詳しくは design-system.md）。
 * extraSection.layout:
 *   - "column"（既定）: 1枚目の下段に3カラム目として収める（所見・追加セクション・コメント）
 *   - "page": 追加セクションだけで2枚目のスライドを作る（表の行数が多い・情報量が多いとき）
 *
 * dataNote（任意）は data.json 側の記録用フィールドとして残せるが、スライド上には表示しない
 * （2026-07-29のユーザー要望で「下部のデータ注記」欄を廃止したため）。
 *
 * 依存: pptxgenjs（プリインストールされていない環境がある。
 * require が失敗したら `npm install pptxgenjs` をこのスクリプトと同じ作業ディレクトリで実行）
 */
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

// 2026-07-30、ユーザー要望で暖色系（クレイ基調）から寒色系（ブルー×オレンジ基調）に変更。
// dataviz スキルの validate_palette.js で検証済み（[#2a78d6, #eb6834] を surface #F7F9FB /
// light モードで検証。CVD分離 ΔE 24.7、通常視ΔE 33.6、いずれも閾値を大きく上回りPASS）。
// 配色を再度変える場合は同じ手順（node scripts/validate_palette.js "<hex,hex>" --mode light
// --surface "<surface>"、dataviz スキル配下）で必ず再検証すること。
const COL = {
  ink: "1F2937",
  inkMuted: "52616B",
  inkFaint: "8A97A0",
  rule: "D9E2EA",
  ruleSoft: "EDF1F5",
  accent: "2A78D6",
  accentSoft: "E4EEFC",
  up: "006300",
  down: "D03B3B",
  sales: "2A78D6",
  orders: "EB6834",
  white: "F7F9FB",
};
const FONT_HEAD = "Noto Sans JP";
const FONT_BODY = "Noto Sans JP";

// 増加は "+"、減少は "▼" で表す。日本の会計資料では▲が「マイナス」を意味する慣習があり、
// 「増加=▲（上向き三角）」という表記は誤読されやすい（2026-07-29のユーザー指摘）。
// ▲は使わず、増加=+、減少=▼、変化なし=－ に統一する。
function fmtDelta(delta, isPt, decimals = 1) {
  if (delta === null || delta === undefined || Number.isNaN(delta)) return "－";
  const sym = delta < 0 ? "▼" : delta > 0 ? "+" : "－";
  const mag = Math.abs(delta).toFixed(decimals);
  return sym + mag + (isPt ? "pt" : "%");
}
function isBad(delta, higherIsBetter) {
  if (delta === null || delta === undefined) return false;
  return higherIsBetter ? delta < 0 : delta > 0;
}
function fmtValue(v, decimals) {
  return v.toLocaleString("ja-JP", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

function deriveKpiDisplay(k) {
  // k: {label, unit, decimals, isRatio, value, target, mom, avg3, avg3Prev, higherIsBetter}
  const dec = k.decimals ?? (k.isRatio ? 1 : 0);
  const out = { label: k.label, valueStr: fmtValue(k.value, dec), unit: k.unit };

  if (!k.isRatio) {
    const targetDelta = k.target != null ? ((k.value - k.target) / k.target) * 100 : null;
    const momDelta = ((k.value - k.mom) / k.mom) * 100;
    const avg3Delta = ((k.avg3 - k.avg3Prev) / k.avg3Prev) * 100;
    out.targetStr = fmtDelta(targetDelta, false);
    out.targetBad = isBad(targetDelta, k.higherIsBetter);
    out.momStr = fmtDelta(momDelta, false);
    out.momBad = isBad(momDelta, k.higherIsBetter);
    out.subStr = `3ヶ月平均 ${fmtValue(k.avg3, dec)}${k.unit}（前期間比 ${fmtDelta(avg3Delta, false)}）`;
  } else {
    const momDelta = k.value - k.mom;
    const avg3Delta = k.avg3 - k.avg3Prev;
    out.targetStr = "－";
    out.targetBad = false;
    out.momStr = fmtDelta(momDelta, true);
    out.momBad = isBad(momDelta, k.higherIsBetter);
    out.subStr = `3ヶ月平均 ${fmtValue(k.avg3, dec)}${k.unit}（前期間比 ${fmtDelta(avg3Delta, true)}）`;
  }
  return out;
}

function addHeader(slide, cfg, ML, MT, CW) {
  slide.addText(cfg.company, {
    x: ML, y: MT, w: 7.0, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 19, bold: true, color: COL.ink, margin: 0, valign: "top",
  });
  slide.addText(
    [
      { text: "対象月　", options: { color: COL.inkMuted } },
      { text: cfg.targetMonth, options: { color: COL.ink, bold: true } },
      { text: cfg.monthSuffix || "実績", options: { color: COL.inkMuted } },
    ],
    { x: ML + 6.5, y: MT + 0.01, w: CW - 6.5, h: 0.2, fontFace: FONT_BODY, fontSize: 10, align: "right", margin: 0 }
  );
  slide.addText(
    [
      { text: "作成日　", options: { color: COL.inkMuted } },
      { text: cfg.createdDate, options: { color: COL.ink } },
    ],
    { x: ML + 6.5, y: MT + 0.22, w: CW - 6.5, h: 0.2, fontFace: FONT_BODY, fontSize: 10, align: "right", margin: 0 }
  );
  slide.addShape("line", { x: ML, y: MT + 0.46, w: CW, h: 0, line: { color: COL.ink, width: 1.25 } });
}

// 追加セクションの中身（表 or 箇条書き）を指定した矩形内に描画する。
// slide2（ページ全面）でも1枚目の下段カラムでも同じ関数で描ける。
function addExtraContent(slide, ext, x, y, w, h, { fontScale = 1 } = {}) {
  if (ext.kind === "bullets" && ext.bullets && ext.bullets.length) {
    slide.addText(
      ext.bullets.map((b) => ({ text: b, options: { color: COL.ink, breakLine: true, bullet: { code: "25CF", indent: 10 * fontScale } } })),
      { x, y, w, h, fontFace: FONT_BODY, fontSize: 8.5 * fontScale, margin: 0, valign: "top", lineSpacingMultiple: 1.1, paraSpaceAfter: 3 }
    );
    return;
  }
  const t = ext.table;
  if (!t || !t.rows || !t.rows.length) return;
  const cols = t.columns;
  const fixedW = cols.reduce((s, c) => s + (c.width || 0), 0);
  const flexCols = cols.filter((c) => !c.width);
  const flexW = flexCols.length ? (w - fixedW) / flexCols.length : 0;
  const colW = cols.map((c) => c.width || flexW);

  const headerRow = cols.map((c) => ({
    text: c.label,
    options: { align: c.align || "left", bold: true, color: COL.white, fill: { color: COL.ink }, fontSize: 7.6 * fontScale },
  }));
  const bodyRows = t.rows.map((r) =>
    cols.map((c) => ({ text: String(r[c.key] ?? ""), options: { align: c.align || "left", color: COL.ink, fontSize: 7.4 * fontScale } }))
  );
  slide.addTable([headerRow, ...bodyRows], {
    x, y, w, h,
    fontFace: FONT_BODY,
    border: { type: "solid", color: COL.rule, pt: 0.5 },
    autoPage: false,
    colW,
    valign: "middle",
    rowH: h / (t.rows.length + 1),
  });
}

// 1社分のスライド（KPIカード＋グラフ＋所見＋追加セクション＋コメント欄。extraSection.layout==="page"
// の場合はさらに2枚目も）を pres に追加する。複数社を束ねるときはこれを会社の数だけ呼ぶ。
function addCompanySlides(pres, cfg) {
  const slide = pres.addSlide();
  slide.background = { color: COL.white };

  const ML = 0.4, MR = 0.4, MT = 0.26;
  const CW = 13.333 - ML - MR;

  const ext = cfg.extraSection || null;
  const extOnPage = ext && ext.layout === "page";
  const extInColumn = ext && !extOnPage;

  // ---------- header ----------
  addHeader(slide, cfg, ML, MT, CW);

  // ---------- KPI row（枚数は cfg.kpis の要素数に自動追従。存在しない指標は入れない） ----------
  const kpiY = MT + 0.54;
  const kpiH = 1.05;
  const kpiGap = 0.14;
  const kpiCount = cfg.kpis.length;
  const kpiW = (CW - (kpiCount - 1) * kpiGap) / kpiCount;

  cfg.kpis.forEach((kRaw, i) => {
    const k = deriveKpiDisplay(kRaw);
    const x = ML + i * (kpiW + kpiGap);
    slide.addShape("roundRect", {
      x, y: kpiY, w: kpiW, h: kpiH, rectRadius: 0.05,
      fill: { color: COL.accentSoft }, line: { type: "none" },
    });
    slide.addText(k.label, {
      x: x + 0.12, y: kpiY + 0.07, w: kpiW - 0.24, h: 0.17,
      fontFace: FONT_BODY, fontSize: 9, bold: true, color: COL.inkMuted, margin: 0,
    });
    slide.addText(
      [
        { text: k.valueStr, options: { fontSize: 17, bold: true, color: COL.ink } },
        { text: " " + k.unit, options: { fontSize: 9, bold: true, color: COL.inkMuted } },
      ],
      { x: x + 0.12, y: kpiY + 0.23, w: kpiW - 0.24, h: 0.28, fontFace: FONT_BODY, margin: 0, valign: "top" }
    );
    slide.addText(
      [
        { text: "目標比 ", options: { color: k.targetBad ? COL.down : COL.inkFaint, bold: k.targetBad } },
        { text: k.targetStr, options: { color: k.targetBad ? COL.down : COL.inkFaint, bold: k.targetBad } },
        { text: "　", options: {} },
        { text: "前月比 ", options: { color: k.momBad ? COL.down : COL.up, bold: true } },
        { text: k.momStr, options: { color: k.momBad ? COL.down : COL.up, bold: true } },
      ],
      { x: x + 0.12, y: kpiY + 0.53, w: kpiW - 0.24, h: 0.18, fontFace: FONT_BODY, fontSize: 8.2, margin: 0 }
    );
    slide.addText(k.subStr, {
      x: x + 0.12, y: kpiY + 0.73, w: kpiW - 0.24, h: 0.3,
      fontFace: FONT_BODY, fontSize: 7.6, color: COL.inkFaint, margin: 0, valign: "top",
    });
  });

  // ---------- chart row（枚数は cfg.charts の要素数に自動追従） ----------
  const chartY = kpiY + kpiH + 0.12;
  const chartH = 1.62;
  const chartGap = 0.18;
  const chartCount = cfg.charts.length;
  const chartW = (CW - (chartCount - 1) * chartGap) / chartCount;

  const axisOpts = {
    catAxisLabelFontFace: FONT_BODY, catAxisLabelFontSize: 6, catAxisLabelColor: COL.inkFaint,
    catAxisLineColor: COL.rule, catAxisLineShow: false,
    valAxisLabelFontFace: FONT_BODY, valAxisLabelFontSize: 6.5, valAxisLabelColor: COL.inkFaint,
    valGridLine: { color: COL.ruleSoft, size: 0.75 }, catGridLine: { style: "none" },
    valAxisLineShow: false, showValAxisTitle: false, showCatAxisTitle: false,
    lineSize: 1.75, lineDataSymbol: "none", showLegend: false, showTitle: false,
    chartArea: { fill: { color: COL.white } }, plotArea: { fill: { color: COL.white } },
  };
  const seriesPalette = [COL.sales, COL.orders];

  cfg.charts.forEach((chart, i) => {
    const x = ML + i * (chartW + chartGap);
    slide.addShape("roundRect", {
      x, y: chartY, w: chartW, h: chartH, rectRadius: 0.05,
      fill: { color: COL.white }, line: { color: COL.rule, width: 1 },
    });
    slide.addText(chart.title, {
      x: x + 0.1, y: chartY + 0.06, w: chartW - 0.2, h: 0.18,
      fontFace: FONT_BODY, fontSize: 9.3, bold: true, color: COL.ink, margin: 0,
    });
    slide.addText(chart.legend, {
      x: x + 0.1, y: chartY + 0.24, w: chartW - 0.2, h: 0.15,
      fontFace: FONT_BODY, fontSize: 7, color: COL.inkMuted, margin: 0,
    });

    const chartData = chart.seriesKeys.map((key, si) => ({
      name: chart.names[si],
      labels: cfg.months,
      values: cfg.series[key],
    }));
    const chartType = chart.type === "area" ? pres.ChartType.area : pres.ChartType.line;
    const opts = {
      x: x + 0.05, y: chartY + 0.42, w: chartW - 0.1, h: chartH - 0.48,
      chartColors: chart.type === "area" ? [seriesPalette[0]] : seriesPalette.slice(0, chart.seriesKeys.length),
      valAxisLabelFormatCode: chart.valFormat || "#,##0",
      ...axisOpts,
    };
    if (chart.type === "area") opts.chartColorsOpacity = 30;
    slide.addChart(chartType, chartData, opts);
  });

  // ---------- bottom row: 所見 / (追加セクション) / コメント ----------
  // データ注記欄は廃止（2026-07-29のユーザー要望）。その分、所見に使える高さを広げてある。
  const botY = chartY + chartH + 0.12;
  const botH = 7.5 - botY - 0.16;
  const botGap = 0.18;
  const avail = CW - (extInColumn ? 2 : 1) * botGap;
  const insW = extInColumn ? avail * 0.27 : avail * 0.55;
  const topW = extInColumn ? avail * 0.43 : 0;
  const comW = extInColumn ? avail * 0.3 : avail * 0.45;

  slide.addText("所見", {
    x: ML, y: botY, w: insW, h: 0.18,
    fontFace: FONT_BODY, fontSize: 10.5, bold: true, color: COL.accent, margin: 0,
  });

  const insightsY = botY + 0.23;
  const insightsH = botY + botH - insightsY;
  const insightsFontSize = extInColumn ? 8.2 : 9.5;
  slide.addText(
    cfg.insights.flatMap((it) => [
      { text: it.lead, options: { bold: true, color: COL.accent, breakLine: false, bullet: { code: "25CF", indent: 10 } } },
      { text: it.rest, options: { color: COL.ink, breakLine: true } },
    ]),
    {
      x: ML, y: insightsY, w: insW, h: insightsH,
      fontFace: FONT_BODY, fontSize: insightsFontSize, margin: 0, valign: "top",
      lineSpacingMultiple: 1.08, paraSpaceAfter: 4,
    }
  );

  // ---------- 追加セクション（会社特有の指標。列レイアウトのときだけここに収める） ----------
  let comX = ML + insW + botGap;
  if (extInColumn) {
    const topX = ML + insW + botGap;
    comX = topX + topW + botGap;

    slide.addText(ext.title, {
      x: topX, y: botY, w: topW, h: 0.18,
      fontFace: FONT_BODY, fontSize: 10.5, bold: true, color: COL.accent, margin: 0,
    });
    if (ext.subtitle) {
      slide.addText(ext.subtitle, {
        x: topX, y: botY + 0.2, w: topW, h: 0.15, fontFace: FONT_BODY, fontSize: 6.8, color: COL.inkMuted, margin: 0,
      });
    }
    const contentY = botY + 0.38;
    const contentH = botY + botH * 0.86 - contentY;
    addExtraContent(slide, ext, topX, contentY, topW, contentH, { fontScale: 1 });

    if (ext.note) {
      slide.addText(ext.note, {
        x: topX, y: contentY + contentH + 0.03, w: topW, h: botY + botH - (contentY + contentH + 0.03),
        fontFace: FONT_BODY, fontSize: 6.5, color: COL.inkFaint, margin: 0, valign: "top", lineSpacingMultiple: 1.02,
      });
    }
  }

  // comment box — deliberately left BLANK (ruled note area). Do not fill this with
  // AI-drafted comments; the user asked for a manual annotation space here.
  slide.addText("コメント", {
    x: comX, y: botY, w: comW, h: 0.18,
    fontFace: FONT_BODY, fontSize: 10.5, bold: true, color: COL.accent, margin: 0,
  });
  slide.addShape("roundRect", {
    x: comX, y: botY + 0.23, w: comW, h: botH - 0.23, rectRadius: 0.05,
    fill: { color: COL.white }, line: { color: COL.rule, width: 1 },
  });
  const comInnerTop = botY + 0.23;
  const comInnerH = botH - 0.23;
  const ruleGap = 0.3;
  const ruleStart = 0.34;
  const ruleCount = Math.floor((comInnerH - ruleStart - 0.12) / ruleGap);
  for (let i = 0; i < ruleCount; i++) {
    const ly = comInnerTop + ruleStart + i * ruleGap;
    slide.addShape("line", {
      x: comX + 0.15, y: ly, w: comW - 0.3, h: 0,
      line: { color: COL.ruleSoft, width: 1 },
    });
  }

  // ---------- 追加セクションが2枚目全面（extraSection.layout === "page"）の場合 ----------
  if (extOnPage) {
    const slide2 = pres.addSlide();
    slide2.background = { color: COL.white };
    addHeader(slide2, cfg, ML, MT, CW);

    const titleY = MT + 0.66;
    slide2.addText(ext.title, {
      x: ML, y: titleY, w: CW, h: 0.32,
      fontFace: FONT_BODY, fontSize: 14, bold: true, color: COL.accent, margin: 0,
    });
    if (ext.subtitle) {
      slide2.addText(ext.subtitle, {
        x: ML, y: titleY + 0.34, w: CW, h: 0.24, fontFace: FONT_BODY, fontSize: 9.5, color: COL.inkMuted, margin: 0,
      });
    }
    const contentY = titleY + 0.68;
    const contentH = 7.5 - 0.3 - contentY - (ext.note ? 0.5 : 0);
    addExtraContent(slide2, ext, ML, contentY, CW, contentH, { fontScale: 1.35 });

    if (ext.note) {
      slide2.addText(ext.note, {
        x: ML, y: contentY + contentH + 0.06, w: CW, h: 0.4,
        fontFace: FONT_BODY, fontSize: 9, color: COL.inkMuted, margin: 0, valign: "top", lineSpacingMultiple: 1.1,
      });
    }
  }
}

function setupPres() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
  pres.theme = { headFontFace: FONT_HEAD, bodyFontFace: FONT_BODY };
  return pres;
}

// 1社分だけの .pptx を作る（従来どおりの使い方）
function build(cfg, outPath) {
  const pres = setupPres();
  addCompanySlides(pres, cfg);
  return pres.writeFile({ fileName: outPath });
}

// 複数社分を1つの .pptx にまとめる（グループ会社を束ねて経営会議資料にする場合）
function buildMulti(configs, outPath) {
  const pres = setupPres();
  configs.forEach((cfg) => addCompanySlides(pres, cfg));
  return pres.writeFile({ fileName: outPath });
}

if (require.main === module) {
  const dataPath = process.argv[2];
  if (!dataPath) {
    console.error("usage: node build_deck.js data.json [out.pptx]");
    console.error("       node build_deck.js bundle.json [out.pptx]   (bundle.json = { \"companies\": [...] })");
    process.exit(1);
  }
  const parsed = JSON.parse(fs.readFileSync(dataPath, "utf8"));
  if (Array.isArray(parsed.companies)) {
    const outPath = process.argv[3] || parsed.outFile || "bundle.pptx";
    buildMulti(parsed.companies, outPath).then(() => console.log("written:", path.resolve(outPath)));
  } else {
    const outPath = process.argv[3] || parsed.outFile || "deck.pptx";
    build(parsed, outPath).then(() => console.log("written:", path.resolve(outPath)));
  }
}

module.exports = { build, buildMulti, deriveKpiDisplay, fmtDelta };
