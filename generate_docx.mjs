import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType, WidthType, ShadingType, PageBreak, BorderStyle, LevelFormat, convertInchesToTwip, Tab, TabStopType, TabStopPosition } from "docx";
import * as fs from "fs";

const outputDir = process.argv[2] || "output/6h_full";
const summary = JSON.parse(fs.readFileSync(`${outputDir}/claude_analysis.json`, "utf-8"));
const chunks = JSON.parse(fs.readFileSync(`${outputDir}/chunk_details.json`, "utf-8"));
let gemini = null;
try { gemini = JSON.parse(fs.readFileSync(`${outputDir}/gemini_visual.json`, "utf-8")); } catch {}

const products = summary.products;
const patterns = summary.pattern_analysis;
const techDist = patterns.technique_distribution;
const info = summary.broadcast_info;
const techJp = { scarcity: "希少性", social_proof: "社会的証明", anchoring: "アンカリング", comparison: "比較", authority: "権威", urgency: "緊急性" };
const techColors = { scarcity: "e15759", social_proof: "4e79a7", anchoring: "f28e2b", comparison: "76b7b2", authority: "59a14f", urgency: "edc948" };
const totalTech = Object.values(techDist).reduce((a, b) => a + b, 0);
// landscape: width=短辺, height=長辺をセットし、orientation: "landscape" でライブラリがスワップ
const DXA_LETTER_W = 12240;
const DXA_LETTER_H = 15840;
// landscape時の実際の幅は長辺15840、マージン左右1260ずつ引くと13320
const COL_W = 13320; // usable width in landscape

function timeToMin(t) {
  const parts = t.split(":");
  return parseInt(parts[0]) + parseInt(parts[1]) / 60;
}

function makeTextRun(text, opts = {}) {
  return new TextRun({ text, font: "Arial", size: opts.size || 18, bold: opts.bold, color: opts.color, italics: opts.italics });
}

function makePara(text, opts = {}) {
  return new Paragraph({
    children: [makeTextRun(text, opts)],
    spacing: { after: opts.after || 60, before: opts.before || 0 },
    alignment: opts.align,
    indent: opts.indent ? { left: opts.indent } : undefined,
  });
}

function makeCell(text, opts = {}) {
  return new TableCell({
    children: [new Paragraph({ children: [makeTextRun(text, { size: opts.size || 16, bold: opts.bold, color: opts.color })], spacing: { after: 20 } })],
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.shading ? { type: ShadingType.CLEAR, fill: opts.shading } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
  });
}

function makeHeaderCell(text, w) {
  return makeCell(text, { bold: true, size: 14, shading: "1a1a2e", color: "ffffff", width: w });
}

function makeTable(headers, rows, widths) {
  const totalW = widths.reduce((a, b) => a + b, 0);
  return new Table({
    columnWidths: widths,
    width: { size: totalW, type: WidthType.DXA },
    rows: [
      new TableRow({ children: headers.map((h, i) => makeHeaderCell(h, widths[i])), tableHeader: true }),
      ...rows.map(row => new TableRow({
        children: row.map((val, i) => makeCell(String(val || ""), { width: widths[i], size: 15 })),
      })),
    ],
  });
}

// ============================================================
const sections = [];

// ===== 表紙 =====
sections.push({
  properties: { page: { size: { width: DXA_LETTER_W, height: DXA_LETTER_H, orientation: "landscape" }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
  children: [
    ...Array(3).fill(makePara("")),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [makeTextRun("ショップチャンネル 6時間放送", { size: 48, bold: true, color: "1a1a2e" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [makeTextRun("分析レポート", { size: 48, bold: true, color: "1a1a2e" })] }),
    makePara(""),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [makeTextRun("セールストーク構造 × 番組構成 × CV導線 の統合分析", { size: 24, color: "666666" })] }),
    ...Array(2).fill(makePara("")),
    ...[
      `商品数: ${products.length} | 放送時間: 360分 | テクニック検出: ${totalTech}件 | セグメント: ${info.total_segments.toLocaleString()}`,
      `分析日: ${info.analysis_date} | URL: ${info.url}`,
      `キャンペーン: ${info.campaign}`,
      `分析手法: Claude Code テキスト分析 + Gemini 2.5 Flash 視覚分析`,
    ].map(t => new Paragraph({ alignment: AlignmentType.CENTER, children: [makeTextRun(t, { size: 18, color: "999999" })] })),
  ],
});

// ===== 本編 =====
const children = [];

// --- 目次 ---
children.push(new Paragraph({ children: [makeTextRun("目次", { size: 28, bold: true, color: "1a1a2e" })], heading: HeadingLevel.HEADING_1 }));
const tocItems = [
  "1. エグゼクティブサマリー",
  "2. 商品タイムライン",
  "3. セールストーク パターン分析（共通構造・テクニック分布・頻出フレーズ）",
  "4. 番組構成・CV導線の設計分析（フェーズ配分・CTA設計パターン）",
  "5. インサイト・示唆（6つの応用知見）",
  "6.【詳細】商品別分析（テクニック実例・原文引用付き）",
  ...products.map((p, i) => `    6.${i + 1} ${p.name.slice(0, 35)}`),
  "7.【詳細】Gemini 視覚分析ハイライト",
];
tocItems.forEach(t => children.push(makePara(t, { size: 18 })));
children.push(new Paragraph({ children: [new PageBreak()] }));

// --- 1. エグゼクティブサマリー ---
children.push(new Paragraph({ children: [makeTextRun("1. エグゼクティブサマリー", { size: 28, bold: true, color: "1a1a2e" })], heading: HeadingLevel.HEADING_1 }));
children.push(makePara(
  `6時間の放送で7つの商品セグメントを分析。1商品あたり平均45-55分のサイクルで「導入→実演デモ(複数ラウンド)→価格発表(アンカリング)→社会的証明→在庫カウントダウン(希少性エスカレーション)→最終CTA→クロージング」の定型構造を持つ。最も多用されるテクニックは希少性（${techDist.scarcity}件）で、リアルタイム在庫数カウントダウンが全セグメントで使われている。全員送料無料キャンペーン（4/10-16）が横断的な緊急性ドライバーとして機能。`, { size: 19 }));

// --- 2. 商品タイムライン ---
children.push(makePara("", { before: 200 }));
children.push(new Paragraph({ children: [makeTextRun("2. 商品タイムライン", { size: 28, bold: true, color: "1a1a2e" })], heading: HeadingLevel.HEADING_1 }));
const tlWidths = [3500, 1800, 900, 4000, 1200, 1500];
const tlRows = products.map(p => [
  p.name.slice(0, 30), `${p.start_time}〜${p.end_time}`,
  `${p.total_duration_minutes || p.duration_minutes || ""}分`,
  (p.price || "").slice(0, 42), p.category || "", (p.brand || "").slice(0, 18),
]);
children.push(makeTable(["商品名", "時間帯", "時間", "価格", "カテゴリ", "ブランド"], tlRows, tlWidths));
children.push(new Paragraph({ children: [new PageBreak()] }));

// --- 3. セールストーク パターン分析 ---
children.push(new Paragraph({ children: [makeTextRun("3. セールストーク パターン分析", { size: 28, bold: true, color: "1a1a2e" })], heading: HeadingLevel.HEADING_1 }));
children.push(new Paragraph({ children: [makeTextRun("3.1 共通トーク構造", { size: 22, bold: true, color: "4e79a7" })], heading: HeadingLevel.HEADING_2 }));
children.push(makePara(patterns.common_structure, { size: 18 }));

children.push(new Paragraph({ children: [makeTextRun("3.2 テクニック分布", { size: 22, bold: true, color: "4e79a7" })], heading: HeadingLevel.HEADING_2 }));
const maxTech = Math.max(...Object.values(techDist));
const techRows = Object.entries(techDist).sort((a, b) => b[1] - a[1]).map(([k, v]) => [
  techJp[k] || k, String(v), `${(v / totalTech * 100).toFixed(1)}%`, "█".repeat(Math.round(v / maxTech * 20)),
]);
children.push(makeTable(["テクニック", "検出数", "割合", "分布"], techRows, [2500, 1200, 1200, 8000]));

children.push(new Paragraph({ children: [makeTextRun("3.3 頻出フレーズ TOP 10", { size: 22, bold: true, color: "4e79a7" })], heading: HeadingLevel.HEADING_2 }));
const phraseRows = (patterns.top_phrases || []).map(ph => [`「${ph.phrase}」`, String(ph.count), ph.context || ""]);
children.push(makeTable(["フレーズ", "回数", "文脈"], phraseRows, [4500, 800, 7600]));
children.push(new Paragraph({ children: [new PageBreak()] }));

// --- 4. 番組構成・CV導線 ---
children.push(new Paragraph({ children: [makeTextRun("4. 番組構成・CV導線の設計分析", { size: 28, bold: true, color: "1a1a2e" })], heading: HeadingLevel.HEADING_1 }));
children.push(new Paragraph({ children: [makeTextRun("4.1 フェーズ別時間配分", { size: 22, bold: true, color: "4e79a7" })], heading: HeadingLevel.HEADING_2 }));
const phaseJp = { intro: "導入", demo: "実演デモ", price_reveal: "価格発表", social_proof: "社会的証明", cta_scarcity: "CTA・在庫訴求", close: "クロージング" };
const phaseRows = Object.entries(patterns.avg_phase_duration_pct || {}).map(([k, v]) => [phaseJp[k] || k, `${v}%`, "█".repeat(Math.round(v / 5))]);
children.push(makeTable(["フェーズ", "時間比率", "分布"], phaseRows, [3500, 1500, 8000]));

children.push(new Paragraph({ children: [makeTextRun("4.2 CTA設計パターン", { size: 22, bold: true, color: "4e79a7" })], heading: HeadingLevel.HEADING_2 }));
const ctaPatterns = [
  ["電話番号", "常時表示。「0120フリーダイヤル」「タッチでショップ2番」で注文切替。全セグメントで使用。"],
  ["Web/QR", "「スマホ・パソコンでショップチャンネルと検索」「QRコード」。電話混雑時の代替チャネルとして案内。"],
  ["在庫カウントダウン", "リアルタイムで「○○点を切りました」を1セグメント10-25回更新。サイズ・色別に詳細報告。チャイム音連動。"],
  ["注文件数報告", "「お電話○○名」「○○件のオーダー」でバンドワゴン効果。最大1万件突破の報告。同時通話数も告知。"],
  ["送料無料キャンペーン", "全セグメントで繰り返し（推定30回以上）。「今日からスタート」「16日まで」「別住所への送付も無料」。"],
  ["ショップチャンネルカード", "5%還元を主要セグメントで告知。送料定額サービスとの併用案内も。"],
  ["マッチングプライス", "複数商品セット購入で割引（メルティリッチダウン: かけ敷き12,960円、かけ2枚13,960円）。"],
];
children.push(makeTable(["CTA種別", "設計パターン"], ctaPatterns, [2500, 10400]));
children.push(new Paragraph({ children: [new PageBreak()] }));

// --- 5. インサイト ---
children.push(new Paragraph({ children: [makeTextRun("5. インサイト・示唆 — 他の販売チャネルに応用可能な知見", { size: 28, bold: true, color: "1a1a2e" })], heading: HeadingLevel.HEADING_1 }));
const insights = [
  ["在庫カウントダウンのリアルタイム性", `「残り○○点」の連続更新が最強の購買ドライバー（${techDist.scarcity}件検出）。1セグメントで10-25回更新される。ECサイトでも在庫数のリアルタイム表示、「残りN点」通知、sold-out表示が購買転換率を大幅に改善する。`],
  ["デモの反復構造（全体の45%）", "1商品につき同じデモを2-3ラウンド繰り返す設計。途中参加の視聴者に対応しつつ、繰り返しで購買意欲を高める二重効果。動画コマースでも「途中から見ても分かる」ループ構成が鍵。"],
  ["アンカリングの二重構造", "「メーカー希望小売価格→SC価格」の空間的アンカリングに加え、「原材料高騰で値上げすべきを据え置き」「30周年特別」の時間的アンカリングを組み合わせ。単純な値引きより「企業努力」のストーリーが効く。"],
  ["社会的証明のライブ感", `注文件数・同時通話数をリアルタイム報告（${techDist.social_proof}件検出）。「お電話350名」「1万件突破」+顧客メッセージ即時読み上げ。バンドワゴン効果とFOMOを同時発動。レビュー数表示やリアルタイム購入通知に応用可能。`],
  ["ゲスト出演の権威効果", `社長・デザイナー・実演販売士が全商品に登場（権威テクニック${techDist.authority}件）。開発背景や職人のこだわりを「本人の口」から語る。インフルエンサーマーケティングの原型。EC上でも創業者ストーリーやメーカー担当者インタビューの埋め込みが有効。`],
  ["送料無料の横断ドライバー設計", "期間限定キャンペーン（推定30回以上言及）が個別商品訴求とは別レイヤーで全セグメントを貫通。「今買う理由」を商品特性に依存せず生成。ECでも「全品送料無料ウィーク」は個別クーポンより効果的な場合がある。"],
];
insights.forEach(([title, body], i) => {
  children.push(new Paragraph({ children: [makeTextRun(`${i + 1}. ${title}`, { size: 21, bold: true, color: "4e79a7" })], spacing: { before: 160, after: 40 } }));
  children.push(makePara(body, { size: 18 }));
});
children.push(new Paragraph({ children: [new PageBreak()] }));

// --- 6. 商品別詳細 ---
children.push(new Paragraph({ children: [makeTextRun("6.【詳細】商品別分析", { size: 28, bold: true, color: "1a1a2e" })], heading: HeadingLevel.HEADING_1 }));
children.push(makePara("各商品のセールストークに使用された説得テクニックを、原文引用とタイムスタンプ付きで記録。", { size: 18 }));

products.forEach((prod, pi) => {
  children.push(new Paragraph({ children: [new PageBreak()] }));
  const ts = prod.techniques_summary || {};
  const dur = prod.total_duration_minutes || prod.duration_minutes || "?";

  children.push(new Paragraph({ children: [makeTextRun(`6.${pi + 1} ${prod.name}`, { size: 24, bold: true, color: "1a1a2e" })], heading: HeadingLevel.HEADING_2 }));

  // 基本情報
  const infoRows = [
    ["ブランド", prod.brand || "", "品番", prod.item_number || "N/A"],
    ["時間帯", `${prod.start_time}〜${prod.end_time} (${dur}分)`, "カテゴリ", prod.category || ""],
    ["価格", prod.price || "", "定価/割引", prod.retail_price || prod.discount || ""],
  ];
  children.push(makeTable(["", "", "", ""], infoRows.map(r => r.map(v => v || "")), [1800, 4600, 1800, 4700]));

  // テクニック分布
  children.push(new Paragraph({ children: [makeTextRun("テクニック分布", { size: 20, bold: true, color: "4e79a7" })], spacing: { before: 120, after: 40 } }));
  const maxV = Math.max(...Object.values(ts), 1);
  const tsRows = ["scarcity", "social_proof", "anchoring", "comparison", "authority", "urgency"]
    .map(k => [techJp[k], String(ts[k] || 0), "█".repeat(Math.round((ts[k] || 0) / maxV * 15))]);
  children.push(makeTable(["テクニック", "件数", "分布"], tsRows, [2500, 1000, 9400]));

  // キーフレーズ
  children.push(new Paragraph({ children: [makeTextRun("キーフレーズ", { size: 20, bold: true, color: "4e79a7" })], spacing: { before: 120, after: 40 } }));
  (prod.key_phrases || []).slice(0, 6).forEach(kp => {
    children.push(makePara(`・「${kp}」`, { size: 17, indent: 200 }));
  });

  // フェーズ構成
  children.push(new Paragraph({ children: [makeTextRun("フェーズ構成", { size: 20, bold: true, color: "4e79a7" })], spacing: { before: 120, after: 40 } }));
  const phRows = (prod.phases || []).map(ph => [ph.type, `${ph.start}〜${ph.end}`, ph.note || ""]);
  if (phRows.length) children.push(makeTable(["フェーズ", "時間帯", "内容"], phRows, [2000, 2000, 8900]));

  // テクニック実例
  children.push(new Paragraph({ children: [makeTextRun("テクニック実例（原文引用）", { size: 20, bold: true, color: "4e79a7" })], spacing: { before: 120, after: 40 } }));
  const startMin = timeToMin(prod.start_time);
  const endMin = timeToMin(prod.end_time);

  for (const techKey of ["scarcity", "social_proof", "anchoring", "comparison", "authority", "urgency"]) {
    const examples = [];
    for (const chunk of chunks) {
      const techs = chunk.techniques || {};
      for (const item of (techs[techKey] || [])) {
        const tsStr = item.timestamp || "";
        const tsMin = parseInt(tsStr.split(":")[0]) || 0;
        if (tsMin >= startMin - 5 && tsMin <= endMin + 10) examples.push(item);
      }
    }
    if (!examples.length) continue;

    children.push(new Paragraph({
      children: [makeTextRun(`■ ${techJp[techKey]}（${examples.length}件）`, { size: 18, bold: true, color: techColors[techKey] })],
      spacing: { before: 80, after: 20 },
    }));
    examples.slice(0, 5).forEach(ex => {
      const quote = ex.quote || ex.text || "";
      const tsStr = ex.timestamp || "";
      const context = ex.context || "";
      let text = `[${tsStr}] ${quote}`;
      if (context) text += ` — ${context}`;
      children.push(makePara(text, { size: 16, color: "555555", indent: 400 }));
    });
  }
});

// --- 7. Gemini 視覚分析 ---
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(new Paragraph({ children: [makeTextRun("7.【詳細】Gemini 視覚分析ハイライト", { size: 28, bold: true, color: "1a1a2e" })], heading: HeadingLevel.HEADING_1 }));
children.push(makePara("60枚のスクリーンショットをGemini 2.5 Flashで分析。テロップ・価格表示・QRコード・画面レイアウト等の視覚要素を抽出。以下は各バッチの分析結果抜粋。", { size: 18 }));

if (gemini) {
  let batchNum = 0;
  for (const r of (gemini.results || [])) {
    const text = r.gemini_analysis || "";
    if (text.length > 100) {
      batchNum++;
      children.push(new Paragraph({
        children: [makeTextRun(`バッチ ${batchNum}: ${r.timestamp_display} — ${r.reason}`, { size: 19, bold: true, color: "4e79a7" })],
        spacing: { before: 120, after: 20 },
      }));
      children.push(makePara(`スクリーンショット: ${r.screenshot || ""} | ${r.description || ""}`, { size: 15, color: "999999" }));
      const preview = text.slice(0, 1500) + (text.length > 1500 ? "\n... (以下省略)" : "");
      children.push(makePara(preview, { size: 15, color: "555555" }));
      if (batchNum >= 8) break;
    }
  }
}

// フッター
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(...Array(5).fill(makePara("")));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [makeTextRun("Claude Code テキスト分析 + Gemini 2.5 Flash 視覚分析", { size: 20, color: "999999" })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [makeTextRun(`生成日: ${info.analysis_date}`, { size: 20, color: "999999" })] }));

sections.push({
  properties: { page: { size: { width: DXA_LETTER_W, height: DXA_LETTER_H, orientation: "landscape" }, margin: { top: 1080, bottom: 1080, left: 1260, right: 1260 } } },
  children,
});

// ===== 生成 =====
const doc = new Document({
  sections,
  numbering: { config: [{ reference: "bullet", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 200 } } } }] }] },
});

const buffer = await Packer.toBuffer(doc);
const docxPath = `${outputDir}/analysis_report.docx`;
fs.writeFileSync(docxPath, buffer);
console.log(`DOCX出力: ${docxPath}`);
