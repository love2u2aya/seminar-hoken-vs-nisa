// セミナー：変額保険・一時払保険（貯蓄/保障）・NISA 4商品比較
// 27枚構成・4:3レイアウト・原色パレット（navy/green/gold + teal）
// 日本語フォント基本48pt / 最小28pt / 助詞（てにをは）は -4pt

const PptxGenJS = require('pptxgenjs');
const pres = new PptxGenJS();
pres.layout = 'LAYOUT_4x3'; // 10 x 7.5 inches
pres.title = '変額・一時払保険・NISA 比較セミナー';

// =========================================================
// カラー・フォント
// =========================================================
// Midnight Gold パレット（支払方法グループ濃淡構成）
//   毎月グループ（明るい金系）  : 変額 EFB509 / NISA F4C430
//   一時払グループ（濃い紺系） : 貯蓄 16253D（midnight）/ 保障 002C54（dusk）
const C = {
  // 基本（タイトルバー・本文・強調共通）
  navy:       '16253D',  // midnight：タイトルバー・本文
  gold:       'EFB509',  // golden  ：結論・強調
  red:        'C53030',  // 注意・リスク
  white:      'FFFFFF',
  offWhite:   'F8F5F2',  // pearl
  gray:       '4A5568',
  lightGray:  'CBD5E0',
  bodyText:   '16253D',
  // 互換用エイリアス（旧slotを残す）
  darkNavy:   '002C54',
  teal:       'EFB509',
  green:      'F4C430',
  lightNavy:  'B8C6D9',
  lightGold:  'D9E2EC',
  lightTeal:  'FFF3C2',
  lightGreen: 'FFFAE0',
};
const F  = 'Meiryo';
const FA = 'Arial';

// =========================================================
// 助詞分割：基本サイズ baseSize、助詞は -4pt
// =========================================================
function jpRuns(text, baseSize) {
  const pSize = Math.max(baseSize - 4, 18);
  const pRe = /(より|から|まで|として|では|には|とは|など|けど|ので|のに|だが|しか|は|が|を|に|で|と|も|へ|や|の)/g;
  const result = [];
  String(text).split('\n').forEach((line, li) => {
    if (li > 0) result.push({ text: '', options: { breakLine: true } });
    let last = 0;
    pRe.lastIndex = 0;
    let m;
    while ((m = pRe.exec(line)) !== null) {
      if (m.index > last) result.push({ text: line.slice(last, m.index), options: { fontSize: baseSize } });
      result.push({ text: m[0], options: { fontSize: pSize } });
      last = m.index + m[0].length;
    }
    if (last < line.length) result.push({ text: line.slice(last), options: { fontSize: baseSize } });
  });
  return result;
}

// =========================================================
// ヘルパー
// =========================================================
function bg(s, color) { s.background = { color: color || C.offWhite }; }

function sectionTitle(s, title, color) {
  const bar = color || C.navy;
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.55, fill: { color: bar }, line: { color: bar } });
  s.addText(jpRuns(title, 28), {
    x: 0.3, y: 0, w: 9.4, h: 0.55,
    fontFace: F, bold: true, color: C.white, margin: 0, valign: 'middle'
  });
}

function emphBox(s, x, y, w, h, text, fontSize, color) {
  const c = color || C.green;
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: c }, line: { color: c } });
  s.addText(jpRuns(text, fontSize || 32), {
    x, y, w, h,
    fontFace: F, bold: true, color: C.white,
    align: 'center', valign: 'middle', margin: 0.05
  });
}

function fourCards(s, y, items) {
  const cardW = 2.25, gap = 0.1;
  const startX = (10 - (cardW * 4 + gap * 3)) / 2;
  items.forEach((it, i) => {
    const x = startX + i * (cardW + gap);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cardW, h: 0.55, fill: { color: it.color }, line: { color: it.color } });
    s.addText(jpRuns(it.title, 24), {
      x, y, w: cardW, h: 0.55,
      fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: y + 0.55, w: cardW, h: 3.7,
      fill: { color: it.lightColor }, line: { color: it.color, width: 1 }
    });
    s.addText(jpRuns(it.body, 24), {
      x: x + 0.1, y: y + 0.65, w: cardW - 0.2, h: 3.5,
      fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 4
    , bold: true });
  });
}

function twoColCards(s, y, leftTitle, leftBody, leftColor, rightTitle, rightBody, rightColor) {
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 4.55, h: 0.6, fill: { color: leftColor }, line: { color: leftColor } });
  s.addText(jpRuns(leftTitle, 24), { x: 0.4, y, w: 4.55, h: 0.6, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: y + 0.6, w: 4.55, h: 4.0, fill: { color: C.white }, line: { color: leftColor, width: 1.5 } });
  s.addText(jpRuns(leftBody, 24), { x: 0.55, y: y + 0.7, w: 4.25, h: 3.8, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 6 , bold: true });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.05, y, w: 4.55, h: 0.6, fill: { color: rightColor }, line: { color: rightColor } });
  s.addText(jpRuns(rightTitle, 24), { x: 5.05, y, w: 4.55, h: 0.6, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.05, y: y + 0.6, w: 4.55, h: 4.0, fill: { color: C.white }, line: { color: rightColor, width: 1.5 } });
  s.addText(jpRuns(rightBody, 24), { x: 5.2, y: y + 0.7, w: 4.25, h: 3.8, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 6 , bold: true });
}

function footer(s, num, total) {
  s.addText(`${num} / ${total}`, {
    x: 9.0, y: 7.2, w: 0.9, h: 0.2,
    fontSize: 10, fontFace: FA, color: C.gray, align: 'right', margin: 0
  });
}

const TOTAL = 27;
let slideNo = 0;
function newSlide(bgColor) {
  slideNo++;
  const s = pres.addSlide();
  bg(s, bgColor || C.offWhite);
  return s;
}

// 商品色：2軸マトリクス
//   色相＝目的（保障=青系 / 貯蓄=黄系）
//   明度＝支払方法（一括=濃 / 毎月=薄）
const PROD = {
  hengaku: { name: '変額保険',         color: '4A6FA5', light: 'E1EAF7' }, // 保障系×毎月：薄い青
  ipHosh:  { name: '一時払・保障重視', color: '16253D', light: 'D9E2EC' }, // 保障系×一括：濃い青
  nisa:    { name: 'NISA',            color: 'EFB509', light: 'FFF3C2' }, // 貯蓄系×毎月：薄い黄
  ipChoSh: { name: '一時払・貯蓄重視', color: 'B07A0C', light: 'F4DDA1' }, // 貯蓄系×一括：濃い黄
};

// ===== Slide 1: 表紙 =====
{
  const s = newSlide(C.navy);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 2.8, w: 10, h: 0.08, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText(jpRuns('保険か、NISAか。\n本当に自分に合うのはどれ？', 40), {
    x: 0.5, y: 1.0, w: 9, h: 1.7,
    fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0
  });
  s.addText(jpRuns('変額保険 × 一時払保険 × NISA\n4商品をやさしく比較', 28), {
    x: 0.5, y: 3.0, w: 9, h: 1.4,
    fontFace: F, color: C.lightGray, align: 'center', valign: 'middle', margin: 0
  });
  s.addText('Money Navigation Seminar', {
    x: 0.5, y: 5.5, w: 9, h: 0.4,
    fontSize: 14, fontFace: FA, italic: true, color: C.gold, align: 'center', margin: 0
  });
}

// ===== Slide 2: 今日のゴール =====
{
  const s = newSlide();
  sectionTitle(s, '今日のゴール');
  const goals = [
    { n: '①', t: '違いがわかる',     d: '4商品の役割と特徴を整理' },
    { n: '②', t: '選び方がわかる',   d: '自分に合う組み合わせを判断' },
    { n: '③', t: '次の一歩がわかる', d: '今日から動ける具体策' },
  ];
  goals.forEach((g, i) => {
    const y = 1.3 + i * 1.85;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y, w: 1.2, h: 1.2, fill: { color: C.gold }, line: { color: C.gold } });
    s.addText(g.n, { x: 0.5, y, w: 1.2, h: 1.2, fontSize: 44, fontFace: FA, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    s.addText(jpRuns(g.t, 32), { x: 1.9, y: y + 0.05, w: 7.6, h: 0.7, fontFace: F, bold: true, color: C.navy, align: 'left', valign: 'middle', margin: 0 });
    s.addText(jpRuns(g.d, 24), { x: 1.9, y: y + 0.7,  w: 7.6, h: 0.5, fontFace: F, color: C.gray, align: 'left', valign: 'middle', margin: 0 });
  });
  footer(s, slideNo, TOTAL);
}

// ===== Slide 3: 悩み =====
{
  const s = newSlide();
  sectionTitle(s, 'こんな悩み、ありませんか？');
  const worries = [
    '保険とNISA、両方やるべき？',
    '変額保険って結局どうなの？',
    '退職金、一括で何かに入れたい',
    '家族に何か残したい気もする',
    'でも増やしたい気持ちもある',
    '何から始めれば正解なの？',
  ];
  worries.forEach((w, i) => {
    const x = 0.4 + (i % 2) * 4.7;
    const y = 1.2 + Math.floor(i / 2) * 1.8;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 4.5, h: 1.5, fill: { color: C.white }, line: { color: C.gold, width: 2 }, rectRadius: 0.1 });
    s.addText(jpRuns(w, 24), { x: x + 0.2, y, w: 4.1, h: 1.5, fontFace: F, color: C.bodyText, bold: true, align: 'left', valign: 'middle', margin: 0.05 });
  });
  footer(s, slideNo, TOTAL);
}

// ===== Slide 4: ポジショニング =====
{
  const s = newSlide();
  sectionTitle(s, '4商品のポジショニング');
  const cx = 5, cy = 4.0, sz = 2.6;
  s.addShape(pres.shapes.LINE, { x: cx - sz, y: cy, w: sz * 2, h: 0, line: { color: C.gray, width: 1.5 } });
  s.addShape(pres.shapes.LINE, { x: cx, y: cy - sz, w: 0, h: sz * 2, line: { color: C.gray, width: 1.5 } });
  s.addText('保障 不要 →', { x: 0.5, y: cy - 0.3, w: 2.5, h: 0.5, fontSize: 14, fontFace: F, color: C.gray, align: 'left' });
  s.addText('← 保障 必要', { x: 7,   y: cy - 0.3, w: 2.5, h: 0.5, fontSize: 14, fontFace: F, color: C.gray, align: 'right' });
  s.addText('一括', { x: cx - 0.5, y: cy - sz - 0.4, w: 1, h: 0.3, fontSize: 14, fontFace: F, color: C.gray, align: 'center' });
  s.addText('毎月', { x: cx - 0.5, y: cy + sz + 0.1, w: 1, h: 0.3, fontSize: 14, fontFace: F, color: C.gray, align: 'center' });
  const places = [
    { p: PROD.nisa,    x: cx - sz * 0.7 - 0.6, y: cy + sz * 0.5 - 0.3 },
    { p: PROD.hengaku, x: cx + sz * 0.5 - 0.6, y: cy + sz * 0.5 - 0.3 },
    { p: PROD.ipChoSh, x: cx - sz * 0.7 - 0.6, y: cy - sz * 0.7 - 0.1 },
    { p: PROD.ipHosh,  x: cx + sz * 0.5 - 0.6, y: cy - sz * 0.7 - 0.1 },
  ];
  places.forEach(pl => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: pl.x, y: pl.y, w: 1.8, h: 0.7, fill: { color: pl.p.color }, line: { color: pl.p.color }, rectRadius: 0.08 });
    s.addText(jpRuns(pl.p.name, 24), { x: pl.x, y: pl.y, w: 1.8, h: 0.7, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  });
  footer(s, slideNo, TOTAL);
}

// ===== Slide 5: 変額保険 =====
{
  const s = newSlide();
  sectionTitle(s, '変額保険とは');
  emphBox(s, 0.4, 0.85, 9.2, 0.7, '保障 ＋ 投資 を 一本化', 30, C.navy);
  const body =
    '・毎月の保険料を株式・債券などで運用\n' +
    '・運用成績で死亡保険金や解約返戻金が変動\n' +
    '・死亡保険金は最低保証あり（基本保険金額）\n' +
    '・運用益は非課税で再投資される\n' +
    '・解約は原則 長期前提（10年以上）';
  s.addText(jpRuns(body, 24), { x: 0.5, y: 1.7, w: 9, h: 4.0, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 8 , bold: true });
  emphBox(s, 0.4, 6.0, 9.2, 0.6, '長期 ＋ 保障 が必要 な人 向け', 22, C.gold);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 6: 一時払（貯蓄） =====
{
  const s = newSlide();
  sectionTitle(s, '一時払い保険（貯蓄重視）');
  emphBox(s, 0.4, 0.85, 9.2, 0.7, 'まとまった資金 を 安全 に 置く', 28, C.navy);
  const body =
    '・契約時に保険料を一括で支払う\n' +
    '・予定利率や為替で運用（外貨建てが主流）\n' +
    '・死亡保険金は払込額と同程度〜やや上\n' +
    '・据置期間後は元本＋利息で受取可能\n' +
    '・短期解約は元本割れリスクあり';
  s.addText(jpRuns(body, 24), { x: 0.5, y: 1.7, w: 9, h: 4.0, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 8 , bold: true });
  emphBox(s, 0.4, 6.0, 9.2, 0.6, '退職金 など まとまった資金 の 運用先', 22, C.gold);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 7: 一時払（保障） =====
{
  const s = newSlide();
  sectionTitle(s, '一時払い保険（保障重視）');
  emphBox(s, 0.4, 0.85, 9.2, 0.7, 'まとまった資金 で 保障 を 大きく', 28, C.navy);
  const body =
    '・一括払いで 死亡保険金が払込額の1.3〜2倍\n' +
    '・相続対策・受取人指定の柔軟性\n' +
    '・健康状態の告知は緩めの商品もある\n' +
    '・運用益より「残す」目的が中心\n' +
    '・解約返戻金は払込額より少ない期間あり';
  s.addText(jpRuns(body, 24), { x: 0.5, y: 1.7, w: 9, h: 4.0, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 8 , bold: true });
  emphBox(s, 0.4, 6.0, 9.2, 0.6, '相続 や 家族 へ 残したい 人 向け', 22, C.gold);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 8: NISA =====
{
  const s = newSlide();
  sectionTitle(s, 'NISAとは');
  emphBox(s, 0.4, 0.85, 9.2, 0.7, '運用益 が ずっと 非課税', 30, C.navy);
  const body =
    '・つみたて投資枠：年120万円・長期インデックス向け\n' +
    '・成長投資枠：年240万円・個別株・ETFも可\n' +
    '・生涯非課税枠 1,800万円（売却で枠復活）\n' +
    '・運用益・配当が非課税\n' +
    '・保障機能 なし（純粋な投資制度）';
  s.addText(jpRuns(body, 24), { x: 0.5, y: 1.7, w: 9, h: 4.0, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 8 , bold: true });
  emphBox(s, 0.4, 6.0, 9.2, 0.6, '長期 で 増やしたい 人 の 王道', 22, C.gold);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 9: 4商品まとめ =====
{
  const s = newSlide();
  sectionTitle(s, '4商品まとめ比較');
  fourCards(s, 1.0, [
    { title: '変額保険',
      body: '【支払】毎月\n【保障】◎ 最低保証あり\n【運用】株式中心\n【流動性】△ 長期前提\n【税】保険料控除\n【向く人】保障＋投資\n両取り したい',
      color: PROD.hengaku.color, lightColor: PROD.hengaku.light },
    { title: '一時払（貯蓄）',
      body: '【支払】一括\n【保障】△ 払込額相当\n【運用】予定利率\n【流動性】△ 据置あり\n【税】満期で課税\n【向く人】退職金 を\n安全 に 運用',
      color: PROD.ipChoSh.color, lightColor: PROD.ipChoSh.light },
    { title: '一時払（保障）',
      body: '【支払】一括\n【保障】◎ 1.3〜2倍\n【運用】少なめ\n【流動性】△ 長期\n【税】相続税の非課税枠\n【向く人】相続・残し たい',
      color: PROD.ipHosh.color, lightColor: PROD.ipHosh.light },
    { title: 'NISA',
      body: '【支払】毎月／自由\n【保障】× なし\n【運用】株式・投信\n【流動性】◎ いつでも\n【税】運用益 非課税\n【向く人】長期 で 純粋に\n増やし たい',
      color: PROD.nisa.color, lightColor: PROD.nisa.light },
  ]);
  emphBox(s, 0.4, 5.9, 9.2, 0.7, '「守る」 か 「増やす」 か、 軸 を 決める', 24, C.navy);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 10: 選び方フロー =====
{
  const s = newSlide();
  sectionTitle(s, '選び方フロー：3つの分岐');
  const steps = [
    { q: 'Q1. 保障 は 必要？',                y: 1.2, c: C.navy },
    { q: 'Q2. 一括 で 払える資金 が ある？',  y: 3.1, c: C.navy },
    { q: 'Q3. 値動き（ブレ） を 許容できる？', y: 5.0, c: C.navy },
  ];
  steps.forEach(st => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: st.y, w: 9, h: 1.4, fill: { color: st.c }, line: { color: st.c }, rectRadius: 0.1 });
    s.addText(jpRuns(st.q, 30), { x: 0.5, y: st.y, w: 9, h: 1.4, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  });
  s.addText('▼', { x: 4.7, y: 2.6, w: 0.6, h: 0.5, fontSize: 28, fontFace: FA, color: C.gray, align: 'center' });
  s.addText('▼', { x: 4.7, y: 4.5, w: 0.6, h: 0.5, fontSize: 28, fontFace: FA, color: C.gray, align: 'center' });
  emphBox(s, 0.5, 6.6, 9, 0.45, 'この 3問 で 4商品 の どれか に 絞り込める', 18, C.navy);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 11: 比較①問題提起 =====
{
  const s = newSlide();
  sectionTitle(s, '比較①：変額保険 vs NISA');
  emphBox(s, 0.4, 0.85, 9.2, 0.8, '判断軸：保障 が 必要 か どうか', 30, C.navy);
  twoColCards(s, 1.85,
    '変額保険',
    '★ 死亡保障 あり\n  → 万一 でも 家族 に お金\n★ 強制的 な 長期積立\n  → 途中解約 が 不利\n★ 投資 は 保険会社経由\n  → コスト やや 高め',
    PROD.hengaku.color,
    'NISA',
    '★ 保障 なし\n  → 自分 が 死ぬと 残るだけ\n★ いつでも 売却可能\n  → 流動性 が 高い\n★ コスト が 低い\n  → 増やす 効率 が 良い',
    PROD.nisa.color
  );
  footer(s, slideNo, TOTAL);
}

// ===== Slide 12: 比較①結論 =====
{
  const s = newSlide();
  sectionTitle(s, '比較①：結論');
  emphBox(s, 0.4, 1.0, 9.2, 1.0, '家族 に 残す必要 が あるなら 変額\nそうでないなら NISA', 26, C.gold);
  const txt =
    '・小さい子・配偶者がいる → 変額（保障の安心）\n' +
    '・独身・子が独立済み → NISA（コスト効率優先）\n' +
    '・「両方欲しい」場合 → 掛捨て＋NISA も 選択肢';
  s.addText(jpRuns(txt, 24), { x: 0.6, y: 2.4, w: 8.8, h: 3.5, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 10 , bold: true });
  emphBox(s, 0.4, 6.1, 9.2, 0.6, '保障 ＝ 変額  /  効率 ＝ NISA', 24, C.navy);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 13: 比較②問題提起 =====
{
  const s = newSlide();
  sectionTitle(s, '比較②：一時払(保障) vs 変額');
  emphBox(s, 0.4, 0.85, 9.2, 0.8, '保障 が 必要 と 決めた人 の 選択', 28, C.navy);
  twoColCards(s, 1.85,
    '一時払・保障重視',
    '★ 一括払い\n  → まとまった資金 が 必要\n★ 保障 が 払込額 の 1.3〜2倍\n  → 相続 で 大きく 残せる\n★ 運用益 は 控えめ\n  → 「残す」 が 主目的',
    PROD.ipHosh.color,
    '変額保険',
    '★ 毎月払い\n  → 現役 でも 始められる\n★ 保障 ＋ 値上がり益\n  → 長期 で 育てる\n★ 解約返戻金 は 変動\n  → 投資 リスクあり',
    PROD.hengaku.color
  );
  footer(s, slideNo, TOTAL);
}

// ===== Slide 14: 比較②結論 =====
{
  const s = newSlide();
  sectionTitle(s, '比較②：結論');
  emphBox(s, 0.4, 1.0, 9.2, 1.0, '一括資金 ＋ 相続 → 一時払(保障)\n毎月 ＋ 育てる → 変額', 26, C.green);
  const txt =
    '・退職金・相続資金 がある → 一時払（保障）\n' +
    '・現役世代で コツコツ → 変額\n' +
    '・税の観点：一時払は 相続税の非課税枠 が 強み';
  s.addText(jpRuns(txt, 24), { x: 0.6, y: 2.4, w: 8.8, h: 3.5, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 10 , bold: true });
  emphBox(s, 0.4, 6.1, 9.2, 0.6, '一括 ＝ 一時払  /  毎月 ＝ 変額', 24, C.navy);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 15: 比較③問題提起 =====
{
  const s = newSlide();
  sectionTitle(s, '比較③：支払方法 で 選ぶ');
  emphBox(s, 0.4, 0.85, 9.2, 0.8, '判断軸：一括 で 払える資金 が ある か', 28, C.navy);
  twoColCards(s, 1.85,
    '一括 で 払える 場合',
    '→ 一時払・保障重視\n  ・効率 が 良い\n  ・相続税 非課税枠 を 使える\n  ・運用 は 保険会社 任せ\n  ・短期解約 は 不利',
    PROD.ipHosh.color,
    '毎月払い しか できない',
    '→ 変額保険\n  ・少額 から 始められる\n  ・時間分散 で リスク低減\n  ・長期保有 が 前提\n  ・途中解約 は 注意',
    PROD.hengaku.color
  );
  footer(s, slideNo, TOTAL);
}

// ===== Slide 16: 比較③結論 =====
{
  const s = newSlide();
  sectionTitle(s, '比較③：結論');
  emphBox(s, 0.4, 1.0, 9.2, 1.0, '資金 が あれば 一時払 が 効率的\n無ければ 変額 で 時間分散', 24, C.gold);
  const txt =
    '・1,000万円以上の余裕資金 → 一時払（保障）\n' +
    '・現役で月3〜5万円 → 変額\n' +
    '・「半分一時払・半分NISA」 など 併用 も 有効';
  s.addText(jpRuns(txt, 24), { x: 0.6, y: 2.4, w: 8.8, h: 3.5, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 10 , bold: true });
  emphBox(s, 0.4, 6.1, 9.2, 0.6, '資金力 で 支払方法 を 決める', 24, C.navy);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 17: 比較④問題提起 =====
{
  const s = newSlide();
  sectionTitle(s, '比較④：一時払(貯蓄) vs NISA');
  emphBox(s, 0.4, 0.85, 9.2, 0.8, '判断軸：値動き（ブレ） を どう 捉える か', 26, C.navy);
  twoColCards(s, 1.85,
    '一時払・貯蓄重視',
    '★ 予定利率 で 安定\n  → 結果 が 読める\n★ ブレ が 小さい\n  → 元本 を 守る\n★ 期待リターン は 控えめ\n  → 大きくは 増えない\n★ 為替リスク は あり',
    PROD.ipChoSh.color,
    'NISA',
    '★ 値動き が 大きい\n  → 増える 可能性 大\n★ 短期 は 元本割れ も\n  → 心臓 に 悪い\n★ 長期 で 平均化\n  → 20年 で プラス確率 高\n★ 流動性 は 抜群',
    PROD.nisa.color
  );
  footer(s, slideNo, TOTAL);
}

// ===== Slide 18: 比較④結論 =====
{
  const s = newSlide();
  sectionTitle(s, '比較④：結論');
  emphBox(s, 0.4, 1.0, 9.2, 1.0, 'ブレ が 怖い → 一時払(貯蓄)\nブレ に 期待 → NISA', 26, C.green);
  const txt =
    '・60代以降・取り崩しが近い → 一時払（貯蓄）\n' +
    '・40〜50代・長期で運用 → NISA\n' +
    '・「半分ずつ」 で ブレ を 平均化 する 手 も';
  s.addText(jpRuns(txt, 24), { x: 0.6, y: 2.4, w: 8.8, h: 3.5, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 10 , bold: true });
  emphBox(s, 0.4, 6.1, 9.2, 0.6, '安定 ＝ 一時払(貯蓄)  /  期待 ＝ NISA', 24, C.navy);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 19: シミュレーター =====
{
  const s = newSlide();
  sectionTitle(s, '実際 に 数字 で 比べてみよう');
  emphBox(s, 0.4, 1.0, 9.2, 1.0, 'オンライン シミュレーター で\nあなた の 条件 を 入力', 28, C.gold);
  const txt =
    '・月額 ／ 一括 金額 を 入力\n' +
    '・想定利回り を 商品別 に 設定\n' +
    '・10年・20年・30年 の 推移 を 比較\n' +
    '・税金 と 手数料 込み の 手取り 表示';
  s.addText(jpRuns(txt, 24), { x: 0.6, y: 2.6, w: 8.8, h: 3.0, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0.05, paraSpaceAfter: 12 , bold: true });
  emphBox(s, 0.4, 6.0, 9.2, 0.6, '※ 配布 URL から アクセス', 20, C.navy);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 20: 数値例 =====
{
  const s = newSlide();
  sectionTitle(s, '数値例：月3万円 × 20年');
  const headers = ['商品', '元本', '想定利回り', '20年後'];
  const rows = [
    ['変額保険',     '720万',   '4%', '約1,100万'],
    ['一時払(貯蓄)', '一括720', '3%', '約1,300万'],
    ['一時払(保障)', '一括720', '2%', '約1,070万＋保障1,500万'],
    ['NISA',         '720万',   '5%', '約1,230万'],
  ];
  const colW = [2.0, 1.6, 1.8, 3.8];
  const x0 = 0.4;
  let yT = 1.2;
  let cx = x0;
  headers.forEach((h, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: yT, w: colW[i], h: 0.6, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(h, { x: cx, y: yT, w: colW[i], h: 0.6, fontSize: 24, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    cx += colW[i];
  });
  yT += 0.6;
  rows.forEach((r, ri) => {
    let cx2 = x0;
    const rowColor = [PROD.hengaku.light, PROD.ipChoSh.light, PROD.ipHosh.light, PROD.nisa.light][ri];
    r.forEach((v, i) => {
      s.addShape(pres.shapes.RECTANGLE, { x: cx2, y: yT, w: colW[i], h: 0.85, fill: { color: rowColor }, line: { color: C.lightGray } });
      s.addText(jpRuns(v, 24), { x: cx2, y: yT, w: colW[i], h: 0.85, fontFace: F, color: C.bodyText, bold: i === 0, align: 'center', valign: 'middle', margin: 0 });
      cx2 += colW[i];
    });
    yT += 0.85;
  });
  emphBox(s, 0.4, 6.3, 9.2, 0.55, '※ 想定値 です。 実際 は 市場 で 変動', 16, C.gray);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 21: ケース別早見表 =====
{
  const s = newSlide();
  sectionTitle(s, 'ケース別 早見表');
  const cases = [
    { c: '30代・小さい子あり',   a: '変額 ＋ NISA',         col: PROD.hengaku.color },
    { c: '40代・住宅ローン中',   a: '掛捨て ＋ NISA',       col: PROD.nisa.color },
    { c: '50代・教育費 落着',    a: '変額 増額 ＋ NISA',    col: PROD.hengaku.color },
    { c: '60代・退職金 受領',    a: '一時払(貯蓄) ＋ NISA', col: PROD.ipChoSh.color },
    { c: '70代・相続 を 意識',   a: '一時払(保障)',         col: PROD.ipHosh.color },
    { c: '独身・保障 不要',      a: 'NISA フル活用',        col: PROD.nisa.color },
  ];
  cases.forEach((cs, i) => {
    const y = 1.0 + i * 0.95;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 4.6, h: 0.85, fill: { color: C.white }, line: { color: C.lightGray } });
    s.addText(jpRuns(cs.c, 24), { x: 0.5, y, w: 4.4, h: 0.85, fontFace: F, color: C.bodyText, align: 'left', valign: 'middle', margin: 0 , bold: true });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.0, y, w: 4.6, h: 0.85, fill: { color: cs.col }, line: { color: cs.col } });
    s.addText(jpRuns(cs.a, 24), { x: 5.0, y, w: 4.6, h: 0.85, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  });
  footer(s, slideNo, TOTAL);
}

// ===== Slide 22: 追加比較①変額 vs 一時払(貯蓄) =====
{
  const s = newSlide();
  sectionTitle(s, '追加比較：変額 vs 一時払(貯蓄)');
  emphBox(s, 0.4, 0.85, 9.2, 0.8, '保障 不要 の 場合 の 投資手段 選択', 26, C.navy);
  twoColCards(s, 1.85,
    '変額保険',
    '★ 毎月 ＋ 株式運用\n  → 長期 の 上昇 を 取りに行く\n★ 保障 も ついてくる\n  → ただし 「不要」 が 前提',
    PROD.hengaku.color,
    '一時払(貯蓄)',
    '★ 一括 ＋ 利率固定\n  → 結果 が 読める\n★ 為替・据置 に 注意\n  → 短期解約 は 不利',
    PROD.ipChoSh.color
  );
  emphBox(s, 0.4, 6.1, 9.2, 0.6, '保障 不要 なら NISA が 第一候補', 22, C.gold);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 23: 追加比較②相続対策 =====
{
  const s = newSlide();
  sectionTitle(s, '相続対策：一時払(保障) vs 一時払(貯蓄)');
  emphBox(s, 0.4, 0.85, 9.2, 0.8, '判断軸：残す か 自分 で 使う か', 26, C.navy);
  twoColCards(s, 1.85,
    '一時払・保障重視',
    '★ 死亡保険金 1.3〜2倍\n  → 大きく 残せる\n★ 生命保険 非課税枠\n  → 500万 × 法定相続人\n★ 受取人指定 が 自由',
    PROD.ipHosh.color,
    '一時払・貯蓄重視',
    '★ 自分 が 受取 前提\n  → 老後資金 として 使う\n★ 死亡保険金 は 払込相当\n  → 相続効果 は 限定的\n★ 解約 で 取り戻せる',
    PROD.ipChoSh.color
  );
  emphBox(s, 0.4, 6.1, 9.2, 0.6, '残す ＝ 保障重視  /  使う ＝ 貯蓄重視', 22, C.gold);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 24: 組み合わせ戦略 =====
{
  const s = newSlide();
  sectionTitle(s, '組み合わせ戦略：3つの王道');
  const combos = [
    { t: '① NISA ＋ 掛捨て保険',     d: '効率 と 保障 を 別々 に 確保\nコスト最小 ・ 自由度 最大',   c: PROD.nisa.color },
    { t: '② NISA ＋ 変額',          d: '増やす ＋ 強制積立 ＋ 保障\n投資慣れ してきた 人 向け',     c: PROD.hengaku.color },
    { t: '③ 一時払(貯蓄) ＋ NISA',  d: '退職金 を 半分 安定 ／ 半分 成長\n60代 の 王道',              c: PROD.ipChoSh.color },
  ];
  combos.forEach((cb, i) => {
    const y = 1.1 + i * 1.85;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.6, fill: { color: cb.c }, line: { color: cb.c } });
    s.addText(jpRuns(cb.t, 24), { x: 0.4, y, w: 9.2, h: 0.6, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: y + 0.6, w: 9.2, h: 1.1, fill: { color: C.white }, line: { color: cb.c, width: 1.5 } });
    s.addText(jpRuns(cb.d, 24), { x: 0.6, y: y + 0.6, w: 9.0, h: 1.1, fontFace: F, color: C.bodyText, align: 'left', valign: 'middle', margin: 0.05 , bold: true });
  });
  emphBox(s, 0.4, 6.7, 9.2, 0.5, '一本 で なく 組み合わせ で 弱点 を 補う', 18, C.navy);
  footer(s, slideNo, TOTAL);
}

// ===== Slide 25: FAQ =====
{
  const s = newSlide();
  sectionTitle(s, 'よくある質問');
  const faqs = [
    { q: 'Q. 変額保険は元本割れしますか？',
      a: 'A. します。 ただし 死亡保険金 は 最低保証 あり。 解約返戻金 は 変動。' },
    { q: 'Q. NISAだけで保障は足りますか？',
      a: 'A. 扶養家族 が いる場合 は 別途 掛捨て を 推奨。 NISA は 投資制度 で 保障 なし。' },
    { q: 'Q. 一時払いは いつ 受け取れますか？',
      a: 'A. 商品 により 据置3〜10年。 期間内 解約 は 元本割れ リスク。' },
    { q: 'Q. 全部 同時 に 始めて いい？',
      a: 'A. 可。 ただし 家計 と 目的 を 整理 してから。 個別相談 を 推奨。' },
  ];
  faqs.forEach((f, i) => {
    const y = 0.85 + i * 1.5;
    s.addText(jpRuns(f.q, 24), { x: 0.4, y, w: 9.2, h: 0.5, fontFace: F, bold: true, color: C.navy, align: 'left', valign: 'middle', margin: 0 });
    s.addText(jpRuns(f.a, 24), { x: 0.6, y: y + 0.5, w: 9.0, h: 0.9, fontFace: F, color: C.bodyText, align: 'left', valign: 'top', margin: 0 , bold: true });
  });
  footer(s, slideNo, TOTAL);
}

// ===== Slide 26: クロージング =====
{
  const s = newSlide(C.navy);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 3.3, w: 10, h: 0.06, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText(jpRuns('「保険 か NISA か」 の 二択 で 終わらない\nあなた に 合う 組み合わせ が ある', 30), {
    x: 0.5, y: 1.2, w: 9, h: 2.0, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0
  });
  s.addText(jpRuns('まず 「保障 が 必要 か」 から 決めましょう', 24), {
    x: 0.5, y: 3.7, w: 9, h: 1.0, fontFace: F, color: C.lightGray, align: 'center', valign: 'middle', margin: 0
  });
  s.addText('Thank you for joining', {
    x: 0.5, y: 5.2, w: 9, h: 0.5, fontSize: 24, fontFace: FA, italic: true, color: C.gold, align: 'center', margin: 0
  });
}

// ===== Slide 27: 個別相談 =====
{
  const s = newSlide();
  sectionTitle(s, '個別相談・特典のご案内');
  emphBox(s, 0.4, 1.0, 9.2, 1.0, '本日参加 の 方 限定\n無料 個別相談 60分 プレゼント', 26, C.gold);
  const lines = [
    '★ あなた の 家計 と 目的 を ヒアリング',
    '★ 4商品 の 中から 最適 な 組み合わせ を 提案',
    '★ シミュレーター で 数字 を 確認',
    '★ 強引 な 勧誘 は 一切 ありません',
  ];
  lines.forEach((ln, i) => {
    s.addText(jpRuns(ln, 24), { x: 0.6, y: 2.5 + i * 0.6, w: 8.8, h: 0.5, fontFace: F, color: C.bodyText, align: 'left', valign: 'middle', margin: 0 , bold: true });
  });
  emphBox(s, 0.4, 5.4, 9.2, 0.8, 'お申込み は 受付 の QR から', 26, C.gold);
  s.addText('Money Navigation Seminar', {
    x: 0.5, y: 6.5, w: 9, h: 0.4, fontSize: 14, fontFace: FA, italic: true, color: C.gray, align: 'center', margin: 0
  });
  footer(s, slideNo, TOTAL);
}

pres.writeFile({ fileName: 'C:/Users/love2/Desktop/Claude/セミナー_保険vsNISA/seminar_slides.pptx' })
  .then(fn => console.log('Generated:', fn, '/ slides:', slideNo));
