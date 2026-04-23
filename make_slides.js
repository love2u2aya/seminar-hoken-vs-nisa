const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = '一時払い保険とNISAの賢い使い分け';

// ── カラーパレット ──
const C = {
  navy:       '1a365d',
  navyMid:    '2a4a7f',
  navyLight:  '3d5a99',
  green:      '276749',
  greenMid:   '2d7d46',
  greenLight: '48bb78',
  gold:       'D4A017',
  goldLight:  'F6E05E',
  white:      'FFFFFF',
  offWhite:   'F7FAFC',
  lightBg:    'EDF2F7',
  gray:       '718096',
  grayLight:  'CBD5E0',
  darkText:   '1a202c',
  bodyText:   '2d3748',
  red:        'C53030',
  redLight:   'FED7D7',
  blueCard:   'EBF8FF',
  greenCard:  'F0FFF4',
  goldCard:   'FFFFF0',
};
const F = 'Meiryo';
const mkSh = () => ({ type: 'outer', blur: 6, offset: 2, angle: 135, color: '000000', opacity: 0.10 });

// ── ヘルパー ──
function darkBg(s) { s.background = { color: C.navy }; }
function lightBg(s) { s.background = { color: C.offWhite }; }

function sectionTitle(s, title, color) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 0.22, w: 0.07, h: 0.65,
    fill: { color: color || C.green }, line: { color: color || C.green }
  });
  s.addText(title, {
    x: 0.65, y: 0.22, w: 9.1, h: 0.65,
    fontSize: 24, fontFace: F, bold: true, color: C.navy, margin: 0, valign: 'middle'
  });
}

function topBar(s, color) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.07,
    fill: { color: color || C.gold }, line: { color: color || C.gold }
  });
}
function bottomBar(s, color) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.525, w: 10, h: 0.1,
    fill: { color: color || C.green }, line: { color: color || C.green }
  });
}

function card(s, x, y, w, h, bg) {
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: bg || C.white },
    line: { color: C.grayLight, width: 1 },
    shadow: mkSh()
  });
}

function caseTag(s, label, color) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 1.02, w: 1.6, h: 0.38,
    fill: { color: color || C.green }, line: { color: color || C.green }
  });
  s.addText(label, {
    x: 0.45, y: 1.02, w: 1.6, h: 0.38,
    fontSize: 13, fontFace: F, bold: true, color: C.white,
    align: 'center', valign: 'middle', margin: 0
  });
}

function twoColCards(s, leftTitle, leftItems, rightTitle, rightItems, leftBg, rightBg) {
  // Left card
  card(s, 0.4, 1.0, 4.4, 3.9, leftBg || C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.0, w: 4.4, h: 0.5, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText(leftTitle, { x: 0.4, y: 1.0, w: 4.4, h: 0.5, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText(leftItems, { x: 0.55, y: 1.6, w: 4.1, h: 3.1, fontSize: 14, fontFace: F, color: C.bodyText, valign: 'top', lineSpacingMultiple: 1.4 });

  // Right card
  card(s, 5.2, 1.0, 4.4, 3.9, rightBg || C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.0, w: 4.4, h: 0.5, fill: { color: C.green }, line: { color: C.green } });
  s.addText(rightTitle, { x: 5.2, y: 1.0, w: 4.4, h: 0.5, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText(rightItems, { x: 5.35, y: 1.6, w: 4.1, h: 3.1, fontSize: 14, fontFace: F, color: C.bodyText, valign: 'top', lineSpacingMultiple: 1.4 });
}

// ── SLIDE 1: タイトル ──────────────────────────────────────────
{
  const s = pres.addSlide();
  darkBg(s);
  topBar(s, C.gold);

  s.addText('一時払い保険とNISAの賢い使い分け', {
    x: 0.6, y: 1.0, w: 8.8, h: 1.6,
    fontSize: 40, fontFace: F, bold: true, color: C.white,
    align: 'center', valign: 'middle'
  });
  s.addText('同じ金額・同じ10年。何が違う？', {
    x: 0.6, y: 2.75, w: 8.8, h: 0.65,
    fontSize: 22, fontFace: F, color: C.goldLight,
    align: 'center'
  });
  s.addShape(pres.shapes.LINE, { x: 2.8, y: 3.55, w: 4.4, h: 0, line: { color: C.navyLight, width: 1 } });

  s.addShape(pres.shapes.RECTANGLE, { x: 2.4, y: 3.7, w: 2.4, h: 0.45, fill: { color: C.navyMid }, line: { color: C.navyLight, width: 1 } });
  s.addText('講師名', { x: 2.4, y: 3.7, w: 2.4, h: 0.45, fontSize: 13, fontFace: F, color: '8899bb', align: 'center', valign: 'middle', margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.7, w: 2.4, h: 0.45, fill: { color: C.navyMid }, line: { color: C.navyLight, width: 1 } });
  s.addText('日付', { x: 5.2, y: 3.7, w: 2.4, h: 0.45, fontSize: 13, fontFace: F, color: '8899bb', align: 'center', valign: 'middle', margin: 0 });

  bottomBar(s, C.green);
}

// ── SLIDE 2: 今日のゴール ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '今日わかること　3つ');

  const items = [
    { num: '①', text: 'なぜ今、ドル資産を持つべきか' },
    { num: '②', text: '一時払い保険とNISA、何が同じで何が違うか' },
    { num: '③', text: 'あなたの状況に合った使い分け方' },
  ];
  items.forEach((item, i) => {
    const y = 1.25 + i * 1.2;
    card(s, 0.5, y, 9.0, 0.95, C.white);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.65, h: 0.95, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(item.num, { x: 0.5, y, w: 0.65, h: 0.95, fontSize: 18, fontFace: F, bold: true, color: C.goldLight, align: 'center', valign: 'middle', margin: 0 });
    s.addText(item.text, { x: 1.3, y: y + 0.05, w: 8.0, h: 0.85, fontSize: 18, fontFace: F, bold: true, color: C.bodyText, valign: 'middle' });
  });
}

// ── SLIDE 3: あなたの本音 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'こんな気持ち、ありませんか？');

  card(s, 0.5, 1.15, 9.0, 1.35, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 0.18, h: 1.35, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('「子どもの学費が最優先…でも老後もそろそろ心配。\nどうすればいいの？」', {
    x: 0.85, y: 1.2, w: 8.4, h: 1.25,
    fontSize: 20, fontFace: F, bold: true, color: C.navy,
    valign: 'middle', lineSpacingMultiple: 1.5
  });

  card(s, 0.5, 2.7, 9.0, 1.0, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.7, w: 0.18, h: 1.0, fill: { color: C.green }, line: { color: C.green } });
  s.addText('「NISAって今から間に合うの？　損したら怖い…」', {
    x: 0.85, y: 2.75, w: 8.4, h: 0.9,
    fontSize: 20, fontFace: F, bold: true, color: C.green,
    valign: 'middle'
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.9, w: 9.0, h: 0.55, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('→ 今日のセミナーで「やることが見えた」状態にします', {
    x: 0.5, y: 3.9, w: 9.0, h: 0.55,
    fontSize: 15, fontFace: F, bold: true, color: C.goldLight,
    align: 'center', valign: 'middle', margin: 0
  });
}

// ── SLIDE 4: 時間はある ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '50代でも、まだ30〜40年ある');

  // タイムライン
  const tl_y = 2.3, tl_x = 0.7, tl_w = 8.6;
  s.addShape(pres.shapes.RECTANGLE, { x: tl_x, y: tl_y + 0.18, w: tl_w, h: 0.06, fill: { color: C.grayLight }, line: { color: C.grayLight } });

  const points = [
    { x: tl_x,            label: '現在\n40〜50代', color: C.navy },
    { x: tl_x + tl_w*0.38, label: '65歳\n老後スタート', color: C.green },
    { x: tl_x + tl_w*0.75, label: '80代', color: C.gray },
    { x: tl_x + tl_w,     label: '90代', color: C.grayLight },
  ];
  points.forEach(p => {
    s.addShape(pres.shapes.OVAL, { x: p.x - 0.15, y: tl_y + 0.04, w: 0.38, h: 0.38, fill: { color: p.color }, line: { color: p.color } });
    s.addText(p.label, { x: p.x - 0.5, y: tl_y + 0.55, w: 1.1, h: 0.7, fontSize: 12, fontFace: F, color: C.bodyText, align: 'center' });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: tl_x, y: tl_y + 0.21, w: tl_w * 0.75, h: 0.0, fill: { color: C.navy }, line: { color: C.navy, width: 3 } });

  card(s, 1.0, 3.6, 8.0, 0.85, C.goldCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 3.6, w: 0.18, h: 0.85, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('「時間がない」のではなく「時間の使い方を変える」タイミング', {
    x: 1.35, y: 3.65, w: 7.5, h: 0.75,
    fontSize: 16, fontFace: F, bold: true, color: C.navy, valign: 'middle'
  });
}

// ── SLIDE 5: ドル資産① ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '円高か円安か、誰にもわからない');

  card(s, 0.4, 1.1, 3.6, 3.5, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 3.6, h: 0.5, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('円高になったら？', { x: 0.4, y: 1.1, w: 3.6, h: 0.5, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('円資産が有利\n日本円で持っていると\n購買力が上がる', { x: 0.55, y: 1.75, w: 3.3, h: 2.7, fontSize: 15, fontFace: F, color: C.bodyText, valign: 'middle', lineSpacingMultiple: 1.5 });

  card(s, 5.95, 1.1, 3.6, 3.5, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.95, y: 1.1, w: 3.6, h: 0.5, fill: { color: C.green }, line: { color: C.green } });
  s.addText('円安になったら？', { x: 5.95, y: 1.1, w: 3.6, h: 0.5, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('外貨資産が有利\n円建て資産の価値が\n目減りする', { x: 6.1, y: 1.75, w: 3.3, h: 2.7, fontSize: 15, fontFace: F, color: C.bodyText, valign: 'middle', lineSpacingMultiple: 1.5 });

  // 中央
  s.addShape(pres.shapes.RECTANGLE, { x: 4.05, y: 1.85, w: 1.85, h: 1.5, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('だから\n「両方持つ」\n＝分散', { x: 4.05, y: 1.85, w: 1.85, h: 1.5, fontSize: 14, fontFace: F, bold: true, color: C.goldLight, align: 'center', valign: 'middle', margin: 0 });
}

// ── SLIDE 6: ドル資産② ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '日本は輸入大国。円安は生活を直撃する');

  const flow = [
    { label: '食料・エネルギー・部品', bg: C.blueCard, tc: C.navy },
    { label: '外貨（ドル）で購入', bg: C.navy, tc: C.white },
    { label: '円安になると\nコスト上昇', bg: C.red, tc: C.white },
    { label: '物価上昇・\n生活費増加', bg: '7B2D2D', tc: C.white },
  ];
  flow.forEach((f, i) => {
    const x = 0.4 + i * 2.38;
    card(s, x, 1.15, 2.1, 1.5, f.bg);
    s.addText(f.label, { x, y: 1.15, w: 2.1, h: 1.5, fontSize: 14, fontFace: F, bold: true, color: f.tc, align: 'center', valign: 'middle', margin: 8 });
    if (i < 3) {
      s.addShape(pres.shapes.RECTANGLE, { x: x + 2.1, y: 1.7, w: 0.28, h: 0.38, fill: { color: C.gold }, line: { color: C.gold } });
      s.addText('▶', { x: x + 2.1, y: 1.7, w: 0.28, h: 0.38, fontSize: 14, fontFace: F, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    }
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.1, w: 9.15, h: 1.4, fill: { color: C.navy }, line: { color: C.navy } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.1, w: 0.12, h: 1.4, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('だからこそ、資産の一部を外貨で持つことが「生活の守り」になる\n\n今日紹介する一時払い保険もNISA（オルカン）も、どちらもドル建て・外貨資産として機能する', {
    x: 0.65, y: 3.15, w: 8.8, h: 1.3,
    fontSize: 14, fontFace: F, color: C.white, valign: 'top', lineSpacingMultiple: 1.4
  });
}

// ── SLIDE 7: NISAを整理する ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'NISAとは？　1枚で整理');

  // 中央の大きなボックス
  card(s, 0.5, 1.1, 9.0, 1.2, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 1.5, h: 1.2, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('NISA\n口座', { x: 0.5, y: 1.1, w: 1.5, h: 1.2, fontSize: 16, fontFace: F, bold: true, color: C.goldLight, align: 'center', valign: 'middle', margin: 0 });
  s.addText('運用益に税金ゼロ　（通常は利益の約20%が税金）', {
    x: 2.2, y: 1.2, w: 7.1, h: 1.0,
    fontSize: 18, fontFace: F, bold: true, color: C.navy, valign: 'middle'
  });

  // 3つの特徴
  const feats = [
    { title: '毎月積立', body: '少額から\nコツコツ積立' },
    { title: '長期運用', body: '最長20年\n非課税で保有' },
    { title: '世界分散\n（オルカン）', body: '全世界株式に\n自動分散' },
  ];
  feats.forEach((f, i) => {
    const x = 0.5 + i * 3.05;
    card(s, x, 2.6, 2.8, 2.6, C.white);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.6, w: 2.8, h: 0.55, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(f.title, { x, y: 2.6, w: 2.8, h: 0.55, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    s.addText(f.body, { x, y: 3.2, w: 2.8, h: 1.9, fontSize: 16, fontFace: F, color: C.bodyText, align: 'center', valign: 'middle' });
  });
}

// ── SLIDE 8: NISAの特徴 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'NISA（オルカン）の特徴');

  // 3つの数字カード
  const stats = [
    { val: '約8%', sub: '想定利率（年率）', color: C.green },
    { val: '1.5倍', sub: '10年後（目安）\n500万→750万', color: C.greenMid },
    { val: 'ブレあり', sub: '大きく上下する\n可能性がある', color: C.red },
  ];
  stats.forEach((st, i) => {
    const x = 0.45 + i * 3.05;
    card(s, x, 1.05, 2.85, 2.0, C.white);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.05, w: 2.85, h: 0.45, fill: { color: st.color }, line: { color: st.color } });
    s.addText(st.val, { x, y: 1.5, w: 2.85, h: 0.95, fontSize: 34, fontFace: F, bold: true, color: st.color, align: 'center', valign: 'middle', margin: 0 });
    s.addText(st.sub, { x, y: 2.55, w: 2.85, h: 0.55, fontSize: 12, fontFace: F, color: C.gray, align: 'center', valign: 'top' });
  });

  // ブレのイメージ（折れ線チャート風）
  const chartData = [
    { name: '楽観', labels: ['0', '2', '4', '6', '8', '10'], values: [500, 610, 740, 900, 1090, 1320] },
    { name: '中央値', labels: ['0', '2', '4', '6', '8', '10'], values: [500, 540, 583, 630, 680, 735] },
    { name: '悲観', labels: ['0', '2', '4', '6', '8', '10'], values: [500, 480, 460, 440, 420, 400] },
  ];
  s.addChart(pres.charts.LINE, chartData, {
    x: 0.45, y: 3.2, w: 9.1, h: 2.2,
    chartColors: [C.greenLight, C.green, C.red],
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    lineSize: 2, lineSmooth: true,
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray,
    valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
    showLegend: true, legendPos: 'r', legendFontSize: 10,
    showTitle: false,
    valAxisMinVal: 300, valAxisMaxVal: 1400,
  });
}

// ── SLIDE 9: 一時払い保険を整理する ──────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '一時払い保険とは？　1枚で整理');

  const flow2 = [
    { label: 'まとまった\nお金を\n一括払い', bg: C.navy, tc: C.white },
    { label: 'ドルで\n複利運用', bg: C.green, tc: C.white },
    { label: '解約返戻金\nとして受取\n（確実に増える）', bg: C.blueCard, tc: C.navy },
  ];
  flow2.forEach((f, i) => {
    const x = 0.6 + i * 3.15;
    card(s, x, 1.2, 2.8, 3.2, f.bg);
    s.addText(f.label, { x, y: 1.2, w: 2.8, h: 3.2, fontSize: 18, fontFace: F, bold: true, color: f.tc, align: 'center', valign: 'middle', margin: 12 });
    if (i < 2) {
      s.addShape(pres.shapes.RECTANGLE, { x: x + 2.8, y: 2.5, w: 0.35, h: 0.55, fill: { color: C.gold }, line: { color: C.gold } });
      s.addText('▶', { x: x + 2.8, y: 2.5, w: 0.35, h: 0.55, fontSize: 18, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    }
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.7, w: 8.85, h: 0.55, fill: { color: C.goldCard.replace('#','') === C.goldCard ? C.goldCard : 'FFFFF0' }, line: { color: C.gold } });
  s.addText('「保険」という名前がついているが、今日は「増やす手段」として考える', {
    x: 0.6, y: 4.7, w: 8.85, h: 0.55,
    fontSize: 14, fontFace: F, bold: true, color: C.navy, align: 'center', valign: 'middle', margin: 0
  });
}

// ── SLIDE 10: 一時払い保険の特徴 ──────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '一時払い保険（ドル建て）の特徴');

  const stats2 = [
    { val: '約4%', sub: '利率（年率）', color: C.navy },
    { val: '1.5倍', sub: '10年後（確実）\n500万→750万', color: C.navy },
    { val: 'ブレなし', sub: '確実に増える\n安心の右肩上がり', color: C.green },
  ];
  stats2.forEach((st, i) => {
    const x = 0.45 + i * 3.05;
    card(s, x, 1.05, 2.85, 2.0, C.white);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.05, w: 2.85, h: 0.45, fill: { color: st.color }, line: { color: st.color } });
    s.addText(st.val, { x, y: 1.5, w: 2.85, h: 0.95, fontSize: 34, fontFace: F, bold: true, color: st.color, align: 'center', valign: 'middle', margin: 0 });
    s.addText(st.sub, { x, y: 2.55, w: 2.85, h: 0.55, fontSize: 12, fontFace: F, color: C.gray, align: 'center', valign: 'top' });
  });

  // 確実な右肩上がりライン
  const insData = [{ name: '一時払い保険', labels: ['0', '2', '4', '6', '8', '10'], values: [500, 541, 584, 631, 682, 735] }];
  s.addChart(pres.charts.LINE, insData, {
    x: 0.45, y: 3.2, w: 9.1, h: 2.2,
    chartColors: [C.navy],
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    lineSize: 3, lineSmooth: true,
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray,
    valGridLine: { color: 'E2E8F0', size: 0.5 }, catGridLine: { style: 'none' },
    showLegend: false, showTitle: false,
    valAxisMinVal: 450, valAxisMaxVal: 800,
    showValue: false,
  });
}

// ── SLIDE 11: 並べてみる ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'どちらも10年で1.5倍。違いは「ブレ」だけ');

  const headers = [
    [{ text: '', options: { fill: { color: 'F7FAFC' } } },
     { text: '一時払い保険', options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 15 } },
     { text: 'NISA（つみたて）', options: { fill: { color: C.green }, color: C.white, bold: true, fontSize: 15 } }],
    [{ text: '利回り', options: { fill: { color: 'EDF2F7' }, bold: true, fontSize: 14, color: C.bodyText } },
     { text: '約4%', options: { fontSize: 14, color: C.navy, bold: true } },
     { text: '約8%', options: { fontSize: 14, color: C.green, bold: true } }],
    [{ text: '10年後', options: { fill: { color: 'EDF2F7' }, bold: true, fontSize: 14, color: C.bodyText } },
     { text: '約1.5倍（確実）', options: { fontSize: 14, color: C.navy, bold: true } },
     { text: '約1.5倍（目安）', options: { fontSize: 14, color: C.green } }],
    [{ text: 'ブレ', options: { fill: { color: 'EDF2F7' }, bold: true, fontSize: 14, color: C.bodyText } },
     { text: 'なし', options: { fontSize: 14, color: C.navy, bold: true } },
     { text: 'あり（大きく）', options: { fontSize: 14, color: C.red, bold: true } }],
    [{ text: '流動性', options: { fill: { color: 'EDF2F7' }, bold: true, fontSize: 14, color: C.bodyText } },
     { text: '低い（途中解約は損）', options: { fontSize: 13, color: C.gray } },
     { text: '高い（いつでも売れる）', options: { fontSize: 13, color: C.gray } }],
  ];
  s.addTable(headers, {
    x: 0.5, y: 1.05, w: 9.0, h: 3.4,
    colW: [2.2, 3.4, 3.4],
    border: { pt: 1, color: 'CBD5E0' },
    align: 'center', valign: 'middle',
    rowH: 0.65,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.65, w: 9.0, h: 0.65, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('「そのお金、減ったら困りますか？」', {
    x: 0.5, y: 4.65, w: 9.0, h: 0.65,
    fontSize: 20, fontFace: F, bold: true, color: C.goldLight,
    align: 'center', valign: 'middle', margin: 0
  });
}

// ── SLIDE 12: デモ① ──────────────────────────────────────
{
  const s = pres.addSlide();
  darkBg(s);
  topBar(s, C.gold);

  s.addText('実際に動かしてみましょう', {
    x: 0.5, y: 0.8, w: 9, h: 0.9,
    fontSize: 34, fontFace: F, bold: true, color: C.white, align: 'center'
  });

  card(s, 1.0, 1.95, 8.0, 1.6, C.navyMid);
  s.addText('500万円・10年\n保険 4%  vs  NISA 8%  で比べると…？', {
    x: 1.0, y: 1.95, w: 8.0, h: 1.6,
    fontSize: 22, fontFace: F, bold: true, color: C.goldLight,
    align: 'center', valign: 'middle', lineSpacingMultiple: 1.5
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: 3.8, w: 5.0, h: 0.7, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('▶　シミュレーター　デモ', {
    x: 2.5, y: 3.8, w: 5.0, h: 0.7,
    fontSize: 20, fontFace: F, bold: true, color: C.navy,
    align: 'center', valign: 'middle', margin: 0
  });

  bottomBar(s, C.green);
}

// ── SLIDE 13: 使い分けの鍵 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '使い分けの判断は、この一問');

  card(s, 0.5, 1.1, 9.0, 1.2, C.navy);
  s.addText('「そのお金、10年後に減っていたら困りますか？」', {
    x: 0.5, y: 1.1, w: 9.0, h: 1.2,
    fontSize: 22, fontFace: F, bold: true, color: C.goldLight,
    align: 'center', valign: 'middle', margin: 0
  });

  // YES
  card(s, 0.5, 2.6, 4.3, 2.7, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 0.5, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('YES（困る）', { x: 0.5, y: 2.6, w: 4.3, h: 0.5, fontSize: 15, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('一時払い保険\nブレない 1.5倍', { x: 0.5, y: 3.15, w: 4.3, h: 2.0, fontSize: 20, fontFace: F, bold: true, color: C.navy, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5 });

  // NO
  card(s, 5.2, 2.6, 4.3, 2.7, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 0.5, fill: { color: C.green }, line: { color: C.green } });
  s.addText('NO（大丈夫）', { x: 5.2, y: 2.6, w: 4.3, h: 0.5, fontSize: 15, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('NISA\nブレても待てる', { x: 5.2, y: 3.15, w: 4.3, h: 2.0, fontSize: 20, fontFace: F, bold: true, color: C.green, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5 });
}

// ── SLIDE 14: 一時払いの入り方 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '一時払い保険の入り方');

  const pts = [
    { num: '1', title: '期間は10年 or 20年で固定', body: '使わない自信がある期間を選ぶ\n「10年は使わない」なら10年。「20年大丈夫」なら20年。' },
    { num: '2', title: '長いほど有利な利率が多い', body: '20年の方が10年より高い利率が設定されていることが一般的。\n長く固定する価値がある。' },
    { num: '3', title: '不利になったら乗り換えでもOK', body: '途中解約は「市場価格調整」が発生する場合あり。\nそれでも条件が良ければ乗り換えを検討してよい。' },
  ];
  pts.forEach((p, i) => {
    const y = 1.1 + i * 1.35;
    card(s, 0.5, y, 9.0, 1.2, C.white);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.7, h: 1.2, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(p.num, { x: 0.5, y, w: 0.7, h: 1.2, fontSize: 24, fontFace: F, bold: true, color: C.goldLight, align: 'center', valign: 'middle', margin: 0 });
    s.addText(p.title, { x: 1.35, y: y + 0.05, w: 8.0, h: 0.4, fontSize: 15, fontFace: F, bold: true, color: C.navy, valign: 'middle' });
    s.addText(p.body, { x: 1.35, y: y + 0.48, w: 8.0, h: 0.65, fontSize: 13, fontFace: F, color: C.gray, valign: 'top' });
  });
}

// ── SLIDE 15: シミュレーション数値 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '500万円を入れたら？');

  const tableRows = [
    [{ text: '', options: { fill: { color: 'EDF2F7' } } },
     { text: '10年後', options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 14 } },
     { text: '20年後', options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 14 } }],
    [{ text: '一時払い保険（4%）', options: { fill: { color: C.blueCard }, bold: true, fontSize: 13, color: C.navy } },
     { text: '約750万円\n＋250万円増', options: { fontSize: 18, bold: true, color: C.navy } },
     { text: '約1,095万円\n＋595万円増', options: { fontSize: 18, bold: true, color: C.navy } }],
    [{ text: 'NISA 中央値（8%）', options: { fill: { color: C.greenCard }, bold: true, fontSize: 13, color: C.green } },
     { text: '約750万円\n（元本360万円）', options: { fontSize: 18, bold: true, color: C.green } },
     { text: '約1,764万円\n（元本720万円）', options: { fontSize: 18, bold: true, color: C.green } }],
  ];
  s.addTable(tableRows, {
    x: 0.5, y: 1.1, w: 9.0, h: 3.4,
    colW: [3.0, 3.0, 3.0],
    border: { pt: 1, color: 'CBD5E0' },
    align: 'center', valign: 'middle', rowH: 1.0,
  });

  s.addText('※ NISAは積立総額＝保険元本と同額の条件で試算　／　NISAの実際の成果は大きくブレます', {
    x: 0.5, y: 4.7, w: 9.0, h: 0.5,
    fontSize: 11, fontFace: F, color: C.gray, align: 'center'
  });
}

// ── SLIDE 16: ケースA 状況 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースA　子どもが大学生');
  caseTag(s, 'CASE A', C.navy);

  card(s, 0.5, 1.55, 9.0, 1.65, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.55, w: 0.15, h: 1.65, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('状況：学費の総額がほぼ見えてきた\n残り期間は1〜4年。ゴールが近い。', {
    x: 0.8, y: 1.6, w: 8.5, h: 1.55,
    fontSize: 18, fontFace: F, color: C.bodyText, valign: 'middle', lineSpacingMultiple: 1.5
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.45, w: 9.0, h: 0.55, fill: { color: C.navyMid }, line: { color: C.navyMid } });
  s.addText('問：学費を払い終えた後、いくら残る？　それをどう老後に活かす？', {
    x: 0.5, y: 3.45, w: 9.0, h: 0.55,
    fontSize: 15, fontFace: F, bold: true, color: C.goldLight, align: 'center', valign: 'middle', margin: 0
  });
}

// ── SLIDE 17: ケースA 結論 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースA　結論');
  caseTag(s, 'CASE A', C.navy);

  twoColCards(s,
    'お金に余裕がある方',
    '保険で確実に増やし\n「使い始める」計画を立てる\n\n・老後の旅行・医療費・リフォーム\n・減らしたくないお金は保険へ\n・NISAは余裕資金として並行',
    'お金が足りない方',
    '覚悟を持って\nNISAで増やすフェーズ\n\n・勉強して始める\n・少額でもOK、まず動く\n・ブレを受け入れる覚悟が必要'
  );
}

// ── SLIDE 18: ケースB 状況 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースB　子どもが中学生');
  caseTag(s, 'CASE B', C.green);

  card(s, 0.5, 1.55, 9.0, 1.65, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.55, w: 0.15, h: 1.65, fill: { color: C.green }, line: { color: C.green } });
  s.addText('状況：進路がまだ読めない。でも時間はある（5〜10年）\n文系か理系か、どの大学か、まだわからない。', {
    x: 0.8, y: 1.6, w: 8.5, h: 1.55,
    fontSize: 18, fontFace: F, color: C.bodyText, valign: 'middle', lineSpacingMultiple: 1.5
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.45, w: 9.0, h: 0.55, fill: { color: C.green }, line: { color: C.green } });
  s.addText('この期間をどう使うかが、老後を左右する', {
    x: 0.5, y: 3.45, w: 9.0, h: 0.55,
    fontSize: 17, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0
  });
}

// ── SLIDE 19: 奨学金 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '「奨学金」は借金ではなく、武器になる');

  // 数字強調
  const nums = [
    { val: '上限3%', sub: '奨学金の上限金利\n（第二種）' },
    { val: '最長20年', sub: '返済可能期間' },
  ];
  nums.forEach((n, i) => {
    card(s, 0.5 + i * 4.6, 1.05, 4.2, 1.5, C.blueCard);
    s.addText(n.val, { x: 0.5 + i * 4.6, y: 1.05, w: 4.2, h: 0.9, fontSize: 32, fontFace: F, bold: true, color: C.navy, align: 'center', valign: 'middle' });
    s.addText(n.sub, { x: 0.5 + i * 4.6, y: 1.95, w: 4.2, h: 0.6, fontSize: 12, fontFace: F, color: C.gray, align: 'center' });
  });

  // フロー
  const fl = [
    { label: '子が奨学金\nで借りる\n（〜3%）', bg: C.navy, tc: C.white },
    { label: '親はその分\n保険・NISA\nで運用（4〜8%）', bg: C.green, tc: C.white },
    { label: '利率の差額\nが家族の\n利益になる', bg: C.gold, tc: C.navy },
  ];
  fl.forEach((f, i) => {
    const x = 0.5 + i * 3.08;
    card(s, x, 2.75, 2.75, 2.0, f.bg);
    s.addText(f.label, { x, y: 2.75, w: 2.75, h: 2.0, fontSize: 14, fontFace: F, bold: true, color: f.tc, align: 'center', valign: 'middle', margin: 8 });
    if (i < 2) {
      s.addShape(pres.shapes.RECTANGLE, { x: x + 2.75, y: 3.5, w: 0.33, h: 0.48, fill: { color: C.grayLight }, line: { color: C.grayLight } });
      s.addText('▶', { x: x + 2.75, y: 3.5, w: 0.33, h: 0.48, fontSize: 14, fontFace: F, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    }
  });

  s.addText('子どもにNISAを覚えさせ、奨学金返済資金として積み立てる', {
    x: 0.5, y: 4.92, w: 9.0, h: 0.45,
    fontSize: 13, fontFace: F, bold: true, color: C.green, align: 'center'
  });
}

// ── SLIDE 20: ケースB 結論 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースB　結論');
  caseTag(s, 'CASE B', C.green);

  twoColCards(s,
    'お金に余裕がある方',
    '老後・夫死後まで\nしっかり試算する\n\n・長期保険（20年）も選択肢に\n・NISAも並行して余裕資金に\n・「意外とお金がいる」を早めに認識',
    'お金が足りない方',
    '保険に入れすぎない\n流動性を確保\n\n・末子8歳以下なら10年保険で\n  大学費用の土台を確実に準備\n・子にNISAを覚えさせ\n  奨学金返済資金として活用'
  );
}

// ── SLIDE 21: デモ② ──────────────────────────────────────
{
  const s = pres.addSlide();
  darkBg(s);
  topBar(s, C.gold);

  s.addText('あなたのケースで動かしてみましょう', {
    x: 0.5, y: 0.8, w: 9, h: 0.9,
    fontSize: 30, fontFace: F, bold: true, color: C.white, align: 'center'
  });

  const cases = [
    { label: 'CASE A', text: '1,000万円・20年で保険に入れたら？' },
    { label: 'CASE B', text: '500万円・10年の保険 vs NISA比較' },
    { label: 'CASE C', text: '月5万円・20年のNISA積立は？' },
  ];
  cases.forEach((c, i) => {
    card(s, 0.5, 1.95 + i * 1.05, 8.95, 0.85, C.navyMid);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.95 + i * 1.05, w: 1.1, h: 0.85, fill: { color: C.gold }, line: { color: C.gold } });
    s.addText(c.label, { x: 0.5, y: 1.95 + i * 1.05, w: 1.1, h: 0.85, fontSize: 13, fontFace: F, bold: true, color: C.navy, align: 'center', valign: 'middle', margin: 0 });
    s.addText(c.text, { x: 1.75, y: 1.95 + i * 1.05, w: 7.5, h: 0.85, fontSize: 15, fontFace: F, color: C.white, valign: 'middle' });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: 5.1, w: 5.0, h: 0.3, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('▶　シミュレーター　デモ', { x: 2.5, y: 5.1, w: 5.0, h: 0.3, fontSize: 14, fontFace: F, bold: true, color: C.navy, align: 'center', valign: 'middle', margin: 0 });

  bottomBar(s, C.green);
}

// ── SLIDE 22: ケースC 状況 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースC　子どもが社会人');
  caseTag(s, 'CASE C', C.greenMid);

  card(s, 0.5, 1.55, 9.0, 1.65, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.55, w: 0.15, h: 1.65, fill: { color: C.greenMid }, line: { color: C.greenMid } });
  s.addText('状況：学費の心配がなくなった。資産形成のフルパワー期\n今まで学費に使っていたお金が丸ごと老後資金に回せる。', {
    x: 0.8, y: 1.6, w: 8.5, h: 1.55,
    fontSize: 18, fontFace: F, color: C.bodyText, valign: 'middle', lineSpacingMultiple: 1.5
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.45, w: 9.0, h: 0.55, fill: { color: C.greenMid }, line: { color: C.greenMid } });
  s.addText('このタイミングを逃さないことが最重要', {
    x: 0.5, y: 3.45, w: 9.0, h: 0.55,
    fontSize: 17, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0
  });
}

// ── SLIDE 23: ケースC 結論 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースC　結論');
  caseTag(s, 'CASE C', C.greenMid);

  twoColCards(s,
    'お金に余裕がある方',
    '「使う」ことも視野に入れる\n\n・老後の旅行・リフォームの計画を\n・使わないお金は長期保険で\n  今の利率を固定する\n・NISAも並行してフル活用',
    'お金が足りない方',
    'とにかく急いで\nNISAを始める\n\n・フルパワーで積み立てる\n・長く働く＝最強の老後対策\n・65歳→70歳の差は数百万円'
  );
}

// ── SLIDE 24: 3ケース早見表 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'まとめ：あなたはどのケース？');

  const tableD = [
    [{ text: '', options: { fill: { color: 'EDF2F7' } } },
     { text: 'お金に余裕がある', options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 14 } },
     { text: 'お金が足りない', options: { fill: { color: '4a6fa5' }, color: C.white, bold: true, fontSize: 14 } }],
    [{ text: 'A 大学生', options: { fill: { color: C.blueCard }, bold: true, fontSize: 13, color: C.navy } },
     { text: '保険で増やし\n使い始める', options: { fontSize: 13, color: C.navy } },
     { text: '覚悟を持って\nNISA', options: { fontSize: 13, color: C.bodyText } }],
    [{ text: 'B 中学生', options: { fill: { color: C.greenCard }, bold: true, fontSize: 13, color: C.green } },
     { text: '老後まで試算\n長期保険＋NISA', options: { fontSize: 13, color: C.navy } },
     { text: '奨学金活用\n流動性確保', options: { fontSize: 13, color: C.bodyText } }],
    [{ text: 'C 社会人', options: { fill: { color: 'F0FFF4' }, bold: true, fontSize: 13, color: C.greenMid } },
     { text: '使う計画＋\n長期保険で固定', options: { fontSize: 13, color: C.navy } },
     { text: 'NISA全力\n長く働く', options: { fontSize: 13, color: C.bodyText } }],
  ];
  s.addTable(tableD, {
    x: 0.5, y: 1.1, w: 9.0, h: 4.1,
    colW: [2.2, 3.4, 3.4],
    border: { pt: 1, color: 'CBD5E0' },
    align: 'center', valign: 'middle', rowH: 0.95,
  });
}

// ── SLIDE 25: 万一のときは ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'おまけ：保険はもうひとつ「得」がある');

  card(s, 0.5, 1.1, 9.0, 2.1, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 0.15, h: 2.1, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('死亡保険金の非課税枠', { x: 0.8, y: 1.15, w: 8.5, h: 0.5, fontSize: 16, fontFace: F, bold: true, color: C.navy, valign: 'middle' });
  s.addText('500万円 × 法定相続人の数  まで非課税', { x: 0.8, y: 1.7, w: 8.5, h: 0.7, fontSize: 24, fontFace: F, bold: true, color: C.navy, valign: 'middle' });
  s.addText('例：相続人3人 → 1,500万円まで相続税ゼロ', { x: 0.8, y: 2.5, w: 8.5, h: 0.55, fontSize: 14, fontFace: F, color: C.gray, valign: 'middle' });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 9.0, h: 1.3, fill: { color: C.navy }, line: { color: C.navy } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 0.15, h: 1.3, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('「増やす」＋「万一のとき非課税で渡せる」\n一時払い保険は二重においしい', {
    x: 0.8, y: 3.55, w: 8.5, h: 1.2,
    fontSize: 18, fontFace: F, bold: true, color: C.white, valign: 'middle', lineSpacingMultiple: 1.5
  });
}

// ── SLIDE 26: クロージング ──────────────────────────────────────
{
  const s = pres.addSlide();
  darkBg(s);
  topBar(s, C.gold);

  s.addText('NISAはひとりでできる。\n保険は違う。', {
    x: 0.5, y: 0.2, w: 9, h: 1.5,
    fontSize: 30, fontFace: F, bold: true, color: C.white,
    align: 'center', lineSpacingMultiple: 1.4
  });

  card(s, 0.4, 1.95, 4.2, 2.3, C.navyMid);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.95, w: 4.2, h: 0.5, fill: { color: C.navyLight }, line: { color: C.navyLight } });
  s.addText('NISA', { x: 0.4, y: 1.95, w: 4.2, h: 0.5, fontSize: 15, fontFace: F, bold: true, color: C.goldLight, align: 'center', valign: 'middle', margin: 0 });
  s.addText('口座開設して\n積立設定するだけ\nひとりでOK', { x: 0.4, y: 2.5, w: 4.2, h: 1.65, fontSize: 16, fontFace: F, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.4 });

  card(s, 5.4, 1.95, 4.2, 2.3, C.navyMid);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.95, w: 4.2, h: 0.5, fill: { color: C.green }, line: { color: C.green } });
  s.addText('一時払い保険', { x: 5.4, y: 1.95, w: 4.2, h: 0.5, fontSize: 15, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('商品・期間・金額は\nあなたの状況次第\n相談相手が必要', { x: 5.4, y: 2.5, w: 4.2, h: 1.65, fontSize: 16, fontFace: F, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.4 });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 4.5, w: 8.0, h: 0.78, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('「その相談相手、いますか？」', {
    x: 1.0, y: 4.5, w: 8.0, h: 0.78,
    fontSize: 26, fontFace: F, bold: true, color: C.navy,
    align: 'center', valign: 'middle', margin: 0
  });
}

// ── SLIDE 27: 個別相談のご案内 ──────────────────────────────────────
{
  const s = pres.addSlide();
  darkBg(s);
  topBar(s, C.gold);

  s.addText('本日の特典', {
    x: 0.5, y: 0.2, w: 9, h: 0.75,
    fontSize: 28, fontFace: F, bold: true, color: C.white, align: 'center'
  });

  // 特典①
  card(s, 0.4, 1.15, 4.2, 3.5, C.navyMid);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 4.2, h: 0.55, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('特典①　シミュレーター', { x: 0.4, y: 1.15, w: 4.2, h: 0.55, fontSize: 14, fontFace: F, bold: true, color: C.navy, align: 'center', valign: 'middle', margin: 0 });
  s.addText('自分の数字で\n「保険 vs NISA」を\n試せるWebアプリ\n\nhoken-nisa-hikaku.netlify.app', {
    x: 0.4, y: 1.75, w: 4.2, h: 2.8,
    fontSize: 15, fontFace: F, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5
  });

  // 特典②
  card(s, 5.4, 1.15, 4.2, 3.5, C.navyMid);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.15, w: 4.2, h: 0.55, fill: { color: C.green }, line: { color: C.green } });
  s.addText('特典②　個別相談', { x: 5.4, y: 1.15, w: 4.2, h: 0.55, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('あなたの状況で\n「どうすればいい？」を\n一緒に考えます\n\n無料・30分\n申込フォームをチャットに', {
    x: 5.4, y: 1.75, w: 4.2, h: 2.8,
    fontSize: 15, fontFace: F, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 4.9, w: 8.0, h: 0.5, fill: { color: C.navyLight }, line: { color: C.navyLight } });
  s.addText('本日はご参加いただき、ありがとうございました', {
    x: 1.0, y: 4.9, w: 8.0, h: 0.5,
    fontSize: 14, fontFace: F, color: C.goldLight, align: 'center', valign: 'middle', margin: 0
  });

  bottomBar(s, C.green);
}

// ── 書き出し ──────────────────────────────────────
pres.writeFile({ fileName: 'C:/Users/love2/Desktop/Claude/セミナー_保険vsNISA/seminar_slides.pptx' })
  .then(() => console.log('完了：seminar_slides.pptx'))
  .catch(e => console.error('エラー：', e));
