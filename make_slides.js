const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = 'LAYOUT_4x3';  // 10 x 7.5 inches（参照ファイルに合わせる）
pres.title = '一時払い保険とNISAの賢い使い分け';

// ── カラーパレット（参照ファイルベース）──
const C = {
  darkNavy:    '002060',   // ダーク背景・最強調
  blue:        '1F497D',   // タイトルバー・ヘッダー
  blueMid:     '2E5E9E',   // サブヘッダー・中間強調
  blueLight:   '4472C4',   // 薄い青要素
  red:         'C00000',   // 強調ボックス（参照ファイルの主要アクセント）
  redMid:      'CC3333',   // 中間赤
  orange:      'FF6600',   // サブアクセント（参照ファイルの橙）
  orangeLight: 'FFE4CC',   // 薄い橙背景
  white:       'FFFFFF',
  offWhite:    'F5F5F5',
  blueCard:    'DCE6F1',   // 薄い青背景カード
  greenCard:   'E2EFDA',   // 薄い緑背景カード（右カード用）
  gray:        '666666',
  grayLight:   'CCCCCC',
  bodyText:    '2D2D2D',
  goldCard:    'FFF2CC',
};
const F  = 'Meiryo';
const FA = 'Arial';
const mkSh = () => ({ type: 'outer', blur: 6, offset: 2, angle: 135, color: '000000', opacity: 0.10 });

// ── ヘルパー ──
function darkBg(s) { s.background = { color: C.darkNavy }; }
function lightBg(s) { s.background = { color: C.white }; }

// 全幅タイトルバー（参照ファイルのパターン：上端全幅）
function sectionTitle(s, title) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.52,
    fill: { color: C.blue }, line: { color: C.blue }
  });
  s.addText(title, {
    x: 0.3, y: 0, w: 9.4, h: 0.52,
    fontSize: 21, fontFace: F, bold: true, color: C.white, margin: 0, valign: 'middle'
  });
}

// ダークスライド用：細いアクセントライン（上端）
function topBar(s, color) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.09,
    fill: { color: color || C.orange }, line: { color: color || C.orange }
  });
}
// ダークスライド用：底辺アクセントライン
function bottomBar(s, color) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 7.37, w: 10, h: 0.13,
    fill: { color: color || C.red }, line: { color: color || C.red }
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

// 赤い強調ボックス（参照ファイルの #C00000 パターン）
function redBox(s, x, y, w, h, text, fontSize) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.red }, line: { color: C.red } });
  s.addText(text, {
    x, y, w, h, fontSize: fontSize || 22, fontFace: F, bold: true,
    color: C.white, align: 'center', valign: 'middle', margin: 0
  });
}

function caseTag(s, label, color) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 1.36, w: 1.6, h: 0.51,
    fill: { color: color || C.blue }, line: { color: color || C.blue }
  });
  s.addText(label, {
    x: 0.45, y: 1.36, w: 1.6, h: 0.51,
    fontSize: 13, fontFace: F, bold: true, color: C.white,
    align: 'center', valign: 'middle', margin: 0
  });
}

function twoColCards(s, leftTitle, leftItems, rightTitle, rightItems) {
  card(s, 0.4, 1.3, 4.4, 5.2, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.3, w: 4.4, h: 0.60, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText(leftTitle, { x: 0.4, y: 1.3, w: 4.4, h: 0.60, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText(leftItems, { x: 0.55, y: 2.0, w: 4.1, h: 4.3, fontSize: 14, fontFace: F, color: C.bodyText, valign: 'top', lineSpacingMultiple: 1.5 });

  card(s, 5.2, 1.3, 4.4, 5.2, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.4, h: 0.60, fill: { color: C.red }, line: { color: C.red } });
  s.addText(rightTitle, { x: 5.2, y: 1.3, w: 4.4, h: 0.60, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText(rightItems, { x: 5.35, y: 2.0, w: 4.1, h: 4.3, fontSize: 14, fontFace: F, color: C.bodyText, valign: 'top', lineSpacingMultiple: 1.5 });
}

// ── SLIDE 1: タイトル ──────────────────────────────────────────
{
  const s = pres.addSlide();
  darkBg(s);
  topBar(s, C.orange);

  s.addText('一時払い保険とNISAの賢い使い分け', {
    x: 0.6, y: 1.33, w: 8.8, h: 2.13,
    fontSize: 40, fontFace: F, bold: true, color: C.white,
    align: 'center', valign: 'middle'
  });
  s.addText('同じ金額・同じ10年。何が違う？', {
    x: 0.6, y: 3.67, w: 8.8, h: 0.87,
    fontSize: 22, fontFace: F, color: C.orangeLight,
    align: 'center'
  });
  s.addShape(pres.shapes.LINE, { x: 2.8, y: 4.73, w: 4.4, h: 0, line: { color: C.blueLight, width: 1 } });

  s.addShape(pres.shapes.RECTANGLE, { x: 2.4, y: 4.93, w: 2.4, h: 0.60, fill: { color: C.blueMid }, line: { color: C.blueLight, width: 1 } });
  s.addText('講師名', { x: 2.4, y: 4.93, w: 2.4, h: 0.60, fontSize: 13, fontFace: F, color: 'AABBCC', align: 'center', valign: 'middle', margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 4.93, w: 2.4, h: 0.60, fill: { color: C.blueMid }, line: { color: C.blueLight, width: 1 } });
  s.addText('日付', { x: 5.2, y: 4.93, w: 2.4, h: 0.60, fontSize: 13, fontFace: F, color: 'AABBCC', align: 'center', valign: 'middle', margin: 0 });

  bottomBar(s, C.red);
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
    const y = 1.67 + i * 1.60;
    card(s, 0.5, y, 9.0, 1.27, C.white);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.87, h: 1.27, fill: { color: C.blue }, line: { color: C.blue } });
    s.addText(item.num, { x: 0.5, y, w: 0.87, h: 1.27, fontSize: 18, fontFace: F, bold: true, color: C.orangeLight, align: 'center', valign: 'middle', margin: 0 });
    s.addText(item.text, { x: 1.55, y: y + 0.07, w: 7.8, h: 1.13, fontSize: 19, fontFace: F, bold: true, color: C.bodyText, valign: 'middle' });
  });
}

// ── SLIDE 3: あなたの本音 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'こんな気持ち、ありませんか？');

  card(s, 0.5, 1.53, 9.0, 1.80, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.53, w: 0.20, h: 1.80, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('「子どもの学費が最優先…でも老後もそろそろ心配。\nどうすればいいの？」', {
    x: 0.87, y: 1.60, w: 8.4, h: 1.67,
    fontSize: 20, fontFace: F, bold: true, color: C.blue,
    valign: 'middle', lineSpacingMultiple: 1.5
  });

  card(s, 0.5, 3.60, 9.0, 1.33, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.60, w: 0.20, h: 1.33, fill: { color: C.blueMid }, line: { color: C.blueMid } });
  s.addText('「NISAって今から間に合うの？　損したら怖い…」', {
    x: 0.87, y: 3.67, w: 8.4, h: 1.20,
    fontSize: 20, fontFace: F, bold: true, color: C.blueMid,
    valign: 'middle'
  });

  redBox(s, 0.5, 5.20, 9.0, 0.73, '→ 今日のセミナーで「やることが見えた」状態にします', 16);
}

// ── SLIDE 4: 時間はある ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '50代でも、まだ30〜40年ある');

  const tl_y = 3.07, tl_x = 0.7, tl_w = 8.6;
  s.addShape(pres.shapes.RECTANGLE, { x: tl_x, y: tl_y + 0.24, w: tl_w, h: 0.08, fill: { color: C.grayLight }, line: { color: C.grayLight } });

  const points = [
    { x: tl_x,             label: '現在\n40〜50代', color: C.blue },
    { x: tl_x + tl_w*0.38, label: '65歳\n老後スタート', color: C.blueMid },
    { x: tl_x + tl_w*0.75, label: '80代', color: C.gray },
    { x: tl_x + tl_w,      label: '90代', color: C.grayLight },
  ];
  points.forEach(p => {
    s.addShape(pres.shapes.OVAL, { x: p.x - 0.19, y: tl_y + 0.05, w: 0.50, h: 0.50, fill: { color: p.color }, line: { color: p.color } });
    s.addText(p.label, { x: p.x - 0.6, y: tl_y + 0.70, w: 1.2, h: 0.93, fontSize: 12, fontFace: F, color: C.bodyText, align: 'center' });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: tl_x, y: tl_y + 0.28, w: tl_w * 0.75, h: 0, fill: { color: C.blue }, line: { color: C.blue, width: 3 } });

  card(s, 1.0, 4.80, 8.0, 1.13, C.goldCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 4.80, w: 0.20, h: 1.13, fill: { color: C.orange }, line: { color: C.orange } });
  s.addText('「時間がない」のではなく「時間の使い方を変える」タイミング', {
    x: 1.35, y: 4.87, w: 7.5, h: 1.00,
    fontSize: 16, fontFace: F, bold: true, color: C.bodyText, valign: 'middle'
  });
}

// ── SLIDE 5: ドル資産① ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '円高か円安か、誰にもわからない');

  card(s, 0.4, 1.47, 3.6, 4.67, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.47, w: 3.6, h: 0.67, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('円高になったら？', { x: 0.4, y: 1.47, w: 3.6, h: 0.67, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('円資産が有利\n日本円で持っていると\n購買力が上がる', { x: 0.55, y: 2.27, w: 3.3, h: 3.60, fontSize: 16, fontFace: F, color: C.bodyText, valign: 'middle', lineSpacingMultiple: 1.6 });

  card(s, 5.95, 1.47, 3.6, 4.67, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.95, y: 1.47, w: 3.6, h: 0.67, fill: { color: C.blueMid }, line: { color: C.blueMid } });
  s.addText('円安になったら？', { x: 5.95, y: 1.47, w: 3.6, h: 0.67, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('外貨資産が有利\n円建て資産の価値が\n目減りする', { x: 6.1, y: 2.27, w: 3.3, h: 3.60, fontSize: 16, fontFace: F, color: C.bodyText, valign: 'middle', lineSpacingMultiple: 1.6 });

  s.addShape(pres.shapes.RECTANGLE, { x: 4.1, y: 2.47, w: 1.75, h: 2.00, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('だから\n「両方持つ」\n＝分散', { x: 4.1, y: 2.47, w: 1.75, h: 2.00, fontSize: 14, fontFace: F, bold: true, color: C.orangeLight, align: 'center', valign: 'middle', margin: 0 });
}

// ── SLIDE 6: ドル資産② ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '日本は輸入大国。円安は生活を直撃する');

  const flow = [
    { label: '食料・エネルギー\n・部品', bg: C.blueCard, tc: C.blue },
    { label: '外貨（ドル）で購入', bg: C.blue, tc: C.white },
    { label: '円安になると\nコスト上昇', bg: C.red, tc: C.white },
    { label: '物価上昇・\n生活費増加', bg: '7B1010', tc: C.white },
  ];
  flow.forEach((f, i) => {
    const x = 0.4 + i * 2.38;
    card(s, x, 1.53, 2.1, 2.00, f.bg);
    s.addText(f.label, { x, y: 1.53, w: 2.1, h: 2.00, fontSize: 14, fontFace: F, bold: true, color: f.tc, align: 'center', valign: 'middle', margin: 8 });
    if (i < 3) {
      s.addShape(pres.shapes.RECTANGLE, { x: x + 2.1, y: 2.20, w: 0.28, h: 0.52, fill: { color: C.orange }, line: { color: C.orange } });
      s.addText('▶', { x: x + 2.1, y: 2.20, w: 0.28, h: 0.52, fontSize: 14, fontFace: F, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    }
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.13, w: 9.15, h: 1.87, fill: { color: C.blue }, line: { color: C.blue } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.13, w: 0.14, h: 1.87, fill: { color: C.orange }, line: { color: C.orange } });
  s.addText('だからこそ、資産の一部を外貨で持つことが「生活の守り」になる\n\n今日紹介する一時払い保険もNISA（オルカン）も、どちらもドル建て・外貨資産として機能する', {
    x: 0.67, y: 4.20, w: 8.8, h: 1.73,
    fontSize: 15, fontFace: F, color: C.white, valign: 'top', lineSpacingMultiple: 1.5
  });
}

// ── SLIDE 7: NISAを整理する ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'NISAとは？　1枚で整理');

  card(s, 0.5, 1.47, 9.0, 1.60, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.47, w: 1.5, h: 1.60, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('NISA\n口座', { x: 0.5, y: 1.47, w: 1.5, h: 1.60, fontSize: 16, fontFace: F, bold: true, color: C.orangeLight, align: 'center', valign: 'middle', margin: 0 });
  s.addText('運用益に税金ゼロ　（通常は利益の約20%が税金）', {
    x: 2.2, y: 1.60, w: 7.1, h: 1.33,
    fontSize: 19, fontFace: F, bold: true, color: C.blue, valign: 'middle'
  });

  const feats = [
    { title: '毎月積立', body: '少額から\nコツコツ積立' },
    { title: '長期運用', body: '最長20年\n非課税で保有' },
    { title: '世界分散\n（オルカン）', body: '全世界株式に\n自動分散' },
  ];
  feats.forEach((f, i) => {
    const x = 0.5 + i * 3.05;
    card(s, x, 3.47, 2.8, 3.47, C.white);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.47, w: 2.8, h: 0.73, fill: { color: C.blue }, line: { color: C.blue } });
    s.addText(f.title, { x, y: 3.47, w: 2.8, h: 0.73, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    s.addText(f.body, { x, y: 4.27, w: 2.8, h: 2.53, fontSize: 17, fontFace: F, color: C.bodyText, align: 'center', valign: 'middle' });
  });
}

// ── SLIDE 8: NISAの特徴 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'NISA（オルカン）の特徴');

  const stats = [
    { val: '約8%', sub: '想定利率（年率）', color: C.blue },
    { val: '1.5倍', sub: '10年後（目安）\n500万→750万', color: C.blueMid },
    { val: 'ブレあり', sub: '大きく上下する\n可能性がある', color: C.red },
  ];
  stats.forEach((st, i) => {
    const x = 0.45 + i * 3.05;
    card(s, x, 1.40, 2.85, 2.67, C.white);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.40, w: 2.85, h: 0.60, fill: { color: st.color }, line: { color: st.color } });
    s.addText(st.val, { x, y: 2.00, w: 2.85, h: 1.27, fontSize: 36, fontFace: F, bold: true, color: st.color, align: 'center', valign: 'middle', margin: 0 });
    s.addText(st.sub, { x, y: 3.33, w: 2.85, h: 0.73, fontSize: 13, fontFace: F, color: C.gray, align: 'center', valign: 'top' });
  });

  const chartData = [
    { name: '楽観', labels: ['0', '2', '4', '6', '8', '10'], values: [500, 610, 740, 900, 1090, 1320] },
    { name: '中央値', labels: ['0', '2', '4', '6', '8', '10'], values: [500, 540, 583, 630, 680, 735] },
    { name: '悲観', labels: ['0', '2', '4', '6', '8', '10'], values: [500, 480, 460, 440, 420, 400] },
  ];
  s.addChart(pres.charts.LINE, chartData, {
    x: 0.45, y: 4.27, w: 9.1, h: 2.93,
    chartColors: [C.blueLight, C.blue, C.red],
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    lineSize: 2, lineSmooth: true,
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray,
    valGridLine: { color: 'E0E0E0', size: 0.5 }, catGridLine: { style: 'none' },
    showLegend: true, legendPos: 'r', legendFontSize: 10,
    showTitle: false, valAxisMinVal: 300, valAxisMaxVal: 1400,
  });
}

// ── SLIDE 9: 一時払い保険を整理する ──────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '一時払い保険とは？　1枚で整理');

  const flow2 = [
    { label: 'まとまった\nお金を\n一括払い', bg: C.blue, tc: C.white },
    { label: 'ドルで\n複利運用', bg: C.blueMid, tc: C.white },
    { label: '解約返戻金\nとして受取\n（確実に増える）', bg: C.blueCard, tc: C.blue },
  ];
  flow2.forEach((f, i) => {
    const x = 0.6 + i * 3.15;
    card(s, x, 1.60, 2.8, 4.27, f.bg);
    s.addText(f.label, { x, y: 1.60, w: 2.8, h: 4.27, fontSize: 18, fontFace: F, bold: true, color: f.tc, align: 'center', valign: 'middle', margin: 12 });
    if (i < 2) {
      s.addShape(pres.shapes.RECTANGLE, { x: x + 2.8, y: 3.33, w: 0.35, h: 0.73, fill: { color: C.orange }, line: { color: C.orange } });
      s.addText('▶', { x: x + 2.8, y: 3.33, w: 0.35, h: 0.73, fontSize: 18, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    }
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 6.20, w: 8.85, h: 0.73, fill: { color: C.goldCard }, line: { color: C.orange, width: 1 } });
  s.addText('「保険」という名前がついているが、今日は「増やす手段」として考える', {
    x: 0.6, y: 6.20, w: 8.85, h: 0.73,
    fontSize: 14, fontFace: F, bold: true, color: C.blue, align: 'center', valign: 'middle', margin: 0
  });
}

// ── SLIDE 10: 一時払い保険の特徴 ──────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '一時払い保険（ドル建て）の特徴');

  const stats2 = [
    { val: '約4%', sub: '利率（年率）', color: C.blue },
    { val: '1.5倍', sub: '10年後（確実）\n500万→750万', color: C.blue },
    { val: 'ブレなし', sub: '確実に増える\n安心の右肩上がり', color: C.blueMid },
  ];
  stats2.forEach((st, i) => {
    const x = 0.45 + i * 3.05;
    card(s, x, 1.40, 2.85, 2.67, C.white);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.40, w: 2.85, h: 0.60, fill: { color: st.color }, line: { color: st.color } });
    s.addText(st.val, { x, y: 2.00, w: 2.85, h: 1.27, fontSize: 36, fontFace: F, bold: true, color: st.color, align: 'center', valign: 'middle', margin: 0 });
    s.addText(st.sub, { x, y: 3.33, w: 2.85, h: 0.73, fontSize: 13, fontFace: F, color: C.gray, align: 'center', valign: 'top' });
  });

  const insData = [{ name: '一時払い保険', labels: ['0', '2', '4', '6', '8', '10'], values: [500, 541, 584, 631, 682, 735] }];
  s.addChart(pres.charts.LINE, insData, {
    x: 0.45, y: 4.27, w: 9.1, h: 2.93,
    chartColors: [C.blue],
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    lineSize: 3, lineSmooth: true,
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray,
    valGridLine: { color: 'E0E0E0', size: 0.5 }, catGridLine: { style: 'none' },
    showLegend: false, showTitle: false,
    valAxisMinVal: 450, valAxisMaxVal: 800,
  });
}

// ── SLIDE 11: 並べてみる ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'どちらも10年で1.5倍。違いは「ブレ」だけ');

  const headers = [
    [{ text: '', options: { fill: { color: C.offWhite } } },
     { text: '一時払い保険', options: { fill: { color: C.blue }, color: C.white, bold: true, fontSize: 15 } },
     { text: 'NISA（つみたて）', options: { fill: { color: C.blueMid }, color: C.white, bold: true, fontSize: 15 } }],
    [{ text: '利回り', options: { fill: { color: C.blueCard }, bold: true, fontSize: 14, color: C.blue } },
     { text: '約4%', options: { fontSize: 15, color: C.blue, bold: true } },
     { text: '約8%', options: { fontSize: 15, color: C.blueMid, bold: true } }],
    [{ text: '10年後', options: { fill: { color: C.blueCard }, bold: true, fontSize: 14, color: C.blue } },
     { text: '約1.5倍（確実）', options: { fontSize: 14, color: C.blue, bold: true } },
     { text: '約1.5倍（目安）', options: { fontSize: 14, color: C.blueMid } }],
    [{ text: 'ブレ', options: { fill: { color: C.blueCard }, bold: true, fontSize: 14, color: C.blue } },
     { text: 'なし', options: { fontSize: 15, color: C.blue, bold: true } },
     { text: 'あり（大きく）', options: { fontSize: 15, color: C.red, bold: true } }],
    [{ text: '流動性', options: { fill: { color: C.blueCard }, bold: true, fontSize: 14, color: C.blue } },
     { text: '低い（途中解約は損）', options: { fontSize: 13, color: C.gray } },
     { text: '高い（いつでも売れる）', options: { fontSize: 13, color: C.gray } }],
  ];
  s.addTable(headers, {
    x: 0.5, y: 1.40, w: 9.0, h: 4.53,
    colW: [2.2, 3.4, 3.4],
    border: { pt: 1, color: 'CCCCCC' },
    align: 'center', valign: 'middle', rowH: 0.87,
  });

  redBox(s, 0.5, 6.20, 9.0, 0.87, '「そのお金、減ったら困りますか？」', 20);
}

// ── SLIDE 12: デモ① ──────────────────────────────────────
{
  const s = pres.addSlide();
  darkBg(s);
  topBar(s, C.orange);

  s.addText('実際に動かしてみましょう', {
    x: 0.5, y: 1.07, w: 9, h: 1.20,
    fontSize: 34, fontFace: F, bold: true, color: C.white, align: 'center'
  });

  card(s, 1.0, 2.60, 8.0, 2.13, C.blueMid);
  s.addText('500万円・10年\n保険 4%  vs  NISA 8%  で比べると…？', {
    x: 1.0, y: 2.60, w: 8.0, h: 2.13,
    fontSize: 22, fontFace: F, bold: true, color: C.orangeLight,
    align: 'center', valign: 'middle', lineSpacingMultiple: 1.5
  });

  redBox(s, 2.5, 5.07, 5.0, 0.93, '▶　シミュレーター　デモ', 20);

  bottomBar(s, C.red);
}

// ── SLIDE 13: 使い分けの鍵 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '使い分けの判断は、この一問');

  card(s, 0.5, 1.47, 9.0, 1.60, C.blue);
  s.addText('「そのお金、10年後に減っていたら困りますか？」', {
    x: 0.5, y: 1.47, w: 9.0, h: 1.60,
    fontSize: 22, fontFace: F, bold: true, color: C.orangeLight,
    align: 'center', valign: 'middle', margin: 0
  });

  card(s, 0.5, 3.47, 4.3, 3.60, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.47, w: 4.3, h: 0.67, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('YES（困る）', { x: 0.5, y: 3.47, w: 4.3, h: 0.67, fontSize: 15, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('一時払い保険\nブレない 1.5倍', { x: 0.5, y: 4.20, w: 4.3, h: 2.67, fontSize: 20, fontFace: F, bold: true, color: C.blue, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5 });

  card(s, 5.2, 3.47, 4.3, 3.60, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.47, w: 4.3, h: 0.67, fill: { color: C.red }, line: { color: C.red } });
  s.addText('NO（大丈夫）', { x: 5.2, y: 3.47, w: 4.3, h: 0.67, fontSize: 15, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('NISA\nブレても待てる', { x: 5.2, y: 4.20, w: 4.3, h: 2.67, fontSize: 20, fontFace: F, bold: true, color: C.red, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5 });
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
    const y = 1.47 + i * 1.80;
    card(s, 0.5, y, 9.0, 1.60, C.white);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.93, h: 1.60, fill: { color: C.blue }, line: { color: C.blue } });
    s.addText(p.num, { x: 0.5, y, w: 0.93, h: 1.60, fontSize: 26, fontFace: FA, bold: true, color: C.orangeLight, align: 'center', valign: 'middle', margin: 0 });
    s.addText(p.title, { x: 1.60, y: y + 0.07, w: 7.7, h: 0.53, fontSize: 16, fontFace: F, bold: true, color: C.blue, valign: 'middle' });
    s.addText(p.body, { x: 1.60, y: y + 0.64, w: 7.7, h: 0.87, fontSize: 13, fontFace: F, color: C.gray, valign: 'top' });
  });
}

// ── SLIDE 15: シミュレーション数値 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '500万円を入れたら？');

  const tableRows = [
    [{ text: '', options: { fill: { color: C.offWhite } } },
     { text: '10年後', options: { fill: { color: C.blue }, color: C.white, bold: true, fontSize: 14 } },
     { text: '20年後', options: { fill: { color: C.blue }, color: C.white, bold: true, fontSize: 14 } }],
    [{ text: '一時払い保険（4%）', options: { fill: { color: C.blueCard }, bold: true, fontSize: 13, color: C.blue } },
     { text: '約750万円\n＋250万円増', options: { fontSize: 18, bold: true, color: C.blue } },
     { text: '約1,095万円\n＋595万円増', options: { fontSize: 18, bold: true, color: C.blue } }],
    [{ text: 'NISA 中央値（8%）', options: { fill: { color: C.greenCard }, bold: true, fontSize: 13, color: C.blueMid } },
     { text: '約750万円\n（元本360万円）', options: { fontSize: 18, bold: true, color: C.blueMid } },
     { text: '約1,764万円\n（元本720万円）', options: { fontSize: 18, bold: true, color: C.blueMid } }],
  ];
  s.addTable(tableRows, {
    x: 0.5, y: 1.47, w: 9.0, h: 4.53,
    colW: [3.0, 3.0, 3.0],
    border: { pt: 1, color: 'CCCCCC' },
    align: 'center', valign: 'middle', rowH: 1.40,
  });

  s.addText('※ NISAは積立総額＝保険元本と同額の条件で試算　／　NISAの実際の成果は大きくブレます', {
    x: 0.5, y: 6.27, w: 9.0, h: 0.67,
    fontSize: 11, fontFace: F, color: C.gray, align: 'center'
  });
}

// ── SLIDE 16: ケースA 状況 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースA　子どもが大学生');
  caseTag(s, 'CASE A', C.blue);

  card(s, 0.5, 2.07, 9.0, 2.20, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.07, w: 0.20, h: 2.20, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('状況：学費の総額がほぼ見えてきた\n残り期間は1〜4年。ゴールが近い。', {
    x: 0.87, y: 2.13, w: 8.4, h: 2.07,
    fontSize: 18, fontFace: F, color: C.bodyText, valign: 'middle', lineSpacingMultiple: 1.5
  });

  redBox(s, 0.5, 4.60, 9.0, 0.73, '問：学費を払い終えた後、いくら残る？　それをどう老後に活かす？', 15);
}

// ── SLIDE 17: ケースA 結論 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースA　結論');
  caseTag(s, 'CASE A', C.blue);

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
  caseTag(s, 'CASE B', C.blueMid);

  card(s, 0.5, 2.07, 9.0, 2.20, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.07, w: 0.20, h: 2.20, fill: { color: C.blueMid }, line: { color: C.blueMid } });
  s.addText('状況：進路がまだ読めない。でも時間はある（5〜10年）\n文系か理系か、どの大学か、まだわからない。', {
    x: 0.87, y: 2.13, w: 8.4, h: 2.07,
    fontSize: 18, fontFace: F, color: C.bodyText, valign: 'middle', lineSpacingMultiple: 1.5
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.60, w: 9.0, h: 0.73, fill: { color: C.blueMid }, line: { color: C.blueMid } });
  s.addText('この期間をどう使うかが、老後を左右する', {
    x: 0.5, y: 4.60, w: 9.0, h: 0.73,
    fontSize: 17, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0
  });
}

// ── SLIDE 19: 奨学金 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, '「奨学金」は借金ではなく、武器になる');

  const nums = [
    { val: '上限3%', sub: '奨学金の上限金利\n（第二種）' },
    { val: '最長20年', sub: '返済可能期間' },
  ];
  nums.forEach((n, i) => {
    card(s, 0.5 + i * 4.6, 1.40, 4.2, 2.00, C.blueCard);
    s.addText(n.val, { x: 0.5 + i * 4.6, y: 1.40, w: 4.2, h: 1.20, fontSize: 34, fontFace: FA, bold: true, color: C.blue, align: 'center', valign: 'middle' });
    s.addText(n.sub, { x: 0.5 + i * 4.6, y: 2.60, w: 4.2, h: 0.80, fontSize: 13, fontFace: F, color: C.gray, align: 'center' });
  });

  const fl = [
    { label: '子が奨学金\nで借りる\n（〜3%）', bg: C.blue, tc: C.white },
    { label: '親はその分\n保険・NISA\nで運用（4〜8%）', bg: C.blueMid, tc: C.white },
    { label: '利率の差額\nが家族の\n利益になる', bg: C.orange, tc: C.white },
  ];
  fl.forEach((f, i) => {
    const x = 0.5 + i * 3.08;
    card(s, x, 3.67, 2.75, 2.67, f.bg);
    s.addText(f.label, { x, y: 3.67, w: 2.75, h: 2.67, fontSize: 14, fontFace: F, bold: true, color: f.tc, align: 'center', valign: 'middle', margin: 8 });
    if (i < 2) {
      s.addShape(pres.shapes.RECTANGLE, { x: x + 2.75, y: 4.67, w: 0.33, h: 0.64, fill: { color: C.grayLight }, line: { color: C.grayLight } });
      s.addText('▶', { x: x + 2.75, y: 4.67, w: 0.33, h: 0.64, fontSize: 14, fontFace: F, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    }
  });

  s.addText('子どもにNISAを覚えさせ、奨学金返済資金として積み立てる', {
    x: 0.5, y: 6.56, w: 9.0, h: 0.60,
    fontSize: 14, fontFace: F, bold: true, color: C.blueMid, align: 'center'
  });
}

// ── SLIDE 20: ケースB 結論 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースB　結論');
  caseTag(s, 'CASE B', C.blueMid);

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
  topBar(s, C.orange);

  s.addText('あなたのケースで動かしてみましょう', {
    x: 0.5, y: 1.07, w: 9, h: 1.20,
    fontSize: 30, fontFace: F, bold: true, color: C.white, align: 'center'
  });

  const cases = [
    { label: 'CASE A', text: '1,000万円・20年で保険に入れたら？' },
    { label: 'CASE B', text: '500万円・10年の保険 vs NISA比較' },
    { label: 'CASE C', text: '月5万円・20年のNISA積立は？' },
  ];
  cases.forEach((c, i) => {
    card(s, 0.5, 2.60 + i * 1.40, 8.95, 1.13, C.blueMid);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.60 + i * 1.40, w: 1.1, h: 1.13, fill: { color: C.orange }, line: { color: C.orange } });
    s.addText(c.label, { x: 0.5, y: 2.60 + i * 1.40, w: 1.1, h: 1.13, fontSize: 13, fontFace: FA, bold: true, color: C.darkNavy, align: 'center', valign: 'middle', margin: 0 });
    s.addText(c.text, { x: 1.75, y: 2.60 + i * 1.40, w: 7.5, h: 1.13, fontSize: 16, fontFace: F, color: C.white, valign: 'middle' });
  });

  redBox(s, 2.5, 6.80, 5.0, 0.40, '▶　シミュレーター　デモ', 14);
  bottomBar(s, C.red);
}

// ── SLIDE 22: ケースC 状況 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースC　子どもが社会人');
  caseTag(s, 'CASE C', C.blueLight);

  card(s, 0.5, 2.07, 9.0, 2.20, C.greenCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.07, w: 0.20, h: 2.20, fill: { color: C.blueLight }, line: { color: C.blueLight } });
  s.addText('状況：学費の心配がなくなった。資産形成のフルパワー期\n今まで学費に使っていたお金が丸ごと老後資金に回せる。', {
    x: 0.87, y: 2.13, w: 8.4, h: 2.07,
    fontSize: 18, fontFace: F, color: C.bodyText, valign: 'middle', lineSpacingMultiple: 1.5
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.60, w: 9.0, h: 0.73, fill: { color: C.blueLight }, line: { color: C.blueLight } });
  s.addText('このタイミングを逃さないことが最重要', {
    x: 0.5, y: 4.60, w: 9.0, h: 0.73,
    fontSize: 17, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0
  });
}

// ── SLIDE 23: ケースC 結論 ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'ケースC　結論');
  caseTag(s, 'CASE C', C.blueLight);

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
    [{ text: '', options: { fill: { color: C.offWhite } } },
     { text: 'お金に余裕がある', options: { fill: { color: C.blue }, color: C.white, bold: true, fontSize: 14 } },
     { text: 'お金が足りない', options: { fill: { color: C.blueMid }, color: C.white, bold: true, fontSize: 14 } }],
    [{ text: 'A 大学生', options: { fill: { color: C.blueCard }, bold: true, fontSize: 13, color: C.blue } },
     { text: '保険で増やし\n使い始める', options: { fontSize: 13, color: C.blue } },
     { text: '覚悟を持って\nNISA', options: { fontSize: 13, color: C.bodyText } }],
    [{ text: 'B 中学生', options: { fill: { color: C.greenCard }, bold: true, fontSize: 13, color: C.blueMid } },
     { text: '老後まで試算\n長期保険＋NISA', options: { fontSize: 13, color: C.blue } },
     { text: '奨学金活用\n流動性確保', options: { fontSize: 13, color: C.bodyText } }],
    [{ text: 'C 社会人', options: { fill: { color: 'E8F5E9' }, bold: true, fontSize: 13, color: C.blueLight } },
     { text: '使う計画＋\n長期保険で固定', options: { fontSize: 13, color: C.blue } },
     { text: 'NISA全力\n長く働く', options: { fontSize: 13, color: C.bodyText } }],
  ];
  s.addTable(tableD, {
    x: 0.5, y: 1.47, w: 9.0, h: 5.47,
    colW: [2.2, 3.4, 3.4],
    border: { pt: 1, color: 'CCCCCC' },
    align: 'center', valign: 'middle', rowH: 1.27,
  });
}

// ── SLIDE 25: 万一のときは ──────────────────────────────────────
{
  const s = pres.addSlide();
  lightBg(s);
  sectionTitle(s, 'おまけ：保険はもうひとつ「得」がある');

  card(s, 0.5, 1.47, 9.0, 2.80, C.blueCard);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.47, w: 0.20, h: 2.80, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('死亡保険金の非課税枠', { x: 0.87, y: 1.53, w: 8.4, h: 0.67, fontSize: 16, fontFace: F, bold: true, color: C.blue, valign: 'middle' });
  s.addText('500万円 × 法定相続人の数  まで非課税', { x: 0.87, y: 2.27, w: 8.4, h: 0.93, fontSize: 24, fontFace: F, bold: true, color: C.blue, valign: 'middle' });
  s.addText('例：相続人3人 → 1,500万円まで相続税ゼロ', { x: 0.87, y: 3.33, w: 8.4, h: 0.73, fontSize: 14, fontFace: F, color: C.gray, valign: 'middle' });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.67, w: 9.0, h: 1.73, fill: { color: C.blue }, line: { color: C.blue } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.67, w: 0.20, h: 1.73, fill: { color: C.orange }, line: { color: C.orange } });
  s.addText('「増やす」＋「万一のとき非課税で渡せる」\n一時払い保険は二重においしい', {
    x: 0.87, y: 4.73, w: 8.4, h: 1.60,
    fontSize: 19, fontFace: F, bold: true, color: C.white, valign: 'middle', lineSpacingMultiple: 1.5
  });
}

// ── SLIDE 26: クロージング ──────────────────────────────────────
{
  const s = pres.addSlide();
  darkBg(s);
  topBar(s, C.orange);

  s.addText('NISAはひとりでできる。\n保険は違う。', {
    x: 0.5, y: 0.27, w: 9, h: 2.00,
    fontSize: 30, fontFace: F, bold: true, color: C.white,
    align: 'center', lineSpacingMultiple: 1.4
  });

  card(s, 0.4, 2.60, 4.2, 3.07, C.blueMid);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.60, w: 4.2, h: 0.67, fill: { color: C.blueLight }, line: { color: C.blueLight } });
  s.addText('NISA', { x: 0.4, y: 2.60, w: 4.2, h: 0.67, fontSize: 15, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('口座開設して\n積立設定するだけ\nひとりでOK', { x: 0.4, y: 3.33, w: 4.2, h: 2.20, fontSize: 16, fontFace: F, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.4 });

  card(s, 5.4, 2.60, 4.2, 3.07, C.blueMid);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 2.60, w: 4.2, h: 0.67, fill: { color: C.red }, line: { color: C.red } });
  s.addText('一時払い保険', { x: 5.4, y: 2.60, w: 4.2, h: 0.67, fontSize: 15, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('商品・期間・金額は\nあなたの状況次第\n相談相手が必要', { x: 5.4, y: 3.33, w: 4.2, h: 2.20, fontSize: 16, fontFace: F, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.4 });

  redBox(s, 1.0, 6.00, 8.0, 1.04, '「その相談相手、いますか？」', 26);
}

// ── SLIDE 27: 個別相談のご案内 ──────────────────────────────────────
{
  const s = pres.addSlide();
  darkBg(s);
  topBar(s, C.orange);

  s.addText('本日の特典', {
    x: 0.5, y: 0.27, w: 9, h: 1.00,
    fontSize: 28, fontFace: F, bold: true, color: C.white, align: 'center'
  });

  card(s, 0.4, 1.53, 4.2, 4.67, C.blueMid);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.53, w: 4.2, h: 0.73, fill: { color: C.orange }, line: { color: C.orange } });
  s.addText('特典①　シミュレーター', { x: 0.4, y: 1.53, w: 4.2, h: 0.73, fontSize: 14, fontFace: F, bold: true, color: C.darkNavy, align: 'center', valign: 'middle', margin: 0 });
  s.addText('自分の数字で\n「保険 vs NISA」を\n試せるWebアプリ\n\nhoken-nisa-hikaku.netlify.app', {
    x: 0.4, y: 2.33, w: 4.2, h: 3.73,
    fontSize: 15, fontFace: F, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5
  });

  card(s, 5.4, 1.53, 4.2, 4.67, C.blueMid);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.53, w: 4.2, h: 0.73, fill: { color: C.red }, line: { color: C.red } });
  s.addText('特典②　個別相談', { x: 5.4, y: 1.53, w: 4.2, h: 0.73, fontSize: 14, fontFace: F, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  s.addText('あなたの状況で\n「どうすればいい？」を\n一緒に考えます\n\n無料・30分\n申込フォームをチャットに', {
    x: 5.4, y: 2.33, w: 4.2, h: 3.73,
    fontSize: 15, fontFace: F, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 6.53, w: 8.0, h: 0.67, fill: { color: C.blueMid }, line: { color: C.blueMid } });
  s.addText('本日はご参加いただき、ありがとうございました', {
    x: 1.0, y: 6.53, w: 8.0, h: 0.67,
    fontSize: 14, fontFace: F, color: C.orangeLight, align: 'center', valign: 'middle', margin: 0
  });

  bottomBar(s, C.red);
}

// ── 書き出し ──────────────────────────────────────
pres.writeFile({ fileName: 'C:/Users/love2/Desktop/Claude/セミナー_保険vsNISA/seminar_slides.pptx' })
  .then(() => console.log('完了：seminar_slides.pptx'))
  .catch(e => console.error('エラー：', e));
