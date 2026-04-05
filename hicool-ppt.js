const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "特普丽新材料 - 深睡科技 HICOOL 2026";
pres.author = "深睡科技";

// ─── Design Tokens ───────────────────────────────────────────────────────────
const C = {
  navy:    "0d1b2a",   // dark slide bg
  primary: "1a3a6b",   // primary brand
  accent:  "00d4aa",   // teal accent
  orange:  "ff6b35",   // orange accent
  light:   "f0f4f8",   // light bg
  white:   "FFFFFF",
  text:    "2c3e50",
  muted:   "7f8c8d",
  silver:  "E2E8F0",
};

// ─── Reusable Helpers ────────────────────────────────────────────────────────
const makeShadow = () => ({ type: "outer", blur: 5, offset: 2, angle: 135, color: "000000", opacity: 0.10 });

// ─── Slide 1: COVER ──────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.navy };

  // left accent bar
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent } });

  // product name
  sl.addText("特普丽新材料", { x: 0.7, y: 1.2, w: 8, h: 0.6, fontSize: 18, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 4 });

  // main title
  sl.addText("深睡科技", { x: 0.7, y: 1.85, w: 8, h: 1.3, fontSize: 72, color: C.white, fontFace: "Georgia", bold: true });

  // tagline
  sl.addText("被动释放 · 持久改善 · 还原自然空气", { x: 0.7, y: 3.2, w: 8, h: 0.5, fontSize: 20, color: C.accent, fontFace: "Calibri" });

  // divider line
  sl.addShape(pres.shapes.LINE, { x: 0.7, y: 3.85, w: 4.5, h: 0, line: { color: C.accent, width: 2 } });

  // three benefit pills
  const pills = ["杀菌 99.9%", "抗病毒 99.1%", "除甲醛 88.7%"];
  pills.forEach((t, i) => {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7 + i * 2.9, y: 4.15, w: 2.7, h: 0.52, fill: { color: C.primary }, rectRadius: 0.1 });
    sl.addText(t, { x: 0.7 + i * 2.9, y: 4.15, w: 2.7, h: 0.52, fontSize: 14, color: C.white, fontFace: "Calibri", align: "center", valign: "middle" });
  });

  // bottom bar
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: C.accent } });
  sl.addText("HICOOL 2026  ·  创业大赛  ·  融资目标 ¥2000万", { x: 0.7, y: 5.2, w: 9, h: 0.425, fontSize: 13, color: C.navy, fontFace: "Calibri", bold: true, valign: "middle" });
}

// ─── Slide 2: MARKET PAIN ────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };

  // top accent strip
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

  // section label
  sl.addText("01  市场痛点", { x: 0.6, y: 0.35, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });

  // headline stat
  sl.addText("3亿人", { x: 0.6, y: 0.8, w: 4, h: 1.1, fontSize: 80, color: C.primary, fontFace: "Georgia", bold: true });
  sl.addText("中国人受睡眠问题困扰", { x: 0.6, y: 1.9, w: 4, h: 0.4, fontSize: 16, color: C.text, fontFace: "Calibri" });

  // three pain cards
  const pains = [
    { icon: "😰", title: "现有助眠方式效果有限", desc: "药物依赖 · 副作用大 · 治标不治本" },
    { icon: "🏥", title: "医院睡眠中心覆盖率低", desc: "资源稀缺 · 费用高昂 · 体验差" },
    { icon: "🌍", title: "环境问题加剧健康风险", desc: "甲醛 · 辐射 · 空气质量差" },
  ];
  pains.forEach((p, i) => {
    const x = 5.3;
    const y = 0.75 + i * 1.55;
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.1, h: 1.35, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h: 1.35, fill: { color: i === 0 ? C.accent : i === 1 ? C.orange : C.primary } });
    sl.addText(p.icon, { x: x + 0.25, y: y + 0.18, w: 0.6, h: 0.6, fontSize: 28 });
    sl.addText(p.title, { x: x + 0.85, y: y + 0.15, w: 3.1, h: 0.45, fontSize: 14, color: C.text, fontFace: "Calibri", bold: true });
    sl.addText(p.desc, { x: x + 0.85, y: y + 0.6, w: 3.1, h: 0.6, fontSize: 12, color: C.muted, fontFace: "Calibri" });
  });

  // bottom insight bar
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.1, w: 10, h: 0.525, fill: { color: C.navy } });
  sl.addText("市场需要一个安全、有效、可持续的环境改善解决方案", { x: 0.6, y: 5.1, w: 9, h: 0.525, fontSize: 14, color: C.white, fontFace: "Calibri", valign: "middle" });
}

// ─── Slide 3: SOLUTION ───────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("02  解决方案", { x: 0.6, y: 0.35, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });

  // big title left
  sl.addText("AMIC", { x: 0.6, y: 0.9, w: 3, h: 1, fontSize: 68, color: C.primary, fontFace: "Georgia", bold: true });
  sl.addText("无源负氧离子技术", { x: 0.6, y: 1.85, w: 4.5, h: 0.55, fontSize: 22, color: C.navy, fontFace: "Calibri", bold: true });
  sl.addShape(pres.shapes.LINE, { x: 0.6, y: 2.5, w: 3, h: 0, line: { color: C.accent, width: 3 } });
  sl.addText("被动释放 · 无需能耗 · 持久改善", { x: 0.6, y: 2.65, w: 4.5, h: 0.4, fontSize: 14, color: C.accent, fontFace: "Calibri" });

  // right: 3 pillars
  const pillars = [
    { icon: "⚗️", title: "材料提纯", color: C.accent, items: ["多道提纯去除有害杂质", "释放浓度安全可控", "森林级浓度无醉氧"] },
    { icon: "🏭", title: "小批量精造", color: "7c3aed", items: ["小批量定制交付", "全流程质量检测", "特普丽50年工艺"] },
    { icon: "🛠️", title: "完整服务方案", color: C.orange, items: ["环境监测部署方案", "3天完成病房改造", "6个月跟踪服务"] },
  ];

  pillars.forEach((p, i) => {
    const x = 5.3 + i * 1.55;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.85, w: 1.45, h: 4.15, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.85, w: 1.45, h: 0.1, fill: { color: p.color } });
    sl.addText(p.icon, { x, y: 1.0, w: 1.45, h: 0.55, fontSize: 26, align: "center" });
    sl.addText(p.title, { x, y: 1.55, w: 1.45, h: 0.45, fontSize: 12, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
    sl.addShape(pres.shapes.LINE, { x: x + 0.3, y: 2.05, w: 0.85, h: 0, line: { color: C.silver, width: 1 } });
    p.items.forEach((item, j) => {
      sl.addText("→  " + item, { x: x + 0.08, y: 2.15 + j * 0.7, w: 1.3, h: 0.65, fontSize: 10, color: C.muted, fontFace: "Calibri" });
    });
  });

  // bottom
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.1, w: 10, h: 0.525, fill: { color: C.primary } });
  sl.addText("安全 · 有效 · 持久 · 无需能耗", { x: 0.6, y: 5.1, w: 9, h: 0.525, fontSize: 14, color: C.white, fontFace: "Calibri", bold: true, valign: "middle", align: "center" });
}

// ─── Slide 4: SCIENCE ────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("03  循证医学验证", { x: 0.6, y: 0.35, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });

  // headline
  sl.addText("17篇国际论文  ·  40对40临床双盲验证", { x: 0.6, y: 0.85, w: 9, h: 0.6, fontSize: 26, color: C.navy, fontFace: "Georgia", bold: true });

  // four evidence cards
  const evidence = [
    { big: "99.9%", label: "抗菌率", sub: "国检权威认证" },
    { big: "99.1%", label: "抗病毒率", sub: "病毒消减实测" },
    { big: "88.7%", label: "甲醛净化率", sub: "甲醛去除率" },
    { big: "93%", label: "功能持久性", sub: "长期有效保持" },
  ];
  evidence.forEach((e, i) => {
    const x = 0.6 + i * 2.35;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.65, w: 2.15, h: 1.7, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.65, w: 2.15, h: 0.1, fill: { color: i < 2 ? C.accent : C.orange } });
    sl.addText(e.big, { x, y: 1.8, w: 2.15, h: 0.8, fontSize: 36, color: C.primary, fontFace: "Georgia", bold: true, align: "center" });
    sl.addText(e.label, { x, y: 2.55, w: 2.15, h: 0.4, fontSize: 14, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
    sl.addText(e.sub, { x, y: 2.9, w: 2.15, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri", align: "center" });
  });

  // comparison table
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.55, w: 8.8, h: 1.4, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText("安全性对比", { x: 0.8, y: 3.65, w: 2, h: 0.35, fontSize: 13, color: C.primary, fontFace: "Calibri", bold: true });
  const rows = [
    ["", "AMIC被动释放", "电晕放电式", "喷洒/雾化"],
    ["臭氧风险", "✅ 无", "⚠️ 可能产生", "❌ 可能产生"],
    ["醉氧风险", "✅ 无", "✅ 无", "⚠️ 高浓度时可能"],
    ["超细颗粒", "✅ 无", "⚠️ 可能产生", "✅ 无"],
  ];
  sl.addTable(rows, {
    x: 0.8, y: 4.0, w: 8.4, colW: [1.8, 2.2, 2.2, 2.2],
    border: { pt: 0.5, color: C.silver },
    fontFace: "Calibri",
    fontSize: 11,
    color: C.text,
    align: "center",
    valign: "middle",
    rowH: 0.3,
  });
}

// ─── Slide 5: PRODUCT ────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("04  产品矩阵", { x: 0.6, y: 0.35, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });

  // three product cards
  const products = [
    { name: "负氧离子板材", price: "¥250-300", unit: "/㎡", badge: "核心产品", color: C.accent, features: ["被动持续释放负氧离子", "抗菌抗病毒除甲醛", "墙面/天花板/地面通用", "持久性93%，长期有效"] },
    { name: "定制壁纸/窗帘", price: "¥150", unit: "/㎡", badge: "定制款", color: "7c3aed", features: ["个性化图案定制", "融入室内设计美学", "窗帘+负氧离子二合一", "已有多次小批量交付"] },
    { name: "完整解决方案", price: "交钥匙", unit: "工程", badge: "服务型", color: C.orange, features: ["环境评估+方案设计", "施工改造（3天/间）", "6个月跟踪监测", "技术支持全流程"] },
  ];

  products.forEach((p, i) => {
    const x = 0.6 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.85, w: 2.9, h: 3.95, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.85, w: 2.9, h: 0.55, fill: { color: p.color } });
    sl.addText(p.badge, { x, y: 0.85, w: 2.9, h: 0.55, fontSize: 12, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });
    sl.addText(p.name, { x: x + 0.15, y: 1.5, w: 2.6, h: 0.45, fontSize: 15, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
    sl.addText(p.price, { x: x + 0.15, y: 1.95, w: 2.6, h: 0.55, fontSize: 28, color: p.color, fontFace: "Georgia", bold: true, align: "center" });
    sl.addText(p.unit, { x: x + 0.15, y: 2.45, w: 2.6, h: 0.3, fontSize: 11, color: C.muted, fontFace: "Calibri", align: "center" });
    sl.addShape(pres.shapes.LINE, { x: x + 0.3, y: 2.85, w: 2.3, h: 0, line: { color: C.silver, width: 1 } });
    p.features.forEach((f, j) => {
      sl.addText("→  " + f, { x: x + 0.15, y: 2.95 + j * 0.45, w: 2.6, h: 0.42, fontSize: 11, color: C.text, fontFace: "Calibri" });
    });
  });

  // price advantage callout
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.9, w: 8.8, h: 0.6, fill: { color: C.navy } });
  sl.addText("💡 价格优势：比进口竞品低 50-70%，国内市场份额 25%，行业第一", { x: 0.8, y: 4.9, w: 8.4, h: 0.6, fontSize: 13, color: C.white, fontFace: "Calibri", valign: "middle" });
}

// ─── Slide 6: BUSINESS MODEL ─────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("05  商业模式", { x: 0.6, y: 0.35, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });

  // B端 card
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.85, w: 4.25, h: 4.15, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.85, w: 4.25, h: 0.5, fill: { color: C.primary } });
  sl.addText("🏢  B端：医院 & 康养机构", { x: 0.8, y: 0.85, w: 4, h: 0.5, fontSize: 13, color: C.white, fontFace: "Calibri", bold: true, valign: "middle" });

  const bItems = [
    "北京协和医院睡眠中心（已签约）",
    "301医院 · 北医三院 · 安贞医院",
    "中海锦年康养 · 房协100+机构",
    "亚朵酒店 · 康铂酒店",
    "香港玛丽医院",
  ];
  bItems.forEach((item, i) => {
    sl.addText("✅  " + item, { x: 0.8, y: 1.5 + i * 0.52, w: 3.9, h: 0.45, fontSize: 12, color: C.text, fontFace: "Calibri" });
  });

  // C端 card
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 0.85, w: 4.25, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 0.85, w: 4.25, h: 0.5, fill: { color: C.orange } });
  sl.addText("👤  C端：全国总代渠道", { x: 5.35, y: 0.85, w: 4, h: 0.5, fontSize: 13, color: C.white, fontFace: "Calibri", bold: true, valign: "middle" });
  sl.addText("朱国勇", { x: 5.35, y: 1.55, w: 3.9, h: 0.4, fontSize: 22, color: C.navy, fontFace: "Georgia", bold: true });
  sl.addText("全国C端总代", { x: 5.35, y: 1.95, w: 3.9, h: 0.3, fontSize: 12, color: C.muted, fontFace: "Calibri" });
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.35, y: 2.35, w: 2.5, h: 0.52, fill: { color: C.orange }, rectRadius: 0.08 });
  sl.addText("已签约 ¥1000万/年", { x: 5.35, y: 2.35, w: 2.5, h: 0.52, fontSize: 13, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });

  // 出海 card
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 3.5, w: 4.25, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 3.5, w: 4.25, h: 0.5, fill: { color: "7c3aed" } });
  sl.addText("🌍  出海：70+国家", { x: 5.35, y: 3.5, w: 4, h: 0.5, fontSize: 13, color: C.white, fontFace: "Calibri", bold: true, valign: "middle" });
  sl.addText("欧洲 · 东南亚 · 中东主力市场", { x: 5.35, y: 4.15, w: 3.9, h: 0.35, fontSize: 12, color: C.text, fontFace: "Calibri" });
  sl.addText("CE / FDA 等国际认证", { x: 5.35, y: 4.5, w: 3.9, h: 0.35, fontSize: 12, color: C.muted, fontFace: "Calibri" });

  // bottom summary
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.1, w: 10, h: 0.525, fill: { color: C.navy } });
  sl.addText("B端深耕  +  C端爆发  +  全球布局 = 三轮驱动增长模型", { x: 0.6, y: 5.1, w: 9, h: 0.525, fontSize: 14, color: C.white, fontFace: "Calibri", bold: true, valign: "middle", align: "center" });
}

// ─── Slide 7: CASES ─────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("06  已落地案例", { x: 0.6, y: 0.35, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("已验证的商业落地", { x: 0.6, y: 0.8, w: 9, h: 0.6, fontSize: 28, color: C.navy, fontFace: "Georgia", bold: true });

  const cases = [
    { name: "北京协和医院", type: "睡眠中心", desc: "战略合作签约", icon: "🏥", color: C.accent },
    { name: "亚朵酒店", type: "连锁酒店", desc: "客房升级改造", icon: "🏨", color: C.primary },
    { name: "康铂酒店", type: "连锁酒店", desc: "客房升级改造", icon: "🏩", color: C.primary },
    { name: "中海锦年", type: "康养机构", desc: "养老机构落地", icon: "🧓", color: C.orange },
  ];

  cases.forEach((c, i) => {
    const x = 0.6 + (i % 2) * 4.6;
    const y = 1.55 + Math.floor(i / 2) * 1.7;
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.3, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h: 1.5, fill: { color: c.color } });
    sl.addText(c.icon, { x: x + 0.25, y: y + 0.25, w: 0.7, h: 0.7, fontSize: 32 });
    sl.addText(c.name, { x: x + 1.05, y: y + 0.2, w: 3.0, h: 0.45, fontSize: 15, color: C.text, fontFace: "Calibri", bold: true });
    sl.addText(c.type, { x: x + 1.05, y: y + 0.65, w: 3.0, h: 0.3, fontSize: 12, color: c.color, fontFace: "Calibri" });
    sl.addText(c.desc, { x: x + 1.05, y: y + 0.95, w: 3.0, h: 0.3, fontSize: 11, color: C.muted, fontFace: "Calibri" });
  });

  // more cases bar
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.9, w: 8.8, h: 0.6, fill: { color: C.navy } });
  sl.addText("20+ 三甲医院正在合作洽谈中（301医院 · 北医三院 · 安贞医院 · 香港玛丽医院）", { x: 0.8, y: 4.9, w: 8.4, h: 0.6, fontSize: 12, color: C.white, fontFace: "Calibri", valign: "middle" });
}

// ─── Slide 8: TECH MOAT ─────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("07  技术护城河", { x: 0.6, y: 0.35, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });

  // 3 pillars full width
  const moats = [
    { icon: "⚗️", title: "材料提纯", color: C.accent, desc: "原材料经多道提纯工艺，有效去除杂质离子，确保负氧离子释放的稳定性与安全性，避免副产物风险。", items: ["多道提纯去除有害杂质", "释放浓度安全可控", "零臭氧、零辐射", "森林级浓度无醉氧"] },
    { icon: "🏭", title: "小批量精造", color: "7c3aed", desc: "成熟的小批量生产体系，灵活响应客户定制需求，已有多次成功交付记录，质量稳定可靠。", items: ["小批量订单交付能力", "支持个性化定制", "全流程质量检测", "特普丽50年工艺保障"] },
    { icon: "🛠️", title: "完整服务方案", color: C.orange, desc: "从环境评估、方案设计、施工部署到长期监测，提供交钥匙工程式完整服务。", items: ["负氧离子微环境监测部署", "整体施工改造（3天/间）", "睡眠监测系统配套", "6个月跟踪监测服务"] },
  ];

  moats.forEach((m, i) => {
    const x = 0.6 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.85, w: 2.9, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.85, w: 2.9, h: 0.1, fill: { color: m.color } });
    sl.addText(m.icon, { x, y: 1.0, w: 2.9, h: 0.6, fontSize: 30, align: "center" });
    sl.addText(m.title, { x, y: 1.6, w: 2.9, h: 0.4, fontSize: 15, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
    sl.addShape(pres.shapes.LINE, { x: x + 0.25, y: 2.05, w: 2.4, h: 0, line: { color: C.silver, width: 1 } });
    sl.addText(m.desc, { x: x + 0.15, y: 2.15, w: 2.6, h: 1.0, fontSize: 10, color: C.muted, fontFace: "Calibri" });
    m.items.forEach((item, j) => {
      sl.addText("→  " + item, { x: x + 0.15, y: 3.2 + j * 0.42, w: 2.6, h: 0.4, fontSize: 10, color: C.text, fontFace: "Calibri" });
    });
  });

  // process flow
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.05, w: 10, h: 0.575, fill: { color: C.navy } });
  sl.addText("实施流程：环境评估 → 材料提纯 → 小批量生产 → 施工部署（3天/间）→ 6个月跟踪监测", { x: 0.6, y: 5.05, w: 9, h: 0.575, fontSize: 12, color: C.white, fontFace: "Calibri", valign: "middle", align: "center" });
}

// ─── Slide 9: TEAM ───────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("08  跨学科团队", { x: 0.6, y: 0.35, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });

  const team = [
    { icon: "🎓", name: "杨起", role: "特普丽创始人", desc: "50年室内材料研发经验\n企业战略决策核心", tag: "战略决策", tagColor: C.accent },
    { icon: "📖", name: "学术带头人", role: "北京化工大学资深教授", desc: "负氧离子技术学术权威\n产学研融合推动者", tag: "学术指导", tagColor: C.primary },
    { icon: "🔬", name: "临床医学团队", role: "睡眠医学专家", desc: "睡眠医学专家\n临床验证设计执行", tag: "循证医学", tagColor: "7c3aed" },
    { icon: "💼", name: "朱国勇", role: "C端全国总代", desc: "全国市场拓展\nC端渠道建设", tag: "已签约 ¥1000万", tagColor: C.orange },
  ];

  team.forEach((t, i) => {
    const x = 0.6 + i * 2.35;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.85, w: 2.15, h: 3.7, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.85, w: 2.15, h: 0.85, fill: { color: C.navy } });
    sl.addText(t.icon, { x, y: 0.9, w: 2.15, h: 0.55, fontSize: 26, align: "center" });
    sl.addText(t.name, { x, y: 1.8, w: 2.15, h: 0.45, fontSize: 15, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
    sl.addText(t.role, { x, y: 2.2, w: 2.15, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri", align: "center" });
    sl.addShape(pres.shapes.LINE, { x: x + 0.3, y: 2.65, w: 1.55, h: 0, line: { color: C.silver, width: 1 } });
    sl.addText(t.desc, { x: x + 0.1, y: 2.75, w: 1.95, h: 0.9, fontSize: 10, color: C.muted, fontFace: "Calibri", align: "center" });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: 3.75, w: 1.65, h: 0.45, fill: { color: t.tagColor }, rectRadius: 0.08 });
    sl.addText(t.tag, { x: x + 0.25, y: 3.75, w: 1.65, h: 0.45, fontSize: 10, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });
  });

  // stats row
  const stats = [
    { val: "17", label: "国际期刊论文" },
    { val: "40对", label: "临床验证" },
    { val: "2", label: "核心发明专利" },
    { val: "50年", label: "行业经验" },
  ];
  stats.forEach((s, i) => {
    const x = 0.6 + i * 2.35;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 4.7, w: 2.15, h: 0.8, fill: { color: C.light } });
    sl.addText(s.val, { x, y: 4.72, w: 2.15, h: 0.45, fontSize: 22, color: C.primary, fontFace: "Georgia", bold: true, align: "center" });
    sl.addText(s.label, { x, y: 5.15, w: 2.15, h: 0.3, fontSize: 10, color: C.muted, fontFace: "Calibri", align: "center" });
  });
}

// ─── Slide 10: COOPERATION ───────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("09  战略合作", { x: 0.6, y: 0.35, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });

  // left: 协和合作
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.85, w: 5.4, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.85, w: 5.4, h: 0.5, fill: { color: C.primary } });
  sl.addText("🏥  与协和共建课题", { x: 0.8, y: 0.85, w: 5, h: 0.5, fontSize: 14, color: C.white, fontFace: "Calibri", bold: true, valign: "middle" });
  sl.addText("北京协和医院睡眠中心", { x: 0.8, y: 1.5, w: 5, h: 0.4, fontSize: 18, color: C.navy, fontFace: "Georgia", bold: true });
  sl.addText([
    { text: "✅ 战略合作签约，共同开展临床研究", options: { breakLine: true } },
    { text: "✅ 探索负氧离子微环境对睡眠障碍患者的干预机制", options: { breakLine: true } },
    { text: "✅ 为DRG支付改革下的创新医疗器械准入提供循证依据", options: {} },
  ], { x: 0.8, y: 2.0, w: 5, h: 1.2, fontSize: 12, color: C.text, fontFace: "Calibri", paraSpaceAfter: 6 });

  // right top: 校企合作
  sl.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 0.85, w: 3.2, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 0.85, w: 3.2, h: 0.5, fill: { color: C.accent } });
  sl.addText("🔬  校企合作", { x: 6.4, y: 0.85, w: 2.8, h: 0.5, fontSize: 14, color: C.white, fontFace: "Calibri", bold: true, valign: "middle" });
  sl.addText("特普丽 × 北京化工大学", { x: 6.4, y: 1.5, w: 2.8, h: 0.4, fontSize: 14, color: C.navy, fontFace: "Calibri", bold: true });
  sl.addText("产学研深度融合\n持续技术迭代升级\n产品创新核心支撑", { x: 6.4, y: 1.95, w: 2.8, h: 1.2, fontSize: 12, color: C.muted, fontFace: "Calibri" });

  // bottom: 认证
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.5, w: 8.8, h: 1.45, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText("🏆  资质与认证", { x: 0.8, y: 3.6, w: 3, h: 0.4, fontSize: 13, color: C.primary, fontFace: "Calibri", bold: true });
  const certs = ["专精特新企业", "CE认证", "FDA注册", "国检报告", "2项发明专利"];
  certs.forEach((c, i) => {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8 + i * 1.75, y: 4.1, w: 1.6, h: 0.65, fill: { color: C.light }, rectRadius: 0.08 });
    sl.addText(c, { x: 0.8 + i * 1.75, y: 4.1, w: 1.6, h: 0.65, fontSize: 11, color: C.text, fontFace: "Calibri", align: "center", valign: "middle" });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.05, w: 10, h: 0.575, fill: { color: C.navy } });
  sl.addText("特普丽 × 北京化工大学  ·  产学研深度融合  ·  专精特新企业认定", { x: 0.6, y: 5.05, w: 9, h: 0.575, fontSize: 13, color: C.white, fontFace: "Calibri", valign: "middle", align: "center" });
}

// ─── Slide 11: VISION ────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.navy };

  // left accent
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent } });

  sl.addText("10  愿景", { x: 0.6, y: 0.4, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("智能睡眠", { x: 0.6, y: 0.9, w: 8, h: 0.9, fontSize: 52, color: C.white, fontFace: "Georgia", bold: true });
  sl.addText("环境改善新材料领导者", { x: 0.6, y: 1.75, w: 8, h: 0.5, fontSize: 20, color: C.accent, fontFace: "Calibri" });

  sl.addShape(pres.shapes.LINE, { x: 0.6, y: 2.4, w: 3, h: 0, line: { color: C.accent, width: 2 } });

  // three vision pillars
  const vision = [
    { icon: "🏥", title: "医疗级睡眠管理", desc: "以医院睡眠中心为核心场景，建立循证医学支撑的智能睡眠干预标准" },
    { icon: "🏠", title: "家庭健康守护", desc: "将医疗级环境改善技术带入千家万户，让每个人都能享受健康空气" },
    { icon: "🌍", title: "全球化布局", desc: "深耕欧洲、东南亚、中东市场，服务全球70+国家，输出中国创新" },
  ];

  vision.forEach((v, i) => {
    const x = 0.6 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 2.7, w: 2.9, h: 2.2, fill: { color: C.primary } });
    sl.addText(v.icon, { x, y: 2.85, w: 2.9, h: 0.55, fontSize: 26, align: "center" });
    sl.addText(v.title, { x, y: 3.45, w: 2.9, h: 0.45, fontSize: 14, color: C.white, fontFace: "Calibri", bold: true, align: "center" });
    sl.addText(v.desc, { x: x + 0.15, y: 3.95, w: 2.6, h: 0.85, fontSize: 10, color: "B0C4DE", fontFace: "Calibri" });
  });

  // bottom
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.1, w: 10, h: 0.525, fill: { color: C.accent } });
  sl.addText("安全 · 有效 · 持久  →  智能睡眠引领者", { x: 0.6, y: 5.1, w: 9, h: 0.525, fontSize: 15, color: C.navy, fontFace: "Calibri", bold: true, valign: "middle", align: "center" });
}

// ─── Slide 12: FUNDING ASK ───────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.navy };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent } });

  sl.addText("融资计划", { x: 0.6, y: 0.4, w: 4, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("¥2000万", { x: 0.6, y: 0.85, w: 6, h: 1.2, fontSize: 80, color: C.white, fontFace: "Georgia", bold: true });
  sl.addText("融资目标  ·  投前估值 ¥1亿", { x: 0.6, y: 2.0, w: 6, h: 0.4, fontSize: 16, color: C.accent, fontFace: "Calibri" });

  sl.addShape(pres.shapes.LINE, { x: 0.6, y: 2.55, w: 8.8, h: 0, line: { color: "2a4a7a", width: 1 } });

  // use of funds
  const funds = [
    { pct: "40%", amount: "¥800万", use: "研发投入", desc: "材料提纯 · 工艺迭代 · 新品开发", color: C.accent },
    { pct: "40%", amount: "¥800万", use: "市场拓展", desc: "医院渠道 · C端总代 · 出海布局", color: C.orange },
    { pct: "20%", amount: "¥400万", use: "团队建设", desc: "核心人才引进 · 运营体系完善", color: "7c3aed" },
  ];

  funds.forEach((f, i) => {
    const x = 0.6 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 2.75, w: 2.9, h: 2.15, fill: { color: C.primary } });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 2.75, w: 2.9, h: 0.1, fill: { color: f.color } });
    sl.addText(f.pct, { x, y: 2.95, w: 2.9, h: 0.65, fontSize: 32, color: f.color, fontFace: "Georgia", bold: true, align: "center" });
    sl.addText(f.amount, { x, y: 3.55, w: 2.9, h: 0.4, fontSize: 18, color: C.white, fontFace: "Calibri", bold: true, align: "center" });
    sl.addText(f.use, { x, y: 3.95, w: 2.9, h: 0.35, fontSize: 13, color: C.white, fontFace: "Calibri", bold: true, align: "center" });
    sl.addText(f.desc, { x: x + 0.1, y: 4.35, w: 2.7, h: 0.5, fontSize: 10, color: "B0C4DE", fontFace: "Calibri", align: "center" });
  });

  // bottom
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.05, w: 10, h: 0.575, fill: { color: C.accent } });
  sl.addText("联系邮箱：contact@deeptech.com  ·  手机：138-xxxx-xxxx", { x: 0.6, y: 5.05, w: 9, h: 0.575, fontSize: 13, color: C.navy, fontFace: "Calibri", bold: true, valign: "middle", align: "center" });
}

// ─── Write file ───────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "/Users/mma/Dropbox/0-睡眠管理/HICOOL_深睡科技_v1.pptx" })
  .then(() => console.log("✅ PPT生成成功: HICOOL_深睡科技_v1.pptx"))
  .catch(e => console.error("❌ 错误:", e));
