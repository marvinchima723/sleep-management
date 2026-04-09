const pptxgen = require("pptxgenjs");
let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "特普丽新材料 - 深睡科技 HICOOL 2026";
pres.author = "深睡科技";

const C = {
  navy:    "0d1b2a",
  primary: "1a3a6b",
  accent:  "00d4aa",
  orange:  "ff6b35",
  light:   "f0f4f8",
  white:   "FFFFFF",
  text:    "2c3e50",
  muted:   "7f8c8d",
  silver:  "E2E8F0",
};
const makeShadow = () => ({ type: "outer", blur: 5, offset: 2, angle: 135, color: "000000", opacity: 0.10 });

// ─── Slide 1: COVER ──────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.navy };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent } });

  sl.addText("特普丽新材料", { x: 0.7, y: 1.2, w: 8, h: 0.6, fontSize: 18, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 4 });
  sl.addText("深睡科技", { x: 0.7, y: 1.85, w: 8, h: 1.3, fontSize: 72, color: C.white, fontFace: "Georgia", bold: true });
  sl.addText("被动释放 · 持久改善 · 还原自然空气", { x: 0.7, y: 3.2, w: 8, h: 0.5, fontSize: 20, color: C.accent, fontFace: "Calibri" });
  sl.addShape(pres.shapes.LINE, { x: 0.7, y: 3.85, w: 4.5, h: 0, line: { color: C.accent, width: 2 } });

  const pills = ["杀菌 99.9%", "抗病毒 99.1%", "除甲醛 88.7%"];
  pills.forEach((t, i) => {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7 + i * 2.9, y: 4.15, w: 2.7, h: 0.52, fill: { color: C.primary }, rectRadius: 0.1 });
    sl.addText(t, { x: 0.7 + i * 2.9, y: 4.15, w: 2.7, h: 0.52, fontSize: 14, color: C.white, fontFace: "Calibri", align: "center", valign: "middle" });
  });

  sl.addText("HICOOL 2026  ·  创业大赛", { x: 0.7, y: 5.1, w: 5, h: 0.35, fontSize: 12, color: C.muted, fontFace: "Calibri" });
  sl.addText("全球首创  ·  引领行业50年  ·  连续13年市场占有率第一", { x: 0.7, y: 4.78, w: 8, h: 0.3, fontSize: 11, color: "4a6fa5", fontFace: "Calibri" });
  sl.addText("汇报人：杨帆", { x: 8.5, y: 5.1, w: 1.2, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri", align: "right" });
}

// ─── Slide 2: MARKET PAIN ───────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 3.8, h: 5.625, fill: { color: C.navy } });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 3.8, h: 0.12, fill: { color: C.accent } });
  sl.addText("01", { x: 0.4, y: 0.4, w: 1, h: 0.4, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 3 });
  sl.addText("市场痛点", { x: 0.4, y: 0.75, w: 3, h: 0.35, fontSize: 12, color: "8fa3bf", fontFace: "Calibri" });
  sl.addText("3", { x: 0.4, y: 1.3, w: 1.5, h: 1.6, fontSize: 110, color: C.white, fontFace: "Georgia", bold: true });
  sl.addText("亿人", { x: 1.85, y: 2.35, w: 1.5, h: 0.6, fontSize: 28, color: C.accent, fontFace: "Calibri", bold: true });
  sl.addText("中国人受睡眠问题困扰", { x: 0.4, y: 3.1, w: 3.1, h: 0.7, fontSize: 13, color: "8fa3bf", fontFace: "Calibri" });
  sl.addShape(pres.shapes.LINE, { x: 0.4, y: 3.85, w: 2.2, h: 0, line: { color: "2a4a7a", width: 1 } });
  sl.addText("市场年规模", { x: 0.4, y: 4.05, w: 3, h: 0.3, fontSize: 11, color: "6b8ab5", fontFace: "Calibri" });
  sl.addText("¥4700亿+", { x: 0.4, y: 4.35, w: 3, h: 0.5, fontSize: 24, color: C.accent, fontFace: "Georgia", bold: true });
  sl.addText("CAGR > 12%，持续高速增长", { x: 0.4, y: 4.85, w: 3.1, h: 0.4, fontSize: 10, color: "6b8ab5", fontFace: "Calibri" });

  const pains = [
    { num: "01", title: "现有助眠方式效果有限", desc: "药物依赖、副作用大、治标不治本。现有助眠产品无法从根本上改善睡眠环境质量。", color: C.accent },
    { num: "02", title: "医院睡眠中心覆盖率极低", desc: "全国睡眠中心数量不足千家，资源稀缺、费用高昂、体验差，绝大部分患者得不到专业帮助。", color: C.orange },
    { num: "03", title: "室内环境污染持续损害健康", desc: "甲醛超标、PM2.5渗入、辐射无处不在——室内环境问题成为睡眠障碍的重要诱因，却长期被忽视。", color: C.primary },
  ];
  pains.forEach((p, i) => {
    const y = 0.4 + i * 1.72;
    sl.addText(p.num, { x: 4.1, y, w: 0.7, h: 0.5, fontSize: 28, color: "D0D8E4", fontFace: "Georgia", bold: true });
    sl.addShape(pres.shapes.OVAL, { x: 4.85, y: y + 0.12, w: 0.18, h: 0.18, fill: { color: p.color } });
    sl.addText(p.title, { x: 5.15, y, w: 4.5, h: 0.45, fontSize: 15, color: C.text, fontFace: "Calibri", bold: true });
    sl.addText(p.desc, { x: 5.15, y: y + 0.48, w: 4.5, h: 1.0, fontSize: 12, color: C.muted, fontFace: "Calibri" });
    if (i < 2) sl.addShape(pres.shapes.LINE, { x: 4.1, y: y + 1.55, w: 5.5, h: 0, line: { color: C.silver, width: 0.5 } });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 4.1, y: 5.1, w: 5.5, h: 0.42, fill: { color: C.navy } });
  sl.addText("核心矛盾：睡眠问题日益严峻，现有解决方案无法满足安全、有效、可持续的市场需求", { x: 4.3, y: 5.1, w: 5.3, h: 0.42, fontSize: 10, color: C.white, fontFace: "Calibri", valign: "middle" });
}

// ─── Slide 3: THEORY ─────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("02", { x: 0.5, y: 0.28, w: 0.6, h: 0.35, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("理论支撑与技术架构", { x: 1.05, y: 0.28, w: 5, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri" });
  sl.addText("负氧离子改善睡眠的四大作用机制", { x: 0.5, y: 0.68, w: 9, h: 0.45, fontSize: 22, color: C.navy, fontFace: "Georgia", bold: true });

  const mechs = [
    { num: "1", title: "神经内分泌调节", desc: "促进褪黑素分泌，调节5-羟色胺等单胺类神经递质，从根本上改善睡眠节律", color: C.accent },
    { num: "2", title: "自主神经平衡", desc: "增强副交感神经活性，降低交感神经兴奋性，减少夜间觉醒次数", color: C.primary },
    { num: "3", title: "环境优化", desc: "吸附清除空气颗粒物，创造洁净睡眠环境，减少过敏与呼吸刺激", color: C.orange },
    { num: "4", title: "抗氧化与抗炎", desc: "减轻与睡眠障碍相关的氧化应激和炎症反应，改善睡眠质量", color: "7c3aed" },
  ];
  mechs.forEach((m, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 4.75;
    const y = 1.25 + row * 1.65;
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h: 1.5, fill: { color: m.color } });
    sl.addText(m.num, { x: x + 0.25, y: y + 0.15, w: 0.55, h: 0.55, fontSize: 22, color: m.color, fontFace: "Georgia", bold: true });
    sl.addText(m.title, { x: x + 0.85, y: y + 0.18, w: 3.5, h: 0.42, fontSize: 15, color: C.text, fontFace: "Calibri", bold: true });
    sl.addText(m.desc, { x: x + 0.85, y: y + 0.6, w: 3.5, h: 0.8, fontSize: 12, color: C.muted, fontFace: "Calibri" });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.6, w: 10, h: 1.025, fill: { color: C.navy } });
  sl.addText("研究启示：", { x: 0.5, y: 4.72, w: 1.5, h: 0.35, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true });
  sl.addText("从美国宾夕法尼亚大学医学院Kornblueh et al 1958年奠基性研究，到中国医科大学Wang & Li 2024年最新发现，负氧离子改善睡眠的证据链从主观报告到客观PSG数据，从单一机制到多靶点通路不断完善。", { x: 2.0, y: 4.72, w: 7.5, h: 0.85, fontSize: 11.5, color: C.white, fontFace: "Calibri" });
}

// ─── Slide 4: SOLUTION ───────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.navy };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 4.2, h: 5.625, fill: { color: "0a1628" } });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.accent } });
  sl.addText("03", { x: 0.45, y: 0.45, w: 1, h: 0.35, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 3 });
  sl.addText("解决方案", { x: 0.45, y: 0.75, w: 3, h: 0.3, fontSize: 11, color: "6b8ab5", fontFace: "Calibri" });
  sl.addText("AMIC", { x: 0.45, y: 1.25, w: 3.5, h: 1.1, fontSize: 72, color: C.white, fontFace: "Georgia", bold: true });
  sl.addText("无源负氧离子技术", { x: 0.45, y: 2.35, w: 3.5, h: 0.45, fontSize: 17, color: C.accent, fontFace: "Calibri", bold: true });
  sl.addShape(pres.shapes.LINE, { x: 0.45, y: 2.95, w: 2.5, h: 0, line: { color: "2a4a7a", width: 1 } });
  sl.addText("被动释放 · 无需能耗 · 持久改善\n还原森林级清新空气，无醉氧风险", { x: 0.45, y: 3.1, w: 3.4, h: 0.75, fontSize: 12, color: "8fa3bf", fontFace: "Calibri" });

  const stats = [
    { val: "99.9%", label: "抗菌率" },
    { val: "93%", label: "功能持久性" },
    { val: "0", label: "能耗·臭氧·辐射" },
  ];
  stats.forEach((s, i) => {
    const y = 3.95 + i * 0.55;
    sl.addText(s.val, { x: 0.45, y, w: 1.8, h: 0.4, fontSize: 16, color: C.accent, fontFace: "Georgia", bold: true });
    sl.addText(s.label, { x: 2.25, y: y + 0.05, w: 1.8, h: 0.35, fontSize: 11, color: "8fa3bf", fontFace: "Calibri" });
  });

  const pillars = [
    {
      num: "01", title: "材料提纯",
      desc: "独创四步法核心提纯工艺——酸溶、载体共沉/选择性吸附/离子交换分离Ra、矿物纯化、再合成——彻底去除放射性镭（Ra），保留并激活负氧离子释放活性，技术壁垒高、难复制。",
      items: ["酸溶将镭转入溶液体系", "载体共沉/选择性吸附/离子交换分离Ra", "纯化主体矿物组分", "重新合成高活性负氧离子粉"],
      color: C.accent,
    },
    {
      num: "02", title: "小批量精造",
      desc: "成熟的小批量生产体系，灵活响应定制需求，已有多次成功交付记录，质量稳定可靠。",
      items: ["小批量订单交付能力", "支持个性化定制", "全流程质量检测"],
      color: "7c3aed",
    },
    {
      num: "03", title: "完整服务方案",
      desc: "从环境评估、方案设计、施工部署到长期监测，提供交钥匙工程式完整服务。",
      items: ["负氧离子微环境监测部署", "整体施工改造（3天/间）", "6个月跟踪监测服务"],
      color: C.orange,
    },
  ];
  pillars.forEach((p, i) => {
    const x = 4.55 + i * 1.82;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.45, w: 1.7, h: 4.6, fill: { color: C.primary } });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.45, w: 1.7, h: 0.1, fill: { color: p.color } });
    sl.addText(p.num, { x, y: 0.65, w: 1.7, h: 0.4, fontSize: 24, color: p.color, fontFace: "Georgia", bold: true, align: "center" });
    sl.addText(p.title, { x, y: 1.1, w: 1.7, h: 0.4, fontSize: 13, color: C.white, fontFace: "Calibri", bold: true, align: "center" });
    sl.addShape(pres.shapes.LINE, { x: x + 0.25, y: 1.55, w: 1.2, h: 0, line: { color: "2a4a7a", width: 0.5 } });
    sl.addText(p.desc, { x: x + 0.1, y: 1.65, w: 1.5, h: 1.2, fontSize: 9.5, color: "a0b4cc", fontFace: "Calibri" });
    p.items.forEach((item, j) => {
      sl.addShape(pres.shapes.OVAL, { x: x + 0.12, y: 2.92 + j * 0.52, w: 0.12, h: 0.12, fill: { color: p.color } });
      sl.addText(item, { x: x + 0.3, y: 2.85 + j * 0.52, w: 1.35, h: 0.5, fontSize: 9.5, color: C.white, fontFace: "Calibri" });
    });
  });
}

// ─── Slide 5: PRODUCT ────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("04", { x: 0.5, y: 0.28, w: 0.6, h: 0.35, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("产品矩阵", { x: 1.05, y: 0.28, w: 3, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri" });
  sl.addText("可调控睡眠环境 · 完整服务体系", { x: 0.5, y: 0.68, w: 9, h: 0.45, fontSize: 22, color: C.navy, fontFace: "Georgia", bold: true });

  const products = [
    {
      badge: "核心产品", name: "负氧离子板材", price: "¥250-300", unit: "/㎡",
      color: C.accent,
      features: ["被动持续释放负氧离子", "抗菌抗病毒除甲醛", "墙面/天花板/地面通用", "持久性93%，长期有效"],
    },
    {
      badge: "定制款", name: "定制壁纸/窗帘", price: "¥150", unit: "/㎡",
      color: "7c3aed",
      features: ["个性化图案定制", "融入室内设计美学", "窗帘+负氧离子二合一", "已有多次小批量交付"],
    },
    {
      badge: "服务型", name: "完整解决方案", price: "交钥匙", unit: "工程",
      color: C.orange,
      features: ["环境评估+方案设计", "施工改造（3天/间）", "6个月跟踪监测", "技术支持全流程"],
    },
  ];
  products.forEach((p, i) => {
    const x = 0.5 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.9, h: 3.55, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.9, h: 0.55, fill: { color: p.color } });
    sl.addText(p.badge, { x, y: 1.2, w: 2.9, h: 0.55, fontSize: 11, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });
    sl.addText(p.name, { x, y: 1.85, w: 2.9, h: 0.45, fontSize: 15, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
    sl.addText(p.price, { x, y: 2.3, w: 2.9, h: 0.45, fontSize: 22, color: p.color, fontFace: "Georgia", bold: true, align: "center" });
    sl.addText(p.unit, { x, y: 2.72, w: 2.9, h: 0.3, fontSize: 11, color: C.muted, fontFace: "Calibri", align: "center" });
    sl.addShape(pres.shapes.LINE, { x: x + 0.3, y: 3.1, w: 2.3, h: 0, line: { color: C.silver, width: 1 } });
    p.features.forEach((f, j) => {
      sl.addText("→  " + f, { x: x + 0.15, y: 3.22 + j * 0.42, w: 2.65, h: 0.4, fontSize: 11, color: C.text, fontFace: "Calibri" });
    });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.88, w: 9, h: 0.62, fill: { color: C.navy } });
  sl.addText("💡  价格优势：比进口竞品低 50-70%  |  国内市场份额 25%，行业第一  |  板材/壁纸/窗帘全线覆盖", { x: 0.7, y: 4.88, w: 8.6, h: 0.62, fontSize: 13, color: C.white, fontFace: "Calibri", valign: "middle" });
}

// ─── Slide 6: BUSINESS MODEL ─────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("05", { x: 0.5, y: 0.28, w: 0.6, h: 0.35, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("商业模式", { x: 1.05, y: 0.28, w: 3, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri" });
  sl.addText("B端 · C端 · 出海三线并进", { x: 0.5, y: 0.68, w: 9, h: 0.45, fontSize: 22, color: C.navy, fontFace: "Georgia", bold: true });

  const bms = [
    {
      icon: "🏢", title: "B端：医院 & 康养", color: C.primary,
      items: ["北京协和医院睡眠中心（已签约）", "301医院 · 北医三院 · 安贞医院", "中海锦年康养 · 房协100+机构", "亚朵酒店 · 康铂酒店", "香港玛丽医院"],
      highlight: "✅ 已签约",
    },
    {
      icon: "👤", title: "C端：全国总代渠道", color: C.accent,
      items: ["朱国勇 全国C端总代", "千万级大V", "已签约 ¥1000万/年", "全国渠道建设", "C端市场全面拓展"],
      highlight: "✅ ¥1000万/年",
    },
    {
      icon: "🌍", title: "出海：70+国家", color: C.orange,
      items: ["欧洲 · 东南亚 · 中东主力市场", "CE / FDA 等国际认证", "产品已出口 70+ 国家", "全球化布局", "输出中国创新"],
      highlight: "✅ 全球覆盖",
    },
  ];
  bms.forEach((b, i) => {
    const x = 0.5 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.9, h: 3.7, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.9, h: 0.12, fill: { color: b.color } });
    sl.addText(b.icon, { x, y: 1.38, w: 2.9, h: 0.55, fontSize: 28, align: "center" });
    sl.addText(b.title, { x, y: 1.95, w: 2.9, h: 0.42, fontSize: 14, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
    sl.addShape(pres.shapes.LINE, { x: x + 0.25, y: 2.42, w: 2.4, h: 0, line: { color: C.silver, width: 1 } });
    b.items.forEach((item, j) => {
      sl.addText("→  " + item, { x: x + 0.15, y: 2.55 + j * 0.42, w: 2.65, h: 0.4, fontSize: 11, color: C.text, fontFace: "Calibri" });
    });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: 4.58, w: 2.3, h: 0.38, fill: { color: b.color }, rectRadius: 0.08 });
    sl.addText(b.highlight, { x: x + 0.3, y: 4.58, w: 2.3, h: 0.38, fontSize: 11, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.0, w: 10, h: 0.625, fill: { color: C.navy } });
  sl.addText("核心优势：高端睡眠空间改造——一日换新，睡眠空间改造，单日完成", { x: 0.5, y: 5.0, w: 9, h: 0.625, fontSize: 14, color: C.white, fontFace: "Calibri", valign: "middle", align: "center" });
}

// ─── Slide 7: CASES ──────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("06", { x: 0.5, y: 0.28, w: 0.6, h: 0.35, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("已落地案例", { x: 1.05, y: 0.28, w: 3, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri" });
  sl.addText("已验证的商业落地 · 持续扩大中", { x: 0.5, y: 0.68, w: 9, h: 0.45, fontSize: 22, color: C.navy, fontFace: "Georgia", bold: true });

  const cases = [
    { icon: "🏥", name: "北京协和医院", sub: "睡眠中心", type: "医院", tag: "战略合作签约", tagColor: C.accent, note: "预期产出SCI文章3-4篇" },
    { icon: "🏨", name: "亚朵酒店", sub: "连锁酒店", type: "酒店", tag: "客房升级改造", tagColor: C.orange, note: "预期改造1000+间/年" },
    { icon: "🏩", name: "康铂酒店", sub: "连锁酒店", type: "酒店", tag: "客房升级改造", tagColor: C.orange, note: "预期改造1000+间/年" },
    { icon: "🧓", name: "中海锦年", sub: "康养机构", type: "康养", tag: "养老机构落地", tagColor: C.primary, note: "预期改造2000+间/年" },
  ];
  cases.forEach((c, i) => {
    const x = 0.5 + i * 2.35;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.15, h: 2.65, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.15, h: 0.1, fill: { color: c.tagColor } });
    sl.addText(c.icon, { x, y: 1.38, w: 2.15, h: 0.6, fontSize: 30, align: "center" });
    sl.addText(c.name, { x, y: 2.0, w: 2.15, h: 0.4, fontSize: 13, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
    sl.addText(c.sub, { x, y: 2.38, w: 2.15, h: 0.3, fontSize: 10, color: C.muted, fontFace: "Calibri", align: "center" });
    sl.addShape(pres.shapes.LINE, { x: x + 0.25, y: 2.75, w: 1.65, h: 0, line: { color: C.silver, width: 1 } });
    sl.addText(c.note, { x, y: 2.85, w: 2.15, h: 0.35, fontSize: 10, color: C.muted, fontFace: "Calibri", align: "center" });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: 3.3, w: 1.75, h: 0.35, fill: { color: c.tagColor }, rectRadius: 0.07 });
    sl.addText(c.tag, { x: x + 0.2, y: 3.3, w: 1.75, h: 0.35, fontSize: 9.5, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.05, w: 9, h: 0.65, fill: { color: C.navy } });
  sl.addText("20+ 三甲医院正在合作洽谈中：301医院 · 北医三院 · 安贞医院 · 香港玛丽医院", { x: 0.7, y: 4.05, w: 8.6, h: 0.65, fontSize: 13, color: C.white, fontFace: "Calibri", valign: "middle" });

  // Stats row
  const stats = [
    { val: "1000+", label: "成功家装案例" },
    { val: "2000+", label: "合作高端酒店" },
    { val: "NO.1", label: "全国市场占有率" },
  ];
  stats.forEach((s, i) => {
    const x = 0.5 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 4.85, w: 2.9, h: 0.7, fill: { color: C.light } });
    sl.addText(s.val, { x, y: 4.88, w: 2.9, h: 0.42, fontSize: 20, color: C.primary, fontFace: "Georgia", bold: true, align: "center" });
    sl.addText(s.label, { x, y: 5.28, w: 2.9, h: 0.25, fontSize: 10, color: C.muted, fontFace: "Calibri", align: "center" });
  });
}

// ─── Slide 8: TECH MOAT ──────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("07", { x: 0.5, y: 0.28, w: 0.6, h: 0.35, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("技术护城河", { x: 1.05, y: 0.28, w: 3, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri" });

  const moats = [
    {
      icon: "⚗️", title: "材料提纯", color: C.accent,
      desc: "独创四步法核心提纯工艺——酸溶、载体共沉/选择性吸附/离子交换分离Ra、矿物纯化、再合成——彻底去除放射性镭（Ra），保留并激活负氧离子释放活性，技术壁垒高、难复制。",
      items: ["酸溶将镭转入溶液体系", "载体共沉/选择性吸附/离子交换分离Ra", "纯化主体矿物组分", "重新合成高活性负氧离子粉", "零臭氧、零辐射", "森林级浓度无醉氧"],
    },
    {
      icon: "🏭", title: "小批量精造", color: "7c3aed",
      desc: "成熟的小批量生产体系，灵活响应客户定制需求，已有多次成功交付记录，质量稳定可靠。",
      items: ["小批量订单交付能力", "支持个性化定制", "全流程质量检测", "特普丽50年工艺保障"],
    },
    {
      icon: "🛠️", title: "完整服务方案", color: C.orange,
      desc: "从环境评估、方案设计、施工部署到长期监测，提供交钥匙工程式完整服务。",
      items: ["负氧离子微环境监测部署", "整体施工改造（3天/间）", "睡眠监测系统配套", "6个月跟踪监测服务"],
    },
  ];
  moats.forEach((m, i) => {
    const x = 0.5 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.72, w: 2.9, h: 3.95, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 0.72, w: 2.9, h: 0.1, fill: { color: m.color } });
    sl.addText(m.icon, { x, y: 0.88, w: 2.9, h: 0.6, fontSize: 30, align: "center" });
    sl.addText(m.title, { x, y: 1.48, w: 2.9, h: 0.4, fontSize: 15, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
    sl.addShape(pres.shapes.LINE, { x: x + 0.25, y: 1.93, w: 2.4, h: 0, line: { color: C.silver, width: 1 } });
    sl.addText(m.desc, { x: x + 0.15, y: 2.03, w: 2.6, h: 1.0, fontSize: 9.5, color: C.muted, fontFace: "Calibri" });
    m.items.forEach((item, j) => {
      sl.addText("→  " + item, { x: x + 0.15, y: 3.1 + j * 0.32, w: 2.6, h: 0.32, fontSize: 9.5, color: C.text, fontFace: "Calibri" });
    });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.77, w: 10, h: 0.855, fill: { color: C.navy } });
  sl.addText("实施流程：环境评估 → 材料提纯 → 小批量生产 → 施工部署（3天/间）→ 6个月跟踪监测", { x: 0.5, y: 4.77, w: 9, h: 0.5, fontSize: 12, color: C.white, fontFace: "Calibri", valign: "middle", align: "center" });
  sl.addText("特普丽 × 北京化工大学  ·  产学研深度融合  ·  专精特新企业认定", { x: 0.5, y: 5.28, w: 9, h: 0.35, fontSize: 11, color: C.accent, fontFace: "Calibri", align: "center" });
}

// ─── Slide 9: TEAM ───────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("08", { x: 0.5, y: 0.28, w: 0.6, h: 0.35, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("跨学科团队", { x: 1.05, y: 0.28, w: 3, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri" });
  sl.addText("产学研医投 · 全链条核心成员", { x: 0.5, y: 0.68, w: 9, h: 0.45, fontSize: 22, color: C.navy, fontFace: "Georgia", bold: true });

  const team = [
    { name: "杨帆", role: "发起人/CEO", org: "北京深睡科技", desc: "澳大利亚纽卡斯尔大学商学硕士；北京特普丽装饰装帧材料有限公司总经理；北京银河家墙纸有限公司CEO", initials: "YF", color: C.primary },
    { name: "李瑞锋", role: "总工程师", org: "北京深睡科技", desc: "中国建筑装饰装修材料协会专家委员会副主任委员；实用新型专利3项+发明专利2项；全国建材科技创新奖", initials: "LRF", color: C.accent },
    { name: "史仲广", role: "材料专家", org: "北京深睡科技", desc: "四川大学纺织与皮革专业客座教授；长安/红旗/赛里斯汽车材料研究院特聘专家；参与定制汽车负离子行业标准", initials: "SZG", color: "7c3aed" },
    { name: "马驰", role: "科研负责人", org: "北京深睡科技", desc: "北京协和医学院科研博士后；香港大学医学院骨科博士；亿元级科研转化项目CEO（红杉资本领投）", initials: "MC", color: C.orange },
    { name: "王磊", role: "医学顾问", org: "北京协和医院睡眠中心", desc: "北京协和医院睡眠医学专家；主导临床验证设计与医学研究合作；为产品功效提供循证支持", initials: "WL", color: C.primary },
    { name: "孙瑜", role: "CHO", org: "北京深睡科技", desc: "主导人才战略与组织文化建设；深度理解创业公司人力资源需求；构建高效协作的核心团队", initials: "SY", color: C.accent },
  ];
  team.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.5 + col * 3.1;
    const y = 1.2 + row * 2.15;
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 2.0, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.85, h: 2.0, fill: { color: t.color } });
    sl.addText(t.initials, { x, y: y + 0.65, w: 0.85, h: 0.7, fontSize: 20, color: C.white, fontFace: "Georgia", bold: true, align: "center", valign: "middle" });
    sl.addText(t.name, { x: x + 1.0, y: y + 0.15, w: 1.8, h: 0.38, fontSize: 15, color: C.text, fontFace: "Calibri", bold: true });
    sl.addText(t.role, { x: x + 1.0, y: y + 0.52, w: 1.8, h: 0.3, fontSize: 10, color: t.color, fontFace: "Calibri", bold: true });
    sl.addText(t.org, { x: x + 1.0, y: y + 0.8, w: 1.8, h: 0.25, fontSize: 9, color: C.muted, fontFace: "Calibri" });
    sl.addShape(pres.shapes.LINE, { x: x + 1.0, y: y + 1.12, w: 1.75, h: 0, line: { color: C.silver, width: 0.5 } });
    sl.addText(t.desc, { x: x + 1.0, y: y + 1.18, w: 1.75, h: 0.75, fontSize: 8.5, color: C.muted, fontFace: "Calibri" });
  });
}

// ─── Slide 10: COOPERATION ───────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("09", { x: 0.5, y: 0.28, w: 0.6, h: 0.35, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("战略合作", { x: 1.05, y: 0.28, w: 3, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri" });
  sl.addText("与顶级机构共建循证医学标准", { x: 0.5, y: 0.68, w: 9, h: 0.45, fontSize: 22, color: C.navy, fontFace: "Georgia", bold: true });

  //协和
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.5, h: 2.8, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.5, h: 0.1, fill: { color: C.accent } });
  sl.addText("🏥", { x: 0.5, y: 1.4, w: 4.5, h: 0.6, fontSize: 30, align: "center" });
  sl.addText("北京协和医院睡眠中心", { x: 0.5, y: 2.05, w: 4.5, h: 0.4, fontSize: 16, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
  sl.addShape(pres.shapes.LINE, { x: 0.7, y: 2.5, w: 4.1, h: 0, line: { color: C.silver, width: 1 } });
  const xiehe = ["✅ 战略合作签约，共同开展临床研究", "✅ 探索负氧离子微环境对睡眠障碍患者的干预机制", "✅ 为DRG支付改革下的创新医疗器械准入提供循证依据"];
  xiehe.forEach((item, i) => { sl.addText(item, { x: 0.7, y: 2.62 + i * 0.42, w: 4.1, h: 0.4, fontSize: 11.5, color: C.text, fontFace: "Calibri" }); });

  //校企
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.2, w: 4.35, h: 2.8, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.2, w: 4.35, h: 0.1, fill: { color: C.primary } });
  sl.addText("🔬", { x: 5.15, y: 1.4, w: 4.35, h: 0.6, fontSize: 30, align: "center" });
  sl.addText("特普丽 × 北京化工大学", { x: 5.15, y: 2.05, w: 4.35, h: 0.4, fontSize: 16, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
  sl.addShape(pres.shapes.LINE, { x: 5.35, y: 2.5, w: 3.95, h: 0, line: { color: C.silver, width: 1 } });
  sl.addText("产学研深度融合，持续技术迭代升级", { x: 5.35, y: 2.62, w: 3.95, h: 0.4, fontSize: 12, color: C.accent, fontFace: "Calibri", bold: true });
  const xieheItems = ["产品创新核心支撑", "AMIC技术迭代升级", "持续研发投入"];
  xieheItems.forEach((item, i) => { sl.addText("→  " + item, { x: 5.35, y: 3.1 + i * 0.38, w: 3.95, h: 0.36, fontSize: 11.5, color: C.text, fontFace: "Calibri" }); });

  // Badges
  const badges = [
    { label: "专精特新小巨人企业", color: C.primary },
    { label: "2项发明专利", color: C.accent },
    { label: "17篇国际论文", color: C.orange },
    { label: "40对40临床验证", color: "7c3aed" },
  ];
  badges.forEach((b, i) => {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5 + i * 2.35, y: 4.2, w: 2.2, h: 0.55, fill: { color: b.color }, rectRadius: 0.1 });
    sl.addText(b.label, { x: 0.5 + i * 2.35, y: 4.2, w: 2.2, h: 0.55, fontSize: 11, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.9, w: 10, h: 0.725, fill: { color: C.navy } });
  sl.addText("特普丽 × 北京化工大学  ·  产学研深度融合  ·  专精特新企业认定  ·  40+三甲医院合作洽谈中", { x: 0.5, y: 4.9, w: 9, h: 0.725, fontSize: 13, color: C.white, fontFace: "Calibri", valign: "middle", align: "center" });
}

// ─── Slide 11: VISION ────────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.navy };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.accent } });
  sl.addText("10", { x: 0.55, y: 0.4, w: 1, h: 0.4, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 3 });
  sl.addText("愿景", { x: 0.55, y: 0.72, w: 3, h: 0.3, fontSize: 11, color: "6b8ab5", fontFace: "Calibri" });
  sl.addText("智能睡眠", { x: 0.55, y: 1.2, w: 8, h: 1.0, fontSize: 64, color: C.white, fontFace: "Georgia", bold: true });
  sl.addText("环境改善新材料领导者", { x: 0.55, y: 2.2, w: 8, h: 0.5, fontSize: 20, color: C.accent, fontFace: "Calibri" });
  sl.addShape(pres.shapes.LINE, { x: 0.55, y: 2.85, w: 4, h: 0, line: { color: "2a4a7a", width: 1 } });

  const visions = [
    { icon: "🏥", title: "医疗级睡眠管理", desc: "以医院睡眠中心为核心场景，建立循证医学支撑的智能睡眠干预标准", color: C.accent },
    { icon: "🏠", title: "家庭健康守护", desc: "将医疗级环境改善技术带入千家万户，让每个人都能享受健康空气", color: "7c3aed" },
    { icon: "🌍", title: "全球化布局", desc: "深耕欧洲、东南亚、中东市场，服务全球70+国家，输出中国创新", color: C.orange },
  ];
  visions.forEach((v, i) => {
    const x = 0.55 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 3.05, w: 2.9, h: 2.1, fill: { color: C.primary } });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 3.05, w: 2.9, h: 0.1, fill: { color: v.color } });
    sl.addText(v.icon, { x, y: 3.2, w: 2.9, h: 0.6, fontSize: 28, align: "center" });
    sl.addText(v.title, { x, y: 3.82, w: 2.9, h: 0.4, fontSize: 13, color: C.white, fontFace: "Calibri", bold: true, align: "center" });
    sl.addText(v.desc, { x: x + 0.15, y: 4.28, w: 2.6, h: 0.8, fontSize: 10.5, color: "a0b4cc", fontFace: "Calibri" });
  });

  sl.addText("安全 · 有效 · 持久  →  智能睡眠引领者", { x: 0.55, y: 5.25, w: 9, h: 0.3, fontSize: 13, color: C.accent, fontFace: "Calibri" });
}

// ─── Slide 12: FUNDING ───────────────────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  sl.addText("11", { x: 0.5, y: 0.28, w: 0.6, h: 0.35, fontSize: 11, color: C.accent, fontFace: "Calibri", bold: true, charSpacing: 2 });
  sl.addText("融资计划", { x: 1.05, y: 0.28, w: 3, h: 0.35, fontSize: 11, color: C.muted, fontFace: "Calibri" });

  // Big number
  sl.addText("¥2000", { x: 0.5, y: 0.85, w: 4.5, h: 1.1, fontSize: 72, color: C.navy, fontFace: "Georgia", bold: true });
  sl.addText("万元", { x: 4.8, y: 1.35, w: 1.5, h: 0.5, fontSize: 24, color: C.navy, fontFace: "Calibri", bold: true });
  sl.addText("目标融资  ·  投前估值 ¥1亿", { x: 0.5, y: 1.98, w: 5, h: 0.4, fontSize: 14, color: C.accent, fontFace: "Calibri", bold: true });
  sl.addShape(pres.shapes.LINE, { x: 0.5, y: 2.48, w: 5, h: 0, line: { color: C.silver, width: 1 } });

  // Funds use
  const funds = [
    { pct: "40%", amount: "¥800万", use: "研发投入", desc: "材料提纯 · 工艺迭代 · 新品开发", color: C.accent },
    { pct: "40%", amount: "¥800万", use: "市场拓展", desc: "医院渠道 · C端总代 · 出海布局", color: C.orange },
    { pct: "20%", amount: "¥400万", use: "团队建设", desc: "核心人才引进 · 运营体系完善", color: "7c3aed" },
  ];
  funds.forEach((f, i) => {
    const x = 0.5 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 2.6, w: 2.9, h: 2.15, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 2.6, w: 2.9, h: 0.85, fill: { color: f.color } });
    sl.addText(f.pct, { x, y: 2.62, w: 2.9, h: 0.55, fontSize: 28, color: C.white, fontFace: "Georgia", bold: true, align: "center" });
    sl.addText(f.amount, { x, y: 3.12, w: 2.9, h: 0.32, fontSize: 12, color: C.white, fontFace: "Calibri", align: "center" });
    sl.addText(f.use, { x, y: 3.55, w: 2.9, h: 0.4, fontSize: 15, color: C.text, fontFace: "Calibri", bold: true, align: "center" });
    sl.addShape(pres.shapes.LINE, { x: x + 0.3, y: 3.98, w: 2.3, h: 0, line: { color: C.silver, width: 1 } });
    sl.addText(f.desc, { x: x + 0.15, y: 4.08, w: 2.6, h: 0.6, fontSize: 11, color: C.muted, fontFace: "Calibri", align: "center" });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.9, w: 10, h: 0.725, fill: { color: C.navy } });
  sl.addText("联系邮箱：echoecho_cn@hotmail.com  ·  手机：13910104381", { x: 0.5, y: 4.9, w: 9, h: 0.725, fontSize: 14, color: C.white, fontFace: "Calibri", valign: "middle", align: "center" });
}

// ─── Write ───────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "/Users/mma/Dropbox/0-睡眠管理/HICOOL_深睡科技_v2.pptx" })
  .then(() => console.log("✅ 生成成功: HICOOL_深睡科技_v2.pptx"))
  .catch(e => console.error("❌ 生成失败:", e));