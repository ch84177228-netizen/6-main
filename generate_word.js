const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType,
  VerticalAlign, PageBreak
} = require("docx");
const fs = require("fs");

const navy       = "2C3E6B";
const pink       = "D95F7A";
const skyblue    = "4A9CC9";
const mint       = "5BAE8A";
const coral      = "E07050";
const lavender   = "7B6CB5";
const gold       = "D4A017";
const red        = "C0392B";
const bgCream    = "FAF6F0";
const bgPink     = "FDE8EC";
const bgBlue     = "E8F4FB";
const bgYellow   = "FEF9E7";

function header(text) {
  return new Paragraph({
    children: [new TextRun({ text, color: "FFFFFF", bold: true, size: 28 })],
    shading: { type: ShadingType.SOLID, color: navy },
    spacing: { before: 240, after: 120 },
    indent: { left: 200, right: 200 },
  });
}

function sectionTitle(num, text, color) {
  return new Paragraph({
    children: [new TextRun({ text: `【知识点${num}】${text}`, color: "FFFFFF", bold: true, size: 24 })],
    shading: { type: ShadingType.SOLID, color },
    spacing: { before: 300, after: 100 },
    indent: { left: 200 },
  });
}

function mnemonic(text, color) {
  return new Paragraph({
    children: [new TextRun({ text: `【口诀】${text}`, color, bold: true, size: 22 })],
    shading: { type: ShadingType.SOLID, color: bgCream },
    spacing: { before: 80, after: 80 },
    indent: { left: 300 },
  });
}

function tip(text) {
  return new Paragraph({
    children: [new TextRun({ text: `📌 诗雨叮嘱：${text}`, color: pink, size: 20 })],
    shading: { type: ShadingType.SOLID, color: bgPink },
    spacing: { before: 80, after: 80 },
    indent: { left: 300, right: 200 },
  });
}

function warn(text) {
  return new Paragraph({
    children: [new TextRun({ text: `⚠️ 易错：${text}`, color: coral, bold: true, size: 20 })],
    shading: { type: ShadingType.SOLID, color: bgYellow },
    spacing: { before: 80, after: 80 },
    indent: { left: 300, right: 200 },
  });
}

function example(text) {
  return new Paragraph({
    children: [new TextRun({ text: `📝 例：${text}`, color: "1A5276", size: 20 })],
    shading: { type: ShadingType.SOLID, color: bgBlue },
    spacing: { before: 80, after: 80 },
    indent: { left: 300, right: 200 },
  });
}

function body(text, color = "2C3E50") {
  return new Paragraph({
    children: [new TextRun({ text, color, size: 22 })],
    spacing: { before: 60, after: 60 },
    indent: { left: 300 },
  });
}

function partTitle(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 28, color: navy })],
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 400, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: navy } },
  });
}

function blank() {
  return new Paragraph({ text: "", spacing: { before: 60, after: 60 } });
}

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 900, bottom: 900, left: 900, right: 900 },
      },
    },
    children: [

      // ── PART A ──────────────────────────────────────────────
      partTitle("PART A  爆款标题方案（3选1）"),

      new Paragraph({
        children: [new TextRun({ text: "方案① 【推荐】数字悬念型", bold: true, color: pink, size: 24 })],
        spacing: { before: 160, after: 60 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "增值税计算题总丢分？销项税额这5个坑必须知道", size: 22 })],
        indent: { left: 300 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "推荐理由：数字「5个坑」制造具体感，「总丢分」戳痛点，高爆款率。", color: "666666", size: 20 })],
        indent: { left: 300 },
        spacing: { after: 120 },
      }),

      new Paragraph({
        children: [new TextRun({ text: "方案② 反常识型", bold: true, color: skyblue, size: 24 })],
        spacing: { before: 100, after: 60 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "含税价×税率=销项税额？这个公式是错的", size: 22 })],
        indent: { left: 300 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "推荐理由：反常识开头，引发点击欲，适合热门选题。", color: "666666", size: 20 })],
        indent: { left: 300 },
        spacing: { after: 120 },
      }),

      new Paragraph({
        children: [new TextRun({ text: "方案③ 场景共鸣型", bold: true, color: mint, size: 24 })],
        spacing: { before: 100, after: 60 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "销项税额老是算错？学姐帮你把坑全踩一遍", size: 22 })],
        indent: { left: 300 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "推荐理由：学姐人设感强，亲切感高，适合涨粉期内容。", color: "666666", size: 20 })],
        indent: { left: 300 },
        spacing: { after: 200 },
      }),

      // ── PART B ──────────────────────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      partTitle("PART B  图文内容稿"),

      header("CPA税法 - 增值税 | 销项税额 · 计税依据 · 差额征税"),

      blank(),
      new Paragraph({
        children: [new TextRun({ text: "增值税计算题，很多人第一步就错了——", size: 22, color: "2C3E50" })],
        indent: { left: 200 },
        spacing: { before: 100, after: 60 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "含税价直接乘税率，算出来的结果，全错。", size: 22, color: "2C3E50" })],
        indent: { left: 200 },
        spacing: { after: 60 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "今天这篇把销项税额的计算逻辑、计税依据的判断，还有差额征税的范围，一次梳理清楚。", size: 22, color: "2C3E50" })],
        indent: { left: 200 },
        spacing: { after: 160 },
      }),

      // 知识点①
      sectionTitle("①", "销项税额基础公式", pink),
      mnemonic("含税价÷（1+税率）×税率；不含税价直接乘", pink),
      body("含税价（价税合计）的处理：先还原为不含税价，再乘税率。"),
      example("含税价113元，税率13%\n→ 不含税价 = 113 ÷ 1.13 ≈ 100元\n→ 销项税额 = 100 × 13% = 13元"),
      body("不含税价的处理：直接乘税率，无需还原。"),
      tip("增值税是\"价外税\"，含税价里的税是\"加进去\"的，不是\"乘进去\"的。所以要先把税剥出来，再算税额。考试里最常见的坑就是含税价直接乘——一步错，全错。"),
      warn("含税价 × 税率 ≠ 销项税额，别被数字迷惑。"),

      blank(),

      // 知识点②
      sectionTitle("②", "视同销售的计税依据", skyblue),
      mnemonic("有同类按同类，无同类用组价；组价=成本×（1+利润率）", skyblue),
      body("判断逻辑："),
      body("1. 有同类货物近期销售价格 → 按同类货物销售价格"),
      body("2. 没有同类 → 按组成计税价格"),
      body("组成计税价格公式："),
      body("一般货物：组价 = 成本 × （1 + 成本利润率）", coral),
      body("消费税应税货物：组价 = 成本 × （1 + 成本利润率）÷ （1 - 消费税税率）", coral),
      example("企业将自产产品用于员工福利，同类产品售价10000元，则按10000元计税；若无同类，成本8000元，利润率10%，则组价 = 8000 × 1.1 = 8800元。"),
      tip("视同销售题目的重点在于\"有没有同类\"，有同类先用同类，别直接算组价。很多同学记反了。"),

      blank(),

      // 知识点③
      sectionTitle("③", "折扣销售的处理", mint),
      mnemonic("同票可扣，另票不扣；备注栏注明不算数", mint),
      body("折扣额与销售额在同一张发票的金额栏注明 → 可按折扣后金额计税"),
      body("分开开票，或只在备注栏写折扣 → 不能扣减，全额计税"),
      example("销售额10000元，在同一张发票金额栏注明折扣1000元，则按9000元计算销项税额。"),
      warn("备注栏写了折扣不算！考试里选项经常拿\"备注栏注明\"来迷惑你，牢记只认金额栏。"),

      blank(),

      // 知识点④
      sectionTitle("④", "差额征税的7种情形", coral),
      mnemonic("旅游建筑金融经纪，客运航代差额记", coral),
      body("7类可以差额征税的情形："),
      body("🔹 旅游服务：扣除住宿、交通、门票等代收转付费用"),
      body("🔹 建筑服务：扣除分包款"),
      body("🔹 金融商品转让：卖出价减去买入价"),
      body("🔹 经纪代理：扣除代为支付的政府性基金"),
      body("🔹 客运场站服务：扣除支付给承运方的费用"),
      body("🔹 航空地面服务：扣除支付给航空公司的费用"),
      body("🔹 其他经纪代理：扣除代收转付款项"),
      tip("差额征税考题常问\"能不能扣\"，把7种情形背熟，其他情形一律不能差额征税，按全额计税。"),
      warn("金融商品转让当期亏损不得跨期结转，这是高频考点，很多人忽略。"),

      blank(),

      // 知识点⑤
      sectionTitle("⑤", "纳税义务发生时间", lavender),
      mnemonic("先开票按票，收钱按收，赊销按合同，预收看发货，委代孰早", lavender),

      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: "销售方式", bold: true, color: "FFFFFF", size: 20 })] })],
                shading: { type: ShadingType.SOLID, color: skyblue },
                verticalAlign: VerticalAlign.CENTER,
              }),
              new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: "纳税义务发生时间", bold: true, color: "FFFFFF", size: 20 })] })],
                shading: { type: ShadingType.SOLID, color: skyblue },
                verticalAlign: VerticalAlign.CENTER,
              }),
            ],
          }),
          ...[
            ["先开发票", "开票当天"],
            ["收到款项", "收款当天"],
            ["赊销/分期收款", "合同约定收款日期"],
            ["预收货款", "货物发出当天"],
            ["委托代销", "收到代销清单或收到全部款项，孰早"],
          ].map(([l, r]) => new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: l, size: 20 })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r, size: 20 })] })] }),
            ],
          })),
        ],
      }),

      blank(),
      warn("预收货款≠收款即确认，货还没发出，销项税额还不确认。考试里最爱考这个。"),

      blank(),

      // 小结
      new Paragraph({
        children: [new TextRun({ text: "🌟 本篇小结", bold: true, color: navy, size: 26 })],
        shading: { type: ShadingType.SOLID, color: bgCream },
        spacing: { before: 200, after: 80 },
        indent: { left: 200 },
      }),
      body("① 含税价要先还原，不能直接乘税率"),
      body("② 折扣只认同一张发票金额栏"),
      body("③ 差额征税只有那7种，其他情形别套用"),
      body("基础打牢，计算题分才拿得稳。💪"),

      blank(),

      // ── PART C ──────────────────────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      partTitle("PART C  ChatGPT出图Prompt（英文）"),

      new Paragraph({
        children: [new TextRun({
          text: `Create a kawaii stationery-style vertical infographic layout (3:4 ratio, 1080×1440px) with a cream/off-white background (#FAF6F0). The design should have a clean, warm, and cute hand-drawn notebook aesthetic.

Layout structure (top to bottom):
1. A navy (#2C3E6B) header banner with white text placeholder area (leave all text areas completely blank — no Chinese or English text)
2. Five colored knowledge section cards, each with:
   - A colored title bar (colors: #D95F7A, #4A9CC9, #5BAE8A, #E07050, #7B6CB5)
   - A light cream mnemonic banner area (blank)
   - Bullet point lines (blank placeholder lines)
   - Small colored accent boxes for examples (light blue #E8F4FB) and tips (light pink #FDE8EC) and warnings (light yellow #FEF9E7)
3. A small summary box at the bottom with a cream background

All text areas must be LEFT BLANK — no words, no characters. Generate only the visual frame and layout. Sticker-like decorative elements (small stars, dots, pencil icons) are welcome. Do not add any watermarks.`,
          size: 20,
          color: "2C3E50",
        })],
        spacing: { before: 120, after: 120 },
        indent: { left: 200, right: 200 },
      }),

      blank(),

      // ── PART D ──────────────────────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      partTitle("PART D  Canva文字排版清单"),

      ...[
        ["页眉", "CPA税法 - 增值税 | 销项税额 · 计税依据 · 差额征税", "白色，粗体，居中，26pt"],
        ["知识点①标题条", "【知识点①】销项税额基础公式", "白色，粗体，22pt，背景#D95F7A"],
        ["口诀①", "【口诀】含税价÷（1+税率）×税率；不含税价直接乘", "深玫红，粗体，20pt"],
        ["例题①", "📝 例：含税价113元，税率13%→销项税额=113÷1.13×13%=13元", "深蓝，20pt，蓝色底"],
        ["诗雨叮嘱①", "📌 诗雨叮嘱：增值税是价外税，含税价里的税是加进去的，要先还原再乘税率。", "玫红，20pt，粉色底"],
        ["易错①", "⚠️ 易错：含税价 × 税率 ≠ 销项税额", "橙色，粗体，20pt，黄色底"],
        ["知识点②标题条", "【知识点②】视同销售的计税依据", "白色，粗体，22pt，背景#4A9CC9"],
        ["口诀②", "【口诀】有同类按同类，无同类用组价；组价=成本×（1+利润率）", "天蓝，粗体，20pt"],
        ["诗雨叮嘱②", "📌 诗雨叮嘱：有同类先用同类价格，别直接算组价，很多人记反了。", "玫红，20pt，粉色底"],
        ["知识点③标题条", "【知识点③】折扣销售的处理", "白色，粗体，22pt，背景#5BAE8A"],
        ["口诀③", "【口诀】同票可扣，另票不扣；备注栏注明不算数", "薄荷绿，粗体，20pt"],
        ["易错③", "⚠️ 易错：备注栏写折扣不算，只认同一张发票金额栏", "橙色，粗体，20pt，黄色底"],
        ["知识点④标题条", "【知识点④】差额征税7种情形", "白色，粗体，22pt，背景#E07050"],
        ["口诀④", "【口诀】旅游建筑金融经纪，客运航代差额记", "橙红，粗体，20pt"],
        ["易错④", "⚠️ 易错：金融商品转让当期亏损不得跨期结转", "橙色，粗体，20pt，黄色底"],
        ["知识点⑤标题条", "【知识点⑤】纳税义务发生时间", "白色，粗体，22pt，背景#7B6CB5"],
        ["口诀⑤", "【口诀】先开票按票，收钱按收，赊销按合同，预收看发货，委代孰早", "紫色，粗体，20pt"],
        ["易错⑤", "⚠️ 易错：预收货款≠收款即确认，货发出才确认销项", "橙色，粗体，20pt，黄色底"],
        ["小结", "🌟 本篇小结：①含税价先还原 ②折扣只认金额栏 ③差额征税只有7种", "深蓝，粗体，20pt，奶油底"],
      ].map(([label, content, style]) => new Paragraph({
        children: [
          new TextRun({ text: `【${label}】`, bold: true, color: navy, size: 20 }),
          new TextRun({ text: ` ${content}`, size: 20, color: "2C3E50" }),
          new TextRun({ text: `  →  ${style}`, size: 18, color: "888888" }),
        ],
        spacing: { before: 80, after: 80 },
        indent: { left: 200 },
      })),

      blank(),

      // ── PART E ──────────────────────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      partTitle("PART E  小红书正文文案 + 话题标签"),

      new Paragraph({
        children: [new TextRun({ text: "标题：增值税计算题总丢分？销项税额这5个坑必须知道", bold: true, size: 24, color: navy })],
        spacing: { before: 120, after: 80 },
        indent: { left: 200 },
      }),

      new Paragraph({
        children: [new TextRun({ text: "正文：", bold: true, size: 22, color: navy })],
        spacing: { before: 80, after: 40 },
        indent: { left: 200 },
      }),

      ...[
        "📚 增值税计算题，很多人第一步就错了——",
        "含税价直接乘税率，算出来的结果，全错。",
        "今天这篇把销项税额的计算逻辑、计税依据的判断，还有差额征税的范围，一次梳理清楚。",
        "",
        "🎯 【知识点①】销项税额基础公式",
        "【口诀】含税价÷（1+税率）×税率；不含税价直接乘",
        "含税价（价税合计）的处理：先还原为不含税价，再乘税率。",
        "📝 例：含税价113元，税率13%→销项税额=113÷1.13×13%=13元",
        "📌 诗雨叮嘱：增值税是\"价外税\"，含税价里的税是\"加进去\"的，先剥出来再算。考试最常见的坑。",
        "⚠️ 易错：含税价 × 税率 ≠ 销项税额",
        "",
        "📋 【知识点②】视同销售的计税依据",
        "【口诀】有同类按同类，无同类用组价；组价=成本×（1+利润率）",
        "有同类货物近期销售价格→按同类价格",
        "没有同类→按组成计税价格（成本×（1+利润率））",
        "📌 诗雨叮嘱：有同类先用同类价格，别直接算组价，很多人记反了。",
        "",
        "📋 【知识点③】折扣销售的处理",
        "【口诀】同票可扣，另票不扣；备注栏注明不算数",
        "折扣额在同一张发票金额栏注明→按折扣后金额计税",
        "分开开票/备注栏写折扣→全额计税",
        "⚠️ 易错：备注栏写了折扣不算！只认同一张发票金额栏。",
        "",
        "🌱 【知识点④】差额征税7种情形",
        "【口诀】旅游建筑金融经纪，客运航代差额记",
        "旅游/建筑/金融商品转让/经纪代理/客运场站/航空地面/其他代理",
        "📌 诗雨叮嘱：除这7种外，其他情形一律全额计税，别乱套。",
        "⚠️ 易错：金融商品转让当期亏损不得跨期结转。",
        "",
        "✨ 【知识点⑤】纳税义务发生时间",
        "【口诀】先开票按票，收钱按收，赊销按合同，预收看发货，委代孰早",
        "⚠️ 易错：预收货款≠收款即确认，货物发出才确认销项。",
        "",
        "🌟 本篇小结",
        "销项税额这一块，记住三件事：",
        "① 含税价要先还原，不能直接乘税率",
        "② 折扣只认同一张发票金额栏",
        "③ 差额征税只有那7种，其他情形别套用",
        "基础打牢，计算题分才拿得稳。💪",
      ].map(line => new Paragraph({
        children: [new TextRun({ text: line, size: 20, color: "2C3E50" })],
        spacing: { before: 40, after: 40 },
        indent: { left: 300 },
      })),

      blank(),
      new Paragraph({
        children: [new TextRun({ text: "话题：#注会cpa #cpa备考 #注会 #CPA税法 #增值税 #会计 #财会 #诗雨会计", size: 20, color: "888888" })],
        indent: { left: 200 },
        spacing: { before: 80 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "时间：", size: 20, color: "888888" })],
        indent: { left: 200 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "账号：06红书店铺｜姗姗 店铺大号", size: 20, color: "888888" })],
        indent: { left: 200 },
        spacing: { after: 200 },
      }),

    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("CPA税法-增值税③销项税额_内容全套.docx", buf);
  console.log("生成成功：CPA税法-增值税③销项税额_内容全套.docx");
});
