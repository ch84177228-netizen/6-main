const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, ShadingType, PageBreak
} = require("docx");
const fs = require("fs");

const NAVY  = "2C3E6B";
const RED   = "C8522A";
const DARK  = "1A1816";
const GRAY  = "6B6560";
const BLUE  = "1A5FA8";
const PINK  = "D95F7A";
const GREEN = "5BAE8A";

const divider = () => new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" } },
  spacing: { before: 160, after: 160 },
});
const blank = () => new Paragraph({ text: "", spacing: { before: 60, after: 60 } });
const h2 = (text, color = RED) => new Paragraph({
  children: [new TextRun({ text, bold: true, size: 26, color })],
  spacing: { before: 240, after: 80 },
});
const p = (text, color = DARK, indent = 300) => new Paragraph({
  children: [new TextRun({ text, size: 22, color })],
  spacing: { before: 60, after: 60 },
  indent: { left: indent },
});

const groups = [
  {
    label: "第1组 · 主号·甄嬛传版",
    account: "06红书店铺|粉色手机",
    color: NAVY,
    title: "收入确认，一下通透了",
    style: "甄嬛传故事类比，老师视角权威讲解",
    topics: "#注会cpa #中级会计 #cpa备考 #中级会计备考 #会计 #财会 #收入确认 #甄嬛传 #诗雨会计",
    scripts: [
      {张: "第1张 · 封面", text: "收入确认，一下通透了\n越看越懂！甄嬛传视角带你入门"},
      {张: "第2张 · 背景铺垫", text: `熹贵妃在宫中卖货，什么时候算【收入】？\n很多考生觉得——钱到手了，收入就确认了。\n但注会/中级会计对收入的要求，比这复杂得多。\n【收到钱】不等于【确认收入】\n收入要等到控制权转移才算数。`},
      {张: "第3张 · 核心逻辑：5步法", text: "新收入准则核心：5步法\n第一步：识别合同\n第二步：识别履约义务（最常考！）\n第三步：确定交易价格\n第四步：分摊交易价格\n第五步：履约时确认收入（控制权何时转移？）\n\n📌 诗雨叮嘱：考试最常考第二步和第五步。"},
      {张: "第4张 · 识别履约义务", text: "熹贵妃接了一个大单：\n→ 给皇上送御膳（送货）\n→ 负责1年的餐后点心（后续服务）\n这是一个合同，但有两个履约义务：\n① 送御膳（某时点转移控制权）\n② 点心服务（1年内持续提供）\n\n⚠️ 易错：误以为一个合同=一个义务，没有拆分。"},
      {张: "第5张 · 时点 vs 时段", text: "控制权在某个时点转移 → 时点法确认\n例：货物交付给买方的那一刻\n\n控制权在一段时间内持续转移 → 时段法确认\n例：建造合同、长期服务合同\n\n判断时段法的3个标准（满足一个即可）：\n① 客户在履约同时即取得并消耗利益\n② 企业履约创造或增强客户控制的资产\n③ 企业履约创造的资产无替代用途，且有权收取款项"},
      {张: "第6张 · 甄嬛传场景举例", text: "甄嬛接了两单：\n单A：给太后送一批绣品，交货即完成。\n→ 时点法，交货时一次确认收入。\n单B：给皇后宫殿每月提供打扫服务，合同期1年。\n→ 时段法，每个月确认当月服务收入。"},
      {张: "第7张 · 高频考点陷阱", text: "❌ 收到预付款就确认收入\n✅ 预付款是合同负债，履约后才转收入\n\n❌ 合同修改后，直接在原合同上改收入\n✅ 要判断修改是单独合同还是原合同变更\n\n❌ 附有销售退回条款，全额确认收入\n✅ 预计会退回的部分，要单独挂退货负债，不算收入"},
      {张: "第8张 · 诗雨叮嘱·总结", text: "收入确认这章，记住三件事：\n① 5步法是框架，第2、5步是重点\n② 时点 vs 时段，看控制权转移性质，不看收钱时间\n③ 预付款、附退回条款、合同修改——三大高频坑\n\n控制权转移了，收入才算。\n钱到手不等于入账。💪"},
    ],
  },
  {
    label: "第2组 · 副号1·蜡笔小新版",
    account: "06红书店铺|黄色手机",
    color: GREEN,
    title: "靠小新家，我终于搞懂收入确认了",
    style: "蜡笔小新故事类比，备考人自述共鸣",
    topics: "#注会cpa #中级会计 #cpa备考 #中级会计备考 #会计 #财会 #收入确认 #蜡笔小新 #诗雨会计",
    scripts: [
      {张: "第1张 · 封面", text: "靠小新家，我终于搞懂收入确认了"},
      {张: "第2张 · 我之前的状态", text: "备考看到收入确认5步法，完全不知道在说什么。\n直到我用小新家类比，才突然开窍。"},
      {张: "第3张 · 小新家类比", text: "美冴妈妈接了一个大装修单：\n→ 负责交付整套装修成果（一次性）\n→ 合同还包含后续3年维保服务\n这是一个合同，但有两个独立的履约义务！\n① 装修交付 → 时点法，验收完成时确认\n② 维保服务 → 时段法，3年内持续确认\n核心：控制权转移，才算确认收入。"},
      {张: "第4张 · 我终于开窍的地方", text: "以前我一直以为：收到钱=确认收入。\n但小新家的例子让我明白：\n收到定金，是合同负债（欠客户的服务）\n服务做完了，才转成收入\n\n广志爸爸预付了装修款，美冴不能马上入账收入。\n必须等装修完成、验收通过，控制权转移给广志家。"},
      {张: "第5张 · 时点 vs 时段怎么判断", text: "一次性交付 → 时点法\n持续提供服务 → 时段法\n\n判断时段法3个标准（满足一个就够）：\n① 客户同步取得并消耗利益\n② 企业在客户控制的资产上履约\n③ 无替代用途，且有权收取完工部分款项"},
      {张: "第6张 · 考试高频坑+总结", text: "❌ 收到预付款直接确认收入\n❌ 附退回条款的全额确认\n❌ 一个合同只有一个履约义务\n这三个坑搞懂之后，错题率降了一半。\n\n控制权转移了，收入才算数。💪"},
    ],
  },
  {
    label: "第3组 · 副号2·甄嬛速记版",
    account: "06红书店铺|蓝色ipad张菊芳",
    color: PINK,
    title: "收入确认考点｜甄嬛传5分钟背完",
    style: "甄嬛传精简速记版，收藏向",
    topics: "#注会cpa #中级会计 #cpa备考 #中级会计备考 #会计 #财会 #收入确认 #甄嬛传 #诗雨会计",
    scripts: [
      {张: "第1张 · 封面", text: "收入确认考点\n甄嬛传5分钟背完\n速记版·建议收藏"},
      {张: "第2张 · 口诀+5步法", text: "【口诀】控制权转移，收入才到位\n\n甄嬛版5步法：\n接单→拆任务→定价格→分钱→交货入账\n\n正式版5步法：\n识别合同→识别履约义务→确定交易价→分摊→履约确认"},
      {张: "第3张 · 时点 vs 时段", text: "时点法：控制权一次性转移（发货、验收）\n时段法：控制权持续转移（服务、建造）\n\n时段法3个标准（满足1个即可）：\n① 客户同步取得并消耗利益\n② 企业创造并增强客户控制的资产\n③ 无替代用途+有权收取款项\n\n⚠️ 易错：时段法不只限于建造合同！"},
      {张: "第4张 · 高频易错对比", text: "❌ 收到预付款=确认收入\n✅ 预付款挂合同负债\n\n❌ 附退回条款全额确认\n✅ 预计退回部分挂退货负债\n\n❌ 合同修改直接改收入\n✅ 先判断是新合同还是变更\n\n❌ 一合同=一义务\n✅ 一合同可含多个独立履约义务"},
      {张: "第5张 · 可变对价+总结", text: "可变对价估计方法：\n→ 期望值法（多种结果加权平均）\n→ 最可能金额法（概率最大的那个）\n\n【必背】\n①有多个履约义务→按独立售价比例分摊\n②时段法进度=实际成本/预算总成本（投入法）\n\n建议收藏，考前回看✅"},
    ],
  },
  {
    label: "第4组 · 副号3·避坑吐槽版",
    account: "06红书店铺|孙文新134 已实名",
    color: "E07050",
    title: "学收入确认，我踩过的3个深坑",
    style: "过来人吐槽口吻，痛点引流",
    topics: "#注会cpa #中级会计 #cpa备考 #中级会计备考 #会计 #财会 #收入确认 #诗雨会计",
    scripts: [
      {张: "第1张 · 封面", text: "学收入确认\n我踩过的3个深坑\n让你少走弯路"},
      {张: "第2张 · 反面教材", text: "收入确认这章，我当初做题错得想砸书。\n看了书，觉得自己懂了。做题，全错。\n反复看，反复错。\n后来才发现：不是我不努力，是踩了坑没发现。"},
      {张: "第3张 · 坑一：预付款直接确认收入", text: "我当初的做法：\n收到客户预付款，贷主营业务收入，一步到位。\n\n真相：\n收到预付款，应该贷合同负债。\n等我把货/服务交付完，才能转成收入。\n\n收到钱不等于赚到钱。"},
      {张: "第4张 · 坑二：退回条款全额确认", text: "我当初的想法：反正货已经卖出去了，先全部确认收入。\n\n真相：\n合同里有退货条款时，预计会退回的部分不算收入，\n要单独挂退货负债。\n\n题目出现【预计退货率30%】，就是在考这个坑。"},
      {张: "第5张 · 坑三+总结", text: "坑三：一个合同只认一个履约义务\n\n真相：\n一份合同可以包含多个独立履约义务，\n比如【卖设备+包3年维修】，设备和维修要分开确认。\n\n三个坑搞清楚，做题正确率直接上去：\n① 预付款 → 合同负债\n② 附退回条款 → 预计退回部分挂退货负债\n③ 一合同 → 可能有多个独立履约义务"},
    ],
  },
  {
    label: "第5组 · 副号4·纯考点总结版",
    account: "06红书店铺|孙文新 小号店铺号",
    color: "7B6CB5",
    title: "收入确认核心考点｜收藏这一篇就够了",
    style: "文字版原稿，清晰简洁考点整理",
    topics: "#注会cpa #中级会计 #cpa备考 #中级会计备考 #会计 #财会 #收入确认 #考点整理 #诗雨会计",
    scripts: [
      {张: "第1张 · 封面", text: "收入确认核心考点\n收藏这一篇就够了\n注会·中级会计"},
      {张: "第2张 · 5步法框架", text: "✅ 核心框架：5步法\n① 识别合同\n② 识别履约义务（一合同可含多个独立义务）\n③ 确定交易价格（含可变对价、重大融资成分）\n④ 将交易价格分摊至各履约义务（按独立售价比例）\n⑤ 履约时确认收入（控制权转移时点）"},
      {张: "第3张 · 时点法 vs 时段法", text: "✅ 时点法：控制权在某时点一次性转移\n✅ 时段法：控制权在一段时间内持续转移\n\n判断时段法3标准（满足1个即可）：\n① 客户同步取得并消耗利益\n② 企业创造并增强客户控制的资产\n③ 无替代用途且有权收取款项"},
      {张: "第4张 · 高频考点", text: "✅ 高频考点\n→ 预付款 = 合同负债，不是收入\n→ 附退回条款：预计退回部分挂退货负债\n→ 合同修改：判断是否构成独立合同\n→ 可变对价：期望值法或最可能金额法\n→ 重大融资成分：调整交易价格\n→ 质保条款：判断保证型还是服务型"},
      {张: "第5张 · 常见错误+总结", text: "❌ 收到预付款=确认收入 → ✅ 挂合同负债\n❌ 附退回条款全额确认 → ✅ 挂退货负债\n❌ 合同修改直接改收入 → ✅ 判断新合同还是变更\n❌ 一合同=一义务 → ✅ 可含多个独立义务\n\n✅ 常考数字：时段法标准3个，满足1个即可\n\n建议收藏，考前回看✅"},
    ],
  },
  {
    label: "第6组 · 副号5·考前冲刺版",
    account: "06红书店铺|姗姗已实名",
    color: "C0392B",
    title: "距离中级考试105天｜收入确认必考考点今天搞定",
    style: "中级考试倒计时，紧迫感强",
    topics: "#中级会计 #中级会计备考 #中级会计考试 #注会cpa #备考冲刺 #收入确认 #9月上岸 #诗雨会计",
    scripts: [
      {张: "第1张 · 封面", text: "距离中级考试还有105天\n收入确认 · 必考考点今天搞定\n9月5日稳稳上岸💪"},
      {张: "第2张 · 为什么必须搞懂", text: "收入确认是中级会计实务高频考点。\n选择题+主观题每年都考，分值稳定。\n5步法是框架，考题围绕第2步（识别履约义务）和第5步（控制权转移时点）出。\n主观题常给一段业务描述，让你判断：\n① 有几个履约义务？\n② 什么时候确认收入？"},
      {张: "第3张 · 高频考点TOP5", text: "① 识别独立履约义务（一合同多义务，分别确认）\n② 时点法 vs 时段法判断\n③ 预付款 → 合同负债，不是收入\n④ 附退回条款的销售处理（挂退货负债）\n⑤ 可变对价的估计方法（期望值/最可能金额）"},
      {张: "第4张 · 选择题陷阱", text: "⚠️ 选择题常见陷阱：\n→ 收到钱 = 确认收入 ❌\n→ 一个合同只有一个履约义务 ❌\n→ 时段法只有建造合同才能用 ❌\n→ 合同修改后直接调整原收入 ❌\n看到这些选项，直接排除。"},
      {张: "第5张 · 备考安排+冲刺总结", text: "📅 建议备考节奏：\n今天：弄懂5步法框架和时点/时段判断\n明天：专项练习预付款、退回条款题型\n后天：做2-3年真题，找题感\n\n记住一句话：控制权转移了，收入才算数。\n\n✅ 5步法框架背下来\n✅ 时点/时段判断搞清楚\n✅ 三大高频坑绕开\n\n9月5日稳稳上岸💪"},
    ],
  },
];

const children = [];

groups.forEach((g, i) => {
  if (i > 0) children.push(new Paragraph({ children: [new PageBreak()] }));

  children.push(new Paragraph({
    children: [new TextRun({ text: g.label, bold: true, size: 30, color: "FFFFFF" })],
    shading: { type: ShadingType.SOLID, color: g.color },
    spacing: { before: 200, after: 160 },
    indent: { left: 200 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: `风格：${g.style}`, size: 20, color: GRAY })],
    spacing: { after: 40 }, indent: { left: 200 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: `账号：${g.account}`, size: 20, color: GRAY })],
    spacing: { after: 200 }, indent: { left: 200 },
  }));

  children.push(h2("📌 标题方案", g.color));
  children.push(p(`推荐标题：${g.title}`));
  children.push(blank());

  children.push(h2("📋 图文脚本", g.color));
  g.scripts.forEach(({ 张: zhang, text }) => {
    children.push(new Paragraph({
      children: [new TextRun({ text: zhang, bold: true, size: 22, color: g.color })],
      spacing: { before: 160, after: 60 }, indent: { left: 200 },
    }));
    text.split("\n").forEach(line => children.push(p(line, DARK, 400)));
    children.push(blank());
  });

  children.push(divider());
  children.push(new Paragraph({
    children: [
      new TextRun({ text: "话题：", size: 20, color: GRAY }),
      new TextRun({ text: g.topics, size: 20, color: BLUE }),
    ],
    spacing: { before: 60 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: "时间：", size: 20, color: GRAY })],
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: `账号：${g.account}`, size: 20, color: GRAY })],
    spacing: { after: 120 },
  }));
});

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1200, bottom: 1200, left: 1200, right: 1200 },
      },
    },
    children: [
      new Paragraph({
        children: [new TextRun({ text: "01_收入确认_6账号完整脚本", bold: true, size: 44, color: NAVY })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 200 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "诗雨学姐·内容运营系统 | 注会·中级会计系列", size: 24, color: GRAY })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
      }),
      divider(),
      ...children,
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("01_收入确认_6账号完整脚本.docx", buf);
  console.log("生成成功：01_收入确认_6账号完整脚本.docx");
});
