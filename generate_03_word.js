const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, ShadingType
} = require("docx");
const fs = require("fs");

const RED  = "C8522A";
const DARK = "1A1816";
const GRAY = "6B6560";
const BLUE = "1A5FA8";
const bgLight = "FFF8F5";

function divider() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" } },
    spacing: { before: 120, after: 120 },
  });
}
function acct() {
  return new Paragraph({
    children: [new TextRun({ text: "账号：06红书店铺|孙咏美1530521", color: GRAY, size: 18 })],
    alignment: AlignmentType.RIGHT,
  });
}
function h2(text) {
  return new Paragraph({
    children: [new TextRun({ text, color: RED, bold: true, size: 26 })],
    spacing: { before: 280, after: 100 },
  });
}
function p(text, color = DARK) {
  return new Paragraph({
    children: [new TextRun({ text, color, size: 22 })],
    spacing: { before: 60, after: 60 },
    indent: { left: 240 },
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
        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
      },
    },
    children: [
      acct(),
      divider(),

      new Paragraph({
        children: [new TextRun({ text: "2026初级会计6月26日查分攻略，手把手教你", bold: true, size: 40, color: DARK })],
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 200 },
      }),

      divider(),

      h2("📱 查分入口（3个都能用）"),
      p("✅ 官方网站：全国会计资格评价网"),
      p("网址：kj.mof.gov.cn", GRAY),
      p("路径：首页 → 成绩查询 → 输入身份证号 + 准考证号", GRAY),
      blank(),
      p("✅ 手机APP：会计资格评价网官方APP"),
      p("下载后登录，点「成绩查询」即可", GRAY),
      blank(),
      p("✅ 微信公众号：全国会计资格评价网"),
      p("关注后绑定身份信息，成绩出来直接推送通知", GRAY),

      h2("📋 查分需要准备什么"),
      p("→ 身份证号（18位）"),
      p("→ 准考证号（保存好，不要丢）"),
      p("→ 报名时设置的密码（忘了可通过身份证找回）"),

      h2("💡 成绩出来后怎么看"),
      p("初级会计分两科："),
      p("→ 初级会计实务（满分100分）"),
      p("→ 经济法基础（满分100分）"),
      blank(),
      p("两科都达到60分为合格。"),
      blank(),
      p("成绩不合格不要慌！"),
      p("→ 单科合格成绩可以保留1年"),
      p("→ 明年只需补考不合格的那科"),

      h2("✨ 查完成绩之后"),
      p("过了：恭喜！去申领电子证书👇"),
      p("全国会计资格评价网 → 证书查询 → 下载电子证书", GRAY),
      p("（电子证书可直接用于求职，效力与纸质证书相同）", GRAY),
      blank(),
      p("没过也没关系："),
      p("→ 复盘失误在哪里"),
      p("→ 2027年备考还来得及，起步早的人更稳"),

      divider(),

      new Paragraph({
        children: [
          new TextRun({ text: "话题：", color: GRAY, size: 18 }),
          new TextRun({ text: "#初级会计 #初级会计备考 #会计小白 #大学生考证 #查分 #会计 #财会", color: BLUE, size: 18 }),
        ],
        spacing: { before: 100 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "时间：", color: GRAY, size: 18 })],
      }),
      new Paragraph({
        children: [new TextRun({ text: "账号：06红书店铺|孙咏美1530521", color: GRAY, size: 18 })],
        spacing: { after: 200 },
      }),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("03-初级专业号/小红书主号_查分攻略.docx", buf);
  console.log("生成成功：03-初级专业号/小红书主号_查分攻略.docx");
});
