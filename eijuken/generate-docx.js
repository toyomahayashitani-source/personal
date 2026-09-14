const path = require('path');
const fs = require('fs');
const D = require(path.join(__dirname, 'node_modules', 'docx'));
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, convertMillimetersToTwip } = D;

const FONT = { ascii: 'ＭＳ 明朝', eastAsia: 'ＭＳ 明朝', hAnsi: 'ＭＳ 明朝' };

// 穴埋め箇所は黄色ハイライトで目立たせる
const ph = (label) => new TextRun({ text: `【${label}】`, highlight: 'yellow' });
const t = (text, opts = {}) => new TextRun({ text, ...opts });

const body = (children, opts = {}) =>
  new Paragraph({ children, spacing: { line: 360, after: 120 }, ...opts });

const blank = () => new Paragraph({ children: [], spacing: { line: 240 } });

const heading = (text) =>
  new Paragraph({
    children: [t(text, { bold: true })],
    spacing: { before: 240, after: 80, line: 300 },
  });

const rightLine = (children) =>
  new Paragraph({ children, alignment: AlignmentType.RIGHT, spacing: { line: 300 } });

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONT, size: 21 } },
    },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: convertMillimetersToTwip(210), height: convertMillimetersToTwip(297) },
          margin: {
            top: convertMillimetersToTwip(25),
            bottom: convertMillimetersToTwip(25),
            left: convertMillimetersToTwip(25),
            right: convertMillimetersToTwip(25),
          },
        },
      },
      children: [
        rightLine([t('2026年　'), ph('　'), t('月　'), ph('　'), t('日')]),
        blank(),
        new Paragraph({ children: [t('法務大臣　殿')], spacing: { line: 300 } }),
        blank(),
        rightLine([t('申請人　氏　　名　'), ph('氏名'), t('　　（自筆署名）')]),
        rightLine([t('国　　籍　'), ph('国籍')]),
        rightLine([t('生年月日　'), ph('西暦　年　月　日')]),
        rightLine([t('住　　所　'), ph('住所')]),
        rightLine([t('在留カード番号　'), ph('番号')]),
        rightLine([t('基礎年金番号　'), ph('番号')]),
        blank(),

        new Paragraph({
          children: [t('理　由　書', { bold: true, size: 28 })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 240, after: 120 },
          heading: HeadingLevel.HEADING_1,
        }),
        new Paragraph({
          children: [t('（国民年金保険料の納付が遅れたことについて）', { size: 21 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 360 },
        }),

        body([
          t('　このたびは永住許可申請にあたり、国民年金保険料の納付に遅れが生じましたので、その経緯及び現在の納付状況について、以下のとおりご説明申し上げます。'),
        ]),

        heading('１　経緯'),
        body([
          t('　私は、2026年2月'), ph('　'), t('日をもって'), ph('前職の会社名'),
          t('を退職し、同年4月'), ph('　'), t('日付で'), ph('現職の会社名'),
          t('に入社いたしました。そのため、2026年3月の1か月間のみ厚生年金保険の被保険者資格を喪失し、国民年金第1号被保険者となりました。この1か月分の国民年金保険料について、納期限までに納付することができませんでした。'),
        ]),

        heading('２　納付が遅れた理由'),
        body([
          t('　転職の前後において、健康保険・年金の切替手続や各種届出が重なっており、2026年3月分の国民年金保険料を自分自身で納付しなければならないことを十分に認識しておりませんでした。また、納付書が手元に届いた時期と各種手続の時期が前後したこともあり、納期限までに納付するに至りませんでした。'),
        ]),
        body([
          t('　納付を意図的に免れようとしたものでは決してなく、制度に対する理解不足と手続上の見落としによるものです。深く反省しております。'),
        ]),

        heading('３　現在の納付状況'),
        body([
          t('　その後、未納となっていることに気づき、2026年'), ph('　'), t('月'), ph('　'),
          t('日に、2026年3月分の国民年金保険料'), ph('金額'),
          t('円を全額納付いたしました。納付済みであることを証する領収証書の写しを添付いたします。'),
        ]),
        body([
          t('　現在は'), ph('現職の会社名'),
          t('において厚生年金保険に加入しており、保険料は給与から適正に控除されております。上記3月分を納付したことにより、'),
          t('現時点で未納となっている国民年金保険料はございません。', { bold: true }),
          t('また、住民税その他の公租公課につきましても、これまで滞納なく納付しております。'),
        ]),

        heading('４　今後について'),
        body([
          t('　今回の件を真摯に反省し、今後は公的義務を確実に履行してまいります。具体的には、転職等により加入状況が変わる際には、年金事務所又は市区町村の窓口で必要な手続を必ず確認するとともに、「ねんきんネット」等により納付状況を定期的に確認し、二度とこのようなことがないよう十分に注意いたします。'),
        ]),
        body([
          t('　以上のとおり、今回の納付の遅れは故意によるものではなく、既に全額を納付済みであることをご理解いただき、何卒永住許可を賜りますようお願い申し上げます。'),
        ]),

        new Paragraph({
          children: [t('以　上')],
          alignment: AlignmentType.RIGHT,
          spacing: { before: 240, after: 360 },
        }),

        heading('【添付書類】'),
        new Paragraph({ children: [t('１　国民年金保険料　領収証書（写し）')], spacing: { line: 300 } }),
        new Paragraph({ children: [t('２　退職日・入社日が分かる書類（退職証明書、健康保険の資格取得を証する書類等）')], spacing: { line: 300 } }),
        new Paragraph({ children: [t('３　被保険者記録照会回答票又は「ねんきんネット」の納付状況画面（直近2年分）')], spacing: { line: 300 } }),
      ],
    },
  ],
});

const out = process.argv[2];
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log('wrote', out, buf.length, 'bytes');
});
