const path = require('path');
const fs = require('fs');
const D = require(path.join(__dirname, 'node_modules', 'docx'));
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, convertMillimetersToTwip } = D;

const FONT = { ascii: 'ＭＳ 明朝', eastAsia: 'ＭＳ 明朝', hAnsi: 'ＭＳ 明朝' };

// 穴埋め箇所・取捨選択する箇所は黄色ハイライトで目立たせる
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

const plain = (children) => new Paragraph({ children, spacing: { line: 300 } });

const item = (children) =>
  new Paragraph({ children, spacing: { line: 300, after: 60 }, indent: { left: 340 } });

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
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
        plain([t('法務大臣　殿')]),
        blank(),
        rightLine([t('申請人　氏　　名　'), ph('妻の氏名'), t('　（国籍　'), ph('国籍'), t('）')]),
        blank(),
        rightLine([t('本書作成者　申請人の夫（身元保証人）')]),
        rightLine([t('氏　　名　'), ph('氏名'), t('　　（自筆署名）')]),
        rightLine([t('生年月日　'), ph('西暦　年　月　日')]),
        rightLine([t('住　　所　'), ph('住所')]),
        rightLine([t('職　　業　個人事業主（フリーランスエンジニア）')]),
        blank(),

        new Paragraph({
          children: [t('理　由　書', { bold: true, size: 28 })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 240, after: 120 },
          heading: HeadingLevel.HEADING_1,
        }),
        new Paragraph({
          children: [t('（住民税の納付が納期限を過ぎたことについて）', { size: 21 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 360 },
        }),

        body([
          t('　私は、永住許可申請人　'), ph('妻の氏名'),
          t('　の夫であり、同人の身元保証人です。このたびの永住許可申請にあたり、私が納付する住民税について、一部の納期限を過ぎてからの納付となったものがございましたので、その経緯及び現在の納付状況について、以下のとおりご説明申し上げます。'),
        ]),

        heading('１　経緯'),
        body([
          t('　私は個人事業主（フリーランスのシステムエンジニア）として稼働しており、住民税は給与からの特別徴収ではなく、市区町村から送付される納付書による普通徴収の方法で納付しております。'),
        ]),
        body([
          t('　このうち、下記の分につきまして、納期限を過ぎてからの納付となりました。'),
        ]),
        item([
          t('・'), ph('　'), t('年度　第'), ph('　'), t('期　／　納期限　'),
          ph('　'), t('年　'), ph('　'), t('月　'), ph('　'), t('日　／　税額　'),
          ph('　'), t('円'),
        ]),
        item([
          ph('該当が複数ある場合はこの行を複製して追記し、1件のみの場合はこの行を削除してください。'),
        ]),

        heading('２　納期限を過ぎた理由'),
        body([
          t('　当該納期限の前後は、担当していた開発案件の繁忙期と重なっており、加えて客先での作業や出張により自宅を離れる期間が続いておりました。そのため、納付書の納期限を十分に確認できず、また金融機関やコンビニエンスストアの窓口で納付する時間を確保できないまま、納期限を経過してしまいました。'),
        ]),
        body([
          t('　資金が不足していたものではなく、また納付を怠る意図があったものでも決してございません。納付書による納付を自己の予定管理に委ねていたことによる不注意であり、深く反省しております。'),
        ]),

        heading('３　現在の納付状況'),
        body([
          t('　納期限後となりましたが、上記の住民税につきましては、下記のとおり既に全額を納付しております。'),
        ]),
        item([
          t('・'), ph('　'), t('年度　第'), ph('　'), t('期　／　納付日　'),
          ph('　'), t('年　'), ph('　'), t('月　'), ph('　'), t('日　／　納付額　'),
          ph('　'), t('円'),
        ]),
        item([
          ph('該当が複数ある場合はこの行を複製して追記し、1件のみの場合はこの行を削除してください。'),
        ]),
        body([
          ph('延滞金が発生した場合はこの一文を残し、発生していない場合は削除してください：また、延滞金についても併せて納付済みです。'),
        ]),
        body([
          t('　現在、'),
          t('未納となっている住民税はございません。', { bold: true }),
          t('納付済みであることを証する納税証明書及び領収証書の写しを添付いたします。なお、所得税その他の公租公課につきましても、滞納はございません。'),
        ]),

        heading('４　今後について'),
        body([
          t('　今回の件を踏まえ、今後は納付書による納付を改め、'),
          t('住民税を口座振替により納付することといたします。', { bold: true }),
          t('具体的には、'), ph('　'), t('年　'), ph('　'),
          t('月までに市区町村の窓口又は金融機関において口座振替の申込手続を行い、納期限に確実に引き落とされる状態にいたします。繁忙期や出張の有無にかかわらず納付が漏れない仕組みに改めることで、二度とこのようなことがないよう徹底いたします。'),
        ]),
        body([
          t('　以上のとおり、今回の納付の遅れは故意によるものではなく、既に全額を納付済みであり、かつ今後の再発防止措置も講じることをご理解いただき、妻　'),
          ph('妻の氏名'),
          t('　の永住許可を賜りますよう、何卒お願い申し上げます。'),
        ]),

        new Paragraph({
          children: [t('以　上')],
          alignment: AlignmentType.RIGHT,
          spacing: { before: 240, after: 360 },
        }),

        heading('【添付書類】'),
        plain([t('１　住民税　納税証明書（未納がないことが分かるもの）')]),
        plain([t('２　住民税　領収証書（写し）')]),
        plain([t('３　口座振替の申込書控え（申込済みの場合）')]),
      ],
    },
  ],
});

const out = process.argv[2];
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log('wrote', out, buf.length, 'bytes');
});
