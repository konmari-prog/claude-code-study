// ========================================
// Claude Code x GAS アンケート集計 デモ用
// ========================================

/** 複数アンケート設定 */
const SURVEY_CONFIGS = [
  {
    title: 'Claude Code勉強会 第8回アンケート',
    description: '第8回「当日資料を秒で作る」のアンケートです。ご回答ありがとうございます。',
    questions: [
      { title: '全体的な満足度を教えてください', type: 'SCALE', required: true, options: { low: 1, high: 5, lowLabel: '不満', highLabel: '満足' } },
      { title: '理解度を教えてください', type: 'SCALE', required: true, options: { low: 1, high: 5, lowLabel: '理解できなかった', highLabel: 'よく理解できた' } },
      { title: '特に良かった点は？（複数選択可）', type: 'CHECKBOX', required: false, options: ['内容の分かりやすさ', '資料の充実度', '講師の説明', '実践的な内容', 'ハンズオンの時間'] },
      { title: '質問・要望があればお書きください', type: 'PARAGRAPH', required: false, options: null },
      { title: '感想をお聞かせください', type: 'PARAGRAPH', required: false, options: null }
    ]
  },
  {
    title: 'Claude Code勉強会 第9回アンケート',
    description: '第9回「GASに触れてみる」のアンケートです。ご回答ありがとうございます。',
    questions: [
      { title: '全体的な満足度を教えてください', type: 'SCALE', required: true, options: { low: 1, high: 5, lowLabel: '不満', highLabel: '満足' } },
      { title: '理解度を教えてください', type: 'SCALE', required: true, options: { low: 1, high: 5, lowLabel: '理解できなかった', highLabel: 'よく理解できた' } },
      { title: '業務で活用できそうですか？', type: 'MULTIPLE_CHOICE', required: true, options: ['すぐに活用したい', 'いずれ活用したい', 'まだわからない', '活用は難しそう'] },
      { title: '次回取り上げてほしいテーマ（複数選択可）', type: 'CHECKBOX', required: false, options: ['Slack連携', 'メール自動化', 'データ分析', 'カレンダー連携', 'その他'] },
      { title: '質問があればお書きください', type: 'PARAGRAPH', required: false, options: null },
      { title: '感想をお聞かせください', type: 'PARAGRAPH', required: false, options: null }
    ]
  },
  {
    title: 'Claude Code 特別講座アンケート',
    description: '特別講座「AIで業務改善」のアンケートです。ご回答ありがとうございます。',
    questions: [
      { title: '満足度を教えてください', type: 'SCALE', required: true, options: { low: 1, high: 5, lowLabel: '不満', highLabel: '満足' } },
      { title: '講座の難易度はいかがでしたか？', type: 'MULTIPLE_CHOICE', required: true, options: ['簡単すぎた', 'ちょうどよかった', '少し難しかった', '難しすぎた'] },
      { title: '質問・改善点があればお書きください', type: 'PARAGRAPH', required: false, options: null },
      { title: '総合評価', type: 'SCALE', required: true, options: { low: 1, high: 5, lowLabel: '低い', highLabel: '高い' } }
    ]
  },
  {
    title: 'Claude Code勉強会 第10回アンケート',
    description: '第10回「Claude Code x GAS」のアンケートです。ご回答ありがとうございます！',
    questions: [
      { title: '今日の満足度を教えてください', type: 'SCALE', required: true, options: { low: 1, high: 5, lowLabel: '不満', highLabel: '満足' } },
      { title: 'GASの理解度はどうでしたか？', type: 'SCALE', required: true, options: { low: 1, high: 5, lowLabel: '理解できなかった', highLabel: 'よく理解できた' } },
      { title: 'clasp pushは成功しましたか？', type: 'MULTIPLE_CHOICE', required: true, options: ['成功した', 'まだできていない', 'エラーが出た'] },
      { title: '今日一番「おおっ」と思ったことは？', type: 'PARAGRAPH', required: false, options: null },
      { title: '質問・感想があればお書きください', type: 'PARAGRAPH', required: false, options: null }
    ]
  }
];

// ========================================
// メニュー
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('アンケート集計');
  menu.addItem('全フォーム作成', 'createAllForms');
  const sub = ui.createMenu('個別フォーム作成');
  SURVEY_CONFIGS.forEach((cfg, i) => {
    sub.addItem(cfg.title, `createForm${i}`);
  });
  menu.addSubMenu(sub);
  menu.addSeparator();
  menu.addItem('ダッシュボードを開く', 'openDashboard');
  menu.addItem('フォームURLを表示', 'showFormUrls');
  menu.addItem('FAQ質問を同期', 'syncFaqSheet');
  menu.addItem('FAQ模範解答を記入', 'fillFaqAnswers');
  menu.addItem('FAQ自動追加トリガーを設定', 'setupFormTrigger');
  menu.addItem('使い方', 'showHelp');
  menu.addSeparator();
  menu.addItem('デモ用サンプルデータ生成', 'generateSampleData');
  menu.addToUi();
}

// 個別フォーム作成用のラッパー
function createForm0() { createSurveyForm(0); }
function createForm1() { createSurveyForm(1); }
function createForm2() { createSurveyForm(2); }
function createForm3() { createSurveyForm(3); }

function openDashboard() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(
    `<script>window.open("${url}");google.script.host.close();</script>`
  ).setWidth(200).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, 'ダッシュボードを開く');
}

function showFormUrls() {
  const props = PropertiesService.getScriptProperties();
  let msg = '';
  SURVEY_CONFIGS.forEach((cfg, i) => {
    const url = props.getProperty(`formUrl_${i}`) || '(未作成)';
    msg += `${cfg.title}\n${url}\n\n`;
  });
  SpreadsheetApp.getUi().alert('フォームURL一覧', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showHelp() {
  const html = HtmlService.createHtmlOutputFromFile('help')
    .setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '使い方');
}

// ========================================
// フォーム生成
// ========================================

/**
 * 指定インデックスのアンケートフォームを作成
 * @param {number} index SURVEY_CONFIGSのインデックス
 * @return {string} フォームURL
 */
function createSurveyForm(index) {
  const cfg = SURVEY_CONFIGS[index];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = FormApp.create(cfg.title);

  form.setDescription(cfg.description);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false);

  cfg.questions.forEach(q => {
    let item;
    switch (q.type) {
      case 'MULTIPLE_CHOICE':
        item = form.addMultipleChoiceItem();
        item.setTitle(q.title).setRequired(q.required);
        item.setChoiceValues(q.options);
        break;
      case 'CHECKBOX':
        item = form.addCheckboxItem();
        item.setTitle(q.title).setRequired(q.required);
        item.setChoiceValues(q.options);
        break;
      case 'SCALE':
        item = form.addScaleItem();
        item.setTitle(q.title).setRequired(q.required);
        item.setBounds(q.options.low, q.options.high);
        item.setLabels(q.options.lowLabel, q.options.highLabel);
        break;
      case 'LIST':
        item = form.addListItem();
        item.setTitle(q.title).setRequired(q.required);
        item.setChoiceValues(q.options);
        break;
      case 'TEXT':
        item = form.addTextItem();
        item.setTitle(q.title).setRequired(q.required);
        break;
      case 'PARAGRAPH':
        item = form.addParagraphTextItem();
        item.setTitle(q.title).setRequired(q.required);
        break;
    }
  });

  const formUrl = form.getPublishedUrl();
  PropertiesService.getScriptProperties().setProperty(`formUrl_${index}`, formUrl);

  SpreadsheetApp.getUi().alert(`${cfg.title} を作成しました\n\n${formUrl}`);
  return formUrl;
}

/** 全フォームを一括作成 */
function createAllForms() {
  SURVEY_CONFIGS.forEach((_, i) => createSurveyForm(i));
  SpreadsheetApp.getUi().alert(`${SURVEY_CONFIGS.length}個のフォームを作成しました`);
}

// ========================================
// WEBアプリ
// ========================================

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'dashboard';
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('アンケート集計ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

// ========================================
// 集計API: Dashboard
// ========================================

function getSurveyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().startsWith('フォームの回答'));

  const surveys = [];
  let totalResponses = 0;

  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0];
    const rows = data.slice(1);
    totalResponses += rows.length;
    const surveyName = sheet.getName().replace('フォームの回答 ', 'アンケート');

    const questions = [];
    headers.forEach((header, colIdx) => {
      if (colIdx === 0) return; // タイムスタンプ列をスキップ
      const values = rows.map(r => r[colIdx]).filter(v => v !== '' && v !== null && v !== undefined);
      if (values.length === 0) return;

      const isScale = values.every(v => typeof v === 'number' || (!isNaN(Number(v)) && String(v).match(/^[1-5]$/)));

      if (isScale) {
        const nums = values.map(Number);
        const avg = nums.reduce((a, b) => a + b, 0) / nums.length;
        const dist = {};
        for (let i = 1; i <= 5; i++) dist[i] = 0;
        nums.forEach(n => { if (dist[n] !== undefined) dist[n]++; });
        questions.push({ title: header, type: 'scale', avg: Math.round(avg * 10) / 10, distribution: dist, count: nums.length });
      } else if (values.some(v => String(v).includes(', '))) {
        const counts = {};
        values.forEach(v => {
          String(v).split(', ').forEach(opt => {
            counts[opt] = (counts[opt] || 0) + 1;
          });
        });
        questions.push({ title: header, type: 'checkbox', counts, total: values.length });
      } else if (values.every(v => String(v).length < 50) && new Set(values.map(String)).size < values.length * 0.8) {
        const counts = {};
        values.forEach(v => { counts[String(v)] = (counts[String(v)] || 0) + 1; });
        questions.push({ title: header, type: 'choice', counts, total: values.length });
      } else {
        questions.push({ title: header, type: 'text', values: values.slice(0, 10).map(String) });
      }
    });

    surveys.push({ name: surveyName, responseCount: rows.length, questions });
  });

  return { totalResponses, surveyCount: surveys.length, surveys, lastUpdate: new Date().toLocaleString('ja-JP') };
}

// ========================================
// 集計API: Hot Words
// ========================================

function getHotWords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().startsWith('フォームの回答'));

  const STOP_WORDS = new Set(['の', 'は', 'が', 'を', 'に', 'で', 'と', 'も', 'か', 'な', 'い', 'る', 'た', 'する', 'ある', 'いる', 'れる', 'です', 'ます', 'こと', 'もの', 'ため', 'よう', 'それ', 'これ', 'この', 'その', 'から', 'まで', 'だけ', 'など', 'して', 'ない', 'なる', 'れた', 'った', 'ので', 'けど', 'ても', 'という', 'について', 'として', 'られ', 'できる', 'ている', 'てい', 'ました', 'でした', 'ありがとう', 'ございました', 'ございます', 'おり', 'られる', 'ところ', 'くださ', 'いただ', 'ですが', 'ですね', 'ました', 'かった', 'ません', 'ない', 'あり', 'なかった', 'ました']);
  const NEGATIVE_WORDS = [
    // 不満・クレーム
    '不満', 'クレーム', '苦情', '訴え', '許せない', '信じられない', 'ありえない', 'ふざけ',
    // 怒り・強い感情
    '怒', '最悪', 'ひどい', '酷い', '失望', 'がっかり', '残念', '呆れ', '腹が立つ', 'イライラ',
    // 品質の低さ
    'わかりにくい', 'わかりづらい', '難しい', 'つまらない', '退屈', '微妙', '中途半端', '物足りない', '期待はずれ',
    // 量・時間の問題
    '長い', '長すぎ', '短い', '短すぎ', '遅い', '遅すぎ', '少ない', '足りない', '多すぎ',
    // 要改善
    '改善', '見直し', 'やめて', 'やめた方', '無駄', '意味がない', '必要ない', '不要',
    // 金銭・信頼
    '返金', '詐欺', '高い', '高すぎ', 'コスパ', '割に合わない', '損', '騙',
    // 対応の悪さ
    '対応が悪い', '無視', '放置', '雑', '適当', '不親切', '冷たい',
    // 体調・負荷
    '疲れ', 'しんどい', 'きつい', 'つらい', '辛い', 'ストレス', '負担'
  ];

  const wordCounts = {};
  const negativeHits = {};
  let totalTexts = 0;

  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0];
    const rows = data.slice(1);
    const surveyName = sheet.getName().replace('フォームの回答 ', 'アンケート');

    headers.forEach((header, colIdx) => {
      if (colIdx === 0) return;
      rows.forEach(row => {
        const val = String(row[colIdx]);
        if (val.length < 5) return;

        totalTexts++;
        const tokens = val.split(/[\s、。！？!?.,;:()（）「」『』\n\r\t]+/).filter(t => t.length >= 2 && !STOP_WORDS.has(t));
        tokens.forEach(token => {
          wordCounts[token] = (wordCounts[token] || 0) + 1;
        });

        NEGATIVE_WORDS.forEach(nw => {
          if (val.includes(nw)) {
            if (!negativeHits[nw]) negativeHits[nw] = [];
            negativeHits[nw].push({ text: val, source: surveyName });
          }
        });
      });
    });
  });

  const topWords = Object.entries(wordCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 30)
    .map(([word, count]) => ({ word, count }));

  const negatives = Object.entries(negativeHits)
    .sort((a, b) => b[1].length - a[1].length)
    .map(([word, hits]) => ({ word, count: hits.length, hits }));

  return { topWords, negatives, totalTexts };
}

// ========================================
// 集計API: Notable Comments
// ========================================

function getNotableComments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().startsWith('フォームの回答'));

  const KEYWORDS = ['質問', '要望', '希望', '提案', 'お願い', 'ほしい', 'してほしい', '教えて'];
  const comments = [];

  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0];
    const rows = data.slice(1);
    const surveyName = sheet.getName().replace('フォームの回答 ', 'アンケート');

    rows.forEach(row => {
      const scaleValues = [];
      headers.forEach((h, ci) => {
        if (ci === 0) return;
        const v = row[ci];
        if (typeof v === 'number' && v >= 1 && v <= 5) scaleValues.push(v);
      });
      const hasLowScore = scaleValues.some(v => v <= 2);

      headers.forEach((header, colIdx) => {
        if (colIdx === 0) return;
        const val = String(row[colIdx]);
        if (val.length < 5) return;

        const types = [];
        if (val.length >= 50) types.push('long');
        if (hasLowScore) types.push('low-score');
        if (KEYWORDS.some(kw => val.includes(kw))) types.push('keyword');

        if (types.length > 0) {
          comments.push({
            text: val,
            source: surveyName,
            question: header,
            types,
            timestamp: row[0] ? new Date(row[0]).toLocaleDateString('ja-JP') : ''
          });
        }
      });
    });
  });

  comments.sort((a, b) => b.types.length - a.types.length);
  return { comments: comments.slice(0, 30), totalFound: comments.length };
}

// ========================================
// 集計API: FAQ
// ========================================

function getFaqData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // FAQシートから回答済みのものを取得
  const faqSheet = ss.getSheetByName('FAQ');
  if (faqSheet) {
    const faqData = faqSheet.getDataRange().getValues();
    if (faqData.length >= 2) {
      const answeredFaqs = [];
      const unansweredCount = { total: 0 };
      for (let i = 1; i < faqData.length; i++) {
        const question = String(faqData[i][0]).trim();
        const answer = String(faqData[i][1]).trim();
        const source = String(faqData[i][2] || '').trim();
        const count = Number(faqData[i][3]) || 1;
        if (question.length < 2) continue;
        if (answer.length >= 2) {
          answeredFaqs.push({ question, answer, source, count });
        } else {
          unansweredCount.total++;
        }
      }
      return { faqs: answeredFaqs, totalQuestions: answeredFaqs.length + unansweredCount.total, unanswered: unansweredCount.total, fromSheet: true };
    }
  }

  // FAQシートがない場合: フォーム回答から質問を収集してFAQシートに書き出す
  return syncFaqSheet();
}

/**
 * フォーム回答から質問を収集し、FAQシートに書き出す
 */
function syncFaqSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().startsWith('フォームの回答'));

  const rawQuestions = [];

  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0];
    const rows = data.slice(1);
    const surveyName = sheet.getName().replace('フォームの回答 ', 'アンケート');

    headers.forEach((header, colIdx) => {
      if (colIdx === 0) return;
      if (!header.includes('質問')) return;

      rows.forEach(row => {
        const val = String(row[colIdx]).trim();
        if (val.length >= 5) {
          rawQuestions.push({ text: val, source: surveyName });
        }
      });
    });
  });

  // 類似質問グルーピング
  const groups = [];
  rawQuestions.forEach(q => {
    const existing = groups.find(g =>
      g.questions[0].text === q.text ||
      (g.questions[0].text.length > 10 && q.text.includes(g.questions[0].text.substring(0, 10)))
    );
    if (existing) {
      existing.questions.push(q);
    } else {
      groups.push({ questions: [q] });
    }
  });

  groups.sort((a, b) => b.questions.length - a.questions.length);

  // FAQシートに書き出す
  let faqSheet = ss.getSheetByName('FAQ');
  if (!faqSheet) {
    faqSheet = ss.insertSheet('FAQ');
  } else {
    faqSheet.clear();
  }

  // ヘッダー
  faqSheet.getRange(1, 1, 1, 4).setValues([['質問', '回答', '出典', '件数']]);
  faqSheet.getRange(1, 1, 1, 4).setFontWeight('bold');

  // データ
  const faqRows = groups.slice(0, 30).map(g => [
    g.questions[0].text,
    '',  // 回答は空欄（手動で記入する）
    [...new Set(g.questions.map(q => q.source))].join(', '),
    g.questions.length
  ]);

  if (faqRows.length > 0) {
    faqSheet.getRange(2, 1, faqRows.length, 4).setValues(faqRows);
  }

  // 列幅調整
  faqSheet.setColumnWidth(1, 400);
  faqSheet.setColumnWidth(2, 400);
  faqSheet.setColumnWidth(3, 150);
  faqSheet.setColumnWidth(4, 60);

  return { faqs: [], totalQuestions: rawQuestions.length, unanswered: rawQuestions.length, fromSheet: false };
}

/**
 * フォーム送信時にFAQシートへ新しい質問を自動追加する
 */
function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let faqSheet = ss.getSheetByName('FAQ');
  if (!faqSheet) return;

  const sheet = e.range.getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = e.range.getRow();
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const surveyName = sheet.getName().replace('フォームの回答 ', 'アンケート');

  // 「質問」を含むヘッダーの列から値を取得
  const faqData = faqSheet.getDataRange().getValues();
  const existingQuestions = faqData.slice(1).map(r => String(r[0]).trim());

  headers.forEach((header, colIdx) => {
    if (colIdx === 0) return;
    if (!String(header).includes('質問')) return;

    const val = String(rowData[colIdx]).trim();
    if (val.length < 5) return;
    if (existingQuestions.includes(val)) return;

    // 新しい質問をFAQシートに追加
    const nextRow = faqSheet.getLastRow() + 1;
    faqSheet.getRange(nextRow, 1, 1, 4).setValues([[val, '', surveyName, 1]]);
  });
}

/**
 * フォーム送信トリガーを設定する
 */
function setupFormTrigger() {
  // 既存のonFormSubmitトリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 新しいトリガーを作成
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  SpreadsheetApp.getUi().alert('フォーム送信トリガーを設定しました。\n新しい質問がフォームに入ると、FAQシートに自動追加されます。');
}

/**
 * FAQシートにデモ用の模範解答を記入する
 */
function fillFaqAnswers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let faqSheet = ss.getSheetByName('FAQ');
  if (!faqSheet) {
    syncFaqSheet();
    faqSheet = ss.getSheetByName('FAQ');
  }
  if (!faqSheet) return;

  const data = faqSheet.getDataRange().getValues();
  const answers = {
    'セキュリティ的に社内データを渡しても大丈夫ですか？': 'Claude Codeはローカルで動作し、入力データはモデルの学習には使用されません。ただし、機密情報を扱う場合は社内のセキュリティポリシーを確認してください。API経由の通信は暗号化されています。',
    'GASの実行時間制限はどう回避できますか？': 'GASの実行時間は1回あたり6分が上限です。大量データを処理する場合は、処理を分割してトリガーで順次実行する方法が有効です。途中経過をスクリプトプロパティに保存し、次の実行で続きから処理します。',
    '他のAIツールとの違いを教えてほしいです。': 'Claude Codeはターミナル上で動作し、ファイルの読み書き・コマンド実行を直接行えるのが特徴です。ChatGPTやGeminiはブラウザ上での対話が中心ですが、Claude Codeは実際のプロジェクトファイルを操作できます。',
    'Slack連携の具体的な手順を知りたいです。': 'GASからSlackに通知を送るには、SlackのIncoming Webhook URLを取得し、UrlFetchApp.fetchでPOSTリクエストを送ります。Claude Codeに「Slack通知するGAS作って」と伝えれば、コードを自動生成してくれます。',
    'スキルを共有する方法を教えてください。': 'スキルフォルダ（SKILL.md + references/）をzipに圧縮して共有します。受け取った人はClaude Codeにzipファイルのパスを渡せば自動でインストールされます。スキルギャラリーにアップロードすることも可能です。',
    'Claude Codeのライセンス費用はどのくらいですか？': 'Claude CodeはClaude Pro（月額$20）またはClaude Max（月額$100/$200）のサブスクリプションで利用できます。Proでも基本機能は使えますが、Maxだとより多くのメッセージを送れます。',
    '大量データの処理はどうすればいいですか？': 'GASでは一括読み込み（getValues）を使い、ループ内でのgetValueは避けてください。数万行を超えるデータは、処理を分割して複数回に分けるか、BigQueryとの連携を検討してください。',
    'CLAUDE.mdのベストプラクティスはありますか？': '自分の役割・プロジェクトの目的・よく使うルールを書きます。「やらないこと」も明記すると効果的です。長くなりすぎず、AIが迷いそうなポイントに絞るのがコツです。'
  };

  for (let i = 1; i < data.length; i++) {
    const question = String(data[i][0]).trim();
    if (answers[question]) {
      faqSheet.getRange(i + 1, 2).setValue(answers[question]);
    }
  }

  SpreadsheetApp.getUi().alert('FAQ模範解答を記入しました');
}

// ========================================
// デモ用: サンプルデータ生成
// ========================================

/**
 * デモ用にサンプル回答データを生成する
 * 各アンケートに10〜15件のダミー回答を作成
 */
function generateSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sampleComments = [
    'Claude Codeの使い方がよくわかりました。実務でもすぐに使えそうです。',
    'ハンズオンの時間がもう少し欲しかったです。',
    'スキルの作り方が特に参考になりました。自分でも作ってみたいと思います。',
    'GASとの連携はとても面白いですね。業務効率化に活用したいです。',
    '説明がわかりやすくて助かりました。',
    '少し難しかったですが、質問しやすい雰囲気だったので良かったです。',
    'もっと応用的な内容も知りたいです。',
    'セキュリティの話が印象に残りました。気をつけないといけないですね。',
    '実際に動くものが作れて達成感がありました。',
    'チームメンバーにも教えたいと思います。',
    '資料がとても充実していて復習しやすいです。',
    'AI活用の具体的なイメージが湧きました。提案してほしいテーマがあります。',
    '改善点としては、ペースがやや速かったかもしれません。',
    'データ分析の自動化について質問があります。どこまでできますか？',
    'Slack連携のやり方も教えてほしいです。'
  ];

  const sampleQuestions = [
    'Claude Codeのライセンス費用はどのくらいですか？',
    'GASの実行時間制限はどう回避できますか？',
    'スキルを共有する方法を教えてください。',
    'セキュリティ的に社内データを渡しても大丈夫ですか？',
    'Slack連携の具体的な手順を知りたいです。',
    'CLAUDE.mdのベストプラクティスはありますか？',
    '他のAIツールとの違いを教えてほしいです。',
    '大量データの処理はどうすればいいですか？',
    ''
  ];

  SURVEY_CONFIGS.forEach((cfg, cfgIndex) => {
    const sheetName = `フォームの回答 ${cfgIndex + 1}`;
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }

    // ヘッダー
    const headers = ['タイムスタンプ', ...cfg.questions.map(q => q.title)];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // サンプルデータ生成
    const numRows = 10 + Math.floor(Math.random() * 6);
    const rows = [];

    for (let i = 0; i < numRows; i++) {
      const row = [new Date(2026, 2, 20 + Math.floor(Math.random() * 7), 17 + Math.floor(Math.random() * 3), Math.floor(Math.random() * 60))];

      cfg.questions.forEach(q => {
        switch (q.type) {
          case 'SCALE':
            // 4と5が多めの分布
            const weights = [0.05, 0.1, 0.15, 0.35, 0.35];
            const rand = Math.random();
            let cumulative = 0;
            let score = 5;
            for (let s = 0; s < weights.length; s++) {
              cumulative += weights[s];
              if (rand < cumulative) { score = s + 1; break; }
            }
            row.push(score);
            break;
          case 'MULTIPLE_CHOICE':
            row.push(q.options[Math.floor(Math.random() * q.options.length)]);
            break;
          case 'CHECKBOX':
            const numSelected = 1 + Math.floor(Math.random() * 3);
            const shuffled = [...q.options].sort(() => Math.random() - 0.5);
            row.push(shuffled.slice(0, numSelected).join(', '));
            break;
          case 'PARAGRAPH':
            if (q.title.includes('質問')) {
              row.push(sampleQuestions[Math.floor(Math.random() * sampleQuestions.length)]);
            } else {
              row.push(sampleComments[Math.floor(Math.random() * sampleComments.length)]);
            }
            break;
          default:
            row.push('');
        }
      });

      rows.push(row);
    }

    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  });

  SpreadsheetApp.getUi().alert(`デモ用サンプルデータを生成しました（${SURVEY_CONFIGS.length}アンケート分）`);
}
