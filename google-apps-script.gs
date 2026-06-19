/**
 * LeaPASS 英語力診断ツール - 結果記録用 Apps Script
 *
 * セットアップ手順:
 * 1. https://sheet.new で新しいGoogleスプレッドシートを作成
 * 2. メニュー「拡張機能」→「Apps Script」を開く
 * 3. デフォルトのコードを全て削除し、このファイルの内容を貼り付け
 * 4. 「デプロイ」→「新しいデプロイ」
 *    - 種類: ウェブアプリ
 *    - 実行ユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. 発行されたウェブアプリのURLをコピー
 * 6. index.html の GAS_URL にそのURLを設定
 */

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.postData.contents);

  if (sheet.getLastRow() === 0) {
    const headers = [
      "タイムスタンプ",
      "LINE表示名", "性別", "年齢", "職業", "学習時間", "学習目的", "TOEIC受験目的", "悩み", "理想の未来", "TOEICへの興味理由", "将来の生活場所", "希望の仕事",
      "推定TOEICスコア", "目標スコア",
      "キャラクタータイプ", "キャラクタータイプの説明",
      "おすすめの勉強方法", "おすすめ教材", "目標達成までの勉強期間", "LeaPASSで補えるところ",
      "正答数(/20)"
    ];
    for (let i = 1; i <= 20; i++) headers.push("Q" + i);
    sheet.appendRow(headers);
  }

  const row = [
    new Date(),
    ...data.profile,
    data.score,
    data.target,
    data.character,
    data.characterDesc,
    data.studyMethods,
    data.materials,
    data.period,
    data.leapassSupport,
    data.correctCount,
  ];

  data.quizResults.forEach(correct => row.push(correct ? "○" : "×"));

  sheet.appendRow(row);

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}
