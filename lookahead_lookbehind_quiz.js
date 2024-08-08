const quizData = [
    {
      "question": "次のログファイルから、エラーメッセージを含むが警告メッセージは含まない行のみを抽出するための正規表現パターンはどれですか？\n\nログファイル例：\n2023-08-09 10:15:30 [INFO] アプリケーションが起動しました\n2023-08-09 10:16:45 [WARNING] メモリ使用率が高くなっています\n2023-08-09 10:17:20 [ERROR] データベース接続に失敗しました\n2023-08-09 10:18:10 [INFO] バックアップ処理を開始します\n2023-08-09 10:19:30 [ERROR] ファイルの書き込みに失敗しました [WARNING] 再試行します",
      "choices": [
        "^.*\\[ERROR\\].*$",
        "^.*\\[ERROR\\].*(?!.*\\[WARNING\\]).*$",
        "^(?=.*\\[ERROR\\])(?!.*\\[WARNING\\]).*$"
      ],
      "correct": 2,
      "explanation": "^(?=.*\\[ERROR\\])(?!.*\\[WARNING\\]).*$ が正解です。この正規表現は、行に [ERROR] を含み（肯定先読み）、かつ [WARNING] を含まない（否定先読み）行にマッチします。これにより、エラーメッセージを含むが警告メッセージは含まない行のみを抽出できます。"
    },
    {
      "question": "以下のHTML文書から、class属性に'highlight'を含むが'hidden'は含まない全ての<div>タグを抽出するための正規表現パターンはどれですか？\n\nHTML例：\n<div class='content highlight'>重要な情報</div>\n<div class='sidebar hidden'>非表示のサイドバー</div>\n<div class='footer highlight hidden'>フッター</div>\n<div class='header highlight'>ヘッダー</div>",
      "choices": [
        "<div class=['\"](?=.*\\bhighlight\\b)(?!.*\\bhidden\\b)[^'\"]*['\"]>[^<]*</div>",
        "<div class=['\"].*highlight.*(?<!hidden)['\"]>[^<]*</div>",
        "<div class=['\"](?:(?!hidden).)*highlight(?:(?!hidden).)*['\"]>[^<]*</div>"
      ],
      "correct": 0,
      "explanation": "<div class=['\"](?=.*\\bhighlight\\b)(?!.*\\bhidden\\b)[^'\"]*['\"]>[^<]*</div> が正解です。この正規表現は、class属性内に'highlight'を含み（肯定先読み）、'hidden'を含まない（否定先読み）<div>タグにマッチします。\\b を使用して、'highlight'と'hidden'が単語の一部ではなく独立した単語としてマッチすることを保証しています。"
    },
    {
      "question": "次の文章から、8桁の数字で、最初の2桁が '19' または '20' で始まる西暦年を表す可能性のある数字のみを抽出するための正規表現パターンはどれですか？ただし、月は01-12、日は01-31の範囲内とします。\n\n文章例：\n私は1985年7月20日に生まれ、2023年4月1日に新しい仕事を始めました。電話番号は090-1234-5678で、クレジットカードの有効期限は20251230です。2099年12月31日の予約も入っています。",
      "choices": [
        "(19|20)\\d{6}",
        "(19|20)\\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\\d|3[01])",
        "(?<!\\d)((?:19|20)\\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\\d|3[01]))(?!\\d)"
      ],
      "correct": 2,
      "explanation": "(?<!\\d)((?:19|20)\\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\\d|3[01]))(?!\\d) が正解です。この正規表現は、19または20で始まる4桁の年、01-12の範囲の月、01-31の範囲の日にマッチします。さらに、前後に数字がないことを確認する否定先読みと否定後読みを使用して、電話番号やクレジットカード番号の一部を誤って抽出することを防いでいます。"
    }
  ]
