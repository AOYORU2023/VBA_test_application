const quizData = [
    // 1 マクロとVBAの概念-2
    // 1-1 用語と概念
    {
        question: "VBAとは何の略称ですか？",
        choices: ["Visual Basic Algorithm", "Visual Basic for Applications", "Visual Basic nalysis"],
        correct: 1,
        explanation: "VBAは「Visual Basic for Applications」の略称です。これはMicrosoft Office製品に組み込まれたプログラミング言語です。"
    },
    {
        question: "VBEとは何の略称ですか？",
        choices: ["Visual Basic Editor", "Visual Basic Engine", "Visual Basic Execution"],
        correct: 0,
        explanation: "VBEは「Visual Basic Editor」の略称です。これはVBAコードを書いたり編集したりするための統合開発環境です。"
    },
  // 1-2 ブックとマクロの関係
    {
        question: "マクロを保存する際、どのような選択肢がありますか？",
        choices: ["現在のブックのみ", "新しいブック", "個人用マクロブック"],
        correct: 2,
        explanation: "マクロは現在のブック、新しいブック、または個人用マクロブックに保存できますが、最も汎用的なのは個人用マクロブックです。これにより、すべてのExcelファイルでマクロを使用できます。"
    },
    // 1-3 マクロとセキュリティ
    {
        question: "Excelのマクロセキュリティレベルで、最も安全なのはどれですか？",
        choices: ["すべてのマクロを有効にする", "デジタル署名付きマクロのみ有効にする", "すべてのマクロを無効にする"],
        correct: 2,
        explanation: "最も安全なセキュリティレベルは「すべてのマクロを無効にする」です。ただし、これにより正当なマクロも実行できなくなるため、通常は「デジタル署名付きマクロのみ有効にする」が推奨されます。"
    }
];
