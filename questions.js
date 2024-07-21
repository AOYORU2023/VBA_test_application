const quizData = [
    // 1 マクロとVBAの概念
    // 1-1 用語と概念
    {
        question: "VBAとは何の略称ですか？",
        choices: ["Visual Basic Algorithm", "Visual Basic for Applications", "Visual Basic Analysis"],
        correct: 1,
        explanation: "VBAは「Visual Basic for Applications」の略称です。これはMicrosoft Office製品に組み込まれたプログラミング言語です。"
    },
    {
        question: "マクロ記録とは何ですか？",
        choices: ["手動で書いたコードを実行すること", "Excelの操作を自動的に記録する機能", "VBAコードをデバッグする機能"],
        correct: 1,
        explanation: "マクロ記録は、ユーザーが行った操作をVBAコードとして自動的に記録する機能です。これにより、繰り返し行う操作を自動化できます。"
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
    },
    // 2 マクロ記録
    // 2-1 マクロ記録とは
    {
        question: "マクロ記録の主な利点は何ですか？",
        choices: ["コードを手動で書く必要がない", "記録されたマクロは常に最適化されている", "記録されたマクロは編集できない"],
        correct: 0,
        explanation: "マクロ記録の主な利点は、VBAコードを手動で書く必要がないことです。ユーザーのアクションが自動的にVBAコードに変換されます。"
    },
    {
        question: "マクロの記録を開始するにはどのメニューを使用しますか？",
        choices: ["ファイル", "挿入", "開発"],
        correct: 2,
        explanation: "マクロの記録は「開発」タブから開始します。"
    },
    {
        question: "マクロ記録の停止はどのように行いますか？",
        choices: ["Excelを終了する", "「マクロの記録」をもう一度クリックする", "記録したコードを削除する"],
        correct: 1,
        explanation: "「マクロの記録」をもう一度クリックすると、マクロ記録が停止します。"
    },
    // 2-2 [マクロの記録] ダイアログボックスの設定項目
    {
        question: "[マクロの記録]ダイアログボックスで設定できる項目は次のうちどれですか？",
        choices: ["マクロ名", "マクロの実行速度", "マクロの対象範囲"],
        correct: 0,
        explanation: "[マクロの記録]ダイアログボックスでは、マクロ名、ショートカットキー、マクロの保存先、説明を設定できます。"
    },
    {
        question: "[マクロの記録]ダイアログボックスで設定できない項目はどれですか？",
        choices: ["ショートカットキー", "マクロの説明", "マクロのデバッグオプション"],
        correct: 2,
        explanation: "[マクロの記録]ダイアログボックスでは、マクロ名、ショートカットキー、マクロの保存先、説明を設定できます。デバッグオプションはこのダイアログでは設定できません。"
    },
    {
        question: "[マクロの記録]ダイアログボックスでショートカットキーを設定する際、どのキーと組み合わせて使用しますか？",
        choices: ["Alt", "Shift", "Ctrl"],
        correct: 2,
        explanation: "[マクロの記録]ダイアログボックスでショートカットキーを設定する際は、Ctrlキーとの組み合わせを使用します。"
    },
    // 2-3 個人用マクロブック
    {
        question: "個人用マクロブックの特徴として正しいのはどれですか？",
        choices: ["特定のExcelファイルでのみ使用できる", "すべてのExcelファイルで使用できる", "インターネット上で自動的に共有される"],
        correct: 1,
        explanation: "個人用マクロブックに保存されたマクロは、すべてのExcelファイルで使用することができます。これにより、頻繁に使用するマクロを簡単に再利用できます。"
    },
    {
        question: "マクロを個人用マクロブックに保存する理由は何ですか？",
        choices: ["コードの読みやすさを向上させるため", "他のブックでも使用できるようにするため", "VBAエディタを使用するため"],
        correct: 1,
        explanation: "個人用マクロブックに保存すると、どのブックからでもそのマクロを使用することができます。"
    },
    // 3 モジュールとプロシージャ
    // 3-1 モジュールとは
    {
        question: "VBAにおけるモジュールの役割は何ですか？",
        choices: ["データを保存する", "プロシージャをグループ化する", "ワークシートを管理する"],
        correct: 1,
        explanation: "モジュールは関連するプロシージャ（SubやFunction）をグループ化するためのコンテナです。これにより、コードを整理し、管理しやすくなります。"
    },
    {
        question: "モジュールを挿入するにはどのメニューを使用しますか？",
        choices: ["ファイル", "挿入", "ツール"],
        correct: 1,
        explanation: "モジュールは「挿入」メニューから挿入できます。"
    },
    // 3-2 プロシージャとは
    {
        question: "VBAには主に2種類のプロシージャがありますが、それらは何ですか？",
        choices: ["Open プロシージャと Close プロシージャ", "Sub プロシージャと Function プロシージャ", "Public プロシージャと Private プロシージャ"],
        correct: 1,
        explanation: "VBAの主な2種類のプロシージャは、Sub プロシージャと Function プロシージャです。Sub は値を返さない処理を、Function は値を返す処理を実行します。"
    },
    {
        question: "プロシージャとは何ですか？",
        choices: ["変数の一種", "関数やサブルーチンのこと", "Excelシートの別名"],
        correct: 1,
        explanation: "プロシージャとは、関数やサブルーチンのことで、一連のVBAコードをまとめたものです。"
    },
    {
        question: "プロシージャを呼び出すにはどのようにしますか？",
        choices: ["プロシージャ名を直接記述する", "セルに入力する", "シートを切り替える"],
        correct: 0,
        explanation: "プロシージャ名を直接記述することで、他のプロシージャを呼び出すことができます。"
    },
    // 3-2 プロシージャとは（追加問題）
    {
        question: "別のプロシージャを呼び出すための正しい方法は？",
        choices: ["Call MyProcedure", "Run MyProcedure", "Execute MyProcedure"],
        correct: 0,
        explanation: "別のプロシージャを呼び出すには、'Call' キーワードを使用するか、単にプロシージャ名を記述します。例：Call MyProcedure または単に MyProcedure"
    },
    {
        question: "VBAでコメントを書く際に使用する記号は？",
        choices: ["//", "/* */", "'"],
        correct: 2,
        explanation: "VBAでは、シングルクォーテーション(')を使用してコメントを書きます。行の先頭に'を付けると、その行全体がコメントとして扱われます。"
    },
    {
        question: "VBAで1行の途中で改行する場合、どの記号を使用しますか？",
        choices: ["_", "\\", "/"],
        correct: 0,
        explanation: "VBAで1行の途中で改行する場合は、アンダースコア(_)を使用します。これにより、コードの可読性を向上させることができます。"
    },
    // 4 VBAの構文
    // 4-1 オブジェクト式
    {
        question: "次のうち、正しいオブジェクト式はどれですか？",
        choices: ["Workbooks(1).Worksheets(1).Range(A1)", "Workbooks(1).Worksheets(1).Range(\"A1\")", "Workbooks[1].Worksheets[1].Range[\"A1\"]"],
        correct: 1,
        explanation: "正しいオブジェクト式は Workbooks(1).Worksheets(1).Range(\"A1\") です。VBAではオブジェクトの階層構造をピリオドで連結し、文字列は二重引用符で囲みます。"
    },
    {
        question: "オブジェクト式とは何ですか？",
        choices: ["Excelのオブジェクトを操作するための文法", "プログラムの終了を示す記号", "変数を初期化するための文法"],
        correct: 0,
        explanation: "オブジェクト式とは、Excelのオブジェクト（例えば、セルやシート）を操作するための文法です。"
    },
    {
        question: "オブジェクトの階層構造とは何ですか？",
        choices: ["オブジェクトが親子関係を持つ構造", "オブジェクトが直列に並んでいる構造", "オブジェクトが単一レベルで存在する構造"],
        correct: 0,
        explanation: "オブジェクトの階層構造とは、オブジェクトが親子関係を持ち、階層的に組織されている構造です。"
    },
    {
        question: "セルの表し方で正しいのはどれですか？",
        choices: ["Cell(1,1)", "Cells(1,1)", "Range(1,1)"],
        correct: 1,
        explanation: "セルを表すには「Cells(行,列)」を使用します。"
    },
    {
        question: "セル範囲を表すにはどのプロパティを使用しますか？",
        choices: ["Range", "Value", "Text"],
        correct: 0,
        explanation: "セル範囲を表すには「Range」プロパティを使用します。"
    },
    // 4-1 オブジェクト式（追加問題）
    {
      question: "VBAでコレクションの要素を参照する正しい方法は？",
      choices: ["Worksheets[1]", "Worksheets(1)", "Worksheets.Item(1)"],
      correct: 1,
      explanation: "VBAでコレクションの要素を参照するには、括弧()を使用します。例えば、Worksheets(1)は最初のワークシートを指します。"
    },
    // 4-2 ステートメント
    {
        question: "次のうち、正しいVBAステートメントはどれですか？",
        choices: ["Cell.Value = 10", "Range(\"A1\").Value = 10", "A1 = 10"],
        correct: 1,
        explanation: "正しいステートメントは Range(\"A1\").Value = 10 です。これはセルA1に値10を代入するステートメントです。"
    },
    // 4-3 関数
    {
        question: "VBAで文字列の長さを取得する関数は？",
        choices: ["Length()", "Len()", "Size()"],
        correct: 1,
        explanation: "文字列の長さを取得するには Len() 関数を使用します。例：stringLength = Len(\"Hello World\")"
    },
    // 4-4 演算子
    {
        question: "VBAで文字列を連結する演算子は？",
        choices: ["+", "&", "||"],
        correct: 1,
        explanation: "VBAで文字列を連結するには & 演算子を使用します。例：\"Hello\" & \" \" & \"World\" は \"Hello World\" となります。"
    },
    // 4-4 演算子（追加問題）
    {
        question: "VBAで等しくないことを表す演算子は？",
        choices: ["!=", "<>", "Not="],
        correct: 1,
        explanation: "VBAで等しくないことを表す演算子は <> です。例：If a <> b Then ..."
    },
    {
        question: "VBAで論理AND演算子として使用されるのは？",
        choices: ["&&", "AND", "&"],
        correct: 1,
        explanation: "VBAでは、論理AND演算子として 'AND' を使用します。例：If a > 0 AND b > 0 Then ..."
    },
    // 5 変数と定数
    // 5-1 変数とは
    {
        question: "VBAにおいて、変数の主な用途は何ですか？",
        choices: ["データを永続的に保存する", "一時的にデータを格納する", "プロシージャを定義する"],
        correct: 1,
        explanation: "変数の主な用途は、プログラムの実行中に一時的にデータを格納することです。これにより、後で使用するためにデータを保持したり、操作したりすることができます。"
    },
    // 5-2 変数を宣言する
    {
        question: "VBAで変数を宣言する際に使用するキーワードは？",
        choices: ["var", "let", "Dim"],
        correct: 2,
        explanation: "VBAでは、変数を宣言する際に Dim キーワードを使用します。例：Dim myVariable As Integer"
    },
    {
        question: "変数を宣言するにはどのキーワードを使用しますか？",
        choices: ["Const", "Dim", "Set"],
        correct: 1,
        explanation: "変数を宣言するには「Dim」キーワードを使用します。"
    },
    // 5-3 変数に代入する
    {
        question: "次のうち、正しい変数への代入方法はどれですか？",
        choices: ["myVar := 10", "myVar == 10", "myVar = 10"],
        correct: 2,
        explanation: "VBAでは、変数に値を代入する際に = 演算子を使用します。例：myVar = 10"
    },
    {
        question: "変数に値を代入するにはどの記号を使用しますか？",
        choices: ["=", ":", ";"],
        correct: 0,
        explanation: "変数に値を代入するには「=」記号を使用します。"
    },
    // 5-4 変数の名前
    {
        question: "VBAで有効な変数名はどれですか？",
        choices: ["1stVariable", "My Variable", "_myVariable"],
        correct: 2,
        explanation: "VBAの変数名はアルファベット、数字、アンダースコアを使用できますが、数字で始めることはできません。また、スペースは使用できません。したがって、_myVariable が有効な変数名です。"
    },
    // 5-5 変数の適用範囲
    {
        question: "VBAにおいて、プロシージャ内で宣言された変数のスコープは？",
        choices: ["グローバル", "モジュールレベル", "ローカル"],
        correct: 2,
        explanation: "プロシージャ内で宣言された変数のスコープはローカルです。つまり、その変数はそのプロシージャ内でのみ使用可能です。"
    },
    // 5-6 定数とは
    {
        question: "VBAで定数を宣言するキーワードは？",
        choices: ["Constant", "Const", "Final"],
        correct: 1,
        explanation: "VBAでは、定数を宣言する際に Const キーワードを使用します。例：Const PI As Double = 3.14159"
    },
    {
        question: "定数を宣言するにはどのキーワードを使用しますか？",
        choices: ["Dim", "Const", "Set"],
        correct: 1,
        explanation: "定数を宣言するには「Const」キーワードを使用します。"
    }
];
