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
    },
    // 6 セルの操作
    // 6-1 セルを操作する
    {
        question: "VBAでセルA1を参照する正しい方法は？",
        choices: ["Cells(A, 1)", "Range(\"A1\")", "A1"],
        correct: 1,
        explanation: "セルA1を参照する正しい方法は Range(\"A1\") です。これはオブジェクト式の一例で、特定のセルを指定するのに使用されます。"
    },
    // 6-2 Value プロパティ
    {
        question: "セルの値を取得するために使用するプロパティは？",
        choices: ["Text", "Formula", "Value"],
        correct: 2,
        explanation: "セルの値を取得するには Value プロパティを使用します。例：cellValue = Range(\"A1\").Value"
    },
    // 6-3 セルの様子を表すプロパティ
    {
        question: "セルに表示されているテキストを取得するプロパティは？",
        choices: ["Value", "Text", "DisplayedText"],
        correct: 1,
        explanation: "セルに表示されているテキストを取得するには Text プロパティを使用します。これは、セルの表示形式に従って整形された文字列を返します。"
    },
    {
        question: "セルの数式を取得するために使うプロパティはどれですか？",
        choices: ["Text", "Value", "Formula"],
        correct: 2,
        explanation: "セルの数式を取得するためには「Formula」プロパティを使用します。"
    },
    // 6-4 別のセルを表すプロパティ
    {
        question: "現在のセルから相対的に別のセルを参照するプロパティは？",
        choices: ["Relative", "Offset", "Move"],
        correct: 1,
        explanation: "Offset プロパティを使用すると、現在のセルから相対的に別のセルを参照できます。例：Range(\"A1\").Offset(1, 1) はB2セルを指します。"
    },
    {
        question: "別のセルを表すために使うプロパティはどれですか？",
        choices: ["Offset", "Resize", "End"],
        correct: 0,
        explanation: "別のセルを表すためには「Offset」プロパティを使用します。"
    },
    {
        question: "セル範囲を選択するために使うメソッドはどれですか？",
        choices: ["Activate", "Select", "Copy"],
        correct: 1,
        explanation: "セル範囲を選択するためには「Select」メソッドを使用します。"
    },
    // Resize プロパティ
    {
        question: "選択範囲を拡大または縮小するプロパティは？",
        choices: ["Expand", "Resize", "Adjust"],
        correct: 1,
        explanation: "Resize プロパティを使用すると、選択範囲を拡大または縮小できます。例：Range(\"A1\").Resize(2, 2) はA1:B2の範囲を選択します。"
    },
    //End プロパティ
    {
        question: "Endプロパティの主な用途は何ですか？",
        choices: ["セル範囲の最後のセルを選択する", "プロシージャを終了する", "ループを終了する"],
        correct: 0,
        explanation: "Endプロパティは、特定の方向の最後の使用されているセルを選択するために使用されます。例：Range(\"A1\").End(xlDown)は、A列の最後の使用されているセルを選択します。"
    },
    // 6-5 セルを表すその他の単語（追加問題）
    {
        question: "現在アクティブなセルを参照するためのプロパティは？",
        choices: ["ActiveCell", "CurrentCell", "SelectedCell"],
        correct: 0,
        explanation: "現在アクティブなセル（選択されているセル）を参照するには、ActiveCell プロパティを使用します。"
    },
    {
        question: "現在選択されている範囲を参照するためのプロパティは？",
        choices: ["ActiveRange", "CurrentSelection", "Selection"],
        correct: 2,
        explanation: "現在選択されている範囲を参照するには、Selection プロパティを使用します。"
    },
    // 6-6 セルのメソッド
    {
        question: "特定のセルを選択状態にするメソッドは？",
        choices: ["Activate", "Select", "Focus"],
        correct: 1,
        explanation: "特定のセルを選択状態にするには Select メソッドを使用します。例：Range(\"A1\").Select"
    },
    // 6-6 セルのメソッド（追加問題）
    {
        question: "セルの内容をクリアするメソッドは？",
        choices: ["Clear", "Delete", "ClearContents"],
        correct: 2,
        explanation: "セルの内容をクリアするには、ClearContents メソッドを使用します。これはセルの書式を保持したまま、内容のみをクリアします。"
    },
    {
        question: "セル（または行、列）を削除するメソッドは？",
        choices: ["Remove", "Delete", "Erase"],
        correct: 1,
        explanation: "セル、行、または列を削除するには、Delete メソッドを使用します。例：Range(\"A1\").Delete"
    },


    // 6-7 複数セル（セル範囲）の指定
    {
        question: "A1からC3までのセル範囲を指定する正しい方法は？",
        choices: ["Range(\"A1-C3\")", "Range(\"A1:C3\")", "Range(A1,C3)"],
        correct: 1,
        explanation: "A1からC3までのセル範囲を指定するには、Range(\"A1:C3\") を使用します。コロン (:) を使って範囲を指定します。"
    },
    // 6-8 行や列の指定
    {
        question: "3行目全体を指定する正しい方法は？",
        choices: ["Range(\"3:3\")", "Rows(3)", "Range(\"A3:XFD3\")"],
        correct: 1,
        explanation: "3行目全体を指定するには Rows(3) を使用します。これは3行目のすべてのセルを選択します。"
    },
    // 7 ステートメント（続き）
    // 7-1 For...Nextステートメント
    {
        question: "1から10まで数えるFor...Nextループの正しい書き方は？",
        choices: ["For i = 1 To 10 ... Next", "For i = 1 Until 10 ... Next", "For i = 1; i <= 10; i++ ... Next"],
        correct: 0,
        explanation: "VBAでは、For i = 1 To 10 ... Next が正しい For...Next ループの書き方です。これは1から10まで繰り返し処理を行います。"
    },
    // 7-2 Ifステートメント
    {
        question: "VBAで複数の条件を持つIf文を書く場合、どのキーワードを使用しますか？",
        choices: ["Else If", "ElseIf", "Elsif"],
        correct: 1,
        explanation: "VBAでは、複数の条件を持つIf文を書く場合、ElseIf キーワードを使用します。例：If ... ElseIf ... Else ... End If"
    },
    // 7-3 Withステートメント
    {
        question: "Withステートメントの主な目的は何ですか？",
        choices: ["コードの実行速度を上げる", "オブジェクト名の繰り返しを避ける", "変数のスコープを制限する"],
        correct: 1,
        explanation: "Withステートメントの主な目的は、オブジェクト名の繰り返しを避けることです。これにより、コードが簡潔になり、読みやすくなります。"
    },
    // 8 関数
    // 8-1 日付や時刻を操作する関数
    {
        question: "現在の日付と時刻を取得する関数は？",
        choices: ["CurrentDateTime()", "Now()", "GetNow()"],
        correct: 1,
        explanation: "現在の日付と時刻を取得するには Now() 関数を使用します。この関数は現在の日付と時刻を返します。"
    },
    // 8-2 文字列を操作する関数
    {
        question: "文字列の一部を取り出す関数は次のうちどれですか？",
        choices: ["Substring()", "Extract()", "Mid()"],
        correct: 2,
        explanation: "VBAでは、Mid() 関数を使用して文字列の一部を取り出します。例：Mid(\"Hello\", 2, 2) は 'el' を返します。"
    },
    // 8-2 文字列を操作する関数（追加問題）
    {
        question: "文字列を小文字に変換する関数は？",
        choices: ["Lower()", "LCase()", "ToLower()"],
        correct: 1,
        explanation: "文字列を小文字に変換するには、LCase() 関数を使用します。例：LCase(\"HELLO\") は \"hello\" を返します。"
    },
    {
        question: "文字列の先頭と末尾の空白を削除する関数は？",
        choices: ["Trim()", "Strip()", "Clean()"],
        correct: 0,
        explanation: "文字列の先頭と末尾の空白を削除するには、Trim() 関数を使用します。例：Trim(\"  Hello  \") は \"Hello\" を返します。"
    },
    {
        question: "文字列内の特定の部分文字列を別の文字列に置換する関数は？",
        choices: ["Substitute()", "Replace()", "Swap()"],
        correct: 1,
        explanation: "文字列内の特定の部分を置換するには、Replace() 関数を使用します。例：Replace(\"Hello World\", \"World\", \"VBA\") は \"Hello VBA\" を返します。"
    },
    {
        question: "文字列の文字コードを変換する関数は？",
        choices: ["Convert()", "StrConv()", "CharConv()"],
        correct: 1,
        explanation: "文字列の文字コードを変換するには、StrConv() 関数を使用します。これは、全角/半角変換や大文字/小文字変換などに使用されます。"
    },
    {
        question: "指定した書式で文字列を整形する関数は？",
        choices: ["Format()", "StringFormat()", "FormatStr()"],
        correct: 0,
        explanation: "指定した書式で文字列を整形するには、Format() 関数を使用します。これは日付、時刻、数値などの書式設定に便利です。"
    },
    // 8-3 数値を操作する関数
    {
        question: "数値を指定した小数点以下の桁数に丸める関数は？",
        choices: ["Round()", "Truncate()", "Floor()"],
        correct: 0,
        explanation: "Round() 関数を使用すると、数値を指定した小数点以下の桁数に丸めることができます。例：Round(3.14159, 2) は 3.14 を返します。"
    },
    {
        question: "数値の絶対値を返す関数はどれですか？",
        choices: ["Int関数", "Round関数", "Abs関数"],
        correct: 2,
        explanation: "Abs関数は、数値の絶対値を返します。"
    },
    // 8-4 データの種類を判定する関数
    {
        question: "変数が数値かどうかを判定する関数は？",
        choices: ["IsNumber()", "IsNumeric()", "IsInteger()"],
        correct: 1,
        explanation: "IsNumeric() 関数を使用すると、変数が数値かどうかを判定できます。これは数値として解釈可能な文字列に対してもTrueを返します。"
    },
    // 8-5 文字列の入出力に関する関数
    {
        question: "ユーザーに入力を求めるダイアログボックスを表示する関数は？",
        choices: ["InputBox()", "GetInput()", "PromptUser()"],
        correct: 0,
        explanation: "InputBox() 関数を使用すると、ユーザーに入力を求めるダイアログボックスを表示できます。例：userInput = InputBox(\"名前を入力してください\")"
    },
    // 9 シートとブックの操作
    // 9-1 シートの操作
    {
        question: "新しいワークシートを追加するメソッドは？",
        choices: ["Sheets.New()", "Worksheets.Add", "AddWorksheet()"],
        correct: 1,
        explanation: "新しいワークシートを追加するには Worksheets.Add メソッドを使用します。例：Worksheets.Add"
    },
    {
        question: "シートを選択するために使うメソッドはどれですか？",
        choices: ["Activate", "Select", "Copy"],
        correct: 1,
        explanation: "シートを選択するためには「Select」メソッドを使用します。"
    },
    {
        question: "シートをコピーするために使うメソッドはどれですか？",
        choices: ["Activate", "Select", "Copy"],
        correct: 2,
        explanation: "シートをコピーするためには「Copy」メソッドを使用します。"
    },
    // 9-2 ブックの操作
    {
        question: "現在のブックを保存するメソッドは？",
        choices: ["ActiveWorkbook.Save", "ThisWorkbook.SaveAs", "Workbooks.Save"],
        correct: 0,
        explanation: "現在のブックを保存するには ActiveWorkbook.Save メソッドを使用します。これは現在アクティブなブックを保存します。"
    },
    {
        question: "新規ブックを挿入するにはどのメソッドを使用しますか？",
        choices: ["Open", "Add", "Insert"],
        correct: 1,
        explanation: "新規ブックを挿入するには「Add」メソッドを使用します。"
    },
    // 10 マクロの実行
    // 10-1 VBEから実行する
    {
        question: "VBEでマクロを実行するショートカットキーは？",
        choices: ["F5", "F9", "F11"],
        correct: 0,
        explanation: "VBEでマクロを実行するショートカットキーは F5 です。カーソルがプロシージャ内にある状態でF5を押すと、そのプロシージャが実行されます。"
    },
    // 10-2 Excelから実行する
    {
        question: "Excelからマクロを実行する際に使用されるタブは？",
        choices: ["ホーム", "挿入", "開発"],
        correct: 2,
        explanation: "Excelからマクロを実行する際は、通常「開発」タブを使用します。このタブにはマクロ関連の機能が集められています。"
    },
    // 10-3 クイックアクセスツールバー(QAT)から実行する
    {
        question: "マクロをクイックアクセスツールバー(QAT)に追加する主な利点は？",
        choices: ["マクロの実行速度が上がる", "簡単にアクセスできる", "マクロが自動的に最適化される"],
        correct: 1,
        explanation: "マクロをQATに追加する主な利点は、簡単にアクセスできることです。頻繁に使用するマクロをQATに追加することで、素早く実行できるようになります。"
    },
    // 10-4 ボタンや図形から実行する
    {
        question: "ワークシート上のボタンにマクロを割り当てる方法として正しいのは？",
        choices: ["ボタンを右クリックし、「マクロの割り当て」を選択する", "VBEでボタンのコードを直接編集する", "「開発」タブから「コントロールの挿入」を選択し、ボタンを配置する"],
        correct: 2,
        explanation: "ワークシート上にマクロを実行するボタンを配置するには、「開発」タブから「コントロールの挿入」を選択し、ボタンを配置します。その後、マクロを割り当てることができます。"
    }
];
