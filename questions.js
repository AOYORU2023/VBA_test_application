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
    // 2-2 [マクロの記録] ダイアログボックスの設定項目
    {
        question: "[マクロの記録]ダイアログボックスで設定できない項目はどれですか？",
        choices: ["マクロ名", "ショートカットキー", "マクロの説明"],
        correct: 1,
        explanation: "ショートカットキーの設定は[マクロの記録]ダイアログボックスではなく、[マクロ]ダイアログボックスで行います。"
    },
    // 2-3 個人用マクロブック
    {
        question: "個人用マクロブックの特徴として正しいのはどれですか？",
        choices: ["特定のExcelファイルでのみ使用できる", "すべてのExcelファイルで使用できる", "インターネット上で自動的に共有される"],
        correct: 1,
        explanation: "個人用マクロブックに保存されたマクロは、すべてのExcelファイルで使用することができます。これにより、頻繁に使用するマクロを簡単に再利用できます。"
    },

    // 3 モジュールとプロシージャ
    // 3-1 モジュールとは
    {
        question: "VBAにおけるモジュールの役割は何ですか？",
        choices: ["データを保存する", "プロシージャをグループ化する", "ワークシートを管理する"],
        correct: 1,
        explanation: "モジュールは関連するプロシージャ（SubやFunction）をグループ化するためのコンテナです。これにより、コードを整理し、管理しやすくなります。"
    },
    // 3-2 プロシージャとは
    {
        question: "VBAには主に2種類のプロシージャがありますが、それらは何ですか？",
        choices: ["Open プロシージャと Close プロシージャ", "Sub プロシージャと Function プロシージャ", "Public プロシージャと Private プロシージャ"],
        correct: 1,
        explanation: "VBAの主な2種類のプロシージャは、Sub プロシージャと Function プロシージャです。Sub は値を返さない処理を、Function は値を返す処理を実行します。"
    },

    // 4 VBAの構文
    // 4-1 オブジェクト式
    {
        question: "次のうち、正しいオブジェクト式はどれですか？",
        choices: ["Workbooks(1).Worksheets(1).Range(A1)", "Workbooks(1).Worksheets(1).Range("A1")", "Workbooks[1].Worksheets[1].Range["A1"]"],
        correct: 1,
        explanation: "正しいオブジェクト式は Workbooks(1).Worksheets(1).Range("A1") です。VBAではオブジェクトの階層構造をピリオドで連結し、文字列は二重引用符で囲みます。"
    },
    // 4-2 ステートメント
    {
        question: "次のうち、正しいVBAステートメントはどれですか？",
        choices: ["Cell.Value = 10", "Range("A1").Value = 10", "A1 = 10"],
        correct: 1,
        explanation: "正しいステートメントは Range("A1").Value = 10 です。これはセルA1に値10を代入するステートメントです。"
    },
    // 4-3 関数
    {
        question: "VBAで文字列の長さを取得する関数は？",
        choices: ["Length()", "Len()", "Size()"],
        correct: 1,
        explanation: "文字列の長さを取得するには Len() 関数を使用します。例：stringLength = Len("Hello World")"
    },
    // 4-4 演算子
    {
        question: "VBAで文字列を連結する演算子は？",
        choices: ["+", "&", "||"],
        correct: 1,
        explanation: "VBAで文字列を連結するには & 演算子を使用します。例："Hello" & " " & "World" は "Hello World" となります。"
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
    // 5-3 変数に代入する
    {
        question: "次のうち、正しい変数への代入方法はどれですか？",
        choices: ["myVar := 10", "myVar == 10", "myVar = 10"],
        correct: 2,
        explanation: "VBAでは、変数に値を代入する際に = 演算子を使用します。例：myVar = 10"
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

    // 6 セルの操作
    // 6-1 セルを操作する
    {
        question: "VBAでセルA1を参照する正しい方法は？",
        choices: ["Cells(A, 1)", "Range("A1")", "A1"],
        correct: 1,
        explanation: "セルA1を参照する正しい方法は Range("A1") です。これはオブジェクト式の一例で、特定のセルを指定するのに使用されます。"
    },
    // 6-2 Value プロパティ
    {
        question: "セルの値を取得するために使用するプロパティは？",
        choices: ["Text", "Formula", "Value"],
        correct: 2,
        explanation: "セルの値を取得するには Value プロパティを使用します。例：cellValue = Range("A1").Value"
    },
    // 6-3 セルの様子を表すプロパティ
    {
        question: "セルに表示されているテキストを取得するプロパティは？",
        choices: ["Value", "Text", "DisplayedText"],
        correct: 1,
        explanation: "セルに表示されているテキストを取得するには Text プロパティを使用します。これは、セルの表示形式に従って整形された文字列を返します。"
    },
    // 6-4 別のセルを表すプロパティ
    {
        question: "現在のセルから相対的に別のセルを参照するプロパティは？",
        choices: ["Relative", "Offset", "Move"],
        correct: 1,
        explanation: "Offset プロパティを使用すると、現在のセルから相対的に別のセルを参照できます。例：Range("A1").Offset(1, 1) はB2セルを指します。"
    },
    // 6-5 Resize プロパティ
    {
        question: "選択範囲を拡大または縮小するプロパティは？",
        choices: ["Expand", "Resize", "Adjust"],
        correct: 1,
        explanation: "Resize プロパティを使用すると、選択範囲を拡大または縮小できます。例：Range("A1").Resize(2, 2) はA1:B2の範囲を選択します。"
    },
    // 6-6 セルのメソッド
    {
        question: "特定のセルを選択状態にするメソッドは？",
        choices: ["Activate", "Select", "Focus"],
        correct: 1,
        explanation: "特定のセルを選択状態にするには Select メソッドを使用します。例：Range("A1").Select"
    },
    // 6-7 複数セル（セル範囲）の指定
    {
        question: "A1からC3までのセル範囲を指定する正しい方法は？",
        choices: ["Range("A1-C3")", "Range("A1:C3")", "Range(A1,C3)"],
        correct: 1,
        explanation: "A1からC3までのセル範囲を指定するには、Range("A1:C3") を使用します。コロン (:) を使って範囲を指定します。"
    },
    // 6-8 行や列の指定
    {
        question: "3行目全体を指定する正しい方法は？",
        choices: ["Range("3:3")", "Rows(3)", "Range("A3:XFD3")"],
        correct: 1,
        explanation: "3行目全体を指定するには Rows(3) を使用します。これは3行目のすべてのセルを選択します。"
    },    
    // 7 ステートメント（続き）
    // 7-1 For...Nextステートメント（続き）
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
        explanation: "VBAでは、Mid() 関数を使用して文字列の一部を取り出します。例：Mid("Hello", 2, 2) は 'el' を返します。"
    },
    // 8-3 数値を操作する関数
    {
        question: "数値を指定した小数点以下の桁数に丸める関数は？",
        choices: ["Round()", "Truncate()", "Floor()"],
        correct: 0,
        explanation: "Round() 関数を使用すると、数値を指定した小数点以下の桁数に丸めることができます。例：Round(3.14159, 2) は 3.14 を返します。"
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
        explanation: "InputBox() 関数を使用すると、ユーザーに入力を求めるダイアログボックスを表示できます。例：userInput = InputBox("名前を入力してください")"
    },

    // 9 シートとブックの操作
    // 9-1 シートの操作
    {
        question: "新しいワークシートを追加するメソッドは？",
        choices: ["Sheets.New()", "Worksheets.Add", "AddWorksheet()"],
        correct: 1,
        explanation: "新しいワークシートを追加するには Worksheets.Add メソッドを使用します。例：Worksheets.Add"
    },
    // 9-2 ブックの操作
    {
        question: "現在のブックを保存するメソッドは？",
        choices: ["ActiveWorkbook.Save", "ThisWorkbook.SaveAs", "Workbooks.Save"],
        correct: 0,
        explanation: "現在のブックを保存するには ActiveWorkbook.Save メソッドを使用します。これは現在アクティブなブックを保存します。"
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
        
        
        
        
        
        
        
        
        
        
