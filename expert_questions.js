const quizData = [
    // 序章 マクロを作れるようになるには
    // 序章-1 技術を使うために必要な考え方
    {
        question: "技術を効果的に使うために最も重要な考え方は次のうちどれですか？",
        choices: ["具体化", "抽象化", "単純化"],
        correct: 1,
        explanation: "抽象化は、複雑な問題や概念を単純化し、本質的な部分に焦点を当てる考え方です。これは技術を効果的に使用するための重要な基礎となります。"
    },
    {
        question: "プログラミングにおいて、大きな問題を小さな部分に分割する考え方を何と呼びますか？",
        choices: ["抽象化", "細分化", "簡略化"],
        correct: 1,
        explanation: "細分化は、大きな問題や課題を小さな、扱いやすい部分に分割する考え方です。これにより、複雑な問題を効率的に解決できます。"
    },
    // 序章-1-1 抽象化
    {
        question: "抽象化の主な目的は何ですか？",
        choices: ["問題を複雑にする", "本質的な部分に焦点を当てる", "すべての詳細を保持する"],
        correct: 1,
        explanation: "抽象化の主な目的は、複雑な問題や概念から不要な詳細を取り除き、本質的な部分に焦点を当てることです。これにより、問題の理解と解決が容易になります。"
    },
    {
        question: "プログラミングにおいて、抽象化の例として最も適切なものはどれですか？",
        choices: ["変数の使用", "無限ループの作成", "すべての処理を1つの関数に記述すること"],
        correct: 0,
        explanation: "変数の使用は抽象化の良い例です。変数は具体的な値を抽象的な名前で表現し、コードの理解と管理を容易にします。"
    },
    // 序章-1-2 細分化
    {
        question: "細分化の利点は何ですか？",
        choices: ["問題を複雑にする", "コードの再利用性を低下させる", "問題を扱いやすい小さな部分に分割する"],
        correct: 2,
        explanation: "細分化の主な利点は、大きな問題を扱いやすい小さな部分に分割することです。これにより、複雑な問題の解決が容易になり、コードの管理も改善されます。"
    },
    {
        question: "VBAにおいて、細分化を実現する主な方法は何ですか？",
        choices: ["変数の使用", "サブプロシージャとファンクションの作成", "無限ループの使用"],
        correct: 1,
        explanation: "VBAにおいて、細分化を実現する主な方法は、サブプロシージャとファンクションを作成することです。これにより、大きな問題を小さな、管理しやすい部分に分割できます。"
    },
    // 序章-1-3 簡略化
    {
        question: "プログラミングにおける簡略化の目的は何ですか？",
        choices: ["コードを複雑にする", "実行速度を遅くする", "コードをより理解しやすく、保守しやすくする"],
        correct: 2,
        explanation: "簡略化の主な目的は、コードをより理解しやすく、保守しやすくすることです。これにより、開発効率が向上し、エラーの可能性も減少します。"
    },
    {
        question: "次のうち、VBAコードの簡略化の例として最も適切なものはどれですか？",
        choices: ["意味のない変数名の使用", "繰り返し処理のためのループの使用", "すべての処理を1つの長い関数に記述すること"],
        correct: 1,
        explanation: "繰り返し処理のためのループの使用は、コードの簡略化の良い例です。同じ処理を何度も書く代わりに、ループを使用することで、コードをより簡潔で理解しやすくできます。"
    },

    // 1 プロシージャ
    {
        question: "VBAにおいて、プロシージャとは何ですか？",
        choices: ["変数の別名", "特定のタスクを実行するコードのまとまり", "Excelのワークシート"],
        correct: 1,
        explanation: "プロシージャは、特定のタスクを実行するためにまとめられたコードの集まりです。VBAでは主にSubプロシージャとFunctionプロシージャの2種類があり、コードの構造化と再利用性の向上に役立ちます。"
    },
    // 1-1 他のプロシージャを呼び出す
    {
        question: "あるプロシージャから別のプロシージャを呼び出すことの利点は何ですか？",
        choices: ["メモリ使用量を増やす", "コードの重複を減らし、保守性を高める", "実行速度を遅くする"],
        correct: 1,
        explanation: "他のプロシージャを呼び出すことで、コードの重複を減らし、保守性を高めることができます。これにより、同じ機能を複数の場所で使い回すことが可能になり、変更が必要な場合も一箇所で済むため、効率的なコード管理が可能になります。"
    },
    // 1-1-1 モジュールレベル変数
    {
        question: "モジュールレベル変数の特徴として正しいものは？",
        choices: [
            "プロシージャ内でのみアクセス可能",
            "モジュール内のすべてのプロシージャからアクセス可能",
            "他のモジュールからも直接アクセス可能"
        ],
        correct: 1,
        explanation: "モジュールレベル変数は、そのモジュール内のすべてのプロシージャからアクセス可能です。これにより、複数のプロシージャ間でデータを共有することができます。ただし、他のモジュールから直接アクセスすることはできません。"
    },
    {
        question: "次のうち、モジュールレベル変数の適切な宣言はどれですか？",
        choices: [
            "Sub MyVar As Integer",
            "Dim MyVar As Integer",
            "Private MyVar As Integer"
        ],
        correct: 2,
        explanation: "モジュールレベル変数は通常、`Private` または `Public` キーワードを使用して宣言します。`Private MyVar As Integer` は、そのモジュール内でのみアクセス可能な整数型変数 MyVar を宣言します。`Dim` はプロシージャ内で使用するのが一般的です。"
    },
    // 1-2 Function プロシージャ
    {
        question: "Functionプロシージャの主な特徴は何ですか？",
        choices: [
            "値を返すことができない",
            "値を返すことができる",
            "常に引数が必要"
        ],
        correct: 1,
        explanation: "Functionプロシージャの主な特徴は、値を返すことができる点です。これにより、計算結果や処理結果を呼び出し元に返すことができ、より柔軟なプログラミングが可能になります。Subプロシージャとは異なり、Functionは必ず何らかの値を返します。"
    },
    {
        question: "次のうち、正しいFunctionプロシージャの使用例はどれですか？",
        choices: [
            "Call MyFunction()",
            "MyFunction()",
            "result = MyFunction()"
        ],
        correct: 2,
        explanation: "Functionプロシージャは値を返すため、通常は戻り値を変数に代入するか、他の式の一部として使用します。`result = MyFunction()` は、MyFunction の戻り値を result 変数に代入する正しい使用例です。`Call` キーワードは主にSubプロシージャで使用します。"
    },
    // 1-3 引数を渡す
    {
        question: "プロシージャに引数を渡す主な目的は何ですか？",
        choices: [
            "プロシージャの実行速度を上げる",
            "プロシージャの再利用性と柔軟性を高める",
            "変数の数を減らす"
        ],
        correct: 1,
        explanation: "引数を渡す主な目的は、プロシージャの再利用性と柔軟性を高めることです。引数を使用することで、同じプロシージャを異なるデータで動作させることができ、コードの重複を減らしつつ、様々な状況に対応できるようになります。"
    },
    // 1-3-1 参照渡しと値渡し
    {
        question: "VBAにおける「値渡し」の特徴として正しいものは？",
        choices: [
            "引数の値が変更されると、元の変数も変更される",
            "引数の値のコピーが渡され、元の変数は影響を受けない",
            "配列のみに適用できる"
        ],
        correct: 1,
        explanation: "値渡しでは、引数の値のコピーがプロシージャに渡されます。そのため、プロシージャ内で引数の値を変更しても、呼び出し元の元の変数には影響しません。これにより、意図しない副作用を防ぐことができます。"
    },
    {
        question: "VBAで参照渡しを明示的に指定するためのキーワードは何ですか？",
        choices: ["ByVal", "ByRef", "ByPointer"],
        correct: 1,
        explanation: "`ByRef` キーワードを使用すると、引数を参照渡しで渡すことを明示的に指定できます。参照渡しでは、引数の参照（メモリ上の位置）が渡されるため、プロシージャ内での変更が元の変数に反映されます。デフォルトでは引数は `ByRef` で渡されますが、明示的に指定することで意図を明確にできます。"
    },
    // 1-4 引数を使わないで値を共有する
    {
        question: "引数を使わずに値を共有する方法として、VBAでよく使われるのは次のうちどれですか？",
        choices: [
            "グローバル変数の使用",
            "静的変数の使用",
            "定数の使用"
        ],
        correct: 0,
        explanation: "グローバル変数（モジュールレベル変数やPublic変数）の使用は、引数を使わずに値を共有する一般的な方法です。ただし、過度の使用はコードの可読性や保守性を低下させる可能性があるため、適切な場面で慎重に使用する必要があります。"
    },
    {
        question: "VBAにおいて、静的変数（Static変数）の特徴として正しいものは？",
        choices: [
            "プロシージャの呼び出し間で値が保持される",
            "すべてのプロシージャから参照できる",
            "定数と同じように値を変更できない"
        ],
        correct: 0,
        explanation: "静的変数（Static変数）は、プロシージャの呼び出し間でその値が保持されます。つまり、プロシージャが終了しても変数の値はリセットされず、次回の呼び出し時にも前回の値を保持しています。これは、カウンターや累積値の管理などに便利ですが、使用には注意が必要です。"
    },

    // 2 変数
    {
        question: "VBAにおいて、変数のデータ型を明示的に指定することの利点は何ですか？",
        choices: [
            "メモリ使用量を増やす",
            "コードの実行速度を向上させ、型の不一致によるエラーを防ぐ",
            "すべての変数を文字列型にする"
        ],
        correct: 1,
        explanation: "変数のデータ型を明示的に指定することで、コードの実行速度が向上し、型の不一致によるエラーを防ぐことができます。また、コードの可読性も向上し、意図しないデータ型の変換を避けることができます。例えば、`Dim age As Integer` のように指定します。"
    },
    // 2-1 配列
    {
        question: "VBAで配列を使用する主な利点は何ですか？",
        choices: [
            "常に単一の値しか格納できない",
            "関連するデータをグループ化し、効率的に管理できる",
            "ループを使用できなくなる"
        ],
        correct: 1,
        explanation: "配列を使用すると、関連するデータをグループ化し、効率的に管理できます。これにより、複数の類似したデータを一つの変数名で扱うことができ、ループを使用して簡単にデータを処理することが可能になります。"
    },
    // 2-1-1 配列を宣言する
    {
        question: "VBAで1から10までの整数を格納する配列を宣言する正しい方法は次のうちどれですか？",
        choices: [
            "Dim numbers(10) As Integer",
            "Dim numbers(1 To 10) As Integer",
            "Dim numbers[1-10] As Integer"
        ],
        correct: 1,
        explanation: "`Dim numbers(1 To 10) As Integer` が正しい宣言方法です。これにより、インデックス1から10までの10個の整数を格納できる配列が作成されます。VBAのデフォルトでは配列は0から始まりますが、この方法で明示的に開始インデックスを指定できます。"
    },
    // 2-1-2 配列を受け取る
    {
        question: "サブプロシージャで配列を引数として受け取る際の正しい宣言は次のうちどれですか？",
        choices: [
            "Sub ProcessArray(arr As Integer)",
            "Sub ProcessArray(arr() As Integer)",
            "Sub ProcessArray(arr[] As Integer)"
        ],
        correct: 1,
        explanation: "`Sub ProcessArray(arr() As Integer)` が正しい宣言です。括弧 `()` は、引数が配列であることを示します。これにより、様々なサイズの整数配列をこのサブプロシージャに渡すことができます。"
    },
    // 2-2 動的配列
    {
        question: "VBAにおいて、動的配列の主な特徴は何ですか？",
        choices: [
            "コンパイル時にサイズが固定される",
            "実行時にサイズを変更できる",
            "要素を追加できない"
        ],
        correct: 1,
        explanation: "動的配列の主な特徴は、実行時にそのサイズを変更できることです。これにより、必要に応じて配列のサイズを拡大または縮小することができ、メモリを効率的に使用できます。`ReDim` ステートメントを使用してサイズを変更します。"
    },
    // 2-2-1 Preserveキーワード
    {
        question: "ReDimステートメントでPreserveキーワードを使用する目的は何ですか？",
        choices: [
            "配列のサイズを固定する",
            "配列の既存のデータを保持したままサイズを変更する",
            "配列のすべての要素を0にリセットする"
        ],
        correct: 1,
        explanation: "`Preserve` キーワードは、配列の既存のデータを保持したままサイズを変更するために使用します。例えば、`ReDim Preserve myArray(1 To newSize)` とすることで、既存のデータを保持しつつ配列のサイズを変更できます。これは特に、配列に格納されたデータを失うことなくサイズを拡張する際に有用です。"
    },
    // 2-3 オブジェクト変数
    {
        question: "VBAでオブジェクト変数を使用する主な目的は何ですか？",
        choices: [
            "数値計算を高速化する",
            "ExcelやWordなどのオブジェクトを参照し操作する",
            "文字列を格納する"
        ],
        correct: 1,
        explanation: "オブジェクト変数の主な目的は、ExcelのセルやWordの段落などのオブジェクトを参照し操作することです。これにより、VBAから様々なアプリケーションのオブジェクトを効率的に制御できます。"
    },
    // 2-3-1 オブジェクト変数を宣言する
    {
        question: "Excelのワークシートを参照するオブジェクト変数を宣言する正しい方法は次のうちどれですか？",
        choices: [
            "Dim ws As Worksheet",
            "Dim ws As New Worksheet",
            "Dim ws = Worksheet"
        ],
        correct: 0,
        explanation: "`Dim ws As Worksheet` が正しい宣言方法です。これにより、`Worksheet` オブジェクトを参照できる変数 `ws` が作成されます。ただし、この時点では `ws` は `Nothing` を参照しているため、使用前に有効なワークシートオブジェクトを割り当てる必要があります。"
    },
    // 2-3-2 オブジェクト変数にオブジェクトを格納する
    {
        question: "宣言したオブジェクト変数にExcelの現在のアクティブシートを割り当てる正しいコードは次のうちどれですか？",
        choices: [
            "Set ws = ActiveSheet",
            "ws = ActiveSheet",
            "Let ws = ActiveSheet"
        ],
        correct: 0,
        explanation: "`Set ws = ActiveSheet` が正しい方法です。オブジェクト変数にオブジェクトを割り当てる際は、必ず `Set` キーワードを使用します。これにより、`ws` 変数に現在のアクティブシートへの参照が格納されます。"
    },
    // 2-4 変数の演算
    {
        question: "VBAで整数型の変数 `x` と浮動小数点型の変数 `y` を加算し、結果を `z` に格納する場合、`z` の適切なデータ型は何ですか？",
        choices: [
            "Integer",
            "Double",
            "Variant"
        ],
        correct: 1,
        explanation: "整数型と浮動小数点型の加算結果を正確に格納するには、`Double` 型が適切です。`Double` は高精度の浮動小数点数を扱えるため、計算結果の精度を損なうことなく格納できます。例: `Dim z As Double: z = x + y`"
    },
    // 2-4-1 カウントする
    {
        question: "VBAでループ内の特定の条件の出現回数をカウントする最も一般的な方法は次のうちどれですか？",
        choices: [
            "配列を使用する",
            "カウンタ変数をインクリメントする",
            "再帰を使用する"
        ],
        correct: 1,
        explanation: "ループ内で特定の条件の出現回数をカウントする最も一般的な方法は、カウンタ変数をインクリメントすることです。例えば、`counter = counter + 1` または `counter += 1` のように、条件が満たされるたびにカウンタを増やします。"
    },
    // 2-4-2 合計する
    {
        question: "配列内の数値を合計する際に使用できるVBAの組み込み関数は次のうちどれですか？",
        choices: [
            "Sum()",
            "Total()",
            "Application.WorksheetFunction.Sum()"
        ],
        correct: 2,
        explanation: "`Application.WorksheetFunction.Sum()` を使用して、配列内の数値を簡単に合計できます。これはExcelのSUM関数をVBA内で直接使用できるようにするものです。例: `total = Application.WorksheetFunction.Sum(myArray)`"
    },
    // 2-5 文字列を結合する
    {
        question: "VBAで文字列を結合する最も効率的な方法は次のうちどれですか？",
        choices: [
            "& 演算子を使用する",
            "+ 演算子を使用する",
            "Join() 関数を使用する"
        ],
        correct: 0,
        explanation: "VBAで文字列を結合する最も一般的で効率的な方法は、`&` 演算子を使用することです。例: `fullName = firstName & \" \" & lastName`。大量の文字列を結合する場合は、`Join()` 関数も効率的ですが、通常の結合操作では `&` が最適です。"
    },

    
    // 3 ステートメント
    {
        question: "VBAにおけるステートメントの主な役割は何ですか？",
        choices: [
            "変数を宣言する",
            "プログラムの実行フローを制御する",
            "関数を定義する"
        ],
        correct: 1,
        explanation: "ステートメントの主な役割は、プログラムの実行フローを制御することです。これには条件分岐、ループ、サブルーチンの終了などが含まれます。適切なステートメントを使用することで、プログラムの動作を効果的に制御できます。"
    },
    // 3-1 Exit ステートメント
    {
        question: "Exit ステートメントの主な用途は何ですか？",
        choices: [
            "プログラム全体を終了する",
            "現在の制御構造（ループやプロシージャ）を即座に抜ける",
            "変数の値を初期化する"
        ],
        correct: 1,
        explanation: "Exit ステートメントの主な用途は、現在の制御構造（ループやプロシージャ）を即座に抜けることです。これにより、特定の条件が満たされた場合に処理を中断し、効率的にコードの実行を制御できます。"
    },
    // 3-1-1 Exit Sub ステートメント/Exit Function ステートメント
    {
        question: "Exit Sub ステートメントと Exit Function ステートメントの違いは何ですか？",
        choices: [
            "違いはない",
            "Exit Sub はサブプロシージャを、Exit Function は関数を終了する",
            "Exit Sub は値を返せないが、Exit Function は値を返せる"
        ],
        correct: 2,
        explanation: "Exit Sub はサブプロシージャを即座に終了し、Exit Function は関数を即座に終了します。主な違いは、Exit Function を使用する場合、関数の戻り値を設定してから終了することができる点です。例えば、`Function MyFunc() As Integer: MyFunc = 5: Exit Function: End Function`"
    },
    // 3-1-2 Exit For ステートメント
    {
        question: "For...Next ループ内で Exit For ステートメントを使用する主な目的は何ですか？",
        choices: [
            "ループのカウンタを0にリセットする",
            "特定の条件が満たされた時点でループを終了する",
            "ループを無限に続ける"
        ],
        correct: 1,
        explanation: "Exit For ステートメントの主な目的は、特定の条件が満たされた時点でFor...Nextループを即座に終了することです。これにより、ループの途中で必要な条件が満たされた場合に、残りの不要な繰り返しをスキップして処理効率を向上させることができます。"
    },
    // 3-1-3 Exit Do ステートメント
    {
        question: "Do...Loop 内で Exit Do ステートメントを使用する際の注意点は何ですか？",
        choices: [
            "使用できるのは1回だけである",
            "必ず条件文と組み合わせて使用する必要がある",
            "ループが無限に続く可能性があるため、適切な条件下でのみ使用する"
        ],
        correct: 2,
        explanation: "Exit Do ステートメントを使用する際の主な注意点は、適切な条件下でのみ使用することです。不適切に使用すると、ループが意図せず終了したり、逆に無限ループになる可能性があります。通常、If文と組み合わせて特定の条件が満たされた場合にのみExit Doを実行するようにします。"
    },
    // 3-2 Select Case ステートメント
    {
        question: "Select Case ステートメントの主な利点は何ですか？",
        choices: [
            "常に If...ElseIf より高速に動作する",
            "複数の条件分岐を簡潔に記述でき、コードの可読性が向上する",
            "無限ループを簡単に作成できる"
        ],
        correct: 1,
        explanation: "Select Case ステートメントの主な利点は、複数の条件分岐を簡潔に記述でき、コードの可読性が向上することです。特に、単一の変数や式に基づいて多くの分岐がある場合に有効です。例: `Select Case score: Case 90 To 100: grade = \"A\": Case 80 To 89: grade = \"B\": ... End Select`"
    },
    // 3-3 Do...Loopステートメント
    {
        question: "Do...Loop ステートメントの中で、ループの継続条件を最初に評価するのは次のうちどれですか？",
        choices: [
            "Do While...Loop",
            "Do...Loop While",
            "Do Until...Loop"
        ],
        correct: 0,
        explanation: "Do While...Loop 構造では、ループの継続条件が最初に評価されます。これにより、条件が最初から偽の場合、ループ本体が一度も実行されない可能性があります。一方、Do...Loop While では、ループ本体が少なくとも1回は実行されます。例: `Do While count < 10: ... Loop`"
    },
    // 3-4 For Each...Nextステートメント
    {
        question: "For Each...Next ステートメントの主な用途は何ですか？",
        choices: [
            "特定の回数だけループを繰り返す",
            "コレクションや配列の各要素に対して処理を行う",
            "無限ループを作成する"
        ],
        correct: 1,
        explanation: "For Each...Next ステートメントの主な用途は、コレクションや配列の各要素に対して処理を行うことです。これにより、要素の数を事前に知る必要なく、すべての要素を簡単に処理できます。例: `For Each cell In Range(\"A1:D10\"): ... Next cell`"
    },
    // 3-4-1 コレクションを操作する
    {
        question: "VBAでコレクションを操作する際のFor Each...Nextの利点は何ですか？",
        choices: [
            "コレクションの要素を逆順に処理できる",
            "インデックスを気にせずに各要素にアクセスできる",
            "コレクションのサイズを動的に変更できる"
        ],
        correct: 1,
        explanation: "For Each...Nextを使用してコレクションを操作する主な利点は、インデックスを気にせずに各要素に簡単にアクセスできることです。これにより、コードがシンプルになり、エラーのリスクも減少します。例: `For Each item In myCollection: Debug.Print item: Next item`"
    },
    // 3-4-2 セル範囲を操作する
    {
        question: "Excelのセル範囲を操作する際、For Each...Nextを使用する利点は何ですか？",
        choices: [
            "セルの値を変更できない",
            "非表示のセルをスキップできる",
            "行や列を意識せずに各セルにアクセスできる"
        ],
        correct: 2,
        explanation: "For Each...Nextを使用してExcelのセル範囲を操作する主な利点は、行や列を意識せずに各セルに簡単にアクセスできることです。これにより、複雑な範囲でも簡潔にコードを書くことができます。例: `For Each cell In Range(\"A1:Z100\"): cell.Value = cell.Value * 2: Next cell`"
    },
    // 3-4-3 配列を操作する
    {
        question: "VBAで配列を操作する際、For Each...Nextを使用する場合の制限は何ですか？",
        choices: [
            "多次元配列には使用できない",
            "配列の要素を変更できない",
            "配列のインデックスを取得できない"
        ],
        correct: 2,
        explanation: "For Each...Nextで配列を操作する際の主な制限は、配列のインデックスを直接取得できないことです。要素の値にはアクセスできますが、その要素が配列の何番目にあるかを知ることはできません。インデックスが必要な場合は、通常のForループを使用する必要があります。"
    },
    // 3-5 Ifステートメント
    {
        question: "VBAのIf...Then...Else構造で、複数の条件を1行で書く場合の正しい方法は次のうちどれですか？",
        choices: [
            "If x > 0 Then y = 1 Else If x < 0 Then y = -1 Else y = 0",
            "If x > 0 Then y = 1 ElseIf x < 0 Then y = -1 Else y = 0",
            "If x > 0: y = 1 ElseIf x < 0: y = -1 Else: y = 0"
        ],
        correct: 0,
        explanation: "1行でIf...Then...Else構造を書く場合、`If x > 0 Then y = 1 Else If x < 0 Then y = -1 Else y = 0` が正しい方法です。この形式では、`ElseIf`ではなく`Else If`を使用し、コロン(`:`)ではなくスペースで区切ります。複数行の場合は異なる構文を使用します。"
    },
    // 3-5-1 複数条件による条件分岐
    {
        question: "VBAで複数の条件を組み合わせて条件分岐を行う際、どの論理演算子を使用しますか？",
        choices: [
            "AND と OR",
            "& と |",
            "+ と -"
        ],
        correct: 0,
        explanation: "VBAで複数の条件を組み合わせる際は、`AND` と `OR` 論理演算子を使用します。`AND` は両方の条件が真の場合に真を返し、`OR` はいずれかの条件が真の場合に真を返します。例: `If age >= 18 AND country = \"Japan\" Then ... `"
    },


    // 4 ファイルの操作
    {
        question: "VBAでファイル操作を行う際、主に使用されるオブジェクトは次のうちどれですか？",
        choices: [
            "File オブジェクト",
            "FileSystemObject オブジェクト",
            "Workbook オブジェクト"
        ],
        correct: 1,
        explanation: "VBAでファイル操作を行う際、主に FileSystemObject (FSO) オブジェクトが使用されます。FSO は、ファイルやフォルダの作成、削除、移動、コピーなど、多様なファイルシステム操作を提供します。使用するには、Microsoft Scripting Runtime ライブラリの参照設定が必要です。"
    },
    // 4-1 ブックを開く
    {
        question: "VBAで既存のExcelブックを開く正しいコードは次のうちどれですか？",
        choices: [
            "Workbooks.Open(\"C:\\path\\to\\file.xlsx\")",
            "OpenWorkbook(\"C:\\path\\to\\file.xlsx\")",
            "Excel.Open(\"C:\\path\\to\\file.xlsx\")"
        ],
        correct: 0,
        explanation: "`Workbooks.Open(\"C:\\path\\to\\file.xlsx\")` が正しいコードです。このメソッドは指定されたパスのExcelブックを開き、新しく開いたWorkbookオブジェクトを返します。オプションのパラメータを使用して、読み取り専用で開くなどの追加設定も可能です。"
    } ,

    // 4-1-1 フォルダー内の複数のブックを開く
    {
        question: "フォルダ内のすべてのExcelファイルを開くためのVBAコードとして正しいものは？",
        choices: [
            "For Each file In Folder.Files: Workbooks.Open file.Path: Next file",
            "For Each file In Directory.Files: OpenWorkbook file.Name: Next file",
            "For Each file In Folder: If Right(file.Name, 4) = \".xls\" Then Workbooks.Open file: Next file"
        ],
        correct: 0,
        explanation: "`For Each file In Folder.Files: Workbooks.Open file.Path: Next file` が正しいアプローチです。このコードはFSOを使用してフォルダ内のファイルを列挙し、各ファイルのパスを使ってWorkbooks.Openメソッドでブックを開きます。ただし、実際の使用時には、ファイル拡張子のチェックを追加するとより安全です。"
    },
    // 4-2 ブックを保存する
    {
        question: "VBAで現在のブックを新しい名前で保存する正しいコードは？",
        choices: [
            "ActiveWorkbook.Save As \"NewFileName.xlsx\"",
            "ThisWorkbook.SaveAs Filename:=\"NewFileName.xlsx\"",
            "Workbooks(1).SaveAs \"NewFileName.xlsx\""
        ],
        correct: 1,
        explanation: "`ThisWorkbook.SaveAs Filename:=\"NewFileName.xlsx\"` が正しいコードです。SaveAsメソッドを使用して、現在のブックを新しい名前で保存します。Filenameパラメータは明示的に指定することをお勧めします。また、ThisWorkbookは現在のVBAプロジェクトが含まれるブックを指します。"
    } ,
    // 4-3 ファイルをコピーする
    {
        question: "VBAでファイルをコピーする際に使用されるFSOのメソッドは？",
        choices: [
            "File.Copy",
            "FSO.CopyFile",
            "Folder.CopyFile"
        ],
        correct: 1,
        explanation: "`FSO.CopyFile` メソッドを使用してファイルをコピーします。使用例：`FSO.CopyFile \"C:\\source\\file.txt\", \"C:\\destination\\file.txt\"` このメソッドは、ソースファイルパスと目的地ファイルパスを引数に取ります。オプションで、上書きを許可するブール値を第3引数に指定できます。"
    },

    // 4-4 フォルダーを操作する
    {
        question: "VBAで新しいフォルダを作成するための正しいコードは？",
        choices: [
            "MkDir \"C:\\NewFolder\"",
            "FSO.CreateFolder \"C:\\NewFolder\"",
            "Folder.Create \"C:\\NewFolder\""
        ],
        correct: 1,
        explanation: "`FSO.CreateFolder \"C:\\NewFolder\"` が正しいコードです。FileSystemObjectの CreateFolder メソッドを使用して新しいフォルダを作成します。このメソッドは指定されたパスに新しいフォルダを作成し、作成されたFolderオブジェクトを返します。既にフォルダが存在する場合はエラーが発生するため、通常はフォルダの存在チェックと組み合わせて使用します。"
    }
    
];
