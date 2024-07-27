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
    },
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
    },
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
    },
    // Additional questions for deeper understanding
    {
        question: "VBAでファイルの存在を確認する最も一般的な方法は？",
        choices: [
            "If File.Exists(filePath) Then",
            "If Dir(filePath) <> \"\" Then",
            "If FSO.FileExists(filePath) Then"
        ],
        correct: 2,
        explanation: "`If FSO.FileExists(filePath) Then` が最も一般的で信頼性の高い方法です。FileSystemObjectの FileExists メソッドは、指定されたパスにファイルが存在するかどうかをブール値で返します。これは、ファイルの存在確認を行う前に必ず使用すべきメソッドです。"
    },
    {
        question: "VBAでファイルパスからファイル名のみを取得する関数は？",
        choices: [
            "Right(filePath, Len(filePath) - InStrRev(filePath, \"\\"))",
            "FSO.GetFileName(filePath)",
            "Mid(filePath, InStrRev(filePath, \"\\") + 1)"
        ],
        correct: 1,
        explanation: "`FSO.GetFileName(filePath)` が正しい方法です。FileSystemObjectの GetFileName メソッドは、フルパスからファイル名部分のみを簡単に抽出します。これは、パスの操作を行う際に非常に便利で、エラーが少ない方法です。"
    },
    {
        question: "VBAでファイルを削除する際の正しいコードは？",
        choices: [
            "Kill filePath",
            "FSO.DeleteFile filePath",
            "File.Delete filePath"
        ],
        correct: 1,
        explanation: "`FSO.DeleteFile filePath` が推奨される方法です。FileSystemObjectの DeleteFile メソッドを使用してファイルを削除します。このメソッドは、指定されたファイルが存在しない場合にエラーを発生させるため、通常はFileExistsメソッドと組み合わせて使用します。`Kill` ステートメントも使用可能ですが、より古い方法です。"
    },

    // 5-1 WorksheetFunctionの使い方
    {
        question: "VBAでExcelのワークシート関数を使用する際、正しい記述は次のうちどれですか？",
        choices: [
            "Excel.WorksheetFunction.Sum(Range(\"A1:A10\"))",
            "Application.WorksheetFunction.Sum(Range(\"A1:A10\"))",
            "Worksheet.Function.Sum(Range(\"A1:A10\"))"
        ],
        correct: 1,
        explanation: "`Application.WorksheetFunction.Sum(Range(\"A1:A10\"))` が正しい記述です。VBAでExcelのワークシート関数を使用する際は、`Application.WorksheetFunction` オブジェクトを通じてアクセスします。これにより、Excelのほとんどのビルトイン関数をVBAコード内で直接使用できます。"
    },
    // 5-2 いろいろな関数
    // 5-2-1 SUM関数
    {
        question: "VBAでSUM関数を使用して、複数の範囲の合計を計算する正しいコードは？",
        choices: [
            "Total = Application.WorksheetFunction.Sum(Range(\"A1:A10\"), Range(\"B1:B10\"))",
            "Total = Application.Sum(Range(\"A1:A10\"), Range(\"B1:B10\"))",
            "Total = WorksheetFunction.Sum(Range(\"A1:A10\") + Range(\"B1:B10\"))"
        ],
        correct: 0,
        explanation: "`Total = Application.WorksheetFunction.Sum(Range(\"A1:A10\"), Range(\"B1:B10\"))` が正しいコードです。SUM関数は複数の引数を取ることができ、それぞれの引数は別々の範囲を指定できます。この方法で、複数の範囲の合計を一度に計算できます。"
    },
    // 5-2-2 COUNTIF関数/SUMIF関数
    {
        question: "VBAでCOUNTIF関数を使用して、特定の条件を満たすセルの数を数える正しいコードは？",
        choices: [
            "Count = Application.WorksheetFunction.CountIf(Range(\"A1:A10\"), \">5\")",
            "Count = Application.CountIf(Range(\"A1:A10\"), \">5\")",
            "Count = WorksheetFunction.CountIf(Range(\"A1:A10\"), \">5\")"
        ],
        correct: 0,
        explanation: "`Count = Application.WorksheetFunction.CountIf(Range(\"A1:A10\"), \">5\")` が正しいコードです。COUNTIF関数は、指定された範囲内で特定の条件を満たすセルの数を返します。この例では、A1:A10の範囲で5より大きい値を持つセルの数をカウントします。"
    },
    // 5-2-3 LARGE関数/SMALL関数
    {
        question: "VBAでLARGE関数を使用して、範囲内の3番目に大きい値を取得する正しいコードは？",
        choices: [
            "ThirdLargest = Application.WorksheetFunction.Large(Range(\"A1:A10\"), 3)",
            "ThirdLargest = Application.Large(Range(\"A1:A10\"), 3)",
            "ThirdLargest = WorksheetFunction.Large(Range(\"A1:A10\"), 3)"
        ],
        correct: 0,
        explanation: "`ThirdLargest = Application.WorksheetFunction.Large(Range(\"A1:A10\"), 3)` が正しいコードです。LARGE関数は、データセット内のn番目に大きい値を返します。この例では、A1:A10の範囲内の3番目に大きい値を取得します。"
    },
    // 5-2-4 VLOOKUP関数
    {
        question: "VBAでVLOOKUP関数を使用して、近似一致で値を検索する正しいコードは？",
        choices: [
            "Result = Application.WorksheetFunction.VLookup(LookupValue, Range(\"A1:B10\"), 2, True)",
            "Result = Application.VLookup(LookupValue, Range(\"A1:B10\"), 2, True)",
            "Result = WorksheetFunction.VLookup(LookupValue, Range(\"A1:B10\"), 2, 1)"
        ],
        correct: 0,
        explanation: "`Result = Application.WorksheetFunction.VLookup(LookupValue, Range(\"A1:B10\"), 2, True)` が正しいコードです。VLOOKUP関数は、テーブルの最左列から検索値を探し、指定された列の値を返します。最後の引数 True は近似一致を意味し、False は完全一致を意味します。"
    },
    // 5-2-5 MATCH関数 + INDEX関数
    {
        question: "VBAでMATCH関数とINDEX関数を組み合わせて使用する正しいコードは？",
        choices: [
            "Result = Application.WorksheetFunction.Index(Range(\"B1:B10\"), Application.WorksheetFunction.Match(LookupValue, Range(\"A1:A10\"), 0))",
            "Result = Application.Index(Range(\"B1:B10\"), Application.Match(LookupValue, Range(\"A1:A10\"), 0))",
            "Result = WorksheetFunction.Index(Range(\"B1:B10\"), WorksheetFunction.Match(LookupValue, Range(\"A1:A10\"), 0))"
        ],
        correct: 0,
        explanation: "`Result = Application.WorksheetFunction.Index(Range(\"B1:B10\"), Application.WorksheetFunction.Match(LookupValue, Range(\"A1:A10\"), 0))` が正しいコードです。この組み合わせは、VLOOKUP関数の代替として使用され、より柔軟な検索が可能です。MATCH関数で検索値の位置を特定し、INDEX関数でその位置に対応する値を返します。"
    },
    // 5-2-6 EOMONTH関数
    {
        question: "VBAでEOMONTH関数を使用して、2か月後の月末日を取得する正しいコードは？",
        choices: [
            "EndDate = Application.WorksheetFunction.EoMonth(Date, 2)",
            "EndDate = Application.EoMonth(Date, 2)",
            "EndDate = WorksheetFunction.EoMonth(Date, 2)"
        ],
        correct: 0,
        explanation: "`EndDate = Application.WorksheetFunction.EoMonth(Date, 2)` が正しいコードです。EOMONTH関数は、指定された月数だけ進んだ（または戻った）月の最終日を返します。この例では、現在の日付から2か月後の月末日を取得します。"
    },
    // Additional questions for deeper understanding
    {
        question: "VBAでワークシート関数を使用する際のエラーハンドリングの best practice は？",
        choices: [
            "On Error Resume Next を使用する",
            "Try...Catch ブロックを使用する",
            "IsError 関数を使用してエラーをチェックする"
        ],
        correct: 2,
        explanation: "VBAでワークシート関数を使用する際は、`IsError` 関数を使用してエラーをチェックすることが best practice です。例: `If IsError(Application.WorksheetFunction.VLookup(...)) Then`. これにより、関数が失敗した場合に適切に処理できます。`On Error Resume Next` は問題を隠す可能性があり、VBAには `Try...Catch` ブロックがありません。"
    },
    {
        question: "VBAで配列数式を使用して複数のセルに結果を返す正しい方法は？",
        choices: [
            "Range(\"A1:A10\").FormulaArray = \"=TRANSPOSE(B1:B10)\"",
            "Range(\"A1:A10\").Formula = \"=TRANSPOSE(B1:B10)\"",
            "Application.Evaluate(\"=TRANSPOSE(B1:B10)\")"
        ],
        correct: 0,
        explanation: "`Range(\"A1:A10\").FormulaArray = \"=TRANSPOSE(B1:B10)\"` が正しい方法です。`FormulaArray` プロパティを使用することで、配列数式をVBAから適用できます。これは、複数のセルに結果を返す数式や、通常のセル参照では計算できない複雑な数式を使用する場合に有効です。"
    },

    // 6-1 セルの検索
    {
        question: "VBAでワークシート内の特定の値を検索する最も一般的なメソッドは何ですか？",
        choices: [
            "Search メソッド",
            "Find メソッド",
            "Lookup メソッド"
        ],
        correct: 1,
        explanation: "Find メソッドが最も一般的に使用されます。このメソッドは Range オブジェクトに対して使用され、指定された条件に一致する最初のセルを返します。例: `Worksheets(\"Sheet1\").Cells.Find(What:=\"検索値\", LookIn:=xlValues)`"
    },
    // 6-1-1 Findメソッド
    {
        question: "Find メソッドで大文字と小文字を区別せずに検索するためのパラメータは？",
        choices: [
            "CaseSensitive:=False",
            "MatchCase:=False",
            "IgnoreCase:=True"
        ],
        correct: 1,
        explanation: "`MatchCase:=False` が正しいパラメータです。これにより、大文字と小文字を区別せずに検索を行います。完全な使用例: `Range(\"A1:A10\").Find(What:=\"検索値\", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)`"
    },
    // 6-1-2 見つからなかったとき
    {
        question: "Find メソッドで検索結果が見つからなかった場合、どうなりますか？",
        choices: [
            "エラーが発生する",
            "Nothing が返される",
            "False が返される"
        ],
        correct: 1,
        explanation: "Find メソッドで検索結果が見つからなかった場合、Nothing が返されます。したがって、結果をチェックする際は通常 `If Not FoundCell Is Nothing Then` のようなコードを使用します。"
    },
    // 6-2 検索結果の操作
    // 6-2-1 見つかったセルを含む行を削除する
    {
        question: "Find メソッドで見つかったセルを含む行を削除する正しいコードは？",
        choices: [
            "FoundCell.EntireRow.Delete",
            "FoundCell.Row.Delete",
            "FoundCell.Delete xlShiftUp"
        ],
        correct: 0,
        explanation: "`FoundCell.EntireRow.Delete` が正しいコードです。EntireRow プロパティは、セルが属する行全体を表し、Delete メソッドはその行を削除します。"
    },
    // 6-2-2 見つかったセルを基点に別のセルを操作する
    {
        question: "Find メソッドで見つかったセルの右隣のセルの値を変更するコードは？",
        choices: [
            "FoundCell.Offset(0, 1).Value = \"新しい値\"",
            "FoundCell.Next.Value = \"新しい値\"",
            "FoundCell.Right.Value = \"新しい値\""
        ],
        correct: 0,
        explanation: "`FoundCell.Offset(0, 1).Value = \"新しい値\"` が正しいコードです。Offset メソッドは、指定されたセルから相対的な位置にあるセルを参照します。ここでは、同じ行（0）で1列右（1）のセルを指定しています。"
    },
    // 6-2-3 見つかったセルを含む行範囲をコピーする
    {
        question: "Find メソッドで見つかったセルを含む行全体を別のワークシートにコピーするコードは？",
        choices: [
            "FoundCell.EntireRow.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "FoundCell.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "FoundCell.Row.Copy Worksheets(\"Sheet2\").Range(\"A1\")"
        ],
        correct: 0,
        explanation: "`FoundCell.EntireRow.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")` が正しいコードです。これにより、見つかったセルを含む行全体が別のワークシートの A1 セルを起点にコピーされます。"
    },
    // 6-3 オートフィルターの操作
    {
        question: "VBAでオートフィルターを適用するメソッドは？",
        choices: [
            "Range.Filter",
            "Range.AutoFilter",
            "Worksheet.SetFilter"
        ],
        correct: 1,
        explanation: "`Range.AutoFilter` メソッドを使用してオートフィルターを適用します。例: `ActiveSheet.Range(\"A1:D10\").AutoFilter` これにより、指定された範囲にオートフィルターが適用されます。"
    },
    // 6-3-1 オートフィルターで特定のセルを残す
    {
        question: "オートフィルターを使用して特定の値のみを表示するコードは？",
        choices: [
            "ActiveSheet.Range(\"A1:D10\").AutoFilter Field:=1, Criteria1:=\"特定の値\"",
            "ActiveSheet.AutoFilter.Filter Field:=1, Value:=\"特定の値\"",
            "ActiveSheet.Range(\"A1:D10\").Filter Value:=\"特定の値\", Column:=1"
        ],
        correct: 0,
        explanation: "`ActiveSheet.Range(\"A1:D10\").AutoFilter Field:=1, Criteria1:=\"特定の値\"` が正しいコードです。これにより、1列目（Field:=1）で「特定の値」に一致するデータのみが表示されます。"
    },
    // 6-3-2 オートフィルターで絞り込む
    {
        question: "オートフィルターで複数の条件を使用して絞り込むコードは？",
        choices: [
            "ActiveSheet.Range(\"A1:D10\").AutoFilter Field:=1, Criteria1:=\">=100\", Operator:=xlAnd, Criteria2:=\"<=200\"",
            "ActiveSheet.Range(\"A1:D10\").AutoFilter Field:=1, Criteria1:=\">=100 AND <=200\"",
            "ActiveSheet.Range(\"A1:D10\").AutoFilter Field:=1, Criteria1:=Array(\">=100\", \"<=200\"), Operator:=xlFilterValues"
        ],
        correct: 0,
        explanation: "`ActiveSheet.Range(\"A1:D10\").AutoFilter Field:=1, Criteria1:=\">=100\", Operator:=xlAnd, Criteria2:=\"<=200\"` が正しいコードです。これにより、1列目の値が100以上200以下のデータのみが表示されます。"
    },
    // 6-3-3 絞り込んだ結果をコピーする
    {
        question: "オートフィルターで絞り込んだ結果を別のワークシートにコピーするコードは？",
        choices: [
            "ActiveSheet.AutoFilter.Range.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "ActiveSheet.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "ActiveSheet.FilteredRange.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")"
        ],
        correct: 1,
        explanation: "`ActiveSheet.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")` が正しいコードです。これにより、フィルター適用後に表示されているセルのみが新しいワークシートにコピーされます。"
    },
    // 6-3-4 絞り込んだ結果をカウントする
    {
        question: "オートフィルターで絞り込んだ結果の行数をカウントするコードは？",
        choices: [
            "FilteredCount = ActiveSheet.AutoFilter.Range.Rows.Count - 1",
            "FilteredCount = ActiveSheet.UsedRange.SpecialCells(xlCellTypeVisible).Rows.Count - 1",
            "FilteredCount = Application.WorksheetFunction.Subtotal(3, ActiveSheet.UsedRange)"
        ],
        correct: 1,
        explanation: "`FilteredCount = ActiveSheet.UsedRange.SpecialCells(xlCellTypeVisible).Rows.Count - 1` が正しいコードです。これにより、フィルター適用後に表示されている行数をカウントします（ヘッダー行を除くため -1 しています）。"
    },
    // 6-3-5 絞り込んだ結果の列を編集する
    {
        question: "オートフィルターで絞り込んだ結果の特定の列に値を設定するコードは？",
        choices: [
            "ActiveSheet.AutoFilter.Range.Columns(2).Value = \"新しい値\"",
            "ActiveSheet.UsedRange.SpecialCells(xlCellTypeVisible).Columns(2).Value = \"新しい値\"",
            "For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeVisible).Columns(2): cell.Value = \"新しい値\": Next cell"
        ],
        correct: 2,
        explanation: "`For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeVisible).Columns(2): cell.Value = \"新しい値\": Next cell` が正しいコードです。これにより、フィルター適用後に表示されている2列目のセルすべてに新しい値が設定されます。直接 Value プロパティを設定すると、非表示のセルも変更される可能性があるため、ループを使用するのが安全です。"
    },

    // 7 データの並べ替え
    {
        question: "Excel VBAでデータを並べ替える主なメソッドは何ですか？",
        choices: [
            "Range.Sort メソッド",
            "Worksheet.SortData メソッド",
            "Application.Sort メソッド"
        ],
        correct: 0,
        explanation: "`Range.Sort` メソッドがExcel VBAでデータを並べ替える主なメソッドです。このメソッドは、指定された範囲のデータを様々な条件に基づいて並べ替えることができます。"
    },
    // 7-1 Excel 2007以降の並べ替え
    {
        question: "Excel 2007以降で、複数の列を基準に並べ替えを行う場合、どのようにコードを記述しますか？",
        choices: [
            "Range.Sort Key1:=Range(\"A1\"), Key2:=Range(\"B1\"), Key3:=Range(\"C1\")",
            "Range.Sort SortFields.Add Key:=Range(\"A1\"), SortFields.Add Key:=Range(\"B1\")",
            "Range.Sort SortField1:=Range(\"A1\"), SortField2:=Range(\"B1\"), SortField3:=Range(\"C1\")"
        ],
        correct: 1,
        explanation: "Excel 2007以降では、`SortFields.Add` メソッドを使用して複数の並べ替え条件を追加します。例: `With ActiveSheet.Sort: .SortFields.Add Key:=Range(\"A1\"): .SortFields.Add Key:=Range(\"B1\"): .SetRange Range(\"A1:C10\"): .Apply: End With`"
    },
    // 7-1-1 難しくなった並べ替え
    {
        question: "Excel 2007以降で並べ替えが難しくなった主な理由は何ですか？",
        choices: [
            "並べ替えの速度が遅くなった",
            "並べ替えの条件設定が複雑になった",
            "並べ替えの機能が削減された"
        ],
        correct: 1,
        explanation: "Excel 2007以降では、より柔軟で強力な並べ替え機能が導入されましたが、それに伴いコードの記述が複雑になりました。特に、複数の条件による並べ替えやカスタム並べ替えリストの使用などが、以前のバージョンよりも複雑になっています。"
    },
    // 7-1-2 並べ替えの条件を指定する
    {
        question: "Excel 2007以降で、並べ替えの順序を降順に指定するコードは？",
        choices: [
            ".SortFields.Add Key:=Range(\"A1\"), SortOn:=xlSortOnValues, Order:=xlDescending",
            ".SortFields.Add Key:=Range(\"A1\"), SortOn:=xlSortOnValues, Order:=xlAscending",
            ".SortFields.Add Key:=Range(\"A1\"), SortOn:=xlSortOnValues, Order:=xlDescend"
        ],
        correct: 0,
        explanation: "正しいコードは `.SortFields.Add Key:=Range(\"A1\"), SortOn:=xlSortOnValues, Order:=xlDescending` です。`Order:=xlDescending` を指定することで、降順での並べ替えを行います。"
    },
    // 7-1-3 並べ替えの挙動を指定して実行する
    {
        question: "Excel 2007以降で、ヘッダー行を含む範囲を並べ替える際の正しいコードは？",
        choices: [
            "With ActiveSheet.Sort: .SetRange Range(\"A1:C10\"): .Header = xlYes: .Apply: End With",
            "With ActiveSheet.Sort: .SetRange Range(\"A1:C10\"): .Header = xlNo: .Apply: End With",
            "With ActiveSheet.Sort: .SetRange Range(\"A1:C10\"): .HeaderRow = True: .Apply: End With"
        ],
        correct: 0,
        explanation: "正しいコードは `With ActiveSheet.Sort: .SetRange Range(\"A1:C10\"): .Header = xlYes: .Apply: End With` です。`.Header = xlYes` を指定することで、1行目をヘッダーとして扱い、並べ替えから除外します。"
    },
    // 7-2 Excel 2003までの並べ替え
    {
        question: "Excel 2003以前のバージョンで、基本的な並べ替えを行うコードは？",
        choices: [
            "Range(\"A1:C10\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending",
            "Range(\"A1:C10\").SortAscending Key:=Range(\"A1\")",
            "Range(\"A1:C10\").Sort SortFields.Add Key:=Range(\"A1\")"
        ],
        correct: 0,
        explanation: "Excel 2003以前では、`Range(\"A1:C10\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending` のように、直接 Sort メソッドを使用し、Key1, Order1 などのパラメータで並べ替え条件を指定します。"
    },
    // 7-2-1 セルのSortメソッド
    {
        question: "Excel 2003以前で、複数の列を基準に並べ替えを行う場合のコードは？",
        choices: [
            "Range(\"A1:C10\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending, Key2:=Range(\"B1\"), Order2:=xlDescending",
            "Range(\"A1:C10\").Sort SortField1:=Range(\"A1\"), SortField2:=Range(\"B1\")",
            "Range(\"A1:C10\").SortMultiple Key1:=Range(\"A1\"), Key2:=Range(\"B1\")"
        ],
        correct: 0,
        explanation: "Excel 2003以前では、`Range(\"A1:C10\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending, Key2:=Range(\"B1\"), Order2:=xlDescending` のように、複数のKeyとOrderを指定して並べ替えを行います。"
    },
    // 7-2-2 漢字を並べ替えるときの注意
    {
        question: "Excel 2003以前で日本語の漢字を正しく並べ替えるために必要な設定は？",
        choices: [
            "Sort メソッドの Orient パラメータを xlSortOrientation に設定する",
            "Sort メソッドの Method パラメータを xlStroke に設定する",
            "Sort メソッドの IgnoreKana パラメータを False に設定する"
        ],
        correct: 1,
        explanation: "日本語の漢字を正しく並べ替えるには、`Sort` メソッドの `Method` パラメータを `xlStroke` に設定します。これにより、漢字の画数に基づいて並べ替えが行われます。例: `Range(\"A1:A10\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending, Method:=xlStroke`"
    },
    // 7-2-3 ふりがなの操作
    {
        question: "Excel VBAでセルのふりがな（ルビ）を取得するプロパティは？",
        choices: [
            "Cell.Phonetic",
            "Cell.Furigana",
            "Cell.Ruby"
        ],
        correct: 0,
        explanation: "セルのふりがな（ルビ）を取得するには `Cell.Phonetic` プロパティを使用します。例えば、`ActiveCell.Phonetic` でアクティブセルのふりがなを取得できます。ふりがなを基準に並べ替えを行う場合、このプロパティを使用して並べ替えのキーを設定します。"
    },
    // Additional questions for deeper understanding
    {
        question: "Excel 2007以降で、カスタムリストを使用して並べ替えを行うコードは？",
        choices: [
            ".SortFields.Add Key:=Range(\"A1\"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=Array(\"月曜\", \"火曜\", \"水曜\")",
            ".SortFields.Add Key:=Range(\"A1\"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=\"月曜,火曜,水曜\"",
            ".SortFields.Add Key:=Range(\"A1\"), SortOn:=xlSortOnValues, Order:=xlCustom, CustomOrder:=Array(\"月曜\", \"火曜\", \"水曜\")"
        ],
        correct: 1,
        explanation: "カスタムリストを使用して並べ替えを行う場合、`.SortFields.Add Key:=Range(\"A1\"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=\"月曜,火曜,水曜\"` のように、CustomOrder パラメータにカンマ区切りの文字列でリストを指定します。"
    },
    {
        question: "Excel VBAで並べ替えを行った後、元の行番号を保持するための最良の方法は？",
        choices: [
            "並べ替え前に新しい列を追加し、行番号を入力する",
            "並べ替え後に行番号を再計算する関数を使用する",
            "並べ替えの代わりに AutoFilter を使用する"
        ],
        correct: 0,
        explanation: "元の行番号を保持する最良の方法は、並べ替え前に新しい列を追加し、そこに行番号を入力することです。例えば、`Range(\"A1:A\" & LastRow).Value = Application.Transpose(Array(1, 2, 3))` のようなコードで連番を入力し、その後でデータと共に並べ替えを行います。これにより、並べ替え後も元の順序を追跡できます。"
    },

    // 8 テーブルの操作
    {
        question: "Excel VBAでテーブルを操作する際に使用するオブジェクトは何ですか？",
        choices: [
            "Table オブジェクト",
            "ListObject オブジェクト",
            "TableObject オブジェクト"
        ],
        correct: 1,
        explanation: "Excel VBAでテーブルを操作する際は、`ListObject` オブジェクトを使用します。これは、Excelのテーブル機能を表すオブジェクトで、テーブルの様々なプロパティやメソッドにアクセスできます。"
    },
    // 8-1 テーブルを特定する
    {
        question: "ワークシート内のすべてのテーブルにアクセスするためのプロパティは何ですか？",
        choices: [
            "Worksheet.Tables",
            "Worksheet.ListObjects",
            "Worksheet.TableObjects"
        ],
        correct: 1,
        explanation: "`Worksheet.ListObjects` プロパティを使用して、ワークシート内のすべてのテーブルにアクセスできます。例えば、`For Each tbl In ActiveSheet.ListObjects` のようにして、シート内の各テーブルに対して処理を行うことができます。"
    },
    // 8-1-1 テーブルのセルから特定する
    {
        question: "特定のセルが属するテーブルを取得するコードは次のうちどれですか？",
        choices: [
            "ActiveCell.Table",
            "ActiveCell.ListObject",
            "ActiveCell.CurrentRegion.ListObject"
        ],
        correct: 1,
        explanation: "`ActiveCell.ListObject` を使用して、アクティブセルが属するテーブルを取得できます。セルがテーブルの一部でない場合は Nothing が返されます。例: `If Not ActiveCell.ListObject Is Nothing Then ... `"
    },
    // 8-1-2 テーブルが存在するシートから特定する
    {
        question: "シート内の最初のテーブルを取得するコードは次のうちどれですか？",
        choices: [
            "ActiveSheet.Tables(1)",
            "ActiveSheet.ListObjects(1)",
            "ActiveSheet.ListObjects.Item(1)"
        ],
        correct: 1,
        explanation: "`ActiveSheet.ListObjects(1)` または `ActiveSheet.ListObjects.Item(1)` を使用して、シート内の最初のテーブルを取得できます。インデックスは1から始まります。テーブルが存在しない場合はエラーが発生するため、通常は存在チェックと組み合わせて使用します。"
    },
    // 8-1-3 Rangeとテーブルの名前で特定する
    {
        question: "テーブル名を使用してテーブルを取得するコードは次のうちどれですか？",
        choices: [
            "ActiveSheet.Tables(\"TableName\")",
            "ActiveSheet.ListObjects(\"TableName\")",
            "Range(\"TableName\")"
        ],
        correct: 1,
        explanation: "`ActiveSheet.ListObjects(\"TableName\")` を使用して、名前を指定してテーブルを取得できます。また、`Range(\"TableName\")` を使用してテーブル全体の範囲を取得することもできますが、これは ListObject オブジェクトではなく Range オブジェクトを返します。"
    },
    // 8-2 テーブルの部位を特定する
    {
        question: "テーブルのデータ部分（ヘッダーを除く）を取得するプロパティは次のうちどれですか？",
        choices: [
            "ListObject.DataBodyRange",
            "ListObject.DataRange",
            "ListObject.InsideRange"
        ],
        correct: 0,
        explanation: "`ListObject.DataBodyRange` プロパティを使用して、テーブルのデータ部分（ヘッダーを除く）を取得できます。これはヘッダー行を含まないテーブルのデータ全体を表す Range オブジェクトを返します。"
    },
    // 8-2-1 見出し（タイトル）行を含むテーブル全体
    {
        question: "見出し（タイトル）行を含むテーブル全体を取得するプロパティは次のうちどれですか？",
        choices: [
            "ListObject.Range",
            "ListObject.EntireTable",
            "ListObject.WholeTable"
        ],
        correct: 0,
        explanation: "`ListObject.Range` プロパティを使用して、見出し（タイトル）行を含むテーブル全体を取得できます。このプロパティは、テーブル全体（ヘッダーを含む）を表す Range オブジェクトを返します。"
    },
    // 8-2-2 見出し（タイトル）行を含まないテーブルのデータ全体
    {
        question: "見出し（タイトル）行を含まないテーブルのデータ全体を取得するプロパティは次のうちどれですか？",
        choices: [
            "ListObject.DataBodyRange",
            "ListObject.DataRange",
            "ListObject.InsideRange"
        ],
        correct: 0,
        explanation: "この質問は 8-2 の質問と同じですが、重要なので再度確認します。`ListObject.DataBodyRange` プロパティを使用して、見出し（タイトル）行を含まないテーブルのデータ全体を取得できます。"
    },
    // 8-2-3 見出し（タイトル）行
    {
        question: "テーブルの見出し（タイトル）行を取得するプロパティは次のうちどれですか？",
        choices: [
            "ListObject.HeaderRow",
            "ListObject.HeaderRange",
            "ListObject.TableTitle"
        ],
        correct: 1,
        explanation: "`ListObject.HeaderRange` プロパティを使用して、テーブルの見出し（タイトル）行を取得できます。このプロパティは、テーブルのヘッダー行を表す Range オブジェクトを返します。"
    },
    // 8-2-4 列
    {
        question: "テーブルの特定の列を取得するメソッドは次のうちどれですか？",
        choices: [
            "ListObject.ListColumns(\"ColumnName\")",
            "ListObject.Columns(\"ColumnName\")",
            "ListObject.TableColumns(\"ColumnName\")"
        ],
        correct: 0,
        explanation: "`ListObject.ListColumns(\"ColumnName\")` を使用して、テーブルの特定の列を取得できます。これは ListColumn オブジェクトを返します。列名または列のインデックスを使用して特定の列にアクセスできます。"
    },
    // 8-2-5 行
    {
        question: "テーブルの特定の行を取得するプロパティは次のうちどれですか？",
        choices: [
            "ListObject.ListRows(index)",
            "ListObject.DataBodyRange.Rows(index)",
            "ListObject.TableRows(index)"
        ],
        correct: 0,
        explanation: "`ListObject.ListRows(index)` を使用して、テーブルの特定の行を取得できます。これは ListRow オブジェクトを返します。インデックスは1から始まり、ヘッダー行は含まれません。"
    },

    // 8-3 構造化参照を使って特定する
    {
        question: "VBAで構造化参照を使用してテーブル全体を参照する正しい方法は？",
        choices: [
            "Range(\"TableName[#All]\")",
            "Range(\"TableName[@]\")",
            "Range(\"TableName[#Data]\")"
        ],
        correct: 0,
        explanation: "`Range(\"TableName[#All]\")` を使用して、構造化参照でテーブル全体（ヘッダーを含む）を参照できます。これは `ListObject.Range` と同等の結果を返します。"
    },
    // 8-3-1 見出し（タイトル）行を含むテーブル全体
    {
        question: "構造化参照を使用してテーブル全体（ヘッダーを含む）を取得するコードは？",
        choices: [
            "Range(\"TableName[#All]\")",
            "Range(\"TableName[#Headers],[#Data]\")",
            "Range(\"TableName[#Everything]\")"
        ],
        correct: 0,
        explanation: "この質問は前問と重複しますが、重要なので再確認します。`Range(\"TableName[#All]\")` を使用して、構造化参照でテーブル全体（ヘッダーを含む）を参照できます。"
    },
    // 8-3-2 見出し（タイトル）行を含まないテーブルのデータ全体
    {
        question: "構造化参照を使用してテーブルのデータ部分（ヘッダーを除く）を取得するコードは？",
        choices: [
            "Range(\"TableName[#Data]\")",
            "Range(\"TableName[@]\")",
            "Range(\"TableName[#Body]\")"
        ],
        correct: 0,
        explanation: "`Range(\"TableName[#Data]\")` を使用して、構造化参照でテーブルのデータ部分（ヘッダーを除く）を参照できます。これは `ListObject.DataBodyRange` と同等の結果を返します。"
    },
    // 8-3-3 列
    {
        question: "構造化参照を使用してテーブルの特定の列を取得するコードは？",
        choices: [
            "Range(\"TableName[ColumnName]\")",
            "Range(\"TableName[[ColumnName]]\")",
            "Range(\"TableName[@ColumnName]\")"
        ],
        correct: 0,
        explanation: "`Range(\"TableName[ColumnName]\")` を使用して、構造化参照でテーブルの特定の列を参照できます。これはヘッダーを含む列全体を返します。"
    },
    // 8-3-4 行
    {
        question: "構造化参照を使用してテーブルの現在の行を取得するコードは？",
        choices: [
            "Range(\"TableName[@]\")",
            "Range(\"TableName[#ThisRow]\")",
            "Range(\"TableName[#CurrentRow]\")"
        ],
        correct: 0,
        explanation: "`Range(\"TableName[@]\")` を使用して、構造化参照でテーブルの現在の行を参照できます。これは現在のレコードを表し、通常はアクティブセルが含まれる行を指します。"
    },
    // 8-4 特定のデータを操作する
    // 8-4-1 テーブル内のデータを探す
    {
        question: "VBAでテーブル内の特定のデータを検索する正しいコードは？",
        choices: [
            "ListObject.DataBodyRange.Find(What:=\"検索値\")",
            "ListObject.ListRows.Find(What:=\"検索値\")",
            "ListObject.Range.Find(What:=\"検索値\")"
        ],
        correct: 0,
        explanation: "`ListObject.DataBodyRange.Find(What:=\"検索値\")` を使用して、テーブル内の特定のデータを検索できます。これはテーブルのデータ部分（ヘッダーを除く）から検索を行います。"
    },
    // 8-4-2 見出し行ごとコピーする
    {
        question: "テーブル全体（ヘッダーを含む）を別のワークシートにコピーするコードは？",
        choices: [
            "ListObject.Range.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "ListObject.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "Range(\"TableName[#All]\").Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")"
        ],
        correct: 0,
        explanation: "`ListObject.Range.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")` を使用して、テーブル全体（ヘッダーを含む）を別のワークシートにコピーできます。構造化参照を使用する場合は、3番目の選択肢も正解となります。"
    },
    // 8-4-3 見出し行を含まないデータだけをコピーする
    {
        question: "テーブルのデータ部分（ヘッダーを除く）を別のワークシートにコピーするコードは？",
        choices: [
            "ListObject.DataBodyRange.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "ListObject.ListRows.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "Range(\"TableName[#Data]\").Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")"
        ],
        correct: 0,
        explanation: "`ListObject.DataBodyRange.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")` を使用して、テーブルのデータ部分（ヘッダーを除く）を別のワークシートにコピーできます。構造化参照を使用する場合は、3番目の選択肢も正解となります。"
    },
    // 8-4-4 Rangeと構造化参照を使ってコピーする
    {
        question: "構造化参照を使用してテーブルのデータ部分をコピーするコードは？",
        choices: [
            "Range(\"TableName[#Data]\").Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "Range(\"TableName[@]\").Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "Range(\"TableName\").Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")"
        ],
        correct: 0,
        explanation: "`Range(\"TableName[#Data]\").Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")` を使用して、構造化参照でテーブルのデータ部分をコピーできます。これは `ListObject.DataBodyRange.Copy` と同等の結果を返します。"
    },
    // 8-4-5 特定の列だけコピーする
    {
        question: "テーブルの特定の列をコピーするコードは？",
        choices: [
            "ListObject.ListColumns(\"ColumnName\").Range.Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "Range(\"TableName[ColumnName]\").Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")",
            "ListObject.Range.Columns(\"ColumnName\").Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")"
        ],
        correct: 1,
        explanation: "`Range(\"TableName[ColumnName]\").Copy Destination:=Worksheets(\"Sheet2\").Range(\"A1\")` を使用して、テーブルの特定の列をコピーできます。また、1番目の選択肢も正解となります。構造化参照を使用する方が、より直感的でエラーが少ないでしょう。"
    },
    // 8-4-6 特定の列だけ書式を設定する
    {
        question: "テーブルの特定の列の書式を設定するコードは？",
        choices: [
            "ListObject.ListColumns(\"ColumnName\").Range.Font.Bold = True",
            "Range(\"TableName[ColumnName]\").Font.Bold = True",
            "ListObject.Range.Columns(\"ColumnName\").Font.Bold = True"
        ],
        correct: 1,
        explanation: "`Range(\"TableName[ColumnName]\").Font.Bold = True` を使用して、テーブルの特定の列の書式を設定できます。また、1番目の選択肢も正解となります。構造化参照を使用する方が、より直感的でエラーが少ないでしょう。"
    },

    // 8-5 行を削除する
    {
        question: "VBAでテーブルの特定の行を削除する正しいコードは？",
        choices: [
            "ListObject.ListRows(1).Delete",
            "ListObject.DataBodyRange.Rows(1).Delete",
            "Range(\"TableName[@]\").Delete"
        ],
        correct: 0,
        explanation: "`ListObject.ListRows(1).Delete` を使用して、テーブルの特定の行（この場合は1行目）を削除できます。ListRows コレクションは1から始まり、ヘッダー行は含まれません。"
    },
    // 8-5-1 テーブルの行全体を削除する
    {
        question: "テーブルの全ての行（データ）を削除するコードは？",
        choices: [
            "ListObject.DataBodyRange.Delete",
            "ListObject.ListRows.Delete",
            "For i = ListObject.ListRows.Count To 1 Step -1: ListObject.ListRows(i).Delete: Next i"
        ],
        correct: 2,
        explanation: "テーブルの全ての行を削除するには、`For i = ListObject.ListRows.Count To 1 Step -1: ListObject.ListRows(i).Delete: Next i` のようなループを使用します。後ろから削除することで、インデックスのずれを防ぎます。単純に `DataBodyRange.Delete` を使用すると、テーブル構造が壊れる可能性があります。"
    },
    // 8-5-2 Rangeと構造化参照を使って削除する
    {
        question: "構造化参照を使用してテーブルの特定の行を削除するコードは？",
        choices: [
            "Range(\"TableName[@]\").Delete",
            "Range(\"TableName[#ThisRow]\").Delete",
            "Range(\"TableName[#Data]\").Rows(1).Delete"
        ],
        correct: 0,
        explanation: "`Range(\"TableName[@]\").Delete` を使用して、構造化参照で現在の行（アクティブセルを含む行）を削除できます。ただし、この方法はアクティブセルの位置に依存するため、注意が必要です。"
    },
    // 8-6 列を挿入する
    {
        question: "VBAでテーブルに新しい列を挿入する正しいコードは？",
        choices: [
            "ListObject.ListColumns.Add",
            "ListObject.Columns.Insert",
            "ListObject.Range.Columns.Insert"
        ],
        correct: 0,
        explanation: "`ListObject.ListColumns.Add` を使用して、テーブルに新しい列を挿入できます。デフォルトでは、新しい列は最後に追加されます。特定の位置に挿入する場合は、位置を指定するパラメータを追加できます。"
    },
    // 8-6-1 テーブルに列を挿入する
    {
        question: "テーブルの特定の位置に新しい列を挿入するコードは？",
        choices: [
            "ListObject.ListColumns.Add(2)",
            "ListObject.ListColumns.Insert(2)",
            "ListObject.ListColumns.Add Position:=2"
        ],
        correct: 2,
        explanation: "`ListObject.ListColumns.Add Position:=2` を使用して、テーブルの特定の位置（この場合は2列目）に新しい列を挿入できます。Position パラメータを使用することで、挿入位置を明示的に指定できます。"
    },
    // 8-6-2 Rangeと構造化参照を使って列を挿入する
    {
        question: "構造化参照を使用してテーブルに新しい列を挿入するコードは？",
        choices: [
            "Range(\"TableName[ColumnName]\").Insert Shift:=xlToRight",
            "Range(\"TableName[[ColumnName]]\").EntireColumn.Insert",
            "Range(\"TableName[#Headers]\").Columns(1).Insert Shift:=xlToRight"
        ],
        correct: 0,
        explanation: "`Range(\"TableName[ColumnName]\").Insert Shift:=xlToRight` を使用して、構造化参照で特定の列の左に新しい列を挿入できます。この方法では、テーブル構造が自動的に更新され、新しい列がテーブルの一部として認識されます。"
    },
    // Additional questions for deeper understanding
    {
        question: "テーブルの特定の列を削除するコードは？",
        choices: [
            "ListObject.ListColumns(\"ColumnName\").Delete",
            "Range(\"TableName[ColumnName]\").Delete Shift:=xlToLeft",
            "ListObject.Range.Columns(\"ColumnName\").Delete"
        ],
        correct: 0,
        explanation: "`ListObject.ListColumns(\"ColumnName\").Delete` を使用して、テーブルの特定の列を削除できます。この方法では、テーブル構造が適切に更新されます。構造化参照を使用する場合は2番目の選択肢も可能ですが、ListColumns を使用する方がより直接的です。"
    },
    {
        question: "テーブルに複数の列を一度に追加するコードは？",
        choices: [
            "ListObject.ListColumns.Add(1, 3)",
            "For i = 1 To 3: ListObject.ListColumns.Add: Next i",
            "ListObject.Range.Columns.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove, Count:=3"
        ],
        correct: 1,
        explanation: "複数の列を一度に追加するには、ループを使用するのが最も確実です。`For i = 1 To 3: ListObject.ListColumns.Add: Next i` のようなコードで、3つの新しい列をテーブルの末尾に追加できます。ListColumns.Add メソッドには、一度に複数の列を追加するオプションがないため、このアプローチが必要です。"
    },


    // 9-1 エラーの種類
    {
        question: "VBAにおける主なエラーの種類は次のうちどれですか？",
        choices: [
            "記述エラー、論理エラー、実行時エラー",
            "構文エラー、コンパイルエラー、ランタイムエラー",
            "タイプエラー、参照エラー、演算エラー"
        ],
        correct: 0,
        explanation: "VBAにおける主なエラーの種類は、記述エラー（構文エラーとも呼ばれる）、論理エラー、実行時エラー（ランタイムエラー）です。これらは、コードの記述時、ロジックの誤り、実行中に発生する問題をそれぞれ表します。"
    },
    // 9-1-1 記述エラー
    {
        question: "次のうち、記述エラー（構文エラー）の例はどれですか？",
        choices: [
            "For i = 1 To 10 Next i",
            "If x > 0 Then y = 2",
            "Dim x As Integer: x = \"Hello\""
        ],
        correct: 0,
        explanation: "`For i = 1 To 10 Next i` は記述エラーの例です。正しくは `For i = 1 To 10: ... : Next i` となります。記述エラーは、VBAの構文規則に違反している場合に発生し、通常はコードを実行する前に検出されます。"
    },
    // 9-1-2 論理エラー
    {
        question: "次のコードの論理エラーを指摘してください：\n`If x > 0 Then\n    y = 2\nElseIf x > 10 Then\n    y = 3\nEnd If`",
        choices: [
            "x が 0 より大きく 10 以下の場合、y が 2 になる",
            "x が 10 より大きい場合、y が 3 になる",
            "x が 10 より大きい場合、ElseIf 文が実行されない"
        ],
        correct: 2,
        explanation: "このコードの論理エラーは、x が 10 より大きい場合でも ElseIf 文が実行されないことです。最初の If 条件 (x > 0) が真の場合、ElseIf 部分は評価されません。正しくは、条件を `If x > 10 Then` と `ElseIf x > 0 Then` の順に並べ替えるべきです。"
    },
    // 9-1-3 実行時エラー
    {
        question: "次のコードを実行した場合、どのような実行時エラーが発生する可能性がありますか？\n`Dim x As Integer\nx = 1000000`",
        choices: [
            "オーバーフローエラー",
            "タイプミスマッチエラー",
            "配列の境界エラー"
        ],
        correct: 0,
        explanation: "このコードではオーバーフローエラーが発生する可能性があります。Integer型の最大値は32,767であり、1,000,000はこの範囲を超えています。実行時エラーは、コードの実行中に発生するエラーで、このような値の範囲の問題や、ゼロによる除算、存在しないオブジェクトの参照などが原因で発生します。"
    },
    // 9-2 エラーへの対応
    // 9-2-1 エラーが発生したら別の処理にジャンプする
    {
        question: "エラーが発生した場合に特定のラベルにジャンプするVBAコードは？",
        choices: [
            "On Error GoTo ErrorHandler",
            "If Error Then GoTo ErrorHandler",
            "Try ... Catch GoTo ErrorHandler"
        ],
        correct: 0,
        explanation: "`On Error GoTo ErrorHandler` を使用して、エラーが発生した場合に特定のラベル（この場合は ErrorHandler）にジャンプするよう指定できます。このステートメントは通常、プロシージャの先頭に配置されます。"
    },
    // 実用的な事例
    {
        question: "次のコードのエラー処理を改善してください：\n`Sub DivideNumbers()\n    Dim x As Integer, y As Integer\n    x = 10\n    y = 0\n    Debug.Print x / y\nEnd Sub`",
        choices: [
            "On Error Resume Next を追加する",
            "On Error GoTo ErrHandler を追加し、エラーハンドラを実装する",
            "Try ... Catch ブロックを使用する"
        ],
        correct: 1,
        explanation: "正しい改善方法は、`On Error GoTo ErrHandler` を追加し、エラーハンドラを実装することです。例：\n`Sub DivideNumbers()\n    On Error GoTo ErrHandler\n    Dim x As Integer, y As Integer\n    x = 10\n    y = 0\n    Debug.Print x / y\n    Exit Sub\nErrHandler:\n    MsgBox \"エラーが発生しました: \" & Err.Description\nEnd Sub`\nこれにより、ゼロ除算エラーを適切に処理できます。"
    },
    // 9-2-2 どんなエラーが発生したか調べる
    {
        question: "VBAでエラーの詳細情報を取得するために使用するオブジェクトは？",
        choices: [
            "Err オブジェクト",
            "Error オブジェクト",
            "Exception オブジェクト"
        ],
        correct: 0,
        explanation: "VBAでは `Err` オブジェクトを使用してエラーの詳細情報を取得します。主なプロパティには、Number（エラー番号）、Description（エラーの説明）、Source（エラーの発生源）があります。"
    },
    // 実用的な事例
    {
        question: "次のコードでエラーが発生した場合、どのようにエラー情報を表示しますか？\n`Sub OpenFile()\n    Open \"C:\\nonexistent.txt\" For Input As #1\nEnd Sub`",
        choices: [
            "MsgBox \"エラーが発生しました\"",
            "Debug.Print Error.Description",
            "MsgBox \"エラー \" & Err.Number & \": \" & Err.Description"
        ],
        correct: 2,
        explanation: "正しいエラー情報の表示方法は `MsgBox \"エラー \" & Err.Number & \": \" & Err.Description` です。これにより、エラー番号とその説明が表示されます。完全なコードは次のようになります：\n`Sub OpenFile()\n    On Error GoTo ErrHandler\n    Open \"C:\\nonexistent.txt\" For Input As #1\n    Exit Sub\nErrHandler:\n    MsgBox \"エラー \" & Err.Number & \": \" & Err.Description\nEnd Sub`"
    },
    // 9-2-3 発生したエラーを無視する
    {
        question: "VBAで特定の行でエラーを無視するためのステートメントは？",
        choices: [
            "On Error Resume Next",
            "Ignore Error",
            "Try ... Catch"
        ],
        correct: 0,
        explanation: "`On Error Resume Next` を使用すると、エラーが発生しても次の行に進みます。ただし、この使用には注意が必要で、エラーを適切に処理せずに無視することは望ましくありません。"
    },
    // 実用的な事例
    {
        question: "次のコードでファイルが存在しない場合のエラーを適切に処理するには？\n`Sub DeleteFile(filePath As String)\n    Kill filePath\nEnd Sub`",
        choices: [
            "On Error Resume Next を使用する",
            "On Error GoTo ErrHandler を使用し、ファイルが存在しない場合はメッセージを表示する",
            "Try ... Catch ブロックを使用する"
        ],
        correct: 1,
        explanation: "適切な処理方法は `On Error GoTo ErrHandler` を使用し、ファイルが存在しない場合はメッセージを表示することです。例：\n`Sub DeleteFile(filePath As String)\n    On Error GoTo ErrHandler\n    Kill filePath\n    Exit Sub\nErrHandler:\n    If Err.Number = 53 Then ' ファイルが見つからない\n        MsgBox \"ファイルが存在しません: \" & filePath\n    Else\n        MsgBox \"エラー \" & Err.Number & \": \" & Err.Description\n    End If\nEnd Sub`"
    },
    // 9-2-4 エラー対策のポイント
    {
        question: "VBAでのエラー対策の重要なポイントは次のうちどれですか？",
        choices: [
            "すべてのエラーを無視する",
            "予期されるエラーを事前にチェックし、適切に処理する",
            "常に On Error Resume Next を使用する"
        ],
        correct: 1,
        explanation: "エラー対策の重要なポイントは、予期されるエラーを事前にチェックし、適切に処理することです。例えば、ファイル操作の前にファイルの存在をチェックしたり、ゼロ除算を避けるために除数が0でないことを確認したりします。適切なエラー処理により、プログラムの信頼性と使いやすさが向上します。"
    },
    // 9-3 データのクレンジング
    // 9-3-1 不正なデータを修正する
    {
        question: "セル内の数値データに文字列が混じっている場合、どのように処理するのが適切ですか？",
        choices: [
            "エラーとして処理し、ユーザーに再入力を求める",
            "Val() 関数を使用して数値部分のみを抽出する",
            "常に0として扱う"
        ],
        correct: 1,
        explanation: "セル内の数値データに文字列が混じっている場合、`Val()` 関数を使用して数値部分のみを抽出するのが適切です。例：`CleanedValue = Val(Cells(1, 1).Value)`。これにより、「100円」のような文字列から「100」という数値を取得できます。ただし、データの性質によっては、エラーとして処理し、ユーザーに再入力を求めることも適切な場合があります。"
    },
    // 9-3-2 半角文字列と全角文字列
    {
        question: "VBAで全角文字を半角文字に変換する関数は？",
        choices: [
            "ConvertToHankaku()",
            "StrConv(文字列, vbNarrow)",
            "ToHalfWidth()"
        ],
        correct: 1,
        explanation: "`StrConv(文字列, vbNarrow)` を使用して、全角文字を半角文字に変換できます。逆に、半角から全角への変換には `StrConv(文字列, vbWide)` を使用します。例：`HalfWidthStr = StrConv(\"ＡＢＣＤ\", vbNarrow)` の結果は \"ABCD\" となります。"
    },
    // 9-3-3 不要な文字を除去する
    {
        question: "文字列から特定の文字を削除するVBA関数は？",
        choices: [
            "RemoveChar()",
            "Replace()",
            "Trim()"
        ],
        correct: 1,
        explanation: "`Replace()` 関数を使用して、文字列から特定の文字を削除できます。例：`CleanStr = Replace(OriginalStr, \"削除したい文字\", \"\")` このように、削除したい文字を空文字列に置き換えることで、その文字を効果的に削除できます。"
    },
    // 実用的な事例
    {
        question: "次のような不適切なISBN番号を正規化するVBA関数を作成してください：「ISBN: 978-4-12-345678-9 」",
        choices: [
            "Function NormalizeISBN(isbn As String) As String\n    NormalizeISBN = Replace(Replace(isbn, \"ISBN: \", \"\"), \"-\", \"\")\nEnd Function",
            "Function NormalizeISBN(isbn As String) As String\n    NormalizeISBN = Mid(isbn, 6, 13)\nEnd Function",
            "Function NormalizeISBN(isbn As String) As String\n    NormalizeISBN = Trim(isbn)\nEnd Function"
        ],
        correct: 0,
        explanation: "正しい関数は次のとおりです：\n`Function NormalizeISBN(isbn As String) As String\n    NormalizeISBN = Replace(Replace(isbn, \"ISBN: \", \"\"), \"-\", \"\")\nEnd Function`\nこの関数は、\"ISBN: \" という接頭辞を削除し、ハイフンを取り除きます。結果として「9784123456789」という正規化されたISBN番号が得られます。"
    },
    // 9-3-4 日付の操作
    {
        question: "VBAで文字列を日付型に変換する関数は？",
        choices: [
            "CDate()",
            "DateValue()",
            "ConvertToDate()"
        ],
        correct: 0,
        explanation: "`CDate()` 関数を使用して、文字列を日付型に変換できます。例：`Dim myDate As Date: myDate = CDate(\"2023/04/01\")` この関数は様々な日付形式を認識しますが、地域設定に依存する場合があるので注意が必要です。"
    },
    {
        question: "日本語の年号表記（例：「令和5年4月1日」）を西暦の日付型に変換する適切な方法は？",
        choices: [
            "CDate() 関数を直接使用する",
            "Replace() 関数で「令和」を削除してから CDate() を使用する",
            "カスタム関数を作成して変換する"
        ],
        correct: 2,
        explanation: "日本語の年号表記を西暦の日付型に変換するには、カスタム関数を作成するのが最適です。例：\n`Function JapaneseEraToDate(eraDate As String) As Date\n    Dim era As String, year As Integer, month As Integer, day As Integer\n    ' 年号、年、月、日を抽出するロジック\n    ' 年号に応じて西暦に変換\n    If era = \"令和\" Then\n        year = year + 2018\n    ElseIf era = \"平成\" Then\n        year = year + 1988\n    ' その他の年号も同様に処理\n    End If\n    JapaneseEraToDate = DateSerial(year, month, day)\nEnd Function`\nこのようなカスタム関数を使用することで、様々な日本の年号表記を適切に処理できます。"
    },
    // 実用的な事例
    {
        question: "次のようなセルの値を日付型に変換し、無効な日付の場合はエラーメッセージを表示する関数を作成してください：\nA1: 2023/4/1\nA2: 2023/13/1\nA3: 不明",
        choices: [
            "Function SafeDateConvert(cellValue As Variant) As Variant\n    On Error GoTo ErrHandler\n    SafeDateConvert = CDate(cellValue)\n    Exit Function\nErrHandler:\n    SafeDateConvert = \"無効な日付\"\nEnd Function",
            "Function SafeDateConvert(cellValue As Variant) As Variant\n    If IsDate(cellValue) Then\n        SafeDateConvert = CDate(cellValue)\n    Else\n        SafeDateConvert = \"無効な日付\"\n    End If\nEnd Function",
            "Function SafeDateConvert(cellValue As Variant) As Variant\n    SafeDateConvert = Format(cellValue, \"yyyy/mm/dd\")\nEnd Function"
        ],
        correct: 1,
        explanation: "正しい関数は次のとおりです：\n`Function SafeDateConvert(cellValue As Variant) As Variant\n    If IsDate(cellValue) Then\n        SafeDateConvert = CDate(cellValue)\n    Else\n        SafeDateConvert = \"無効な日付\"\n    End If\nEnd Function`\nこの関数は、まず `IsDate()` 関数を使用して値が有効な日付かどうかをチェックします。有効な場合は `CDate()` で変換し、無効な場合は「無効な日付」というメッセージを返します。これにより、A1は正しく日付型に変換され、A2とA3は「無効な日付」と表示されます。"
    },
    // エラー対策の総合的な理解を問う質問
    {
        question: "大量のデータを処理するVBAマクロを作成する際、エラー処理と性能の両方を考慮した最適なアプローチは？",
        choices: [
            "すべての処理で On Error Resume Next を使用し、エラーを無視する",
            "各処理ステップでエラーチェックを行い、問題があればすぐに処理を中断する",
            "予期されるエラーを事前にチェックし、バッチ処理でエラーログを作成しながら続行する"
        ],
        correct: 2,
        explanation: "大量のデータを処理する場合、予期されるエラーを事前にチェックし、バッチ処理でエラーログを作成しながら続行するアプローチが最適です。例：\n`Sub ProcessLargeData()\n    Dim errorLog As String\n    For Each cell In Range(\"A1:A1000\")\n        If Not IsNumeric(cell.Value) Then\n            errorLog = errorLog & \"行 \" & cell.Row & \": 数値でないデータ\" & vbNewLine\n        Else\n            ' データ処理\n        End If\n    Next cell\n    If Len(errorLog) > 0 Then\n        ' エラーログをファイルに保存または表示\n    End If\nEnd Sub`\nこのアプローチにより、処理を中断することなくエラーを記録し、後で一括して確認・修正することができます。また、`Application.ScreenUpdating = False` や `Application.Calculation = xlCalculationManual` などを使用して処理速度を向上させることも重要です。"
    },
    // データクレンジングの総合的な理解を問う質問
    {
        question: "顧客データベースのクレンジングを行うVBAマクロを作成する際、考慮すべき重要な点は何ですか？（複数選択可）",
        choices: [
            "全ての文字列データを大文字に統一する",
            "電話番号の形式を統一し、不要な文字（ハイフンやスペース）を除去する",
            "住所データの都道府県名を正規化する（例：「東京都」と「東京」を統一）",
            "氏名のミドルネームを削除する"
        ],
        correct: [1, 2],
        explanation: "顧客データベースのクレンジングでは、以下の点が特に重要です：\n1. 電話番号の形式統一と不要な文字の除去：これにより、検索や照合が容易になります。\n2. 住所データの正規化：都道府県名や市区町村名の表記を統一することで、データの一貫性が向上します。\n\n全ての文字列データを大文字に統一することは、場合によっては情報の損失につながる可能性があります。また、氏名のミドルネームの削除は、個人を特定する上で重要な情報を失う可能性があるため、通常は推奨されません。\n\n実装例：\n`Sub CleanCustomerData()\n    Dim cell As Range\n    For Each cell In Range(\"電話番号列\")\n        cell.Value = CleanPhoneNumber(cell.Value)\n    Next cell\n    For Each cell In Range(\"住所列\")\n        cell.Value = NormalizeAddress(cell.Value)\n    Next cell\nEnd Sub\n\nFunction CleanPhoneNumber(phone As String) As String\n    CleanPhoneNumber = Replace(Replace(phone, \"-\", \"\"), \" \", \"\")\nEnd Function\n\nFunction NormalizeAddress(address As String) As String\n    ' 都道府県名を正規化するロジック\n    ' 例: \"東京\" を \"東京都\" に変換\nEnd Function`"
    },

    // 10-1 デバッグとは
    {
        question: "VBAにおけるデバッグの主な目的は何ですか？",
        choices: [
            "コードの実行速度を向上させること",
            "プログラム内のエラーを特定し修正すること",
            "コードの行数を減らすこと"
        ],
        correct: 1,
        explanation: "デバッグの主な目的は、プログラム内のエラーを特定し修正することです。これには、構文エラー、ランタイムエラー、論理エラーなど、様々な種類のエラーの検出と修正が含まれます。デバッグは、プログラムが正しく動作することを確認するための重要なプロセスです。"
    },
    // 10-1-1 文法エラーと論理エラー
    {
        question: "次のうち、論理エラーの例はどれですか？",
        choices: [
            "For i = 1 To 10 Next i",
            "Dim x As Integer: x = \"Hello\"",
            "If totalAmount > 1000 Then discount = 0.05 Else discount = 0.1"
        ],
        correct: 2,
        explanation: "3番目の選択肢が論理エラーの例です。このコードは構文的には正しいですが、論理的に誤っています。通常、総額が大きいほど割引率が高くなるべきですが、このコードでは逆になっています。文法エラー（構文エラー）はコンパイル時に検出されますが、論理エラーは実行時の結果が意図したものと異なる場合に発見されます。"
    },
    // 10-2 イミディエイトウィンドウ
    {
        question: "VBAのイミディエイトウィンドウの主な用途は何ですか？",
        choices: [
            "コードを編集する",
            "変数の値を確認したり、簡単なコードを実行したりする",
            "エラーメッセージを表示する"
        ],
        correct: 1,
        explanation: "イミディエイトウィンドウの主な用途は、変数の値を確認したり、簡単なコードを実行したりすることです。デバッグ中に変数の状態を確認したり、小さなコード片をテストしたりするのに非常に便利なツールです。また、Debug.Print ステートメントの出力先としても使用されます。"
    },
    // 10-2-1 イミディエイトウィンドウへの出力
    {
        question: "VBAコードからイミディエイトウィンドウに出力を送るために使用するステートメントは？",
        choices: [
            "Console.WriteLine",
            "MsgBox",
            "Debug.Print"
        ],
        correct: 2,
        explanation: "`Debug.Print` ステートメントを使用して、VBAコードからイミディエイトウィンドウに出力を送ることができます。例えば、`Debug.Print \"現在の値: \" & myVariable` のように使用します。これは変数の値や処理の進行状況を確認するのに非常に役立ちます。"
    },
    // 10-3 マクロを一時停止する
    {
        question: "VBAでマクロの実行を一時停止する方法として、正しいものはどれですか？",
        choices: [
            "ブレークポイントを設定する",
            "Pause ステートメントを使用する",
            "Wait ステートメントを使用する"
        ],
        correct: 0,
        explanation: "マクロの実行を一時停止する最も一般的な方法は、ブレークポイントを設定することです。ブレークポイントは、コードエディタ内の特定の行に設定でき、その行に到達すると実行が一時停止します。これにより、その時点での変数の状態を確認したり、ステップ実行を開始したりできます。"
    },
    // 10-3-1 ブレークポイント
    {
        question: "VBAでブレークポイントを設定する手順は次のうちどれですか？",
        choices: [
            "コードの行の先頭に 'Break' と入力する",
            "デバッグメニューから 'ブレークポイントの設定' を選択する",
            "コードエディタで該当する行の左端の余白をクリックする"
        ],
        correct: 2,
        explanation: "VBAでブレークポイントを設定するには、コードエディタで該当する行の左端の余白をクリックします。正しく設定されると、その行に赤い点が表示されます。また、F9キーを押すことでも、カーソルがある行にブレークポイントを設定できます。"
    },
    // 10-3-2 Stopステートメント
    {
        question: "VBAの Stop ステートメントの効果は何ですか？",
        choices: [
            "プログラムを完全に終了する",
            "プログラムの実行を一時停止し、デバッグモードに入る",
            "エラーメッセージを表示する"
        ],
        correct: 1,
        explanation: "`Stop` ステートメントは、プログラムの実行を一時停止し、デバッグモードに入ります。これはコード内に直接記述できるブレークポイントのようなものです。ただし、`Stop` ステートメントはデバッグ目的でのみ使用し、最終的なコードからは削除すべきです。"
    },
    // 10-4 ステップ実行
    {
        question: "VBAのステップ実行で、現在の行をスキップして次の行に移動するためのキーボードショートカットは？",
        choices: [
            "F8",
            "F10",
            "Shift + F8"
        ],
        correct: 1,
        explanation: "F10キーを使用して、現在の行をスキップして次の行に移動します（ステップオーバー）。F8キーはステップイン（関数やプロシージャ内部に入る）、Shift + F8キーはステップアウト（現在の関数やプロシージャから抜ける）に使用されます。ステップ実行は、コードの動作を1行ずつ確認するのに非常に有用です。"
    },
    // 10-5 デバッグでよく使う関数
    {
        question: "デバッグ中に変数の型を確認するのに適したVBA関数は？",
        choices: [
            "TypeName()",
            "VarType()",
            "IsObject()"
        ],
        correct: 0,
        explanation: "`TypeName()` 関数は、変数の型を文字列として返すため、デバッグ中に変数の型を確認するのに適しています。例えば、`Debug.Print TypeName(myVariable)` とすることで、myVariableの型（例：「Integer」、「String」、「Object」など）をイミディエイトウィンドウに表示できます。"
    },
    // 10-5-1 IsNumeric関数
    {
        question: "IsNumeric 関数の主な用途は何ですか？",
        choices: [
            "変数が数値型かどうかを判断する",
            "文字列を数値に変換する",
            "数値の範囲をチェックする"
        ],
        correct: 0,
        explanation: "`IsNumeric` 関数の主な用途は、変数や式が数値として評価できるかどうかを判断することです。これは、ユーザー入力の検証や、数値演算を行う前のチェックに非常に有用です。例えば、`If IsNumeric(userInput) Then...` のように使用します。`IsNumeric` は文字列型の数値（\"123\"など）に対してもTrueを返すことに注意してください。"
    },
    // 実践的な問題
    {
        question: "次のVBAコードのデバッグ方法として最適なアプローチは何ですか？\n`Sub ProcessData()\n    Dim i As Integer\n    For i = 1 To 100\n        Cells(i, 1).Value = i * 2\n        If i = 50 Then\n            ' ここで何か問題が発生している可能性がある\n        End If\n    Next i\nEnd Sub`",
        choices: [
            "For ループの直前にブレークポイントを設定し、F5キーで実行を再開しながら i の値を監視する",
            "If i = 50 Then の行にブレークポイントを設定し、その時点での変数の状態を確認する",
            "各行に Debug.Print ステートメントを追加して、すべての変数の値を出力する"
        ],
        correct: 1,
        explanation: "この場合、`If i = 50 Then` の行にブレークポイントを設定するのが最適なアプローチです。問題が発生する可能性があるとコメントされている箇所で実行を一時停止し、その時点での変数の状態（特に i と Cells(i, 1).Value）を確認できます。ここから、ステップ実行（F8キー）を使用して、コードの動作を詳細に観察できます。全ての反復でデバッグ出力を行うよりも効率的で、問題の箇所を正確に特定できます。"
    }

];
