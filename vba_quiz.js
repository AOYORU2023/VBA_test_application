const quizData = [
    {
        question: "VBAで変数を宣言するキーワードは何ですか？",
        choices: ["var", "dim", "let"],
        correct: 1,
        explanation: "VBAでは、'Dim'キーワードを使用して変数を宣言します。例: Dim myVariable As Integer"
    },
    {
        question: "VBAでループを終了させるステートメントは？",
        choices: ["End Loop", "Exit For", "Break"],
        correct: 1,
        explanation: "'Exit For'を使用してForループを途中で終了させることができます。While/Do...Whileループの場合は'Exit Do'を使用します。"
    },
    {
        question: "VBAで配列の最初のインデックスは通常何ですか？",
        choices: ["0", "1", "任意に設定可能"],
        correct: 0,
        explanation: "VBAでは、デフォルトで配列の最初のインデックスは0です。ただし、'Option Base 1'を使用して1から始めることも可能です。"
    }
];

let currentQuestion = 0;

function loadQuestion() {
    const question = quizData[currentQuestion];
    document.getElementById('question').innerHTML = `<h2>${question.question}</h2>`;
    const choicesHtml = question.choices.map((choice, index) => 
        `<button onclick="checkAnswer(${index})">${choice}</button>`
    ).join('');
    document.getElementById('choices').innerHTML = choicesHtml;
    document.getElementById('result').innerHTML = '';
    document.getElementById('explanation').innerHTML = '';
    document.getElementById('nextButton').style.display = 'none';
}

function checkAnswer(choiceIndex) {
    const question = quizData[currentQuestion];
    const resultDiv = document.getElementById('result');
    const explanationDiv = document.getElementById('explanation');
    
    if (choiceIndex === question.correct) {
        resultDiv.innerHTML = '<h3 style="color: green;">正解です！</h3>';
    } else {
        resultDiv.innerHTML = '<h3 style="color: red;">不正解です。</h3>';
    }
    
    explanationDiv.innerHTML = `<p><strong>解説:</strong> ${question.explanation}</p>`;
    document.getElementById('nextButton').style.display = 'block';
}

function nextQuestion() {
    currentQuestion++;
    if (currentQuestion < quizData.length) {
        loadQuestion();
    } else {
        document.getElementById('quiz').innerHTML = '<h2>クイズ終了！お疲れ様でした。</h2>';
    }
}

document.getElementById('nextButton').addEventListener('click', nextQuestion);

loadQuestion();