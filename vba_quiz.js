const quizData = [
    {
        question: "VBA�ŕϐ���錾����L�[���[�h�͉��ł����H",
        choices: ["var", "dim", "let"],
        correct: 1,
        explanation: "VBA�ł́A'Dim'�L�[���[�h���g�p���ĕϐ���錾���܂��B��: Dim myVariable As Integer"
    },
    {
        question: "VBA�Ń��[�v���I��������X�e�[�g�����g�́H",
        choices: ["End Loop", "Exit For", "Break"],
        correct: 1,
        explanation: "'Exit For'���g�p����For���[�v��r���ŏI�������邱�Ƃ��ł��܂��BWhile/Do...While���[�v�̏ꍇ��'Exit Do'���g�p���܂��B"
    },
    {
        question: "VBA�Ŕz��̍ŏ��̃C���f�b�N�X�͒ʏ퉽�ł����H",
        choices: ["0", "1", "�C�ӂɐݒ�\"],
        correct: 0,
        explanation: "VBA�ł́A�f�t�H���g�Ŕz��̍ŏ��̃C���f�b�N�X��0�ł��B�������A'Option Base 1'���g�p����1����n�߂邱�Ƃ��\�ł��B"
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
        resultDiv.innerHTML = '<h3 style="color: green;">�����ł��I</h3>';
    } else {
        resultDiv.innerHTML = '<h3 style="color: red;">�s�����ł��B</h3>';
    }
    
    explanationDiv.innerHTML = `<p><strong>���:</strong> ${question.explanation}</p>`;
    document.getElementById('nextButton').style.display = 'block';
}

function nextQuestion() {
    currentQuestion++;
    if (currentQuestion < quizData.length) {
        loadQuestion();
    } else {
        document.getElementById('quiz').innerHTML = '<h2>�N�C�Y�I���I�����l�ł����B</h2>';
    }
}

document.getElementById('nextButton').addEventListener('click', nextQuestion);

loadQuestion();