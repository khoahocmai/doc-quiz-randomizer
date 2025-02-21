**doc-quiz-randomizer**

A Node.js utility for generating a randomized quiz document from a .docx file containing questions. This tool reads questions and multiple-choice options from a source file (e.g., questions.docx), randomizes the order of both the questions and the answer options, and then outputs a new .docx file with the following characteristics:

- **Randomized Questions:** Each question is re-numbered and its answer options are shuffled (the correct answer is not indicated within the question).
- **Answer Summary:** At the end of the document, a summary section lists the correct answer(s) for each question (displayed in white text).
- **Custom Styling:** All text in the document is rendered in Arial font at 14pt size.

**Features**

- Parses a .docx file to extract questions and multiple-choice answers.
- Randomizes the order of questions and answer options.
- Generates a new .docx file (randoms.docx) with:
  - Questions and their shuffled answer choices.
  - An answer summary section at the end (answers displayed in white).

**Prerequisites**

- **Node.js** (version 14 or higher)
- A source .docx file (questions.docx) that contains your quiz questions following the expected format.

**Installation**

1. Clone this repository:

git clone https://github.com/khoahocmai/doc-quiz-randomizer.git

cd doc-quiz-randomizer

1. Install dependencies:

npm install

**Input Format**

Your **questions.docx** file should adhere to the following format:

1. **Question Format:**
   1. Each question begins with a number (e.g., 1. , 2. ) followed by the question text ending with a ‘?’ or ‘:’.
1. **Answer Options:**
   1. Each answer option starts with a capital letter (e.g., A. , B. , C. , etc.), followed by the answer text and ending with ;[\*].
   1. The correct answer is marked with an extra equals sign at the end (e.g., ;[\*] =).

**Example:**

1\. What is the capital of France?

A. Berlin;[\*]

B. Madrid;[\*]

C. Paris;[\*] =

D. Rome;[\*]

**Usage**

1. Place your questions.docx file inside a folder.
1. Run the script using the command line with the following parameters:
   1. **directory**: The path to your folder containing questions.docx.
   1. **count**: The number of questions to select randomly.

For example, to select 5 random questions:

npm run dev directory="path/to/your/folder" count=5

The script will generate a file named randoms.docx in the same folder.

**Output**

The generated **randoms.docx** file contains:

- **Randomized Questions:**
  Each question is displayed with its randomized answer options (labeled A, B, C, …) without any indication of which option is correct.
- **Answer Summary:**
  At the end of the document, a summary is appended in white text showing the correct answer for each question in the following format:

answers:[

{Q: 1; A: A},

{Q: 2; A: B},

]

**Dependencies**

- **mammoth**: Extracts raw text from .docx files.
- **docx**: Generates .docx documents with custom styling.
- **fs** and **path**: Node.js modules for file system and path handling.

**License**

This project is licensed under the MIT License.

**Author**

👨‍💻 [Created by khoahocmai](https://github.com/khoahocmai)
