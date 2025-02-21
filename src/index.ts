import mammoth from "mammoth";
import fs from "fs";
import path from "path";
import { Document, Packer, Paragraph, TextRun } from "docx";

type Option = { text: string; isCorrect: boolean };
type Question = { question: string; options: Option[] };

/**
 * Hàm tiện ích tạo TextRun với style: font Arial, cỡ chữ 14 (28 half-point)
 * và màu chữ có thể tùy chỉnh (mặc định là đen "000000").
 */
function createTextRun(text: string, color: string = "000000"): TextRun {
  return new TextRun({
    text,
    font: "Arial",
    size: 28, // 14pt = 14*2 = 28 half-points
    color,
  });
}

/**
 * Hàm parseFile: trích xuất nội dung văn bản thô từ file .docx và phân tích thành danh sách câu hỏi.
 * Mỗi dòng được định dạng như:
 * "1. Câu hỏi? A. Phương án 1;[*] B. Phương án 2;[*] = C. Phương án 3;[*]"
 */
async function parseFile(filePath: string): Promise<Question[]> {
  const result = await mammoth.extractRawText({ path: filePath });
  const lines = result.value.split("\n");
  const questions: Question[] = [];

  for (const line of lines) {
    // Regex tách câu hỏi và phần đáp án (giả sử đáp án bắt đầu từ "A.")
    const questionMatch = line.match(/^(\d+\.)\s+(.*?[?:])\s*(A\..*)$/);
    if (questionMatch) {
      // Lấy phần câu hỏi, loại bỏ số thứ tự ban đầu để tái đánh số sau
      const questionText = questionMatch[2].trim();
      const optionsText = questionMatch[3];

      // Tách các phương án trả lời bằng regex
      const optionRegex = /([A-Z])\.\s*(.*?)\s*;\[\*\](\s*=)?/g;
      const options: Option[] = [];
      let match;
      while ((match = optionRegex.exec(optionsText)) !== null) {
        const text = match[2].trim();
        const isCorrect = !!match[3]; // Có dấu "=" thì là đáp án đúng
        options.push({ text, isCorrect });
      }
      questions.push({ question: questionText, options });
    }
  }

  if (questions.length === 0) {
    console.warn(`Không tìm thấy câu hỏi hợp lệ trong file ${filePath}.`);
  }
  return questions;
}

/**
 * Hàm shuffleArray: Tráo thứ tự một mảng theo thuật toán Fisher-Yates.
 */
function shuffleArray<T>(array: T[]): T[] {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

/**
 * Hàm getShuffledAnswerLetter: Tráo thứ tự các phương án trả lời và xác định ký tự của đáp án đúng.
 * Nếu có nhiều đáp án đúng thì nối chúng bằng dấu phẩy.
 * Trả về cả danh sách các phương án đã tráo thứ tự và ký tự đáp án đúng.
 */
function getShuffledAnswerLetter(options: Option[]): {
  shuffledOptions: Option[];
  answerLetter: string;
} {
  const shuffled = shuffleArray([...options]);
  const correctLetters: string[] = [];
  shuffled.forEach((opt, index) => {
    if (opt.isCorrect) {
      // Gán ký tự dựa trên vị trí sau khi tráo: A, B, C, ...
      correctLetters.push(String.fromCharCode(65 + index));
    }
  });
  return { shuffledOptions: shuffled, answerLetter: correctLetters.join(", ") };
}

/**
 * Hàm main: Xử lý tham số dòng lệnh, đọc file questions.docx, chọn ngẫu nhiên số câu hỏi theo yêu cầu,
 * và tạo file randoms.docx theo định dạng yêu cầu.
 *
 * Mỗi câu hỏi sẽ hiển thị:
 *   Số thứ tự và nội dung câu hỏi.
 *   Danh sách các lựa chọn theo thứ tự ngẫu nhiên (đánh số lại từ A, B, C, …)
 *
 * Ở cuối file sẽ có phần tổng hợp đáp án dưới dạng:
 *   answers:[
 *   {Q: 1; A: A},
 *   {Q: 2; A: B},
 *   ]
 */
async function main() {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    displayUsage();
    process.exit(0);
  }

  // Lấy tham số: directory và count
  const folderPathArg = args.find((arg) => arg.startsWith("directory="));
  const countArg = args.find((arg) => arg.startsWith("count="));
  if (!folderPathArg || !countArg) {
    console.error(
      "Vui lòng cung cấp đường dẫn folder và số lượng câu hỏi (count)."
    );
    displayUsage();
    process.exit(1);
  }

  const folderPath = folderPathArg.split("=")[1];
  const count = parseInt(countArg.split("=")[1], 10);
  if (isNaN(count) || count <= 0) {
    console.error("Tham số count phải là số nguyên dương.");
    process.exit(1);
  }

  const resolvedFolderPath = path.resolve(folderPath);
  if (
    !fs.existsSync(resolvedFolderPath) ||
    !fs.statSync(resolvedFolderPath).isDirectory()
  ) {
    console.error(
      "Đường dẫn không hợp lệ. Đường dẫn phải bằng tiếng Anh và không chứa khoảng trắng."
    );
    process.exit(1);
  }

  const questionFilePath = path.join(resolvedFolderPath, "questions.docx");
  if (!fs.existsSync(questionFilePath)) {
    console.error("File questions.docx không tồn tại trong folder đã cho.");
    process.exit(1);
  }

  try {
    const questions = await parseFile(questionFilePath);
    if (questions.length === 0) {
      console.error("Không có câu hỏi nào được tìm thấy trong file.");
      process.exit(1);
    }

    // Tráo thứ tự các câu hỏi và chọn số lượng câu hỏi theo tham số (nếu count > tổng số thì chọn tất cả)
    const shuffledQuestions = shuffleArray([...questions]);
    const selectedQuestions = shuffledQuestions.slice(
      0,
      Math.min(count, shuffledQuestions.length)
    );

    let answerSummary: { Q: number; A: string }[] = [];
    let paragraphs: Paragraph[] = [];

    // Xử lý từng câu hỏi: thêm câu hỏi (bao gồm các đáp án) vào file và lưu đáp án đúng
    selectedQuestions.forEach((q, index) => {
      const questionIndex = index + 1;

      // Gọi hàm getShuffledAnswerLetter để tráo thứ tự các phương án và lấy đáp án đúng
      const { shuffledOptions, answerLetter } = getShuffledAnswerLetter(
        q.options
      );

      // Thêm câu hỏi
      const questionLine = `${questionIndex}. ${q.question.replace(
        /^\d+\.\s*/,
        ""
      )}`;
      paragraphs.push(
        new Paragraph({ children: [createTextRun(questionLine)] })
      );

      // Thêm danh sách các đáp án (không đánh dấu đáp án nào là đúng)
      shuffledOptions.forEach((option, optIndex) => {
        const letter = String.fromCharCode(65 + optIndex);
        const optionLine = `${letter}. ${option.text};[*]`;
        paragraphs.push(
          new Paragraph({ children: [createTextRun(optionLine)] })
        );
      });

      // Thêm khoảng trắng giữa các câu hỏi
      paragraphs.push(new Paragraph({ children: [createTextRun("")] }));

      // Lưu lại đáp án đúng của câu hỏi (theo thứ tự sau khi tráo)
      answerSummary.push({ Q: questionIndex, A: answerLetter });
    });

    // Thêm phần tổng hợp đáp án vào cuối file
    paragraphs.push(
      new Paragraph({ children: [createTextRun("answers:[", "FFFFFF")] })
    );
    answerSummary.forEach((item) => {
      const answerLine = `{Q: ${item.Q}; A: ${item.A}},`;
      paragraphs.push(
        new Paragraph({ children: [createTextRun(answerLine, "FFFFFF")] })
      );
    });
    paragraphs.push(
      new Paragraph({ children: [createTextRun("]", "FFFFFF")] })
    );

    // Tạo file DOCX
    const doc = new Document({ sections: [{ children: paragraphs }] });
    const buffer = await Packer.toBuffer(doc);
    const outputFilePath = path.join(resolvedFolderPath, "randoms.docx");
    fs.writeFileSync(outputFilePath, buffer);
  } catch (error: any) {
    console.error("Đã có lỗi xảy ra:", error.message);
    console.error("Stack trace:", error.stack);
  }
}

/**
 * Hàm hiển thị hướng dẫn sử dụng
 */
function displayUsage() {
  console.log("Usage:");
  console.log(
    '  npm run dev directory="path/to/your/directory" count=NUMBER_OF_QUESTIONS'
  );
}

main();
