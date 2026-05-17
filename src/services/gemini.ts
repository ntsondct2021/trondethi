import { GoogleGenAI } from "@google/genai";

export const generateQuestionsWithAI = async (prompt: string) => {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error("GEMINI_API_KEY is not configured in environment variables.");
  }

  const ai = new GoogleGenAI({ apiKey });

  const systemInstruction = `
    Bạn là một chuyên gia soạn đề thi trắc nghiệm. 
    Nhiệm vụ của bạn là tạo ra các câu hỏi trắc nghiệm dựa trên nội dung người dùng cung cấp.
    Mỗi câu hỏi phải có 4 phương án (A, B, C, D) và 1 đáp án đúng.
    Kết quả trả về phải là một mảng JSON các đối tượng có cấu trúc:
    {
      "content": "Nội dung câu hỏi",
      "option_a": "Phương án A",
      "option_b": "Phương án B",
      "option_c": "Phương án C",
      "option_d": "Phương án D",
      "correct_answer": "A" | "B" | "C" | "D",
      "topic": "Môn học/Chuyên đề",
      "difficulty": "Dễ" | "Trung bình" | "Khó",
      "explanation": "Giải thích ngắn gọn tại sao chọn đáp án đó"
    }
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
      config: {
        systemInstruction,
        responseMimeType: "application/json",
      },
    });

    const responseText = response.text;
    if (!responseText) throw new Error("AI không trả về nội dung.");
    return JSON.parse(responseText);
  } catch (error) {
    console.error("Error with AI response:", error);
    throw error;
  }
};
