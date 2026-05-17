export type QuestionType = 'SINGLE_CHOICE' | 'TRUE_FALSE' | 'SHORT_ANSWER';

export interface Question {
  id?: number;
  type: QuestionType;
  content: string;
  option_a?: string;
  option_b?: string;
  option_c?: string;
  option_d?: string;
  correct_answer: string;
  topic: string;
  difficulty: 'Dễ' | 'Trung bình' | 'Khó';
  image_url?: string;
  explanation?: string;
}

export interface ExamStructure {
  total_questions: number;
  part1_count: number;
  part2_count: number;
  part3_count: number;
  difficulty_ratios: {
    'Dễ': number;
    'Trung bình': number;
    'Khó': number;
  };
  topic_distribution: Record<string, number>;
}

export interface GeneratedExam {
  code: string;
  part1: MixedQuestion[];
  part2: MixedQuestion[];
  part3: MixedQuestion[];
}

export interface MixedQuestion extends Question {
  original_id: number;
  shuffled_options: {
    label: string; // A, B, C, D
    text: string;
    is_correct: boolean;
  }[];
  new_correct_answer: string;
}
