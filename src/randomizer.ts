import { Question, ExamStructure, GeneratedExam, MixedQuestion, QuestionType } from './types';

export const mixExams = (
  allQuestions: Question[],
  structure: ExamStructure,
  examCodes: string[]
): GeneratedExam[] => {
  const exams: GeneratedExam[] = [];

  // Helper to shuffle an array
  const shuffle = <T>(array: T[]): T[] => {
    const newArr = [...array];
    for (let i = newArr.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [newArr[i], newArr[j]] = [newArr[j], newArr[i]];
    }
    return newArr;
  };

  const selectQuestionsForExam = (): { part1: Question[], part2: Question[], part3: Question[] } => {
    const selected: Question[] = [];
    const usedIds = new Set<number>();

    // 1. Group questions by Topic and Type
    const poolByTopicAndType: Record<string, Record<string, Question[]>> = {};
    allQuestions.forEach(q => {
      if (!poolByTopicAndType[q.topic]) {
        poolByTopicAndType[q.topic] = { 'SINGLE_CHOICE': [], 'TRUE_FALSE': [], 'SHORT_ANSWER': [] };
      }
      poolByTopicAndType[q.topic][q.type].push(q);
    });

    // Shuffle all pools
    Object.keys(poolByTopicAndType).forEach(topic => {
      Object.keys(poolByTopicAndType[topic]).forEach(type => {
        poolByTopicAndType[topic][type] = shuffle(poolByTopicAndType[topic][type]);
      });
    });

    const topicCounts = { ...structure.topic_distribution };
    const partCounts = {
      'SINGLE_CHOICE': structure.part1_count,
      'TRUE_FALSE': structure.part2_count,
      'SHORT_ANSWER': structure.part3_count
    };

    const result = {
      'SINGLE_CHOICE': [] as Question[],
      'TRUE_FALSE': [] as Question[],
      'SHORT_ANSWER': [] as Question[]
    };

    // 2. Greedy selection to satisfy Topic Distribution
    const topics = Object.keys(topicCounts).filter(t => topicCounts[t] > 0);
    
    // Iterate through parts first to ensure we fill them
    const types: QuestionType[] = ['SINGLE_CHOICE', 'TRUE_FALSE', 'SHORT_ANSWER'];
    
    // First pass: try to satisfy both topic and part counts
    for (const type of types) {
      for (const topic of topics) {
        const neededForTopic = topicCounts[topic];
        const neededForPart = partCounts[type] - result[type].length;
        if (neededForTopic <= 0 || neededForPart <= 0) continue;

        const available = poolByTopicAndType[topic][type];
        const toPick = Math.min(available.length, neededForTopic, neededForPart);
        
        for (let i = 0; i < toPick; i++) {
          const q = available.pop()!;
          result[type].push(q);
          topicCounts[topic]--;
          usedIds.add(q.id!);
        }
      }
    }

    // Second pass: if some parts are still missing questions, fill them from any topic
    for (const type of types) {
      const remainingPool = shuffle(allQuestions.filter(q => q.type === type && !usedIds.has(q.id!)));
      while (result[type].length < partCounts[type] && remainingPool.length > 0) {
        const q = remainingPool.pop()!;
        result[type].push(q);
        usedIds.add(q.id!);
      }
    }

    return {
      part1: shuffle(result['SINGLE_CHOICE']),
      part2: shuffle(result['TRUE_FALSE']),
      part3: shuffle(result['SHORT_ANSWER'])
    };
  };

  for (const examCode of examCodes) {
    const { part1, part2, part3 } = selectQuestionsForExam();

    // Helper to mix options
    const mixOptions = (q: Question, shuffleOptions: boolean): MixedQuestion => {
      if (shuffleOptions && q.type === 'SINGLE_CHOICE') {
        const options = [
          { label: 'A', text: q.option_a || '', original: 'A' },
          { label: 'B', text: q.option_b || '', original: 'B' },
          { label: 'C', text: q.option_c || '', original: 'C' },
          { label: 'D', text: q.option_d || '', original: 'D' },
        ];
        const shuffledOptions = shuffle(options);
        let newCorrectLabel = '';
        const finalOptions = shuffledOptions.map((opt, idx) => {
          const label = String.fromCharCode(65 + idx);
          const isCorrect = opt.original === q.correct_answer;
          if (isCorrect) newCorrectLabel = label;
          return { label, text: opt.text, is_correct: isCorrect };
        });
        return { ...q, original_id: q.id!, shuffled_options: finalOptions, new_correct_answer: newCorrectLabel };
      } else {
        return { 
          ...q, 
          original_id: q.id!, 
          shuffled_options: [], 
          new_correct_answer: q.correct_answer 
        };
      }
    };

    exams.push({
      code: examCode,
      part1: part1.map(q => mixOptions(q, true)),
      part2: part2.map(q => mixOptions(q, false)),
      part3: part3.map(q => mixOptions(q, false)),
    });
  }

  return exams;
};
