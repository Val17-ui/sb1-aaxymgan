export interface Question {
  question: string;
  correctAnswer: boolean;
  duration?: number;
  imageUrl?: string;
}

export interface ConfigOptions {
  pollStartMode: 'Automatic' | 'Manual';
  answersBulletStyle: 'ppBulletAlphaUCParenRight' | 'ppBulletAlphaUCPeriod' | 'ppBulletArabicParenRight' | 'ppBulletArabicPeriod';
  chartValueLabelFormat: 'Response_Count' | 'Percentage_Integer' | 'Percentage_One_Decimal' | 'Percentage_Two_Decimal' | 'Percentage_Three_Decimal';
  pollTimeLimit: number;
  pollCountdownStartMode: 'Automatic' | 'Manual';
  pollMultipleResponse: number;
}

export interface GenerationOptions {
  fileName?: string;
  defaultDuration?: number;
  ombeaConfig?: ConfigOptions;
}

export interface ImageDimensions {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface DownloadedImage {
  fileName: string;
  data: ArrayBuffer;
  width: number;
  height: number;
  dimensions: ImageDimensions;
}