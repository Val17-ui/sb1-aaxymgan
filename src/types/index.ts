export interface Question {
  question: string;
  correctAnswer: boolean;
  duration?: number;
  imageUrl?: string;
}

export interface GenerationOptions {
  fileName?: string;
  defaultDuration?: number;
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