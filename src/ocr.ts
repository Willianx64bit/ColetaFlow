import { registerPlugin } from '@capacitor/core';

export interface OcrRecognitionResult {
  text: string;
}

export interface OcrPlugin {
  recognize(options: { imageBase64: string }): Promise<OcrRecognitionResult>;
}

export const Ocr = registerPlugin<OcrPlugin>('Ocr');
