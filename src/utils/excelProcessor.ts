import ExcelJS from 'exceljs';
import { Question } from '../types';

export async function processExcel(file: File): Promise<Question[]> {
  const workbook = new ExcelJS.Workbook();
  const arrayBuffer = await file.arrayBuffer();
  
  await workbook.xlsx.load(arrayBuffer);
  
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('Aucune feuille de calcul trouvée dans le fichier Excel');
  }

  const questions: Question[] = [];
  
  // Parcourir les lignes (en commençant à 2 pour ignorer l'en-tête)
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Ignorer l'en-tête
    
    // Extraire les valeurs des cellules
    const question = row.getCell(1).value?.toString() || '';
    const answer = row.getCell(2).value?.toString() || '';
    const durationValue = row.getCell(3).value;
    const duration = durationValue ? parseInt(durationValue.toString()) : undefined;
    const imageUrl = row.getCell(4).value?.toString() || undefined;
    
    // Valider que c'est une question valide
    if (!question.trim()) {
      console.warn(`Ligne ${rowNumber} ignorée : pas de question`);
      return;
    }
    
    // Déterminer la bonne réponse (Vrai/Faux)
    const correctAnswer = answer === 'Vrai' || 
                         answer === 'VRAI' || 
                         answer === 'vrai' || 
                         answer === 'True' || 
                         answer === 'TRUE' || 
                         answer === '1';
    
    questions.push({
      question: question.trim(),
      correctAnswer,
      duration: duration || undefined,
      imageUrl: imageUrl?.trim() || undefined
    });
  });
  
  console.log(`${questions.length} questions extraites du fichier Excel`);
  return questions;
}

// Version alternative qui utilise les en-têtes de colonnes
export async function processExcelWithHeaders(file: File): Promise<Question[]> {
  const workbook = new ExcelJS.Workbook();
  const arrayBuffer = await file.arrayBuffer();
  
  await workbook.xlsx.load(arrayBuffer);
  
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('Aucune feuille de calcul trouvée dans le fichier Excel');
  }

  const questions: Question[] = [];
  
  // Obtenir les en-têtes de la première ligne
  const headerRow = worksheet.getRow(1);
  const headers: { [key: string]: number } = {};
  
  headerRow.eachCell((cell, colNumber) => {
    const header = cell.value?.toString().toLowerCase() || '';
    headers[header] = colNumber;
  });
  
  // Trouver les colonnes par leurs noms possibles
  const questionCol = headers['question'] || headers['questions'] || 1;
  const answerCol = headers['réponse'] || headers['réponse correcte'] || headers['answer'] || 2;
  const durationCol = headers['durée'] || headers['temps'] || headers['duration'] || 3;
  const imageCol = headers['image'] || headers['image url'] || headers['url image'] || headers['lien image'] || 4;
  
  // Parcourir les lignes de données
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Ignorer l'en-tête
    
    const question = row.getCell(questionCol).value?.toString() || '';
    const answer = row.getCell(answerCol).value?.toString() || '';
    const duration = row.getCell(durationCol).value;
    const imageUrl = row.getCell(imageCol).value?.toString();
    
    // Valider que c'est une question valide
    if (!question.trim()) {
      console.warn(`Ligne ${rowNumber} ignorée : pas de question`);
      return;
    }
    
    // Déterminer la bonne réponse
    const correctAnswer = answer === 'Vrai' || 
                         answer === 'VRAI' || 
                         answer === 'vrai' || 
                         answer === 'True' || 
                         answer === 'TRUE' || 
                         answer === '1';
    
    // Parser la durée
    let parsedDuration: number | undefined;
    if (duration) {
      const durationNum = typeof duration === 'number' ? duration : parseInt(duration.toString());
      parsedDuration = isNaN(durationNum) ? undefined : durationNum;
    }
    
    questions.push({
      question: question.trim(),
      correctAnswer,
      duration: parsedDuration,
      imageUrl: imageUrl?.trim() || undefined
    });
  });
  
  console.log(`${questions.length} questions extraites du fichier Excel`);
  return questions;
}