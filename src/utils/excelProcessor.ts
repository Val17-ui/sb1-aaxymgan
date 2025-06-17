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
  
  // Obtenir les en-têtes de la première ligne
  const headerRow = worksheet.getRow(1);
  const headers: { [key: string]: number } = {};
  const optionColumns: { colNumber: number; letter: string }[] = [];
  
  // Analyser les en-têtes pour trouver toutes les colonnes
  headerRow.eachCell((cell, colNumber) => {
    const header = cell.value?.toString() || '';
    const normalizedHeader = header.toLowerCase().trim();
    
    // Stocker la position de chaque en-tête
    headers[normalizedHeader] = colNumber;
    
    // Détecter les colonnes d'options (Option1, OptionA, etc.)
    const optionMatch = header.match(/^option\s*([a-zA-Z0-9]+)$/i);
    if (optionMatch) {
      optionColumns.push({
        colNumber,
        letter: optionMatch[1].toUpperCase()
      });
    }
  });
  
  // Trier les colonnes d'options par leur position
  optionColumns.sort((a, b) => a.colNumber - b.colNumber);
  
  // Identifier les colonnes importantes
  const questionCol = headers['question'] || headers['questions'] || 1;
  const correctAnswerCol = headers['bonnereponse'] || headers['bonne reponse'] || 
                          headers['bonne réponse'] || headers['correctanswer'] || 
                          headers['correct answer'] || headers['réponse correcte'] || 0;
  const imageCol = headers['image'] || headers['image url'] || headers['url image'] || 
                   headers['lien image'] || headers['imageurl'] || 0;
  
  // Valider qu'on a au moins les colonnes essentielles
  if (!headers['question'] && !headers['questions']) {
    throw new Error('Colonne "Question" non trouvée. Assurez-vous que la première ligne contient les en-têtes.');
  }
  
  if (optionColumns.length === 0) {
    throw new Error('Aucune colonne d\'option trouvée. Les colonnes doivent être nommées "OptionA", "Option1", etc.');
  }
  
  console.log(`Colonnes détectées: ${optionColumns.length} options, BonneReponse: ${correctAnswerCol > 0 ? 'Oui' : 'Non'}, Image: ${imageCol > 0 ? 'Oui' : 'Non'}`);
  
  // Parcourir les lignes de données
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Ignorer l'en-tête
    
    const questionText = row.getCell(questionCol).value?.toString() || '';
    
    // Valider que c'est une question valide
    if (!questionText.trim()) {
      console.warn(`Ligne ${rowNumber} ignorée : pas de question`);
      return;
    }
    
    // Extraire toutes les options non vides
    const options: string[] = [];
    for (const optionCol of optionColumns) {
      const cellValue = row.getCell(optionCol.colNumber).value;
      let optionValue = '';
      
      // Gérer les cas spéciaux d'Excel uniquement pour les booléens
      if (typeof cellValue === 'boolean') {
        // Excel a converti en booléen, on doit deviner la langue voulue
        // On regarde les autres options pour détecter la langue
        const otherOptions: string[] = [];
        for (const col of optionColumns) {
          const val = row.getCell(col.colNumber).value;
          if (typeof val === 'string') {
            otherOptions.push(val.toLowerCase());
          }
        }
        
        // Détecter si c'est en français ou anglais
        const isFrench = otherOptions.some(opt => 
          opt.includes('vrai') || opt.includes('faux') || 
          opt.includes('oui') || opt.includes('non')
        );
        
        if (isFrench) {
          optionValue = cellValue ? 'Vrai' : 'Faux';
        } else {
          optionValue = cellValue ? 'True' : 'False';
        }
      } else if (cellValue !== null && cellValue !== undefined) {
        // Garder la valeur telle quelle si ce n'est pas un booléen
        optionValue = cellValue.toString();
      }
      
      if (optionValue.trim()) {
        options.push(optionValue.trim());
      } else {
        // Si on trouve une cellule vide, on arrête
        break;
      }
    }
    
    // Valider qu'on a au moins une option
    if (options.length === 0) {
      console.warn(`Ligne ${rowNumber} ignorée : aucune option trouvée pour la question "${questionText}"`);
      return;
    }
    
    // Parser la bonne réponse si elle existe
    let correctAnswerIndex: number | undefined;
    if (correctAnswerCol > 0) {
      const correctAnswerValue = row.getCell(correctAnswerCol).value?.toString() || '';
      if (correctAnswerValue.trim()) {
        correctAnswerIndex = parseCorrectAnswer(correctAnswerValue.trim());
        
        // Valider que l'index est valide
        if (correctAnswerIndex !== undefined && (correctAnswerIndex < 0 || correctAnswerIndex >= options.length)) {
          console.warn(`Ligne ${rowNumber} : Bonne réponse "${correctAnswerValue}" invalide pour ${options.length} options`);
          correctAnswerIndex = undefined;
        }
      }
    }
    
    // Extraire l'URL de l'image si présente
    const imageUrl = imageCol > 0 ? (row.getCell(imageCol).value?.toString()?.trim() || undefined) : undefined;
    
    // Ajouter la question
    questions.push({
      question: questionText.trim(),
      options,
      correctAnswerIndex,
      imageUrl
    });
    
    // Log pour debug
    console.log(`Question ${rowNumber}: "${questionText.substring(0, 30)}..." - ${options.length} options${correctAnswerIndex !== undefined ? ` (réponse: ${correctAnswerIndex + 1})` : ' (sondage)'}`);
  });
  
  console.log(`\n${questions.length} questions extraites du fichier Excel`);
  
  // Statistiques finales
  const stats = {
    total: questions.length,
    avecBonneReponse: questions.filter(q => q.correctAnswerIndex !== undefined).length,
    avecImage: questions.filter(q => q.imageUrl).length,
    parNombreOptions: {} as { [key: number]: number }
  };
  
  questions.forEach(q => {
    const count = q.options.length;
    stats.parNombreOptions[count] = (stats.parNombreOptions[count] || 0) + 1;
  });
  
  console.log('Statistiques:');
  console.log(`- Questions avec bonne réponse: ${stats.avecBonneReponse}`);
  console.log(`- Questions avec image: ${stats.avecImage}`);
  console.log(`- Distribution des options:`, stats.parNombreOptions);
  
  return questions;
}

/**
 * Parse la valeur de la bonne réponse pour obtenir l'index
 * Accepte: "1", "2", "A", "B", "Option1", "OptionA", etc.
 */
function parseCorrectAnswer(value: string): number | undefined {
  const normalized = value.toUpperCase().trim();
  
  // Cas 1: Numéro direct (1, 2, 3...)
  const numberMatch = normalized.match(/^(\d+)$/);
  if (numberMatch) {
    const num = parseInt(numberMatch[1]);
    // Convertir en index 0-based
    return num > 0 ? num - 1 : undefined;
  }
  
  // Cas 2: Lettre (A, B, C...)
  const letterMatch = normalized.match(/^([A-Z])$/);
  if (letterMatch) {
    // A=0, B=1, C=2, etc.
    return letterMatch[1].charCodeAt(0) - 'A'.charCodeAt(0);
  }
  
  // Cas 3: Format "Option" + numéro/lettre
  const optionMatch = normalized.match(/^OPTION\s*([A-Z0-9]+)$/);
  if (optionMatch) {
    const suffix = optionMatch[1];
    
    // Si c'est un nombre
    if (/^\d+$/.test(suffix)) {
      const num = parseInt(suffix);
      return num > 0 ? num - 1 : undefined;
    }
    
    // Si c'est une lettre
    if (/^[A-Z]$/.test(suffix)) {
      return suffix.charCodeAt(0) - 'A'.charCodeAt(0);
    }
  }
  
  // Cas non reconnu
  console.warn(`Format de bonne réponse non reconnu: "${value}"`);
  return undefined;
}

// Export pour les tests
export { parseCorrectAnswer };