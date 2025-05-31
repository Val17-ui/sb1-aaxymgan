import JSZip from 'jszip';
import { saveAs } from 'file-saver';

interface Question {
  question: string;
  correctAnswer?: boolean;
  duration?: number;
}

interface GenerationOptions {
  fileName?: string;
}

// Fonction utilitaire pour échapper le XML
function escapeXml(unsafe: string): string {
  return unsafe.replace(/[<>&'"]/g, function (c: string): string {
    switch (c) {
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '&': return '&amp;';
      case '"': return '&quot;';
      case "'": return '&apos;';
      default: return c;
    }
  });
}

// Génère le XML d'une nouvelle slide basée sur un layout
function createSlideXml(question: string, slideLayoutRelId: string = "rId1"): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>

      <!-- Zone de titre -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="title"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="685800" y="914400"/>
            <a:ext cx="10820400" cy="1371600"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="fr-FR"/>
              <a:t>${escapeXml(question)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- Zone de contenu/réponses -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Content"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="body" idx="1"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="685800" y="2743200"/>
            <a:ext cx="10820400" cy="3429000"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle>
            <a:lvl1pPr>
              <a:buFont typeface="Arial"/>
              <a:buAutoNum type="arabicPeriod"/>
            </a:lvl1pPr>
          </a:lstStyle>
          <a:p>
            <a:pPr lvl="0"/>
            <a:r>
              <a:rPr lang="fr-FR"/>
              <a:t>Vrai</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="0"/>
            <a:r>
              <a:rPr lang="fr-FR"/>
              <a:t>Faux</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
    <p:clrMapOvr>
      <a:masterClrMapping/>
    </p:clrMapOvr>
  </p:cSld>
</p:sld>`;
}

// Génère le fichier .rels pour une slide
function createSlideRelsXml(slideLayoutRelId: string = "rId1"): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${slideLayoutRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
}

// Met à jour le fichier [Content_Types].xml pour inclure les nouvelles slides
function updateContentTypes(originalContent: string, newSlideCount: number, existingSlideCount: number): string {
  let updatedContent = originalContent;
  
  // Ajouter les nouvelles slides dans [Content_Types].xml
  const insertPoint = updatedContent.lastIndexOf('</Types>');
  let newOverrides = '';
  
  for (let i = existingSlideCount + 1; i <= existingSlideCount + newSlideCount; i++) {
    newOverrides += `\n  <Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  }
  
  return updatedContent.slice(0, insertPoint) + newOverrides + '\n' + updatedContent.slice(insertPoint);
}

// Met à jour presentation.xml pour inclure les nouvelles slides
function updatePresentationXml(originalContent: string, newSlideCount: number, existingSlideCount: number): string {
  // Trouver la section sldIdLst
  const sldIdLstStart = originalContent.indexOf('<p:sldIdLst>');
  const sldIdLstEnd = originalContent.indexOf('</p:sldIdLst>') + '</p:sldIdLst>'.length;
  
  if (sldIdLstStart === -1 || sldIdLstEnd === -1) {
    throw new Error('Structure presentation.xml invalide - section sldIdLst introuvable');
  }
  
  // Extraire la section existante
  const beforeSldIdLst = originalContent.slice(0, sldIdLstStart);
  const existingSldIdLst = originalContent.slice(sldIdLstStart, sldIdLstEnd);
  const afterSldIdLst = originalContent.slice(sldIdLstEnd);
  
  // Créer les nouvelles entrées de slides
  let newSlideEntries = '';
  for (let i = existingSlideCount + 1; i <= existingSlideCount + newSlideCount; i++) {
    const slideId = 255 + i; // ID unique pour chaque slide
    newSlideEntries += `\n    <p:sldId id="${slideId}" r:id="rIdSlide${i}"/>`;
  }
  
  // Insérer les nouvelles slides avant la fermeture de sldIdLst
  const updatedSldIdLst = existingSldIdLst.replace('</p:sldIdLst>', newSlideEntries + '\n  </p:sldIdLst>');
  
  return beforeSldIdLst + updatedSldIdLst + afterSldIdLst;
}

// Met à jour presentation.xml.rels pour inclure les relations vers les nouvelles slides
function updatePresentationRels(originalContent: string, newSlideCount: number, existingSlideCount: number): string {
  const insertPoint = originalContent.lastIndexOf('</Relationships>');
  let newRelationships = '';
  
  for (let i = existingSlideCount + 1; i <= existingSlideCount + newSlideCount; i++) {
    newRelationships += `\n  <Relationship Id="rIdSlide${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i}.xml"/>`;
  }
  
  return originalContent.slice(0, insertPoint) + newRelationships + '\n' + originalContent.slice(insertPoint);
}

// Compte le nombre de slides existantes dans le modèle
function countExistingSlides(zip: JSZip): number {
  let count = 0;
  zip.folder('ppt/slides')?.forEach((relativePath) => {
    if (relativePath.match(/^slide\d+\.xml$/) && !relativePath.includes('_rels')) {
      count++;
    }
  });
  return count;
}

// Validation des données d'entrée
function validateQuestions(questions: Question[]): void {
  if (!Array.isArray(questions) || questions.length === 0) {
    throw new Error('Au moins une question est requise');
  }
  
  questions.forEach((question, index) => {
    if (!question.question || typeof question.question !== 'string' || question.question.trim() === '') {
      throw new Error(`Question ${index + 1}: Le texte de la question est requis`);
    }
    
    if (question.duration && (typeof question.duration !== 'number' || question.duration <= 0)) {
      throw new Error(`Question ${index + 1}: La durée doit être un nombre positif`);
    }
  });
}

// Fonction principale de génération
export async function generatePPTX(
  templateFile: File, 
  questions: Question[], 
  options: GenerationOptions = {}
): Promise<void> {
  try {
    console.log('Validation des données...');
    validateQuestions(questions);
    
    console.log('Chargement du modèle...');
    
    // Charger le modèle PowerPoint
    const templateZip = await JSZip.loadAsync(templateFile);
    
    // Compter les slides existantes
    const existingSlideCount = countExistingSlides(templateZip);
    console.log(`Slides existantes dans le modèle: ${existingSlideCount}`);
    console.log(`Nouvelles slides à créer: ${questions.length}`);
    
    // Créer une copie du modèle
    const outputZip = new JSZip();
    
    // Copier tout le contenu du modèle
    const copyPromises: Promise<void>[] = [];
    templateZip.forEach((relativePath, file) => {
      if (!file.dir) {
        const promise = file.async('blob').then(content => {
          outputZip.file(relativePath, content);
        });
        copyPromises.push(promise);
      }
    });
    
    await Promise.all(copyPromises);
    console.log('Modèle copié');
    
    // Créer les nouvelles slides
    console.log('Création des nouvelles slides...');
    
    for (let i = 0; i < questions.length; i++) {
      const slideNumber = existingSlideCount + i + 1;
      const question = questions[i];
      
      // Créer le fichier slide XML
      const slideXml = createSlideXml(question.question);
      outputZip.file(`ppt/slides/slide${slideNumber}.xml`, slideXml);
      
      // Créer le fichier .rels pour la slide
      const slideRelsXml = createSlideRelsXml();
      outputZip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, slideRelsXml);
      
      console.log(`Slide ${slideNumber} créée: ${question.question.substring(0, 50)}...`);
    }
    
    // Mettre à jour [Content_Types].xml
    console.log('Mise à jour des métadonnées...');
    const contentTypesFile = outputZip.file('[Content_Types].xml');
    if (contentTypesFile) {
      const contentTypesContent = await contentTypesFile.async('string');
      const updatedContentTypes = updateContentTypes(contentTypesContent, questions.length, existingSlideCount);
      outputZip.file('[Content_Types].xml', updatedContentTypes);
    }
    
    // Mettre à jour presentation.xml
    const presentationFile = outputZip.file('ppt/presentation.xml');
    if (presentationFile) {
      const presentationContent = await presentationFile.async('string');
      const updatedPresentation = updatePresentationXml(presentationContent, questions.length, existingSlideCount);
      outputZip.file('ppt/presentation.xml', updatedPresentation);
    }
    
    // Mettre à jour presentation.xml.rels
    const presentationRelsFile = outputZip.file('ppt/_rels/presentation.xml.rels');
    if (presentationRelsFile) {
      const presentationRelsContent = await presentationRelsFile.async('string');
      const updatedPresentationRels = updatePresentationRels(presentationRelsContent, questions.length, existingSlideCount);
      outputZip.file('ppt/_rels/presentation.xml.rels', updatedPresentationRels);
    }
    
    // Générer le fichier final
    console.log('Génération du fichier final...');
    const outputBlob = await outputZip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    const fileName = options.fileName || `Questions_${new Date().toISOString().slice(0, 10)}.pptx`;
    saveAs(outputBlob, fileName);
    
    console.log(`Fichier généré avec succès: ${fileName}`);
    console.log(`Total des slides: ${existingSlideCount + questions.length}`);
    
  } catch (error: any) {
    console.error('Erreur lors de la génération:', error);
    throw new Error(`Génération échouée: ${error.message}`);
  }
}

// Fonction utilitaire pour tester avec des données d'exemple
export function createTestQuestions(): Question[] {
  return [
    { question: "Paris est-elle la capitale de la France ?" },
    { question: "Le soleil tourne-t-il autour de la Terre ?" },
    { question: "L'eau bout-elle à 100°C au niveau de la mer ?" },
    { question: "JavaScript est-il un langage de programmation ?" },
    { question: "Les pingouins vivent-ils au pôle Nord ?" }
  ];
}

// Exemple d'utilisation avec les nouvelles options
export const handleGeneratePPTX = async (templateFile: File, questions: Question[]) => {
  try {
    await generatePPTX(templateFile, questions, {
      fileName: 'Mon_Quiz_Personnalise.pptx'
    });
  } catch (error: any) {
    console.error('Erreur:', error);
    alert(`Erreur lors de la génération: ${error.message}`);
  }
};