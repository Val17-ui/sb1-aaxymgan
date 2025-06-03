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

// Fonction pour générer un GUID unique (format UUID v4)
function generateGUID(): string {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16).toUpperCase();
  });
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

// Génère le XML d'une nouvelle slide OMBEA
function createSlideXml(question: string, slideNumber: number, duration: number = 30): string {
  // Utiliser slideNumber pour éviter l'avertissement
  const slideComment = `<!-- Slide ${slideNumber} -->`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
${slideComment}
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
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
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Titre 1"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="title"/>
            <p:custDataLst>
              <p:tags r:id="rId2"/>
            </p:custDataLst>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="fr-FR" dirty="0" smtClean="0"/>
              <a:t>${escapeXml(question)}</a:t>
            </a:r>
            <a:endParaRPr lang="fr-FR" dirty="0"/>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Espace réservé du texte 2"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="body" idx="1"/>
            <p:custDataLst>
              <p:tags r:id="rId3"/>
            </p:custDataLst>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="1600200"/>
            <a:ext cx="4572000" cy="4525963"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="514350" indent="-514350">
              <a:buAutoNum type="arabicPeriod"/>
            </a:pPr>
            <a:r>
              <a:rPr lang="fr-FR" dirty="0" smtClean="0"/>
              <a:t>Vrai</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr marL="514350" indent="-514350">
              <a:buAutoNum type="arabicPeriod"/>
            </a:pPr>
            <a:r>
              <a:rPr lang="fr-FR" dirty="0" smtClean="0"/>
              <a:t>Faux</a:t>
            </a:r>
            <a:endParaRPr lang="fr-FR" dirty="0"/>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="OMBEA Countdown"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr>
            <p:custDataLst>
              <p:tags r:id="rId4"/>
            </p:custDataLst>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="317500" y="5715000"/>
            <a:ext cx="1524000" cy="769441"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr vert="horz" rtlCol="0" anchor="ctr" anchorCtr="1">
            <a:spAutoFit/>
          </a:bodyPr>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="fr-FR" sz="4400" smtClean="0"/>
              <a:t>30</a:t>
            </a:r>
            <a:endParaRPr lang="fr-FR" sz="4400"/>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
    <p:custDataLst>
      <p:tags r:id="rId1"/>
    </p:custDataLst>
    <p:extLst>
      <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
        <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="${Math.floor(Math.random() * 10000000000)}"/>
      </p:ext>
    </p:extLst>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
  <p:timing>
    <p:tnLst>
      <p:par>
        <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/>
      </p:par>
    </p:tnLst>
  </p:timing>
</p:sld>`;
}

// Génère le fichier .rels pour une slide OMBEA
function createSlideRelsXml(slideNumber: number, existingSlideCount: number): string {
  const baseTagId = (existingSlideCount + slideNumber - 1) * 4 + 2;
  
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagId}.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagId + 1}.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagId + 2}.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagId + 3}.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout12.xml"/>
</Relationships>`;
}

// Génère les fichiers tags OMBEA
function createTagFiles(slideNumber: number, existingSlideCount: number, correctAnswer: boolean, duration: number = 30): { [key: string]: string } {
  const baseTagId = (existingSlideCount + slideNumber - 1) * 4 + 2;
  const slideGuid = generateGUID();
  const points = correctAnswer ? "1.00,0.00" : "0.00,1.00";
  
  const tags: { [key: string]: string } = {};
  
  // Tag principal de la slide
  tags[`tag${baseTagId}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tag name="OR_SLIDE_GUID" val="${slideGuid}"/>
  <p:tag name="OR_OFFICE_MAJOR_VERSION" val="14"/>
  <p:tag name="OR_POLL_START_MODE" val="Automatic"/>
  <p:tag name="OR_CHART_VALUE_LABEL_FORMAT" val="Response_Count"/>
  <p:tag name="OR_CHART_RESPONSE_DENOMINATOR" val="Responses"/>
  <p:tag name="OR_CHART_FIXED_RESPONSE_DENOMINATOR" val="100"/>
  <p:tag name="OR_CHART_COLOR_MODE" val="Color_Scheme"/>
  <p:tag name="OR_CHART_APPLY_OMBEA_TEMPLATE" val="True"/>
  <p:tag name="OR_POLL_DEFAULT_ANSWER_OPTION" val="None"/>
  <p:tag name="OR_SLIDE_TYPE" val="OR_QUESTION_SLIDE"/>
  <p:tag name="OR_ANSWERS_BULLET_STYLE" val="ppBulletArabicPeriod"/>
  <p:tag name="OR_POLL_FLOW" val="Automatic"/>
  <p:tag name="OR_CHART_DISPLAY_MODE" val="Automatic"/>
  <p:tag name="OR_POLL_TIME_LIMIT" val="${duration}"/>
  <p:tag name="OR_POLL_COUNTDOWN_START_MODE" val="Automatic"/>
  <p:tag name="OR_POLL_MULTIPLE_RESPONSES" val="1"/>
  <p:tag name="OR_POLL_DUPLICATES_ALLOWED" val="False"/>
  <p:tag name="OR_CATEGORIZING" val="False"/>
  <p:tag name="OR_PRIORITY_RANKING" val="False"/>
  <p:tag name="OR_IS_POLLED" val="False"/>
</p:tagLst>`;
  
  // Tag du titre
  tags[`tag${baseTagId + 1}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tag name="OR_SHAPE_TYPE" val="OR_TITLE"/>
</p:tagLst>`;
  
  // Tag des réponses
  tags[`tag${baseTagId + 2}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tag name="OR_SHAPE_TYPE" val="OR_ANSWERS"/>
  <p:tag name="OR_ANSWER_POINTS" val="${points}"/>
  <p:tag name="OR_ANSWERS_TEXT" val="Vrai Faux"/>
  <p:tag name="OR_EXCEL_ANSWER_COLORS" val="-10838489,-14521195"/>
</p:tagLst>`;
  
  // Tag du countdown
  tags[`tag${baseTagId + 3}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tag name="OR_SHAPE_TYPE" val="OR_COUNTDOWN"/>
</p:tagLst>`;
  
  return tags;
}

// Crée le tag1.xml global pour OMBEA
function createGlobalTag(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tag name="OR_PRESENTATION" val="True"/>
</p:tagLst>`;
}

// Vérifie si slideLayout12.xml existe dans le template
async function ensureSlideLayout12Exists(zip: JSZip): Promise<void> {
  const layout12 = zip.file('ppt/slideLayouts/slideLayout12.xml');
  if (!layout12) {
    // Créer slideLayout12.xml s'il n'existe pas
    const slideLayout12Content = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="tx" preserve="1">
  <p:cSld name="Titre et texte">
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
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Titre 1"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="title"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="fr-FR" smtClean="0"/>
              <a:t>Modifiez le style du titre</a:t>
            </a:r>
            <a:endParaRPr lang="fr-FR"/>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Espace réservé du texte 2"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="body" idx="1"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr lvl="0"/>
            <a:r>
              <a:rPr lang="fr-FR" smtClean="0"/>
              <a:t>Modifiez les styles du texte du masque</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="1"/>
            <a:r>
              <a:rPr lang="fr-FR" smtClean="0"/>
              <a:t>Deuxième niveau</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="2"/>
            <a:r>
              <a:rPr lang="fr-FR" smtClean="0"/>
              <a:t>Troisième niveau</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="3"/>
            <a:r>
              <a:rPr lang="fr-FR" smtClean="0"/>
              <a:t>Quatrième niveau</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="4"/>
            <a:r>
              <a:rPr lang="fr-FR" smtClean="0"/>
              <a:t>Cinquième niveau</a:t>
            </a:r>
            <a:endParaRPr lang="fr-FR"/>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Espace réservé de la date 3"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="dt" sz="half" idx="10"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:fld id="{ABB4FD2C-0372-488A-B992-EB1BD753A34A}" type="datetimeFigureOut">
              <a:rPr lang="fr-FR" smtClean="0"/>
              <a:t>28/05/2025</a:t>
            </a:fld>
            <a:endParaRPr lang="fr-FR"/>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="5" name="Espace réservé du pied de page 4"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="ftr" sz="quarter" idx="11"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:endParaRPr lang="fr-FR"/>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="6" name="Espace réservé du numéro de diapositive 5"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="sldNum" sz="quarter" idx="12"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:fld id="{CD42254F-ACD2-467B-9045-5226EEC3B6AB}" type="slidenum">
              <a:rPr lang="fr-FR" smtClean="0"/>
              <a:t>‹N°›</a:t>
            </a:fld>
            <a:endParaRPr lang="fr-FR"/>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
    <p:extLst>
      <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
        <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="2131546393"/>
      </p:ext>
    </p:extLst>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>`;
    
    zip.file('ppt/slideLayouts/slideLayout12.xml', slideLayout12Content);
    
    // Créer aussi le fichier .rels correspondant
    const slideLayout12RelsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
    
    zip.file('ppt/slideLayouts/_rels/slideLayout12.xml.rels', slideLayout12RelsContent);
  }
}
// Met à jour presentation.xml.rels pour inclure les relations vers les nouvelles slides et tag1
// Remplacez la fonction updatePresentationRels par celle-ci :
function updatePresentationRels(originalContent: string, newSlideCount: number, existingSlideCount: number): string {
  let updatedContent = originalContent;
  
  // Ne PAS ajouter de relation vers tag1.xml dans presentation.xml.rels
  // Le tag1.xml est référencé depuis presentation.xml mais pas ici
  
  const insertPoint = updatedContent.lastIndexOf('</Relationships>');
  let newRelationships = '';
  
  // Trouver le plus grand rId existant
  const rIdMatches = originalContent.match(/rId(\d+)/g) || [];
  let maxRId = 0;
  rIdMatches.forEach(match => {
    const num = parseInt(match.replace('rId', ''));
    if (num > maxRId) maxRId = num;
  });
  
  // Ajouter les relations vers les nouvelles slides avec des rId séquentiels
  for (let i = 0; i < newSlideCount; i++) {
    const slideNum = existingSlideCount + i + 1;
    const rId = maxRId + i + 1;
    newRelationships += `\n  <Relationship Id="rId${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNum}.xml"/>`;
  }
  
  return updatedContent.slice(0, insertPoint) + newRelationships + '\n' + updatedContent.slice(insertPoint);
}

// Met à jour le slideMaster pour inclure slideLayout12
async function updateSlideMasterRels(zip: JSZip): Promise<void> {
  const masterRelsFile = zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels');
  if (masterRelsFile) {
    let content = await masterRelsFile.async('string');
    
    // Vérifier si slideLayout12 est déjà référencé
    if (!content.includes('slideLayout12.xml')) {
      // Trouver le plus grand rId pour les layouts
      const layoutMatches = content.match(/rId(\d+).*slideLayout/g) || [];
      let maxLayoutRId = 0;
      layoutMatches.forEach(match => {
        const num = parseInt(match.match(/rId(\d+)/)?.[1] || '0');
        if (num > maxLayoutRId) maxLayoutRId = num;
      });
      
      const insertPoint = content.lastIndexOf('</Relationships>');
      const newRel = `\n  <Relationship Id="rId${maxLayoutRId + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout12.xml"/>`;
      content = content.slice(0, insertPoint) + newRel + '\n' + content.slice(insertPoint);
      
      zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', content);
    }
  }
}

// Met à jour le slideMaster pour inclure la référence à slideLayout12
async function updateSlideMaster(zip: JSZip): Promise<void> {
  const masterFile = zip.file('ppt/slideMasters/slideMaster1.xml');
  if (masterFile) {
    let content = await masterFile.async('string');
    
    // Vérifier si slideLayout12 est déjà dans la liste
    if (!content.includes('2147483660')) {
      // Trouver la section sldLayoutIdLst
      const layoutIdLstEnd = content.indexOf('</p:sldLayoutIdLst>');
      if (layoutIdLstEnd > -1) {
        // Trouver le dernier layout ID
        const layoutMatches = content.match(/sldLayoutId id="(\d+)"/g) || [];
        let maxLayoutNum = 0;
        layoutMatches.forEach(match => {
          const num = parseInt(match.match(/id="(\d+)"/)?.[1] || '0');
          if (num > maxLayoutNum) maxLayoutNum = num;
        });
        
        const newLayoutId = `\n    <p:sldLayoutId id="2147483660" r:id="rId12"/>`;
        content = content.slice(0, layoutIdLstEnd) + newLayoutId + '\n  ' + content.slice(layoutIdLstEnd);
        
        zip.file('ppt/slideMasters/slideMaster1.xml', content);
      }
    }
  }
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

// Compte le nombre de tags existants
function countExistingTags(zip: JSZip): number {
  let count = 0;
  zip.folder('ppt/tags')?.forEach((relativePath) => {
    if (relativePath.match(/^tag\d+\.xml$/)) {
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
// Met à jour core.xml avec les nouvelles métadonnées
async function updateCoreXml(zip: JSZip, slideCount: number): Promise<void> {
  const coreFile = zip.file('docProps/core.xml');
  if (coreFile) {
    let content = await coreFile.async('string');
    
    // Mettre à jour le titre avec gestion du pluriel
    const title = `Quiz OMBEA - ${slideCount} question${slideCount > 1 ? 's' : ''}`;
    content = content.replace(/<dc:title>.*?<\/dc:title>/, `<dc:title>${escapeXml(title)}</dc:title>`);
    
    // Mettre à jour la date de modification
    const now = new Date().toISOString();
    content = content.replace(/<dcterms:modified.*?>.*?<\/dcterms:modified>/, 
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`);
    
    zip.file('docProps/core.xml', content);
  }
}

// Met à jour app.xml avec le nombre de slides
async function updateAppXml(zip: JSZip, totalSlides: number, questions: Question[]): Promise<void> {
  const appFile = zip.file('docProps/app.xml');
  if (appFile) {
    let content = await appFile.async('string');
    
    // Calculer les mots et paragraphes
    let totalWords = 0;
    let totalParagraphs = 0;
    questions.forEach(q => {
      totalWords += q.question.split(/\s+/).length + 2; // +2 pour "Vrai" et "Faux"
      totalParagraphs += 3; // 1 pour le titre, 2 pour les réponses
    });
    
    // Mettre à jour les champs
    content = content.replace(/<Slides>\d+<\/Slides>/, `<Slides>${totalSlides}</Slides>`);
    content = content.replace(/<Words>\d+<\/Words>/, `<Words>${totalWords}</Words>`);
    content = content.replace(/<Paragraphs>\d+<\/Paragraphs>/, `<Paragraphs>${totalParagraphs}</Paragraphs>`);
    content = content.replace(/<TotalTime>\d+<\/TotalTime>/, `<TotalTime>2</TotalTime>`);
    
    zip.file('docProps/app.xml', content);
  }
}
// Fonction principale de génération avec support OMBEA
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
    
    // Compter les slides et tags existants
    const existingSlideCount = countExistingSlides(templateZip);
    const existingTagCount = countExistingTags(templateZip);
    
    console.log(`Slides existantes dans le modèle: ${existingSlideCount}`);
    console.log(`Tags existants dans le modèle: ${existingTagCount}`);
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
    
    // S'assurer que slideLayout12 existe
    await ensureSlideLayout12Exists(outputZip);
    
    // Mettre à jour le slideMaster pour inclure slideLayout12
    await updateSlideMasterRels(outputZip);
    await updateSlideMaster(outputZip);
    
    // Créer le dossier tags s'il n'existe pas
    outputZip.folder('ppt/tags');
    
    // Créer le tag global OMBEA si pas déjà présent
    if (existingTagCount === 0) {
      outputZip.file('ppt/tags/tag1.xml', createGlobalTag());
    }
    
    // Créer les nouvelles slides OMBEA
    console.log('Création des nouvelles slides OMBEA...');
    for (let i = 0; i < questions.length; i++) {
      const slideNumber = existingSlideCount + i + 1;
      const question = questions[i];
      const correctAnswer = question.correctAnswer !== undefined ? question.correctAnswer : false;
      const duration = question.duration || 30;
      
      // Créer le fichier slide XML
      const slideXml = createSlideXml(question.question, slideNumber, 30);
      outputZip.file(`ppt/slides/slide${slideNumber}.xml`, slideXml);
      
      // Créer le fichier .rels pour la slide
      const slideRelsXml = createSlideRelsXml(i + 1, existingSlideCount);
      outputZip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, slideRelsXml);
      
      // Créer les fichiers tags pour cette slide
      const tags = createTagFiles(i + 1, existingSlideCount, correctAnswer, duration);
      Object.entries(tags).forEach(([fileName, content]) => {
        outputZip.file(`ppt/tags/${fileName}`, content);
      });
      
      console.log(`Slide OMBEA ${slideNumber} créée: ${question.question.substring(0, 50)}...`);
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
        // Mettre à jour core.xml et app.xml
        await updateCoreXml(outputZip, questions.length);
        await updateAppXml(outputZip, existingSlideCount + questions.length, questions);
    // Générer le fichier final
    console.log('Génération du fichier final...');
    const outputBlob = await outputZip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    const fileName = options.fileName || `Questions_OMBEA_${new Date().toISOString().slice(0, 10)}.pptx`;
    saveAs(outputBlob, fileName);
    
    console.log(`Fichier OMBEA généré avec succès: ${fileName}`);
    console.log(`Total des slides: ${existingSlideCount + questions.length}`);
    console.log(`Total des tags: ${existingTagCount + 1 + (questions.length * 4)}`);
  } catch (error: any) {
    console.error('Erreur lors de la génération:', error);
    throw new Error(`Génération échouée: ${error.message}`);
  }
}


function updateContentTypes(originalContent: string, newSlideCount: number, existingSlideCount: number): string {
  let updatedContent = originalContent;
  
  // Vérifier si slideLayout12 est déjà présent, sinon l'ajouter
  if (!updatedContent.includes('slideLayout12.xml')) {
    const layoutInsertPoint = updatedContent.lastIndexOf('slideLayout11.xml');
    if (layoutInsertPoint > -1) {
      const endOfLayout11 = updatedContent.indexOf('/>', layoutInsertPoint) + 2;
      const layout12Override = `\n  <Override PartName="/ppt/slideLayouts/slideLayout12.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`;
      updatedContent = updatedContent.slice(0, endOfLayout11) + layout12Override + updatedContent.slice(endOfLayout11);
    }
  }
  
  // Ajouter les nouvelles slides et tags dans [Content_Types].xml
  const insertPoint = updatedContent.lastIndexOf('</Types>');
  let newOverrides = '';
  
  // Ajouter les slides
  for (let i = existingSlideCount + 1; i <= existingSlideCount + newSlideCount; i++) {
    newOverrides += `\n  <Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  }
  
  // Ajouter le tag global si pas déjà présent
  if (!updatedContent.includes('tag1.xml')) {
    newOverrides += `\n  <Override PartName="/ppt/tags/tag1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/>`;
  }
  
  // Ajouter les tags pour chaque nouvelle slide (4 tags par slide)
  for (let i = 0; i < newSlideCount; i++) {
    const baseTagId = (existingSlideCount + i) * 4 + 2; // Commence à tag6 pour slide2
    for (let j = 0; j < 4; j++) {
      newOverrides += `\n  <Override PartName="/ppt/tags/tag${baseTagId + j}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/>`;
    }
  }
  
  return updatedContent.slice(0, insertPoint) + newOverrides + '\n' + updatedContent.slice(insertPoint);
}

// Met à jour presentation.xml pour inclure les nouvelles slides et les tags OMBEA
function updatePresentationXml(originalContent: string, newSlideCount: number, existingSlideCount: number): string {
  let updatedContent = originalContent;
   
  // Trouver la section sldIdLst
  const sldIdLstStart = updatedContent.indexOf('<p:sldIdLst>');
  const sldIdLstEnd = updatedContent.indexOf('</p:sldIdLst>') + '</p:sldIdLst>'.length;
  
  if (sldIdLstStart === -1 || sldIdLstEnd === -1) {
    throw new Error('Structure presentation.xml invalide - section sldIdLst introuvable');
  }
  
 // Extraire la section existante
 const beforeSldIdLst = updatedContent.slice(0, sldIdLstStart);
 const existingSldIdLst = updatedContent.slice(sldIdLstStart, sldIdLstEnd);
 const afterSldIdLst = updatedContent.slice(sldIdLstEnd);
 
 // Trouver le plus grand rId existant dans presentation.xml.rels
 const presentationRelsFile = updatedContent; // Nous avons besoin de connaître maxRId
 let maxExistingRId = 2; // Par défaut, slide1 a rId2
 
 // Créer les nouvelles entrées de slides avec les bons rId
 let newSlideEntries = '';
 for (let i = 0; i < newSlideCount; i++) {
   const slideNum = existingSlideCount + i + 1;
   const slideId = 256 + slideNum; // 257 pour slide2, 258 pour slide3, etc.
   const rId = maxExistingRId + i + 1; // rId3 pour slide2, rId4 pour slide3, etc.
   newSlideEntries += `\n    <p:sldId id="${slideId}" r:id="rId${rId}"/>`;
 }
 
 // Insérer les nouvelles slides avant la fermeture de sldIdLst
 const updatedSldIdLst = existingSldIdLst.replace('</p:sldIdLst>', newSlideEntries + '\n  </p:sldIdLst>');
 
 return beforeSldIdLst + updatedSldIdLst + afterSldIdLst;
}
// Fonction utilitaire pour tester avec des données d'exemple
export function createTestQuestions(): Question[] {
  return [
    { question: "Paris est-elle la capitale de la France ?", correctAnswer: true },
    { question: "Le soleil tourne-t-il autour de la Terre ?", correctAnswer: false },
    { question: "L'eau bout-elle à 100°C au niveau de la mer ?", correctAnswer: true },
    { question: "JavaScript est-il un langage de programmation ?", correctAnswer: true },
    { question: "Les pingouins vivent-ils au pôle Nord ?", correctAnswer: false }
  ];
}

// Exemple d'utilisation avec les nouvelles options
export const handleGeneratePPTX = async (templateFile: File, questions: Question[]) => {
  try {
    await generatePPTX(templateFile, questions, {
      fileName: 'Quiz_OMBEA_Interactif.pptx'
    });
  } catch (error: any) {
    console.error('Erreur:', error);
    alert(`Erreur lors de la génération: ${error.message}`);
  }
};