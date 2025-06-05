import JSZip from 'jszip';
import { saveAs } from 'file-saver';

// ========== INTERFACES ==========
interface Question {
  question: string;
  correctAnswer?: boolean;
  duration?: number;
}

interface GenerationOptions {
  fileName?: string;
}

interface TagInfo {
  tagNumber: number;
  fileName: string;
  content: string;
}

interface RIdMapping {
  rId: string;
  type: string;
  target: string;
}

interface AppXmlMetadata {
  totalSlides: number;
  totalWords: number;
  totalParagraphs: number;
  slideTitles: string[];
}

// ========== FONCTIONS UTILITAIRES ==========

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
  const tagsFolder = zip.folder('ppt/tags');
  if (tagsFolder) {
    tagsFolder.forEach((relativePath) => {
      if (relativePath.match(/^tag\d+\.xml$/)) {
        count++;
      }
    });
  }
  return count;
}

// Vérifie si tag1.xml existe déjà
async function checkTag1Exists(zip: JSZip): Promise<boolean> {
  const tag1 = zip.file('ppt/tags/tag1.xml');
  return tag1 !== null;
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

// ========== GESTION DES LAYOUTS ==========

// Trouve le prochain slideLayout disponible après les existants
async function findNextAvailableSlideLayoutId(zip: JSZip): Promise<{ layoutId: number, layoutFileName: string, rId: string }> {
  const masterRelsFile = zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels');
  if (!masterRelsFile) {
    throw new Error('slideMaster1.xml.rels non trouvé');
  }
  
  const masterRelsContent = await masterRelsFile.async('string');
  
  // Trouver tous les slideLayouts existants
  const layoutMatches = masterRelsContent.match(/slideLayout(\d+)\.xml/g) || [];
  let maxLayoutNum = 0;
  
  layoutMatches.forEach(match => {
    const num = parseInt(match.match(/slideLayout(\d+)\.xml/)?.[1] || '0');
    if (num > maxLayoutNum) maxLayoutNum = num;
  });
  
  // Le prochain layout sera maxLayoutNum + 1
  const nextLayoutNum = maxLayoutNum + 1;
  
  // Trouver le prochain rId disponible dans slideMaster1.xml.rels
  const rIdMatches = masterRelsContent.match(/rId(\d+)/g) || [];
  let maxRId = 0;
  
  rIdMatches.forEach(match => {
    const num = parseInt(match.replace('rId', ''));
    if (num > maxRId) maxRId = num;
  });
  
  return {
    layoutId: nextLayoutNum,
    layoutFileName: `slideLayout${nextLayoutNum}.xml`,
    rId: `rId${maxRId + 1}`
  };
}
// ========== GESTION DES LAYOUTS (SUITE) ==========

// Créer ou vérifier l'existence d'un slideLayout OMBEA
async function ensureOmbeaSlideLayoutExists(zip: JSZip): Promise<{ layoutFileName: string, layoutRId: string }> {
  // CHANGEMENT : Au lieu de chercher un layout existant (qui pourrait être le mauvais),
  // on va toujours créer un nouveau layout OMBEA pour être sûr qu'il soit compatible
  
  console.log('Création d\'un layout OMBEA dédié...');
  const { layoutId, layoutFileName, rId } = await findNextAvailableSlideLayoutId(zip);
  
  // Contenu du slideLayout OMBEA avec la structure spécifique pour les questions
  const slideLayoutContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
  
  // Créer le fichier slideLayout
  zip.file(`ppt/slideLayouts/${layoutFileName}`, slideLayoutContent);
  
  // Créer le fichier .rels correspondant
  const slideLayoutRelsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
  
  zip.file(`ppt/slideLayouts/_rels/${layoutFileName}.rels`, slideLayoutRelsContent);
  
  // Mettre à jour slideMaster1.xml.rels
  await updateSlideMasterRelsForNewLayout(zip, layoutFileName, rId);
  
  // Mettre à jour slideMaster1.xml
  await updateSlideMasterForNewLayout(zip, layoutId, rId);
  
  // Mettre à jour [Content_Types].xml
  await updateContentTypesForNewLayout(zip, layoutFileName);
  
  console.log(`Layout OMBEA créé : ${layoutFileName} avec ${rId}`);
  
  return {
    layoutFileName: layoutFileName,
    layoutRId: rId
  };
}

// Mettre à jour slideMaster1.xml.rels pour le nouveau layout
async function updateSlideMasterRelsForNewLayout(zip: JSZip, layoutFileName: string, rId: string): Promise<void> {
  const masterRelsFile = zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels');
  if (masterRelsFile) {
    let content = await masterRelsFile.async('string');
    
    // Ajouter la nouvelle relation avant </Relationships>
    const insertPoint = content.lastIndexOf('</Relationships>');
    const newRel = `\n  <Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/${layoutFileName}"/>`;
    content = content.slice(0, insertPoint) + newRel + '\n' + content.slice(insertPoint);
    
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', content);
  }
}

// Mettre à jour slideMaster1.xml pour le nouveau layout
async function updateSlideMasterForNewLayout(zip: JSZip, layoutId: number, rId: string): Promise<void> {
  const masterFile = zip.file('ppt/slideMasters/slideMaster1.xml');
  if (masterFile) {
    let content = await masterFile.async('string');
    
    // Trouver la section sldLayoutIdLst
    const layoutIdLstEnd = content.indexOf('</p:sldLayoutIdLst>');
    if (layoutIdLstEnd > -1) {
      // Générer un ID unique pour le layout (commencer à 2147483649)
      const layoutIdValue = 2147483648 + layoutId;
      
      const newLayoutId = `\n    <p:sldLayoutId id="${layoutIdValue}" r:id="${rId}"/>`;
      content = content.slice(0, layoutIdLstEnd) + newLayoutId + '\n  ' + content.slice(layoutIdLstEnd);
      
      zip.file('ppt/slideMasters/slideMaster1.xml', content);
    }
  }
}

// Mettre à jour [Content_Types].xml pour le nouveau layout
async function updateContentTypesForNewLayout(zip: JSZip, layoutFileName: string): Promise<void> {
  const contentTypesFile = zip.file('[Content_Types].xml');
  if (contentTypesFile) {
    let content = await contentTypesFile.async('string');
    
    // Vérifier si le layout est déjà dans Content_Types
    if (!content.includes(layoutFileName)) {
      // Trouver un bon endroit pour insérer (après les autres layouts)
      const lastLayoutIndex = content.lastIndexOf('slideLayout');
      if (lastLayoutIndex > -1) {
        const endOfLastLayout = content.indexOf('/>', lastLayoutIndex) + 2;
        const newOverride = `\n  <Override PartName="/ppt/slideLayouts/${layoutFileName}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`;
        content = content.slice(0, endOfLastLayout) + newOverride + content.slice(endOfLastLayout);
      }
      
      zip.file('[Content_Types].xml', content);
    }
  }
}

// ========== CRÉATION DES SLIDES ==========

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
              <a:t>${duration}</a:t>
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
// ========== GESTION DES TAGS ==========

// Génère le tag1.xml global pour OMBEA (toujours tag1)
function createGlobalTag(): TagInfo {
  return {
    tagNumber: 1,
    fileName: 'tag1.xml',
    content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tag name="OR_PRESENTATION" val="True"/>
</p:tagLst>`
  };
}

// Calcule le numéro de base pour les tags d'une slide
function calculateBaseTagNumber(slideNumber: number, isFirstSlideEver: boolean): number {
  // Si c'est la première slide du document, les tags commencent à 2 (après tag1 global)
  // Sinon, chaque slide a 4 tags, donc : 2 + (slideNumber - 1) * 4
  if (isFirstSlideEver) {
    return 2; // tag2, tag3, tag4, tag5 pour la première slide
  }
  return 2 + (slideNumber - 1) * 4;
}

// Génère les 4 fichiers tags pour une slide OMBEA
function createSlideTagFiles(
  slideNumber: number, 
  isFirstSlideEver: boolean,
  correctAnswer: boolean, 
  duration: number = 30
): TagInfo[] {
  const baseTagNumber = calculateBaseTagNumber(slideNumber, isFirstSlideEver);
  const slideGuid = generateGUID();
  const points = correctAnswer ? "1.00,0.00" : "0.00,1.00";
  
  const tags: TagInfo[] = [];
  
  // Tag principal de la slide (configuration OMBEA)
  tags.push({
    tagNumber: baseTagNumber,
    fileName: `tag${baseTagNumber}.xml`,
    content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</p:tagLst>`
  });
  
  // Tag du titre
  tags.push({
    tagNumber: baseTagNumber + 1,
    fileName: `tag${baseTagNumber + 1}.xml`,
    content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tag name="OR_SHAPE_TYPE" val="OR_TITLE"/>
</p:tagLst>`
  });
  
  // Tag des réponses avec les points
  tags.push({
    tagNumber: baseTagNumber + 2,
    fileName: `tag${baseTagNumber + 2}.xml`,
    content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tag name="OR_SHAPE_TYPE" val="OR_ANSWERS"/>
  <p:tag name="OR_ANSWER_POINTS" val="${points}"/>
  <p:tag name="OR_ANSWERS_TEXT" val="Vrai&#13;Faux"/>
  <p:tag name="OR_EXCEL_ANSWER_COLORS" val="-10838489,-14521195"/>
</p:tagLst>`
  });
  
  // Tag du countdown
  tags.push({
    tagNumber: baseTagNumber + 3,
    fileName: `tag${baseTagNumber + 3}.xml`,
    content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tag name="OR_SHAPE_TYPE" val="OR_COUNTDOWN"/>
</p:tagLst>`
  });
  
  return tags;
}

// Génère le fichier .rels pour une slide OMBEA avec les bons tags
function createSlideRelsXml(
  slideNumber: number, 
  isFirstSlideEver: boolean,
  layoutFileName: string
): string {
  const baseTagNumber = calculateBaseTagNumber(slideNumber, isFirstSlideEver);
  
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagNumber}.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagNumber + 1}.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagNumber + 2}.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagNumber + 3}.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/${layoutFileName}"/>
</Relationships>`;
}
// ========== GESTION DES RID ==========

// Analyse et extrait tous les rId existants d'un fichier .rels
function extractExistingRIds(relsContent: string): RIdMapping[] {
  const mappings: RIdMapping[] = [];
  const relationshipRegex = /<Relationship\s+Id="(rId\d+)"\s+Type="([^"]+)"\s+Target="([^"]+)"/g;
  
  let match;
  while ((match = relationshipRegex.exec(relsContent)) !== null) {
    mappings.push({
      rId: match[1],
      type: match[2],
      target: match[3]
    });
  }
  
  return mappings;
}

// Trouve le prochain rId disponible
function getNextAvailableRId(existingRIds: string[]): string {
  let maxId = 0;
  
  existingRIds.forEach(rId => {
    const num = parseInt(rId.replace('rId', ''));
    if (num > maxId) maxId = num;
  });
  
  return `rId${maxId + 1}`;
}

// ========== MISES À JOUR XML ==========

// Met à jour presentation.xml avec les nouvelles slides
// Remplacer la fonction updatePresentationXmlWithSlides par celle-ci :

function updatePresentationXmlWithSlides(
  originalContent: string, 
  existingSlideCount: number,
  slideRIdMappings: { slideNumber: number; rId: string }[],
  tag1RId: string // Nouveau paramètre pour passer le vrai rId de tag1
): string {
  let updatedContent = originalContent;
  
  // Ajouter la référence custDataLst pour tag1 si elle n'existe pas
  if (!updatedContent.includes('<p:custDataLst>')) {
    const insertPoint = updatedContent.lastIndexOf('</p:presentation>');
    // Utiliser le vrai rId au lieu de rIdTag1
    const custDataLst = `\n  <p:custDataLst>\n    <p:tags r:id="${tag1RId}"/>\n  </p:custDataLst>`;
    updatedContent = updatedContent.slice(0, insertPoint) + custDataLst + '\n' + updatedContent.slice(insertPoint);
  } else {
    // Si custDataLst existe, s'assurer que le bon rId est utilisé
    updatedContent = updatedContent.replace(/r:id="rIdTag1"/, `r:id="${tag1RId}"`);
  }
  
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
  
  // Trouver le plus grand ID de slide existant
  const slideIdMatches = existingSldIdLst.match(/id="(\d+)"/g) || [];
  let maxSlideId = 256; // Valeur par défaut
  
  slideIdMatches.forEach(match => {
    const id = parseInt(match.match(/id="(\d+)"/)?.[1] || '0');
    if (id > maxSlideId) maxSlideId = id;
  });
  
  // Créer les nouvelles entrées de slides
  let newSlideEntries = '';
  slideRIdMappings.forEach(mapping => {
    const slideId = maxSlideId + mapping.slideNumber - existingSlideCount;
    newSlideEntries += `\n    <p:sldId id="${slideId}" r:id="${mapping.rId}"/>`;
  });
  
  // Insérer les nouvelles slides avant la fermeture de sldIdLst
  const updatedSldIdLst = existingSldIdLst.replace('</p:sldIdLst>', newSlideEntries + '\n  </p:sldIdLst>');
  
  return beforeSldIdLst + updatedSldIdLst + afterSldIdLst;
}

// Modifier aussi updatePresentationRelsWithMappings pour retourner le rId de tag1 :

function updatePresentationRelsWithMappings(
  originalContent: string,
  newSlideCount: number,
  existingSlideCount: number,
  hasTag1: boolean
): { updatedContent: string; slideRIdMappings: { slideNumber: number; rId: string }[]; tag1RId: string } {
  // Extraire tous les rId existants
  const existingMappings = extractExistingRIds(originalContent);
  const existingRIds = existingMappings.map(m => m.rId);
  
  let updatedContent = originalContent;
  const slideRIdMappings: { slideNumber: number; rId: string }[] = [];
  
  // Ajouter tag1.xml si nécessaire
  let rIdTag1 = existingMappings.find(m => m.target.includes('tag1.xml'))?.rId;
  
  if (!hasTag1 && !rIdTag1) {
    rIdTag1 = getNextAvailableRId(existingRIds);
    existingRIds.push(rIdTag1);
    
    const insertPoint = updatedContent.lastIndexOf('</Relationships>');
    const newRel = `\n  <Relationship Id="${rIdTag1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="tags/tag1.xml"/>`;
    updatedContent = updatedContent.slice(0, insertPoint) + newRel + updatedContent.slice(insertPoint);
  }
  
  // Ajouter les relations pour les nouvelles slides
  let newRelationships = '';
  for (let i = 1; i <= newSlideCount; i++) {
    const slideNumber = existingSlideCount + i;
    const newRId = getNextAvailableRId(existingRIds);
    existingRIds.push(newRId);
    
    slideRIdMappings.push({
      slideNumber: slideNumber,
      rId: newRId
    });
    
    newRelationships += `\n  <Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNumber}.xml"/>`;
  }
  
  // Insérer les nouvelles relations
  if (newRelationships) {
    const insertPoint = updatedContent.lastIndexOf('</Relationships>');
    updatedContent = updatedContent.slice(0, insertPoint) + newRelationships + '\n' + updatedContent.slice(insertPoint);
  }
  
  return { updatedContent, slideRIdMappings, tag1RId: rIdTag1 || 'rId10' };
}

// Met à jour [Content_Types].xml avec tous les nouveaux éléments
function updateContentTypesComplete(
  originalContent: string, 
  newSlideCount: number, 
  existingSlideCount: number,
  layoutFileName: string,
  totalTags: number
): string {
  let updatedContent = originalContent;
  
  // Vérifier et ajouter le slideLayout si nécessaire
  if (!updatedContent.includes(layoutFileName)) {
    // Trouver un bon endroit pour insérer (après les autres layouts ou avant slides)
    const layoutInsertRegex = /(<Override[^>]*slideLayout\d+\.xml"[^>]*\/>)/g;
    const matches = Array.from(updatedContent.matchAll(layoutInsertRegex));
    
    if (matches.length > 0) {
      const lastMatch = matches[matches.length - 1];
      const insertPoint = lastMatch.index! + lastMatch[0].length;
      const newOverride = `\n  <Override PartName="/ppt/slideLayouts/${layoutFileName}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`;
      updatedContent = updatedContent.slice(0, insertPoint) + newOverride + updatedContent.slice(insertPoint);
    }
  }
  
  // Préparer les nouvelles entrées
  let newOverrides = '';
  
  // Ajouter les nouvelles slides
  for (let i = existingSlideCount + 1; i <= existingSlideCount + newSlideCount; i++) {
    if (!updatedContent.includes(`slide${i}.xml`)) {
      newOverrides += `\n  <Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
    }
  }
  
  // Ajouter tous les tags manquants
  for (let i = 1; i <= totalTags; i++) {
    if (!updatedContent.includes(`tag${i}.xml`)) {
      newOverrides += `\n  <Override PartName="/ppt/tags/tag${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/>`;
    }
  }
  
  // Insérer toutes les nouvelles entrées avant </Types>
  if (newOverrides) {
    const insertPoint = updatedContent.lastIndexOf('</Types>');
    updatedContent = updatedContent.slice(0, insertPoint) + newOverrides + '\n' + updatedContent.slice(insertPoint);
  }
  
  return updatedContent;
}
// ========== GESTION DE APP.XML ==========

// Calcule les métadonnées pour app.xml
function calculateAppXmlMetadata(
  existingSlideCount: number,
  questions: Question[]
): AppXmlMetadata {
  // Calculer les mots et paragraphes
  let totalWords = 0;
  let totalParagraphs = 0;
  const slideTitles: string[] = [];
  
  questions.forEach(q => {
    // Compter les mots dans la question + "Vrai" et "Faux"
    const questionWords = q.question.trim().split(/\s+/).filter(word => word.length > 0).length;
    totalWords += questionWords + 2; // +2 pour "Vrai" et "Faux"
    
    // Paragraphes : 1 pour le titre, 2 pour les réponses
    totalParagraphs += 3;
    
    // Ajouter le titre de la slide
    slideTitles.push(q.question);
  });
  
  return {
    totalSlides: existingSlideCount + questions.length,
    totalWords,
    totalParagraphs,
    slideTitles
  };
}

// Met à jour app.xml avec la structure correcte
async function updateAppXml(
  zip: JSZip, 
  metadata: AppXmlMetadata
): Promise<void> {
  const appFile = zip.file('docProps/app.xml');
  if (!appFile) {
    console.warn('app.xml non trouvé, création d\'un nouveau fichier');
    createNewAppXml(zip, metadata);
    return;
  }
  
  let content = await appFile.async('string');
  
  // Mettre à jour les champs simples
  content = updateSimpleFields(content, metadata);
  
  // Mettre à jour HeadingPairs et TitlesOfParts
  content = updateHeadingPairsAndTitles(content, metadata);
  
  zip.file('docProps/app.xml', content);
}

// Met à jour les champs simples dans app.xml
function updateSimpleFields(content: string, metadata: AppXmlMetadata): string {
  let updated = content;
  
  // Slides
  updated = updated.replace(/<Slides>\d+<\/Slides>/, `<Slides>${metadata.totalSlides}</Slides>`);
  
  // Words
  updated = updated.replace(/<Words>\d+<\/Words>/, `<Words>${metadata.totalWords}</Words>`);
  
  // Paragraphs
  updated = updated.replace(/<Paragraphs>\d+<\/Paragraphs>/, `<Paragraphs>${metadata.totalParagraphs}</Paragraphs>`);
  
  // TotalTime (garder la valeur existante ou mettre 2 par défaut)
  if (!updated.includes('<TotalTime>')) {
    const propertiesEnd = updated.indexOf('</Properties>');
    const totalTimeTag = '\n  <TotalTime>2</TotalTime>';
    updated = updated.slice(0, propertiesEnd) + totalTimeTag + '\n' + updated.slice(propertiesEnd);
  }
  
  // Company (s'assurer qu'elle existe)
  if (!updated.includes('<Company')) {
    // Insérer après TitlesOfParts ou HeadingPairs
    const insertPoint = updated.indexOf('</TitlesOfParts>');
    if (insertPoint > -1) {
      const companyTag = '\n  <Company/>';
      updated = updated.slice(0, insertPoint + '</TitlesOfParts>'.length) + companyTag + updated.slice(insertPoint + '</TitlesOfParts>'.length);
    }
  }
  
  return updated;
}

// Met à jour HeadingPairs et TitlesOfParts avec la structure correcte
function updateHeadingPairsAndTitles(content: string, metadata: AppXmlMetadata): string {
  let updated = content;
  
  // Extraire les titres existants si présents
  const existingTitles: string[] = [];
  const titlesMatch = content.match(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/);
  
  if (titlesMatch) {
    const titlesContent = titlesMatch[0];
    const titleRegex = /<vt:lpstr>([^<]+)<\/vt:lpstr>/g;
    let match;
    
    while ((match = titleRegex.exec(titlesContent)) !== null) {
      existingTitles.push(match[1]);
    }
  }
  
  // Filtrer pour ne garder que les titres non-slides (polices, thèmes, etc.)
  const nonSlideTitles = existingTitles.filter(title => 
    !title.includes('Slide ') && 
    !title.includes('Diapositive ') &&
    title !== 'PowerPoint Presentation' &&
    !metadata.slideTitles.some(st => title.includes(st.substring(0, 20)))
  );
  
  // Construire la nouvelle structure HeadingPairs
  const headingPairs = buildHeadingPairs(nonSlideTitles, metadata.slideTitles);
  
  // Construire la nouvelle structure TitlesOfParts
  const titlesOfParts = buildTitlesOfParts(nonSlideTitles, metadata.slideTitles);
  
  // Remplacer HeadingPairs
  const headingPairsRegex = /<HeadingPairs>[\s\S]*?<\/HeadingPairs>/;
  if (headingPairsRegex.test(updated)) {
    updated = updated.replace(headingPairsRegex, headingPairs);
  } else {
    // Insérer HeadingPairs si absent
    const insertPoint = updated.indexOf('<TitlesOfParts>');
    if (insertPoint > -1) {
      updated = updated.slice(0, insertPoint) + headingPairs + '\n  ' + updated.slice(insertPoint);
    }
  }
  
  // Remplacer TitlesOfParts
  const titlesOfPartsRegex = /<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/;
  if (titlesOfPartsRegex.test(updated)) {
    updated = updated.replace(titlesOfPartsRegex, titlesOfParts);
  } else {
    // Insérer TitlesOfParts si absent
    const insertPoint = updated.indexOf('</HeadingPairs>') + '</HeadingPairs>'.length;
    updated = updated.slice(0, insertPoint) + '\n  ' + titlesOfParts + updated.slice(insertPoint);
  }
  
  return updated;
}

// Construit la structure HeadingPairs correcte
function buildHeadingPairs(nonSlideTitles: string[], slideTitles: string[]): string {
  const pairs: string[] = [];
  
  // Ajouter les paires pour les éléments non-slides (polices, thèmes, etc.)
  if (nonSlideTitles.some(t => t.includes('Police') || t.includes('Font'))) {
    pairs.push(`
      <vt:variant>
        <vt:lpstr>Polices utilisées</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>2</vt:i4>
      </vt:variant>`);
  }
  
  if (nonSlideTitles.some(t => t.includes('Thème') || t.includes('Theme'))) {
    pairs.push(`
      <vt:variant>
        <vt:lpstr>Thème</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>1</vt:i4>
      </vt:variant>`);
  }
  
  // Ajouter la paire pour les titres de diapositives
  if (slideTitles.length > 0) {
    pairs.push(`
      <vt:variant>
        <vt:lpstr>Titres des diapositives</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>${slideTitles.length}</vt:i4>
      </vt:variant>`);
  }
  
  const vectorSize = pairs.length * 2; // Chaque paire compte pour 2 éléments
  
  return `<HeadingPairs>
    <vt:vector size="${vectorSize}" baseType="variant">${pairs.join('')}
    </vt:vector>
  </HeadingPairs>`;
}

// Construit la structure TitlesOfParts correcte
function buildTitlesOfParts(nonSlideTitles: string[], slideTitles: string[]): string {
  const allTitles: string[] = [];
  
  // Ajouter d'abord les titres non-slides dans l'ordre approprié
  if (nonSlideTitles.some(t => t.includes('Arial'))) {
    allTitles.push('Arial');
  }
  if (nonSlideTitles.some(t => t.includes('Calibri'))) {
    allTitles.push('Calibri');
  }
  if (nonSlideTitles.some(t => t.includes('Thème'))) {
    allTitles.push('Thème Office');
  }
  
  // Ajouter les titres des slides
  slideTitles.forEach(title => {
    // Tronquer les titres trop longs
    const truncatedTitle = title.length > 100 ? title.substring(0, 97) + '...' : title;
    allTitles.push(escapeXml(truncatedTitle));
  });
  
  const vectorContent = allTitles.map(title => 
    `\n      <vt:lpstr>${title}</vt:lpstr>`
  ).join('');
  
  return `<TitlesOfParts>
    <vt:vector size="${allTitles.length}" baseType="lpstr">${vectorContent}
    </vt:vector>
  </TitlesOfParts>`;
}

// Crée un nouveau fichier app.xml si nécessaire
function createNewAppXml(zip: JSZip, metadata: AppXmlMetadata): void {
  const headingPairs = buildHeadingPairs(['Arial', 'Calibri', 'Thème Office'], metadata.slideTitles);
  const titlesOfParts = buildTitlesOfParts(['Arial', 'Calibri', 'Thème Office'], metadata.slideTitles);
  
  const appXmlContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>2</TotalTime>
  <Words>${metadata.totalWords}</Words>
  <Application>Microsoft Office PowerPoint</Application>
  <PresentationFormat>Affichage à l'écran (4:3)</PresentationFormat>
  <Paragraphs>${metadata.totalParagraphs}</Paragraphs>
  <Slides>${metadata.totalSlides}</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  ${headingPairs}
  ${titlesOfParts}
  <Company/>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>14.0000</AppVersion>
</Properties>`;
  
  zip.file('docProps/app.xml', appXmlContent);
}

// ========== GESTION DE CORE.XML ==========

// Met à jour core.xml avec les nouvelles métadonnées
// Remplacer la fonction updateCoreXml existante par celle-ci :

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
    
    // NE PAS ajouter dcterms:created si elle existe déjà
    // (Le code original ajoutait une deuxième balise, ce qui causait la corruption)
    
    zip.file('docProps/core.xml', content);
  }
}
// ========== FONCTION PRINCIPALE ==========

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
    const hasExistingTag1 = await checkTag1Exists(templateZip);
    
    console.log(`Slides existantes dans le modèle: ${existingSlideCount}`);
    console.log(`Tags existants dans le modèle: ${existingTagCount}`);
    console.log(`Tag1 existe: ${hasExistingTag1}`);
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
    
    // S'assurer qu'un layout OMBEA existe
    const { layoutFileName, layoutRId } = await ensureOmbeaSlideLayoutExists(outputZip);
    console.log(`Layout OMBEA: ${layoutFileName} (${layoutRId})`);
    
    // Créer le dossier tags s'il n'existe pas
    outputZip.folder('ppt/tags');
    
    // Créer le tag global OMBEA si pas déjà présent
    let totalTagsCreated = existingTagCount;
    if (!hasExistingTag1) {
      const globalTag = createGlobalTag();
      outputZip.file(`ppt/tags/${globalTag.fileName}`, globalTag.content);
      totalTagsCreated = 1; // Reset à 1 car on commence avec tag1
      console.log('Tag global OMBEA créé');
    }
    
    // Créer les nouvelles slides OMBEA
    console.log('Création des nouvelles slides OMBEA...');
    for (let i = 0; i < questions.length; i++) {
      const slideNumber = existingSlideCount + i + 1;
      const question = questions[i];
      const correctAnswer = question.correctAnswer !== undefined ? question.correctAnswer : false;
      const duration = question.duration || 30;
      
      // Créer le fichier slide XML
      const slideXml = createSlideXml(question.question, slideNumber, duration);
      outputZip.file(`ppt/slides/slide${slideNumber}.xml`, slideXml);
      
      // Créer le fichier .rels pour la slide
      const isFirstSlideEver = (existingSlideCount === 0 && i === 0);
      const slideRelsXml = createSlideRelsXml(i + 1, isFirstSlideEver, layoutFileName);
      outputZip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, slideRelsXml);
      
      // Créer les fichiers tags pour cette slide
      const tags = createSlideTagFiles(i + 1, isFirstSlideEver, correctAnswer, duration);
      tags.forEach(tag => {
        outputZip.file(`ppt/tags/${tag.fileName}`, tag.content);
        totalTagsCreated = Math.max(totalTagsCreated, tag.tagNumber);
      });
      
      console.log(`Slide OMBEA ${slideNumber} créée: ${question.question.substring(0, 50)}...`);
    }
    
    console.log(`Total des tags créés: ${totalTagsCreated}`);
    
    // Mettre à jour [Content_Types].xml
    console.log('Mise à jour des métadonnées...');
    const contentTypesFile = outputZip.file('[Content_Types].xml');
    if (contentTypesFile) {
      const contentTypesContent = await contentTypesFile.async('string');
      const updatedContentTypes = updateContentTypesComplete(
        contentTypesContent, 
        questions.length, 
        existingSlideCount,
        layoutFileName,
        totalTagsCreated
      );
      outputZip.file('[Content_Types].xml', updatedContentTypes);
    }
    
    // Mettre à jour presentation.xml.rels
    const presentationRelsFile = outputZip.file('ppt/_rels/presentation.xml.rels');
    if (presentationRelsFile) {
      const presentationRelsContent = await presentationRelsFile.async('string');
      const { updatedContent: updatedPresentationRels, slideRIdMappings, tag1RId } = updatePresentationRelsWithMappings(
        presentationRelsContent,
        questions.length,
        existingSlideCount,
        !hasExistingTag1
      );
      outputZip.file('ppt/_rels/presentation.xml.rels', updatedPresentationRels);
      
      // Mettre à jour presentation.xml avec les mappings corrects et le bon rId pour tag1
      const presentationFile = outputZip.file('ppt/presentation.xml');
      if (presentationFile) {
        const presentationContent = await presentationFile.async('string');
        const updatedPresentation = updatePresentationXmlWithSlides(
          presentationContent,
          existingSlideCount,
          slideRIdMappings,
          tag1RId // Passer le vrai rId de tag1
        );
        outputZip.file('ppt/presentation.xml', updatedPresentation);
      }
    }
    
    // Mettre à jour core.xml
    await updateCoreXml(outputZip, questions.length);
    
    // Calculer et mettre à jour app.xml
    const appMetadata = calculateAppXmlMetadata(existingSlideCount, questions);
    await updateAppXml(outputZip, appMetadata);
    
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
    console.log(`Total des tags: ${totalTagsCreated}`);
  } catch (error: any) {
    console.error('Erreur lors de la génération:', error);
    throw new Error(`Génération échouée: ${error.message}`);
  }
}

// ========== FONCTIONS UTILITAIRES POUR TEST ==========

// Fonction utilitaire pour tester avec des données d'exemple
export function createTestQuestions(): Question[] {
  return [
    { question: "Paris est-elle la capitale de la France ?", correctAnswer: true, duration: 30 },
    { question: "Le soleil tourne-t-il autour de la Terre ?", correctAnswer: false, duration: 30 },
    { question: "L'eau bout-elle à 100°C au niveau de la mer ?", correctAnswer: true, duration: 45 },
    { question: "JavaScript est-il un langage de programmation ?", correctAnswer: true, duration: 20 },
    { question: "Les pingouins vivent-ils au pôle Nord ?", correctAnswer: false, duration: 30 }
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

// ========== EXPORTS ==========

export type {
  Question,
  GenerationOptions,
  TagInfo,
  RIdMapping,
  AppXmlMetadata
};
