import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import format from 'xml-formatter';

interface Question {
  question: string;
  correctAnswer: boolean;
  imagePath?: string;
  duration?: number;
}

const baseContentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`;

const basePresentation = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle/>
  <p:custDataLst>
    <p:tags r:id="rId3"/>
  </p:custDataLst>
</p:presentation>`;

const baseRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const basePresentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="/ppt/slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="/ppt/slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="tags/tag1.xml"/>
</Relationships>`;

const baseSlide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
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
              <a:rPr lang="en-US" dirty="0" smtClean="0"/>
              <a:t>Question Title</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
  <p:custDataLst>
    <p:tags r:id="rId1"/>
  </p:custDataLst>
</p:sld>`;

const baseAppProps = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office PowerPoint</Application>
  <AppVersion>16.0000</AppVersion>
</Properties>`;

const baseCoreProps = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
                   xmlns:dc="http://purl.org/dc/elements/1.1/" 
                   xmlns:dcterms="http://purl.org/dc/terms/" 
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>OMBEA Interactive Presentation</dc:title>
  <dc:creator>OMBEA PowerPoint Generator</dc:creator>
  <cp:lastModifiedBy>OMBEA PowerPoint Generator</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>
</cp:coreProperties>`;

const createSlideRelsXml = (slideNumber: number) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${(slideNumber - 1) * 4 + 2}.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${(slideNumber - 1) * 4 + 3}.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${(slideNumber - 1) * 4 + 4}.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${(slideNumber - 1) * 4 + 5}.xml"/>
</Relationships>`;

const createTagPresentationXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:tag name="OR_PRESENTATION" val="1"/>
</p:tagLst>`;

const createTagQuestionXml = (slideNumber: number, question: string, correctAnswer: boolean, duration: number) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:tag name="OR_SLIDE_TYPE" val="OR_QUESTION_SLIDE"/>
  <p:tag name="OR_QUESTION_TEXT" val="${question}"/>
  <p:tag name="OR_ANSWERS_TEXT" val="Vrai|Faux"/>
  <p:tag name="OR_ANSWER_POINTS" val="${correctAnswer ? '1.00,0.00' : '0.00,1.00'}"/>
  <p:tag name="OR_POLL_TIME_LIMIT" val="${duration}"/>
  <p:tag name="OR_CHART_COLOR_MODE" val="Color_Scheme"/>
  <p:tag name="OR_SHAPE_TYPE" val="OR_TITLE|OR_ANSWERS|OR_COUNTDOWN"/>
</p:tagLst>`;

const createTagTitleXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:tag name="OR_SHAPE_TYPE" val="OR_TITLE"/>
</p:tagLst>`;

const createTagAnswersXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:tag name="OR_SHAPE_TYPE" val="OR_ANSWERS"/>
</p:tagLst>`;

const createTagCountdownXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:tag name="OR_SHAPE_TYPE" val="OR_COUNTDOWN"/>
</p:tagLst>`;

export async function generatePPTX(templateFile: File | null, questions: Question[], defaultDuration: number): Promise<void> {
  return new Promise(async (resolve, reject) => {
    try {
      const zip = new JSZip();
      
      // Create base directory structure
      zip.folder('ppt');
      zip.folder('ppt/slides');
      zip.folder('ppt/slides/_rels');
      zip.folder('ppt/tags');
      zip.folder('ppt/_rels');
      zip.folder('ppt/theme');
      zip.folder('ppt/slideLayouts');
      zip.folder('ppt/slideMasters');
      zip.folder('docProps');
      zip.folder('_rels');
      
      // Add base files
      zip.file('[Content_Types].xml', baseContentTypes);
      zip.file('_rels/.rels', baseRels);
      zip.file('docProps/app.xml', baseAppProps);
      zip.file('docProps/core.xml', baseCoreProps);
      zip.file('ppt/presentation.xml', basePresentation);
      zip.file('ppt/_rels/presentation.xml.rels', basePresentationRels);
      
      // Add presentation tag file (tag1.xml)
      zip.file('ppt/tags/tag1.xml', createTagPresentationXml());
      
      // Process each question and create slides
      for (let i = 0; i < questions.length; i++) {
        const slideNumber = i + 1;
        const question = questions[i];
        const duration = question.duration || defaultDuration;
        
        // Add slide files
        zip.file(`ppt/slides/slide${slideNumber}.xml`, baseSlide);
        zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, createSlideRelsXml(slideNumber));
        
        // Calculate tag indices for this slide
        const baseTagIndex = (slideNumber - 1) * 4 + 2; // Start from tag2 for first slide
        
        // Add tag files
        zip.file(`ppt/tags/tag${baseTagIndex}.xml`, createTagQuestionXml(slideNumber, question.question, question.correctAnswer, duration));
        zip.file(`ppt/tags/tag${baseTagIndex + 1}.xml`, createTagTitleXml());
        zip.file(`ppt/tags/tag${baseTagIndex + 2}.xml`, createTagAnswersXml());
        zip.file(`ppt/tags/tag${baseTagIndex + 3}.xml`, createTagCountdownXml());
      }
      
      // Generate the new PPTX file
      const outputZip = await zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      });
      
      // Save the file
      saveAs(outputZip, `OMBEA_Questions_${new Date().toISOString().slice(0, 10)}.pptx`);
      
      resolve();
    } catch (error) {
      console.error('Error generating PPTX:', error);
      reject(error);
    }
  });
}