import JSZip from 'jszip';
import { saveAs } from 'file-saver';

interface Question {
  question: string;
  correctAnswer: boolean;
  imagePath?: string;
  duration?: number;
}
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

const defaultPresPropsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;

const defaultViewPropsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr/>
  <p:slideViewPr/>
  <p:notesTextViewPr/>
</p:viewPr>`;

const createContentTypes = (questionCount: number) => {
  let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="fntdata" ContentType="application/vnd.ms-opentype"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout12.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`;

  // Add types for each slide
  for (let i = 1; i <= questionCount; i++) {
    contentTypes += `
  <Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  }

  // Add types for all tags (1 presentation + 4 per question)
  const totalTags = 1 + (questionCount * 4);
  for (let i = 1; i <= totalTags; i++) {
    contentTypes += `
  <Override PartName="/ppt/tags/tag${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/>`;
  }

  contentTypes += `
</Types>`;
  return contentTypes;
};

const baseRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const baseAppProps = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>OMBEA PowerPoint Generator</Application>
  <AppVersion>16.0000</AppVersion>
</Properties>`;

const defaultThemeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Thème Office">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1>
<a:sysClr val="windowText" lastClr="000000"/>
</a:dk1>
<a:lt1>
<a:sysClr val="window" lastClr="FFFFFF"/>
</a:lt1>
<a:dk2>
<a:srgbClr val="1F497D"/>
</a:dk2>
<a:lt2>
<a:srgbClr val="EEECE1"/>
</a:lt2>
<a:accent1>
<a:srgbClr val="4F81BD"/>
</a:accent1>
<a:accent2>
<a:srgbClr val="C0504D"/>
</a:accent2>
<a:accent3>
<a:srgbClr val="9BBB59"/>
</a:accent3>
<a:accent4>
<a:srgbClr val="8064A2"/>
</a:accent4>
<a:accent5>
<a:srgbClr val="4BACC6"/>
</a:accent5>
<a:accent6>
<a:srgbClr val="F79646"/>
</a:accent6>
<a:hlink>
<a:srgbClr val="0000FF"/>
</a:hlink>
<a:folHlink>
<a:srgbClr val="800080"/>
</a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont>
<a:latin typeface="Calibri"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
<a:font script="Hang" typeface="맑은 고딕"/>
<a:font script="Hans" typeface="宋体"/>
<a:font script="Hant" typeface="新細明體"/>
<a:font script="Arab" typeface="Times New Roman"/>
<a:font script="Hebr" typeface="Times New Roman"/>
<a:font script="Thai" typeface="Angsana New"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="MoolBoran"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Times New Roman"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:majorFont>
<a:minorFont>
<a:latin typeface="Calibri"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
<a:font script="Hang" typeface="맑은 고딕"/>
<a:font script="Hans" typeface="宋体"/>
<a:font script="Hant" typeface="新細明體"/>
<a:font script="Arab" typeface="Arial"/>
<a:font script="Hebr" typeface="Arial"/>
<a:font script="Thai" typeface="Cordia New"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="DaunPenh"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Arial"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="50000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="35000">
<a:schemeClr val="phClr">
<a:tint val="37000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:tint val="15000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="1"/>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:shade val="51000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="80000">
<a:schemeClr val="phClr">
<a:shade val="93000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="94000"/>
<a:satMod val="135000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="0"/>
</a:gradFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr">
<a:shade val="95000"/>
<a:satMod val="105000"/>
</a:schemeClr>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="38000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
<a:scene3d>
<a:camera prst="orthographicFront">
<a:rot lat="0" lon="0" rev="0"/>
</a:camera>
<a:lightRig rig="threePt" dir="t">
<a:rot lat="0" lon="0" rev="1200000"/>
</a:lightRig>
</a:scene3d>
<a:sp3d>
<a:bevelT w="63500" h="25400"/>
</a:sp3d>
</a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="40000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="40000">
<a:schemeClr val="phClr">
<a:tint val="45000"/>
<a:shade val="99000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="20000"/>
<a:satMod val="255000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
</a:path>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="80000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="30000"/>
<a:satMod val="200000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
</a:path>
</a:gradFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>`;

const defaultSlideMasterRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

const defaultSlideMasterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" acc1="acc1" acc2="acc2" acc3="acc3" acc4="acc4" acc5="acc5" acc6="acc6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649"/>
  </p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`;

const defaultSlideLayoutRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

const defaultSlideLayoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="tx" preserve="1">
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

      <!-- Titre -->
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

      <!-- Corps du texte -->
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
          <a:p><a:pPr lvl="0"/><a:r><a:rPr lang="fr-FR" smtClean="0"/><a:t>Modifiez les styles du texte du masque</a:t></a:r></a:p>
          <a:p><a:pPr lvl="1"/><a:r><a:rPr lang="fr-FR" smtClean="0"/><a:t>Deuxième niveau</a:t></a:r></a:p>
          <a:p><a:pPr lvl="2"/><a:r><a:rPr lang="fr-FR" smtClean="0"/><a:t>Troisième niveau</a:t></a:r></a:p>
          <a:p><a:pPr lvl="3"/><a:r><a:rPr lang="fr-FR" smtClean="0"/><a:t>Quatrième niveau</a:t></a:r></a:p>
          <a:p><a:pPr lvl="4"/><a:r><a:rPr lang="fr-FR" smtClean="0"/><a:t>Cinquième niveau</a:t></a:r><a:endParaRPr lang="fr-FR"/></a:p>
        </p:txBody>
      </p:sp>

      <!-- Date -->
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

      <!-- Pied de page -->
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
          <a:p><a:endParaRPr lang="fr-FR"/></a:p>
        </p:txBody>
      </p:sp>

      <!-- Numéro de diapositive -->
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


const defaultTableStylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">
  <a:tblStyle styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" styleName="Medium Style 2 - Accent 1">
    <a:wholeTbl>
      <a:tcTxStyle>
        <a:fontRef idx="minor">
          <a:schemeClr val="lt1"/>
        </a:fontRef>
        <a:schemeClr val="dk1"/>
      </a:tcTxStyle>
      <a:tcStyle>
        <a:tcBdr>
          <a:left>
            <a:ln w="12700">
              <a:solidFill>
                <a:schemeClr val="accent1"/>
              </a:solidFill>
            </a:ln>
          </a:left>
          <a:right>
            <a:ln w="12700">
              <a:solidFill>
                <a:schemeClr val="accent1"/>
              </a:solidFill>
            </a:ln>
          </a:right>
          <a:top>
            <a:ln w="12700">
              <a:solidFill>
                <a:schemeClr val="accent1"/>
              </a:solidFill>
            </a:ln>
          </a:top>
          <a:bottom>
            <a:ln w="12700">
              <a:solidFill>
                <a:schemeClr val="accent1"/>
              </a:solidFill>
            </a:ln>
          </a:bottom>
        </a:tcBdr>
        <a:fill>
          <a:solidFill>
            <a:schemeClr val="lt1"/>
          </a:solidFill>
        </a:fill>
      </a:tcStyle>
    </a:wholeTbl>
  </a:tblStyle>
</a:tblStyleLst>`;

// FIXED: Core Properties with correct template literal
const createCoreProps = () => {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
                   xmlns:dc="http://purl.org/dc/elements/1.1/" 
                   xmlns:dcterms="http://purl.org/dc/terms/" 
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>OMBEA Interactive Presentation</dc:title>
  <dc:creator>OMBEA PowerPoint Generator</dc:creator>
  <cp:lastModifiedBy>OMBEA PowerPoint Generator</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
};

// FIXED: All interpolations corrected
const createSlideXml = (question: string, duration: number) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr>
          <p:custDataLst><p:tags r:id="rId1"/></p:custDataLst>
        </p:nvPr>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>

      <!-- Titre -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title Placeholder"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr>
            <p:ph type="title"/>
            <p:custDataLst><p:tags r:id="rId2"/></p:custDataLst>
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
              <a:rPr lang="en-US"/>
              <a:t>${escapeXml(question)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- Réponses -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Answers Placeholder"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr>
            <p:ph type="body" idx="1"/>
            <p:custDataLst><p:tags r:id="rId3"/></p:custDataLst>
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
              <a:rPr lang="en-US"/>
              <a:t>Vrai</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="0"/>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>Faux</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- Compte à rebours -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Countdown Timer"/>
          <p:cNvSpPr/>
          <p:nvPr>
            <p:custDataLst><p:tags r:id="rId4"/></p:custDataLst>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="10973100" y="457200"/>
            <a:ext cx="914400" cy="457200"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>${String(duration)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
    <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  </p:cSld>
</p:sld>`;

// CORRECTION 4: Ajouter le fichier tableStyles.xml manquant

const createSlideRelsXml = (slideNumber: number) => {
  const baseTagIndex = 2 + (slideNumber - 1) * 4;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <!-- Tag global de la présentation -->
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag1.xml"/>

  <!-- Tag de la question (slide) -->
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagIndex}.xml"/>

  <!-- Tag du titre -->
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagIndex + 1}.xml"/>

  <!-- Tag des réponses -->
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagIndex + 2}.xml"/>

  <!-- Tag du timer -->
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag${baseTagIndex + 3}.xml"/>

  <!-- Slide layout -->
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout12.xml"/>
</Relationships>`;
};


const createTagPresentationXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:tag name="OR_PRESENTATION" val="1"/>
</p:tagLst>`;

const createTagQuestionXml = (duration: number) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:tag name="OR_SLIDE_TYPE" val="OR_QUESTION_SLIDE"/>
  <p:tag name="OR_POLL_TIME_LIMIT" val="${duration}"/>
  <p:tag name="OR_CHART_COLOR_MODE" val="Color_Scheme"/>
</p:tagLst>`;

const createTagTitleXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:tag name="OR_SHAPE_TYPE" val="OR_TITLE"/>
</p:tagLst>`;

const createTagAnswersXml = (correctAnswer: boolean) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:tag name="OR_SHAPE_TYPE" val="OR_ANSWERS"/>
  <p:tag name="OR_ANSWERS_TEXT" val="Vrai Faux"/>
  <p:tag name="OR_ANSWER_POINTS" val="${correctAnswer ? '1.00,0.00' : '0.00,1.00'}"/>
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
      
      // Create base directory structure proactively
      zip.folder('ppt');
      zip.folder('ppt/slides');
      zip.folder('ppt/slides/_rels');
      zip.folder('ppt/tags');
      zip.folder('ppt/_rels');
      zip.folder('ppt/theme');
      zip.folder('ppt/slideLayouts');
      zip.folder('ppt/slideLayouts/_rels'); // Proactively create for default
      zip.folder('ppt/slideMasters');
      zip.folder('ppt/slideMasters/_rels'); 
      zip.folder('ppt/fonts');
      zip.folder('docProps');
      zip.folder('_rels');
      
      if (templateFile) {
        const templateZip = await JSZip.loadAsync(templateFile);

        const filesToCopyDirectly = [
          'ppt/tableStyles.xml',
          'ppt/presProps.xml',
          'ppt/viewProps.xml'
        ];
        for (const filePath of filesToCopyDirectly) {
          const file = templateZip.file(filePath);
          if (file) {
            const content = await file.async('blob');
            zip.file(filePath, content);
          }
        }
        
        const foldersToCopy = ['ppt/theme', 'ppt/slideLayouts', 'ppt/slideMasters', 'ppt/fonts'];
        for (const folderPath of foldersToCopy) {
            const folder = templateZip.folder(folderPath);
            if (folder) {
                const promises: Promise<void>[] = [];
                folder.forEach((_relativePath, fileEntry) => { 
                    if (!fileEntry.dir) {
                        const promise = fileEntry.async('blob').then(content => {
                          zip.file(`${folderPath}/${_relativePath}`, content);
                        });
                        promises.push(promise);
                    }
                });
                await Promise.all(promises);
            }
        }
      }

      // Add fixed base files
      // CHANGEMENT 1: Utiliser la fonction de Content Types dynamique
      zip.file('[Content_Types].xml', createContentTypes(questions.length));
      
      zip.file('_rels/.rels', baseRels);
      zip.file('docProps/app.xml', baseAppProps);
      
      // CHANGEMENT 2: Utiliser la fonction corrigée pour core props
      zip.file('docProps/core.xml', createCoreProps());
      
      // CHANGEMENT 3: Ajouter tableStyles.xml obligatoire
      if (!zip.file('ppt/tableStyles.xml')) {
        zip.file('ppt/tableStyles.xml', defaultTableStylesXml);
      }
      
      zip.file('ppt/tags/tag1.xml', createTagPresentationXml());

      // Add default theme if not provided by template
      if (!zip.file('ppt/theme/theme1.xml')) {
        zip.file('ppt/theme/theme1.xml', defaultThemeXml);
      }

      // Add default slide master and its rels if not provided by template
      if (!zip.file('ppt/slideMasters/slideMaster1.xml')) {
        zip.file('ppt/slideMasters/slideMaster1.xml', defaultSlideMasterXml);
      }
      if (!zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels')) {
        zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', defaultSlideMasterRelsXml);
      }

      // Add default slide layout and its rels if not provided by template
      if (!zip.file('ppt/slideLayouts/slideLayout12.xml')) { // Changed from slideLayout1.xml
        zip.file('ppt/slideLayouts/slideLayout12.xml', defaultSlideLayoutXml); // Changed from slideLayout1.xml
      }
      if (!zip.file('ppt/slideLayouts/_rels/slideLayout12.xml.rels')) { // Changed from slideLayout1.xml.rels
        zip.folder('ppt/slideLayouts/_rels'); // Ensure folder exists if not already
        zip.file('ppt/slideLayouts/_rels/slideLayout12.xml.rels', defaultSlideLayoutRelsXml); // Changed from slideLayout1.xml.rels
      }
      // Add default presProps if not provided by template or if not copied
      if (!zip.file('ppt/presProps.xml')) {                         
        zip.file('ppt/presProps.xml', defaultPresPropsXml);         
      }

      // Add default viewProps if not provided by template or if not copied
      if (!zip.file('ppt/viewProps.xml')) {                        
        zip.file('ppt/viewProps.xml', defaultViewPropsXml);          
      }
      
      const presentationMasterRelId = "rIdMaster1"; 
      const presentationTagRelId = "rIdPresTag"; 

      let sldIdLstContent = '';
      let presRelXmlContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="../docProps/core.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="../docProps/app.xml"/>
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="../ppt/presentation.xml"/>
        <Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
        <Relationship Id="rIdPresTag" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="tags/tag1.xml"/>
        <Relationship Id="rIdTheme1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
        <Relationship Id="rIdTableStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>`;
      
      // Loop through questions to generate slides and their relationships
      // CHANGEMENT 4: Utiliser la fonction de slide corrigée
      for (let i = 0; i < questions.length; i++) {
        const slideNumber = i + 1;
        const questionItem = questions[i]; 
        const duration = questionItem.duration || defaultDuration;
        const slideRelId = `rIdSlide${slideNumber}`; 
        const slidePersistId = 255 + slideNumber; 

        sldIdLstContent += `<p:sldId id="${slidePersistId}" r:id="${slideRelId}"/>`;
        presRelXmlContent += `
  <Relationship Id="${slideRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNumber}.xml"/>`;
      
        // CHANGEMENT 5: Utiliser createSlideXml au lieu de baseSlide
        zip.file(`ppt/slides/slide${slideNumber}.xml`, createSlideXml(questionItem.question, duration));
        zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, createSlideRelsXml(slideNumber));
        
        // Tags pour chaque slide
        const baseTagIndex = 2 + (slideNumber - 1) * 4;
        zip.file(`ppt/tags/tag${baseTagIndex}.xml`, createTagQuestionXml(duration));
        zip.file(`ppt/tags/tag${baseTagIndex + 1}.xml`, createTagTitleXml());
        zip.file(`ppt/tags/tag${baseTagIndex + 2}.xml`, createTagAnswersXml(questionItem.correctAnswer));
        zip.file(`ppt/tags/tag${baseTagIndex + 3}.xml`, createTagCountdownXml());
      }
      
      presRelXmlContent += `
</Relationships>`;
      zip.file('ppt/_rels/presentation.xml.rels', presRelXmlContent);

      const dynamicPresentationXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="${presentationMasterRelId}"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>${sldIdLstContent}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle/>
  <p:custDataLst>
    <p:tags r:id="${presentationTagRelId}"/>
  </p:custDataLst>
</p:presentation>`;
      zip.file('ppt/presentation.xml', dynamicPresentationXml);
      
      const outputZip = await zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      });
      
      saveAs(outputZip, `OMBEA_Questions_${new Date().toISOString().slice(0, 10)}.pptx`);
    resolve();      
    } catch (error) {
      console.error('Erreur lors de la génération du PPTX :', error);
      reject(error);
    }
  });
}