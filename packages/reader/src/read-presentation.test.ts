import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { readPresentation } from './read-presentation.js';

const encoder = new TextEncoder();

function xml(strings: TemplateStringsArray, ...values: unknown[]): Uint8Array {
  return encoder.encode(String.raw({ raw: strings }, ...values));
}

const CONTENT_TYPES = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="wav" ContentType="audio/wav"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;

const ROOT_RELS = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const PRESENTATION_XML = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:notesMasterIdLst><p:notesMasterId r:id="rId3"/></p:notesMasterIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle>
    <a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr>
  </p:defaultTextStyle>
</p:presentation>`;

const PRESENTATION_RELS = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>
</Relationships>`;

const THEME_XML = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;

const SLIDE_MASTER_XML = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld name="Master">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:p><a:r><a:t>Master title</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr><a:defRPr sz="4400" b="1"/></a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr>
    </p:bodyStyle>
  </p:txStyles>
</p:sldMaster>`;

const SLIDE_MASTER_RELS = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

const SLIDE_LAYOUT_XML = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" matchingName="Title Slide" showMasterSp="0">
  <p:cSld name="Title Slide">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

const SLIDE_LAYOUT_RELS = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

const SLIDE_XML = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title 1"/></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:p><a:r><a:t>Hello, presentation!</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="3" name="Picture 1"/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId2"/></p:blipFill>
        <p:spPr/>
      </p:pic>
    </p:spTree>
  </p:cSld>
  <p:transition spd="med" advClick="0" advTm="4000">
    <p:fade thruBlk="1"/>
    <p:sndAc>
      <p:stSnd loop="0"><p:snd r:embed="rId4"/></p:stSnd>
    </p:sndAc>
  </p:transition>
  <p:timing>
    <p:tnLst>
      <p:par>
        <p:cTn id="1" nodeType="tmRoot">
          <p:childTnLst>
            <p:animEffect transition="in" filter="fade">
              <p:cBhvr>
                <p:cTn id="2" dur="500"/>
                <p:tgtEl><p:spTgt spid="2"/></p:tgtEl>
              </p:cBhvr>
            </p:animEffect>
          </p:childTnLst>
        </p:cTn>
      </p:par>
    </p:tnLst>
  </p:timing>
</p:sld>`;

const SLIDE_RELS = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="../media/audio1.wav"/>
</Relationships>`;

const NOTES_MASTER_XML = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:notesMaster>`;

const NOTES_SLIDE_XML = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:p><a:r><a:t>Speaker notes</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>`;

const NOTES_SLIDE_RELS = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/>
</Relationships>`;

const IMAGE_BYTES = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const AUDIO_BYTES = new Uint8Array([0x52, 0x49, 0x46, 0x46]);

function buildFixturePptx(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': CONTENT_TYPES,
    '_rels/.rels': ROOT_RELS,
    'ppt/presentation.xml': PRESENTATION_XML,
    'ppt/_rels/presentation.xml.rels': PRESENTATION_RELS,
    'ppt/theme/theme1.xml': THEME_XML,
    'ppt/slideMasters/slideMaster1.xml': SLIDE_MASTER_XML,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': SLIDE_MASTER_RELS,
    'ppt/slideLayouts/slideLayout1.xml': SLIDE_LAYOUT_XML,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': SLIDE_LAYOUT_RELS,
    'ppt/slides/slide1.xml': SLIDE_XML,
    'ppt/slides/_rels/slide1.xml.rels': SLIDE_RELS,
    'ppt/notesMasters/notesMaster1.xml': NOTES_MASTER_XML,
    'ppt/notesSlides/notesSlide1.xml': NOTES_SLIDE_XML,
    'ppt/notesSlides/_rels/notesSlide1.xml.rels': NOTES_SLIDE_RELS,
    'ppt/media/image1.png': IMAGE_BYTES,
    'ppt/media/audio1.wav': AUDIO_BYTES,
  });
}

describe('readPresentation (synthetic end-to-end fixture)', () => {
  it('assembles the full object graph from a minimal in-memory .pptx', () => {
    const presentation = readPresentation(buildFixturePptx());

    expect(presentation.slideSize).toEqual({ width: 12192000, height: 6858000 });
    expect(presentation.notesSize).toEqual({ width: 6858000, height: 9144000 });

    expect(presentation.slideMasters).toHaveLength(1);
    const master = presentation.slideMasters[0]!;
    expect(master.theme.name).toBe('Office Theme');
    expect(master.theme.colorScheme.accent1).toEqual({ type: 'srgb', value: '4F81BD' });
    expect(master.layouts).toHaveLength(1);
    expect(master.textStyles?.titleStyle?.levels[0]).toEqual({
      runProperties: { fontSize: 4400, bold: true },
    });
    expect(master.textStyles?.bodyStyle?.levels[0]).toEqual({
      runProperties: { fontSize: 2800 },
    });
    expect(presentation.defaultTextStyle?.levels[0]).toEqual({
      runProperties: { fontSize: 1800 },
    });

    expect(presentation.slides).toHaveLength(1);
    const slide = presentation.slides[0]!;

    // The slide's layout must be the exact same object instance owned by its master, not a
    // separately re-parsed copy.
    expect(slide.layout).toBe(master.layouts[0]);
    expect(slide.layout.type).toBe('title');
    expect(slide.layout.matchingName).toBe('Title Slide');
    expect(slide.layout.master).toBe(master);

    const [titleShape, picture] = slide.commonSlideData.shapeTree;
    expect(titleShape).toMatchObject({
      kind: 'shape',
      nonVisual: { name: 'Title 1' },
    });
    if (titleShape?.kind === 'shape') {
      expect(titleShape.textBody?.paragraphs[0]?.runs[0]).toEqual({
        kind: 'run',
        text: 'Hello, presentation!',
      });
    }
    expect(picture).toMatchObject({
      kind: 'picture',
      nonVisual: { name: 'Picture 1' },
      image: { contentType: 'image/png' },
    });
    if (picture?.kind === 'picture') {
      expect(picture.image.data).toEqual(IMAGE_BYTES);
    }

    expect(slide.transition).toEqual({
      speed: 'med',
      advanceOnClick: false,
      advanceAfter: 4000,
      effect: { kind: 'fade', throughBlack: true },
      sound: {
        kind: 'play',
        sound: { contentType: 'audio/wav', data: AUDIO_BYTES },
        loop: false,
      },
    });

    expect(slide.timing?.timeNodeTree).toMatchObject({
      kind: 'par',
      common: { id: 1, role: 'tmRoot' },
      children: [
        {
          kind: 'animEffect',
          transition: 'in',
          filter: 'fade',
          target: { kind: 'shape', shapeId: 2 },
        },
      ],
    });

    expect(presentation.notesMaster?.commonSlideData.shapeTree).toEqual([]);
    expect(presentation.notesSlides).toHaveLength(1);
    const notesSlide = presentation.notesSlides[0]!;
    expect(notesSlide.master).toBe(presentation.notesMaster);
    // The notes slide must reuse the exact Slide instance built for presentation.slides, not a
    // second parse of slide1.xml.
    expect(notesSlide.slide).toBe(slide);
    expect(notesSlide.commonSlideData.shapeTree[0]).toMatchObject({
      kind: 'shape',
      nonVisual: { name: 'Notes Placeholder' },
    });
  });
});
