import { describe, expect, it } from 'vitest';

import type { MediaResolver } from '../drawingml/fill.js';
import { parseXml } from '../xml/parse.js';
import { parseShapeTree } from './shape-tree.js';

const noMedia: MediaResolver = () => undefined;
const withImage: MediaResolver = (id) =>
  id === 'rId1' ? { contentType: 'image/png', data: new Uint8Array([9]) } : undefined;

describe('presentationml/shape-tree', () => {
  it('parses a mix of shape kinds in document order and recurses into groups', () => {
    const [spTree] = parseXml(
      `<p:spTree xmlns:p="p" xmlns:a="a" xmlns:r="r">
        <p:nvGrpSpPr><p:cNvPr id="1" name="tree"/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Title 1"/></p:nvSpPr>
          <p:spPr/>
        </p:sp>
        <p:cxnSp>
          <p:nvCxnSpPr><p:cNvPr id="3" name="Connector 1"/></p:nvCxnSpPr>
          <p:spPr/>
          <p:stCxn id="2" idx="1"/>
          <p:endCxn id="5" idx="3"/>
        </p:cxnSp>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="4" name="Chart 1"/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic>
        </p:graphicFrame>
        <p:grpSp>
          <p:nvGrpSpPr><p:cNvPr id="5" name="Group 1"/></p:nvGrpSpPr>
          <p:grpSpPr><a:xfrm><a:off x="1" y="1"/><a:ext cx="2" cy="2"/></a:xfrm></p:grpSpPr>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="6" name="Nested 1"/></p:nvSpPr>
            <p:spPr/>
          </p:sp>
        </p:grpSp>
      </p:spTree>`,
    );

    const nodes = parseShapeTree(spTree!, noMedia);
    expect(nodes.map((n) => n.kind)).toEqual(['shape', 'connector', 'graphicFrame', 'group']);

    const connector = nodes[1];
    expect(connector).toMatchObject({
      startConnection: { shapeId: 2, connectionSiteIndex: 1 },
      endConnection: { shapeId: 5, connectionSiteIndex: 3 },
    });

    const graphicFrame = nodes[2];
    expect(graphicFrame).toMatchObject({ graphic: { type: 'chart' } });

    const group = nodes[3];
    expect(group).toMatchObject({
      kind: 'group',
      transform: { offset: { x: 1, y: 1 }, extents: { width: 2, height: 2 } },
    });
    if (group?.kind === 'group') {
      expect(group.children).toHaveLength(1);
      expect(group.children[0]?.nonVisual.name).toBe('Nested 1');
    }
  });

  it('drops a picture whose image relationship cannot be resolved, keeps one that can', () => {
    const [spTree] = parseXml(
      `<p:spTree xmlns:p="p" xmlns:a="a" xmlns:r="r">
        <p:pic>
          <p:nvPicPr><p:cNvPr id="2" name="Broken"/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdMissing"/></p:blipFill>
          <p:spPr/>
        </p:pic>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="3" name="Picture 1"/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
          <p:spPr/>
        </p:pic>
      </p:spTree>`,
    );

    const nodes = parseShapeTree(spTree!, withImage);
    expect(nodes).toHaveLength(1);
    expect(nodes[0]).toMatchObject({ kind: 'picture', nonVisual: { name: 'Picture 1' } });
  });

  it("parses a picture whose blip has no direct r:embed, only an extLst-nested extension blip (PowerPoint's SVG icons)", () => {
    const [spTree] = parseXml(
      `<p:spTree xmlns:p="p" xmlns:a="a" xmlns:r="r">
        <p:pic>
          <p:nvPicPr><p:cNvPr id="2" name="Icon 1"/></p:nvPicPr>
          <p:blipFill>
            <a:blip>
              <a:extLst>
                <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
                  <asvg:svgBlip xmlns:asvg="a" r:embed="rId1"/>
                </a:ext>
              </a:extLst>
            </a:blip>
          </p:blipFill>
          <p:spPr/>
        </p:pic>
      </p:spTree>`,
    );

    const nodes = parseShapeTree(spTree!, withImage);
    expect(nodes).toHaveLength(1);
    expect(nodes[0]).toMatchObject({ kind: 'picture', nonVisual: { name: 'Icon 1' } });
  });

  it('drops a group or graphicFrame missing a required transform', () => {
    const [spTree] = parseXml(
      `<p:spTree xmlns:p="p" xmlns:a="a">
        <p:grpSp><p:nvGrpSpPr/><p:grpSpPr/></p:grpSp>
        <p:graphicFrame><p:nvGraphicFramePr/></p:graphicFrame>
      </p:spTree>`,
    );
    expect(parseShapeTree(spTree!, noMedia)).toEqual([]);
  });

  it("parses a shape's p:style fillRef/lnRef/fontRef (PowerPoint's Shape Styles gallery, no explicit spPr fill/line or run colour)", () => {
    const [spTree] = parseXml(
      `<p:spTree xmlns:p="p" xmlns:a="a">
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Oval 1"/></p:nvSpPr>
          <p:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr>
          <p:style>
            <a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>
            <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
            <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
            <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
          </p:style>
        </p:sp>
      </p:spTree>`,
    );
    const [shape] = parseShapeTree(spTree!, noMedia);
    expect(shape).toMatchObject({
      style: {
        fillRef: { index: 1, color: { type: 'scheme', value: 'accent1' } },
        lineRef: {
          index: 2,
          color: { type: 'scheme', value: 'accent1', transforms: { shade: 15000 } },
        },
        fontRef: { collection: 'minor', color: { type: 'scheme', value: 'lt1' } },
      },
    });
  });

  it('parses fontRef with idx="major"/"none", and omits an unrecognized idx value', () => {
    const [spTree] = parseXml(
      `<p:spTree xmlns:p="p" xmlns:a="a">
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Shape 1"/></p:nvSpPr>
          <p:spPr/>
          <p:style><a:fontRef idx="major"><a:schemeClr val="dk1"/></a:fontRef></p:style>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="3" name="Shape 2"/></p:nvSpPr>
          <p:spPr/>
          <p:style><a:fontRef idx="none"><a:schemeClr val="dk1"/></a:fontRef></p:style>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="4" name="Shape 3"/></p:nvSpPr>
          <p:spPr/>
          <p:style><a:fontRef idx="bogus"><a:schemeClr val="dk1"/></a:fontRef></p:style>
        </p:sp>
      </p:spTree>`,
    );
    const [major, none, bogus] = parseShapeTree(spTree!, noMedia);
    expect(major).toMatchObject({
      style: { fontRef: { collection: 'major', color: { type: 'scheme', value: 'dk1' } } },
    });
    expect(none).toMatchObject({
      style: { fontRef: { collection: 'none', color: { type: 'scheme', value: 'dk1' } } },
    });
    expect(bogus).not.toHaveProperty('style');
  });

  it('parses a picture/connector p:style, and omits style entirely when absent', () => {
    const [spTree] = parseXml(
      `<p:spTree xmlns:p="p" xmlns:a="a" xmlns:r="r">
        <p:pic>
          <p:nvPicPr><p:cNvPr id="2" name="Picture 1"/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
          <p:spPr/>
          <p:style><a:fillRef idx="3"><a:schemeClr val="accent2"/></a:fillRef></p:style>
        </p:pic>
        <p:cxnSp>
          <p:nvCxnSpPr><p:cNvPr id="3" name="Connector 1"/></p:nvCxnSpPr>
          <p:spPr/>
        </p:cxnSp>
      </p:spTree>`,
    );
    const [picture, connector] = parseShapeTree(spTree!, withImage);
    expect(picture).toMatchObject({
      style: { fillRef: { index: 3, color: { type: 'scheme', value: 'accent2' } } },
    });
    expect(connector).not.toHaveProperty('style');
  });
});
