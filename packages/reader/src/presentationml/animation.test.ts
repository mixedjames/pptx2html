import { describe, expect, it } from 'vitest';

import { parseXml } from '../xml/parse.js';
import { parseSlideTiming } from './animation.js';

function firstNode(xml: string) {
  const [root] = parseXml(xml);
  return root!;
}

const NS = 'xmlns:p="p" xmlns:a="a"';

describe('parseSlideTiming', () => {
  it('returns undefined when there is no p:timing element', () => {
    expect(parseSlideTiming(undefined)).toBeUndefined();
  });

  it('returns undefined for an empty p:timing', () => {
    const node = firstNode(`<p:timing ${NS}/>`);
    expect(parseSlideTiming(node)).toBeUndefined();
  });

  it('parses a full click-triggered fade-in as a par/seq/animEffect tree', () => {
    const node = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:par>
            <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
              <p:childTnLst>
                <p:seq concurrent="1">
                  <p:cTn id="2" nodeType="mainSeq">
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="3" presetClass="entr" presetID="10" presetSubtype="0" nodeType="clickEffect" fill="hold" restart="whenNotActive" autoRev="1" spd="150%" accel="50000" decel="20000" repeatCount="2000">
                          <p:stCondLst>
                            <p:cond delay="0" evt="onClick"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>
                          </p:stCondLst>
                          <p:endCondLst>
                            <p:cond delay="indefinite" tn="3"/>
                          </p:endCondLst>
                          <p:childTnLst>
                            <p:animEffect transition="in" filter="fade">
                              <p:cBhvr>
                                <p:cTn id="4" dur="500"/>
                                <p:tgtEl><p:spTgt spid="2"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                            </p:animEffect>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:seq>
              </p:childTnLst>
            </p:cTn>
          </p:par>
        </p:tnLst>
      </p:timing>`,
    );

    const timing = parseSlideTiming(node);
    expect(timing?.timeNodeTree).toEqual({
      kind: 'par',
      common: { id: 1, duration: 'indefinite', role: 'tmRoot' },
      children: [
        {
          kind: 'seq',
          common: { id: 2, role: 'mainSeq' },
          concurrent: true,
          children: [
            {
              kind: 'par',
              common: {
                id: 3,
                presetClass: 'entr',
                presetId: 10,
                presetSubtype: 0,
                role: 'clickEffect',
                fill: 'hold',
                restart: 'whenNotActive',
                autoReverse: true,
                speed: 150000,
                accelerate: 50000,
                decelerate: 20000,
                repeatCount: 2000,
                startConditions: [{ delay: 0, event: 'onClick', target: { kind: 'slide' } }],
                endConditions: [{ delay: 'indefinite', triggerTimeNodeId: 3 }],
              },
              children: [
                {
                  kind: 'animEffect',
                  common: { id: 4, duration: 500 },
                  target: { kind: 'shape', shapeId: 2 },
                  attributeNames: ['style.opacity'],
                  transition: 'in',
                  filter: 'fade',
                },
              ],
            },
          ],
        },
      ],
    });
  });

  it('parses p:set with a literal string value', () => {
    const node = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:set>
            <p:cBhvr>
              <p:cTn id="1"/>
              <p:tgtEl><p:spTgt spid="5"/></p:tgtEl>
              <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
            </p:cBhvr>
            <p:to><p:strVal val="visible"/></p:to>
          </p:set>
        </p:tnLst>
      </p:timing>`,
    );

    expect(parseSlideTiming(node)?.timeNodeTree).toEqual({
      kind: 'set',
      common: { id: 1 },
      target: { kind: 'shape', shapeId: 5 },
      attributeNames: ['style.visibility'],
      to: 'visible',
    });
  });

  it('parses p:set with a colour value', () => {
    const node = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:set>
            <p:cBhvr><p:cTn id="1"/></p:cBhvr>
            <p:to><p:clrVal><a:srgbClr val="FF0000"/></p:clrVal></p:to>
          </p:set>
        </p:tnLst>
      </p:timing>`,
    );

    expect(parseSlideTiming(node)?.timeNodeTree).toEqual({
      kind: 'set',
      common: { id: 1 },
      to: { type: 'srgb', value: 'FF0000' },
    });
  });

  it('parses p:anim with from/to/by attributes and a keyframe list', () => {
    const node = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:anim calcmode="lin" valueType="num" from="0.5" to="1" by="0.1">
            <p:cBhvr><p:cTn id="1"/></p:cBhvr>
            <p:tavLst>
              <p:tav tm="0"><p:val><p:fltVal val="0.5"/></p:val></p:tav>
              <p:tav tm="100000" fmla="#ppt_x"><p:val><p:fltVal val="1"/></p:val></p:tav>
            </p:tavLst>
          </p:anim>
        </p:tnLst>
      </p:timing>`,
    );

    expect(parseSlideTiming(node)?.timeNodeTree).toEqual({
      kind: 'anim',
      common: { id: 1 },
      calcMode: 'lin',
      valueType: 'num',
      from: '0.5',
      to: '1',
      by: '0.1',
      keyframes: [
        { time: 0, value: '0.5' },
        { time: 100000, value: '1', formula: '#ppt_x' },
      ],
    });
  });

  it('parses p:animClr from/to colours and colour-space/direction', () => {
    const node = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:animClr clrSpc="hsl" dir="cw">
            <p:cBhvr><p:cTn id="1"/></p:cBhvr>
            <p:from><a:srgbClr val="000000"/></p:from>
            <p:to><a:srgbClr val="FFFFFF"/></p:to>
          </p:animClr>
        </p:tnLst>
      </p:timing>`,
    );

    expect(parseSlideTiming(node)?.timeNodeTree).toEqual({
      kind: 'animClr',
      common: { id: 1 },
      colorSpace: 'hsl',
      direction: 'cw',
      from: { type: 'srgb', value: '000000' },
      to: { type: 'srgb', value: 'FFFFFF' },
    });
  });

  it('parses p:animMotion path/origin/rAng and from/to/by points', () => {
    const node = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:animMotion origin="layout" path="M 0 0 L 0.5 0.5 E" rAng="60000">
            <p:cBhvr><p:cTn id="1"/></p:cBhvr>
            <p:from x="0" y="0"/>
            <p:to x="50000" y="50000"/>
          </p:animMotion>
        </p:tnLst>
      </p:timing>`,
    );

    expect(parseSlideTiming(node)?.timeNodeTree).toEqual({
      kind: 'animMotion',
      common: { id: 1 },
      origin: 'layout',
      path: 'M 0 0 L 0.5 0.5 E',
      pathRotation: 60000,
      from: { x: 0, y: 0 },
      to: { x: 50000, y: 50000 },
    });
  });

  it('parses p:animRot from/to/by angle attributes', () => {
    const node = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:animRot from="0" to="5400000" by="1800000">
            <p:cBhvr><p:cTn id="1"/></p:cBhvr>
          </p:animRot>
        </p:tnLst>
      </p:timing>`,
    );

    expect(parseSlideTiming(node)?.timeNodeTree).toEqual({
      kind: 'animRot',
      common: { id: 1 },
      from: 0,
      to: 5400000,
      by: 1800000,
    });
  });

  it('parses p:animScale from/to/by points', () => {
    const node = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:animScale>
            <p:cBhvr><p:cTn id="1"/></p:cBhvr>
            <p:from x="100000" y="100000"/>
            <p:to x="150000" y="150000"/>
          </p:animScale>
        </p:tnLst>
      </p:timing>`,
    );

    expect(parseSlideTiming(node)?.timeNodeTree).toEqual({
      kind: 'animScale',
      common: { id: 1 },
      from: { x: 100000, y: 100000 },
      to: { x: 150000, y: 150000 },
    });
  });

  it('parses p:cmd type/cmd', () => {
    const node = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:cmd type="call" cmd="playAnimation">
            <p:cBhvr><p:cTn id="1"/></p:cBhvr>
          </p:cmd>
        </p:tnLst>
      </p:timing>`,
    );

    expect(parseSlideTiming(node)?.timeNodeTree).toEqual({
      kind: 'cmd',
      common: { id: 1 },
      type: 'call',
      command: 'playAnimation',
    });
  });

  it('parses p:audio and p:video as media nodes targeting a shape', () => {
    const node = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:video>
            <p:cMediaNode>
              <p:cTn id="1"/>
              <p:tgtEl><p:spTgt spid="9"/></p:tgtEl>
            </p:cMediaNode>
          </p:video>
        </p:tnLst>
      </p:timing>`,
    );

    expect(parseSlideTiming(node)?.timeNodeTree).toEqual({
      kind: 'video',
      common: { id: 1 },
      target: { kind: 'shape', shapeId: 9 },
    });
  });

  it('parses spTgt/bg and spTgt/txEl (paragraph and character range) targets', () => {
    const bg = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:set>
            <p:cBhvr><p:cTn id="1"/><p:tgtEl><p:spTgt spid="2"><p:bg/></p:spTgt></p:tgtEl></p:cBhvr>
          </p:set>
        </p:tnLst>
      </p:timing>`,
    );
    expect(parseSlideTiming(bg)?.timeNodeTree).toMatchObject({
      target: { kind: 'shapeBackground', shapeId: 2 },
    });

    const pRg = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:set>
            <p:cBhvr><p:cTn id="1"/><p:tgtEl><p:spTgt spid="2"><p:txEl><p:pRg st="0" end="1"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr>
          </p:set>
        </p:tnLst>
      </p:timing>`,
    );
    expect(parseSlideTiming(pRg)?.timeNodeTree).toMatchObject({
      target: { kind: 'shapeText', shapeId: 2, paragraphRange: [0, 1] },
    });

    const charRg = firstNode(
      `<p:timing ${NS}>
        <p:tnLst>
          <p:set>
            <p:cBhvr><p:cTn id="1"/><p:tgtEl><p:spTgt spid="2"><p:txEl><p:charRg st="3" end="7"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr>
          </p:set>
        </p:tnLst>
      </p:timing>`,
    );
    expect(parseSlideTiming(charRg)?.timeNodeTree).toMatchObject({
      target: { kind: 'shapeText', shapeId: 2, characterRange: [3, 7] },
    });
  });

  describe('build list', () => {
    it('parses a bldP defaulting to a whole build when @build is omitted', () => {
      const node = firstNode(`<p:timing ${NS}><p:bldLst><p:bldP spid="2"/></p:bldLst></p:timing>`);
      expect(parseSlideTiming(node)?.buildList).toEqual([
        { kind: 'paragraph', shapeId: 2, buildType: 'whole' },
      ]);
    });

    it('parses a by-paragraph bldP with its level/background/reverse attributes', () => {
      const node = firstNode(
        `<p:timing ${NS}>
          <p:bldLst>
            <p:bldP spid="3" build="p" bldLvl="2" animBg="0" autoUpdateAnimBg="0" rev="1"/>
          </p:bldLst>
        </p:timing>`,
      );
      expect(parseSlideTiming(node)?.buildList).toEqual([
        {
          kind: 'paragraph',
          shapeId: 3,
          buildType: 'byParagraph',
          buildLevel: 2,
          animateBackground: false,
          autoUpdateAnimBg: false,
          reverse: true,
        },
      ]);
    });

    it('parses bldDgm/bldChart/bldGraphic, preserving their raw build attribute', () => {
      const node = firstNode(
        `<p:timing ${NS}>
          <p:bldLst>
            <p:bldDgm spid="4" build="whole"/>
            <p:bldChart spid="5" build="series"/>
            <p:bldGraphic spid="6"/>
          </p:bldLst>
        </p:timing>`,
      );
      expect(parseSlideTiming(node)?.buildList).toEqual([
        { kind: 'diagram', shapeId: 4, buildType: 'whole' },
        { kind: 'chart', shapeId: 5, buildType: 'series' },
        { kind: 'graphic', shapeId: 6 },
      ]);
    });
  });
});
