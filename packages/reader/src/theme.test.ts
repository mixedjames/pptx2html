import { describe, expect, it } from 'vitest';
import { emptyTheme, parseTheme } from './theme.js';

const THEME_XML = `<a:theme xmlns:a="a" name="Office Theme">
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
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="67000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="81000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
        <a:noFill/>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;

describe('parseTheme', () => {
  it("parses the fmtScheme's fillStyleLst into exactly 3 Fills, in order", () => {
    const theme = parseTheme(THEME_XML);
    expect(theme.formatScheme.fillStyles).toHaveLength(3);
    expect(theme.formatScheme.fillStyles[0]).toMatchObject({
      type: 'solid',
      color: { type: 'scheme', value: 'phClr' },
    });
    expect(theme.formatScheme.fillStyles[1]).toMatchObject({ type: 'gradient' });
    expect(theme.formatScheme.fillStyles[2]).toEqual({ type: 'none' });
  });

  it("parses the fmtScheme's lnStyleLst into exactly 3 Lines, in order", () => {
    const theme = parseTheme(THEME_XML);
    expect(theme.formatScheme.lineStyles).toHaveLength(3);
    expect(theme.formatScheme.lineStyles.map((line) => line.width)).toEqual([12700, 19050, 25400]);
    expect(theme.formatScheme.lineStyles[0]?.fill).toMatchObject({
      type: 'solid',
      color: { type: 'scheme', value: 'phClr' },
    });
  });

  it('returns empty fill/line style lists when fmtScheme has no fillStyleLst/lnStyleLst', () => {
    const theme = parseTheme('<a:theme xmlns:a="a"><a:themeElements/></a:theme>');
    expect(theme.formatScheme.fillStyles).toEqual([]);
    expect(theme.formatScheme.lineStyles).toEqual([]);
  });

  it('emptyTheme also has empty fill/line style lists', () => {
    expect(emptyTheme().formatScheme).toEqual({ name: '', fillStyles: [], lineStyles: [] });
  });
});
