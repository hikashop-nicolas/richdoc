import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocxAdapter, docxToParts, htmlToDocx } from "./index";
import { readChartData } from "./chart-data";
import { specOf } from "./chart-spec";

const pts = (vals: (string | number)[]) => `<c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}`;
const ser = (i: number, name: string, col: string, vals: number[]) =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sheet1!$${col}$1</c:f><c:strCache>${pts([name])}</c:strCache></c:strRef></c:tx>` +
  `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache>${pts(["Q1", "Q2", "Q3"])}</c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>Sheet1!$${col}$2:$${col}$4</c:f><c:numCache><c:formatCode>General</c:formatCode>${pts(vals)}</c:numCache></c:numRef></c:val></c:ser>`;
const CHART = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, "North", "B", [10, 20, 30])}${ser(1, "South", "C", [5, 6, 7])}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart><c:externalData r:id="rId1"/></c:chartSpace>`;
const XLSX = zipSync({
  "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
  "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`),
  "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2"><v>10</v></c></row></sheetData></worksheet>`),
});
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const CT = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`;
function docx(withChart = true): Uint8Array {
  const run = withChart ? '<w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId7"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' : "<w:r><w:t>text</w:t></w:r>";
  return zipSync({
    "[Content_Types].xml": strToU8(CT),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(`<?xml version="1.0"?><w:document ${NS}><w:body><w:p>${run}</w:p><w:sectPr/></w:body></w:document>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${withChart ? '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>' : ""}</Relationships>`),
    ...(withChart ? {
      "word/charts/chart1.xml": strToU8(CHART),
      "word/charts/_rels/chart1.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>`),
      "word/embeddings/Microsoft_Excel_Worksheet.xlsx": XLSX,
    } : {}),
  });
}
const asDom = (html: string) => {
  const host = document.createElement("div");
  host.innerHTML = html;
  return host;
};

describe("docx charts", () => {
  it("draws a chart from its part and saves it untouched", () => {
    const d = docx();
    const { body } = docxToParts(d);
    expect(body).toMatch(/<span class="docx-chart" contenteditable="false" data-docx-xml="[^"]+" data-chart-part="word\/charts\/chart1.xml" data-chart-xml="[^"]+"><img src="data:image\/svg\+xml;charset=utf-8,%3Csvg[^"]*Sales[^"]*" alt="" style="width:576px;height:336px">/);
    const out = unzipSync(htmlToDocx(body, d));
    expect(strFromU8(out["word/charts/chart1.xml"]!)).toBe(CHART);
    expect(Array.from(out["word/embeddings/Microsoft_Excel_Worksheet.xlsx"]!)).toEqual(Array.from(XLSX));
  });

  it("edits a Word chart's data in its caches and embedded workbook", () => {
    const d = docx();
    const hooks = createDocxAdapter(d).charts!;
    const host = asDom(docxToParts(d).body);
    const chart = host.querySelector<HTMLElement>(".docx-chart")!;
    const m = hooks.read(chart)!;
    expect(m.fixed).toBe(true);
    expect(m.series[0]!.values).toEqual([10, 20, 30]);
    m.title = "Revenue";
    m.series[0]!.values[1] = 99;
    hooks.update(chart, m);
    expect(chart.querySelector("img")!.getAttribute("src")).toContain("Revenue");
    const out = unzipSync(htmlToDocx(host.innerHTML, d));
    const back = readChartData(strFromU8(out["word/charts/chart1.xml"]!))!;
    expect(back.title).toBe("Revenue");
    expect(back.series[0]!.values).toEqual([10, 99, 30]);
    const sheet = strFromU8(unzipSync(out["word/embeddings/Microsoft_Excel_Worksheet.xlsx"]!)["xl/worksheets/sheet1.xml"]!);
    expect(sheet).toContain('<c r="B3"><v>99</v></c>');
  });

  it("inserts a new chart part with its relationship and content type", () => {
    const d = docx(false);
    const hooks = createDocxAdapter(d).charts!;
    const chart = hooks.create();
    const m = hooks.read(chart)!;
    expect(m.fixed).toBe(false);
    m.type = "pie";
    m.title = "Share";
    m.categories = ["A", "B"];
    m.series = [{ name: "S", values: [3, 1] }];
    hooks.update(chart, m);
    const host = asDom("<p>text </p>");
    host.querySelector("p")!.appendChild(chart);
    const out = unzipSync(htmlToDocx(host.innerHTML, d));
    const part = strFromU8(out["word/charts/chart1.xml"]!);
    expect(specOf(part)).toMatchObject({ type: "pie", title: "Share", categories: ["A", "B"] });
    expect(strFromU8(out["word/_rels/document.xml.rels"]!)).toMatch(/Type="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships\/chart" Target="charts\/chart1.xml"/);
    expect(strFromU8(out["[Content_Types].xml"]!)).toContain('PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"');
    const xml = strFromU8(out["word/document.xml"]!);
    expect(xml).toMatch(/<c:chart [^>]*r:id="rId\d+"\/>/);
    // A theme-less package gets the Office theme, so other readers colour the chart.
    expect(strFromU8(out["word/theme/theme1.xml"]!)).toContain('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>');
    expect(strFromU8(out["word/_rels/document.xml.rels"]!)).toContain('Target="theme/theme1.xml"');
    expect(strFromU8(out["[Content_Types].xml"]!)).toContain('PartName="/word/theme/theme1.xml"');
    // Reopened, it reads back as a chart richdoc can edit freely.
    const re = asDom(docxToParts(htmlToDocx(host.innerHTML, d)).body).querySelector<HTMLElement>(".docx-chart")!;
    expect(createDocxAdapter(d).charts!.read(re)).toMatchObject({ fixed: false, type: "pie" });
  });
});
