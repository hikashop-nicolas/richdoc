import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readChartData, updateWorkbook, writeChartData } from "./chart-data";

const pts = (vals: (string | number)[]) => `<c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}`;
const ser = (i: number, name: string, col: string, vals: number[]) =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sheet1!$${col}$1</c:f><c:strCache>${pts([name])}</c:strCache></c:strRef></c:tx>` +
  `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache>${pts(["Q1", "Q2", "Q3"])}</c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>Sheet1!$${col}$2:$${col}$4</c:f><c:numCache><c:formatCode>General</c:formatCode>${pts(vals)}</c:numCache></c:numRef></c:val></c:ser>`;
const CHART = `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/>${ser(0, "North", "B", [10, 20, 30])}${ser(1, "South", "C", [5, 6, 7])}</c:barChart></c:plotArea></c:chart><c:externalData r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></c:chartSpace>`;
const SHEET = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="B1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>10</v></c><c r="C2" s="3"><v>5</v></c></row></sheetData></worksheet>`;
const XLSX = zipSync({
  "xl/workbook.xml": strToU8(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>`),
  "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`),
  "xl/worksheets/sheet1.xml": strToU8(SHEET),
});

describe("chart data", () => {
  it("reads the table from the caches", () => {
    expect(readChartData(CHART)).toEqual({ title: "Sales", categories: ["Q1", "Q2", "Q3"], series: [{ name: "North", values: [10, 20, 30] }, { name: "South", values: [5, 6, 7] }] });
  });

  it("writes edited values, names, labels and the title back into the caches", () => {
    const xml = writeChartData(CHART, { title: "Revenue", categories: ["Jan", "Feb", "Mar"], series: [{ name: "East", values: [1, null, 3.5] }, { name: "South", values: [5, 6, 7] }] });
    const back = readChartData(xml)!;
    expect(back.title).toBe("Revenue");
    expect(back.categories).toEqual(["Jan", "Feb", "Mar"]);
    expect(back.series[0]).toEqual({ name: "East", values: [1, null, 3.5] });
    expect(xml).toContain("<c:formatCode>General</c:formatCode>");
    expect(xml).toContain('<c:ptCount val="3"/>'); // the shape is kept
    expect(xml).toContain("<c:f>Sheet1!$B$2:$B$4</c:f>");
  });

  it("writes the same edit into the embedded workbook's cells", () => {
    const data = { categories: ["Jan", "Feb", "Mar"], series: [{ name: "East", values: [1, 2, 3] }, { name: "South", values: [5, 6, 7.5] }] };
    const files = unzipSync(updateWorkbook(XLSX, CHART, data));
    const sheet = strFromU8(files["xl/worksheets/sheet1.xml"]!);
    expect(sheet).toContain('<c r="B1" t="inlineStr"><is><t>East</t></is></c>');
    expect(sheet).toContain('<c r="A3" t="inlineStr"><is><t>Feb</t></is></c>');
    expect(sheet).toContain('<c r="C2" s="3"><v>5</v></c>'); // style kept
    expect(sheet).toContain('<c r="C4"><v>7.5</v></c>');
    expect(sheet).toMatch(/<row r="2">.*<row r="3">.*<row r="4">/); // rows created in order
    expect(strFromU8(files["xl/workbook.xml"]!)).toContain('<calcPr fullCalcOnLoad="1" calcId="191029"/>');
  });

  it("returns undefined for junk", () => {
    expect(readChartData("<nope")).toBeUndefined();
    expect(readChartData("<c:chartSpace xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart'/>")).toBeUndefined();
  });
});
