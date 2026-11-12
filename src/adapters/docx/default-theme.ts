// The Office theme a package gets when richdoc adds content that takes its colours from the theme
// (charts) and the package has none: Word assumes this theme, other readers fall back to black.
import { strFromU8, strToU8 } from "fflate";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const srgb = (v: string) => `<a:srgbClr val="${v}"/>`;
const sys = (v: string, last: string) => `<a:sysClr val="${v}" lastClr="${last}"/>`;
const phFill = (mods: string) => `<a:solidFill><a:schemeClr val="phClr">${mods}</a:schemeClr></a:solidFill>`;
const line = (w: number) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;

export const DEFAULT_THEME =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="${A}" name="Office Theme"><a:themeElements>` +
  `<a:clrScheme name="Office"><a:dk1>${sys("windowText", "000000")}</a:dk1><a:lt1>${sys("window", "FFFFFF")}</a:lt1>` +
  `<a:dk2>${srgb("44546A")}</a:dk2><a:lt2>${srgb("E7E6E6")}</a:lt2><a:accent1>${srgb("4472C4")}</a:accent1><a:accent2>${srgb("ED7D31")}</a:accent2>` +
  `<a:accent3>${srgb("A5A5A5")}</a:accent3><a:accent4>${srgb("FFC000")}</a:accent4><a:accent5>${srgb("5B9BD5")}</a:accent5><a:accent6>${srgb("70AD47")}</a:accent6>` +
  `<a:hlink>${srgb("0563C1")}</a:hlink><a:folHlink>${srgb("954F72")}</a:folHlink></a:clrScheme>` +
  `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
  `<a:fmtScheme name="Office"><a:fillStyleLst>${phFill("")}${phFill('<a:tint val="50000"/>')}${phFill('<a:shade val="80000"/>')}</a:fillStyleLst>` +
  `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
  `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
  `<a:bgFillStyleLst>${phFill("")}${phFill('<a:tint val="95000"/>')}${phFill('<a:shade val="90000"/>')}</a:bgFillStyleLst></a:fmtScheme>` +
  `</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;

const THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const THEME_CT = "application/vnd.openxmlformats-officedocument.theme+xml";

/** Add word/theme/theme1.xml (with its relationship and content type) when the package has no theme. */
export function ensureTheme(files: Record<string, Uint8Array>): void {
  const relsKey = "word/_rels/document.xml.rels";
  const rels = files[relsKey] ? strFromU8(files[relsKey]!) : "";
  if (/relationships\/theme"/.test(rels) || files["word/theme/theme1.xml"]) return;
  files["word/theme/theme1.xml"] = strToU8(DEFAULT_THEME);
  const doc = new DOMParser().parseFromString(rels || `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`, "application/xml");
  const root = doc.documentElement;
  const used = new Set(Array.from(root.getElementsByTagName("Relationship")).map((r) => r.getAttribute("Id")));
  let n = 1;
  while (used.has(`rId${n}`)) n++;
  const rel = doc.createElementNS(root.namespaceURI, "Relationship");
  rel.setAttribute("Id", `rId${n}`);
  rel.setAttribute("Type", THEME_REL);
  rel.setAttribute("Target", "theme/theme1.xml");
  root.appendChild(rel);
  files[relsKey] = strToU8(new XMLSerializer().serializeToString(doc));
  const ct = files["[Content_Types].xml"];
  if (ct) {
    const c = new DOMParser().parseFromString(strFromU8(ct), "application/xml");
    const ov = c.createElementNS(c.documentElement.namespaceURI, "Override");
    ov.setAttribute("PartName", "/word/theme/theme1.xml");
    ov.setAttribute("ContentType", THEME_CT);
    c.documentElement.appendChild(ov);
    files["[Content_Types].xml"] = strToU8(new XMLSerializer().serializeToString(c));
  }
}
