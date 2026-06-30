// Field decoration: the pass that recomputes every computed field's text after a reflow. It owns
// page-number / page-count fields, caption sequence numbers, cross-reference text (heading / bookmark
// / caption text, page number, or above/below), and the table-of-contents rebuild. Pure DOM over the
// editable body; the only outward effect is asking for one more reflow when a TOC's height changed.
import { t } from "../i18n";

export interface FieldsDeps {
  doc: HTMLElement; // the editable body
  scheduleReflow: () => void; // a TOC rebuild can change page heights; settle pagination once more
  formatPage: (n: number) => string; // render a 1-based page number per the page-number restart/format
}

export function setupFields(deps: FieldsDeps) {
  const { doc, scheduleReflow, formatPage } = deps;

  // Flatten a fragment to one line, inserting a space at each block boundary so a range bookmark that
  // spans paragraphs does not run its text together ("a gammadelta b"); whitespace is then collapsed.
  const BLOCK_TAGS = /^(P|H[1-6]|LI|DIV|TD|TH|BLOCKQUOTE|PRE|TR)$/;
  const fragText = (node: Node): string => {
    let out = "";
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 3) { out += child.textContent ?? ""; continue; }
      if (child.nodeType !== 1) continue;
      const el = child as HTMLElement;
      const block = BLOCK_TAGS.test(el.tagName);
      if (block && out && !out.endsWith(" ")) out += " ";
      out += fragText(el);
      if (block) out += " ";
    }
    return out;
  };
  // The display text for a cross-reference to `target`: a heading's text, a range bookmark's spanned
  // text, or the bookmark name as a fallback for a point bookmark.
  const xrefTargetText = (target: HTMLElement, name: string): string => {
    if (/^H[1-6]$/.test(target.tagName)) return (target.textContent ?? "").trim() || name;
    const id = target.getAttribute("data-rdoc-bm-id");
    const end = id ? Array.from(doc.querySelectorAll<HTMLElement>(".docx-bookmark-end")).find((e) => e.getAttribute("data-rdoc-bm-id") === id) : null;
    if (end && target.compareDocumentPosition(end) & Node.DOCUMENT_POSITION_FOLLOWING) {
      const r = document.createRange();
      r.setStartAfter(target);
      r.setEndBefore(end);
      const txt = fragText(r.cloneContents()).replace(/\s+/g, " ").trim();
      if (txt) return txt;
    }
    return name;
  };
  const tocSig = new WeakMap<Element, string>();
  const decorateFields = (cardCount: number, pageStep: number, vertical: boolean): void => {
    for (const f of Array.from(doc.querySelectorAll<HTMLElement>('.docx-field[data-field="NUMPAGES"]'))) f.textContent = String(cardCount);
    for (const f of Array.from(doc.querySelectorAll<HTMLElement>('.docx-field[data-field="PAGE"]')))
      f.textContent = formatPage(vertical ? 1 : Math.max(1, Math.floor(f.offsetTop / pageStep) + 1));
    // Caption numbers: number each sequence (data-seq) on its own, in document order.
    const seqCounts = new Map<string, number>();
    for (const s of Array.from(doc.querySelectorAll<HTMLElement>('.docx-field[data-field="seq"]'))) {
      const n = (seqCounts.get(s.getAttribute("data-seq") || "") ?? 0) + 1;
      seqCounts.set(s.getAttribute("data-seq") || "", n);
      if (s.textContent !== String(n)) s.textContent = String(n);
    }
    // Cross-references: recompute each xref's text from its target (a heading or bookmark carrying
    // the matching data-rdoc-bm) - its text, or its page number for the "page" format.
    const bmTarget = (name: string) => Array.from(doc.querySelectorAll<HTMLElement>("[data-rdoc-bm]")).find((e) => e.getAttribute("data-rdoc-bm") === name) ?? null;
    for (const x of Array.from(doc.querySelectorAll<HTMLElement>(".docx-xref"))) {
      const name = x.getAttribute("data-rdoc-xref");
      const target = name ? bmTarget(name) : null;
      if (!target) continue;
      const fmt = x.getAttribute("data-rdoc-xref-fmt");
      let text: string;
      if (fmt === "page") text = vertical ? "1" : String(Math.max(1, Math.floor(target.offsetTop / pageStep) + 1));
      else if (fmt === "direction") text = x.compareDocumentPosition(target) & Node.DOCUMENT_POSITION_FOLLOWING ? t("refBelow") : t("refAbove");
      else text = xrefTargetText(target, name!);
      if (text && x.textContent !== text) x.textContent = text;
    }
    let needReflow = false;
    for (const toc of Array.from(doc.querySelectorAll<HTMLElement>(".docx-field-toc"))) {
      const headings = Array.from(doc.querySelectorAll<HTMLElement>("h1,h2,h3")).filter((h) => !h.closest(".docx-field-toc"));
      const pageOf = (el: HTMLElement) => (vertical ? "" : String(Math.max(1, Math.floor(el.offsetTop / pageStep) + 1)));
      const sig = `${cardCount}|` + headings.map((h) => `${h.tagName}:${h.textContent}:${pageOf(h)}`).join("|");
      if (tocSig.get(toc) === sig) continue; // unchanged: don't rebuild (and don't loop reflow)
      tocSig.set(toc, sig);
      needReflow = true;
      toc.replaceChildren();
      if (!headings.length) {
        const e = document.createElement("div");
        e.className = "docx-field-toc-empty";
        e.textContent = t("tocEmpty");
        toc.appendChild(e);
        continue;
      }
      const title = document.createElement("div");
      title.className = "docx-field-toc-title";
      title.textContent = t("tocTitle");
      toc.appendChild(title);
      for (const h of headings) {
        const row = document.createElement("div");
        row.className = `docx-field-toc-row toc-${h.tagName.toLowerCase()}`;
        const txt = document.createElement("span");
        txt.className = "docx-field-toc-text";
        txt.textContent = h.textContent || "";
        const pg = document.createElement("span");
        pg.className = "docx-field-toc-page";
        pg.textContent = pageOf(h);
        row.append(txt, pg);
        toc.appendChild(row);
      }
    }
    if (needReflow) scheduleReflow(); // TOC height changed; one more pass settles pagination
  };

  return { decorateFields };
}
