// The contract between the shared engine (core/editor.ts) and a per-format adapter.
// The engine knows nothing about OOXML/ODF; the adapter knows nothing about the toolbar.

/** One emoji reaction on a comment, with the people who reacted. */
export interface CommentReaction {
  emoji: string;
  people: string[];
}

/** A single comment (top-level or a reply). */
export interface CommentEntry {
  id: string;
  author: string;
  date: string;
  text: string;
  reactions: CommentReaction[];
  paraId: string; // threading key (last paragraph id of the comment)
}

/** A top-level comment plus its nested replies. */
export interface CommentThread extends CommentEntry {
  replies: CommentEntry[];
  resolved: boolean;
}

/** How a floating image sits relative to the text. Absent on the element = "inline". */
export type ImageWrap = "square" | "tight" | "topbottom" | "behind" | "front";

/** A floating image's wrap mode, alignment, absolute px offset (behind/front), and the
 *  distance (px) kept clear of the text on each side. */
export interface ImageLayout {
  wrap: ImageWrap;
  align: "left" | "center" | "right";
  x: number; // px offset from the anchor (behind/front only)
  y: number; // px offset from the anchor (behind/front only)
  dist?: { t: number; r: number; b: number; l: number }; // wrap padding; omitted = adapter default
}

/** A named default page size used when a document carries no geometry of its own. */
export type PageSizeName = "a4" | "letter";

/** Physical page size and margins, in CSS pixels at 96 dpi. A view concern only. */
export interface PageGeometry {
  widthPx: number;
  heightPx: number;
  margin: { top: number; right: number; bottom: number; left: number };
  /** Top-to-bottom, right-to-left text (Japanese tategaki) when true. */
  vertical?: boolean;
  /** Horizontal right-to-left text (Arabic/Hebrew) when true. Ignored if `vertical`. */
  rtl?: boolean;
  /** Number of text columns on the (single) section; >1 renders multi-column. */
  columns?: number;
  /** Gap between columns, in px. */
  columnGapPx?: number;
  /** Different first-page header/footer (docx w:titlePg, odt first-page master). */
  titlePage?: boolean;
  /** Different odd & even page header/footer (docx w:evenAndOddHeaders, odt *-left). */
  evenOdd?: boolean;
}

/** The editable model the engine renders: body + header/footer HTML, comments, fonts. */
export interface RichDoc {
  body: string;
  header: string;
  footer: string;
  headerPath?: string; // adapter-specific key of the header part, for write-back
  footerPath?: string;
  /** First-page and even-page header/footer variants (the default lives in header/footer above).
      Present only when the document declares them; `path` is the adapter write-back key. */
  headerFirst?: { html: string; path?: string };
  footerFirst?: { html: string; path?: string };
  headerEven?: { html: string; path?: string };
  footerEven?: { html: string; path?: string };
  /** Distinct per-section header/footer parts, keyed by an opaque key that a section's boundary
      paragraph carries (data-rdoc-secheaderkey / data-rdoc-secfooterkey). `path` is the
      adapter-specific write-back key (docx: the real part path; odt: "header@<master>"). */
  sectionBands?: Record<string, { html: string; path: string }>;
  /** Footnote / endnote bodies, keyed by the id their inline reference (data-fn-id) carries. */
  notes?: Note[];
  comments: CommentThread[];
  fontCss?: string; // @font-face rules for embedded fonts
  fontUrls?: string[]; // blob URLs to revoke on destroy
  defaultFont?: string;
  page?: PageGeometry; // the document's own page size/margins, if it declares them
  paragraphStyles?: ParagraphStyle[]; // named paragraph styles, for the style picker
  characterStyles?: ParagraphStyle[]; // named character styles, for the character-style picker
  styleDefs?: StyleDef[]; // each named style's resolved CSS, so its definition can be edited
  styleCss?: string; // CSS rules giving each named style its appearance in the editor
  /** The document's footnote/endnote body style as inheritable CSS declarations (font family,
      size, line-height, colour), applied to the note area so footnotes render at the document's
      footnote size/font instead of a hardcoded default. Empty when the document defines none. */
  noteCss?: string;
}

/** A named style's resolved appearance, so the editor can prefill the edit dialog. */
export interface StyleDef {
  id: string;
  kind: "paragraph" | "character";
  css: Record<string, string>;
}

/** A named paragraph style the user can pick from the style dropdown. */
export interface ParagraphStyle {
  id: string; // the style id written back (docx w:pStyle val / odt style:name)
  name: string; // display name shown in the dropdown
}

/** Structured side-channel of comment edits the engine collects and the adapter applies. */
export interface CommentEdits {
  reactions: { commentId: string; emoji: string; person: string; date: string }[];
  replies: { id: string; paraId: string; parentParaId: string; author: string; date: string; text: string }[];
  done: Map<string, boolean>; // thread paraId -> resolved
  deletedComments: string[]; // comment ids removed from the document
}

/** Metadata for a comment the user adds in the editor. */
export interface NewCommentMeta {
  id: string;
  author: string;
  date: string;
  text: string;
  paraId: string;
}

/** The three markers the adapter builds for a new comment (engine inserts them). */
export interface CommentMarkers {
  start: HTMLElement;
  end: HTMLElement;
  ref: HTMLElement;
}

/** Which engine features a format supports; the toolbar hides what is false. */
export interface Capabilities {
  comments: boolean;
  /** Threaded replies on comments (off for formats with no reply model, e.g. ODF). */
  commentReplies: boolean;
  /** Emoji reactions on comments (off for formats that cannot store them, e.g. ODF). */
  commentReactions: boolean;
  trackChanges: boolean;
  images: boolean;
  /** Inserting new tables (existing tables are always editable when present). */
  tables: boolean;
  headerFooter: boolean;
  pageBreak: boolean;
  /** Text + background colour controls (needs inline run colour/shading on save). */
  textColor: boolean;
  /** Font family + size pickers (needs run-level font props on save). */
  fontControls: boolean;
  /** Paragraph alignment buttons. */
  alignment: boolean;
  /** Vertical (top-to-bottom, right-to-left) writing is supported / rendered. */
  verticalText: boolean;
  /** Insertable fields: page number, page count, table of contents. */
  fields: boolean;
  /** Inserting / editing equations (needs math round-trip; docx OMML for now, odt deferred). */
  equations: boolean;
  /** Section-break authoring, and which paragraph carries a section's geometry:
      "trailing" = the last paragraph of the section (docx w:sectPr / data-rdoc-secbreak),
      "leading"  = the first paragraph of the section (odt master-page / data-rdoc-secstart),
      false      = sections are not authorable. */
  sections: "trailing" | "leading" | false;
}

/** A section's page geometry in px (the on-the-wire shape stashed on a section-boundary
    paragraph as data-rdoc-secbreak / data-rdoc-secstart). */
export interface SecGeom {
  w: number;
  h: number;
  mt: number;
  mr: number;
  mb: number;
  ml: number;
  cols?: number;
  colGap?: number;
  vertical?: boolean; // tategaki (vertical-rl) writing for this section
  rtl?: boolean; // horizontal right-to-left
}

/** A footnote / endnote body. The inline reference in the body HTML carries the matching id. */
export interface Note {
  id: string;
  kind: "footnote" | "endnote";
  html: string;
}

/** A style the user authored in-session, to be added to the document's stylesheet on save.
    `css` uses the same property vocabulary the reader emits (text-align, margin-left "Npx",
    line-height, margin-top/bottom, font-weight, font-style, text-decoration, color "#hex",
    font-size "Npt", font-family). */
export interface NewStyle {
  id: string;
  name: string;
  kind: "paragraph" | "character";
  css: Record<string, string>;
}

/** The irreducible per-format layer: parse, serialize, comment markers, feature flags. */
export interface Adapter {
  original: Uint8Array;
  read(): RichDoc;
  /** Serialize edits. `page` is passed when the user changed the page geometry (margins);
      `newStyles` are styles authored in-session, to add to the stylesheet. */
  write(bodyHtml: string, parts: { path: string; html: string }[], edits: CommentEdits, page?: PageGeometry, newStyles?: NewStyle[], notes?: Note[]): Uint8Array;
  newCommentMarkers(meta: NewCommentMeta): CommentMarkers;
  capabilities: Capabilities;
}

export interface EditorOptions {
  onChange?: () => void;
  /** Author name stamped on comments and tracked changes added in the editor. */
  author?: string;
  /** ISO date string for added comments (injected so the build stays deterministic). */
  now?: string;
  /** Page size used only when the document carries no geometry of its own. Default "a4". */
  defaultPageSize?: PageSizeName;
  /** Paginated view (page cards + reflow) vs a single continuous page. Default true. */
  paginated?: boolean;
  /** Initial zoom factor (1 = 100%); omit for fit-to-width. */
  zoom?: number;
}

export interface RichEditor {
  getBytes(): Promise<Uint8Array>;
  isDirty(): boolean;
  destroy(): void;
}
