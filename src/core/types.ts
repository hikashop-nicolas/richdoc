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

/** A named default page size used when a document carries no geometry of its own. */
export type PageSizeName = "a4" | "letter";

/** Physical page size and margins, in CSS pixels at 96 dpi. A view concern only. */
export interface PageGeometry {
  widthPx: number;
  heightPx: number;
  margin: { top: number; right: number; bottom: number; left: number };
}

/** The editable model the engine renders: body + header/footer HTML, comments, fonts. */
export interface RichDoc {
  body: string;
  header: string;
  footer: string;
  headerPath?: string; // adapter-specific key of the header part, for write-back
  footerPath?: string;
  comments: CommentThread[];
  fontCss?: string; // @font-face rules for embedded fonts
  fontUrls?: string[]; // blob URLs to revoke on destroy
  defaultFont?: string;
  page?: PageGeometry; // the document's own page size/margins, if it declares them
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
  trackChanges: boolean;
  images: boolean;
  headerFooter: boolean;
  pageBreak: boolean;
  /** Text + background colour controls (needs inline run colour/shading on save). */
  textColor: boolean;
  /** Font family + size pickers (needs run-level font props on save). */
  fontControls: boolean;
  /** Paragraph alignment buttons. */
  alignment: boolean;
}

/** The irreducible per-format layer: parse, serialize, comment markers, feature flags. */
export interface Adapter {
  original: Uint8Array;
  read(): RichDoc;
  write(bodyHtml: string, parts: { path: string; html: string }[], edits: CommentEdits): Uint8Array;
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
}

export interface RichEditor {
  getBytes(): Promise<Uint8Array>;
  isDirty(): boolean;
  destroy(): void;
}
