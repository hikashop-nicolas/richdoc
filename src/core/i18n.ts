// Self-contained i18n for docxedit so the library is a complete multilingual product on
// its own. Detects the locale from the browser / device preferred-languages list (base
// language, first match), English fallback. Adding a language = add a dict to LOCALES;
// hosts may force one via setLocale().

type Dict = Record<string, string>;

const en: Dict = {
  toolbar: "Document formatting tools",
  documentText: "Document text",
  paragraphStyle: "Paragraph style",
  styleParagraph: "Paragraph",
  styleH1: "Heading 1",
  styleH2: "Heading 2",
  styleH3: "Heading 3",
  bold: "Bold",
  italic: "Italic",
  underline: "Underline",
  bulletedLabel: "• List",
  bulleted: "Bulleted list",
  numberedLabel: "1. List",
  numbered: "Numbered list",
  link: "Link",
  linkAria: "Add or edit link",
  linkPrompt: "Link URL (empty to remove):",
  alignLeft: "Align left",
  alignCenter: "Center",
  alignRight: "Align right",
  alignJustify: "Justify",
  textColor: "Text colour",
  highlight: "Highlight",
  none: "None",
  hlYellow: "Yellow",
  hlGreen: "Green",
  hlCyan: "Cyan",
  hlMagenta: "Magenta",
  hlBlue: "Blue",
  hlRed: "Red",
  font: "Font",
  size: "Size",
  header: "Header",
  footer: "Footer",
  editHeaderFooter: "Click to edit",
  pageBreak: "Page break",
  insertPageBreak: "Insert page break",
  insertImage: "Insert image",
  addComment: "Add comment",
  commentPrompt: "Comment:",
  commentSelect: "Select some text first.",
  more: "more",
  moreTools: "More tools",
  zoomIn: "Zoom in",
  zoomOut: "Zoom out",
  zoomReset: "Reset zoom (fit width)",
  addReaction: "Add reaction",
  reply: "Reply",
  send: "Send",
  resolve: "Resolve",
  deleteComment: "Delete",
  deleteImage: "Delete image",
  suggesting: "Suggesting mode (track changes)",
  accept: "Accept",
  reject: "Reject",
  acceptAll: "Accept all changes",
  rejectAll: "Reject all changes",
};

const fr: Dict = {
  toolbar: "Outils de mise en forme du document",
  documentText: "Texte du document",
  paragraphStyle: "Style de paragraphe",
  styleParagraph: "Paragraphe",
  styleH1: "Titre 1",
  styleH2: "Titre 2",
  styleH3: "Titre 3",
  bold: "Gras",
  italic: "Italique",
  underline: "Souligné",
  bulletedLabel: "• Liste",
  bulleted: "Liste à puces",
  numberedLabel: "1. Liste",
  numbered: "Liste numérotée",
  link: "Lien",
  linkAria: "Ajouter ou modifier un lien",
  linkPrompt: "URL du lien (vide pour retirer) :",
  alignLeft: "Aligner à gauche",
  alignCenter: "Centrer",
  alignRight: "Aligner à droite",
  alignJustify: "Justifier",
  textColor: "Couleur du texte",
  highlight: "Surlignage",
  none: "Aucun",
  hlYellow: "Jaune",
  hlGreen: "Vert",
  hlCyan: "Cyan",
  hlMagenta: "Magenta",
  hlBlue: "Bleu",
  hlRed: "Rouge",
  font: "Police",
  size: "Taille",
  header: "En-tête",
  footer: "Pied de page",
  editHeaderFooter: "Cliquer pour modifier",
  pageBreak: "Saut de page",
  insertPageBreak: "Insérer un saut de page",
  insertImage: "Insérer une image",
  addComment: "Ajouter un commentaire",
  commentPrompt: "Commentaire :",
  commentSelect: "Sélectionnez d'abord du texte.",
  more: "plus",
  moreTools: "Plus d'outils",
  zoomIn: "Zoom avant",
  zoomOut: "Zoom arrière",
  zoomReset: "Réinitialiser le zoom (ajuster la largeur)",
  addReaction: "Ajouter une réaction",
  reply: "Répondre",
  send: "Envoyer",
  resolve: "Résoudre",
  deleteComment: "Supprimer",
  deleteImage: "Supprimer l'image",
  suggesting: "Mode suggestion (suivi des modifications)",
  accept: "Accepter",
  reject: "Refuser",
  acceptAll: "Accepter toutes les modifications",
  rejectAll: "Refuser toutes les modifications",
};

const LOCALES: Record<string, Dict> = { en, fr };

let active: Dict | null = null;

function detect(): Dict {
  const prefs = (typeof navigator !== "undefined" && navigator.languages) || ["en"];
  for (const tag of prefs) {
    const base = tag.toLowerCase().split("-")[0]!;
    if (LOCALES[base]) return LOCALES[base]!;
  }
  return en;
}

/** Force a locale (host escape hatch). Unknown codes fall back to English. */
export function setLocale(code: string): void {
  active = LOCALES[code.toLowerCase().split("-")[0]!] ?? en;
}

export function t(key: string, params?: Record<string, string | number>): string {
  if (!active) active = detect();
  let s = active[key] ?? en[key] ?? key;
  if (params) s = s.replace(/\{(\w+)\}/g, (_, k: string) => (k in params ? String(params[k]) : `{${k}}`));
  return s;
}
