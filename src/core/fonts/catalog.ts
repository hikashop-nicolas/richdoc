// Font name lists: the always-offered base fonts, the installed-font candidates probed by
// detection, and the curated Google Fonts catalog for "Get more fonts".

/** Offered by default when the browser can draw them (filtered by detection). */
export const BASE_FONTS = [
  "Arial", "Calibri", "Cambria", "Century", "Courier New", "Georgia", "Helvetica", "Tahoma",
  "Times New Roman", "Trebuchet MS", "Verdana", "Segoe UI", "Roboto", "Noto Sans", "Noto Serif",
  "Liberation Sans", "Liberation Serif", "Liberation Mono", "DejaVu Sans", "DejaVu Serif",
];

/** Common installed families across desktop and mobile systems, Office, Adobe and LibreOffice. */
export const LOCAL_CANDIDATES: string[] = [
  // macOS
  "Helvetica", "Helvetica Neue", "Arial", "Arial Black", "Arial Narrow", "Arial Rounded MT Bold",
  "Avenir", "Avenir Next", "Avenir Next Condensed", "Baskerville", "Big Caslon", "Bodoni 72",
  "Bradley Hand", "Brush Script MT", "Chalkboard", "Chalkboard SE", "Chalkduster", "Charter",
  "Cochin", "Comic Sans MS", "Copperplate", "Courier", "Courier New", "Didot", "Futura", "Geneva",
  "Georgia", "Gill Sans", "Herculanum", "Hoefler Text", "Impact", "Lucida Grande", "Luminari",
  "Marker Felt", "Menlo", "Monaco", "Noteworthy", "Optima", "Palatino", "Papyrus", "Phosphate",
  "Rockwell", "Savoye LET", "SF Pro", "SF Pro Display", "SF Pro Text", "SF Mono", "Skia",
  "Snell Roundhand", "Tahoma", "Times", "Times New Roman", "Trattatello", "Trebuchet MS",
  "Verdana", "Zapfino", "American Typewriter", "Andale Mono", "Apple Chancery", "Athelas",
  "Seravek", "Superclarendon", "Iowan Old Style", "Kefa", "Mishafi", "Damascus", "Geeza Pro",
  "Baghdad", "Al Nile", "Al Bayan", "Farah", "Diwan Kufi", "Arial Hebrew", "Corsiva Hebrew",
  "Raanana", "New Peninsula", "Kohinoor Devanagari", "ITF Devanagari", "Devanagari Sangam MN",
  "Thonburi", "Krungthep", "Silom", "Sathu", "Ayuthaya",
  // macOS CJK
  "Apple SD Gothic Neo", "AppleGothic", "AppleMyungjo", "Hiragino Sans", "Hiragino Kaku Gothic ProN",
  "Hiragino Kaku Gothic Pro", "Hiragino Mincho ProN", "Hiragino Mincho Pro", "Hiragino Maru Gothic ProN",
  "Hiragino Sans GB", "Osaka", "Toppan Bunkyu Mincho", "Toppan Bunkyu Gothic",
  "Tsukushi A Round Gothic", "Tsukushi B Round Gothic", "Klee", "YuGothic", "YuMincho", "Kyokasho",
  "PingFang SC", "PingFang TC", "PingFang HK", "Heiti SC", "Heiti TC", "STHeiti", "Songti SC",
  "Songti TC", "STSong", "Kaiti SC", "Kaiti TC", "STKaiti", "Lantinghei SC", "Hannotate SC",
  "HanziPen SC", "Libian SC", "Weibei SC", "Wawati SC", "Xingkai SC", "Yuanti SC", "BiauKai",
  "LiSong Pro", "LiHei Pro", "Apple LiGothic", "Apple LiSung", "Gungseo", "PCMyungjo", "HeadLineA",
  // Windows and Office
  "Segoe UI", "Segoe Print", "Segoe Script", "Segoe UI Emoji", "Segoe UI Symbol", "Calibri",
  "Calibri Light", "Cambria", "Cambria Math", "Candara", "Consolas", "Constantia", "Corbel",
  "Franklin Gothic Medium", "Gabriola", "Garamond", "Lucida Console", "Lucida Sans Unicode",
  "Lucida Handwriting", "Microsoft Sans Serif", "Palatino Linotype", "Book Antiqua",
  "Bookman Old Style", "Century", "Century Gothic", "Century Schoolbook", "Sylfaen", "Symbol",
  "Webdings", "Wingdings", "Ink Free", "Bahnschrift", "MV Boli", "Ebrima", "Gadugi", "Nirmala UI",
  "Leelawadee UI", "Leelawadee", "Mangal", "Aparajita", "Kokila", "Utsaah", "Javanese Text",
  "Myanmar Text", "Mongolian Baiti", "Microsoft Himalaya", "Microsoft Yi Baiti", "Microsoft Tai Le",
  "Microsoft New Tai Lue", "Microsoft PhagsPa", "Traditional Arabic", "Simplified Arabic",
  "Arabic Typesetting", "Sakkal Majalla", "Andalus", "Aldhabi", "Urdu Typesetting", "David",
  "Miriam", "FrankRuehl", "Narkisim", "Gisha", "Levenim MT", "Aharoni", "Tunga", "Latha", "Vrinda",
  "Shruti", "Raavi", "Gautami", "Kartika", "Kalinga", "Iskoola Pota", "DaunPenh", "Khmer UI",
  "Lao UI", "DokChampa", "Plantagenet Cherokee", "Estrangelo Edessa", "Tw Cen MT", "Gill Sans MT",
  "Agency FB", "Algerian", "Bauhaus 93", "Bell MT", "Berlin Sans FB", "Bernard MT Condensed",
  "Bodoni MT", "Britannic Bold", "Broadway", "Calisto MT", "Castellar", "Centaur", "Chiller",
  "Colonna MT", "Cooper Black", "Copperplate Gothic Bold", "Curlz MT", "Edwardian Script ITC",
  "Elephant", "Engravers MT", "Felix Titling", "Footlight MT Light", "Forte", "Freestyle Script",
  "French Script MT", "Gigi", "Goudy Old Style", "Haettenschweiler", "Harlow Solid Italic",
  "Harrington", "High Tower Text", "Imprint MT Shadow", "Informal Roman", "Jokerman", "Juice ITC",
  "Kristen ITC", "Kunstler Script", "Lucida Bright", "Lucida Calligraphy", "Lucida Fax", "Magneto",
  "Maiandra GD", "Matura MT Script Capitals", "Mistral", "Modern No. 20", "Monotype Corsiva",
  "Niagara Engraved", "OCR A Extended", "Old English Text MT", "Onyx", "Palace Script MT",
  "Parchment", "Perpetua", "Playbill", "Poor Richard", "Pristina", "Rage Italic", "Ravie",
  "Script MT Bold", "Showcard Gothic", "Snap ITC", "Stencil", "Tempus Sans ITC", "Viner Hand ITC",
  "Vivaldi", "Vladimir Script", "Wide Latin", "Aptos", "Aptos Display", "Aptos Narrow",
  "Aptos Serif", "Aptos Mono", "Grandview", "Seaford", "Skeena", "Tenorite", "Bierstadt",
  "Sitka Text", "Cascadia Code", "Cascadia Mono",
  // Windows CJK
  "MS Gothic", "MS PGothic", "MS UI Gothic", "MS Mincho", "MS PMincho", "Meiryo", "Meiryo UI",
  "Yu Gothic", "Yu Gothic UI", "Yu Mincho", "BIZ UDGothic", "BIZ UDPGothic", "BIZ UDMincho",
  "BIZ UDPMincho", "UD Digi Kyokasho N-R", "HGGothicM", "HGMinchoE", "HGMaruGothicMPRO",
  "HGSoeiKakugothicUB", "Microsoft YaHei", "Microsoft YaHei UI", "Microsoft JhengHei",
  "Microsoft JhengHei UI", "SimSun", "NSimSun", "SimHei", "KaiTi", "FangSong", "DengXian",
  "MingLiU", "PMingLiU", "MingLiU-ExtB", "DFKai-SB", "Malgun Gothic", "Gulim", "GulimChe", "Dotum",
  "DotumChe", "Batang", "BatangChe", "Gungsuh",
  // Linux, LibreOffice bundled, ChromeOS, Android
  "Liberation Sans", "Liberation Serif", "Liberation Mono", "Liberation Sans Narrow", "DejaVu Sans",
  "DejaVu Serif", "DejaVu Sans Mono", "DejaVu Sans Condensed", "Noto Sans", "Noto Serif",
  "Noto Sans Mono", "Noto Mono", "Noto Sans Arabic", "Noto Naskh Arabic", "Noto Kufi Arabic",
  "Noto Sans Hebrew", "Noto Serif Hebrew", "Noto Sans Devanagari", "Noto Sans Thai",
  "Noto Sans Armenian", "Noto Sans Georgian", "Noto Sans Symbols", "Noto Color Emoji",
  "Noto Sans CJK JP", "Noto Sans CJK SC", "Noto Sans CJK TC", "Noto Sans CJK KR", "Noto Sans CJK HK",
  "Noto Serif CJK JP", "Noto Serif CJK SC", "Noto Serif CJK TC", "Noto Serif CJK KR", "Noto Sans JP",
  "Noto Serif JP", "Noto Sans SC", "Noto Sans TC", "Noto Sans KR", "Source Han Sans",
  "Source Han Serif", "Source Han Sans JP", "Source Han Serif JP", "Source Han Sans SC",
  "Source Han Sans K", "Source Sans Pro", "Source Serif Pro", "Source Code Pro", "Carlito", "Caladea",
  "Gelasio", "Alef", "Amiri", "Amiri Quran", "David CLM", "David Libre", "Frank Ruehl CLM",
  "Miriam CLM", "Miriam Mono CLM", "Rubik", "Reem Kufi", "Scheherazade", "Scheherazade New", "Karla",
  "Linux Libertine G", "Linux Biolinum G", "Linux Libertine O", "Linux Biolinum O",
  "Libertinus Serif", "Libertinus Sans", "Libertinus Mono", "EmojiOne Color", "OpenSymbol",
  "Gentium Basic", "Gentium Book Basic", "Gentium Plus", "KacstBook", "KacstOffice",
  "Lohit Devanagari", "Lohit Tamil", "Lohit Bengali", "Ubuntu", "Ubuntu Mono", "Ubuntu Condensed",
  "Cantarell", "Droid Sans", "Droid Serif", "Droid Sans Mono", "Droid Sans Fallback",
  "Bitstream Vera Sans", "Bitstream Vera Serif", "Bitstream Vera Sans Mono", "FreeSans", "FreeSerif",
  "FreeMono", "Nimbus Sans", "Nimbus Roman", "Nimbus Mono PS", "URW Bookman", "URW Gothic", "C059",
  "P052", "Z003", "TeX Gyre Heros", "TeX Gyre Termes", "TeX Gyre Pagella", "TeX Gyre Bonum",
  "TeX Gyre Schola", "TeX Gyre Cursor", "Latin Modern Roman", "Inconsolata", "Fira Sans", "Fira Mono",
  "Fira Code", "Hack", "Roboto", "Roboto Mono", "Roboto Slab", "Roboto Condensed", "Open Sans",
  "Lato", "Montserrat", "Oswald", "Raleway", "Poppins", "Inter", "IBM Plex Sans", "IBM Plex Serif",
  "IBM Plex Mono", "JetBrains Mono", "Arimo", "Tinos", "Cousine", "Google Sans", "TakaoGothic",
  "TakaoPGothic", "TakaoMincho", "IPAGothic", "IPAPGothic", "IPAMincho", "IPAPMincho", "IPAexGothic",
  "IPAexMincho", "VL Gothic", "VL PGothic", "WenQuanYi Micro Hei", "WenQuanYi Zen Hei",
  "AR PL UMing CN", "AR PL UKai CN", "UnBatang", "UnDotum", "Baekmuk Gulim", "Sarabun", "Waree",
  "Garuda", "Loma", "Kinnari", "Norasi",
  // Adobe
  "Adobe Garamond Pro", "Minion Pro", "Myriad Pro", "Adobe Caslon Pro", "Adobe Jenson Pro",
  "Trajan Pro", "Kozuka Gothic Pr6N", "Kozuka Mincho Pr6N", "Kozuka Gothic Pro", "Kozuka Mincho Pro",
  "Adobe Ming Std", "Adobe Song Std", "Adobe Heiti Std", "Adobe Kaiti Std", "Adobe Fangsong Std",
  "Adobe Myungjo Std", "Adobe Gothic Std", "Adobe Arabic", "Adobe Hebrew", "Adobe Devanagari",
  "Birch Std", "Blackoak Std", "Brush Script Std", "Chaparral Pro", "Charlemagne Std", "Cooper Std",
  "Giddyup Std", "Hobo Std", "Letter Gothic Std", "Lithos Pro", "Mesquite Std", "Nueva Std",
  "OCR A Std", "Orator Std", "Poplar Std", "Prestige Elite Std", "Rosewood Std", "Stencil Std",
  "Tekton Pro",
];

/** Catalog groups, each with an i18n label key. */
export type WebFontGroup =
  | "compat" | "sans" | "serif" | "mono" | "display" | "hand"
  | "ja" | "zh" | "ko" | "ar" | "he" | "cyr" | "deva" | "th";

/** v: the styles the family has: b = bold 700, i = italic. */
export interface WebFont { family: string; group: WebFontGroup; v: "" | "b" | "i" | "bi" }

// Compact source: one string per group, "Family:v" separated by "|" ("Family" alone = regular only).
const SRC: Record<WebFontGroup, string> = {
  compat: "Arimo:bi|Tinos:bi|Cousine:bi|Carlito:bi|Caladea:bi|Gelasio:bi",
  sans: "Roboto:bi|Open Sans:bi|Lato:bi|Montserrat:bi|Inter:bi|Noto Sans:bi|Source Sans 3:bi|Nunito:bi|Poppins:bi|Raleway:bi|Work Sans:bi|Ubuntu:bi|Fira Sans:bi|Mulish:bi|Karla:bi|Barlow:bi|DM Sans:bi|Manrope:b|Quicksand:b|Josefin Sans:bi|Oswald:b|Cabin:bi|Libre Franklin:bi|IBM Plex Sans:bi|Figtree:bi|Lexend:b|Atkinson Hyperlegible:bi|Archivo:bi|Outfit:b|Titillium Web:bi|Exo 2:bi|Public Sans:bi|Red Hat Display:bi|Jost:bi|Comfortaa:b|Varela Round",
  serif: "Merriweather:bi|Lora:bi|Playfair Display:bi|Noto Serif:bi|Libre Baskerville:bi|EB Garamond:bi|Crimson Pro:bi|Source Serif 4:bi|Cormorant Garamond:bi|Bitter:bi|Domine:b|Spectral:bi|Cardo:bi|Alegreya:bi|Libre Caslon Text:bi|Old Standard TT:bi|Zilla Slab:bi|Roboto Slab:b|Arvo:bi|Vollkorn:bi|IBM Plex Serif:bi|Literata:bi|DM Serif Display:i",
  mono: "Roboto Mono:bi|Source Code Pro:bi|Fira Code:b|JetBrains Mono:bi|IBM Plex Mono:bi|Inconsolata:b|Space Mono:bi|Ubuntu Mono:bi|Courier Prime:bi|Noto Sans Mono:b",
  display: "Bebas Neue|Anton|Abril Fatface|Lobster|Righteous|Alfa Slab One|Bangers|Fredoka:b|Archivo Black|Cinzel:b|Comic Neue:bi",
  hand: "Dancing Script:b|Pacifico|Caveat:b|Great Vibes|Satisfy|Indie Flower|Shadows Into Light|Amatic SC:b|Kalam:b|Patrick Hand|Sacramento|Permanent Marker|Gloria Hallelujah|Architects Daughter|Homemade Apple",
  ja: "Noto Sans JP:b|Noto Serif JP:b|M PLUS 1p:b|M PLUS Rounded 1c:b|Zen Maru Gothic:b|Zen Kaku Gothic New:b|Zen Old Mincho:b|Klee One|Yuji Syuku|Yuji Mai|Kaisei Decol:b|Kaisei Opti:b|Shippori Mincho:b|Sawarabi Mincho|Sawarabi Gothic|Kosugi Maru|BIZ UDPGothic:b|BIZ UDPMincho:b|Dela Gothic One|Hina Mincho|Kiwi Maru|Yomogi|Murecho:b",
  zh: "Noto Sans SC:b|Noto Serif SC:b|Noto Sans TC:b|Noto Serif TC:b|ZCOOL XiaoWei|Ma Shan Zheng|LXGW WenKai TC:b",
  ko: "Noto Sans KR:b|Noto Serif KR:b|Nanum Myeongjo:b|Nanum Gothic:b|Nanum Pen Script|Gowun Dodum|Black Han Sans",
  ar: "Noto Naskh Arabic:b|Noto Kufi Arabic:b|Amiri:bi|Cairo:b|Tajawal:b|Almarai:b|Scheherazade New:b",
  he: "Noto Sans Hebrew:b|Frank Ruhl Libre:b|Heebo:b|Assistant:b|Rubik:bi|David Libre:b|Alef:b",
  cyr: "PT Sans:bi|PT Serif:bi|PT Mono|Yeseva One|Russo One",
  deva: "Noto Sans Devanagari:b|Hind:b|Mukta:b|Tiro Devanagari Hindi:i",
  th: "Sarabun:bi|Prompt:bi|Noto Sans Thai:b|Kanit:bi",
};

export const WEB_FONTS: WebFont[] = (Object.keys(SRC) as WebFontGroup[]).flatMap((group) =>
  SRC[group].split("|").map((s) => {
    const [family, v = ""] = s.split(":");
    return { family: family!, group, v: v as WebFont["v"] };
  }),
);

/** Preview text for a group whose script a family name in Latin letters would not show. */
export const GROUP_SAMPLE: Partial<Record<WebFontGroup, string>> = {
  ja: "あア漢字", zh: "中文字体", ko: "한국어", ar: "العربية", he: "עברית", cyr: "Кириллица", deva: "हिन्दी", th: "ภาษาไทย",
};
