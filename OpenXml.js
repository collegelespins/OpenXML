class Ox {
  static AO = "application/vnd.openxmlformats-officedocument.";
  static AP = "application/vnd.openxmlformats-package.";
  static AX = "application/xml";
  static RL =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  static SH = "http://schemas.openxmlformats.org/";

  static AO_EXPR = Ox.AO + "extended-properties+xml";
  static AO_PACK = Ox.AO + "package";

  static AO_SLOU = Ox.AO + "presentationml.slideLayout+xml";
  static AO_PPTX = Ox.AO + "presentationml.presentation.main+xml";
  static AO_SLID = Ox.AO + "presentationml.slide+xml";
  static AO_SLDM = Ox.AO + "presentationml.slideMaster+xml";
  static AO_TSTY = Ox.AO + "presentationml.tableStyles+xml";

  static AO_SHAR = Ox.AO + "spreadsheetml.sharedStrings+xml";
  static AO_BOOK = Ox.AO + "spreadsheetml.sheet.main+xml";
  static AO_SHET = Ox.AO + "spreadsheetml.worksheet+xml";
  static AO_XLSX = Ox.AO + "spreadsheetml.sheet.main+xml";
  static AO_STYX = Ox.AO + "spreadsheetml.styles+xml";

  static AO_THEM = Ox.AO + "theme+xml";
  static AO_DOCX = Ox.AO + "wordprocessingml.document.main+xml";
  static AO_ENDN = Ox.AO + "wordprocessingml.endnotes+xml";
  static AO_FONT = Ox.AO + "wordprocessingml.fontTable+xml";
  static AO_FOOT = Ox.AO + "wordprocessingml.footer+xml";
  static AO_HEAD = Ox.AO + "wordprocessingml.header+xml";
  static AO_STYL = Ox.AO + "wordprocessingml.styles+xml";
  static AO_SETS = Ox.AO + "wordprocessingml.settings+xml";
  static AO_WEBS = Ox.AO + "wordprocessingml.webSettings+xml";

  static RL_ENDN = Ox.RL + "endnotes";
  static RL_EXPR = Ox.RL + "extended-properties";
  static RL_FOOT = Ox.RL + "footer";
  static RL_FONT = Ox.RL + "fontTable";
  static RL_HEAD = Ox.RL + "header";
  static RL_MAIN = Ox.RL + "officeDocument";
  static RL_SLID = Ox.RL + "slide";
  static RL_SLOU = Ox.RL + "slideLayout";
  static RL_STYL = Ox.RL + "styles";
  static RL_SETS = Ox.RL + "settings";
  static RL_SLDM = Ox.RL + "slideMaster";
  static RL_SHAR = Ox.RL + "sharedStrings";
  static RL_THEM = Ox.RL + "theme";

  static RL_WEBS = Ox.RL + "webSettings";
  static RL_SHET = Ox.RL + "worksheet";

  static SH_EXPR = Ox.SH + "officeDocument/2006/extended-properties";
  static SH_DRAW = Ox.SH + "drawingml/2006/main";
  static SH_WDRW = Ox.SH + "drawingml/2006/wordprocessingDrawing";
  static SH_MARK = Ox.SH + "markup-compatibility/2006";
  static SH_VTYP = Ox.SH + "officeDocument/2006/docPropsVTypes";
  static SH_MATH = Ox.SH + "officeDocument/2006/math";
  static SH_OFRL = Ox.SH + "officeDocument/2006/relationships";
  static SH_PACK = Ox.SH + "officeDocument/2006/relationships/package";
  static SH_PPTX = Ox.SH + "presentationml/2006/main";
  static SH_CTYP = Ox.SH + "package/2006/content-types";
  static SH_CORE = Ox.SH + "package/2006/metadata/core-properties";
  static SH_RCOR = Ox.SH +
    "package/2006/relationships/metadata/core-properties";
  static SH_THUM = Ox.SH + "package/2006/relationships/metadata/thumbnail";
  static SH_PKRL = Ox.SH + "package/2006/relationships";
  static SH_SLIB = Ox.SH + "schemaLibrary/2006/main";
  static SH_XLSX = Ox.SH + "spreadsheetml/2006/main";
  static SH_DOCX = Ox.SH + "wordprocessingml/2006/main";

  static AP_CORE = Ox.AP + "core-properties+xml";
  static AP_RELS = Ox.AP + "relationships+xml";

  static DC_TYPE = "http://purl.org/dc/dcmitype/";
  static DC_ELEM = "http://purl.org/dc/elements/1.1/";
  static DC_TERM = "http://purl.org/dc/terms/";
  static SO_WDML = "http://schemas.microsoft.com/office/word/2006/wordml";
  static SO_OFFI = "urn:schemas-microsoft-com:office:office";
  static SO_WORD = "urn:schemas-microsoft-com:office:word";
  static SO_VML = "urn:schemas-microsoft-com:vml";
  static W3_SXML = "http://www.w3.org/2001/XMLSchema-instance";

  static AO_SP = [
    Ox.AO_BOOK,
    Ox.AO_DOCX,
    Ox.AO_ENDN,
    Ox.AO_EXPR,
    Ox.AO_FONT,
    Ox.AO_FOOT,
    Ox.AO_HEAD,
    Ox.AO_PACK,
    Ox.AO_PPTX,
    Ox.AO_SETS,
    Ox.AO_SHAR,
    Ox.AO_SHET,
    Ox.AO_SLDM,
    Ox.AO_SLID,
    Ox.AO_SLOU,
    Ox.AO_STYL,
    Ox.AO_THEM,
    Ox.AO_WEBS,
    Ox.AO_XLSX,
  ];
  static RL_SP = [
    Ox.RL_ENDN,
    Ox.RL_EXPR,
    Ox.RL_FONT,
    Ox.RL_FOOT,
    Ox.RL_HEAD,
    Ox.RL_MAIN,
    Ox.RL_SETS,
    Ox.RL_SHAR,
    Ox.RL_SHET,
    Ox.RL_SLDM,
    Ox.RL_SLID,
    Ox.RL_SLOU,
    Ox.RL_STYL,
    Ox.RL_THEM,
    Ox.RL_WEBS,
  ];
  static SH_SP = [
    Ox.SH_CORE,
    Ox.SH_CTYP,
    Ox.SH_DOCX,
    Ox.SH_DRAW,
    Ox.SH_EXPR,
    Ox.SH_MARK,
    Ox.SH_MATH,
    Ox.SH_OFRL,
    Ox.SH_PACK,
    Ox.SH_PKRL,
    Ox.SH_PPTX,
    Ox.SH_RCOR,
    Ox.SH_SLIB,
    Ox.SH_VTYP,
    Ox.SH_WDRW,
    Ox.SH_XLSX,
  ];
  static ALT_SP = [
    Ox.AP_CORE,
    Ox.AP_RELS,
    Ox.DC_ELEM,
    Ox.DC_TERM,
    Ox.DC_TYPE,
    Ox.SO_OFFI,
    Ox.SO_VML,
    Ox.SO_WDML,
    Ox.SO_WORD,
    Ox.W3_SXML,
  ];
}

class OxPart {
  constructor(pack, name, data) {
    this.name = name;
    this.data = data;
    this.package = pack;
    this.package.parts.push(this);
  }
  get length() {
    return this.data.length;
  }
  get text() {
    return new TextDecoder().decode(this.data);
  }
  get xml() {
    return new DOMParser().parseFromString(this.text, "application/xml");
  }
  get index() {
    return this.package.parts.indexOf(this);
  }
}

class OxPackage {
  static TYPE = Ox.AO_PACK;
  static REL_TYPE = Ox.SH_PACK;
  parts = [];
  constructor(file) {
    this.file = file;
    consoleLog("pink",`Fichier ${this.officeType} "${this.name}" (${this.size} octets) ${this.date.toDateString()}`);
    let reader = new FileReader();
    reader.addEventListener("loadend", () => {
      let parts = UZIP.parse(reader.result);
      for (let part in parts) {
        new OxPart(this, part, parts[part]);
      }
      this.parse();
      this.show();
    });
    reader.readAsArrayBuffer(file);
  }
  text(partName) {
    let part = this.parts.find((p) => p.name == partName);
    if (part == undefined) return "";
    return part.text;
  }
  xml(partName) {
    let part = this.parts.find((p) => p.name == partName);
    return part.xml;
  }
  parse(){

  }
  show() {
    // implémenté chez les enfants (Word, Excel, PowerPoint)
  }
  get officeType() {
    switch (this.file.type.substring(46)) {
      case "presentationml.presentation":
        return "PowerPoint";
      case "wordprocessingml.document":
        return "Word";
      case "spreadsheetml.sheet":
        return "Excel";
      default:
        return "inconnu";
    }
  }
  get name() {
    return this.file.name;
  }
  get type() {
    return this.file.type;
  }
  get size() {
    return this.file.size;
  }
  get date() {
    return new Date(this.file.lastModified);
  }
}

class OxWord extends OxPackage {
  constructor(file) {
    super(file);
  }
  show() {
    consoleLog("#6666FF", "Il y a", this.parts.length, "parties dans ce document", this.officeType);
    consoleLog("#CCCCFF", this.text("word/document.xml"));
  }
}

class OxExcel extends OxPackage {
  constructor(file) {
    super(file);
    // http://poi.apache.org/components/spreadsheet/quick-guide.html
    // https://www.codeproject.com/Articles/371203/Creating-basic-Excel-workbook-with-Open-XML (format, langue, monnaie, date)

  }
  show() {
    consoleLog("#66FF66", "Il y a", this.parts.length,"parties dans ce document",this.officeType);
    consoleLog("#CCFFCC", this.text("xl/workbook.xml"));
    consoleLog("#AAFFAA", this.text("xl/sharedStrings.xml"));
    consoleLog("#CCFFCC", this.text("xl/worksheets/sheet1.xml"));
  }
}

class OxPowerPoint extends OxPackage {
  constructor(file) {
    super(file);
    // http://poi.apache.org/components/slideshow/xslf-cookbook.html
  }
  show() {
    consoleLog("orange", "Il y a",this.parts.length, "parties dans ce document",this.officeType);
    consoleLog("white", this.text("ppt/presentation.xml"));
  }
}

