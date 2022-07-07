const DEFAULT_STYLE = "numberHeadingsAdd3WithLinksOrNot"; // "H1-H3";
const DEFAULT_PREFIX = "Chapter";

const INTERNAL_LINK_MARKER = "<HEADING_LINK>";
const BROKEN_INTERNAL_LINK_MARKER = "<BROKEN_HEADING_LINK>";

const LINK_MARK_STYLE_NEW = new Object();
LINK_MARK_STYLE_NEW[DocumentApp.Attribute.FOREGROUND_COLOR] = "#ff0000";
LINK_MARK_STYLE_NEW[DocumentApp.Attribute.BACKGROUND_COLOR] = "#ffffff";
LINK_MARK_STYLE_NEW[DocumentApp.Attribute.BOLD] = true;