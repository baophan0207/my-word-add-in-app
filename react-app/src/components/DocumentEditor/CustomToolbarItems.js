const undoItem = {
  prefixIcon: "e-de-ctnr-undo",
  tooltipText: "Undo",
  id: "Undo",
};

const redoItem = {
  prefixIcon: "e-de-ctnr-redo",
  tooltipText: "Redo",
  id: "Redo",
};

const boldItem = {
  prefixIcon: "e-de-ctnr-bold",
  tooltipText: "Bold (Ctrl+B)",
  id: "Bold",
};

const italicItem = {
  prefixIcon: "e-de-ctnr-italic",
  tooltipText: "Italic (Ctrl+I)",
  id: "Italic",
};

const underlineItem = {
  prefixIcon: "e-de-ctnr-underline",
  tooltipText: "Underline (Ctrl+U)",
  id: "Underline",
};

const alignLeftItem = {
  prefixIcon: "e-de-ctnr-alignleft",
  tooltipText: "Align Left",
  id: "AlignLeft",
};

const alignCenterItem = {
  prefixIcon: "e-de-ctnr-aligncenter",
  tooltipText: "Align Center",
  id: "AlignCenter",
};

const alignRightItem = {
  prefixIcon: "e-de-ctnr-alignright",
  tooltipText: "Align Right",
  id: "AlignRight",
};

const bulletListItem = {
  prefixIcon: "e-de-ctnr-bullets",
  tooltipText: "Bullet List",
  id: "BulletList",
};

const numberedListItem = {
  prefixIcon: "e-de-ctnr-numbering",
  tooltipText: "Numbered List",
  id: "NumberedList",
};

const decreaseIndentItem = {
  prefixIcon: "e-de-ctnr-decreaseindent",
  tooltipText: "Decrease Indent",
  id: "DecreaseIndent",
};

const increaseIndentItem = {
  prefixIcon: "e-de-ctnr-increaseindent",
  tooltipText: "Increase Indent",
  id: "IncreaseIndent",
};

const highlightColorItem = {
  prefixIcon: "e-de-ctnr-highlight",
  tooltipText: "Highlight Color",
  id: "HighlightColor",
  template:
    '<button id="HighlightColor" class="e-tbar-btn" title="Highlight Color"><span class="e-de-ctnr-highlight e-icons"></span></button>',
};

const fontColorItem = {
  prefixIcon: "e-de-ctnr-fontcolor",
  tooltipText: "Font Color",
  id: "FontColor",
  template:
    '<button id="FontColor" class="e-tbar-btn" title="Font Color"><span class="e-de-ctnr-fontcolor e-icons"></span></button>',
};

const commentItem = {
  prefixIcon: "e-de-cnt-cmt-add",
  tooltipText: "Add Comment",
  id: "AddComment",
};

const trackChangesItem = {
  prefixIcon: "e-de-cnt-track",
  tooltipText: "Track Changes",
  id: "TrackChanges",
};

// Define dropdowns as templates (we'll implement these separately using DOM elements)
const fontFamilyItem = {
  id: "FontFamily",
  tooltipText: "Font Family",
  template: '<div id="fontFamily" class="custom-dropdown"></div>',
};

const fontSizeItem = {
  id: "FontSize",
  tooltipText: "Font Size",
  template: '<div id="fontSize" class="custom-dropdown"></div>',
};

const printItem = {
  prefixIcon: "e-print e-icons",
  tooltipText: "Print",
  id: "Print",
};

const spellCheckItem = {
  prefixIcon: "e-spell-check",
  tooltipText: "Spell Check",
  id: "SpellCheck",
};

const lineHeightItem = {
  prefixIcon: "e-de-ctnr-linespacing",
  tooltipText: "Line Spacing",
  id: "LineHeight",
  template:
    '<button id="LineHeight" class="e-tbar-btn" title="Line Spacing"><span class="e-de-ctnr-linespacing e-icons"></span></button>',
};

export {
  undoItem,
  redoItem,
  boldItem,
  italicItem,
  underlineItem,
  alignLeftItem,
  alignCenterItem,
  alignRightItem,
  bulletListItem,
  numberedListItem,
  decreaseIndentItem,
  increaseIndentItem,
  highlightColorItem,
  fontColorItem,
  commentItem,
  trackChangesItem,
  fontFamilyItem,
  fontSizeItem,
  printItem,
  spellCheckItem,
  lineHeightItem,
};
