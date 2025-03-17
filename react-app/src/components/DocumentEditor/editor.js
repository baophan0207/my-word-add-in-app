import * as React from "react";
import {
  DocumentEditorContainerComponent,
  DocumentEditorComponent,
  Selection,
  Editor,
  EditorHistory,
  ContextMenu,
  SfdtExport,
} from "@syncfusion/ej2-react-documenteditor";
import {
  ToolbarComponent,
  ItemDirective,
  ItemsDirective,
} from "@syncfusion/ej2-react-navigations";
import { ComboBoxComponent } from "@syncfusion/ej2-react-dropdowns";
import { ColorPickerComponent } from "@syncfusion/ej2-react-inputs";
import { DropDownButtonComponent } from "@syncfusion/ej2-react-splitbuttons";

// Inject required modules into DocumentEditorComponent
DocumentEditorComponent.Inject(
  Selection,
  Editor,
  EditorHistory,
  ContextMenu,
  SfdtExport
);

// Define font styles and sizes outside the class as they are constant
const fontStyle = [
  "Algerian",
  "Arial",
  "Calibri",
  "Cambria",
  "Cambria Math",
  "Candara",
  "Courier New",
  "Georgia",
  "Impact",
  "Segoe Print",
  "Segoe Script",
  "Segoe UI",
  "Symbol",
  "Times New Roman",
  "Verdana",
  "Windings",
];
const fontSize = [
  "8",
  "9",
  "10",
  "11",
  "12",
  "14",
  "16",
  "18",
  "20",
  "22",
  "24",
  "26",
  "28",
  "36",
  "48",
  "72",
  "96",
];

// Line spacing items
const lineSpacingItems = [
  { text: "Single" },
  { text: "1.15" },
  { text: "1.5" },
  { text: "Double" },
];

// Define the App class component
class EditorComponent extends React.Component {
  constructor(props) {
    super(props);
    this.documenteditor = null; // Will be set via ref
  }

  // Lifecycle method to replace useEffect for initialization
  componentDidMount() {
    if (this.documenteditor) {
      // Enable editor history for undo/redo functionality
      this.documenteditor.enableEditorHistory = true;

      this.documenteditor.selectionChange = () => {
        setTimeout(() => {
          this.onSelectionChange();
        }, 20);
      };
    }
  }

  // Handle selection changes in the document editor
  onSelectionChange() {
    if (this.documenteditor && this.documenteditor.selection) {
      // Update character formatting options (bold, italic, etc.)
      this.enableDisableFontOptions();

      // Update paragraph formatting options (alignment, etc.)
      this.updateParagraphFormatState();

      // Update undo/redo states
      this.updateUndoRedoState();
    }
  }

  // Enable or disable undo/redo buttons based on history state
  updateUndoRedoState() {
    // Check undo state
    const undoBtn = document.getElementById("undo");
    if (undoBtn) {
      undoBtn.disabled = !this.documenteditor.editorHistory.canUndo;
      undoBtn.classList.toggle(
        "e-disabled",
        !this.documenteditor.editorHistory.canUndo
      );
    }

    // Check redo state
    const redoBtn = document.getElementById("redo");
    if (redoBtn) {
      redoBtn.disabled = !this.documenteditor.editorHistory.canRedo;
      redoBtn.classList.toggle(
        "e-disabled",
        !this.documenteditor.editorHistory.canRedo
      );
    }
  }

  // Enable or disable toolbar buttons based on character formatting
  enableDisableFontOptions() {
    const characterFormat = this.documenteditor.selection.characterFormat;
    const properties = [
      characterFormat.bold,
      characterFormat.italic,
      characterFormat.underline,
      characterFormat.strikethrough,
    ];
    const toggleBtnId = ["bold", "italic", "underline", "strikethrough"];
    for (let i = 0; i < properties.length; i++) {
      this.changeActiveState(properties[i], toggleBtnId[i]);
    }
  }

  // Update paragraph formatting state in toolbar
  updateParagraphFormatState() {
    const paragraphFormat = this.documenteditor.selection.paragraphFormat;
    const toggleBtnId = [
      "AlignLeft",
      "AlignCenter",
      "AlignRight",
      "Justify",
      "ShowParagraphMark",
    ];

    // Remove toggle state from all alignment buttons
    for (let i = 0; i < toggleBtnId.length; i++) {
      const toggleBtn = document.getElementById(toggleBtnId[i]);
      if (toggleBtn) {
        toggleBtn.classList.remove("e-btn-toggle");
      }
    }

    // Add toggle state based on selection paragraph format
    if (paragraphFormat.textAlignment === "Left") {
      const alignLeftBtn = document.getElementById("AlignLeft");
      if (alignLeftBtn) alignLeftBtn.classList.add("e-btn-toggle");
    } else if (paragraphFormat.textAlignment === "Right") {
      const alignRightBtn = document.getElementById("AlignRight");
      if (alignRightBtn) alignRightBtn.classList.add("e-btn-toggle");
    } else if (paragraphFormat.textAlignment === "Center") {
      const alignCenterBtn = document.getElementById("AlignCenter");
      if (alignCenterBtn) alignCenterBtn.classList.add("e-btn-toggle");
    } else {
      const justifyBtn = document.getElementById("Justify");
      if (justifyBtn) justifyBtn.classList.add("e-btn-toggle");
    }

    // Update paragraph mark button state
    if (this.documenteditor.documentEditorSettings.showHiddenMarks) {
      const paraMarkBtn = document.getElementById("ShowParagraphMark");
      if (paraMarkBtn) paraMarkBtn.classList.add("e-btn-toggle");
    }
  }

  // Update the active state of toolbar buttons
  changeActiveState(property, btnId) {
    const toggleBtn = document.getElementById(btnId);
    if (!toggleBtn) return;

    if (
      (typeof property === "boolean" && property) ||
      (typeof property === "string" && property !== "None")
    ) {
      toggleBtn.classList.add("e-btn-toggle");
    } else if (toggleBtn.classList.contains("e-btn-toggle")) {
      toggleBtn.classList.remove("e-btn-toggle");
    }
  }

  // Handle toolbar button clicks
  toolbarButtonClick(arg) {
    if (!this.documenteditor) return;

    switch (arg.item.id) {
      // Undo and Redo operations
      case "undo":
        this.documenteditor.editorHistory.undo();
        this.updateUndoRedoState();
        break;
      case "redo":
        this.documenteditor.editorHistory.redo();
        this.updateUndoRedoState();
        break;

      // Character formatting
      case "bold":
        this.documenteditor.editor.toggleBold();
        break;
      case "italic":
        this.documenteditor.editor.toggleItalic();
        break;
      case "underline":
        this.documenteditor.editor.toggleUnderline("Single");
        break;
      case "strikethrough":
        this.documenteditor.editor.toggleStrikethrough();
        break;
      case "subscript":
        this.documenteditor.editor.toggleSubscript();
        break;
      case "superscript":
        this.documenteditor.editor.toggleSuperscript();
        break;

      // Paragraph formatting
      case "AlignLeft":
        this.documenteditor.editor.toggleTextAlignment("Left");
        break;
      case "AlignRight":
        this.documenteditor.editor.toggleTextAlignment("Right");
        break;
      case "AlignCenter":
        this.documenteditor.editor.toggleTextAlignment("Center");
        break;
      case "Justify":
        this.documenteditor.editor.toggleTextAlignment("Justify");
        break;
      case "IncreaseIndent":
        this.documenteditor.editor.increaseIndent();
        break;
      case "DecreaseIndent":
        this.documenteditor.editor.decreaseIndent();
        break;
      case "ClearFormat":
        this.documenteditor.editor.clearFormatting();
        break;
      case "ShowParagraphMark":
        this.documenteditor.documentEditorSettings.showHiddenMarks =
          !this.documenteditor.documentEditorSettings.showHiddenMarks;
        this.updateParagraphFormatState();
        break;
      default:
        break;
    }
  }

  // Change font family of selected text
  changeFontFamily(args) {
    if (!this.documenteditor) return;
    this.documenteditor.selection.characterFormat.fontFamily = args.value;
    this.documenteditor.focusIn();
  }

  // Change font size of selected text
  changeFontSize(args) {
    if (!this.documenteditor) return;
    this.documenteditor.selection.characterFormat.fontSize = args.value;
    this.documenteditor.focusIn();
  }

  // Change font color of selected text
  changeFontColor(args) {
    if (!this.documenteditor) return;
    this.documenteditor.selection.characterFormat.fontColor =
      args.currentValue.hex;
    this.documenteditor.focusIn();
  }

  // Apply line spacing
  lineSpacingAction(args) {
    if (!this.documenteditor) return;

    const text = args.item.text;
    switch (text) {
      case "Single":
        this.documenteditor.selection.paragraphFormat.lineSpacing = 1;
        break;
      case "1.15":
        this.documenteditor.selection.paragraphFormat.lineSpacing = 1.15;
        break;
      case "1.5":
        this.documenteditor.selection.paragraphFormat.lineSpacing = 1.5;
        break;
      case "Double":
        this.documenteditor.selection.paragraphFormat.lineSpacing = 2;
        break;
    }

    setTimeout(() => {
      this.documenteditor.focusIn();
    }, 30);
  }

  // Template for font color picker in toolbar
  contentTemplate1() {
    return (
      <ColorPickerComponent
        showButtons={true}
        value="#000000"
        change={this.changeFontColor.bind(this)}
      />
    );
  }

  // Template for font family dropdown in toolbar
  contentTemplate2() {
    return (
      <ComboBoxComponent
        dataSource={fontStyle}
        change={this.changeFontFamily.bind(this)}
        index={2}
        allowCustom={true}
        showClearButton={false}
      />
    );
  }

  // Template for font size dropdown in toolbar
  contentTemplate3() {
    return (
      <ComboBoxComponent
        dataSource={fontSize}
        change={this.changeFontSize.bind(this)}
        index={2}
        allowCustom={true}
        showClearButton={false}
      />
    );
  }

  // Template for line spacing dropdown in toolbar
  contentTemplate4() {
    return (
      <DropDownButtonComponent
        items={lineSpacingItems}
        iconCss="e-de-ctnr-linespacing e-icons"
        select={this.lineSpacingAction.bind(this)}
      />
    );
  }

  // Render the component
  render() {
    return (
      <div>
        <ToolbarComponent
          id="toolbar"
          clicked={this.toolbarButtonClick.bind(this)}
        >
          <ItemsDirective>
            {/* Undo and Redo buttons */}
            <ItemDirective
              id="undo"
              prefixIcon="e-de-ctnr-undo"
              tooltipText="Undo"
            />
            <ItemDirective
              id="redo"
              prefixIcon="e-de-ctnr-redo"
              tooltipText="Redo"
            />
            <ItemDirective type="Separator" />

            {/* Character formatting */}
            <ItemDirective
              id="bold"
              prefixIcon="e-de-ctnr-bold"
              tooltipText="Bold"
            />
            <ItemDirective
              id="italic"
              prefixIcon="e-de-ctnr-italic"
              tooltipText="Italic"
            />
            <ItemDirective
              id="underline"
              prefixIcon="e-de-ctnr-underline"
              tooltipText="Underline"
            />
            <ItemDirective
              id="strikethrough"
              prefixIcon="e-de-ctnr-strikethrough"
              tooltipText="Strikethrough"
            />
            <ItemDirective
              id="subscript"
              prefixIcon="e-de-ctnr-subscript"
              tooltipText="Subscript"
            />
            <ItemDirective
              id="superscript"
              prefixIcon="e-de-ctnr-superscript"
              tooltipText="Superscript"
            />
            <ItemDirective type="Separator" />
            <ItemDirective template={this.contentTemplate1.bind(this)} />
            <ItemDirective type="Separator" />
            <ItemDirective template={this.contentTemplate2.bind(this)} />
            <ItemDirective template={this.contentTemplate3.bind(this)} />
            <ItemDirective type="Separator" />

            {/* Paragraph formatting */}
            <ItemDirective
              id="AlignLeft"
              prefixIcon="e-de-ctnr-alignleft e-icons"
              tooltipText="Align Left"
            />
            <ItemDirective
              id="AlignCenter"
              prefixIcon="e-de-ctnr-aligncenter e-icons"
              tooltipText="Align Center"
            />
            <ItemDirective
              id="AlignRight"
              prefixIcon="e-de-ctnr-alignright e-icons"
              tooltipText="Align Right"
            />
            <ItemDirective
              id="Justify"
              prefixIcon="e-de-ctnr-justify e-icons"
              tooltipText="Justify"
            />
            <ItemDirective type="Separator" />
            <ItemDirective
              id="IncreaseIndent"
              prefixIcon="e-de-ctnr-increaseindent e-icons"
              tooltipText="Increase Indent"
            />
            <ItemDirective
              id="DecreaseIndent"
              prefixIcon="e-de-ctnr-decreaseindent e-icons"
              tooltipText="Decrease Indent"
            />
            <ItemDirective type="Separator" />
            <ItemDirective template={this.contentTemplate4.bind(this)} />
            <ItemDirective
              id="ClearFormat"
              prefixIcon="e-de-ctnr-clearall e-icons"
              tooltipText="Clear Formatting"
            />
            <ItemDirective type="Separator" />
            <ItemDirective
              id="ShowParagraphMark"
              prefixIcon="e-de-e-paragraph-mark e-icons"
              tooltipText="Show the hidden characters like spaces, tab, paragraph marks, and breaks.(Ctrl + *)"
            />
          </ItemsDirective>
        </ToolbarComponent>

        <DocumentEditorContainerComponent
          id="container"
          height={"100%"}
          width={"100%"}
          ref={(scope) => {
            this.documenteditor = scope;
          }}
          //   isReadOnly={false}
          //   enableSelection={true}
          //   enableEditor={true}
          //   enableEditorHistory={true}
          //   enableContextMenu={true}
          //   enableTableDialog={true}enableToolbar={true}
          // toolbarItems={items}
          // toolbarClick={this.onToolbarClick}
          showPropertiesPane={false}
          //   documentChange={(args) => {
          //     this.setState({ documentModified: true });
          //   }}
          //   contentChange={(args) => {
          //     this.setState({ documentModified: true });
          //   }}
          serviceUrl="https://services.syncfusion.com/react/production/api/documenteditor/"
        />
      </div>
    );
  }
}

export default EditorComponent;
