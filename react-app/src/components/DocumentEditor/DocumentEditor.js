import * as React from "react";
import {
  DocumentEditorContainerComponent,
  Toolbar,
  OptionsPane,
} from "@syncfusion/ej2-react-documenteditor";
import "./DocumentEditor.css";

DocumentEditorContainerComponent.Inject(Toolbar, OptionsPane);

class DocumentEditor extends React.Component {
  constructor(props) {
    super(props);

    this.state = {
      documents: [],
      selectedDocument: null,
      status: "",
      autoSaveEnabled: false,
      lastSaved: null,
      isSaving: false,
      documentModified: false,
      showFontColorPicker: false,
      showHighlightColorPicker: false,
      fontColorPickerPosition: { top: 0, left: 0 },
      highlightColorPickerPosition: { top: 0, left: 0 },
    };

    // Refs
    this.editorRef = React.createRef();
    this.autoSaveIntervalRef = null;
    this.fontColorButtonRef = React.createRef();
    this.highlightColorButtonRef = React.createRef();

    // Bind methods to this
    this.loadDocuments = this.loadDocuments.bind(this);
    this.openDocument = this.openDocument.bind(this);
    this.saveDocument = this.saveDocument.bind(this);
    this.setupEditorEvents = this.setupEditorEvents.bind(this);
    this.setupAutoSave = this.setupAutoSave.bind(this);
    this.cleanupAutoSave = this.cleanupAutoSave.bind(this);
    this.onToolbarClick = this.onToolbarClick.bind(this);
    this.toggleFontColorPicker = this.toggleFontColorPicker.bind(this);
    this.toggleHighlightColorPicker =
      this.toggleHighlightColorPicker.bind(this);
    this.applyFontColor = this.applyFontColor.bind(this);
    this.applyHighlightColor = this.applyHighlightColor.bind(this);
    this.closeColorPickers = this.closeColorPickers.bind(this);
    this.ensureSelection = this.ensureSelection.bind(this);
    this.debugSelectionState = this.debugSelectionState.bind(this);
    this.applyColorWithDOMCheck = this.applyColorWithDOMCheck.bind(this);
    this.updateFormatButtonStates = this.updateFormatButtonStates.bind(this);
    this.updateButtonToggleState = this.updateButtonToggleState.bind(this);
  }

  // Custom toolbar items definition
  get customToolbarItems() {
    // Define custom toolbar items for text formatting
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

    // Return the array of toolbar items
    return [
      // First group: Undo, Redo
      "Undo",
      "Redo",
      // "|",

      // Second group: Font controls
      fontFamilyItem,
      fontSizeItem,
      // "|",

      // Third group: Text formatting
      boldItem,
      italicItem,
      underlineItem,
      // "|",
      fontColorItem,
      highlightColorItem,
      // "|",

      // Fourth group: Paragraph alignment
      alignLeftItem,
      alignCenterItem,
      alignRightItem,
      // "|",

      // Fifth group: Lists and indentation
      bulletListItem,
      numberedListItem,
      decreaseIndentItem,
      increaseIndentItem,
      // "|",

      // Sixth group: Insert items
      "Image",
      "Table",
      // "|",

      // Seventh group: Collaboration
      commentItem,
      trackChangesItem,
      // "|",

      // Eighth group: Find
      "Find",
    ];
  }

  componentDidMount() {
    this.loadDocuments();
    this.setupEditorEvents();
    this.setupAutoSave();

    // Add event listener to close color pickers when clicking outside
    document.addEventListener("mousedown", this.closeColorPickers);

    // We'll initialize custom dropdowns like font family and size after the editor is ready
    setTimeout(() => {
      this.initializeCustomToolbarItems();

      // Set references to the color buttons
      this.fontColorButtonRef.current = document.getElementById("FontColor");
      this.highlightColorButtonRef.current =
        document.getElementById("HighlightColor");

      // Initialize format button states after editor is ready
      this.updateFormatButtonStates();
    }, 500);
  }

  componentDidUpdate(prevProps, prevState) {
    // Setup editor events if editor ref changes
    if (this.editorRef.current !== prevState.editorRef) {
      this.setupEditorEvents();
    }

    // Setup auto-save if relevant state changes
    if (
      prevState.selectedDocument !== this.state.selectedDocument ||
      prevState.isSaving !== this.state.isSaving ||
      prevState.documentModified !== this.state.documentModified
    ) {
      this.setupAutoSave();
    }
  }

  componentWillUnmount() {
    // Remove the event listener
    document.removeEventListener("mousedown", this.closeColorPickers);

    this.cleanupAutoSave();
  }

  initializeCustomToolbarItems() {
    // This method will initialize any custom dropdown items
    // that we need to handle separately from standard toolbar items
    const editorContainer = this.editorRef.current;

    if (!editorContainer) return;

    // Get references to the placeholder elements in the toolbar
    const fontFamilyContainer = document.getElementById("fontFamily");
    const fontSizeContainer = document.getElementById("fontSize");

    if (fontFamilyContainer) {
      // Create a select element for font family
      const fontFamilySelect = document.createElement("select");
      fontFamilySelect.className = "gdocs-toolbar-select";

      // Add common font families
      const fontFamilies = [
        "Arial",
        "Calibri",
        "Cambria",
        "Courier New",
        "Georgia",
        "Times New Roman",
        "Trebuchet MS",
        "Verdana",
      ];

      fontFamilies.forEach((font) => {
        const option = document.createElement("option");
        option.value = font;
        option.textContent = font;
        fontFamilySelect.appendChild(option);
      });

      // Add event listener
      fontFamilySelect.addEventListener("change", () => {
        if (editorContainer.documentEditor) {
          // The correct way to apply character format in Syncfusion DocumentEditor
          // is to set properties directly on the selection's characterFormat
          editorContainer.documentEditor.selection.characterFormat.fontFamily =
            fontFamilySelect.value;
          // Make sure to focus back on the editor
          editorContainer.documentEditor.focusIn();
        }
      });

      fontFamilyContainer.appendChild(fontFamilySelect);
    }

    if (fontSizeContainer) {
      // Create a select element for font size
      const fontSizeSelect = document.createElement("select");
      fontSizeSelect.className = "gdocs-toolbar-select";

      // Add common font sizes
      const fontSizes = [
        8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
      ];

      fontSizes.forEach((size) => {
        const option = document.createElement("option");
        option.value = size;
        option.textContent = size;
        fontSizeSelect.appendChild(option);
      });

      // Set default size
      fontSizeSelect.value = 11;

      // Add event listener
      fontSizeSelect.addEventListener("change", () => {
        if (editorContainer.documentEditor) {
          // The correct way to apply font size in Syncfusion DocumentEditor
          editorContainer.documentEditor.selection.characterFormat.fontSize =
            parseFloat(fontSizeSelect.value);
          // Make sure to focus back on the editor
          editorContainer.documentEditor.focusIn();
        }
      });

      fontSizeContainer.appendChild(fontSizeSelect);
    }

    // Add code to update the dropdowns when selection changes
    if (editorContainer.documentEditor) {
      editorContainer.documentEditor.selectionChange = () => {
        // Update font family dropdown
        if (fontFamilyContainer && fontFamilyContainer.firstChild) {
          const currentFontFamily =
            editorContainer.documentEditor.selection.characterFormat.fontFamily;
          if (currentFontFamily) {
            fontFamilyContainer.firstChild.value = currentFontFamily;
          }
        }

        // Update font size dropdown
        if (fontSizeContainer && fontSizeContainer.firstChild) {
          const currentFontSize =
            editorContainer.documentEditor.selection.characterFormat.fontSize;
          if (currentFontSize) {
            fontSizeContainer.firstChild.value = currentFontSize.toString();
          }
        }
      };
    }
  }

  setupEditorEvents() {
    if (this.editorRef.current && this.editorRef.current.documentEditor) {
      // Listen for content changes
      this.editorRef.current.documentEditor.contentChange = () => {
        this.setState({ documentModified: true });
      };

      // Add selection change event to update formatting button states
      this.editorRef.current.documentEditor.selectionChange = () => {
        // Update toggle button states when selection changes
        this.updateFormatButtonStates();
      };

      // Get document change status
      const documentStatus = this.editorRef.current.documentEditor;
      console.log("Document status:", documentStatus);
    }
  }

  // Add this new method to update format button states
  updateFormatButtonStates() {
    if (!this.editorRef.current || !this.editorRef.current.documentEditor)
      return;

    const documentEditor = this.editorRef.current.documentEditor;
    const container = this.editorRef.current.element;
    if (!container) return;

    // Look for toolbar element that contains the buttons
    const toolbar =
      container.querySelector(".e-toolbar") ||
      container.querySelector(".e-de-toolbar");
    if (!toolbar) return;

    try {
      // Get current formatting state from the selection
      const selection = documentEditor.selection;
      if (!selection || !selection.characterFormat) return;

      const charFormat = selection.characterFormat;
      const paraFormat = selection.paragraphFormat;

      // Check bold state
      const isBold = charFormat.bold;
      const boldButton = toolbar.querySelector('[id$="Bold"], [title="Bold"]');
      this.updateButtonToggleState(boldButton, isBold);

      // Check italic state
      const isItalic = charFormat.italic;
      const italicButton = toolbar.querySelector(
        '[id$="Italic"], [title="Italic"]'
      );
      this.updateButtonToggleState(italicButton, isItalic);

      // Check underline state
      const isUnderline = charFormat.underline !== "None";
      const underlineButton = toolbar.querySelector(
        '[id$="Underline"], [title="Underline"]'
      );
      this.updateButtonToggleState(underlineButton, isUnderline);

      // Check alignment states
      if (paraFormat) {
        const alignment = paraFormat.textAlignment;

        const leftAlignButton = toolbar.querySelector(
          '[id$="AlignLeft"], [title="Align Left"]'
        );
        this.updateButtonToggleState(leftAlignButton, alignment === "Left");

        const centerAlignButton = toolbar.querySelector(
          '[id$="AlignCenter"], [title="Align Center"]'
        );
        this.updateButtonToggleState(centerAlignButton, alignment === "Center");

        const rightAlignButton = toolbar.querySelector(
          '[id$="AlignRight"], [title="Align Right"]'
        );
        this.updateButtonToggleState(rightAlignButton, alignment === "Right");

        // Check list states
        const hasBulletList = paraFormat.listType === "Bullet";
        const bulletListButton = toolbar.querySelector(
          '[id$="BulletList"], [title="Bullet List"]'
        );
        this.updateButtonToggleState(bulletListButton, hasBulletList);

        const hasNumberedList = paraFormat.listType === "Numbered";
        const numberedListButton = toolbar.querySelector(
          '[id$="NumberedList"], [title="Numbered List"]'
        );
        this.updateButtonToggleState(numberedListButton, hasNumberedList);
      }
    } catch (error) {
      console.log("Error updating format button states:", error);
    }
  }

  // Helper method to update button toggle state
  updateButtonToggleState(button, isActive) {
    if (!button) return;

    // Add/remove toggle state class
    if (isActive) {
      button.classList.add("e-active", "e-btn-toggle");
    } else {
      button.classList.remove("e-active", "e-btn-toggle");
    }

    // Update aria-pressed attribute for accessibility
    button.setAttribute("aria-pressed", isActive ? "true" : "false");

    // Update parent container if needed
    const parent = button.closest(".e-toolbar-item");
    if (parent) {
      if (isActive) {
        parent.classList.add("e-active");
      } else {
        parent.classList.remove("e-active");
      }
    }
  }

  setupAutoSave() {
    // Clear any existing interval
    this.cleanupAutoSave();

    const { selectedDocument, isSaving, documentModified } = this.state;

    // Set up auto-save if a document is selected
    if (selectedDocument) {
      this.autoSaveIntervalRef = setInterval(() => {
        // Only save if the editor exists, not already saving, and document has changes
        if (this.editorRef.current && !isSaving && documentModified) {
          this.saveDocument(true); // Pass true to indicate it's an auto-save
        }
      }, 10000); // Auto-save every 10 seconds
    }
  }

  cleanupAutoSave() {
    if (this.autoSaveIntervalRef) {
      clearInterval(this.autoSaveIntervalRef);
      this.autoSaveIntervalRef = null;
    }
  }

  async loadDocuments() {
    try {
      const response = await fetch(
        `http://${process.env.REACT_APP_HOST}:${process.env.REACT_APP_NODE_SERVER_PORT}/api/documents`
      );
      const documents = await response.json();
      this.setState({ documents });
    } catch (error) {
      console.error("Error loading documents:", error);
      this.setState({ status: "Error loading documents" });
    }
  }

  async openDocument(doc) {
    try {
      this.setState({ status: `Opening document "${doc.name}"...` });

      // Check if the URL is properly formatted
      let documentUrl = doc.url;

      console.log("Fetching document from:", documentUrl);

      // Fetch the document content from the server
      const response = await fetch(documentUrl);

      if (!response.ok) {
        throw new Error(
          `Server returned ${response.status}: ${response.statusText}`
        );
      }

      const blob = await response.blob();

      // Open the document in the editor
      if (this.editorRef.current) {
        try {
          // Enable local storage if needed
          this.editorRef.current.documentEditor.enableLocalStorage = true;

          // Convert blob to base64 string
          const reader = new FileReader();
          reader.onload = () => {
            try {
              const base64data = reader.result.split(",")[1];

              // Use the correct method to open the document based on type
              const fileName = doc.name.toLowerCase();

              // The DocumentEditor supports different import types, but we need to use the correct method
              if (
                fileName.endsWith(".docx") ||
                fileName.endsWith(".doc") ||
                fileName.endsWith(".rtf") ||
                fileName.endsWith(".txt")
              ) {
                // Use the open method with the appropriate format parameter
                this.editorRef.current.documentEditor.open(
                  base64data,
                  this.getFormatType(fileName)
                );

                this.setState({
                  selectedDocument: doc,
                  status: `Document "${doc.name}" opened successfully`,
                  documentModified: false,
                });

                // Clear status after a few seconds
                setTimeout(() => this.setState({ status: "" }), 3000);
              } else {
                throw new Error(
                  "Unsupported file format. Please use .docx, .doc, .rtf, or .txt files."
                );
              }
            } catch (parseError) {
              console.error("Error parsing document:", parseError);
              this.setState({
                status: `Error parsing document: ${parseError.message}`,
              });
            }
          };

          reader.onerror = (error) => {
            console.error("Error reading file:", error);
            this.setState({ status: "Error reading file. Please try again." });
          };

          reader.readAsDataURL(blob);
        } catch (openError) {
          console.error("Error opening document in editor:", openError);
          this.setState({
            status: `Error opening document: ${openError.message}`,
          });
        }
      }
    } catch (error) {
      console.error("Error fetching document:", error);
      this.setState({ status: `Error fetching document: ${error.message}` });
    }
  }

  // Helper method to determine the format type for the document editor
  getFormatType(fileName) {
    if (fileName.endsWith(".docx")) {
      return "Docx";
    } else if (fileName.endsWith(".doc")) {
      return "Doc";
    } else if (fileName.endsWith(".rtf")) {
      return "Rtf";
    } else if (fileName.endsWith(".txt")) {
      return "Txt";
    } else {
      return "Docx"; // Default format
    }
  }

  async saveDocument(isAutoSave = false) {
    const { selectedDocument } = this.state;

    if (!selectedDocument || !this.editorRef.current) {
      this.setState({ status: "No document selected to save" });
      return;
    }

    try {
      this.setState({ isSaving: true });

      if (!isAutoSave) {
        this.setState({
          status: `Saving document "${selectedDocument.name}"...`,
        });
      } else {
        this.setState({ status: `Auto-saving "${selectedDocument.name}"...` });
      }

      // Get document content as blob without triggering download
      this.editorRef.current.documentEditor
        .saveAsBlob("Docx")
        .then(async (blob) => {
          // Create form data
          const formData = new FormData();
          const file = new File([blob], selectedDocument.name, {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });

          formData.append("document", file);
          formData.append(
            "metadata",
            JSON.stringify({
              documentName: selectedDocument.name,
              timestamp: new Date().toISOString(),
              eventType: isAutoSave ? "auto-save" : "save",
            })
          );

          // Send to server
          const response = await fetch(
            `http://${process.env.REACT_APP_HOST}:${process.env.REACT_APP_NODE_SERVER_PORT}/api/upload-document-with-metadata`,
            {
              method: "POST",
              body: formData,
            }
          );

          const result = await response.json();

          if (response.ok) {
            const currentTime = new Date();
            this.setState({
              lastSaved: currentTime,
              documentModified: false, // Reset modified flag after successful save
            });

            if (!isAutoSave) {
              this.setState({
                status: `Document "${selectedDocument.name}" saved successfully`,
              });
              // Clear status after a few seconds for manual saves
              setTimeout(() => this.setState({ status: "" }), 3000);
            } else {
              this.setState({
                status: `Auto-saved at ${currentTime.toLocaleTimeString()}`,
              });
            }

            // Refresh document list
            await this.loadDocuments();
          } else {
            throw new Error(result.error || "Unknown error occurred");
          }
        });
    } catch (error) {
      console.error("Error saving document:", error);
      this.setState({ status: `Error saving document: ${error.message}` });
    } finally {
      this.setState({ isSaving: false });
    }
  }

  // Handle toolbar button clicks
  onToolbarClick(args) {
    const documentEditor = this.editorRef.current?.documentEditor;
    if (!documentEditor) return;

    // Process click based on the item ID
    switch (args.item.id) {
      case "Undo":
        // After default handler runs, refresh container and update buttons
        setTimeout(() => {
          // Safely call refresh with additional checks
          if (
            this.editorRef.current &&
            typeof this.editorRef.current.refresh === "function" &&
            this.editorRef.current.optionsPaneModule
          ) {
            this.editorRef.current.refresh();
          }
        }, 100); // Increase timeout to ensure component is ready
        break;
      case "Redo":
        // After default handler runs, refresh container and update buttons
        setTimeout(() => {
          // Safely call refresh with additional checks
          if (
            this.editorRef.current &&
            typeof this.editorRef.current.refresh === "function" &&
            this.editorRef.current.optionsPaneModule
          ) {
            this.editorRef.current.refresh();
          }
        }, 100); // Increase timeout to ensure component is ready
        break;
      case "Bold":
        documentEditor.editor.toggleBold();
        // Update format button states after toggling
        setTimeout(() => this.updateFormatButtonStates(), 0);
        break;
      case "Italic":
        documentEditor.editor.toggleItalic();
        setTimeout(() => this.updateFormatButtonStates(), 0);
        break;
      case "Underline":
        documentEditor.editor.toggleUnderline("Single");
        setTimeout(() => this.updateFormatButtonStates(), 0);
        break;
      case "AlignLeft":
        documentEditor.editor.toggleTextAlignment("Left");
        setTimeout(() => this.updateFormatButtonStates(), 0);
        break;
      case "AlignCenter":
        documentEditor.editor.toggleTextAlignment("Center");
        setTimeout(() => this.updateFormatButtonStates(), 0);
        break;
      case "AlignRight":
        documentEditor.editor.toggleTextAlignment("Right");
        setTimeout(() => this.updateFormatButtonStates(), 0);
        break;
      case "DecreaseIndent":
        documentEditor.editor.decreaseIndent();
        break;
      case "IncreaseIndent":
        documentEditor.editor.increaseIndent();
        break;
      case "BulletList":
        // Apply a simple bullet list
        documentEditor.editor.applyBullet("•", "Symbol");
        setTimeout(() => this.updateFormatButtonStates(), 0);
        break;
      case "NumberedList":
        // Apply numbered list with the default format
        documentEditor.editor.applyNumbering("%1.", "Arabic");
        setTimeout(() => this.updateFormatButtonStates(), 0);
        break;
      case "HighlightColor":
        // Show highlight color picker instead of applying a fixed color
        this.toggleHighlightColorPicker();
        break;
      case "FontColor":
        // Show font color picker instead of applying a fixed color
        this.toggleFontColorPicker();
        break;
      case "AddComment":
        documentEditor.editor.insertComment("");
        break;
      case "TrackChanges":
        documentEditor.enableTrackChanges = !documentEditor.enableTrackChanges;
        // Update button appearance to indicate toggle state
        break;
      default:
        // Let the default handler handle other built-in actions
        break;
    }
  }

  toggleFontColorPicker() {
    if (!this.fontColorButtonRef.current) return;

    const rect = this.fontColorButtonRef.current.getBoundingClientRect();
    this.setState({
      showFontColorPicker: !this.state.showFontColorPicker,
      showHighlightColorPicker: false,
      fontColorPickerPosition: {
        top: rect.bottom + window.scrollY,
        left: rect.left + window.scrollX,
      },
    });
  }

  toggleHighlightColorPicker() {
    if (!this.highlightColorButtonRef.current) return;

    const rect = this.highlightColorButtonRef.current.getBoundingClientRect();
    this.setState({
      showHighlightColorPicker: !this.state.showHighlightColorPicker,
      showFontColorPicker: false,
      highlightColorPickerPosition: {
        top: rect.bottom + window.scrollY,
        left: rect.left + window.scrollX,
      },
    });
  }

  closeColorPickers(event) {
    // If this is a mousedown event and it's inside the color picker, don't close
    if (event && event.target) {
      const fontColorPicker = document.querySelector(".color-picker-popup");
      if (fontColorPicker && fontColorPicker.contains(event.target)) {
        return;
      }
    }

    this.setState({
      showFontColorPicker: false,
      showHighlightColorPicker: false,
    });
  }

  applyFontColor(color) {
    this.applyColorWithDOMCheck("font", color);
  }

  applyHighlightColor(color) {
    this.applyColorWithDOMCheck("highlight", color);
  }

  applyColorWithDOMCheck(action, color) {
    const documentEditor = this.editorRef.current?.documentEditor;
    if (!documentEditor) return;

    // Store the selection info before applying
    const selectionBefore = {
      text: documentEditor.selection.text,
      isEmpty: documentEditor.selection.isEmpty,
    };

    console.log(`Attempting to apply ${action} color:`, color);
    console.log("Selection before:", selectionBefore);

    try {
      // Apply the formatting
      if (action === "font") {
        const colorValue = color === "none" ? "empty" : color;
        documentEditor.selection.characterFormat.fontColor = colorValue;
      } else {
        const highlightValue =
          color === "none" ? "NoColor" : this.getHighlightColorName(color);
        documentEditor.selection.characterFormat.highlightColor =
          highlightValue;
      }

      // Check if selection is still valid after applying
      setTimeout(() => {
        const selectionAfter = {
          text: documentEditor.selection.text,
          isEmpty: documentEditor.selection.isEmpty,
        };
        console.log("Selection after:", selectionAfter);

        // If selection is empty after formatting, try reselecting
        if (selectionAfter.isEmpty && !selectionBefore.isEmpty) {
          console.log(
            "Selection lost after applying color, attempting to restore"
          );
          documentEditor.selection.selectAll();
          documentEditor.selection.fireSelectionChanged(true);
        }

        documentEditor.focusIn();

        // Close the color picker after applying the color
        this.setState({
          showFontColorPicker: false,
          showHighlightColorPicker: false,
        });
      }, 50);
    } catch (error) {
      console.error(`Error in applyColorWithDOMCheck (${action}):`, error);
    }
  }

  getHighlightColorName(hexColor) {
    // Map hex colors to Syncfusion highlight color names
    const colorMap = {
      "#ffff00": "Yellow",
      "#00ff00": "BrightGreen",
      "#00ffff": "Turquoise",
      "#ff00ff": "Pink",
      "#0000ff": "Blue",
      "#ff0000": "Red",
      "#000080": "DarkBlue",
      "#008080": "Teal",
      "#008000": "Green",
      "#800080": "Violet",
      "#800000": "DarkRed",
      "#808000": "DarkYellow",
      "#808080": "Gray50",
      "#c0c0c0": "Gray25",
      "#000000": "Black",
    };

    // For custom colors not in our map, try to match to the closest color
    const lowerHex = hexColor.toLowerCase();

    // Direct match
    if (colorMap[lowerHex]) {
      return colorMap[lowerHex];
    }

    // For colors not in our predefined list, use Yellow as default
    // This is a limitation of Syncfusion - it only supports specific highlight colors
    console.log("Using default Yellow highlight for custom color:", hexColor);
    return "Yellow";
  }

  ensureSelection() {
    const documentEditor = this.editorRef.current?.documentEditor;
    if (!documentEditor) return false;

    // If no text is selected, try multiple approaches
    if (documentEditor.selection && documentEditor.selection.isEmpty) {
      try {
        // Method 1: Select current word
        documentEditor.selection.selectCurrentWord();

        // If still empty, try method 2: Select all
        if (documentEditor.selection.isEmpty) {
          console.log("Word selection failed, trying to select all");
          documentEditor.selection.selectAll();
        }

        // Log what we got
        console.log("Selection after ensure:", {
          text: documentEditor.selection.text,
          isEmpty: documentEditor.selection.isEmpty,
        });

        return !documentEditor.selection.isEmpty;
      } catch (error) {
        console.error("Selection error:", error);
        return false;
      }
    }

    return true;
  }

  debugSelectionState() {
    const documentEditor = this.editorRef.current?.documentEditor;
    if (!documentEditor || !documentEditor.selection) return;

    console.log("Selection state:", {
      isEmpty: documentEditor.selection.isEmpty,
      text: documentEditor.selection.text,
      characterFormat: documentEditor.selection.characterFormat,
    });
  }

  render() {
    const {
      documents,
      selectedDocument,
      status,
      autoSaveEnabled,
      lastSaved,
      isSaving,
      documentModified,
      showFontColorPicker,
      showHighlightColorPicker,
      fontColorPickerPosition,
      highlightColorPickerPosition,
    } = this.state;

    // Common color palette for both pickers
    const colorPalette = [
      { color: "none", label: "No Color" },
      { color: "#000000", label: "Black" },
      { color: "#808080", label: "Gray" },
      { color: "#ff0000", label: "Red" },
      { color: "#ff8000", label: "Orange" },
      { color: "#ffff00", label: "Yellow" },
      { color: "#00ff00", label: "Green" },
      { color: "#00ffff", label: "Cyan" },
      { color: "#0000ff", label: "Blue" },
      { color: "#8000ff", label: "Purple" },
      { color: "#ff00ff", label: "Magenta" },
      { color: "#800000", label: "Dark Red" },
      { color: "#808000", label: "Olive" },
      { color: "#008000", label: "Dark Green" },
      { color: "#008080", label: "Teal" },
      { color: "#000080", label: "Navy" },
    ];

    return (
      <div className="document-editor-container">
        <div className="document-list-panel">
          <h3>Available Documents</h3>
          {documents.length === 0 ? (
            <p>No documents available</p>
          ) : (
            <ul className="document-list">
              {documents.map((doc) => (
                <li
                  key={doc.id}
                  className={selectedDocument?.id === doc.id ? "selected" : ""}
                  onClick={() => this.openDocument(doc)}
                >
                  {doc.name}
                  {selectedDocument?.id === doc.id && documentModified && (
                    <span className="modified-indicator">*</span>
                  )}
                </li>
              ))}
            </ul>
          )}
          {status && <div className="status-message">{status}</div>}

          <div className="button-group">
            <button
              className="save-button"
              onClick={() => this.saveDocument(false)}
              disabled={!selectedDocument || isSaving || !documentModified}
            >
              {isSaving && !autoSaveEnabled ? "Saving..." : "Save Document"}
            </button>
          </div>

          {lastSaved && (
            <div className="last-saved-info">
              Last saved: {lastSaved.toLocaleTimeString()}
              {documentModified && (
                <span className="modified-indicator"> (modified)</span>
              )}
            </div>
          )}
        </div>

        <div className="editor-panel">
          <div className="document-title">
            {selectedDocument ? selectedDocument.name : "Untitled Document"}
          </div>
          <DocumentEditorContainerComponent
            ref={this.editorRef}
            id="container"
            height={"calc(100% - 40px)"}
            width={"100%"}
            enableToolbar={true}
            toolbarItems={this.customToolbarItems}
            toolbarClick={this.onToolbarClick}
            showPropertiesPane={false}
            // documentChange={(args) => {
            //   this.setState({ documentModified: true });
            // }}
            contentChange={(args) => {
              this.setState({ documentModified: true });
            }}
            serviceUrl="https://services.syncfusion.com/js/production/api/documenteditor/"
          />
        </div>

        {/* Enhanced Font Color Picker Popup */}
        {showFontColorPicker && (
          <div
            className="color-picker-popup"
            style={{
              display: "block",
              top: fontColorPickerPosition.top,
              left: fontColorPickerPosition.left,
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="color-picker-header">
              <span className="color-picker-title">Text Color</span>
              <button
                className="color-picker-close"
                onClick={this.closeColorPickers}
              >
                ×
              </button>
            </div>
            <div className="color-palette">
              {colorPalette.map((item) => (
                <div
                  key={item.color}
                  className={`color-cell ${
                    item.color === "none" ? "no-color-cell" : ""
                  }`}
                  style={{
                    backgroundColor:
                      item.color !== "none" ? item.color : "white",
                  }}
                  title={item.label}
                  onClick={(e) => {
                    e.stopPropagation(); // Prevent event bubbling
                    this.applyFontColor(item.color);
                  }}
                />
              ))}
            </div>
            {/* Add custom color input */}
            <div className="custom-color-section">
              <input
                type="color"
                className="custom-color-input"
                onInput={(e) => this.applyFontColor(e.target.value)}
                onClick={(e) => e.stopPropagation()}
              />
              <span>Custom</span>
            </div>
            <div className="color-picker-footer">
              <button
                className="apply-color-button"
                onClick={this.closeColorPickers}
              >
                Done
              </button>
            </div>
          </div>
        )}

        {/* Enhanced Highlight Color Picker Popup */}
        {showHighlightColorPicker && (
          <div
            className="color-picker-popup"
            style={{
              display: "block",
              top: highlightColorPickerPosition.top,
              left: highlightColorPickerPosition.left,
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="color-picker-header">
              <span className="color-picker-title">Highlight Color</span>
              <button
                className="color-picker-close"
                onClick={this.closeColorPickers}
              >
                ×
              </button>
            </div>
            <div className="color-palette">
              {colorPalette.map((item) => (
                <div
                  key={item.color}
                  className={`color-cell ${
                    item.color === "none" ? "no-color-cell" : ""
                  }`}
                  style={{
                    backgroundColor:
                      item.color !== "none" ? item.color : "white",
                  }}
                  title={item.label}
                  onClick={(e) => {
                    e.stopPropagation(); // Prevent event bubbling
                    this.applyHighlightColor(item.color);
                  }}
                />
              ))}
            </div>
            {/* Add custom color input */}
            <div className="custom-color-section">
              <input
                type="color"
                className="custom-color-input"
                onInput={(e) => this.applyHighlightColor(e.target.value)}
                onClick={(e) => e.stopPropagation()}
              />
              <span>Custom</span>
            </div>
            <div className="color-picker-footer">
              <button
                className="apply-color-button"
                onClick={this.closeColorPickers}
              >
                Done
              </button>
            </div>
          </div>
        )}
      </div>
    );
  }
}

export default DocumentEditor;
