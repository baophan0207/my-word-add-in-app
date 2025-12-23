import React, { useState, useEffect, useRef } from "react";
import Selection from "../BasicComponents/Selection/Selection";
import TextArea from "../BasicComponents/TextArea/TextArea";
import Button from "../BasicComponents/Button/Button";
import "./Enhancements.scss";

/* global console, setTimeout, Word */

// Helper function to format date as "mm/dd/yyyy hh:mm AM/PM"
const formatDate = (date) => {
  const dateStr = date.toLocaleDateString("en-US", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  });

  const timeStr = date.toLocaleTimeString("en-US", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
  });

  return `${dateStr} ${timeStr}`;
};

const formatTimeDisplay = (date, thresholdSeconds = 60) => {
  const now = new Date();
  const diffInSeconds = Math.abs((now - date) / 1000);

  if (diffInSeconds < thresholdSeconds) {
    return "Just now";
  }

  return formatDate(date);
};

const Enhancements = () => {
  const [selectedSection, setSelectedSection] = useState(null);
  const [textContent, setTextContent] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);
  const [lastUpdatedTime, setLastUpdatedTime] = useState(new Date(2025, 12, 21));

  // Track previously highlighted section to clear it when changing selection
  const previousHighlightRef = useRef(null);

  const sections = [
    { label: "Title", value: "title" },
    { label: "Abstract", value: "abstract" },
    { label: "Description", value: "description" },
    { label: "Technical Field", value: "technical-field" },
    { label: "Claims", value: "claims" },
  ];

  // Function to highlight section title in the Word document
  const highlightSectionInDocument = async (sectionValue) => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;

        // Clear ALL previous highlights by searching for all section titles
        for (const section of sections) {
          const searchText = section.label;
          const results = body.search(searchText, { matchCase: false, matchWholeWord: true });
          context.load(results, "items");
          await context.sync();

          // Clear highlights from found sections
          for (let i = 0; i < results.items.length; i++) {
            results.items[i].font.highlightColor = null;
          }
        }
        await context.sync();

        // Now highlight the new section
        const targetSection = sections.find((s) => s.value === sectionValue);
        if (!targetSection) return;

        const searchText = targetSection.label;
        if (!searchText) return;

        // Search for the section title in the document
        const searchResults = body.search(searchText, { matchCase: false, matchWholeWord: true });
        context.load(searchResults, "items");
        await context.sync();

        if (searchResults.items.length > 0) {
          // Load font properties to find the bold/heading version
          for (let i = 0; i < searchResults.items.length; i++) {
            context.load(searchResults.items[i], "font/bold, style");
          }
          await context.sync();

          // Find the first bold occurrence
          let targetRange = searchResults.items[0];
          for (let i = 0; i < searchResults.items.length; i++) {
            if (searchResults.items[i].font.bold === true) {
              targetRange = searchResults.items[i];
              break;
            }
          }

          // Apply highlight color (Yellow)
          targetRange.font.highlightColor = "Yellow";
          targetRange.select();

          await context.sync();

          previousHighlightRef.current = sectionValue;
          console.log(`Highlighted: ${searchText}`);
        } else {
          console.log(`Section "${searchText}" not found in document`);
        }
      });
    } catch (error) {
      console.error("Error highlighting section:", error);
    }
  };

  const handleSectionChange = (value) => {
    setSelectedSection(value);

    // Highlight the corresponding section in the document
    if (value) {
      highlightSectionInDocument(value);
    }
  };

  const handleTextChange = (value) => {
    setTextContent(value);
  };

  const handleEnhanceClick = () => {
    if (isProcessing) return;

    setIsProcessing(true);
    setTimeout(() => {
      setIsProcessing(false);
      setLastUpdatedTime(new Date());
    }, 5000);
  };

  return (
    <div className="enhancements-container">
      <div className="enhancements-main">
        {/* Section selection */}
        <div className="enhancements-content">
          <div className="enhancements-title">Choose Section</div>
          <div className="enhancements-subtitle">Choose the sections you want to enhance</div>
          <Selection
            list={sections}
            selected={selectedSection}
            onChange={handleSectionChange}
            placeholder="- Choose Section -"
          />
        </div>

        {/* Instructions textarea */}
        <div className="enhancements-content flex-1">
          <div className="enhancements-title">Instruction</div>
          <div className="enhancements-subtitle">Enter how you want to modify</div>
          <div className="textarea-container">
            <TextArea
              value={textContent}
              onChange={handleTextChange}
              placeholder="Enter your instructions here..."
              fullHeight={true}
              scrollable={true}
              height={100}
            />
          </div>
        </div>

        {/* Action button */}
        <div className="button-container">
          <Button
            type="primary"
            onClick={handleEnhanceClick}
            disabled={isProcessing || !selectedSection || !textContent}
          >
            {isProcessing ? "Enhancing..." : "Enhance with AI"}
          </Button>
        </div>
      </div>

      {/* Footer with status */}
      <div className="document-status-container">
        {isProcessing ? (
          <>
            <span>Document Updating ...</span>
          </>
        ) : (
          <>
            <span>Document Last Updated:</span>
            <span className={`enhancements-document-status-time`}>{formatTimeDisplay(lastUpdatedTime)}</span>
          </>
        )}
      </div>
    </div>
  );
};

export default Enhancements;
