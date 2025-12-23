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
  const [sections, setSections] = useState([]);
  const [selectedSection, setSelectedSection] = useState(null);
  const [textContent, setTextContent] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);
  const [lastUpdatedTime, setLastUpdatedTime] = useState(new Date(2025, 12, 21));

  // Track previously highlighted section to clear it when changing selection
  const previousHighlightRef = useRef(null);

  useEffect(() => {
    extractContentByHeadings();
  }, []);

  // Function to highlight section title in the Word document
  const highlightSectionInDocument = async (sectionValue) => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items, style, text");

        await context.sync();

        // Clear all previous highlights globally
        body.font.highlightColor = null;

        let startIndex = -1;
        let endIndex = -1;

        // Regex to match "Heading 1" through "Heading 6", with optional space
        const headingRegex = /^Heading\s?[1-6]$/i;

        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          const isHeading = headingRegex.test(para.style);

          if (startIndex === -1) {
            // Find start of the section
            if (para.text === sectionValue && isHeading) {
              startIndex = i;
            }
          } else {
            // Find end of the section (next heading)
            if (isHeading) {
              endIndex = i - 1;
              break;
            }
          }
        }

        if (startIndex !== -1) {
          // If no next heading found, section goes to end of document
          if (endIndex === -1) {
            endIndex = paragraphs.items.length - 1;
          }

          // Adjust start index to skip the heading itself
          const contentStartIndex = startIndex + 1;

          if (contentStartIndex <= endIndex) {
            const startRange = paragraphs.items[contentStartIndex].getRange("Start");
            const endRange = paragraphs.items[endIndex].getRange("End");
            const sectionRange = startRange.expandTo(endRange);

            sectionRange.select();
            sectionRange.font.highlightColor = "Yellow";

            await context.sync();

            previousHighlightRef.current = sectionValue;
            console.log(`Highlighted content of section: ${sectionValue}`);
          } else {
            console.log(`Section "${sectionValue}" has no content to highlight.`);
            paragraphs.items[startIndex].select();
            await context.sync();
          }
        } else {
          console.log(`Section "${sectionValue}" not found.`);
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

  async function extractContentByHeadings() {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items, style, text");

      await context.sync();

      let currentSection = null;
      const sections = [];

      // Regex to match "Heading 1" through "Heading 6", with optional space
      const headingRegex = /^Heading\s?([1-6])$/i;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const match = headingRegex.exec(para.style);

        if (match) {
          // Found a heading (Level 1-6)
          const level = parseInt(match[1]);

          currentSection = {
            title: para.text,
            level: level,
            content: [],
          };
          sections.push(currentSection);
        } else if (currentSection) {
          // Add content to current section
          currentSection.content.push(para.text);
        }
      }

      const convertSectionToSelection = (section) => {
        return {
          label: section.title,
          value: section.title,
          level: section.level,
          content: section.content,
        };
      };

      setSections(sections.map(convertSectionToSelection));
    });
  }

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
