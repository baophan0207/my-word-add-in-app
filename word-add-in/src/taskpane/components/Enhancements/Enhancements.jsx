import React, { useState } from "react";
import Selection from "../BasicComponents/Selection/Selection";
import TextArea from "../BasicComponents/TextArea/TextArea";
import Button from "../BasicComponents/Button/Button";
import "./Enhancements.scss";

/* global console, setTimeout */

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

const Enhancements = () => {
  const [selectedSection, setSelectedSection] = useState(null);
  const [textContent, setTextContent] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);
  const [timeDisplay, setTimeDisplay] = useState(formatDate(new Date()));

  const sections = [
    { label: "Title", value: "title" },
    { label: "Abstract", value: "abstract" },
    { label: "Description", value: "description" },
    { label: "Technical Field", value: "technical-field" },
    { label: "Claims", value: "claims" },
  ];

  const handleSectionChange = (value) => {
    setSelectedSection(value);
    // eslint-disable-next-line no-console
    console.log("Selected section:", value);
  };

  const handleTextChange = (value) => {
    setTextContent(value);
  };

  const handleEnhanceClick = () => {
    if (isProcessing) return;

    // Start processing
    setIsProcessing(true);

    // Simulate processing for 5 seconds
    setTimeout(() => {
      // Processing done
      setIsProcessing(false);
      setTimeDisplay("Just Now");

      // Clear form inputs
      setSelectedSection(null);
      setTextContent("");

      // After 3 more seconds, show the actual time
      setTimeout(() => {
        setTimeDisplay(formatDate(new Date()));
      }, 3000);

      console.log("Enhancement completed!");
      console.log("Section:", selectedSection);
      console.log("Content:", textContent);
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
              maxHeight={400}
            />
          </div>
        </div>

        {/* Action button */}
        <div className="button-container">
          <Button
            label={isProcessing ? "Enhancing..." : "Enhance with AI"}
            onClick={handleEnhanceClick}
            disabled={isProcessing || !selectedSection || !textContent}
          />
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
            <span className={`enhancements-document-status-time`}>{timeDisplay}</span>
          </>
        )}
      </div>
    </div>
  );
};

export default Enhancements;
