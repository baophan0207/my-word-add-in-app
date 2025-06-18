import React, { useState, useRef, useEffect, useCallback } from "react";
import PropTypes from "prop-types";
import "./TextArea.scss";

/* global setTimeout, clearTimeout */

const TextArea = ({ value, onChange, placeholder, width, height, maxHeight, className, fullHeight, scrollable }) => {
  const [textValue, setTextValue] = useState(value);
  const textareaRef = useRef(null);
  // Use this ref to track resize timeout
  const timeoutRef = useRef(null);

  // Update local state when prop changes
  useEffect(() => {
    setTextValue(value);
    if (!fullHeight) {
      adjustHeight();
    }
  }, [value, fullHeight]);

  // Auto-resize functionality
  const adjustHeight = useCallback(() => {
    if (textareaRef.current && !fullHeight) {
      textareaRef.current.style.height = "auto";
      const scrollHeight = textareaRef.current.scrollHeight;
      const newHeight = scrollHeight;

      // Apply max height if provided
      if (maxHeight) {
        const maxHeightValue = typeof maxHeight === "number" ? maxHeight : parseInt(maxHeight, 10);
        textareaRef.current.style.height = `${Math.min(newHeight, maxHeightValue)}px`;
      } else {
        textareaRef.current.style.height = `${newHeight}px`;
      }
    }
  }, [maxHeight, fullHeight]);

  // Apply initial height and adjust on value change
  useEffect(() => {
    if (value !== textValue) {
      setTextValue(value || "");
    }

    // Use a small timeout to ensure content is rendered before measuring
    // eslint-disable-next-line no-undef
    const timer = setTimeout(() => {
      adjustHeight();
    }, 0);

    // eslint-disable-next-line no-undef
    return () => clearTimeout(timer);
  }, [value, adjustHeight, textValue]);

  // Handle debounced resize
  const debouncedAdjustHeight = useCallback(() => {
    // Skip resize if using full height mode
    if (fullHeight) return;

    // Clear previous timeout if exists
    if (timeoutRef.current) {
      clearTimeout(timeoutRef.current);
    }

    // Set new timeout
    timeoutRef.current = setTimeout(() => {
      adjustHeight();
    }, 150); // 150ms debounce
  }, [adjustHeight, fullHeight, timeoutRef]);

  // Handle input changes with immediate update and debounced resize
  const handleChange = useCallback(
    (e) => {
      const newValue = e.target.value;
      // Update state immediately for responsive typing
      setTextValue(newValue);

      // Debounce the height adjustment
      debouncedAdjustHeight();

      // Call onChange prop if provided
      if (onChange) {
        onChange(newValue);
      }
    },
    [debouncedAdjustHeight, onChange]
  );

  // Calculate style based on props
  const textareaStyle = {
    width: width ? (typeof width === "number" ? `${width}px` : width) : "100%",
    minHeight: height ? (typeof height === "number" ? `${height}px` : height) : "80px",
    maxHeight: maxHeight ? (typeof maxHeight === "number" ? `${maxHeight}px` : maxHeight) : undefined,
  };

  // Define CSS classes for the textarea
  const textareaClasses = ["auto-resize-textarea", fullHeight ? "full-height" : "", scrollable ? "scrollable" : ""]
    .filter(Boolean)
    .join(" ");

  return (
    <div className={`textarea-container ${className}`}>
      <textarea
        ref={textareaRef}
        className={textareaClasses}
        value={textValue}
        onChange={handleChange}
        placeholder={placeholder}
        style={textareaStyle}
      />
    </div>
  );
};

TextArea.propTypes = {
  value: PropTypes.string,
  onChange: PropTypes.func,
  placeholder: PropTypes.string,
  width: PropTypes.oneOfType([PropTypes.string, PropTypes.number]),
  height: PropTypes.oneOfType([PropTypes.string, PropTypes.number]),
  maxHeight: PropTypes.oneOfType([PropTypes.string, PropTypes.number]),
  className: PropTypes.string,
  fullHeight: PropTypes.bool,
  scrollable: PropTypes.bool,
};

TextArea.defaultProps = {
  value: "",
  placeholder: "",
  className: "",
  fullHeight: false,
  scrollable: true,
};

export default TextArea;
