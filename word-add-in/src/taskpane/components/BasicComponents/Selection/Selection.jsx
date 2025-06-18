import React, { useState, useRef, useEffect } from "react";
import PropTypes from "prop-types";
import "./Selection.scss";

const Selection = ({ list, selected, onChange, placeholder }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [selectedItem, setSelectedItem] = useState(selected);
  const dropdownRef = useRef(null);

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };

    // eslint-disable-next-line no-undef
    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      // eslint-disable-next-line no-undef
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  // Update local state when prop changes
  useEffect(() => {
    setSelectedItem(selected);
  }, [selected]);

  const handleSelect = (item) => {
    setSelectedItem(item);
    setIsOpen(false);
    if (onChange) {
      onChange(item);
    }
  };

  const toggleDropdown = () => {
    setIsOpen(!isOpen);
  };

  // Find selected item's label
  const getSelectedLabel = () => {
    if (!selectedItem && placeholder) {
      return placeholder;
    }

    const item = list.find((item) => item.value === selectedItem);
    return item ? item.label : placeholder || "Select an option";
  };

  // Check if showing placeholder or actual value
  const isShowingPlaceholder = () => {
    if (!selectedItem) return true;
    const item = list.find((item) => item.value === selectedItem);
    return !item;
  };

  return (
    <div className="selection-container" ref={dropdownRef}>
      <div className={`selection-header ${isOpen ? "active" : ""}`} onClick={toggleDropdown}>
        <div className={`selection-value ${isShowingPlaceholder() ? "" : "selected-value"}`}>{getSelectedLabel()}</div>
        <div className="selection-arrow">{isOpen ? "▲" : "▼"}</div>
      </div>

      {isOpen && (
        <div className="selection-dropdown">
          {list.length > 0 ? (
            <>
              <div className="selection-dropdown-header">{placeholder}</div>
              {list.map((item) => (
                <div
                  key={item.value}
                  className={`selection-item ${item.value === selectedItem ? "selected" : ""}`}
                  onClick={() => handleSelect(item.value)}
                >
                  {item.label}
                </div>
              ))}
            </>
          ) : (
            <div className="selection-no-options">No options available</div>
          )}
        </div>
      )}
    </div>
  );
};

Selection.propTypes = {
  list: PropTypes.arrayOf(
    PropTypes.shape({
      label: PropTypes.string.isRequired,
      value: PropTypes.oneOfType([PropTypes.string, PropTypes.number]).isRequired,
    })
  ).isRequired,
  selected: PropTypes.oneOfType([PropTypes.string, PropTypes.number]),
  onChange: PropTypes.func,
  placeholder: PropTypes.string,
};

Selection.defaultProps = {
  list: [],
  placeholder: "Select an option",
};

export default Selection;
