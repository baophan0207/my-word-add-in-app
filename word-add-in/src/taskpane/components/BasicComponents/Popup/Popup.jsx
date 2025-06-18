import React from "react";
import PropTypes from "prop-types";
import "./Popup.scss";

const Popup = ({ logo, name, onClose, children, open }) => {
  if (!open) return null;

  return (
    <div className="popup-overlay">
      <div className="popup-container">
        <div className="popup-header">
          <div className="popup-header-left">
            {logo && <img src={logo} alt="Logo" className="popup-logo" />}
            {name && <h2 className="popup-name">{name}</h2>}
          </div>
          <button className="popup-close-button" onClick={onClose} aria-label="Close">
            x
          </button>
        </div>
        <div className="popup-content">{children}</div>
      </div>
    </div>
  );
};

Popup.propTypes = {
  logo: PropTypes.string,
  name: PropTypes.string,
  onClose: PropTypes.func.isRequired,
  children: PropTypes.node,
  open: PropTypes.bool.isRequired,
};

Popup.defaultProps = {
  open: false,
};

export default Popup;
