import React from "react";
import PropTypes from "prop-types";
import "./Button.scss";

const Button = ({ label, onClick, disabled = false, width, height, className = "" }) => {
  const buttonStyle = {
    width: width ? (typeof width === "number" ? `${width}px` : width) : undefined,
    height: height ? (typeof height === "number" ? `${height}px` : height) : undefined,
  };

  return (
    <button
      className={`custom-button ${disabled ? "disabled" : ""} ${className}`}
      onClick={disabled ? undefined : onClick}
      disabled={disabled}
      style={buttonStyle}
    >
      {label}
    </button>
  );
};

Button.propTypes = {
  label: PropTypes.string.isRequired,
  onClick: PropTypes.func.isRequired,
  disabled: PropTypes.bool,
  width: PropTypes.oneOfType([PropTypes.string, PropTypes.number]),
  height: PropTypes.oneOfType([PropTypes.string, PropTypes.number]),
  className: PropTypes.string,
};

Button.defaultProps = {
  disabled: false,
  className: "",
  width: undefined,
  height: undefined,
};

export default Button;
