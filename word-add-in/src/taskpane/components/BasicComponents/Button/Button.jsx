import React from "react";
import PropTypes from "prop-types";
import Icon from "../../IconLibrary/Icon";
import "./Button.scss";

const getMenuClass = (icon, disabled) => {
  let negativeIconList = ["delete", "remove", "close"];
  return disabled
    ? "menu-button disabled"
    : negativeIconList.includes(icon ? icon.toLowerCase() : "")
      ? "menu-button delete-btn"
      : "menu-button";
};

const Button = ({
  label,
  onClick,
  disabled = false,
  width,
  height,
  className = "",
  type = "primary",
  icon = "",
  iconSize,
  children,
}) => {
  const buttonStyle = {
    width: width ? (typeof width === "number" ? `${width}px` : width) : undefined,
    height: height ? (typeof height === "number" ? `${height}px` : height) : undefined,
  };

  let buttonClassName = "";

  switch (type) {
    case "primary":
      buttonClassName = "main-cta-btn";
      break;
    case "secondary":
      buttonClassName = "main-secondary-btn";
      break;
    case "icon":
      buttonClassName = getMenuClass(icon, disabled);
      break;
    default:
      buttonClassName = "main-secondary-btn";
  }

  if (className) {
    buttonClassName = `${buttonClassName} ${className}`;
  }

  if (type === "icon") {
    return (
      <button
        className={buttonClassName}
        onClick={disabled ? undefined : onClick}
        disabled={disabled}
        style={buttonStyle}
      >
        {icon && <Icon icon={icon} size={icon === "close" ? 8 : iconSize ? iconSize : 14} />}
        {children}
      </button>
    );
  }

  return (
    <button
      className={buttonClassName}
      onClick={disabled ? undefined : onClick}
      disabled={disabled}
      style={buttonStyle}
    >
      {label || children}
    </button>
  );
};

Button.propTypes = {
  label: PropTypes.string,
  onClick: PropTypes.func,
  disabled: PropTypes.bool,
  width: PropTypes.oneOfType([PropTypes.string, PropTypes.number]),
  height: PropTypes.oneOfType([PropTypes.string, PropTypes.number]),
  className: PropTypes.string,
  type: PropTypes.oneOf(["primary", "secondary", "icon"]),
  icon: PropTypes.string,
  iconSize: PropTypes.number,
  children: PropTypes.node,
};

Button.defaultProps = {
  disabled: false,
  className: "",
  width: undefined,
  height: undefined,
  type: "primary",
  icon: "",
};

export default Button;
