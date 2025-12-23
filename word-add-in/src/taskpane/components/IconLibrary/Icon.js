/*

 Copyright (c) 2016-2023, AnyGen AI Inc.  All rights reserved.


 IMPORTANT - PLEASE READ THIS CAREFULLY BEFORE ATTEMPTING TO USE ANY SOFTWARE,

 DOCUMENTATION, OR OTHER MATERIALS.

 This software is the confidential and proprietary information of AnyGen AI Inc

 ("Confidential Information") and is protected by applicable copyright or other

 intellectual property laws and treaties. All title and ownership rights in and

 to the software (including but not limited to any source code, images,

 photographs, animations, video, audio, music, text embedded in the software),

 the intellectual property embodied in the software, and any trademarks or

 service marks of AnyGen AI Inc. that are used in connection with the

 software, are and shall at all times remain exclusively owned by AnyGen AI,

 Inc. and its licensors.  Under no circumstances shall you disclose such

 Confidential Information and trade secrets, distribute, disclose or otherwise

 make available to any third party any portion of the software's source code

 and shall use it only in accordance with the terms of the license agreement

 enclosed with this product or as entered into with AnyGen AI, Inc.


 You are prohibited from any attempt to disassemble the code, or attempt in

 any manner to reconstruct, discover, reuse or modify any source code or

 underlying algorithms of the software.


 THIS SOFTWARE IS PROVIDED "AS IS" AND THERE ARE NO WARRANTIES, CLAIMS OR

 REPRESENTATIONS MADE BY AnyGen AI, INC., OR ITS LICENSORS, SUBSIDIARIES

 AND AFFILIATES, EITHER EXPRESS, IMPLIED, OR STATUTORY, INCLUDING WARRANTIES

 OF QUALITY, PERFORMANCE, NON-INFRINGEMENT, MERCHANTABILITY, OR FITNESS FOR

 A PARTICULAR PURPOSE, NOR ARE THERE ANY WARRANTIES CREATED BY COURSE OF

 DEALING, COURSE OF PERFORMANCE, OR TRADE USAGE. AnyGen AI, INC. DOES NOT

 WARRANT THAT THIS SOFTWARE WILL MEET ANY CLIENT'S NEEDS OR BE FREE FROM

 ERRORS, OR THAT THE OPERATION OF THE SOFTWARE WILL BE UNINTERRUPTED.
*/
import React from "react";
import Icons from "./iconLibrary";
import PropTypes from "prop-types";
import { isValidData } from "../BasicComponents/CommonMethods";

const defaultStyles = { display: "inline-block", verticalAlign: "middle" };
const defaultDisabledStyles = { display: "inline-block", fill: "#cdcdcd !important", verticalAlign: "middle" };

const Icon = ({
  size,
  title,
  tooltipSize,
  disabled = false,
  stroke,
  fillRule,
  icon,
  path,
  className,
  style,
  viewBox,
  onClick,
  onMouseOver,
}) => {
  let styles = { ...defaultStyles, ...style };
  if (disabled) {
    styles = { ...defaultDisabledStyles, ...style };
  }
  if (onMouseOver !== undefined && onMouseOver !== null) {
    return (
      <svg
        className={className}
        style={styles}
        viewBox={viewBox}
        width={`${size}px`}
        height={`${size}px`}
        xmlns="https://www.w3.org/2000/svg"
        // smlnsXlink="https://www.w3.org/1999/xlink"
        onMouseOver={onMouseOver}
        onClick={onClick}
        // disabled={disabled}
      >
        {fillRule !== "" ? (
          <path stroke={stroke} fillRule={fillRule} d={isValidData(path) ? path : Icons[icon]} />
        ) : (
          <path stroke={stroke} d={isValidData(path) ? path : Icons[icon]} />
        )}
      </svg>
    );
  } else {
    return (
      <svg
        className={className}
        style={styles}
        viewBox={viewBox}
        width={`${size}px`}
        height={`${size}px`}
        xmlns="https://www.w3.org/2000/svg"
        // smlnsXlink="https://www.w3.org/1999/xlink"
        onClick={onClick}
        // disabled={disabled}
      >
        {fillRule !== "" ? (
          <path stroke={stroke} fillRule={fillRule} d={isValidData(path) ? path : Icons[icon]} />
        ) : (
          <path stroke={stroke} d={isValidData(path) ? path : Icons[icon]} />
        )}
      </svg>
    );
  }
};
Icon.defaultProps = {
  size: 16,
  fill: "#ffffff",
  stroke: "",
  viewBox: "0 0 24 24",
  style: {},
  className: "",
  title: "",
  fillRule: "",
  tooltipSize: 300,
};
Icon.propTypes = {
  size: PropTypes.number.isRequired,
  tooltipSize: PropTypes.number,
  fill: PropTypes.string.isRequired,
  icon: PropTypes.string.isRequired,
  title: PropTypes.string,
  fillRule: PropTypes.oneOf(["evenodd", "inherit", "nonezero"]),
  viewBox: PropTypes.string,
  style: PropTypes.object,
  className: PropTypes.string,
  onClick: PropTypes.func,
};
export default Icon;
