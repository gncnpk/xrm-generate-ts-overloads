// ==UserScript==
// @name         Microsoft Power Platform/Dynamics 365 CE - Generate TypeScript Definitions
// @namespace    https://github.com/gncnpk/xrm-generate-ts-overloads
// @author       Gavin Canon-Phratsachack (https://github.com/gncnpk)
// @version      1.942
// @license      GPL-3.0
// @description  Automatically creates TypeScript type definitions compatible with @types/xrm by extracting form attributes and controls from Dynamics 365/Power Platform model-driven applications.
// @match        https://*.dynamics.com/main.aspx?appid=*&pagetype=entityrecord&etn=*&id=*
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  const generateComments = function (fieldName, fieldValue) {
    return `\n/** ${fieldName}: ${fieldValue} */`;
  };
  // Create a button element and style it to be fixed in the bottom-right corner.
  const btn = document.createElement("button");
  btn.textContent = "Generate TypeScript Definitions";
  btn.style.position = "fixed";
  btn.style.bottom = "20px";
  btn.style.right = "20px";
  btn.style.padding = "10px";
  btn.style.backgroundColor = "#007ACC";
  btn.style.color = "#fff";
  btn.style.border = "none";
  btn.style.borderRadius = "5px";
  btn.style.cursor = "pointer";
  btn.style.zIndex = 10000;
  document.body.appendChild(btn);

  btn.addEventListener("click", () => {
    // Mapping objects for Xrm attribute and control types.
    var attributeTypeMapping = {
      boolean: "Xrm.Attributes.BooleanAttribute",
      datetime: "Xrm.Attributes.DateAttribute",
      decimal: "Xrm.Attributes.NumberAttribute",
      double: "Xrm.Attributes.NumberAttribute",
      integer: "Xrm.Attributes.NumberAttribute",
      lookup: "Xrm.Attributes.LookupAttribute",
      memo: "Xrm.Attributes.StringAttribute",
      money: "Xrm.Attributes.NumberAttribute",
      multiselectoptionset: "Xrm.Attributes.MultiselectOptionSetAttribute",
      optionset: "Xrm.Attributes.OptionSetAttribute",
      string: "Xrm.Attributes.StringAttribute",
    };

    var controlTypeMapping = {
      standard: "Xrm.Controls.StandardControl",
      iframe: "Xrm.Controls.IframeControl",
      lookup: "Xrm.Controls.LookupControl",
      optionset: "Xrm.Controls.OptionSetControl",
      "customsubgrid:MscrmControls.Grid.GridControl":
        "Xrm.Controls.GridControl",
      subgrid: "Xrm.Controls.GridControl",
      timelinewall: "Xrm.Controls.TimelineWall",
      quickform: "Xrm.Controls.QuickFormControl",
    };

    var specificControlTypeMapping = {
      boolean: "Xrm.Controls.BooleanControl",
      datetime: "Xrm.Controls.DateControl",
      decimal: "Xrm.Controls.NumberControl",
      double: "Xrm.Controls.NumberControl",
      integer: "Xrm.Controls.NumberControl",
      lookup: "Xrm.Controls.LookupControl",
      memo: "Xrm.Controls.StringControl",
      money: "Xrm.Controls.NumberControl",
      multiselectoptionset: "Xrm.Controls.MultiselectOptionSetControl",
      optionset: "Xrm.Controls.OptionSetControl",
      string: "Xrm.Controls.StringControl",
    };

    // Object to hold the type information.
    const typeInfo = {
      attributes: {},
      controls: {},
      possibleTypes: {},
      enums: {},
      quickViewControls: {},
    };

    // Loop through all controls on the form.
    if (
      typeof Xrm !== "undefined" &&
      Xrm.Page &&
      typeof Xrm.Page.getControl === "function"
    ) {
      Xrm.Page.getControl().forEach((ctrl) => {
        const ctrlType = ctrl.getControlType();
        const mappedType = controlTypeMapping[ctrlType];
        if (mappedType) {
          typeInfo.controls[ctrl.getName()] = mappedType;
          typeInfo.possibleTypes[ctrl.getName()] = [];
          typeInfo.possibleTypes[ctrl.getName()].push(mappedType);
        }
      });
    } else {
      alert("Xrm.Page is not available on this page.");
      return;
    }

    // Loop through all Quick View controls on the form.
    if (typeof Xrm.Page.ui.quickForms.get === "function") {
      Xrm.Page.ui.quickForms.get().forEach((ctrl) => {
        const ctrlType = ctrl.getControlType();
        const mappedType = controlTypeMapping[ctrlType];
        if (mappedType) {
          typeInfo.possibleTypes[ctrl.getName()] =
            typeInfo.possibleTypes[ctrl.getName()] ?? [];
          typeInfo.possibleTypes[ctrl.getName()].push(mappedType);
        }
        ctrl.getControl().forEach((subCtrl) => {
          if (typeof subCtrl.getAttribute !== "function") {
            return;
          }
          const subCtrlAttrType = subCtrl.getAttribute().getAttributeType();
          const mappedControlType =
            specificControlTypeMapping[subCtrlAttrType] ??
            controlTypeMapping[subCtrl.getControlType()];
          if (mappedControlType) {
            typeInfo.quickViewControls[subCtrl.getName()] = mappedControlType;
          }
        });
      });
    }

    // Loop through all tabs and sections on the form.
    if (typeof Xrm.Page.ui.tabs.get === "function") {
      Xrm.Page.ui.tabs.get().forEach((tab) => {
        typeInfo.possibleTypes[tab.getName()] =
          typeInfo.possibleTypes[tab.getName()] ?? [];
        typeInfo.possibleTypes[tab.getName()].push("Xrm.Controls.Tab");
        tab.sections.forEach((section) => {
          typeInfo.possibleTypes[section.getName()] =
            typeInfo.possibleTypes[section.getName()] ?? [];
          typeInfo.possibleTypes[section.getName()].push(
            "Xrm.Controls.Section"
          );
        });
      });
    }

    // Loop through all attributes on the form.
    if (typeof Xrm.Page.getAttribute === "function") {
      Xrm.Page.getAttribute().forEach((attr) => {
        const attrType = attr.getAttributeType();
        const attrName = attr.getName();
        const mappedType = attributeTypeMapping[attrType];
        const mappedControlType = specificControlTypeMapping[attrType];
        if (mappedType) {
          typeInfo.attributes[attrName] = mappedType;
          typeInfo.controls[attrName] = mappedControlType;
          typeInfo.possibleTypes[attrName] = [];
          if (attrType !== "optionset") {
            typeInfo.possibleTypes[attrName].push(mappedType);
            typeInfo.possibleTypes[attrName].push(mappedControlType);
          }
        }
        if (
          attr.getAttributeType() === "optionset" &&
          attr.controls.get().length > 0
        ) {
          const enumValues = attr.getOptions();
          if (enumValues) {
            typeInfo.enums[attrName] = { attribute: "", values: [] };
            typeInfo.enums[attrName].values = enumValues;
            typeInfo.enums[attrName].attribute = attrName;
            typeInfo.attributes[attrName] = `${attrName}_attribute`;
            typeInfo.possibleTypes[attrName].push(`${attrName}_attribute`);
            typeInfo.possibleTypes[attrName].push(mappedControlType);
          }
        }
      });
    }

    // Loop through all subgrids on the form.
    if (typeof Xrm.Page.getControl === "function") {
      Xrm.Page.getControl().forEach((ctrl) => {
        if (
          ctrl.getControlType() === "subgrid" ||
          ctrl.getControlType() ===
            "customsubgrid:MscrmControls.Grid.GridControl"
        ) {
          const gridRow = ctrl.getGrid().getRows().get(0);
          if (gridRow !== null) {
            gridRow.data.entity.attributes.forEach((attr) => {
              const attrType = attr.getAttributeType();
              const attrName = attr.getName();
              const mappedType = attributeTypeMapping[attrType];
              const mappedControlType = specificControlTypeMapping[attrType];
              if (mappedType) {
                typeInfo.attributes[attrName] = mappedType;
                typeInfo.possibleTypes[attrName] = [];
                if (attrType !== "optionset") {
                  typeInfo.possibleTypes[attrName].push(mappedType);
                }
              }
              if (
                attr.getAttributeType() === "optionset" &&
                attr.controls.get().length > 0
              ) {
                const enumValues = attr.getOptions();
                if (enumValues) {
                  typeInfo.enums[attrName] = { attribute: "", values: [] };
                  typeInfo.enums[attrName].values = enumValues;
                  typeInfo.enums[attrName].attribute = attrName;
                  typeInfo.attributes[attrName] = `${attrName}_attribute`;
                  typeInfo.possibleTypes[attrName].push(
                    `${attrName}_attribute`
                  );
                }
              }
            });
          }
        }
      });
    }

    // Build the TypeScript overload string.
    let outputTS = `// This file is generated automatically.
// It extends the Xrm.FormContext interface with overloads for getAttribute and getControl.
// Do not modify this file manually.

`;
    let extendedAttributeTypes = [];
    for (const [originalEnumName, enumValues] of Object.entries(
      typeInfo.enums
    )) {
      let enumName = `${originalEnumName}_enum`;
      let enumTemplate = [];
      let textLiteralTypes = [];
      let valueLiteralTypes = [];
      for (const enumValue of enumValues.values) {
        enumTemplate.push(
          `   ${enumValue.text.replace(/\W/g, "").replace(/[0-9]/g, "")} = ${
            enumValue.value
          }`
        );
        textLiteralTypes.push(`"${enumValue.text}"`);
        valueLiteralTypes.push(
          `${enumName}.${enumValue.text
            .replace(/\W/g, "")
            .replace(/[0-9]/g, "")}`
        );
      }
      outputTS += `
const enum ${enumName} {
${enumTemplate.join(",\n")}
}
`;
      outputTS += `
interface ${enumValues.attribute}_value extends Xrm.OptionSetValue {
text: ${textLiteralTypes.join(" | ")};
value: ${valueLiteralTypes.join(" | ")};  
}
`;
      outputTS += `
interface ${enumValues.attribute}_attribute extends Xrm.Attributes.OptionSetAttribute {
  getOptions(): ${enumValues.attribute}_value[];
  getSelectedOption(): ${enumValues.attribute}_value | null;
  getValue(): ${enumName} | null;
  setValue(value: ${enumName} | null): void;
}
`;
      extendedAttributeTypes.push(`${enumValues.attribute}_attribute`);
    }
    outputTS += `

type extendedAttributeTypes = Xrm.Attributes.SpecificAttributeTypes | ${extendedAttributeTypes.join(
      " | "
    )}

declare namespace Xrm {

interface Entity {
    attributes: Collection.ItemCollection<extendedAttributeTypes>;
  }

namespace Controls {
    interface QuickFormControl {`;

    for (const [quickViewControlName, quickViewControlType] of Object.entries(
      typeInfo.quickViewControls
    )) {
      outputTS += `\ngetControl(controlName: "${quickViewControlName}"): ${quickViewControlType};\n`;
    }

    outputTS += `}
  }
  `;
    outputTS += `    namespace Collection {
        interface ItemCollection<T> {
`;
    for (const [possibleTypeName, possibleTypesArray] of Object.entries(
      typeInfo.possibleTypes
    )) {
      let possibleTypeTemplate = "";
      for (const possibleType of possibleTypesArray) {
        possibleTypeTemplate += ` TSubType extends ${possibleType} ? ${possibleType} :`;
      }
      outputTS += `\nget<TSubType extends T>(itemName: "${possibleTypeName}"):${possibleTypeTemplate} never;\n`;
    }
    outputTS += `
    }
}`;
    outputTS += `
  interface FormContext {
`;
    for (const [attributeName, attributeType] of Object.entries(
      typeInfo.attributes
    )) {
      outputTS += `\ngetAttribute(attributeName: "${attributeName}"): ${attributeType};\n`;
    }
    for (const [controlName, controlType] of Object.entries(
      typeInfo.controls
    )) {
      outputTS += `\ngetControl(controlName: "${controlName}"): ${controlType};\n`;
    }
    outputTS += `  }
}
`;

    // Create a new window with a textarea showing the output.
    // The textarea is set to readonly to prevent editing.
    const w = window.open(
      "",
      "_blank",
      "width=600,height=400,menubar=no,toolbar=no,location=no,resizable=yes"
    );
    if (w) {
      w.document.write(
        "<html><head><title>TypeScript Definitions</title></head><body>"
      );
      w.document.write(
        '<textarea readonly style="width:100%; height:90%;">' +
          outputTS +
          "</textarea>"
      );
      w.document.write("</body></html>");
      w.document.close();
    } else {
      // Fallback to prompt if popups are blocked.
      prompt("Copy the TypeScript definition:", outputTS);
    }
  });
})();
