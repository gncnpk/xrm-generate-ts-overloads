// ==UserScript==
// @name         Microsoft Power Platform/Dynamics 365 CE - Generate TypeScript Definitions
// @namespace    https://github.com/gncnpk/xrm-generate-ts-overloads
// @author       Gavin Canon-Phratsachack (https://github.com/gncnpk)
// @version      1.984
// @license      GPL-3.0
// @description  Automatically creates TypeScript type definitions compatible with @types/xrm by extracting form attributes and controls from Dynamics 365/Power Platform model-driven applications.
// @match        https://*.dynamics.com/main.aspx?appid=*&pagetype=entityrecord&etn=*&id=*
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  const groupItemsByType = (items) => {
    return Object.entries(items).reduce((acc, [itemName, itemType]) => {
        if (!acc[itemType]) {
            acc[itemType] = [];
        }
        acc[itemType].push(itemName);
        return acc;
    }, {});
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
      subGrids: {},
      quickViews: {},
      formAttributes: {},
      formControls: {},
      possibleEnums: [],
      formTabs: {},
      formEnums: {},
      quickViewControls: {},
      quickViewAttributes: {}
    };
    class Subgrid {
      constructor() {
        this.attributes = {};
        this.enums = {};
      }
    }
    class QuickForm {
      constructor() {
        this.attributes = {};
        this.controls = {};
        this.enums = {};
      }
    }

    class Tab {
      constructor() {
        this.sections = {};
      }
    }

    const currentFormName = Xrm.Page.ui.formSelector
      .getCurrentItem()
      .getLabel();

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
          typeInfo.formControls[ctrl.getName()] = mappedType;
        }
      });
    } else {
      alert("Xrm.Page is not available on this page.");
      return;
    }

    // Loop through all tabs and sections on the form.
    if (typeof Xrm.Page.ui.tabs.get === "function") {
      Xrm.Page.ui.tabs.get().forEach((tab) => {
        let formTab = (typeInfo.formTabs[tab.getName()] = new Tab());
        tab.sections.forEach((section) => {
          formTab.sections[section.getName()] = `${section.getName()}_section`;
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
          typeInfo.formAttributes[attrName] = mappedType;
          attr.controls.forEach((ctrl) => {
            typeInfo.formControls[ctrl.getName()] = mappedControlType;
          });
        }
        if (
          attr.getAttributeType() === "optionset" &&
          attr.controls.get().length > 0
        ) {
          const enumValues = attr.getOptions();
          if (enumValues) {
            typeInfo.formEnums[attrName] = { attribute: "", values: [] };
            typeInfo.formEnums[attrName].values = enumValues;
            typeInfo.formEnums[attrName].attribute = attrName;
            typeInfo.formAttributes[attrName] = `${attrName}_attribute`;
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
          const gridName = ctrl.getName();
          let subgrid = (typeInfo.subGrids[gridName] = new Subgrid());
          typeInfo.formControls[gridName] = `${gridName}_gridcontrol`;
          if (gridRow !== null) {
            gridRow.data.entity.attributes.forEach((attr) => {
              const attrType = attr.getAttributeType();
              const attrName = attr.getName();
              const mappedType = attributeTypeMapping[attrType];
              if (mappedType) {
                subgrid.attributes[attrName] = mappedType;
              }
              if (
                attr.getAttributeType() === "optionset" &&
                attr.controls.get().length > 0
              ) {
                const enumValues = attr.getOptions();
                if (enumValues) {
                  subgrid.enums[attrName] = { attribute: "", values: [] };
                  subgrid.enums[attrName].values = enumValues;
                  subgrid.enums[attrName].attribute = attrName;
                  subgrid.attributes[attrName] = `${attrName}_attribute`;
                }
              }
            });
          }
        }
      });
    }

    // Loop through all Quick View controls and attributes on the form.
    if (typeof Xrm.Page.ui.quickForms.get === "function") {
      Xrm.Page.ui.quickForms.get().forEach((ctrl) => {
        const quickViewName = ctrl.getName();
        let quickView = (typeInfo.quickViews[quickViewName] = new QuickForm());
        const ctrlType = ctrl.getControlType();
        typeInfo.formControls[
          quickViewName
        ] = `${quickViewName}_quickformcontrol`;
        ctrl.getControl().forEach((subCtrl) => {
          if (typeof subCtrl.getAttribute !== "function") {
            return;
          }
          const subCtrlAttrType = subCtrl.getAttribute().getAttributeType();
          const mappedControlType =
            specificControlTypeMapping[subCtrlAttrType] ??
            controlTypeMapping[subCtrl.getControlType()];
          if (mappedControlType) {
            quickView.controls[subCtrl.getName()] = mappedControlType;
          }
        });
        ctrl.getAttribute().forEach((attr) => {
          const attrType = attr.getAttributeType();
          const attrName = attr.getName();
          const mappedType = attributeTypeMapping[attrType];
          if (mappedType) {
            quickView.attributes[attrName] = mappedType;
          }
          if (attrType === "optionset" && attr.controls.get().length > 0) {
            const enumValues = attr.getOptions();
            if (enumValues) {
              quickView.enums[attrName] = { attribute: "", values: [] };
              quickView.enums[attrName].values = enumValues;
              quickView.enums[attrName].attribute = attrName;
              quickView.attributes[attrName] = `${attrName}_attribute`;
            }
          }
        });
      });
    }

    // Build the TypeScript overload string.
    let outputTS = `// These TypeScript definitions were generated automatically on: ${new Date().toDateString()}\n`;
    for (const [originalEnumName, enumValues] of Object.entries(
      typeInfo.formEnums
    )) {
      if (typeInfo.possibleEnums.includes(originalEnumName)) {
        continue;
      }
      typeInfo.possibleEnums.push(originalEnumName);
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
interface ${enumValues.attribute}_attribute extends Xrm.Attributes.OptionSetAttribute<${enumName}> {
   `
      valueLiteralTypes.forEach((value, index) => { 
        outputTS+= `getOption(value: ${value}): {text: ${textLiteralTypes[index]}, value: ${value}};\n`;
      })
  outputTS+=`getOption(value: ${enumValues.attribute}_value['value']): ${enumValues.attribute}_value | null;
  getOptions(): ${enumValues.attribute}_value[];
  getSelectedOption(): ${enumValues.attribute}_value | null;
  getValue(): ${enumName} | null;
  setValue(value: ${enumName} | null): void;
  getText(): ${enumValues.attribute}_value['text'] | null;
}
`;
    }
    for (const [subgridName, subgrid] of Object.entries(typeInfo.subGrids)) {
      for (const [enumName, enumValues] of Object.entries(subgrid.enums)) {
        if (typeInfo.possibleEnums.includes(enumName)) {
          continue;
        }
        typeInfo.possibleEnums.push(enumName);
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
            `${enumName}_enum.${enumValue.text
              .replace(/\W/g, "")
              .replace(/[0-9]/g, "")}`
          );
        }
        outputTS += `
const enum ${enumName}_enum {
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
interface ${enumValues.attribute}_attribute extends Xrm.Attributes.OptionSetAttribute<${enumName}_enum> {
`
      valueLiteralTypes.forEach((value, index) => { 
        outputTS+= `getOption(value: ${value}): {text: ${textLiteralTypes[index]}, value: ${value}};\n`;
      })
  outputTS+=`
  getOption(value: ${enumValues.attribute}_value['value']): ${enumValues.attribute}_value | null;
  getOptions(): ${enumValues.attribute}_value[];
  getSelectedOption(): ${enumValues.attribute}_value | null;
  getValue(): ${enumName}_enum | null;
  setValue(value: ${enumName}_enum | null): void;
  getText(): ${enumValues.attribute}_value['text'] | null;
  }
  `;
      }

      outputTS += `
  interface ${subgridName}_attributes extends Xrm.Collection.ItemCollection<${new Set(
        Object.values(subgrid.attributes)
      )
        .map((attrType) => `${attrType}`)
        .join(" | ")}> {`;
        for (const [attrType, attrNames] of Object.entries(
          groupItemsByType(subgrid.attributes)
        )) {
          outputTS += `get(itemName: "${attrNames.join("\" | \"")}"): ${attrType};\n`;
        }
      outputTS += `}
  `;

      outputTS += `

  interface ${subgridName}_entity extends Xrm.Entity {
    attributes: ${subgridName}_attributes;
  }
  interface ${subgridName}_data extends Xrm.Data {
    entity: ${subgridName}_entity;
  }
  
  interface ${subgridName}_gridrow extends Xrm.Controls.Grid.GridRow {
    data: ${subgridName}_data;
  }
  
  interface ${subgridName}_grid extends Xrm.Controls.Grid {
    getRows(): Xrm.Collection.ItemCollection<${subgridName}_gridrow>;
  }
  
  interface ${subgridName}_gridcontrol extends Xrm.Controls.GridControl {
    getGrid(): ${subgridName}_grid;
  }
  interface ${subgridName}_context extends Xrm.FormContext {`;
  for (const [attrType, attrNames] of Object.entries(
    groupItemsByType(subgrid.attributes)
  )) {
    outputTS += `getAttribute(attributeName: "${attrNames.join("\" | \"")}"): ${attrType};\n`;
  }
    outputTS += `getAttribute(attributeNameOrIndex: string | number): ${subgridName}_attributes | null;`
      outputTS += `getAttribute(delegateFunction?): ${subgridName}_attributes[];`;
      outputTS += `
  }
  `;
      outputTS += `
  interface ${subgridName}_eventcontext extends Xrm.Events.EventContext {
    getFormContext(): ${subgridName}_context;
}`;
    }
    for (const [quickViewName, quickView] of Object.entries(
      typeInfo.quickViews
    )) {
      for (const [enumName, enumValues] of Object.entries(quickView.enums)) {
        if (typeInfo.possibleEnums.includes(enumName)) {
          continue;
        }
        typeInfo.possibleEnums.push(enumName);
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
            `${enumName}_enum.${enumValue.text
              .replace(/\W/g, "")
              .replace(/[0-9]/g, "")}`
          );
        }
        outputTS += `
const enum ${enumName}_enum {
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
interface ${enumValues.attribute}_attribute extends Xrm.Attributes.OptionSetAttribute<${enumName}_enum> {
`
      valueLiteralTypes.forEach((value, index) => { 
        outputTS+= `getOption(value: ${value}): {text: ${textLiteralTypes[index]}, value: ${value}};\n`;
      })
  outputTS+=`
  getOption(value: ${enumValues.attribute}_value['value']): ${enumValues.attribute}_value | null;
  getOptions(): ${enumValues.attribute}_value[];
  getSelectedOption(): ${enumValues.attribute}_value | null;
  getValue(): ${enumName}_enum | null;
  setValue(value: ${enumName}_enum | null): void;
  getText(): ${enumValues.attribute}_value['text'] | null;
      }
  `;
      }
      outputTS += `type ${quickViewName}_controls_types = ${new Set(
        Object.values(quickView.controls)
      )
        .map((controlType) => `${controlType}`)
        .join(" | ")}\n`;
      outputTS += `
  type ${quickViewName}_attributes_types = ${new Set(
        Object.values(quickView.attributes)
      )
        .map((attrType) => `${attrType}`)
        .join(" | ")}\n`;
        outputTS += `type ${quickViewName}_controls_literals = ${new Set(
          Object.keys(quickView.controls)
        )
          .map((controlName) => `"${controlName}"`)
          .join(" | ")}\n`;
        outputTS += `
    type ${quickViewName}_attributes_literals = ${new Set(
          Object.keys(quickView.attributes)
        )
          .map((attrName) => `"${attrName}"`)
          .join(" | ")}\n`;
  
      outputTS += `
  interface ${quickViewName}_quickformcontrol extends Xrm.Controls.QuickFormControl {
    `;
      for (const [controlType, controlNames] of Object.entries(
        groupItemsByType(quickView.controls)
      )) {
        outputTS += `getControl(controlName: "${controlNames.join("\" | \"")}"): ${controlType};\n`;
      }
      outputTS += `
    getControl(): ${quickViewName}_controls_types[]
    getControl(controlNameOrIndex: string | number): ${quickViewName}_controls_types | null;
    `;
    for (const [attrType, attrNames] of Object.entries(
      groupItemsByType(quickView.attributes)
    )) {
      outputTS += `getAttribute(attributeName: "${attrNames.join("\" | \"")}"): ${attrType};\n`;
    }
      outputTS += `
      getAttribute(attributeNameOrIndex: string | number): ${quickViewName}_attributes_types | null;
    getAttribute(): ${quickViewName}_attributes_types[];
  
  }`;
    }
    for (const [tabName, tab] of Object.entries(typeInfo.formTabs)) {
      outputTS += `
interface ${tabName}_sections extends Xrm.Collection.ItemCollection<Xrm.Controls.Section> {`;
        outputTS += `get(itemName:"${Object.keys(tab.sections).join("\" | \"")}"): Xrm.Controls.Section;\n`;
      
      outputTS += `}`;

      outputTS += `
interface ${tabName}_tab extends Xrm.Controls.Tab {
  sections: ${tabName}_sections;
    }`;
    }
    outputTS += `
interface ${currentFormName}_tabs extends Xrm.Collection.ItemCollection<${new Set(
      Object.keys(typeInfo.formTabs)
    )
      .map((tabName) => `${tabName}_tab`)
      .join(" | ")}> {`;
    for (const [itemName, itemType] of Object.entries(typeInfo.formTabs)) {
      outputTS += `get(itemName:"${itemName}"): ${itemName}_tab;\n`;
    }
    outputTS += `}`;
    outputTS += `
interface ${currentFormName}_quickforms extends Xrm.Collection.ItemCollection<${new Set(
      Object.keys(typeInfo.quickViews)
    )
      .map((quickViewName) => `${quickViewName}_quickformcontrol`)
      .join(" | ")}> {`;
    for (const [itemName, itemType] of Object.entries(typeInfo.quickViews)) {
      outputTS += `get(itemName:"${itemName}"): ${itemName}_quickformcontrol;\n`;
    }
    outputTS += `}`;
    outputTS += `
    type ${currentFormName}_attributes_types = ${new Set(
      Object.values(typeInfo.formAttributes)
    )
      .map((attrType) => `${attrType}`)
      .join(" | ")}\n`;
    outputTS += `
    type ${currentFormName}_attributes_literals = ${new Set(
      Object.keys(typeInfo.formAttributes)
    )
      .map((attrName) => `"${attrName}"`)
      .join(" | ")}\n`;
      
    outputTS += `
    type ${currentFormName}_controls_types = ${new Set(
      Object.values(typeInfo.formControls)
    )
      .map((ctrlType) => `${ctrlType}`)
      .join(" | ")}\n`
    outputTS += `
    type ${currentFormName}_controls_literals = ${new Set(
      Object.keys(typeInfo.formControls)
    )
      .map((ctrlName) => `"${ctrlName}"`)
      .join(" | ")}\n`;
    outputTS += `
    interface ${currentFormName}_ui extends Xrm.Ui {
      quickForms: ${currentFormName}_quickforms | null;
      tabs: ${currentFormName}_tabs;

    }
    `;
    outputTS += `
    interface ${currentFormName}_context extends Xrm.FormContext {
    ui: ${currentFormName}_ui;
  `;
  for (const [controlType, controlNames] of Object.entries(
    groupItemsByType(typeInfo.formControls)
  )) {
    outputTS += `getControl(controlName: "${controlNames.join("\" | \"")}"): ${controlType};\n`;
  }
    outputTS += `
    getControl(controlNameOrIndex: string | number): ${currentFormName}_controls_types | null;
    getControl(delegateFunction?): ${currentFormName}_controls_types[];
    `
    for (const [attrType, attrNames] of Object.entries(
      groupItemsByType(typeInfo.formAttributes)
    )) {
      outputTS += `getAttribute(attributeName: "${attrNames.join("\" | \"")}"): ${attrType};\n`;
    }
    outputTS += `
    getAttribute(attributeNameOrIndex: string | number): ${currentFormName}_attributes_types | null;
    getAttribute(delegateFunction?): ${currentFormName}_attributes_types[];
  }`;
    outputTS += `
  interface ${currentFormName}_eventcontext extends Xrm.Events.EventContext {
    getFormContext(): ${currentFormName}_context;
}`;

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
