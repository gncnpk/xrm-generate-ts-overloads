// ==UserScript==
// @name         Microsoft Power Platform/Dynamics 365 CE - Generate TypeScript Overload Signatures
// @namespace    https://github.com/gncnpk/xrm-generate-ts-overloads
// @author       Gavin Canon-Phratsachack (https://github.com/gncnpk)
// @version      1.7
// @license      GPL-3.0
// @description  Automatically creates TypeScript type definitions compatible with @types/xrm by extracting form attributes and controls from Dynamics 365/Power Platform model-driven applications.
// @match        https://*.dynamics.com/main.aspx?appid=*&pagetype=entityrecord&etn=*&id=*
// @grant        none
// ==/UserScript==
 
(function() {
    'use strict';
 
    // Create a button element and style it to be fixed in the bottom-right corner.
    const btn = document.createElement('button');
    btn.textContent = 'Generate TypeScript Signatures';
    btn.style.position = 'fixed';
    btn.style.bottom = '20px';
    btn.style.right = '20px';
    btn.style.padding = '10px';
    btn.style.backgroundColor = '#007ACC';
    btn.style.color = '#fff';
    btn.style.border = 'none';
    btn.style.borderRadius = '5px';
    btn.style.cursor = 'pointer';
    btn.style.zIndex = 10000;
    document.body.appendChild(btn);
 
    btn.addEventListener('click', () => {
        // Mapping objects for Xrm attribute and control types.
        var attributeTypeMapping = {
            "boolean": "Xrm.Attributes.BooleanAttribute",
            "datetime": "Xrm.Attributes.DateAttribute",
            "decimal": "Xrm.Attributes.NumberAttribute",
            "double": "Xrm.Attributes.NumberAttribute",
            "integer": "Xrm.Attributes.NumberAttribute",
            "lookup": "Xrm.Attributes.LookupAttribute",
            "memo": "Xrm.Attributes.StringAttribute",
            "money": "Xrm.Attributes.NumberAttribute",
            "multiselectoptionset": "Xrm.Attributes.MultiselectOptionSetAttribute",
            "optionset": "Xrm.Attributes.OptionSetAttribute",
            "string": "Xrm.Attributes.StringAttribute"
        };
 
        var controlTypeMapping = {
            "standard": "Xrm.Controls.StandardControl",
            "iframe": "Xrm.Controls.IframeControl",
            "lookup": "Xrm.Controls.LookupControl",
            "optionset": "Xrm.Controls.OptionSetControl",
            "customsubgrid:MscrmControls.Grid.GridControl": "Xrm.Controls.GridControl",
            "subgrid": "Xrm.Controls.GridControl",
            "timelinewall": "Xrm.Controls.TimelineWall",
            "quickform": "Xrm.Controls.QuickFormControl"
        };

        var specificControlTypeMapping = {
            "boolean": "Xrm.Controls.BooleanControl",
            "datetime": "Xrm.Controls.DateControl",
            "decimal": "Xrm.Controls.NumberControl",
            "double": "Xrm.Controls.NumberControl",
            "integer": "Xrm.Controls.NumberControl",
            "lookup": "Xrm.Controls.LookupControl",
            "memo": "Xrm.Controls.StringControl",
            "money": "Xrm.Controls.NumberControl",
            "multiselectoptionset": "Xrm.Controls.MultiselectOptionSetControl",
            "optionset": "Xrm.Controls.OptionSetControl",
            "string": "Xrm.Controls.StringControl"
        }
 
        // Object to hold the type information.
        const typeInfoTemplate = { attributes: {}, controls: {}, possibleTypes: {}, enums: {} }
        const entityTypeInfos = {};
 
        // Loop through all controls on the form.
        if (typeof Xrm !== 'undefined' && Xrm.Page && typeof Xrm.Page.getControl === 'function') {
            let entity = Xrm.Page.data.entity.getEntityName();
            let typeInfo = entityTypeInfos[entity] = entityTypeInfos[entity] ?? typeInfoTemplate;
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

        // Loop through all enums on the form.
        if (typeof Xrm.Page.getAttribute === 'function') {
            let entity = Xrm.Page.data.entity.getEntityName();
            let typeInfo = entityTypeInfos[entity] = entityTypeInfos[entity] ?? typeInfoTemplate;
            Xrm.Page.getAttribute().forEach((attr) => {
                if (attr.getAttributeType() === "optionset" && attr.controls.get().length > 0) {
                    const enumValues = attr.getOptions();
                    if (enumValues) {
                        typeInfo.enums[attr.controls.get(0).getLabel().replace(/\s+/g, '')] = enumValues;
                    }
                }
            });
        }

        // Loop through all Quick View controls on the form.
        if (typeof Xrm.Page.ui.quickForms.get === 'function') {
            let entity = Xrm.Page.data.entity.getEntityName();
            let typeInfo = entityTypeInfos[entity] = entityTypeInfos[entity] ?? typeInfoTemplate;
            Xrm.Page.ui.quickForms.get().forEach((ctrl) => {
                const ctrlType = ctrl.getControlType();
                const mappedType = controlTypeMapping[ctrlType];
                if (mappedType) {
                    typeInfo.possibleTypes[ctrl.getName()] = typeInfo.possibleTypes[ctrl.getName()] ?? [];
                    typeInfo.possibleTypes[ctrl.getName()].push(mappedType);
                }
            });
        }

        // Loop through all tabs and sections on the form.
        if (typeof Xrm.Page.ui.tabs.get === 'function') {
            let entity = Xrm.Page.data.entity.getEntityName();
            let typeInfo = entityTypeInfos[entity] = entityTypeInfos[entity] ?? typeInfoTemplate;
            Xrm.Page.ui.tabs.get().forEach((tab) => {
                typeInfo.possibleTypes[tab.getName()] = typeInfo.possibleTypes[tab.getName()] ?? [];
                typeInfo.possibleTypes[tab.getName()].push("Xrm.Controls.Tab");
                tab.sections.forEach((section) => {
                    typeInfo.possibleTypes[section.getName()] = typeInfo.possibleTypes[section.getName()] ?? [];
                    typeInfo.possibleTypes[section.getName()].push("Xrm.Controls.Section");   
                });
            });
        }

        // Loop through all attributes on the form.
        if (typeof Xrm.Page.getAttribute === 'function') {
            let entity = Xrm.Page.data.entity.getEntityName();
            let typeInfo = entityTypeInfos[entity] = entityTypeInfos[entity] ?? typeInfoTemplate;
            Xrm.Page.getAttribute().forEach((attr) => {
                const attrType = attr.getAttributeType();
                const mappedType = attributeTypeMapping[attrType];
                const mappedControlType = specificControlTypeMapping[attrType];
                if (mappedType) {
                    typeInfo.attributes[attr.getName()] = mappedType;
                    typeInfo.controls[attr.getName()] = mappedControlType;
                    typeInfo.possibleTypes[attr.getName()] = [];
                    typeInfo.possibleTypes[attr.getName()].push(mappedType);
                    typeInfo.possibleTypes[attr.getName()].push(mappedControlType);
                }
            });
        }
 
        // Build the TypeScript overload string.
        let outputTS = `// This file is generated automatically.
// It extends the Xrm.FormContext interface with overloads for getAttribute and getControl.
// Do not modify this file manually.
`
for(const [entityName, typeInfo] of Object.entries(entityTypeInfos)) {
    for(const [enumName, enumValues] of Object.entries(typeInfo.enums)) {
        let enumTemplate = [];
        for(const enumValue of enumValues) {
            enumTemplate.push(`   ${enumValue.text.replace(/\W/g, '')} = ${enumValue.value}`);
        }
outputTS += `
/**
 * Entity: ${entityName}
 */`
outputTS += `
const enum ${enumName} {
${enumTemplate.join(",\n")}
}
`
    }
}
outputTS += `
declare namespace Xrm {
    namespace Collection {
        interface ItemCollection<T> {
`
for(const [entityName, typeInfo] of Object.entries(entityTypeInfos)) {
    for (const [possibleTypeName, possibleTypesArray] of Object.entries(typeInfo.possibleTypes)) {
        let possibleTypeTemplate = "";
        for (const possibleType of possibleTypesArray) {
            possibleTypeTemplate += ` TSubType extends ${possibleType} ? ${possibleType} :`;
        }
        outputTS += `
            /**
             * Entity: ${entityName}
             */
            `
        outputTS += ` get<TSubType extends T>(itemName: "${possibleTypeName}"):${possibleTypeTemplate} never;\n`;
    }
}
outputTS += `
    }
}`
outputTS += `
  interface FormContext {
`;
for(const [entityName, typeInfo] of Object.entries(entityTypeInfos)) {
    for (const [attributeName, attributeType] of Object.entries(typeInfo.attributes)) {
        outputTS += `
        /**
         * Entity: ${entityName}
         */
        `
        outputTS += ` getAttribute(attributeName: "${attributeName}"): ${attributeType};\n`;
    }
    for (const [controlName, controlType] of Object.entries(typeInfo.controls)) {
        outputTS += `
        /**
         * Entity: ${entityName}
         */
        `
        outputTS += ` getControl(controlName: "${controlName}"): ${controlType};\n`;
    }
} 
        outputTS += `  }
}
`;
 
        // Create a new window with a textarea showing the output.
        // The textarea is set to readonly to prevent editing.
        const w = window.open('', '_blank', 'width=600,height=400,menubar=no,toolbar=no,location=no,resizable=yes');
        if (w) {
            w.document.write('<html><head><title>TypeScript Overload Signatures</title></head><body>');
            w.document.write('<textarea readonly style="width:100%; height:90%;">' + outputTS + '</textarea>');
            w.document.write('</body></html>');
            w.document.close();
        } else {
            // Fallback to prompt if popups are blocked.
            prompt("Copy the TypeScript definition:", outputTS);
        }
    });
})();
