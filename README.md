# xrm-generate-ts-overloads

Automatically creates TypeScript type definitions compatible with [@types/xrm](https://www.npmjs.com/package/@types/xrm) by extracting form attributes and controls from Dynamics 365/Power Platform model-driven applications.

You can find the source code repository here: [xrm-generate-ts-overloads](https://github.com/gncnpk/xrm-generate-ts-overloads)

## Prerequsites

* A userscript manager such as [Tampermonkey](https://www.tampermonkey.net/) (Chrome, Microsoft Edge, Safari, Opera Next, and Firefox) or [Greasemonkey](https://www.greasespot.net/) (Firefox)

## Usage

1. Download and install the user script.
2. Load any Dynamics 365 or Model-driven application form.
3. Click on the blue Generate TypeScript Definitions button in the bottom right corner.
4. Create a d.ts file in your project folder.
5. Copy the generated TypeScript defintiions to your d.ts file.

**Definitions for subgrids & quick views are generated only if the subgrid/quick view is loaded on the form.**

**Subgrids must have atleast one row of data for definitions to be generated.**

## Features

### Generated Form-Specific Definitions

#### Contexts
* [gridContext](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/clientapi-grid-context)
* [formContext](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/clientapi-form-context)
* [executionContext](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/clientapi-execution-context)

#### Method Parameter/Return Types
* [getAttribute()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes)
* [getControl()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/controls/getcontrol)
* [attributes.get()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes)
* [controls.get()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/controls)
* [sections.get()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/formcontext-ui-tab-sections)
* [tabs.get()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/formcontext-ui-tabs)
* [getOption()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes#choices-and-choice-column-types)
* [getOptions()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes#choices-and-choice-column-types)
* [getSelectedOption()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes#choices-and-choice-column-types)
* [getValue()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes)
* [getText()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes#choices-and-choice-column-types)
* [setValue()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes)

#### Enums
* OptionSet Options

#### Interfaces
* [OptionSet Attributes](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes#choices-and-choice-column-types)
* [OptionSet Values](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes#choices-and-choice-column-types)

#### Subgrid
* Attributes

#### Quick Forms/Views
* Attributes
* [Controls](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/formcontext-ui-quickforms)

#### Literal and Type Unions
* Form Attributes
* Form Controls
* Form Tabs
* Form Quick Views
* Subgrid Attributes
* Quick View Attributes
* Quick View Controls


