# xrm-generate-ts-overloads

Automatically creates TypeScript type definitions compatible with [@types/xrm](https://www.npmjs.com/package/@types/xrm) by extracting form attributes and controls from Dynamics 365/Power Platform model-driven applications.

You can find the source code here: [xrm-generate-ts-overloads](https://github.com/gncnpk/xrm-generate-ts-overloads)

## Features

### Generated Form-Specific Definitions

#### Method Return Types
* [getAttribute()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes)
* [getControl()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/controls/getcontrol)
* [data.attributes.get()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes)
* [controls.get()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/controls)
* [sections.get()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/formcontext-ui-tab-sections)
* [ui.tabs.get()](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/formcontext-ui-tabs)

#### Enums
* OptionSet Options

#### Interfaces
* [OptionSet Attributes](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes#choices-and-choice-column-types)
* [OptionSet Values](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/attributes#choices-and-choice-column-types)

#### Subgrid
* Attributes

#### Quick Forms/Views
* Controls


