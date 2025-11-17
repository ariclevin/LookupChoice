/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    lookupControl: ComponentFramework.PropertyTypes.LookupProperty;
    sortByName: ComponentFramework.PropertyTypes.EnumProperty<"1" | "0">;
    mruSize: ComponentFramework.PropertyTypes.WholeNumberProperty;
    attributemask: ComponentFramework.PropertyTypes.StringProperty;
}
export interface IOutputs {
    lookupControl?: ComponentFramework.LookupValue[];
}
