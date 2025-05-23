import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { OptionSetComponent, IOptionSetProps } from "./OptionSetControl";
import * as React from "react";

const SmallFormFactorMaxWidth = 350;
const enum FormFactors {
  Unknown = 0,
  Desktop = 1,
  Tablet = 2,
  Phone = 3,
}

export class OptionSetControl implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private selectedValue: number | undefined;
    private context: ComponentFramework.Context<IInputs>;
    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        'use strict';
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;
        this.context.mode.trackContainerResize(true);
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        'use strict';
        const {optionSetValue, configuration } = context.parameters;
        let disabled = context.mode.isControlDisabled;
        let masked = false;
        if(optionSetValue.security){
            disabled = disabled || !optionSetValue.security.editable;
            masked = !optionSetValue.security.readable;
        }
        const props: IOptionSetProps = { 
            label: optionSetValue.attributes?.DisplayName,
            value: optionSetValue.raw,
            options: optionSetValue.attributes?.Options,
            iconConfig: configuration.raw,
            onChange: this.onChange,
            disabled: disabled,
            masked: masked,
            fieldRequiredLevel: optionSetValue.attributes?.RequiredLevel,
            formFactor: context.client.getFormFactor() === FormFactors.Phone || context.mode.allocatedWidth < SmallFormFactorMaxWidth ? 'small' : 'large',
        };
        return React.createElement(
            OptionSetComponent, props
        );
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return { optionSetValue: this.selectedValue ?? 'null'} as IOutputs;
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }

    private onChange = (newValue: number | undefined) : void =>{
        'use strict';
        if(newValue === -1){
            this.selectedValue = undefined;
            this.context.parameters.optionSetValue.raw = null;
        } else{
            this.selectedValue = newValue;
        }
        this.notifyOutputChanged();
    }
}
