import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { sortDataSet, IDataSetProps } from "./SortDataSet";
import * as React from "react";


export class OrderLineItems implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private sequenceColumnTarget: any;
    private optionSetColorDefined?: any;
    private optionSetColorDefinedMetadata?: any;
    private entityType?: any;


    /**
     * Empty constructor.
     */
    constructor() { }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    public async init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ):Promise<void>{
        "use strict";
        this.notifyOutputChanged = notifyOutputChanged;
        this.entityType = context.parameters.DataSet.getTargetEntityType();
        this.sequenceColumnTarget = context.parameters.DataSet.columns.find((column)=>column.alias ==="OrderColumn");
        this.optionSetColorDefined = context.parameters.DataSet.columns.find((column)=>column.alias ==="OptionSetColorDefined");
        
        // Method for set all varaibles that handled in Async operation, and call the updateView method
        this.SetVariableValues();
    }
    
    async SetVariableValues(): Promise<void> {
        "use strict";
        this.optionSetColorDefinedMetadata = await this.GetAttributeMetadata(this.entityType,this.optionSetColorDefined.name);
        this.getOutputs();
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        "use strict";
        const props: IDataSetProps ={
                                     dataSet: context.parameters.DataSet,
                                     sequenceColumn: this.sequenceColumnTarget,
                                     optionSetMetadata: this.optionSetColorDefinedMetadata,
                                     optionSetColumn: this.optionSetColorDefined
                                    }
        return React.createElement(
            sortDataSet, props
        );
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        "use strict";
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        "use strict";
        // Add code to cleanup control if necessary
    }

    public GetAttributeMetadata(entityLogicName: string, attributeLoficalName: string)  {
        "use strict";
        return new Promise(function (resolve, reject){
            if(entityLogicName === null || attributeLoficalName === null) {
                resolve(null);
            }else{
                const OptionMetaData: [any] = [null];
                const Utility = (Xrm as any).Utility;
                const globalContext = Utility.getGlobalContext();
                const serverURL = globalContext.getClientUrl();
        
                const Query = "EntityDefinitions(LogicalName='"+entityLogicName +"')/Attributes(LogicalName='"+attributeLoficalName+"')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)";
                const req = new XMLHttpRequest();
                req.open("GET", serverURL + "/api/data/v9.2/" + Query, true);
                req.setRequestHeader("Accept", "application/json");
                req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                req.setRequestHeader("OData-MaxVersion", "4.0");
                req.setRequestHeader("OData-Version", "4.0");
                req.setRequestHeader("Prefer", "odata.include-annotations=*");
                req.onreadystatechange =  async()=>{
                    "use strict";
                    if (req.readyState === 4 /* complete */) {
                        req.onreadystatechange = null;
                        if (req.status === 200) {
                            const data = JSON.parse(req.response);
                            if (data !== null && data.OptionSet !== null && data.OptionSet !== undefined) {
                                data.OptionSet.Options.forEach((option:any) => {
                                    OptionMetaData.push( {
                                        value: option.Value,
                                        label: option.Label.UserLocalizedLabel.Label,
                                        color: option.Color
                                    });
                                });
                                OptionMetaData.shift();
                                resolve(OptionMetaData);
                            }
                        }
                    }
                };
                req.send();
            }
        });
    }
}


