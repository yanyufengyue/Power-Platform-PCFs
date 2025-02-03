import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { gridDataset, IGridDatasetProps } from "./GridDataset";
import * as React from "react";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;

export class FancyGrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private optionsetMetadata?: any;
    private targetEntityType?: any;

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
    ): Promise<void> {
        this.notifyOutputChanged = notifyOutputChanged;
        this.targetEntityType = context.parameters.DataSet.getTargetEntityType();
        
        // Method for set all varaibles that handling in Async operation
        this.SetVariableValues();
    }

    async SetVariableValues(): Promise<void> {
        "use strict";
        this.optionsetMetadata = await getOptionSetMetadata(this.targetEntityType);
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        const props: IGridDatasetProps = { 
            dataset: context.parameters.DataSet,
            optionsetMetadata: this.optionsetMetadata,
            initialPageNumber: context.parameters.DataSet.paging.firstPageNumber
        }   
        return React.createElement(gridDataset, props);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}

function getOptionSetMetadata(targetEntityType: any): any {
    "use strict"
    return new Promise(function (resolve, reject){
        const Utility = (Xrm as any).Utility;
        const globalContext = Utility.getGlobalContext();
        const serverURL = globalContext.getClientUrl();
    
        const Query = "EntityDefinitions(LogicalName='"+targetEntityType +"')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";
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
                    const data =JSON.parse(req.response);
                    resolve(data.value);             
                }
            }
        };
        req.send();
    });
}

