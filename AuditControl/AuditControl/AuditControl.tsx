import * as React from 'react';
import { DetailsList, Link, PrimaryButton, SelectionMode, Selection, Stack} from '@fluentui/react';
import { IInputs } from './generated/ManifestTypes';
import * as Helper from './helper';

export interface IAuditControlProps {
    context: ComponentFramework.Context<IInputs>;
}

export const auditHistory = React.memo(({context}: IAuditControlProps): JSX.Element =>{
    const [items, setItems] = React.useState<any>([]);
    const [columns, setColumns] = React.useState<any>([]); 
    const [auditDataSet, setAuditDataSet] = React.useState<any>();
    const [auditMetadata, setAuditMetadata] = React.useState<any>();
    const [currentEntityAttributes, setCurrentEntityAttributes] = React.useState<any>();
    const [selectedRecords, setSelectedRecords] = React.useState<any>([]);
    const pageContext: any = Xrm.Utility.getPageContext();
    const appUrl = Xrm.Utility.getGlobalContext().getCurrentAppUrl();
    const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl(); 
    // Loading all required data
    React.useEffect(()=>{
        const fetchData = {
            "objectid": pageContext.input.entityId
          };
        let fetchXml = [
            "<fetch>",
            "  <entity name='audit'>",
            "    <filter>",
            "      <condition attribute='objectid' operator='eq' value='", fetchData.objectid/*00000000-0000-0000-0000-000000000000*/, "'/>",
            "    </filter>",
            "    <order attribute='createdon' descending='true'/>",
            "  </entity>",
            "</fetch>"
            ].join("");
        fetchXml = "?fetchXml=" + encodeURIComponent(fetchXml);
        context.webAPI.retrieveMultipleRecords("audit", fetchXml).then((response)=>{
            setAuditDataSet(response.entities);
        });

        context.utils.getEntityMetadata("audit", ["changedata","operation","userid", "createdon"]).then((response)=>{setAuditMetadata(response)});

        //Xrm.Utility.getEntityMetadata(pageContext.input.entityName).then(function successCallback (response){setCurrentEntityAttributes(response)});
        const apiUrl = clientUrl + "/api/data/v9.0/EntityDefinitions(LogicalName='" + pageContext.input.entityName + "')";
        fetch(apiUrl, {
            method: "GET",
            headers: {
                "Accept": "application/json",
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0",
                "Content-Type": "application/json"
            }
        })
        .then(response => response.json())  // Parse the JSON response
        .then(data => {
            if (data) {
                const attributesUrl = apiUrl + "/Attributes";
                fetch(attributesUrl, {
                    method: "GET",
                    headers: {
                        "Accept": "application/json",
                        "OData-MaxVersion": "4.0",
                        "OData-Version": "4.0",
                        "Content-Type": "application/json"
                    }
                })
                    .then(response => response.json())
                    .then(attributesData => {
                        if (attributesData.value && attributesData.value.length > 0) {
                            setCurrentEntityAttributes(attributesData);
                        } 
                    })
                    .catch(error => {
                        context.navigation.openErrorDialog({details:"Error in retrieve attributes", message:error});
                    });
            }
        })
        .catch(error => {
            context.navigation.openErrorDialog({details:"Error in retrieve entityMetadata", message:error});
        });
        
    }, [context])

    // Set columns and items
    React.useEffect(()=>{
        "use strict";
        if(auditMetadata && auditDataSet && currentEntityAttributes){
            setColumns((auditMetadata as any).Attributes.getAll().map((attribute: any)=>{
                return{
                    name: attribute.attributeDescriptor.DisplayName,
                    fieldName: attribute.attributeDescriptor.LogicalName,
                    Key: attribute.attributeDescriptor.LogicalName,
                    isResizable: true,
                    dataType: attribute.attributeDescriptor.Type === "memo" ? "string" : attribute.attributeDescriptor.Type
                }
            }));

            const myItems = auditDataSet.map((record: any)=>{
                const attributes = (auditMetadata as any).Attributes.getAll().map((attribute: any)=>{
                    const logicalName = attribute.attributeDescriptor.LogicalName;
                    return{[logicalName] : record[logicalName]}
                });
                return Object.assign({
                    Key: record.auditid,
                    raw: record
                },...attributes)
            });
            setItems(myItems);         
        }
    }, [auditMetadata,auditDataSet, currentEntityAttributes])

    const renderItemColumn = (item?: any, index?: number, column?: any) =>{
        "use strict";
        let componentUrl, link
        switch(column?.dataType){
            case 'lookup':
                componentUrl = "&pagetype=entityrecord&etn=" + item.raw["_"+column.fieldName+'_value@Microsoft.Dynamics.CRM.lookuplogicalname'] + "&id="+item.raw["_"+column.fieldName+'_value'];
                link = appUrl + componentUrl;
                return<Link href={link}>{item.raw["_"+column.fieldName+'_value@OData.Community.Display.V1.FormattedValue']}</Link>
            case 'Lookup.Owner':
                componentUrl = "&pagetype=entityrecord&etn=" + item.raw._record.fields[column.fieldName].reference.etn + "&id="+item.raw._record.fields[column.fieldName].reference.id.guid;
                link = appUrl + componentUrl;
                return<Link href={link}>{item[column.fieldName]}</Link>
            case 'Lookup.Simple':
                componentUrl = "&pagetype=entityrecord&etn=" + item.raw._record.fields[column.fieldName].reference.etn + "&id="+item.raw._record.fields[column.fieldName].reference.id.guid;
                link = appUrl + componentUrl;
                return<Link href={link}>{item[column.fieldName]}</Link>
            case 'picklist':
            case 'datetime':
                return<>{item.raw[column.fieldName+'@OData.Community.Display.V1.FormattedValue']}</>
            default:
                if(column.fieldName == 'changedata' && currentEntityAttributes && currentEntityAttributes.value.length > 0){
                    return (Helper.handleChangeData(currentEntityAttributes, item[column.fieldName]))
                }else{
                    return <>{item[column.fieldName]}</>;
                }
        }       
    }

    const selection= new Selection({
        onSelectionChanged: () =>{
          "use strict";
          const selectedItems = selection.getSelection();
          setSelectedRecords(selectedItems);
        }
    })

    const onClickButton = (buttonId: string)=>{
        "use strict";
        switch (buttonId){
          case "revert":
            return () =>{
              "use strict";
              if(selectedRecords.length === 1){
                const sortedAtrributes:any = [];
                currentEntityAttributes.value.forEach((attribute: any) => {
                    if(attribute.AttributeOf === null) sortedAtrributes.push(attribute);
                });
                const data = JSON.parse(selectedRecords[0].changedata).changedAttributes;
                const myData = data.map((e:any)=>{
                    if(sortedAtrributes.length > 0){
                        for (const element of sortedAtrributes) {
                            if(element.LogicalName ===e.logicalName){
                                if(element.AttributeType === "Lookup"){
                                    return{[e.oldValue.split(',')[0]+'@odata.bind']: '/'+element.Targets[0] +"s("+ e.oldValue.split(',')[1]+")"};
                                }else{
                                    return {[e.logicalName] : e.oldValue};
                                }
                            }
                        }
                    }
                    else{
                        return {[e.logicalName] : e.oldValue}
                    }
                });
                const revertData = Object.assign({},...myData)
                context.webAPI.updateRecord(pageContext.input.entityName,pageContext.input.entityId,revertData);
              }
            };
            break;
          default:
            return ()=>{"use strict";};
        }
    }

    return(<>
        {(!auditMetadata || !auditDataSet || auditDataSet?.length === 0) && <>No valid Audit History available</>}
        {auditMetadata && auditDataSet && auditDataSet.length > 0 &&
            (<>
                <Stack style={{width: '100%'}}>
                    <Stack.Item>
                    <div style={{width: '100%'}}>
                        <DetailsList
                            items={items}
                            columns={columns}
                            selectionPreservedOnEmptyClick={true}
                            onRenderItemColumn={renderItemColumn}
                            selectionMode={SelectionMode.single}
                            selection={selection}>          
                        </DetailsList>
                        </div>
                    </Stack.Item>
                    <Stack.Item>
                        <div>
                        <PrimaryButton id="revert" text="Revert Change" onClick={onClickButton("revert")} disabled={selectedRecords.length<=0} style={{"padding":"1px"}}/>
                        </div>
                    </Stack.Item>
                </Stack>
            </>)
        }    
    </>)
});
auditHistory.displayName = auditHistory.displayName ?? "auditHistoryControl";