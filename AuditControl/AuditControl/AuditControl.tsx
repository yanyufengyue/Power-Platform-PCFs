import * as React from 'react';
import { DetailsList, Link, PrimaryButton, SelectionMode, Selection, Stack, Modal, StackItem, IconButton, Panel, Checkbox} from '@fluentui/react';
import { IInputs } from './generated/ManifestTypes';
import * as Helper from './helper';
import { useBoolean } from '@fluentui/react-hooks';
import { Suspense } from 'react';

export interface IAuditControlProps {
    context: ComponentFramework.Context<IInputs>;
}

export const auditHistory = React.memo(({context}: IAuditControlProps): JSX.Element =>{
    const [items, setItems] = React.useState<any>([]);
    const [columns, setColumns] = React.useState<any>([]); 
    const [auditDataSet, setAuditDataSet] = React.useState<any>();
    const [auditMetadata, setAuditMetadata] = React.useState<any>();
    const [currentEntityAttributes, setCurrentEntityAttributes] = React.useState<any>();
    const [sortedEntityAttributes, setSortedEntityAttributes] = React.useState<any>();
    const [entityOptionSetMetadata, setEntityOptionSetMetadata] = React.useState<any[]>([]);
    const [manyToOneRelationship, setManyToOneRelationship] = React.useState<any[]>([]);
    const [selectedRecords, setSelectedRecords] = React.useState<any>([]);
    const pageContext: any = Xrm.Utility.getPageContext();
    const appUrl = Xrm.Utility.getGlobalContext().getCurrentAppUrl();
    const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl(); 
    const [isSettingModalOpen, {setTrue: showSettingModal, setFalse: hideSettingModal}] = useBoolean(false);
    const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(false);
    const [checkBoxes, setCheckBoxes] = React.useState<JSX.Element>();
    const [attributes, setAttributes] = React.useState<any>([]);
    const [selectedAttributes, setSelectedAttributes] = React.useState<any>([]);
    // Loading all required data
    React.useEffect(()=>{

        const fetchData = {
            "objectid": pageContext.input.entityId,
            "operation": "2"
          };
        let fetchXml = [
            "<fetch>",
            "  <entity name='audit'>",
            "    <filter>",
            "      <condition attribute='objectid' operator='eq' value='", fetchData.objectid/*00000000-0000-0000-0000-000000000000*/, "'/>",
            "      <condition attribute='operation' operator='eq' value='", fetchData.operation/*2*/, "'/>",
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
        const attrRefUrl = apiUrl + "/ManyToOneRelationships";
        fetch(attrRefUrl, {
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
                setManyToOneRelationship([...data.value]);
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
                    context.navigation.openErrorDialog({details:error.message, message:"Error in retrieve entity attributes"});
                });
                const multiselectOptionsetMetadataUrl = apiUrl + "/Attributes/Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";
                fetch(multiselectOptionsetMetadataUrl, {
                    method: "GET",
                    headers: {
                        "Accept": "application/json",
                        "OData-MaxVersion": "4.0",
                        "OData-Version": "4.0",
                        "Content-Type": "application/json"
                    }
                })
                .then(response => response.json())
                .then(optionsetMetaData => {
                    if (optionsetMetaData.value && optionsetMetaData.value.length > 0) {
                        setEntityOptionSetMetadata(prevState=>[...prevState, ...optionsetMetaData.value]);
                    } 
                })
                .catch(error => {
                    context.navigation.openErrorDialog({details:"Error in retrieve MultiSelectPicklistAttributeMetadata", message:error.message});
                });

                const optionsetMetadataUrl = apiUrl + "/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";
                fetch(optionsetMetadataUrl, {
                    method: "GET",
                    headers: {
                        "Accept": "application/json",
                        "OData-MaxVersion": "4.0",
                        "OData-Version": "4.0",
                        "Content-Type": "application/json"
                    }
                })
                .then(response => response.json())
                .then(optionsetMetaData => {
                    if (optionsetMetaData.value && optionsetMetaData.value.length > 0) {
                        setEntityOptionSetMetadata(prevState=>[...prevState, ...optionsetMetaData.value]);
                    } 
                })
                .catch(error => {
                    context.navigation.openErrorDialog({details:error.message, message:"Error in retrieve PicklistAttributeMetadata"});
                });
            }
        })
        .catch(error => {
            context.navigation.openErrorDialog({details:error.message, message:"Error in retrieve entityMetadata"});
        });
        
    }, [context])
    React.useEffect(()=>{
        if(currentEntityAttributes){
            const sortedAtrributes:any = [];
            currentEntityAttributes.value.forEach((attribute: any) => {
                if(attribute.AttributeOf === null) sortedAtrributes.push(attribute);
            });
            setSortedEntityAttributes(sortedAtrributes);
        }
    },[currentEntityAttributes])
    // Set columns and items
    React.useEffect(()=>{
        "use strict";
        if(auditMetadata && auditDataSet && entityOptionSetMetadata && currentEntityAttributes){
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
    }, [auditMetadata,auditDataSet, entityOptionSetMetadata, currentEntityAttributes])

    React.useEffect(()=>{
        "use strict";
        if(selectedRecords.length === 1){
            const data = JSON.parse(selectedRecords[0].changedata).changedAttributes;
            const myAttributes = data.map((e:any)=>{
              if(sortedEntityAttributes.length > 0){
                  for (const element of sortedEntityAttributes) {
                      if(element.LogicalName ===e.logicalName){
                            return(            
                                element.DisplayName.UserLocalizedLabel.Label
                            )
                        }
                    }
                }
                else{
                    return(            
                        e.logicalName
                    )
                }
            });

            const mySelectedAttributes:any = [];
            data.forEach((e:any)=>{
                if(sortedEntityAttributes.length > 0){
                    for (const element of sortedEntityAttributes) {
                        if(element.LogicalName ===e.logicalName && element.AttributeTypeName?.Value !== "FileType" && element.AttributeTypeName?.Value !=="UniqueidentifierType"){
                            mySelectedAttributes.push(element.DisplayName.UserLocalizedLabel.Label);
                        }
                    }
                }
                else{
                    mySelectedAttributes.push(e.logicalName)
                }
              });

            setAttributes([...myAttributes]);
            setSelectedAttributes([...mySelectedAttributes]);
        }
    },[selectedRecords])

    React.useEffect(()=>{
        if(selectedRecords.length > 0 && attributes && selectedAttributes){
            const data = JSON.parse(selectedRecords[0]?.changedata)?.changedAttributes;
            const myCheckBoxes = data.map((e:any)=>{
                if(sortedEntityAttributes.length > 0){
                    for (const element of sortedEntityAttributes) {
                        if(element.LogicalName ===e.logicalName){
                              return(            
                                  <Stack.Item key={element.LogicalName}>
                                      <Checkbox label={element.DisplayName.UserLocalizedLabel.Label} ariaLabel={element.DisplayName.UserLocalizedLabel.Label} name={element.LogicalName} checked={selectedAttributes.some((col:any) => col === element.DisplayName.UserLocalizedLabel.Label)} disabled ={element.AttributeTypeName?.Value === "FileType" || element.AttributeTypeName?.Value ==="UniqueidentifierType"} onChange={_onChange}/>
                                  </Stack.Item>
                                )
                            }
                        }
                    }
                else{
                    return(            
                        <Stack.Item key={e.logicalName}>
                            <Checkbox label={e.logicalName} ariaLabel={e.logicalName} name={e.logicalName} checked={selectedAttributes.some((col:any) => col === e.logicalName)} onChange={_onChange}/>
                        </Stack.Item>
                    )
                }
              });
            setCheckBoxes(myCheckBoxes)
        }
    },[attributes, selectedAttributes])
    
    function _onChange(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean) {
        let updatedSelectedAttributes: any = [...selectedAttributes];
        const attributeName = (ev?.target as HTMLInputElement).ariaLabel; 
        if(isChecked){
        // If isChecked is true, add the column to the grid
          if (!updatedSelectedAttributes.some((attr:any) => attr === attributeName)) {
              updatedSelectedAttributes.push(attributeName);            
          }
        } else {
          // If isChecked is false, remove the column from the grid
          updatedSelectedAttributes = updatedSelectedAttributes.filter((attr: any) => attr !== attributeName);
        }
        setSelectedAttributes(updatedSelectedAttributes);
    } 
    const renderItemColumn = (item?: any, index?: number, column?: any) =>{
        "use strict";
        let componentUrl, link;
        switch(column?.dataType){
            case 'lookup':
                componentUrl = "&pagetype=entityrecord&etn=" + item.raw["_"+column.fieldName+'_value@Microsoft.Dynamics.CRM.lookuplogicalname'] + "&id="+item.raw["_"+column.fieldName+'_value'];
                link = appUrl + componentUrl;
                return<Link href={link}>{item.raw["_"+column.fieldName+'_value@OData.Community.Display.V1.FormattedValue']}</Link>
            case 'Lookup.Owner':
            case 'Lookup.Simple':
                componentUrl = "&pagetype=entityrecord&etn=" + item.raw._record.fields[column.fieldName].reference.etn + "&id="+item.raw._record.fields[column.fieldName].reference.id.guid;
                link = appUrl + componentUrl;
                return<Link href={link}>{item[column.fieldName]}</Link>
            case 'picklist':
            case 'datetime':
                return<>{item.raw[column.fieldName+'@OData.Community.Display.V1.FormattedValue']}</>
            default:
                if(column.fieldName == 'changedata'){
                    return (
                        <Helper.ChangeDataCell
                            key={`${item.id}-${column.fieldName}`}
                            item={item}
                            column={column}
                            currentEntityAttributes={currentEntityAttributes}
                            entityOptionSetMetadata={entityOptionSetMetadata}
                        />
                    );
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
                const selectedAttrs = sortedEntityAttributes.filter(
                    (entityAttribute:any) => selectedAttributes.some((att:any)=>att === entityAttribute?.DisplayName?.UserLocalizedLabel?.Label && (entityAttribute.AttributeTypeName?.Value !== "UniqueidentifierType" && entityAttribute.AttributeTypeName?.Value !== "FileType"))
                );
                const changedAttributes = JSON.parse(selectedRecords[0].changedata).changedAttributes;
                const myData = changedAttributes.map((e:any)=>{
                    for (const seltAttr of selectedAttrs) {
                        if(seltAttr.LogicalName === e.logicalName){
                            if(e.oldValue){
                                if(seltAttr.AttributeTypeName?.Value === "LookupType" || seltAttr.AttributeTypeName?.Value === "CustomerType"){      
                                    const attributeReferencingMetadata = manyToOneRelationship.filter((attrRef:any) =>attrRef.ReferencingAttribute === seltAttr.LogicalName && attrRef.ReferencedEntity === e.oldValue.split(',')[0]);                          
                                    return{[attributeReferencingMetadata[0].ReferencingEntityNavigationPropertyName + "@odata.bind"]: '/'+e.oldValue.split(',')[0] +"s("+ e.oldValue.split(',')[1]+")"};                                    
                                }else if(seltAttr.AttributeTypeName?.Value === "OwnerType"){
                                    const attributeReferencingMetadata = manyToOneRelationship.filter((attrRef:any) =>attrRef.ReferencingAttribute === seltAttr.LogicalName);                          
                                    return{[attributeReferencingMetadata[0].ReferencingEntityNavigationPropertyName + "@odata.bind"]: '/'+e.oldValue.split(',')[0] +"s("+ e.oldValue.split(',')[1]+")"};         
                                }                              
                                else if(seltAttr.AttributeTypeName?.Value === "MoneyType"){
                                    return{[seltAttr.LogicalName]:parseFloat(e.oldValue)}
                                }  
                                else if(seltAttr.AttributeType === "Integer"){
                                    return{[seltAttr.LogicalName]:parseInt(e.oldValue)}
                                }
                                else if(seltAttr.AttributeTypeName?.Value === "DoubleType"){
                                    return{[seltAttr.LogicalName]:parseFloat(e.oldValue)}
                                }                                
                                else if(seltAttr.AttributeTypeName?.Value === "DecimalType"){
                                    return{[seltAttr.LogicalName]:parseFloat(e.oldValue)}
                                }                                
                                else if(seltAttr.AttributeTypeName?.Value === "MultiSelectPicklistType"){
                                    return{[seltAttr.LogicalName]:JSON.parse(e.oldValue)?.toString()}
                                }
                                else{
                                    return {[seltAttr.LogicalName] : e.oldValue};
                                }
                            }else{
                                if(seltAttr.AttributeTypeName?.Value === "LookupType" || seltAttr.AttributeTypeName?.Value === "CustomerType"){                                
                                    const attributeReferencingMetadata = manyToOneRelationship.filter((attrRef:any) =>attrRef.ReferencingAttribute === seltAttr.LogicalName && attrRef.ReferencedEntity === e.newValue.split(',')[0]);
                                    return{[attributeReferencingMetadata[0].ReferencingEntityNavigationPropertyName]: null};                                    
                                } else if(seltAttr.AttributeTypeName?.Value === "OwnerType"){
                                    const attributeReferencingMetadata = manyToOneRelationship.filter((attrRef:any) =>attrRef.ReferencingAttribute === seltAttr.LogicalName);
                                    return{[attributeReferencingMetadata[0].ReferencingEntityNavigationPropertyName]: null}; 
                                }
                                else{
                                    return {[seltAttr.LogicalName] : null};
                                }    
                            }
                        }
                    }                   
                });
                const revertData = Object.assign({},...myData)
                context.webAPI.updateRecord(pageContext.input.entityName,pageContext.input.entityId,revertData).then(
                    function succes(result){
                        hideSettingModal();
                    },
                    function error (er:any){
                        hideSettingModal();
                        context.navigation.openErrorDialog({details:er.message, message:"Error in reverting changes"});
                    }
                );             
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
                <Modal
                titleAriaId='controlSettingModal'
                isOpen={isSettingModalOpen}
                containerClassName={Helper.getModalContentStyles().container}>         
                    <div className={Helper.getModalContentStyles().header}>
                        <h2 className={Helper.getModalContentStyles().heading}>
                        Review data to update
                        </h2>
                        <IconButton
                        styles={Helper.getIconStyle()}
                        iconProps={{iconName:'Cancel'}}
                        ariaLabel="Close popup modal"
                        title='Close the Settings'
                        onClick={hideSettingModal}
                        />
                    </div>
                    <div className={Helper.getModalContentStyles().body}>
                        <p>By default, all Supported Columns (Not supported columns will be disabled from selection) in the selected row will be reverted. If you wish to proceed, click on Confirm button. Otherwise, Please click on the Modify button to view items to be reverted, and remove the item from the selected list.</p>
                    </div>
                    <div className={Helper.getModalContentStyles().body}>                       
                        <PrimaryButton id="cancel" text="Cancel" onClick={hideSettingModal} disabled={selectedRecords.length<=0} style={{"padding":"5px"}}/>
                        <PrimaryButton id="modify" text="Modify" onClick={openPanel} style={{"padding":"1px"}}/>
                        <PrimaryButton id="revert" text="Confirm" onClick={onClickButton("revert")} disabled={selectedAttributes.length<=0} style={{"padding":"5px", float: 'right'}}/>
                    </div>
                </Modal>
                <Panel
                headerText="Select/Remove"
                // this prop makes the panel non-modal
                isBlocking={true}
                isOpen={isOpen}
                onDismiss={dismissPanel}
                closeButtonAriaLabel="Close"
                >
                    <Stack tokens={{childrenGap: 10}}>
                    {checkBoxes}
                    </Stack>
                </Panel>
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
                        <PrimaryButton id="openModal" text="Revert Change" onClick={showSettingModal} disabled={selectedAttributes.length<=0} style={{"padding":"1px"}}/>
                        </div>
                    </Stack.Item>
                </Stack>
            </>)
        }    
    </>)
});
auditHistory.displayName = auditHistory.displayName ?? "auditHistoryControl";