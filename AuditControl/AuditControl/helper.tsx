import * as React from "react";
import { IInputs } from "./generated/ManifestTypes";
import { createTheme, FontWeights, getTheme, Link, mergeStyleSets } from "@fluentui/react";

const optionSetMetadata:any = [];
const entityPrimaryField = new Map<string, any>();
const appUrl = Xrm.Utility.getGlobalContext().getCurrentAppUrl();
const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl(); 

export const handleChangeDataAsync = async (currentEntityAttributes: any,JSONdata: string, entityOptionsetMetadata?: any) => {
  const sortedAttributes = currentEntityAttributes.value?.filter((attribute: any) => attribute.AttributeOf === null) ?? [];
  const data = JSON.parse(JSONdata).changedAttributes;
  const tableData: any[] = [];
  for(const element of sortedAttributes){
    for(const e of data) {
        if(element.LogicalName ===e.logicalName){
            // Create rows with data for the table
            const oldValue = await formattingData(element, e.oldValue, entityOptionsetMetadata);
            const newValue = await formattingData(element, e.newValue, entityOptionsetMetadata);
            if(element.AttributeTypeName?.Value === "FileType"){
              tableData.push(
                {fieldName:element.DisplayName.UserLocalizedLabel.Label,
                  oldValue:oldValue,
                  newValue:newValue
                }
            );
            }else{
              tableData.push(
                  {fieldName:element.DisplayName.UserLocalizedLabel.Label,
                    oldValue:oldValue,
                    newValue:newValue
                  }
              );
            }

        }
      }
  }
  
  return(<div>
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead>
          <tr style={{background: 'lightblue'}}>
            <th style={{width:'20%'}}>Changed Column</th>
            <th style={{width:'40%'}}>Old Value</th>
            <th style={{width:'40%'}}>New Value</th>
          </tr>
        </thead>
        <tbody>
          {tableData.map((row: any, index: any) => (
            <tr key={index}>
              <td>{row.fieldName}</td>
              <td>{row.oldValue}</td>
              <td>{row.newValue}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export async function formattingData(attributeMetadata: any, attributeValue: any, entityOptionsetMetadata?: any) {
  let [entityName, entityId] = ['', ''];

  if (attributeValue && attributeMetadata) {
      switch (attributeMetadata.AttributeType) {
          case 'DateTime':
                return handleDataTime(attributeMetadata,attributeValue);
          case 'Picklist':
            return handleOptionSet(attributeMetadata, attributeValue,entityOptionsetMetadata);
          case 'Virtual':
            return handleVitualField(attributeMetadata, attributeValue, entityOptionsetMetadata);
          case 'Customer':
          case 'Lookup.Owner':
          case 'Lookup':
          case 'Lookup.Simple':
          case 'Owner':
              [entityName, entityId] = attributeValue.split(',');
              return await handleLookup(entityName, entityId);
          case 'String':
            case 'Money':
            case 'Decimal':
            case 'Double':
            case 'Boolean':
            case 'Integer':
            case 'State':
            case 'Status':
          default:
              return attributeValue;
      }
  } else {
      return <></>;
  }
}

async function handleLookup(entityName: string, entityId: string) {
  let displayName = 'N/A';
  let primaryField = 'Not Empty';
  const link = `${appUrl}&pagetype=entityrecord&etn=${entityName}&id=${entityId}`;

  if (!entityPrimaryField.has(entityName)) {
      primaryField = await getPrimaryColumn(entityName);
  } else {
      primaryField = entityPrimaryField.get(entityName)[0]?.LogicalName;
  }

  if (primaryField !== "Not Empty") {
      displayName = await getDisplayName(entityName, entityId, primaryField);
  }

  return <a href={link}>{displayName}</a>;
}

function handleDataTime (attributeMetadata:any ,attributeValue: any){
  if(attributeMetadata.Format === "DateAndTime"){
    if(attributeMetadata.DateTimeBehavior?.Value === "TimeZoneIndependent"){
      // A TimeZoneIndependent DateTime field stores the value exactly as entered without conversion to UTC or user local time zones.
      return attributeValue;
    }
    if(attributeMetadata.DateTimeBehavior?.Value === "UserLocal"){
      return convertUTCToLocal(attributeValue, false);
    }
  }
  
  if(attributeMetadata.Format === "DateOnly"){
    if(attributeMetadata.DateTimeBehavior?.Value === "TimeZoneIndependent"){
      // A TimeZoneIndependent DateTime field stores the value exactly as entered without conversion to UTC or user local time zones.
      return attributeValue?.split(' ')[0];
    }
    if(attributeMetadata.DateTimeBehavior?.Value === "UserLocal"){
      return convertUTCToLocal(attributeValue, true);
    }
  }
}

function convertUTCToLocal(utcData:any, dateOnly: boolean){
  // Create a Date object using the UTC value
  const utcDataObject = new Date(utcData);
  const utcDateTime = new Date(Date.UTC(utcDataObject.getFullYear(),utcDataObject.getMonth(),utcDataObject.getDate(), utcDataObject.getHours(),utcDataObject.getMinutes(),utcDataObject.getSeconds()));
  // Use toLocaleString() to get the local time
  let localDate = utcDateTime.toLocaleString();
  if(dateOnly) localDate = utcDateTime.toLocaleDateString()

  return localDate;
}

function handleOptionSet(attributeMetadata: any, attributeValue: any, entityOptionsetMetadata?: any){
  let formattedValue = attributeValue;
  if(entityOptionsetMetadata && attributeValue && attributeMetadata){
    for(const optionSetMeteData of entityOptionsetMetadata){
      if(optionSetMeteData.MetadataId === attributeMetadata.MetadataId){
        for(const option of optionSetMeteData.OptionSet.Options){
          if(option.Value?.toString() === attributeValue?.toString()){
            formattedValue = option.Label.UserLocalizedLabel.Label;
          }
        }
      }
    }
  }
  return formattedValue;
}

function handleVitualField(attributeMetadata: any, attributeValue: any, entityOptionsetMetadata:any){
  if(attributeValue && attributeMetadata.AttributeTypeName?.Value === "MultiSelectPicklistType"){
    const valueArray = JSON.parse(attributeValue);
    let options:any =[];
    for(const optionSetMeteData of entityOptionsetMetadata){
      if(optionSetMeteData.MetadataId === attributeMetadata.MetadataId){
        options = optionSetMeteData.OptionSet?.Options;
      }
    }
    if(options.length > 0){
      const formattedString:any = [];
      valueArray.forEach((str:any)=>{
        options.forEach((option:any)=>{
          if(str.toString() === option.Value?.toString()){
            formattedString.push(option.Label.UserLocalizedLabel.Label);
          }
        })
      });
      attributeValue = formattedString.toString();
    }
  }
  return attributeValue;
}
async function getPrimaryColumn(entityLogicalName: string) {
  const ODataAPI = `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes?$select=SchemaName,LogicalName&$filter=IsPrimaryName eq true`;

  const response = await fetch(ODataAPI, {
      method: "GET",
      headers: {
          "Accept": "application/json",
          "Content-Type": "application/json; charset=utf-8",
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          "Prefer": "odata.include-annotations=*",
      },
  });

  if (response.ok) {
      const data = await response.json();
      entityPrimaryField.set(entityLogicalName, data.value);
      return data.value[0]?.LogicalName;
  } else {
      throw new Error(`Error: ${response.status} - ${response.statusText}`);
  }
}

async function getDisplayName(entityLogicalName: string, entityId: string, primaryField: string) {
  const response: any = await Xrm.WebApi.retrieveRecord(entityLogicalName, entityId, `?$select=${primaryField}`);
  return response[primaryField] || 'N/A';
}

export function getModalContentStyles(){
  const theme = createTheme();
  const contentStyles = mergeStyleSets({
    container: {
      display: 'flex',
      flexFlow: 'column wrap',
      width: '40%',
      alignItems: 'stretch',
    },
    header: [
      // eslint-disable-next-line @typescript-eslint/no-deprecated
      theme.fonts.xxLarge,
      {
        flex: '1 1 auto',
        borderTop: `4px solid ${theme.palette.themePrimary}`,
        color: theme.palette.neutralPrimary,
        display: 'flex',
        alignItems: 'center',
        fontWeight: FontWeights.semibold,
        padding: '12px 12px 14px 24px',
      },
    ],
    heading: {
      color: theme.palette.neutralPrimary,
      fontWeight: FontWeights.semibold,
      fontSize: 'inherit',
      margin: '0',
    },
    body: {
      flex: '4 4 auto',
      padding: '0 24px 24px 24px',
      overflowY: 'hidden',
      selectors: {
        p: { margin: '14px 0' },
        'p:first-child': { marginTop: 0 },
        'p:last-child': { marginBottom: 0 },
      },
    },
  });

  return contentStyles;
}

export function getIconStyle(){
  const theme = getTheme();
  const iconButtonStyles = {
    root: {
      color: theme.palette.neutralPrimary,
      marginLeft: 'auto',
      marginTop: '4px',
      marginRight: '2px',
    },
    rootHovered: {
      color: theme.palette.neutralDark,
    },
  };

  return iconButtonStyles;
}
