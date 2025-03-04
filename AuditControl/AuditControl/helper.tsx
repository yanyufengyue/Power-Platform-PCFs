import * as React from "react";
import { IInputs } from "./generated/ManifestTypes";


export function handleChangeData(currentEntityAttributes: any, JSONdata: string): React.JSX.Element{
    const sortedAtrributes:any = [];
    currentEntityAttributes.value.forEach((attribute: any) => {
        if(attribute.AttributeOf === null) sortedAtrributes.push(attribute);
    });
    const data = JSON.parse(JSONdata).changedAttributes;
    const tableData:any =[];
    sortedAtrributes.forEach((element: any) => {
        data.forEach((e:any)=>{
            if(element.LogicalName ===e.logicalName){
                // Create rows with data for the table
                tableData.push(
                    {fieldName:element.DisplayName.UserLocalizedLabel.Label,
                     oldValue:e.oldValue,
                     newValue:e.newValue
                    }
                );
            }
        })
    });
    return(<div>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr style={{background: 'lightblue'}}>
              <th>Changed Column</th>
              <th>Old Value</th>
              <th>New Value</th>
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
}