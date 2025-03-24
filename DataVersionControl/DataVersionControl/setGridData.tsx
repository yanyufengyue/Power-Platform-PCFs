import * as React from "react";
import { DetailsList, SelectionMode, Stack, IDragDropContext, Toggle, FontIcon, values, StackItem, DetailsListLayoutMode, IStackTokens, Checkbox, IDetailsListStyles, IComboBoxOption } from '@fluentui/react';
import * as Papa from "papaparse";
export function generateGridData(noteAttachments: any, selectedComboOptions: IComboBoxOption[], columns: []): JSX.Element {
  const gridStyles: Partial<IDetailsListStyles> = {
    contentWrapper: {
      flex: '1 1 auto',
      overflow: 'scroll',
      height: 500,
    },

    headerWrapper: {
      flex: '0 0 auto',
    },
    root: {
      overflowX: 'scroll',
      selectors: {
        '& [role=grid]': {
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'start',
        },
      },
    },
  };
  const detailLists = noteAttachments.map((attachment: any) =>{
    if(selectedComboOptions.some((file:any, index:any)=>file.key === attachment.annotationid)){
        let dataSet:any = [];
        const base64Body = attachment.documentbody;
        const binaryStringText = atob(base64Body);
        Papa.parse(binaryStringText, {
          complete: (result:any) => {
            dataSet = result
          },
          header: true, // If your CSV has headers
          skipEmptyLines: true, // Skip empty lines
          dynamicTyping: true, // Automatically converts types
        });
      
        const items = dataSet.data.map((row: any, index:any)=>{
          const attributes = dataSet.meta.fields.map((column:any)=>{
              return{
                  [column]:row[column]
              }
          });
          return Object.assign({
              Key: index,
              raw: row
          },...attributes)
        });
      
        return(
          <Stack.Item key={attachment.annotationid} style={{width:100/selectedComboOptions.length+'%', border: '1px solid rgb(164, 169, 177)'}}>
            <DetailsList
              items={items}
              columns={columns}
              styles={gridStyles}
              selectionPreservedOnEmptyClick={true}
              layoutMode={DetailsListLayoutMode.fixedColumns}
              selectionMode={SelectionMode.single}>    
            </DetailsList>
          </Stack.Item>           
        )
      }
    }
  )
  return(detailLists)
}