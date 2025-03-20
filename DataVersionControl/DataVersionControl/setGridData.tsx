import * as React from "react";
import { DetailsList, SelectionMode, Stack, IDragDropContext, Toggle, FontIcon, values, StackItem, DetailsListLayoutMode, IStackTokens, Checkbox, IDetailsListStyles } from '@fluentui/react';
export function generateGridData(dataSet: any, columns: []): JSX.Element {

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

    return(                        
    <DetailsList
        items={items}
        columns={columns}
        styles={gridStyles}
        selectionPreservedOnEmptyClick={true}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        selectionMode={SelectionMode.single}>    
    </DetailsList>
    )
}