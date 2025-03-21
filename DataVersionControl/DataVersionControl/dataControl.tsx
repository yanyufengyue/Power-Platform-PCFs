import * as React from 'react';
import { IInputs } from './generated/ManifestTypes';
import { getNoteAttachment } from './helper';
import * as Papa from 'papaparse';
import { generateGridData } from './setGridData';
import { useBoolean } from '@fluentui/react-hooks';
import { Checkbox, DefaultButton, Dropdown, IDropdownOption, IDropdownStyles, IStackTokens, Panel, Stack, StackItem } from '@fluentui/react';

export interface IDataControlProps {
  context: ComponentFramework.Context<IInputs>;
}

export const DataControl = React.memo(({context}:IDataControlProps): JSX.Element =>{
  const pageContext: any = Xrm.Utility.getPageContext();
  const [noteAttachments, setNoteAttachments] = React.useState<any>();
  
  const [parsedCSV, setParsedCSV] = React.useState<object>();
  const [columns, setColumns] = React.useState<any>([]);
  const [gridColumn, setGridColumns] = React.useState<any>([]);
  const [detailList, setDetailList] = React.useState<JSX.Element>();
  const [checkBoxes, setCheckBoxes] = React.useState<JSX.Element>();
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(false);

  const [option, setOption] = React.useState<IDropdownOption>();
  const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 300 },
  };
  const options: IDropdownOption[] = [
    { key: 0, text: 'Function Locations' },
    { key: 1, text: 'Maintenace Assets' },
    { key: 3, text: 'Asset Hierarchy' },
    { key: 4, text: 'Asset Group' },
    { key: 2, text: 'Asset Attributes' },
  ];
  function optionOnChange(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number){
    setOption(option);
  }

  React.useEffect(()=>{

    const fetchAttachments = async () => {
      const attachments = await getNoteAttachment(pageContext.input.entityId, context, option);
      setNoteAttachments(attachments);
    };

    if(option){
      fetchAttachments();
    }
  },[option])

  React.useEffect(()=>{
    if(noteAttachments){
      const base64Body = noteAttachments[0].documentbody;
      const binaryStringText = atob(base64Body);
      Papa.parse(binaryStringText, {
        complete: (result:any) => {
          setParsedCSV(result);
        },
        header: true, // If your CSV has headers
        skipEmptyLines: true, // Skip empty lines
        dynamicTyping: true, // Automatically converts types
      });
    }
  },[noteAttachments])

  React.useEffect(()=>{
    if(parsedCSV){
      const columnMetadata =(parsedCSV as any).meta.fields.map((item : any)=>{
        return{
            name: item,
            fieldName: item,
            Key: item,
            isResizable: true,
            dataType: "string"
        }
      })
      setColumns(columnMetadata);
      setGridColumns(columnMetadata);
    }
  }, [parsedCSV]);

  React.useEffect(()=>{
    if(parsedCSV && gridColumn){
      const itemList =generateGridData(parsedCSV, gridColumn);
      setDetailList(itemList);
    }
  }, [gridColumn]);

  function _onChange(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean) {
    let updatedGridColumns: any = [...gridColumn];
    const columnName = (ev?.target as HTMLInputElement).ariaLabel; 
    if(isChecked){
    // If isChecked is true, add the column to the grid
      const columnToAdd = columns.find((col: any) => col.name === columnName);
      if (columnToAdd && !gridColumn.some((col:any) => col.name === columnName)) {
        const sequencedColumn:any = [];
        columns.forEach((col:any)=>{
          // Adding the column back in it's orginal sequence
          if(col.name === columnName || gridColumn.some((column:any)=> column.name === col.name)){
            sequencedColumn.push(col);
          }
        });
        updatedGridColumns = sequencedColumn;
      }
    } else {
      // If isChecked is false, remove the column from the grid
      updatedGridColumns = updatedGridColumns.filter((col: any) => col.name !== columnName);
    }
    setGridColumns(updatedGridColumns);
  } 

  React.useEffect(()=>{
    const myCheckBoxes = columns.map((item:any, index: any)=>{
      return(            
        <Stack.Item key={index}>
          <Checkbox label={item.name} ariaLabel={item.name} name={index} checked={gridColumn.some((col:any) => col.name === item.name)} onChange={_onChange}/>
        </Stack.Item>
      )
    })
    setCheckBoxes(myCheckBoxes);
  },[columns, gridColumn])

  return(
  <div style={{width: '100%'}}>
    <Stack verticalFill>
      <StackItem>
        <Dropdown placeholder="Select an option" label="Select a table to explore" options={options} styles={dropdownStyles} onChange={optionOnChange} />
        <DefaultButton text="Update Columns in View" onClick={openPanel} />
      </StackItem>
      <StackItem>
        <Stack horizontal>
          <StackItem style={{width:'50%'}}>{detailList}</StackItem>
          <StackItem style={{width:'50%'}}>{detailList}</StackItem>
        </Stack>
      </StackItem>
    </Stack>
    <Panel
      headerText="Select/Remove"
      // this prop makes the panel non-modal
      isBlocking={false}
      isOpen={isOpen}
      onDismiss={dismissPanel}
      closeButtonAriaLabel="Close"
    >
      <Stack tokens={{childrenGap: 10}}>
        {checkBoxes}
      </Stack>
    </Panel>
  </div>)
});

DataControl.displayName = DataControl.displayName ?? "DataControl";