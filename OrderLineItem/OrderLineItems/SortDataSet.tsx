import * as React from 'react';
import { DetailsList, IDragDropEvents ,IDragDropContext, Selection, DetailsRow, PrimaryButton, IDetailsListStyles, DetailsListLayoutMode, SelectionMode, Link, styled} from '@fluentui/react';
import { getTheme, mergeStyles} from '@fluentui/react/lib/Styling';
"use strict";

type DataSet = ComponentFramework.PropertyTypes.DataSet;

export interface IDataSetProps {
  dataSet?: DataSet;
  sequenceColumn?: any;
  optionSetMetadata?: any;
  optionSetColumn?: any;
}

export const sortDataSet = React.memo(({dataSet, sequenceColumn, optionSetMetadata,optionSetColumn}: IDataSetProps) : JSX.Element =>{
  const [items, setItems] = React.useState<any>([]);
  const [colums, setColums] = React.useState<any>([]);
  const [selectedRecords, setSelectedRecords] = React.useState<any>([]);

  let draggedItem:any;
  const theme = getTheme();
  const dragEnterClass = mergeStyles({
    backgroundColor: theme.palette.neutralLight,
  });

  const selection= new Selection({
    onSelectionChanged: () =>{
      "use strict";
      const selectedItems = selection.getSelection()
      setSelectedRecords(selectedItems);
      if(selectedItems.length === 0){
        dataSet?.clearSelectedRecordIds();
      }else{
        const itemIds: any = [];
        selectedItems.forEach((item: any) =>{
          itemIds.push(item.Key);
        });
        dataSet?.setSelectedRecordIds(itemIds);
      }
    }
  })

  React.useEffect(()=>{
    "use strict";
    setColums(dataSet?.columns.map((column) =>{
      return{
        name: column.displayName,
        fieldName: column.name,
        minWidth: column.visualSizeFactor,
        Key: column.name,
        isResizable: true,
        dataType: column.dataType,
        showSortIconWhenUnsorted: true
      }
    }));

    const myItems = dataSet?.sortedRecordIds.map((id)=>{
      const record = dataSet.records[id];
      const attributes = dataSet.columns.map((column)=>{
        return {[column.name] : record.getFormattedValue(column.name)}
      });
      return Object.assign({
        Key: record.getRecordId(),
        raw: record
      }, ...attributes)
    });
    myItems?.sort((a:any, b:any) =>{
      return Number(a[sequenceColumn.name]) > Number(b[sequenceColumn.name]) ? 1 : -1;
    });
    setItems(myItems);
  },[dataSet])

  //Action of users double-click on a row
  const onItemInvoke = (item: any, index?: number, event?: Event): void =>{
    "use strict";
    const pageInput: Xrm.Navigation.PageInputEntityRecord = {
      pageType: "entityrecord",
      entityName: item.raw.getNamedReference()._etn,
      entityId: item.raw.getNamedReference()._id //replace with actual ID
    };
    const navigationOptions: Xrm.Navigation.NavigationOptions = {
        target: 2,
        height: {value: 80, unit:"%"},
        width: {value: 70, unit:"%"},
        position: 1
    };
    Xrm.Navigation.navigateTo(pageInput, navigationOptions);
  };

  const onDragRowItem =():IDragDropEvents=> {
    "use strict";
    return{
      canDrop: (dropContext?: IDragDropContext, dragContext?: IDragDropContext) => {
        return true;
      },
      canDrag: (item?: any) => {
        return true;
      },
      onDragEnter: (item?: any, event?: DragEvent) => {
        // return string is the css classes that will be added to the entering element.
        return dragEnterClass;
      },
      onDragLeave: (item?: any, event?: DragEvent) => {
        return;
      },
      onDrop: (item?: any, event?: DragEvent) => {
        if (draggedItem) {
          insertBeforeItem(item);
        }
      },
      onDragStart: (item?: any, itemIndex?: number, selectedItems?: any[], event?: MouseEvent) => {
        draggedItem = Array(item);
      },
      onDragEnd: (item?: any, event?: DragEvent) => {
        draggedItem = undefined;
      },  
    }
  }

  const insertBeforeItem =(item: any) => {
    "use strict";
    const draggedItems = selectedRecords.find((record:any) => record.Key === item.Key) !== undefined ? draggedItem : selectedRecords;
    const insertIndex = items.indexOf(item);
    const filteredItems = items.filter((itm:any) => draggedItems.indexOf(itm) === -1);

    filteredItems.splice(insertIndex, 0, ...draggedItems);
    setItems(filteredItems);
  }

  const renderRow = (props: any) => {
    "use strict";
    const customStyles: React.CSSProperties = {};
    customStyles.backgroundColor = 'transparent'; // Default background color
    const value: string = props?.item?.[optionSetColumn?.name] ? props?.item?.[optionSetColumn?.name] : '';
    if(optionSetMetadata!== undefined && optionSetMetadata.length > 0){
      optionSetMetadata.forEach((element:any) => {
        if(element.label === value)
          customStyles.backgroundColor = element.color;
      });
    }
    return <DetailsRow {...props} styles={{ root: customStyles }} />;
  }

  const renderItemColumn = (item?: any, index?: number, column?: any) =>{
    "use strict";
    const appUrl = Xrm.Utility.getGlobalContext().getCurrentAppUrl(); 
    let componentUrl, link
    switch(column?.dataType){
      case 'Lookup.Simple':
        componentUrl = "&pagetype=entityrecord&etn=" + item.raw._record.fields[column.fieldName].reference.etn + "&id="+item.raw._record.fields[column.fieldName].reference.id.guid;
        link = appUrl + componentUrl;
        return<Link href={link}>{item[column.fieldName]}</Link>
      case 'Lookup.Owner':
        componentUrl = "&pagetype=entityrecord&etn=" + item.raw._record.fields[column.fieldName].reference.etn + "&id="+item.raw._record.fields[column.fieldName].reference.id.guid;
        link = appUrl + componentUrl;
        return<Link href={link}>{item[column.fieldName]}</Link>
      default:
        return <>{item[column.fieldName]}</>;
    }
  }

  const onClickReorderButton = (buttonId: string)=>{
    "use strict";
    switch (buttonId){
      case "updateorder":
        return () =>{
          "use strict";
          if(items){
            const promises =[];
            for(let i=0; i< items.length; i++){
              promises.push(
                new Promise((resolve, reject) => {
                  Xrm.WebApi.updateRecord(items[i].raw.getNamedReference()._etn,items[i].raw.getNamedReference()._id,{[sequenceColumn.name]: i+1}).then(
                    function success(){
                      resolve(i);
                    },
                    function (error){
                      reject(error);
                    }
                  )
                })
              );
            }
            Promise.all(promises).then(()=>{dataSet?.refresh()});
          }
        };
        break;
      case "addnew":
        return () =>{
          "use strict";
          const totalNumber = dataSet?.paging.totalResultCount ?? 0;
          const parentEntityName =Xrm.Utility.getPageContext().input.entityName;
          const parentEntityId = Xrm.Utility.getPageContext().input.pageType === "entityrecord" ? (Xrm.Utility.getPageContext().input as Xrm.EntityFormPageContext).entityId : "00000000-0000-0000-0000-000000000000";
          const pageInput: Xrm.Navigation.PageInputEntityRecord = {
            pageType: "entityrecord",
            entityName: dataSet!.getTargetEntityType(),
            createFromEntity:{
             id: parentEntityId!,
             entityType: parentEntityName
            },
            data: {[sequenceColumn.name]: totalNumber + 1}
          };
          const navigationOptions: Xrm.Navigation.NavigationOptions = {
              target: 2,
              height: {value: 80, unit:"%"},
              width: {value: 70, unit:"%"},
              position: 1
          };
          Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(()=>{
            dataSet?.refresh();
          })
        };
        break;
      case "increaserow":
        return () =>{
          "use strict";
          const firstRowHeight = document.getElementById('aa')?.querySelector('.ms-DetailsRow')?.getBoundingClientRect().height;
          const contentWrapper = document.getElementById('aa')?.querySelector('.ms-DetailsList-contentWrapper')?.getBoundingClientRect().height;
          if(firstRowHeight === undefined || contentWrapper === undefined) return;
          const currentRowCount = contentWrapper! / firstRowHeight!;
          (document.getElementById('aa')?.querySelector('.ms-DetailsList-contentWrapper') as HTMLElement).style.height = ((firstRowHeight as number) * (currentRowCount + 1)).toString()+"px";
        }
        break;
      case "loadnextpage":
        return () => {
          "use strict";
          if (dataSet?.paging.hasNextPage) dataSet.paging.loadNextPage(true);
        }
        break;
      default:
        return ()=>{};
    }
  }
  
  // Function to calculate row height dynamically based on the first row's height
  const calculateHeight = () => {
    "use strict";
    const firstRow = document.getElementById('aa')?.querySelector('.ms-DetailsRow');
    //const rowHeader = document.getElementById('aa')?.querySelector('.ms-DetailsHeader');
    if (firstRow) {
      return((firstRow.getBoundingClientRect().height * dataSet?.paging.pageSize!));
    }
  };

  const gridStyles: Partial<IDetailsListStyles> = {
    contentWrapper: {
      flex: '1 1 auto',
      overflow: 'auto',
      height: calculateHeight(),
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

  return (
    <div id='aa' className='full-width-grid'>
      <DetailsList
        isHeaderVisible={true}
        layoutMode={DetailsListLayoutMode.justified}
        items={items}
        columns={colums}
        selectionPreservedOnEmptyClick={true}
        selectionMode={SelectionMode.multiple}
        selection={selection}
        onRenderRow={renderRow}
        onRenderItemColumn={renderItemColumn}
        onItemInvoked={onItemInvoke}
        dragDropEvents={onDragRowItem()}
        styles={gridStyles}>
      </DetailsList>
      <div>      
        <div style={{"float":"left","width":"50%"}}>
          <PrimaryButton id="updateorder" text="Update Order" onClick={onClickReorderButton("updateorder")} style={{"padding":"1px"}}/>
          <PrimaryButton id ="addnew" text="Add New" onClick={onClickReorderButton("addnew")} style={{"margin":"0px 2px 0px 2px", "padding":"1px"}}/>
          <PrimaryButton id ="increaserow" text="Increase Row" onClick={onClickReorderButton("increaserow")} style={{"padding":"1px"}}/>
        </div>
        <div style={{"float":"right","width":"50%"}}>  
          <PrimaryButton id ="loadnextpage" disabled={!dataSet?.paging.hasNextPage} text="Load Next Page" onClick={onClickReorderButton("loadnextpage")} style={{"float":"right","margin":"0px 0px 0px 10px", "padding":"1px"}}/>
          <p style={{"float":"right",}}>Loaded: {[items.length]} / {[dataSet?.paging.totalResultCount]} </p>
        </div>
      </div>

    </div>
  )
});


sortDataSet.displayName = sortDataSet.displayName ? sortDataSet.displayName:"sortDataSet";