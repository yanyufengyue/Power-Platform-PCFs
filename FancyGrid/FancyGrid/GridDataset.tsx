import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IDetailsListStyles, Link, SelectionMode, Selection, Stack, IconButton, IDragDropEvents, IDragDropContext, mergeStyles } from '@fluentui/react';
import * as Helper from './Helper';
import { getTheme } from '@fluentui/react/lib/Styling';
"use strict";

type Dataset = ComponentFramework.PropertyTypes.DataSet;

export interface IGridDatasetProps {
  dataset?: Dataset;
  optionsetMetadata?: any;
  initialPageNumber: number;
}

export const gridDataset = React.memo(({dataset, optionsetMetadata, initialPageNumber}: IGridDatasetProps): JSX.Element =>{
  if(dataset?.paging.totalResultCount === undefined || dataset?.paging.totalResultCount === 0) return<>No records available</>;
  const [items, setItems] = React.useState<any>([]);
  const [columns, setColumns] = React.useState<any>([]);
  const [currentPage, setCurrentPage] = React.useState<number>(initialPageNumber);
  const [hasPreviousPage, setHasPreviousPage] = React.useState<boolean>();
  const [hasNextPage, setHasNextPage] = React.useState<boolean>();
  const [selectedRows, setSelectedRows] = React.useState<any>([]);
  const [showSelections, setShowSelections] = React.useState<string>();
  const [pageType, setPageType] = React.useState(Xrm.Utility.getPageContext().input.pageType);
  const gridDivId = 'fancyGrid';
  const footerStackId = 'footerStack';
  let draggedItem:any;
  const windowSize= Helper.useWindowSize();
  // Re-render the grid by refresh the dataset when optionsetMetadata is ready
  React.useEffect(()=>{"use strict"; dataset?.refresh();},[optionsetMetadata]);
  // Re-render the grid records when the currentPage number is updated with new values
  React.useEffect(()=>{dataset.paging.loadExactPage(currentPage);}, [currentPage]);
  React.useEffect(()=>{if(selectedRows?.length > 0){setShowSelections('inline')} else{setShowSelections('none')}}, [selectedRows]);
  React.useEffect(()=>{
    "use strict";
    setColumns(dataset?.columns.map((column)=>{
      return{
        name: column.displayName,
        fieldName: column.name,
        minWidth: column.visualSizeFactor,
        Key: column.name,
        isResizable: true,
        dataType: column.dataType,
      }
    }));

    const myItems = dataset?.sortedRecordIds.map((id)=>{
      const record = dataset.records[id];
      const attributes = dataset?.columns.map((column)=>{
        return{[column.name] : record.getFormattedValue(column.name)}
      });
      return Object.assign({
        Key: record.getRecordId(),
        raw: record
      }, ...attributes)
    });

    setItems(myItems);
    setHasNextPage(dataset.paging.hasNextPage);
    setHasPreviousPage(dataset.paging.hasPreviousPage);
    // When dataset is updated(subgrid refresh), re-set the current page to the initalized page number in the dataset
    setCurrentPage(initialPageNumber);
  },[dataset]);

  React.useEffect(()=>{
    if (pageType === 'entitylist'){
      const newHeight=()=>{
        const footerStack = document.getElementById(footerStackId);
        const subgridHeaderWrapper = document.getElementById(gridDivId)?.querySelector('.ms-DetailsList-headerWrapper');
        return((((document.getElementById(gridDivId)?.getBoundingClientRect().height as number) as number) - (footerStack?.getBoundingClientRect().height as number) - (subgridHeaderWrapper?.getBoundingClientRect().height as number) - 18));
      };
      (document.getElementById(gridDivId)?.querySelector('.ms-DetailsList-contentWrapper') as HTMLElement).style.height = newHeight.toString()+"px";
    }
  }, [windowSize])
  
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
      case 'OptionSet':
        if(optionsetMetadata && item[column.fieldName]){
          return Helper.handleOptionSetColumn(column,item, optionsetMetadata);
        }else{
          return <>{item[column.fieldName]}</>;
        }
        break;
      default:
        return <>{item[column.fieldName]}</>;
    }
  };
    // Function to calculate row height dynamically based on the first row's height
    const calculateHeight = () => {
      "use strict";
      const firstRow = document.getElementById(gridDivId)?.querySelector('.ms-DetailsRow');
      if(pageType === "entityrecord"){
        if (firstRow) {
          return((firstRow.getBoundingClientRect().height * dataset?.paging.pageSize));
        }
      }else if(pageType === 'entitylist'){
        const footerStack = document.getElementById(footerStackId);
        const subgridHeaderWrapper = document.getElementById(gridDivId)?.querySelector('.ms-DetailsList-headerWrapper');
        return(((document.getElementById(gridDivId)?.getBoundingClientRect().height as number) - (footerStack?.getBoundingClientRect().height as number) - (subgridHeaderWrapper?.getBoundingClientRect().height as number) - 18));
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
  
  const selection = new Selection({
    onSelectionChanged: () =>{
      "use strict";
      
      const selectedItems = selection.getSelection()
      if(selectedItems.length === 0) dataset?.clearSelectedRecordIds();
      const itemIds: any = [];
      selectedItems.forEach((item: any) =>{itemIds.push(item.Key);});
      dataset?.setSelectedRecordIds(itemIds);
      setSelectedRows(selectedItems);
    }
  });
  
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
        return mergeStyles({
          backgroundColor: getTheme().palette.neutralLight,
        });
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
    const draggedItems = selectedRows.find((record:any) => record.Key === item.Key) !== undefined ? draggedItem : selectedRows;
    const insertIndex = items.indexOf(item);
    const filteredItems = items.filter((itm:any) => draggedItems.indexOf(itm) === -1);

    filteredItems.splice(insertIndex, 0, ...draggedItems);
    setItems(filteredItems);
  }

  const onFirstPage = () =>{
    "use strict";
    setCurrentPage(1);
  };
  const onPreviousPage = () =>{
    "use strict";
    setCurrentPage(currentPage - 1);
  };
  const onNextPage = () =>{
    "use strict";
    setCurrentPage(currentPage + 1);
  };

  function onItemInvoke(item?: any, index?: number | undefined, ev?: Event | undefined): void {
    "use strict";
    if(pageType === "entitylist"){
      dataset?.openDatasetItem(item.raw.getNamedReference());
    }
    else{
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
    }
  }

  return(
    <div id={gridDivId} style={{width: '100%'}}>
      <DetailsList
        items={items}
        columns={columns}
        styles={gridStyles}
        selectionPreservedOnEmptyClick={true}
        selectionMode={SelectionMode.multiple}
        selection={selection}
        onItemInvoked={onItemInvoke}
        layoutMode={DetailsListLayoutMode.justified}
        onRenderItemColumn={renderItemColumn}
        dragDropEvents={onDragRowItem()}>
      </DetailsList>
      <Stack id={footerStackId} horizontal style={{width: '100%'}}>
        <Stack.Item align='center' style={{width:'50%'}}>
          <Stack horizontal style={{float: 'left'}}>
            <Stack.Item align='center'>
              <>Rows: {dataset.paging.totalResultCount === -1 ? '5000+' : dataset.paging.totalResultCount}</>
            </Stack.Item>
            <Stack.Item align='center' style={{paddingLeft:'10px', display:showSelections}}>
              <>Selected: {selectedRows.length}</>
            </Stack.Item>
          </Stack>
        </Stack.Item>
        <Stack.Item style={{width: '50%'}}>
          <Stack horizontal style={{float: 'right'}}>
            <Stack.Item align='center' style={{paddingRight: '2px'}}>
              <>Page: {currentPage} / {Math.ceil(dataset.paging.totalResultCount/dataset.paging.pageSize)}</>
            </Stack.Item>
            <IconButton
              autoFocus={false}
              alt="First Page"
              iconProps={{ iconName: 'Rewind' }}
              disabled={!hasPreviousPage}
              onClick={onFirstPage}
            />
            <IconButton
              autoFocus={false}
              alt="Previous Page"
              iconProps={{ iconName: 'Previous' }}
              disabled={!hasPreviousPage}
              onClick={onPreviousPage}
            />
            <IconButton
              autoFocus={false}
              alt="Next Page"
              iconProps={{ iconName: 'Next' }}
              disabled={!hasNextPage}
              onClick={onNextPage}
            />
          </Stack>
        </Stack.Item>
      </Stack>
    </div>
  )
});

gridDataset.displayName = gridDataset.displayName ? gridDataset.displayName:"sortDataSet";
