import { FontIcon } from "@fluentui/react";
import * as React from "react";


export function handleOptionSetColumn (column?: any, item?: any, optionsetMetadata?: any){
    "use strict";
    const optionsMetadata = optionsetMetadata.filter((ele: any)=>ele.LogicalName === column.fieldName);
    const optionValueColor = optionsMetadata[0].OptionSet.Options.filter((ele: any)=>item[column.fieldName] === ele.Label.UserLocalizedLabel.Label);
    return <><FontIcon aria-label="Dictionary" iconName="CheckBoxfill" style={{color:optionValueColor[0].Color}}></FontIcon> {item[column.fieldName]}</>;
}

export function useWindowSize() {
    const [size, setSize] = React.useState<any>();
    React.useEffect(() => {
      function updateSize() {
        setSize(window.innerHeight);
      }
      window.addEventListener('resize', updateSize);
      updateSize();
      return () => window.removeEventListener('resize', updateSize);
    }, []);
    return size;
  }