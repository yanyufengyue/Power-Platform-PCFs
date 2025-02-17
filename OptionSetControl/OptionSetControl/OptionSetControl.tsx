import * as React from 'react';
import { ChoiceGroup, IChoiceGroupOption, Icon } from '@fluentui/react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';

export interface IOptionSetProps {
  label: string | undefined;
  value: number | null;
  options: ComponentFramework.PropertyHelper.OptionMetadata[] | undefined;
  iconConfig: string | null;
  onChange: (newValue: number | undefined) => void;
  disabled: boolean;
  masked: boolean;
  formFactor: 'small' | 'large';
  fieldRequiredLevel: number | undefined;
}

export const OptionSetComponent = React.memo((props: IOptionSetProps) =>{
  const {label, value, options, iconConfig, onChange, disabled, masked, formFactor, fieldRequiredLevel} = props;
  const[items, setItems] = React.useState<any>({});
  const[selectedValue, setSelectedValue] = React.useState<any>(value);
  React.useEffect(()=>{setSelectedValue(value)}, [value])
  React.useMemo(()=>{
    const myOptions = options;
    let mockOptions = true;
    myOptions?.forEach((e)=>{if(e.Value === -1){mockOptions = false}});
    if(mockOptions && fieldRequiredLevel !== 1 && fieldRequiredLevel !== 2) myOptions?.unshift({Label:'---Select---', Value: -1, Color: ''});
    
    const myItems =() => {
      'use strict';
      let iconMapping: Record<number, string> = {};
      let configError: string | undefined;
      if (iconConfig) {
        try {
          iconMapping = JSON.parse(iconConfig) as Record<number, string>;
        } 
        catch {
          configError = `Invalid configuration: '${iconConfig}'`;
        }
      }

      return {
          error: configError,
          choices: myOptions?.map((item) => {
              return {
                  key: item.Value.toString(),
                  value: item.Value,
                  text: item.Label.includes('---Select---') ? 'Clear Value' : item.Label,
                  iconProps: { iconName: item.Label.includes('---Select---')? 'Clear' :iconMapping[item.Value] },
                  disabled: disabled
              } as IChoiceGroupOption;
          }),
          options: myOptions?.map((item) => {
            return {
                key: item.Value.toString(),
                data: { value: item.Value, icon: iconMapping[item.Value] },
                text: item.Label,
            } as IDropdownOption;
          }),
      };
    }
    setItems(myItems);
  },[options,iconConfig, fieldRequiredLevel]);
  
  const onChangeChoiceGroup = React.useCallback((ev?: unknown, option?: IChoiceGroupOption): void => {
    'use strict';
      onChange(option ? (option.value as number) : undefined);
      setSelectedValue(option !== undefined ? option?.value?.toString() : undefined);
  },[onChange]);

  const iconStyles = {marginRight: '8px'};
  const onRenderOption = (option?: IDropdownOption): JSX.Element =>{
    'use strict';
    if (option) {
      return (
        <div>
            {option.data && option.data.icon && (
              <Icon
                  style={iconStyles}
                  iconName={option.data.icon}
                  aria-hidden="true"
                  title={option.data.icon} />
            )}
            <span>{option?.text}</span>
        </div>
      );
    }
    return <></>;
  }

  const onRenderTitle = (options?: IDropdownOption[]): JSX.Element => {
    'use strict';
    if (options) {
        return onRenderOption(options[0]);
    }
    return <></>;
  }

  const onChangeDropDown = React.useCallback((ev: unknown, option?: IDropdownOption): void => {
    'use strict';
      onChange(option ? (option.data.value as number) : undefined);
      setSelectedValue(option !== undefined ? option?.data?.value.toString() : undefined);
    },[onChange])


  return(
    <>
      {items.error}
      {masked && '****'}
      {formFactor === 'large' && !items.error && !masked && (
        <ChoiceGroup
          //label={label}
          options={items.choices}
          selectedKey= {selectedValue?.toString()}
          onChange={onChangeChoiceGroup}
        />
      )}
      {formFactor === 'small' && !items.error && !masked && (
        <Dropdown
          placeholder={'---'}
          //label={label}
          ariaLabel={label}
          options={items.options}
          selectedKey={selectedValue?.toString()}
          disabled={disabled}
          onRenderOption={onRenderOption}
          onRenderTitle={onRenderTitle}
          onChange={onChangeDropDown}
        />
      )}
    </>
  )
});
OptionSetComponent.displayName = 'OptionSetComponent';