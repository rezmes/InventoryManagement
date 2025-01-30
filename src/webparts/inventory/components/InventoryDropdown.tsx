import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

export interface IInventoryDropdownProps {
  items: IDropdownOption[];
  selectedItem: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
  placeholder?: string;
}

class InventoryDropdown extends React.Component<IInventoryDropdownProps, {}> {
  public render(): React.ReactElement<IInventoryDropdownProps> {
    const { items, selectedItem, onChange, placeholder } = this.props;
    const placeHolderText = placeholder || "Select an item";

    return (
      <Dropdown
        placeHolder={placeHolderText}
        options={items}
        onChanged={onChange}
        selectedKey={selectedItem || null}
      />
    );
  }
}

export default InventoryDropdown;
