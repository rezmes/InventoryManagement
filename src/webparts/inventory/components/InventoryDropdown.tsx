import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react";

export interface IInventoryDropdownProps {
  items: IDropdownOption[];
  selectedItem: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
}

class InventoryDropdown extends React.Component<IInventoryDropdownProps, {}> {
  render() {
    const { items, selectedItem, onChange } = this.props;
    const placeHolderText =
      items.length === 0 ? "No items available" : "Select an item";
    return (
      <Dropdown
        placeHolder={placeHolderText}
        options={items}
        onChanged={onChange}
        selectedKey={selectedItem}
      />
    );
  }
}

export default InventoryDropdown;
