import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

export interface IInventoryDropdownProps {
  items: IDropdownOption[];
  selectedItem: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
}

class InventoryDropdown extends React.Component<IInventoryDropdownProps, {}> {
  handleChange = (
    event?: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.props.onChange(option);
    }
  };

  render() {
    const { items, selectedItem, onChange } = this.props;
    const placeHolderText =
      items.length === 0 ? "No items available" : "Select an item";
    console.log("Dropdown props:", this.props); // Log props
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
