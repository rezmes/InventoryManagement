// src\webparts\inventory\components\InventoryDropdown.tsx

import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

export interface IInventoryDropdownProps {
  items: IDropdownOption[];
  selectedItem: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
  placeholder?: string;
}

export interface IInventoryDropdownState {
  dropdownWidth: number | "auto";
}

class InventoryDropdown extends React.Component<
  IInventoryDropdownProps,
  IInventoryDropdownState
> {
  state: IInventoryDropdownState = {
    dropdownWidth: "auto",
  };

  componentDidMount() {
    this.calculateDropdownWidth();
  }

  componentDidUpdate(prevProps: IInventoryDropdownProps) {
    if (prevProps.items !== this.props.items) {
      this.calculateDropdownWidth();
    }
  }

  private calculateDropdownWidth() {
    // Create an offscreen span to measure text width
    const span = document.createElement("span");
    // Apply similar styles to those used in the Dropdown for consistency
    span.style.fontFamily = "IRANSansXFaNum, faSegoe UI, sans-serif";
    span.style.fontSize = "14px";
    span.style.visibility = "hidden";
    span.style.whiteSpace = "nowrap";
    document.body.appendChild(span);

    let maxWidth = 0;
    // Measure each item's text width.
    this.props.items.forEach((item) => {
      span.innerText = item.text;
      const width = span.getBoundingClientRect().width;
      if (width > maxWidth) {
        maxWidth = width;
      }
    });
    document.body.removeChild(span);

    // Add extra pixels to account for padding, dropdown arrow, etc.
    const extraPadding = 40;
    this.setState({ dropdownWidth: maxWidth + extraPadding });
  }

  public render(): React.ReactElement<IInventoryDropdownProps> {
    const { items, selectedItem, onChange, placeholder } = this.props;
    const placeHolderText = placeholder || "انتخاب آیتم";
  // Compute the correct key: if selectedItem is a string, find the matching option's key.
  let computedSelectedKey = null;
  if (selectedItem !== undefined && selectedItem !== null) {
    if (typeof selectedItem === "number") {
      computedSelectedKey = selectedItem;
    } else if (typeof selectedItem === "string") {
      for (let i = 0; i < items.length; i++) {
        if (items[i].text === selectedItem) {
          computedSelectedKey = items[i].key;
          break;
        }
      }
    }
  }
    return (
      <Dropdown
        placeHolder={placeHolderText}
        options={items}
        onChanged={onChange}
        selectedKey={selectedItem || null}
        // Apply the computed width via styles
        style={{
          dropdown: {
            width:
              typeof this.state.dropdownWidth === "number"
                ? `${this.state.dropdownWidth}px`
                : "auto",
          },
          dropdownOptionText: {
            whiteSpace: "nowrap",
          },
          callout: {
            // You might also want the callout to mimic that width
            width:
              typeof this.state.dropdownWidth === "number"
                ? `${this.state.dropdownWidth}px`
                : "auto",
          },
        }}
      />
    );
  }
}

export default InventoryDropdown;
