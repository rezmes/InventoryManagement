import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react";

export interface ICustomDropdownProps {
  options: IDropdownOption[];
  selectedKey: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
  placeholder: string;
}

class CustomDropdown extends React.Component<ICustomDropdownProps, {}> {
  render() {
    const { options, selectedKey, onChange, placeholder } = this.props;

    return (
      <Dropdown
        placeHolder={placeholder}
        options={options}
        onChanged={onChange}
        selectedKey={selectedKey}
      />
    );
  }
}

export default CustomDropdown;
