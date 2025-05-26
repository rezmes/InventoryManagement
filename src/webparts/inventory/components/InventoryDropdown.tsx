// // // // src\webparts\inventory\components\InventoryDropdown.tsx

// import * as React from "react";
// import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

// export interface IInventoryDropdownProps {
//   items: IDropdownOption[];
//   selectedItem: string | number | undefined;
//   onChange: (option?: IDropdownOption) => void;
//   placeholder?: string;
// }

// export interface IInventoryDropdownState {
//   dropdownWidth: number | "auto";
//   filteredItems: IDropdownOption[];
//   isDropdownOpen: boolean;
//   searchText: string;
//   searchByAssetNumber: boolean;
// }

// class InventoryDropdown extends React.Component<
//   IInventoryDropdownProps,
//   IInventoryDropdownState
// > {
//   private dropdownRef: HTMLDivElement | null = null;
//   private searchBoxRef: HTMLInputElement | null = null;

//   constructor(props: IInventoryDropdownProps) {
//     super(props);

//     // Sort items alphabetically by text
//     const sortedItems = [...props.items].sort((a, b) =>
//       a.text.localeCompare(b.text)
//     );

//     this.state = {
//       dropdownWidth: "auto",
//       filteredItems: sortedItems,
//       isDropdownOpen: false,
//       searchText: "",
//       searchByAssetNumber: false,
//     };
//   }

//   componentDidMount() {
//     this.calculateDropdownWidth();
//     document.addEventListener("click", this.handleDocumentClick);
//   }

//   componentDidUpdate(prevProps: IInventoryDropdownProps) {
//     if (prevProps.items !== this.props.items) {
//       // Sort items alphabetically by text
//       const sortedItems = [...this.props.items].sort((a, b) =>
//         a.text.localeCompare(b.text)
//       );

//       this.setState({ filteredItems: sortedItems });
//       this.calculateDropdownWidth();
//     }
//   }

//   componentWillUnmount() {
//     document.removeEventListener("click", this.handleDocumentClick);
//   }

//   private handleDocumentClick = (event: MouseEvent) => {
//     if (this.dropdownRef && !this.dropdownRef.contains(event.target as Node)) {
//       this.setState({ isDropdownOpen: false });
//     }
//   };

//   private calculateDropdownWidth() {
//     // Create an offscreen span to measure text width
//     const span = document.createElement("span");
//     span.style.fontFamily = "IRANSansXFaNum, faSegoe UI, sans-serif";
//     span.style.fontSize = "14px";
//     span.style.visibility = "hidden";
//     span.style.whiteSpace = "nowrap";
//     document.body.appendChild(span);

//     let maxWidth = 0;
//     this.props.items.forEach((item) => {
//       // Measure both item text and asset number (if available)
//       span.innerText = item.text;
//       let width = span.getBoundingClientRect().width;
//       if (width > maxWidth) {
//         maxWidth = width;
//       }

//       if (item.data && item.data.assetNumber) {
//         span.innerText = `${item.data.assetNumber} - ${item.text}`;
//         width = span.getBoundingClientRect().width;
//         if (width > maxWidth) {
//           maxWidth = width;
//         }
//       }
//     });
//     document.body.removeChild(span);

//     const extraPadding = 40;
//     this.setState({ dropdownWidth: maxWidth + extraPadding });
//   }

//   private handleSearchChange = (newValue?: string) => {
//     const searchText = newValue || "";
//     const { searchByAssetNumber } = this.state;

//     // Sort items alphabetically by text
//     const sortedItems = [...this.props.items].sort((a, b) =>
//       a.text.localeCompare(b.text)
//     );

//     let filteredItems = sortedItems;
//     if (searchText) {
//       filteredItems = sortedItems.filter((item) => {
//         if (searchByAssetNumber) {
//           // Search by asset number
//           return (
//             item.data &&
//             item.data.assetNumber &&
//             item.data.assetNumber
//               .toLowerCase()
//               .indexOf(searchText.toLowerCase()) === 0
//           );
//         } else {
//           // Search by item name
//           return (
//             item.text.toLowerCase().indexOf(searchText.toLowerCase()) === 0
//           );
//         }
//       });
//     }

//     this.setState({
//       searchText,
//       filteredItems,
//     });
//   };

//   private toggleSearchMode = () => {
//     this.setState(
//       (prevState) => ({
//         searchByAssetNumber: !prevState.searchByAssetNumber,
//         searchText: "",
//         filteredItems: [...this.props.items].sort((a, b) =>
//           a.text.localeCompare(b.text)
//         ),
//       }),
//       () => {
//         if (this.searchBoxRef) {
//           this.searchBoxRef.focus();
//         }
//       }
//     );
//   };

//   private toggleDropdown = () => {
//     this.setState(
//       (prevState) => {
//         // If opening the dropdown, reset to sorted full list
//         const sortedItems = prevState.isDropdownOpen
//           ? prevState.filteredItems
//           : [...this.props.items].sort((a, b) => a.text.localeCompare(b.text));

//         return {
//           isDropdownOpen: !prevState.isDropdownOpen,
//           searchText: "",
//           filteredItems: sortedItems,
//         };
//       },
//       () => {
//         if (this.state.isDropdownOpen && this.searchBoxRef) {
//           this.searchBoxRef.focus();
//         }
//       }
//     );
//   };

//   private handleOptionClick = (option: IDropdownOption) => {
//     this.props.onChange(option);
//     this.setState({
//       isDropdownOpen: false,
//       searchText: "",
//     });
//   };

//   private setDropdownRef = (ref: HTMLDivElement) => {
//     this.dropdownRef = ref;
//   };

//   private setSearchBoxRef = (ref: HTMLInputElement) => {
//     this.searchBoxRef = ref;
//   };

//   public render(): React.ReactElement<IInventoryDropdownProps> {
//     const { selectedItem, placeholder } = this.props;
//     const { filteredItems, isDropdownOpen, searchByAssetNumber } = this.state;
//     const placeHolderText = placeholder || "انتخاب آیتم";

//     // Find the selected item text and asset number
//     let selectedText = placeHolderText;
//     let selectedAssetNumber = "";
//     if (selectedItem !== undefined && selectedItem !== null) {
//       for (let i = 0; i < this.props.items.length; i++) {
//         if (this.props.items[i].key === selectedItem) {
//           selectedText = this.props.items[i].text;

//           selectedAssetNumber =
//             (this.props.items[i].data &&
//               this.props.items[i].data.assetNumber) ||
//             "";

//           break;
//         }
//       }
//     }

//     const dropdownWidth =
//       typeof this.state.dropdownWidth === "number"
//         ? `${this.state.dropdownWidth}px`
//         : "auto";

//     return (
//       <div ref={this.setDropdownRef} style={{ position: "relative" }}>
//         <div
//           onClick={this.toggleDropdown}
//           style={{
//             width: dropdownWidth,
//             height: "32px",
//             border: "1px solid #a6a6a6",
//             padding: "0 8px",
//             display: "flex",
//             alignItems: "center",
//             justifyContent: "space-between",
//             cursor: "pointer",
//             backgroundColor: "white",
//           }}
//         >
//           <span>
//             {selectedAssetNumber
//               ? `${selectedAssetNumber} - ${selectedText}`
//               : selectedText}
//           </span>
//           <span style={{ fontSize: "10px" }}>▼</span>
//         </div>

//         {isDropdownOpen && (
//           <div
//             style={{
//               position: "absolute",
//               top: "34px",
//               left: "0",
//               width: dropdownWidth,
//               maxHeight: "300px",
//               overflowY: "auto",
//               backgroundColor: "white",
//               border: "1px solid #a6a6a6",
//               zIndex: 1000,
//               boxShadow: "0 2px 4px rgba(0, 0, 0, 0.2)",
//             }}
//           >
//             <div style={{ padding: "8px" }}>
//               <div style={{ display: "flex", marginBottom: "4px" }}>
//                 <input
//                   ref={this.setSearchBoxRef}
//                   type="text"
//                   placeholder={
//                     searchByAssetNumber
//                       ? "جستجو با کد دارایی..."
//                       : "جستجو با نام..."
//                   }
//                   value={this.state.searchText}
//                   onChange={(e) => this.handleSearchChange(e.target.value)}
//                   style={{
//                     width: "100%",
//                     padding: "4px",
//                     border: "1px solid #a6a6a6",
//                   }}
//                 />
//               </div>
//               <div>
//                 <label
//                   style={{
//                     display: "flex",
//                     alignItems: "center",
//                     cursor: "pointer",
//                   }}
//                 >
//                   <input
//                     type="checkbox"
//                     checked={searchByAssetNumber}
//                     onChange={this.toggleSearchMode}
//                   />
//                   <span style={{ marginRight: "4px", fontSize: "12px" }}>
//                     جستجو با کد دارایی
//                   </span>
//                 </label>
//               </div>
//             </div>

//             <div>
//               {filteredItems.map((option) => (
//                 <div
//                   key={option.key.toString()}
//                   onClick={() => this.handleOptionClick(option)}
//                   style={{
//                     padding: "8px",
//                     cursor: "pointer",
//                     backgroundColor:
//                       selectedItem === option.key ? "#f3f2f1" : "transparent",
//                     whiteSpace: "nowrap",
//                     overflow: "hidden",
//                     textOverflow: "ellipsis",
//                   }}
//                 >
//                   {option.data && option.data.assetNumber
//                     ? `${option.data.assetNumber} - ${option.text}`
//                     : option.text}
//                 </div>
//               ))}
//               {filteredItems.length === 0 && (
//                 <div style={{ padding: "8px", color: "#666" }}>
//                   نتیجه‌ای یافت نشد
//                 </div>
//               )}
//             </div>
//           </div>
//         )}
//       </div>
//     );
//   }
// }

// export default InventoryDropdown;

import * as React from "react";
import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

export interface IInventoryDropdownProps {
  items: IDropdownOption[];
  selectedItem: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
  placeholder?: string;
}

export interface IInventoryDropdownState {
  dropdownWidth: number | "auto";
  filteredItems: IDropdownOption[];
  isDropdownOpen: boolean;
  searchText: string;
}

class InventoryDropdown extends React.Component<
  IInventoryDropdownProps,
  IInventoryDropdownState
> {
  private dropdownRef: HTMLDivElement | null = null;
  private searchBoxRef: HTMLInputElement | null = null;

  constructor(props: IInventoryDropdownProps) {
    super(props);

    const sortedItems = [...props.items].sort((a, b) =>
      a.text.localeCompare(b.text)
    );

    this.state = {
      dropdownWidth: "auto",
      filteredItems: sortedItems,
      isDropdownOpen: false,
      searchText: "",
    };
  }

  componentDidMount() {
    this.calculateDropdownWidth();
    document.addEventListener("click", this.handleDocumentClick);
  }

  componentDidUpdate(prevProps: IInventoryDropdownProps) {
    if (prevProps.items !== this.props.items) {
      const sortedItems = [...this.props.items].sort((a, b) =>
        a.text.localeCompare(b.text)
      );

      this.setState({ filteredItems: sortedItems });
      this.calculateDropdownWidth();
    }
  }

  componentWillUnmount() {
    document.removeEventListener("click", this.handleDocumentClick);
  }

  private handleDocumentClick = (event: MouseEvent) => {
    if (this.dropdownRef && !this.dropdownRef.contains(event.target as Node)) {
      this.setState({ isDropdownOpen: false });
    }
  };

  private calculateDropdownWidth() {
    const span = document.createElement("span");
    span.style.fontFamily = "IRANSansXFaNum, faSegoe UI, sans-serif";
    span.style.fontSize = "14px";
    span.style.visibility = "hidden";
    span.style.whiteSpace = "nowrap";
    document.body.appendChild(span);

    let maxWidth = 0;
    this.props.items.forEach((item) => {
      span.innerText = item.text;
      let width = span.getBoundingClientRect().width;
      if (width > maxWidth) {
        maxWidth = width;
      }

      if (item.data && item.data.assetNumber) {
        span.innerText = item.data.assetNumber + " - " + item.text;
        width = span.getBoundingClientRect().width;
        if (width > maxWidth) {
          maxWidth = width;
        }
      }
    });
    document.body.removeChild(span);

    const extraPadding = 40;
    this.setState({ dropdownWidth: maxWidth + extraPadding });
  }

  private handleSearchChange = (newValue?: string) => {
    const searchText = newValue || "";
    const sortedItems = [...this.props.items].sort((a, b) =>
      a.text.localeCompare(b.text)
    );

    let filteredItems = sortedItems;
    if (searchText) {
      filteredItems = sortedItems.filter((item) =>
        item.text.toLowerCase().indexOf(searchText.toLowerCase()) === 0
      );
    }

    this.setState({
      searchText,
      filteredItems,
    });
  };

  private toggleDropdown = () => {
    this.setState(
      (prevState) => {
        const sortedItems = prevState.isDropdownOpen
          ? prevState.filteredItems
          : [...this.props.items].sort((a, b) => a.text.localeCompare(b.text));

        return {
          isDropdownOpen: !prevState.isDropdownOpen,
          searchText: "",
          filteredItems: sortedItems,
        };
      },
      () => {
        if (this.state.isDropdownOpen && this.searchBoxRef) {
          this.searchBoxRef.focus();
        }
      }
    );
  };

  private handleOptionClick = (option: IDropdownOption) => {
    this.props.onChange(option);
    this.setState({
      isDropdownOpen: false,
      searchText: "",
    });
  };

  private setDropdownRef = (ref: HTMLDivElement) => {
    this.dropdownRef = ref;
  };

  private setSearchBoxRef = (ref: HTMLInputElement) => {
    this.searchBoxRef = ref;
  };

  public render(): React.ReactElement<IInventoryDropdownProps> {
    const { selectedItem, placeholder } = this.props;
    const { filteredItems, isDropdownOpen } = this.state;
    const placeHolderText = placeholder || "انتخاب آیتم";

    let selectedText = placeHolderText;
    let selectedAssetNumber = "";
    if (selectedItem !== undefined && selectedItem !== null) {
      for (let i = 0; i < this.props.items.length; i++) {
        const item = this.props.items[i];
        if (
          (item.data && item.data.assetNumber === selectedItem) ||
          item.key === selectedItem
        ) {
          selectedText = item.text;
          selectedAssetNumber = item.data && item.data.assetNumber ? item.data.assetNumber : "";
          break;
        }
      }
    }

    const dropdownWidth =
      typeof this.state.dropdownWidth === "number"
        ? this.state.dropdownWidth + "px"
        : "auto";

    return (
      <div ref={this.setDropdownRef} style={{ position: "relative" }}>
        <div
          onClick={this.toggleDropdown}
          style={{
            width: dropdownWidth,
            height: "32px",
            border: "1px solid #a6a6a6",
            padding: "0 8px",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            cursor: "pointer",
            backgroundColor: "white",
          }}
        >
          <span>
            {selectedAssetNumber
              ? selectedAssetNumber + " - " + selectedText
              : selectedText}
          </span>
          <span style={{ fontSize: "10px" }}>▼</span>
        </div>

        {isDropdownOpen && (
          <div
            style={{
              position: "absolute",
              top: "34px",
              left: "0",
              width: dropdownWidth,
              maxHeight: "300px",
              overflowY: "auto",
              backgroundColor: "white",
              border: "1px solid #a6a6a6",
              zIndex: 1000,
              boxShadow: "0 2px 4px rgba(0, 0, 0, 0.2)",
            }}
          >
            <div style={{ padding: "8px" }}>
              <div style={{ display: "flex", marginBottom: "4px" }}>
                <input
                  ref={this.setSearchBoxRef}
                  type="text"
                  placeholder={placeHolderText}
                  value={this.state.searchText}
                  onChange={(e) => this.handleSearchChange(e.target.value)}
                  style={{
                    width: "100%",
                    padding: "4px",
                    border: "1px solid #a6a6a6",
                  }}
                />
              </div>
            </div>

            <div>
              {filteredItems.map((option) => (
                <div
                  key={option.key.toString()}
                  onClick={() => this.handleOptionClick(option)}
                  style={{
                    padding: "8px",
                    cursor: "pointer",
                    backgroundColor:
                      selectedItem === option.key ||
                      selectedItem === (option.data && option.data.assetNumber)
                        ? "#f3f2f1"
                        : "transparent",
                    whiteSpace: "nowrap",
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                  }}
                >
                  {option.data && option.data.assetNumber
                    ? option.data.assetNumber + " - " + option.text
                    : option.text}
                </div>
              ))}
              {filteredItems.length === 0 && (
                <div style={{ padding: "8px", color: "#666" }}>
                  نتیجه‌ای یافت نشد
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    );
  }
}

export default InventoryDropdown;