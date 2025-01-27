import * as React from "react";
import Select from "react-select";
import { IInventoryProps } from "./IInventoryProps";
import {
  SPHttpClient,
  SPHttpClientBatchConfiguration,
  ISPHttpClientOptions,
  SPHttpClientBatch,
} from "@microsoft/sp-http";
import CustomDropdown from "./CustomDropdown";
import { IDropdownOption } from "office-ui-fabric-react";
export interface InventoryItem {
  value: any;
  itemId: string;
  quantity: number;
  notes: string;
}

export interface IInventoryState {
  transactionType: string;
  formNumber: number | null;
  transactionDate: string;
  items: InventoryItem[];
  itemOptions: { value: string; label: string }[];
  isFormActive: boolean;
}

export default class Inventory extends React.Component<
  IInventoryProps,
  IInventoryState
> {
  constructor(props: IInventoryProps) {
    super(props);

    this.state = {
      transactionType: "",
      formNumber: null,
      transactionDate: new Date().toISOString().substring(0, 10),
      items: [],
      itemOptions: [],
      isFormActive: false,
    };
  }

  componentDidMount() {
    this.fetchInventoryItems();
  }

  private fetchInventoryItems = () => {
    const { spHttpClient, siteUrl } = this.props;

    const url = `${siteUrl}/_api/web/lists/GetByTitle('InventoryItems')/items?$select=ID,Title`;

    spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        if (response.ok) {
          return response.json();
        } else {
          return response.json().then((error) => {
            throw new Error(`Error: ${error.error.message}`);
          });
        }
      })
      .then((data) => {
        // console.log("API response:", data); // Log the entire response
        if (data && data.value) {
          const options = data.value.map((item: any) => ({
            value: item.ID,
            label: item.Title,
          }));
          this.setState({ itemOptions: options });
        } else {
          throw new Error("Unexpected response structure");
        }
      })
      .catch((error) => {
        console.error("There was a problem with the fetch operation:", error);
      });
  };
  handleTransactionTypeChange = (
    event: React.ChangeEvent<HTMLInputElement>
  ): void => {
    this.setState({ transactionType: event.target.value });
  };

  private createForm = () => {
    this.getLastFormNumber()
      .then((lastFormNumber) => {
        const newFormNumber = lastFormNumber + 1;
        this.setState({ formNumber: newFormNumber, isFormActive: true });
      })
      .catch((error) => {
        console.error("Error getting last form number:", error);
      });
  };

  private getLastFormNumber = () => {
    const { spHttpClient, siteUrl } = this.props;

    const url = `${siteUrl}/_api/web/lists/GetByTitle('InventoryTransaction')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`;

    return spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        if (response.ok) {
          return response.json();
        } else {
          return response.json().then((error) => {
            throw new Error(`Error: ${error.error.message}`);
          });
        }
      })
      .then((data) => {
        if (data && data.value && data.value.length > 0) {
          return data.value[0].FormNumber;
        } else {
          return 0; // No previous form numbers, start with 0
        }
      });
  };

  private handleSubmit = () => {
    const { spHttpClient, siteUrl } = this.props;
    const { formNumber, transactionDate, transactionType, items } = this.state;

    if (!siteUrl) {
      console.error("Site URL is not defined");
      return;
    }

    console.log("Submitting batch to:", `${siteUrl}/_api/$batch`);

    // Initialize the batch
    const batch = spHttpClient.beginBatch();

    // Iterate over items and add each one to the batch
    items.forEach((item) => {
      const body = JSON.stringify({
        FormNumber: formNumber,
        TransactionDate: transactionDate,
        TransactionType: transactionType,
        ItemId: item.value, // Assuming "value" corresponds to an ID in the list
        Quantity: transactionType === "Out" ? -item.quantity : item.quantity,
        Notes: item.notes, // Optional
      });

      // Add the request to the batch
      batch.post(
        `${siteUrl}/_api/web/lists/GetByTitle('InventoryTransaction')/items`,
        SPHttpClientBatch.configurations.v1, // Correct usage
        {
          headers: {
            "Content-Type": "application/json",
          },
          body,
        }
      );
    });

    // Execute the batch
    batch
      .execute()
      .then(() => {
        console.log("Batch successfully executed");
      })
      .catch((error) => {
        console.error("Batch execution failed:", error);
      });
  };

  // private handleSubmit = () => {
  //   const { spHttpClient, siteUrl } = this.props;
  //   const { formNumber, transactionDate, transactionType, items } = this.state;

  //   const batch = spHttpClient.beginBatch();

  //   items.forEach((item) => {
  //     const body = JSON.stringify({
  //       FormNumber: formNumber,
  //       TransactionDate: transactionDate,
  //       TransactionType: transactionType,
  //       ItemId: item.value,
  //       Quantity: transactionType === "Out" ? -item.quantity : item.quantity,
  //       Notes: item.notes,
  //     });
  //     const batchConfig = new SPHttpClientBatchConfiguration(
  //       SPHttpClient.configurations.v1
  //     );
  //     batch.post(
  //       `${siteUrl}/_api/web/lists/GetByTitle('InventoryTransaction')/items`,
  //       batchConfig,
  //       {
  //         headers: {
  //           "Content-Type": "application/json",
  //         },
  //         body,
  //       }
  //     );
  //   });

  //   batch
  //     .execute()
  //     .then(() => {
  //       console.log("Items successfully added to InventoryTransaction list");
  //       // Reset form or show success message
  //     })
  //     .catch((error) => {
  //       console.error(
  //         "Error adding items to InventoryTransaction list:",
  //         error
  //       );
  //     });
  // };

  // private handleTransactionTypeChange = (event: React.ChangeEvent<HTMLInputElement>) => {
  //   this.setState({ transactionType: event.target.value });
  // };

  // private handleItemChange = (index: number, selectedOption: any) => {
  //   const items = [...this.state.items];
  //   items[index] = {
  //     ...items[index],
  //     itemId: selectedOption.value,
  //   };
  //   this.setState({ items });
  // };

  // private handleQuantityChange = (index: number, quantity: string) => {
  //   const items = [...this.state.items];
  //   items[index].quantity = parseInt(quantity, 10);
  //   this.setState({ items });
  // };

  // private handleNotesChange = (index: number, notes: string) => {
  //   const items = [...this.state.items];
  //   items[index].notes = notes;
  //   this.setState({ items });
  // };

  handleItemChange = (index: number, option?: IDropdownOption) => {
    const items = [...this.state.items];
    if (option) {
      items[index] = { ...items[index], itemId: option.key as string };
      this.setState({ items });
    }
  };

  handleQuantityChange = (index: number, value: string) => {
    const items = [...this.state.items];
    items[index].quantity = parseInt(value, 10);
    this.setState({ items });
  };

  handleNotesChange = (index: number, value: string) => {
    const items = [...this.state.items];
    items[index].notes = value;
    this.setState({ items });
  };

  render() {
    const { description } = this.props;
    const {
      transactionType,
      formNumber,
      transactionDate,
      items,
      itemOptions,
      isFormActive,
    } = this.state;

    return (
      <div>
        <h1>{description}</h1>
        <div>
          <label>
            <input
              type="radio"
              value="In"
              checked={transactionType === "In"}
              onChange={this.handleTransactionTypeChange}
            />{" "}
            In
          </label>
          <label>
            <input
              type="radio"
              value="Out"
              checked={transactionType === "Out"}
              onChange={this.handleTransactionTypeChange}
            />{" "}
            Out
          </label>
          <button disabled={!transactionType} onClick={this.createForm}>
            Create Form
          </button>
        </div>

        {isFormActive && (
          <div>
            <h3>Form Number: {formNumber}</h3>
            <div>
              <label>Date:</label>
              <input
                type="date"
                value={transactionDate}
                onChange={(e) =>
                  this.setState({ transactionDate: e.target.value })
                }
              />
            </div>
            <div>
              {items.map((item, index) => (
                <div key={index}>
                  <Select
                    name={`item-${index}`}
                    value={itemOptions.filter(
                      (option) => option.value === item.itemId
                    )}
                    options={itemOptions}
                    onChange={(selectedOption) =>
                      this.handleItemChange(index, selectedOption)
                    }
                  />
                  <input
                    type="number"
                    value={item.quantity}
                    onChange={(e) =>
                      this.handleQuantityChange(index, e.target.value)
                    }
                  />
                  <input
                    type="text"
                    value={item.notes}
                    onChange={(e) =>
                      this.handleNotesChange(index, e.target.value)
                    }
                  />
                </div>
              ))}
              <button onClick={this.handleSubmit}>Submit</button>
            </div>
          </div>
        )}
      </div>
    );
  }

  // createForm = () => {
  //   this.getLastFormNumber()
  //     .then((lastFormNumber) => {
  //       this.setState({
  //         formNumber: lastFormNumber + 1,
  //         isFormActive: true,
  //       });
  //     })
  //     .catch((error) =>
  //       console.error("Error fetching last form number:", error)
  //     );
  // };

  // getLastFormNumber = () => {
  //   return fetch(
  //     `${this.props.siteUrl}/_api/web/lists/GetByTitle('InventoryTransaction')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`,
  //     {
  //       headers: {
  //         Accept: "application/json;odata=verbose",
  //       },
  //       credentials: "same-origin",
  //     }
  //   )
  //     .then((response) => response.json())
  //     .then((data) => {
  //       const items = data.d.results;
  //       return items.length > 0 ? items[0].FormNumber : 0;
  //     });
  // };

  // addItem = () => {
  //   const items = [...this.state.items, { itemId: "", quantity: 0, notes: "" }];
  //   this.setState({ items });
  // };

  // removeItem = (index: number) => {
  //   const items = this.state.items.slice();
  //   items.splice(index, 1);
  //   this.setState({ items });
  // };

  // handleItemChange = (
  //   index: number,
  //   option: { value: string; label: string }
  // ) => {
  //   const items = [...this.state.items];
  //   items[index].itemId = option.value;
  //   this.setState({ items });
  // };

  // handleQuantityChange = (
  //   index: number,
  //   event: React.ChangeEvent<HTMLInputElement>
  // ) => {
  //   const items = [...this.state.items];
  //   items[index].quantity = parseInt(event.target.value, 10) || 0;
  //   this.setState({ items });
  // };

  // render() {
  //   const {
  //     transactionType,
  //     formNumber,
  //     transactionDate,
  //     items,
  //     itemOptions,
  //     isFormActive,
  //   } = this.state;

  //   return (
  //     <div>
  //       <div>
  //         <h1>{this.props.description}</h1>
  //       </div>
  //       <div>
  //         <label>
  //           <input
  //             type="radio"
  //             value="In"
  //             checked={transactionType === "In"}
  //             onChange={this.handleTransactionTypeChange}
  //           />{" "}
  //           In
  //         </label>
  //         <label>
  //           <input
  //             type="radio"
  //             value="Out"
  //             checked={transactionType === "Out"}
  //             onChange={this.handleTransactionTypeChange}
  //           />{" "}
  //           Out
  //         </label>
  //         <button disabled={!transactionType} onClick={this.createForm}>
  //           Create Form
  //         </button>
  //       </div>

  //       {isFormActive && (
  //         <div>
  //           <h3>Form Number: {formNumber}</h3>
  //           <div>
  //             <label>Date:</label>
  //             <input
  //               type="date"
  //               value={transactionDate}
  //               onChange={(e) =>
  //                 this.setState({ transactionDate: e.target.value })
  //               }
  //             />
  //           </div>
  //           <div>
  //             {items.map((item, index) => (
  //               <div key={index}>
  //                 <Select
  //                   name={`item-${index}`}
  //                   value={
  //                     itemOptions.filter(
  //                       (option) => option.value === item.itemId
  //                     )[0] || null
  //                   }
  //                   options={itemOptions}
  //                   onChange={(option) => this.handleItemChange(index, option)}
  //                 />
  //                 <input
  //                   type="number"
  //                   value={item.quantity}
  //                   onChange={(e) => this.handleQuantityChange(index, e)}
  //                 />
  //                 <button onClick={() => this.removeItem(index)}>Remove</button>
  //               </div>
  //             ))}
  //             <button onClick={this.addItem}>Add Item</button>
  //           </div>
  //         </div>
  //       )}
  //     </div>
  //   );
  // }
}
