declare interface IInventoryWebPartStrings {
  AssetNumber: ReactNode;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  InventoryItemsListNameLabel: string; // Add this line
  InventoryTransactionListNameLabel: string; // Add this line
  InventoryManagement: string;
  CreateForm: string;
  Date: string;
  FormNumber: string;
  TransactionType: string;
  In: string;
  Out: string;
  Item: string;
  Quantity: string;
  Notes: string;
  Actions: string;
  AddRow: string;
  Submit: string;
  Cancel: string;
  Required: string;
  Remove: string;
  IssuedReturnedBy: string; // <-- Add this line
  AssetNumber: string; // Added for AssetNumber label
}

declare module 'InventoryWebPartStrings' {
  const strings: IInventoryWebPartStrings;
  export = strings;
}
