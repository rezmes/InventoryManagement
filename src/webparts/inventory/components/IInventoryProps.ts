import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

export interface IInventoryProps {
  description: string;
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  inventoryListName: string; // Add this
  transactionListName: string; // Add this
}

