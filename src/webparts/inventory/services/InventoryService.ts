// src\webparts\inventory\services\InventoryService.ts
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export class InventoryService {
  getPersonnel(): { Id: string; Title: string; }[] | PromiseLike<{ Id: string; Title: string; }[]> {
    throw new Error('Method not implemented.');
  }
  private spHttpClient: SPHttpClient;
  private siteUrl: string;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this.spHttpClient = spHttpClient;
    this.siteUrl = siteUrl;
  }


  // In InventoryService.ts
public async getInventoryItems(listName: string): Promise<any[]> {
  const url = `${this.siteUrl}/_api/web/lists/GetByTitle('${listName}')/items?$select=Title,ID,AssetNumber`;

  const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!response.ok) {
    const error = await response.json();
    throw new Error(`Error fetching inventory items: ${error.error.message}`);
  }
  const data = await response.json();
  return data.value || [];
}


  public async getMechanicPersonnel(listName: string, fieldName: string): Promise<any[]> {
    // Extract the root URL dynamically
    const urlObj = new URL(this.siteUrl);
    const rootUrl = `${urlObj.protocol}//${urlObj.host}`; // This gives "http://<root>"

    // Append the hardcoded path
    const targetSiteUrl = `${rootUrl}/pwa/manufacP`;

    // Build the REST API endpoint
    const apiUrl = `${targetSiteUrl}/_api/web/lists/GetByTitle('${listName}')/items?$select=Id,${fieldName}`;

    const response: SPHttpClientResponse = await this.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Error fetching mechanic personnel: ${error.error.message}`);
    }

    const data = await response.json();
    return data.value || [];
  }



  public async getLastFormNumber(listName: string): Promise<number> {
    const url = `${this.siteUrl}/_api/web/lists/GetByTitle('${listName}')/items?$select=FormNumber&$orderby=FormNumber desc&$top=1`;

    const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Error fetching last form number: ${error.error.message}`);
    }
    const data = await response.json();
    return data && data.value && data.value.length > 0 ? parseInt(data.value[0].FormNumber, 10) || 0 : 0;
  }

  public async getItemTitle(listName: string, itemId: number): Promise<string> {
    const url = `${this.siteUrl}/_api/web/lists/GetByTitle('${listName}')/items(${itemId})?$select=Title`;

    const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Error fetching item title: ${error.error.message}`);
    }
    const data = await response.json();
    return data.Title;
  }

  public async submitTransaction(listName: string, item: any, requestDigest: string): Promise<SPHttpClientResponse> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    return this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': requestDigest
      },
      body: JSON.stringify(item)
    });
  }

  public async getRequestDigest(): Promise<string> {
    const url = `${this.siteUrl}/_api/contextinfo`;

    const response: SPHttpClientResponse = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      }
    });
    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Error fetching request digest: ${error.error.message}`);
    }
    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }
}
