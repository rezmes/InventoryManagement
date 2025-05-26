// // // src\webparts\inventory\components\Inventory.tsx

// import * as React from 'react'
// import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown'
// import {
//   PrimaryButton,
// } from 'office-ui-fabric-react'
// import * as moment from 'moment-jalaali'
// import { IInventoryProps } from './IInventoryProps'
// import InventoryDropdown from './InventoryDropdown'
// import { IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox'
// import { InventoryService } from '../services/InventoryService'
// import * as strings from 'InventoryWebPartStrings'
// import { findOption } from '../utils/findOption'

// interface IInventoryItem {
//   ID: number
//   Title: string
//   AssetNumber: string | undefined
// }

// export interface InventoryItem {
//   itemId: string
//   quantity: number
//   notes: string | undefined
//   assetNumber: string | undefined
// }

// export interface IInventoryState {
//   itemOptions: IComboBoxOption[]
//   assetNumberOptions: IComboBoxOption[]
//   mechanicDropdownOptions: IDropdownOption[]
//   selectedItem: string | number | undefined
//   selectedAssetNumber: string | number | undefined
//   formNumber: number | undefined
//   transactionType: string
//   transactionDate: string
//   items: Array<{ itemId: number; quantity: number; notes: string; assetNumber: string }>
//   rows: Array<{
//     issuedReturnedBy: string | number | undefined
//     itemId: number | undefined
//     assetNumber: string | undefined
//     quantity: number
//     notes: string
//   }>
//   inventoryItems: Array<{ key: number; text: string; assetNumber: string }>
//   isFormActive: boolean
//   formValid: boolean
// }

// export default class Inventory extends React.Component<IInventoryProps, IInventoryState> {
//   private inventoryService: InventoryService

//   constructor(props: IInventoryProps) {
//     super(props)
//     this.inventoryService = new InventoryService(props.spHttpClient, props.siteUrl)
//     this.state = {
//       transactionType: '',
//       formNumber: undefined,
//       transactionDate: moment().format('jYYYY/jM/jD'),
//       items: [],
//       rows: [],
//       inventoryItems: [],
//       mechanicDropdownOptions: [],
//       itemOptions: [],
//       assetNumberOptions: [],
//       isFormActive: false,
//       selectedItem: undefined,
//       selectedAssetNumber: undefined,
//       formValid: true
//     }
//   }

  

//   public componentDidMount(): void {
//     console.log('Component mounted, fetching inventory items...')
//     this.fetchInventoryItems()
//     this.fetchMechanicPersonnel()
//   }

//   private fetchInventoryItems = async (): Promise<void> => {
//     const { inventoryItemsListName } = this.props
//     try {
//       const items: IInventoryItem[] = await this.inventoryService.getInventoryItems(inventoryItemsListName)
//       const itemOptions: IComboBoxOption[] = items.map((item: IInventoryItem) => ({
//         key: item.ID,
//         text: item.Title,
//         data: { assetNumber: item.AssetNumber }
//       }))
//       const assetNumberOptions: IComboBoxOption[] = items
//         .filter((item: IInventoryItem) => item.AssetNumber)
//         .map((item: IInventoryItem) => ({
//           key: item.AssetNumber,
//           text: item.AssetNumber,
//           data: { itemId: item.ID, itemTitle: item.Title }
//         }))
//       console.log('Fetched item options:', itemOptions)
//       console.log('Fetched asset number options:', assetNumberOptions)
//       this.setState({ itemOptions, assetNumberOptions })
//     } catch (error) {
//       console.error('Error fetching inventory items:', error)
//       this.setState({ itemOptions: [], assetNumberOptions: [] })
//     }
//   }

//   private createForm = async () => {
//     try {
//       const lastFormNumber = await this.inventoryService.getLastFormNumber(
//         this.props.inventoryTransactionListName
//       );
//       this.setState({
//         formNumber: lastFormNumber + 1,
//         isFormActive: true,
//         rows: [
//           { itemId: null, assetNumber: null, quantity: 1, notes: "", issuedReturnedBy: null },
//         ],
//       });
//     } catch (error) {
//       console.error("Error getting last form number:", error);
//     }
//   };

//   private handleSubmit = async (): Promise<void> => {
//     const { siteUrl, inventoryTransactionListName } = this.props
//     const { rows, formNumber, transactionType, transactionDate } = this.state
  
//     if (!this.validateForm()) {
//       console.log('Form is invalid.')
//       return
//     }
  
//     try {
//       const digestResponse: any = await fetch('/_api/siteusers', {
//         method: 'POST',
//         headers: { Accept: 'application/json;odata=verbose' }
//       })
//       const digestData: any = await digestResponse.json()
//       const requestDigest: string = digestData.d
  
//       const transactionDateISO: string = moment().format('YYYY-MM-DD')
  
//       const requests: any[] = await Promise.all(
//         rows.map(async (row) => {
//           const itemTitle: string = await this.inventoryService.getItemTitle(this.props.inventoryItemsListName, row.itemId!)
//           const quantity: number = transactionType === 'Out' ? -Math.abs(row.quantity) : row.quantity
//           const selectedOption: IDropdownOption | undefined = findOption(this.state.mechanicDropdownOptions, row.issuedReturnedBy)
//           const personnelText: string = selectedOption ? selectedOption.text : ''
//           const item: any = {
//             __metadata: { type: 'SP.Data.' + inventoryTransactionListName + 'ListItem' },
//             FormNumber: formNumber,
//             ItemNameId: row.itemId,
//             Title: itemTitle,
//             AssetNumber: row.assetNumber,
//             Quantity: quantity,
//             IssuedReturnedBy: personnelText,
//             Notes: row.notes,
//             TransactionType: transactionType,
//             TransactionDate: transactionDateISO
//           }
//           console.log('Submitting payload:', JSON.stringify(item))
//           return fetch(
//             siteUrl + '/_api/web/lists/getbytitle(\'' + inventoryTransactionListName + '\')/items',
//             {
//               method: 'POST',
//               headers: {
//                 Accept: 'application/json;odata=verbose',
//                 'Content-Type': 'application/json;odata=verbose',
//                 'X-RequestDigest': requestDigest
//               },
//               body: JSON.stringify(item)
//             }
//           )
//         })
//       )
  
//       for (const response of requests) {
//         if (!response.ok) {
//           const errorText: string = await response.text()
//           throw new Error(errorText)
//         }
//       }
  
//       console.log('All requests successful!')
//       this.resetForm()
//     } catch (error) {
//       console.error('Error submitting transactions:', error)
//     }
//   }
//   private fetchMechanicPersonnel = async () => {
//     try {
//       const items = await this.inventoryService.getMechanicPersonnel(
//         "پرسنل معاونت مکانیک",
//         "LastNameFirstName"
//       );
//       const options: IComboBoxOption[] = items.map((item: any) => ({
//         key: item.Id,
//         text: item.LastNameFirstName,
//       }));
//       console.log("Fetched mechanic personnel options:", options);
//       this.setState({ mechanicDropdownOptions: options });
//     } catch (error) {
//       console.error("Error fetching mechanic personnel:", error);
//       this.setState({ mechanicDropdownOptions: [] });
//     }
//   };

//   private validateForm = (): boolean => {
//     const { rows } = this.state;
//     const isValid = rows.every((row) => row.itemId !== null && row.assetNumber !== null && row.quantity >= 1);
//     this.setState({ formValid: isValid });
//     return isValid;
//   };

//   private handleTransactionTypeChange = (
//     event: React.ChangeEvent<HTMLInputElement>
//   ) => {
//     const transactionType = event.target.value;
//     console.log("Transaction Type Changed:", transactionType);
//     this.setState({ transactionType });
//   };

//   private handleRowChange = (index: number, field: string, value: any, option: IDropdownOption) => {
//     const rows = [...this.state.rows];
//     if (field === "itemId" && option) {
//       rows[index] = {
//         ...rows[index],
//         itemId: Number(option.key),
//         assetNumber: option.data && option.data.assetNumber ? option.data.assetNumber : null,
//       };
//     } else if (field === "assetNumber" && option) {
//       rows[index] = {
//         ...rows[index],
//         assetNumber: value,
//         itemId: option.data && option.data.itemId ? Number(option.data.itemId) : null,
//       };
//     } else {
//       rows[index] = { ...rows[index], [field]: value };
//     }
//     this.setState({ rows }, this.validateForm);
//   };

//   private addRow = () => {
//     this.setState(
//       (prevState) => ({
//         rows: [...prevState.rows, { itemId: null, assetNumber: null, quantity: 1, notes: "", issuedReturnedBy: null }],
//       }),
//       this.validateForm
//     );
//   };

//   private removeRow = (index: number) => {
//     this.setState(
//       (prevState) => ({
//         rows: prevState.rows.filter((_, i) => i !== index),
//       }),
//       this.validateForm
//     );
//   };

//   private resetForm = () => {
//     this.setState({
//       transactionType: "",
//       formNumber: null,
//       transactionDate: moment().format("jYYYY/jM/jD"),
//       rows: [],
//       isFormActive: false,
//       selectedItem: undefined,
//       selectedAssetNumber: undefined,
//       formValid: true,
//     });
//   };
//   public render(): React.ReactElement<IInventoryProps> {
//     const {
//       itemOptions,
//       assetNumberOptions,
//       isFormActive,
//       formNumber,
//       transactionType,
//       transactionDate,
//       rows,
//       formValid,
//     } = this.state;
//     return (
//       <div>
//         <h2>{strings.InventoryManagement}</h2>
//         {!isFormActive && (
//           <div>
//             <div>
//               <label>
//                 <input
//                   type='radio'
//                   name='transactionType'
//                   value='In'
//                   checked={transactionType === 'In'}
//                   onChange={this.handleTransactionTypeChange}
//                   aria-checked={transactionType === 'In'}
//                 />
//                 {strings.In}
//               </label>
//             </div>
//             <div>
//               <label>
//                 <input
//                   type='radio'
//                   name='transactionType'
//                   value='Out'
//                   checked={transactionType === 'Out'}
//                   onChange={this.handleTransactionTypeChange}
//                   aria-checked={transactionType === 'Out'}
//                 />
//                 {strings.Out}
//               </label>
//             </div>
//             <PrimaryButton
//               text={strings.CreateForm}
//               onClick={this.createForm}
//               disabled={!transactionType}
//             />
//           </div>
//         )}

//         {isFormActive && (
//           <div>
//             <h3>{strings.FormNumber}: {formNumber}</h3>
//             <div>
//               <label>{strings.Date}:</label>
//               <input
//                 type='text'
//                 value={transactionDate}
//                 onChange={(event) =>
//                   this.setState({
//                     transactionDate: event.target.value || moment().format('jYYYY/jM/jD')
//                   })
//                 }
//               />
//             </div>
//             <div>
//               <label>
//                 {strings.TransactionType}: {transactionType === 'In' ? strings.In : strings.Out}
//               </label>
//             </div>
//             <table>
//               <thead>
//                 <tr>
//                   <th>{strings.AssetNumber}</th>
//                   <th>{strings.Item}</th>
//                   <th>{strings.Quantity}</th>
//                   <th>{strings.IssuedReturnedBy}</th>
//                   <th>{strings.Notes}</th>
//                   <th>{strings.Actions}</th>
//                 </tr>
//               </thead>
//               <tbody>
//                 {rows.map((row, index) => (
//                   <tr key={index}>
//                     <td>
//                       <InventoryDropdown
//                         items={assetNumberOptions}
//                         selectedItem={row.assetNumber}
//                         onChange={(option) =>
//                           this.handleRowChange(
//                             index,
//                             'assetNumber',
//                             option && option.data && option.data.itemId ? option.text : undefined,
//                             option
//                           )
//                         }
//                         placeholder='انتخاب کد دارایی'
//                       />
//                       {row.assetNumber === undefined && (
//                         <span style={{ color: 'red' }}>{strings.Required}</span>
//                       )}
//                     </td>
//                     <td>
//                       <InventoryDropdown
//                         items={itemOptions}
//                         selectedItem={row.itemId}
//                         onChange={(option) =>
//                           this.handleRowChange(
//                             index,
//                             'itemId',
//                             option ? Number(option.key) : undefined,
//                             option
//                           )
//                         }
//                         placeholder='انتخاب آیتم'
//                       />
//                       {row.itemId === undefined && (
//                         <span style={{ color: 'red' }}>{strings.Required}</span>
//                       )}
//                     </td>
//                     <td>
//                       <input
//                         type='number'
//                         value={row.quantity.toString()}
//                         onChange={(event) =>
//                           this.handleRowChange(
//                             index,
//                             'quantity',
//                             Math.max(parseInt(event.target.value, 10), 1),
//                             undefined
//                           )
//                         }
//                         min='1'
//                         aria-valuemin={1}
//                         aria-valuenow={row.quantity}
//                         aria-valuemax={1000}
//                       />
//                     </td>
//                     <td>
//                       <InventoryDropdown
//                         items={this.state.mechanicDropdownOptions}
//                         selectedItem={row.issuedReturnedBy}
//                         onChange={(option) =>
//                           this.handleRowChange(
//                             index,
//                             'issuedReturnedBy',
//                             option ? option.key : undefined,
//                             undefined
//                           )
//                         }
//                         placeholder='انتخاب فرد'
//                       />
//                     </td>
//                     <td>
//                       <input
//                         type='text'
//                         value={row.notes}
//                         onChange={(event) =>
//                           this.handleRowChange(index, 'notes', event.target.value,undefined)
//                         }
//                       />
//                     </td>
//                     <td>
//                       <PrimaryButton
//                         text={strings.Remove}
//                         onClick={() => this.removeRow(index)}
//                       />
//                     </td>
//                   </tr>
//                 ))}
//               </tbody>
//             </table>

//             <PrimaryButton text={strings.AddRow} onClick={this.addRow} />
//             <PrimaryButton
//               text={strings.Submit}
//               onClick={this.handleSubmit}
//               disabled={!formValid}
//             />
//             <PrimaryButton text={strings.Cancel} onClick={this.resetForm} />
//           </div>
//         )}
//       </div>
//     )
//   }
// }


import * as React from 'react'
import {
  IDropdownOption,
  PrimaryButton,
} from 'office-ui-fabric-react'
import * as moment from 'moment-jalaali'
import { IInventoryProps } from './IInventoryProps'
import InventoryDropdown from './InventoryDropdown'
import { IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox'
import { InventoryService } from '../services/InventoryService'
import * as strings from 'InventoryWebPartStrings'
import { findOption } from '../utils/findOption'

interface IInventoryItem {
  ID: number
  Title: string
  AssetNumber: string | undefined
}

interface InventoryItem {
  itemId: string
  quantity: number
  notes: string | undefined
  assetNumber: string | undefined
}

interface IInventoryState {
  itemOptions: IComboBoxOption[]
  assetNumberOptions: IComboBoxOption[]
  mechanicDropdownOptions: IDropdownOption[]
  selectedItem: string | number | undefined
  selectedAssetNumber: string | number | undefined
  formNumber: number | undefined
  transactionType: string
  transactionDate: string
  items: Array<{ itemId: number; quantity: number; notes: string; assetNumber: string }>
  rows: Array<{
    issuedReturnedBy: string | number | undefined
    itemId: number | undefined
    assetNumber: string | undefined
    quantity: number
    notes: string
  }>
  inventoryItems: Array<{ key: number; text: string; assetNumber: string }>
  isFormActive: boolean
  formValid: boolean
}

export default class Inventory extends React.Component<IInventoryProps, IInventoryState> {
  private inventoryService: InventoryService

  constructor(props: IInventoryProps) {
    super(props)
    this.inventoryService = new InventoryService(props.spHttpClient, props.siteUrl)
    this.state = {
      transactionType: '',
      formNumber: undefined,
      transactionDate: moment().format('jYYYY/jM/jD'),
      items: [],
      rows: [],
      inventoryItems: [],
      mechanicDropdownOptions: [],
      itemOptions: [],
      assetNumberOptions: [],
      isFormActive: false,
      selectedItem: undefined,
      selectedAssetNumber: undefined,
      formValid: true
    }
  }

  public componentDidMount(): void {
    console.log('Component mounted, fetching inventory items...')
    this.fetchInventoryItems()
    this.fetchMechanicPersonnel()
  }

  private fetchInventoryItems = async (): Promise<void> => {
    const { inventoryItemsListName } = this.props
    try {
      const items: IInventoryItem[] = await this.inventoryService.getInventoryItems(inventoryItemsListName)
      const itemOptions: IComboBoxOption[] = items.map((item: IInventoryItem) => ({
        key: item.ID,
        text: item.Title,
        data: { assetNumber: item.AssetNumber }
      }))
      const assetNumberOptions: IComboBoxOption[] = items
        .filter((item: IInventoryItem) => item.AssetNumber)
        .map((item: IInventoryItem) => ({
          key: item.AssetNumber,
          text: item.AssetNumber,
          data: { itemId: item.ID, itemTitle: item.Title }
        }))
      console.log('Fetched item options:', itemOptions)
      console.log('Fetched asset number options:', assetNumberOptions)
      this.setState({ itemOptions, assetNumberOptions })
    } catch (error) {
      console.error('Error fetching inventory items:', error)
      this.setState({ itemOptions: [], assetNumberOptions: [] })
    }
  }

  private handleSubmit = async (): Promise<void> => {
    const { siteUrl, inventoryTransactionListName } = this.props
    const { rows, formNumber, transactionType, transactionDate } = this.state

    if (!this.validateForm()) {
      console.log('Form is invalid.')
      return
    }

    try {
      const digestResponse: Response = await fetch(`${siteUrl}/_api/contextinfo`, {
        method: 'POST',
        headers: { Accept: 'application/json;odata=verbose' }
      })
      const digestData: { d: { GetContextWebInformation: { FormDigestValue: string } } } = await digestResponse.json()
      const requestDigest: string = digestData.d.GetContextWebInformation.FormDigestValue

      const transactionDateISO: string = moment(transactionDate, 'jYYYY/jM/jD').toISOString()

      const requests: Response[] = await Promise.all(
        rows.map(async (row) => {
          const itemTitle: string = await this.inventoryService.getItemTitle(this.props.inventoryItemsListName, row.itemId!)
          const quantity: number = transactionType === 'Out' ? -Math.abs(row.quantity) : row.quantity
          const selectedOption: IDropdownOption | undefined = findOption(this.state.mechanicDropdownOptions, row.issuedReturnedBy)
          const personnelText: string = selectedOption ? selectedOption.text : ''
          const item: any = {
            __metadata: { type: 'SP.Data.' + inventoryTransactionListName + 'ListItem' },
            FormNumber: formNumber,
            ItemNameId: row.itemId,
            Title: itemTitle,
            AssetNumber: row.assetNumber,
            Quantity: quantity,
            IssuedReturnedBy: personnelText,
            Notes: row.notes,
            TransactionType: transactionType,
            TransactionDate: transactionDateISO
          }
          console.log('Submitting payload:', JSON.stringify(item))
          return fetch(
            siteUrl + '/_api/web/lists/getbytitle(\'' + inventoryTransactionListName + '\')/items',
            {
              method: 'POST',
              headers: {
                Accept: 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'X-RequestDigest': requestDigest
              },
              body: JSON.stringify(item)
            }
          )
        })
      )

      for (const response of requests) {
        if (!response.ok) {
          const errorText: string = await response.text()
          throw new Error(errorText)
        }
      }

      console.log('All requests successful!')
      this.resetForm()
    } catch (error) {
      console.error('Error submitting transactions:', error)
    }
  }

  private validateForm = (): boolean => {
    const { rows } = this.state
    const isValid: boolean = rows.every((row) => row.itemId !== undefined && row.assetNumber !== undefined && row.quantity >= 1)
    this.setState({ formValid: isValid })
    return isValid
  }

  private handleTransactionTypeChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const transactionType: string = event.target.value
    console.log('Transaction Type Changed:', transactionType)
    this.setState({ transactionType })
  }

  private handleRowChange = (index: number, field: string, value: any, option: IComboBoxOption | undefined): void => {
    const rows = [...this.state.rows]
    if (field === 'itemId' && option) {
      rows[index] = {
        ...rows[index],
        itemId: Number(option.key),
        assetNumber: option.data && option.data.assetNumber ? option.data.assetNumber : undefined
      }
    } else if (field === 'assetNumber' && option) {
      rows[index] = {
        ...rows[index],
        assetNumber: option.text,
        itemId: option.data && option.data.itemId ? Number(option.data.itemId) : undefined
      }
    } else {
      rows[index] = { ...rows[index], [field]: value }
    }
    this.setState({ rows }, this.validateForm)
  }

  private addRow = (): void => {
    this.setState(
      (prevState) => ({
        rows: [
          ...prevState.rows,
          { itemId: undefined, assetNumber: undefined, quantity: 1, notes: '', issuedReturnedBy: '' }
        ]
      }),
      this.validateForm
    )
  }

  private removeRow = (index: number): void => {
    this.setState(
      (prevState) => ({
        rows: prevState.rows.filter((_, i) => i !== index)
      }),
      this.validateForm
    )
  }

  private resetForm = (): void => {
    this.setState({
      transactionType: '',
      formNumber: undefined,
      transactionDate: moment().format('jYYYY/jM/jD'),
      items: [],
      rows: [],
      selectedItem: undefined,
      selectedAssetNumber: undefined,
      formValid: true
    })
  }

  private createForm = (): void => {
    this.setState({
      isFormActive: true,
      formNumber: Math.floor(Math.random() * 100000) + 1
    })
  }

  private fetchMechanicPersonnel = async (): Promise<void> => {
    try {
      const personnel: Array<{ Id: string; Title: string }> = await this.inventoryService.getPersonnel()
      const mechanicDropdownOptions: IDropdownOption[] = personnel.map((item) => ({
        key: item.Id,
        text: item.Title
      }))
      this.setState({ mechanicDropdownOptions })
    } catch (error) {
      console.error('Error fetching mechanic personnel:', error)
      this.setState({ mechanicDropdownOptions: [] })
    }
  }

  public render(): React.ReactElement<IInventoryProps> {
    const {
      itemOptions,
      assetNumberOptions,
      isFormActive,
      formNumber,
      transactionType,
      transactionDate,
      rows,
      formValid
    } = this.state
    return (
      <div>
        <h2>{strings.InventoryItem}</h2>
        {!isFormActive && (
          <div>
            <div>
              <label>
                <input
                  type='radio'
                  name='transactionType'
                  value='In'
                  checked={transactionType === 'In'}
                  onChange={this.handleTransactionTypeChange}
                  aria-checked={transactionType === 'In'}
                />
                {strings.In}
              </label>
            </div>
            <div>
              <label>
                <input
                  type='radio'
                  name='transactionType'
                  value='Out'
                  checked={transactionType === 'Out'}
                  onChange={this.handleTransactionTypeChange}
                  aria-checked={transactionType === 'Out'}
                />
                {strings.Out}
              </label>
            </div>
            <PrimaryButton
              text={strings.CreateForm}
              onClick={this.createForm}
              disabled={!transactionType}
            />
          </div>
        )}

        {isFormActive && (
          <div>
            <h3>{strings.FormNumber}: {formNumber}</h3>
            <div>
              <label>{strings.Date}:</label>
              <input
                type='text'
                value={transactionDate}
                onChange={(event) =>
                  this.setState({
                    transactionDate: event.target.value || moment().format('jYYYY/jM/jD')
                  })
                }
              />
            </div>
            <div>
              <label>
                {strings.TransactionType}: {transactionType === 'In' ? strings.In : strings.Out}
              </label>
            </div>
            <table>
              <thead>
                <tr>
                  <th>{strings.AssetNumber}</th>
                  <th>{strings.Item}</th>
                  <th>{strings.Quantity}</th>
                  <th>{strings.IssuedReturnedBy}</th>
                  <th>{strings.Notes}</th>
                  <th>{strings.Actions}</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => (
                  <tr key={index}>
                    <td>
                      <InventoryDropdown
                        items={assetNumberOptions}
                        selectedItem={row.assetNumber}
                        onChange={(option) =>
                          this.handleRowChange(index, 'assetNumber', option ? option.text : undefined, option)
                        placeholder='Select Asset Number'
                      />
                      {row.assetNumber === undefined && (
                        <span style={{ color: 'red' }}>{strings.Required}</span>
                      )}
                    </td>
                    <td>
                      <InventoryDropdown
                        items={itemOptions}
                        selectedItem={row.itemId}
                        onChange={(option) =>
                          this.handleRowChange(index, 'itemId', option ? Number(option.key) : undefined, option)
                        }
                        placeholder='Select Item'
                      />
                      {row.itemId === undefined && (
                        <span style={{ color: 'red' }}>{strings.Required}</span>
                      )}
                    </td>
                    <td>
                      <input
                        type='number'
                        value={row.quantity.toString()}
                        onChange={(event) =>
                          this.handleRowChange(
                            index,
                            'quantity',
                            Math.max(parseInt(event.target.value, 10) || 1, 1),
                            undefined
                          )
                        }
                        min='1'
                        aria-label='Quantity'
                        aria-valuemin={1}
                        aria-valuenow={row.quantity}
                      />
                    </td>
                    <td>
                      <InventoryDropdown
                        items={this.state.mechanicDropdownOptions}
                        selectedItem={row.issuedReturnedBy}
                        onChange={(option) =>
                          this.handleRowChange(index, 'issuedReturnedBy', option ? option.key : undefined, option)
                        }
                        placeholder='Select Personnel'
                      />
                    </td>
                    <td>
                      <input
                        type='text'
                        value={row.notes}
                        onChange={(event) =>
                          this.handleRowChange(index, 'notes', event.target.value,undefined)
                        }
                        aria-label='Notes'
                      />
                    </td>
                    <td>
                      <PrimaryButton
                        text={strings.Remove}
                        onClick={() => this.removeRow(index)}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <PrimaryButton text={strings.AddRow} onClick={this.addRow} />
            <PrimaryButton
              text={strings.Submit}
              onClick={this.handleSubmit}
              disabled={!formValid}
            />
            <PrimaryButton text={strings.Cancel} onClick={this.resetForm} />
          </div>
        )}
      </div>
    )
  }
}